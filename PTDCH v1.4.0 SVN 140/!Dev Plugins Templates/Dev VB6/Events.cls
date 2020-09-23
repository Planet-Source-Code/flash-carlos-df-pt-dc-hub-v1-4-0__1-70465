VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------
'Purpose:   Plugin events Class interface
'           Called when the hub owner presses the envents
'
'Params:    VarExpr(): Param Array
'
'Returns:   Depends of the envents
'           Events with returns:
'                   -PreDataArrival
'                         = Empty: the data won't be processed
'                        <> Empty: the data will be processed
'                   -FailedConf
'                         = True:  the data won't be processed
'                         = False: the data will be processed
'                   -PreConnectionRequest
'                         = True: the data won't be processed
'                         = False: the data will be processed
'------------------------------------------------------------------
Option Explicit

'--------------------------------------------------------------
'Triggered when hub start serving
'--------------------------------------------------------------
Public Sub StartedServing()
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.StartedServing()"
End Sub

'--------------------------------------------------------------
'Triggered when hub stop serving
'--------------------------------------------------------------
Public Sub StoppedServing()
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.StoppedServing()"
End Sub

'--------------------------------------------------------------
'Triggered when a non registered user has logged in.
'--------------------------------------------------------------
Public Sub UserConnected(ByRef objUser As Variant)
1:    On Error GoTo Err

3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.UserConnected(""" & objUser.sName & """)"
End Sub

'--------------------------------------------------------------
'Triggered when a registered user has logged in.
'--------------------------------------------------------------
Public Sub RegConnected(ByRef objUser As Variant)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.RegConnected(""" & objUser.sName & """)"
End Sub

'--------------------------------------------------------------
'Triggered when an Operator has logged in.
'--------------------------------------------------------------
Public Sub OpConnected(ByRef objUser As Variant)
1:    On Error GoTo Err

3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.OpConnected(""" & objUser.sName & """)"
End Sub

'--------------------------------------------------------------
'Triggered when a user has logged out.
'--------------------------------------------------------------
Public Sub UserQuit(ByRef objUser As Variant)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.UserQuit(""" & objUser.sName & """)"
End Sub

'--------------------------------------------------------------
'Triggered when a new permanent ban is added to the hub's database.( colIPBans.Add )
'--------------------------------------------------------------
Public Sub AddedPermBan(ByRef strIP As String, ByRef strNick As String, ByRef strBannedBy As String, ByRef strReason As String)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.AddedPermBan(""" & strIP & """)"
End Sub

'--------------------------------------------------------------
'Triggered when a new temporary ban is added to the hub's database.( colIPBans.Add )
'--------------------------------------------------------------
Public Sub AddedTempBan(ByRef strIP As String, ByRef lngMinutes As Long, ByRef strNick As String, ByRef strBannedBy As String, ByRef strReason As String)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.AddedPermBan(""" & strIP & """)"
End Sub

'--------------------------------------------------------------
'Triggered when a new user is added to the hub's database. ( colRegistered.Add )
'--------------------------------------------------------------
Public Sub AddedRegisteredUser(ByRef strName As String, ByRef strPassword As String, ByRef intClass As Integer, ByRef strAdminName As String, ByRef lngMin As Long)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.AddedRegisteredUser(""" & strName & """)"
End Sub

'--------------------------------------------------------------
'Triggered when a new user is removed from the hub's database. ( colRegistered.Rem )
'--------------------------------------------------------------
Public Sub RemovedRegisteredUser(ByRef strName As String)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.AddedRegisteredUser(""" & strName & """)"
End Sub

'--------------------------------------------------------------
'Triggered when there is a new connection to the listening port(s).
'--------------------------------------------------------------
Public Sub AttemptedConnection(ByRef strIP As String)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.AttemptedConnection(""" & strIP & """)"
End Sub

Public Function PreConnectionRequest(ByRef wskShock As Variant, ByRef requestID As Long) As Boolean
'------------------------------------------------------------------
'Purpose:   Bypass hub's new connection requests
'           Triggered when a new connection initiate
'
'Params:    wskShock as Sockets collection
'           requestID as winSocks request ID
'
'Comments:  PreConnectionRequest = False: hub will resume processing
'           PreConnectionRequest = True: Processing halt here.
'           Hub and other plugins/script won't have a chance to process the connection request.
'
'  If you choose to not let the user in, PreConnectionRequest = True,
'  then you MUST close the socket in here before exiting.
'------------------------------------------------------------------
    
      'Implicite default is False, but it does not hurt to specify it.
17:   PreConnectionRequest = False

19:   Exit Function

21:
Err:
22:   PreConnectionRequest = False
23:   HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.PreConnectionRequest(""" & wskShock.RemoteHostIP & ", " & requestID & """)"
End Function

Public Function PreDataArrival(ByRef objUser As Variant, ByRef strCommand As String) As String
'------------------------------------------------------------------
'Purpose:   Allow data string to be modified before the hub see it
'           Triggered whenever the hub receive data from an accepted connection
'
'Params:    objUser: The user m_sending the data.
'                    strCommand: The data received. (Might have already been modify by an other plugin.)
'
'Comments:  PreDataArrival = Empty : Prevent any further processing by the Hub and other plugin/script.
'           PreDataArrival = [same or modified string] : The data will be further processed by the hub and plugins/scripts

'------------------------------------------------------------------
12:    On Error GoTo Err

       'DO NOT remove this line unless you know what you are doing.
15:    PreDataArrival = strCommand

17:    Exit Function

19:
Err:
20:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.PreDataArrival(""" & objUser.sName & ", " & strCommand & """)"
21:    PreDataArrival = strCommand
End Function

'--------------------------------------------------------------
'Triggered whenever a connected user send data.
'--------------------------------------------------------------
Public Sub DataArrival(ByRef objUser As Variant, ByRef strData As String)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.DataArrival(""" & objUser.sName & ", " & strData & """)"
End Sub

'--------------------------------------------------------------
'Triggered when the hub command prefix is detected in a user's message to the hub.
'--------------------------------------------------------------
Public Sub CustComArrival(ByRef objUser As Variant, ByRef objCommand As Variant, ByRef strMessage As String, ByRef blnMainChat As Boolean)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.CustComArrival(""" & objUser.sName & ", " & objCommand & ", " & blnMainChat & """)"
End Sub

'--------------------------------------------------------------
'Triggered when the hub command prefix is detected in a user's message to the hub.
'--------------------------------------------------------------
Public Sub StartedRedirecting()
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.StartedRedirecting()"
End Sub

'--------------------------------------------------------------
'Triggered when m_sending a Mass message. (Tru interface)
'--------------------------------------------------------------
Public Sub MassMessage(ByRef strMessage As String)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.MassMessage(""" & strMessage & """)"
End Sub

Public Function FailedConf(ByRef objUser As Variant, ByRef intType As Variant) As Boolean
'------------------------------------------------------------------
'Purpose:   Allow to bypass the hub's rules settings
'           Triggered whenever a user fail a hub's rules.
'
'Params:    objUser: The user m_sending the data.
'           intType: Integer telling what rule the user failled
'
'Comments:  See: Enum enuAlert in the help file.
'
'           False = Let the hub take care of the user,
'           message the user,  banning and/or disconnecting him.
'------------------------------------------------------------------
13:    On Error GoTo Err

      'Implicite default is False, but it does not hurt to specify it.
16:    FailedConf = False

18:    Exit Function

20:
Err:
21:   FailedConf = False
22:   HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.FailedConf(""" & objUser.sName & ", " & intType & """)"
End Function

'--------------------------------------------------------------
'Triggered when this plugin is stop. Do your cleanup here.
'--------------------------------------------------------------
Public Sub UnloadMain()
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.UnloadMain()"
End Sub

'--------------------------------------------------------------
'Triggered when this plugin is start
'--------------------------------------------------------------
Public Sub SubMain()
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.SubMain()"
End Sub

Public Sub tmrAPI_Timer(ByRef Index As Integer)
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.tmrAPI_Timer(" & Index & ")"
End Sub

'--------------------------------------------------------------
'Triggered when owners ask for this plugin to be reset.
'--------------------------------------------------------------
Public Sub Reload()
1:    On Error GoTo Err
    
3:    Exit Sub

5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.Reload()"
End Sub

'--------------------------------------------------------------
'Triggered whenever hub's serving state is changed.
'--------------------------------------------------------------
Public Sub Switch(ByRef blnState As Boolean)
1:   On Error GoTo Err
    
    'DO NOT remove this code unless you know what you are doing.
    'Set plugin state
5:   g_Enabled = blnState
      
     frmHub.PlgRefreshGUI
     
7:   Exit Sub
8:
Err:
9:   HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.Switch(""" & blnState & """)"
End Sub
