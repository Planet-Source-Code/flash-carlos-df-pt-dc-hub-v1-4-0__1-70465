VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'
'---------------------------------------------------------------------------------------
'Hello this is the add-ins template that you can use to make your own add-ins for PTDCH
'Most of the template has already been setup up for you and is read to use.
'Just some little notes incase you have problums running any add-ins
'Each add-ins class must have it's propertie Name as "Main"
'When your ready to compile your add-ins. the filename must be the same as
'you have the project name in Project->Properies
'For example this projects name is called PTDCH_Plugin_Template
'So when you compile you must compile it to PTDCH_Plugin_Template.dll

'Well that all you need to know for now.
'Ok installing the add-ins is easy all you need to do is drop the dll into
'the "Plugins" folder the when the IDE start your plug-ins

'WARNING: The NAME of the DLL file has to be the same of the PROJECT NAME(Project/Properies)!!
'         This because, the plugins are installed automatically, the plugin
'         server is basead in file name off the *.dll :)
'---------------------------------------------------------------------------------------
Option Explicit

'Set plugin properties
Private Const mName          As String = "PTDCH Plugin Template"
Private Const mVersion       As String = "1.00"
Private Const mAuthor        As String = "Author Name"
Private Const mDescription   As String = "Plugin Description"
Private Const mRelease       As Date = #1/20/2008 7:15:40 PM#
Private Const mComments      As String = "Plugin Comments"
Private Const mUseSetup      As Boolean = True 'Set True if availabed plugin settings
Private Const mUseEvents     As Boolean = True 'Set True if availabed plugin events

Private Sub Class_Initialize()
    '------------------------------------------------------------------
    'Purpose:   When the class is loaded
    '           The frmHide is loaded for the class to
    '           stay active process
    '
    'Params:    none
    '
    'Returns:   none
    '------------------------------------------------------------------
    Set Events = New cEvents
    Call Load(frmHide)
End Sub

Private Sub Class_Terminate()
    '------------------------------------------------------------------
    'Purpose:   When the class is Terminate
    '           Clear the active processes
    '
    'Params:    none
    '
    'Returns:   none
    '------------------------------------------------------------------
     Set Events = Nothing
    Call Unload(frmHide)
    Set frmHide = Nothing
End Sub

Public Function LoadPlug() As Integer
    '------------------------------------------------------------------
    'Purpose:   This just allows the IDE to know
    '           if the plug-ins was loaded
    '
    'Params:    none
    '
    'Returns:   1 = plug-ins was loaded
    '------------------------------------------------------------------
     LoadPlug = 1
End Function

'===================START Plugins properties===========================
'----------------------------------------------------------------------
'Purpose:   Get the Plugin properties
'
'Params:    Ex: plugin_name(array_index).Name
'
'Returns:   Property value
'----------------------------------------------------------------------
Public Property Get Name() As String
    Name = mName
End Property

Public Property Get Version() As String
    Version = mVersion
End Property

Public Property Get Author() As String
    Author = mAuthor
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Get ReleaseDate() As String
    ReleaseDate = mRelease
End Property

Public Property Get Comments() As String
    Comments = mComments
End Property

Public Property Get UseSetup() As Boolean
    UseSetup = mUseSetup
End Property

Public Property Get UseEvents() As Boolean
    UseEvents = mUseEvents
End Property

Public Property Get Enabled() As Boolean
    Enabled = g_Enabled
End Property
'====================END Plugins properties============================

Public Sub Ini(ByRef objGlobal() As Object)
    '------------------------------------------------------------------
    'Purpose:   When the plugin is loaded
    '
    'Params:
    '           objGlobal(): Array objects link to core
    '
    'Returns:   none
    '------------------------------------------------------------------

    On Error GoTo Err
 
    'Set link objects for all forms
    Set frmHub = objGlobal(0)
    Set g_colForms = objGlobal(1)
    
    'Set link objects for all classes
    Set g_objChatRoom = objGlobal(2)
    Set g_objFunctions = objGlobal(3)
    '
    Set g_colIPBans = objGlobal(4)
    Set g_colUsers = objGlobal(5)
    Set g_colCommands = objGlobal(6)
    Set g_colMessages = objGlobal(7)
    Set g_colLanguages = objGlobal(8)
    Set g_colSWinsocks = objGlobal(9)
    Set g_colSVariables = objGlobal(10)
    '
    Set g_objRegistered = objGlobal(11)
    Set g_objSettings = objGlobal(12)
    Set g_objRegExps = objGlobal(13)
    Set g_objScheduler = objGlobal(14)
    Set g_objFileAccess = objGlobal(15)
    Set g_objStatus = objGlobal(16)
    
    Set g_objHighlighter = objGlobal(17)
    Set g_objSCI = objGlobal(18)
    Set g_objComDialog = objGlobal(19)
    Set g_objSQLite = objGlobal(20)
    Set g_objActiveX = objGlobal(21)
    Set g_colDictionary = objGlobal(22)
    Set g_objTimer = objGlobal(23)
    Set g_objTimersCol = objGlobal(24)

    G_PLGPATH = App.Path
    G_APPPATH = frmHub.GetVariable("G_APPPATH")

    'Set desabled by defaut (NOT turn True unless you know what you are doing.)
    g_Enabled = False
    
    Exit Sub

Err:
    HandleError Err.Number, Err.Description, Erl & "|" & "Main.Ini()"
End Sub

Public Sub Term()
    '------------------------------------------------------------------
    'Purpose:   When the plugin is terminate (used at Unload frmHub)
    '
    'Params:    none
    '
    'Returns:   none
    '------------------------------------------------------------------
    RunEvent "UnloadMain"
End Sub

Public Sub LoadForm()
    '------------------------------------------------------------------
    'Purpose:   When the plugin setup is loaded (used RTB menu or
    '           Plugin Config (frmHub Tab Misc/Plugins))
    'Params:    none
    '
    'Returns:   none
    '------------------------------------------------------------------
    On Error GoTo Err

    frmMain.Show vbModal
    
    Exit Sub
Err:
    HandleError Err.Number, Err.Description, Erl & "|" & "Main.LoadForm()"
End Sub

Public Function RunEvent(ParamArray VarExpr() As Variant) As Variant
    '------------------------------------------------------------------
    'Purpose:   This runs the plugin events
    '           Called when the hub owner presses the envents
    '
    'Params:    VarExpr(): Param Array
    '
    'Returns:   Depends of the envent
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
    
    On Error GoTo Err
    
       Select Case LCase$(CStr(VarExpr(0)))
            Case "startedserving"
                Events.StartedServing
            Case "stoppedserving"
                Events.StoppedServing
            Case "userconnected"
                Events.UserConnected VarExpr(1)
            Case "regconnected"
                Events.RegConnected VarExpr(1)
            Case "opconnected"
                Events.OpConnected VarExpr(1)
            Case "userquit"
                Events.UserQuit VarExpr(1)
            Case "addedpermban"
                Events.AddedPermBan CStr(VarExpr(1)), CStr(VarExpr(2)), CStr(VarExpr(3)), CStr(VarExpr(4))
            Case "addedtempban"
                Events.AddedTempBan CStr(VarExpr(1)), CLng(VarExpr(2)), CStr(VarExpr(3)), CStr(VarExpr(4)), CStr(VarExpr(5))
            Case "addedregisteredUser"
                Events.AddedRegisteredUser CStr(VarExpr(1)), CStr(VarExpr(2)), CInt(VarExpr(3)), CStr(VarExpr(4)), CLng(VarExpr(5))
            Case "removedregistereduser"
                Events.RemovedRegisteredUser CStr(VarExpr(1))
            Case "attemptedconnection"
                Events.AttemptedConnection CStr(VarExpr(1))
            Case "startedredirecting"
                Events.StartedRedirecting
            Case "predataarrival"
                RunEvent = CVar(Events.PreDataArrival(VarExpr(1), CStr(VarExpr(2))))
            Case "dataarrival"
                Events.DataArrival VarExpr(1), CStr(VarExpr(2))
            Case "custcomarrival"
                Events.CustComArrival VarExpr(1), VarExpr(2), CStr(VarExpr(3)), CBool(VarExpr(4))
            Case "massmessage"
                Events.MassMessage CStr(VarExpr(1))
            Case "failedconf"
                RunEvent = CVar(Events.FailedConf(VarExpr(1), CInt(VarExpr(2))))
            Case "unloadmain"
                Events.UnloadMain
            Case "submain"
                Events.SubMain
            Case "reload"
                Events.Reload
            Case "switch"
                Events.Switch CBool(VarExpr(1))
            Case "preconnectionrequest"
                RunEvent = CBool(Events.PreConnectionRequest(VarExpr(1), CLng(VarExpr(2))))
            Case "tmrAPI_Timer"
                Events.tmrAPI_Timer CInt(VarExpr(1))
    End Select

    Exit Function
Err:
    HandleError Err.Number, Err.Description, Erl & "|" & "cEvents.RunEvent(""" & CStr(VarExpr(0)) & """)"
End Function
