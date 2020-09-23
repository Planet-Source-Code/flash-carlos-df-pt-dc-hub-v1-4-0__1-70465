Imports System.Runtime.InteropServices
#Region " Help "
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
'you have the NAMESPACE!! -->Namespace DevNET
'For example this projects name is called DevNET
'So when you compile you must compile it to DevNET.dll

'Well that all you need to know for now.
'Ok installing the add-ins is easy all you need to do is drop the dll into
'the "Plugins" folder the when the IDE start your plug-ins

'WARNING: The NAME of the DLL file has to be the same of the NAMESPACE!!
'         This because, the plugins are installed automatically, the plugin
'         server is basead in file name off the *.dll :)
'---------------------------------------------------------------------------------------
#End Region
Namespace DevNET
    'Using a .NET Library in PTDCH!!
    'In the Project Options you need to set the flag register for COM Interop
    'This enables VS.NET to create a .tlb File
    'Generate a GUID in registry Format via the extras menue and use this guid as attribute
    <Guid("7FFBB260-AA83-4064-B9EE-ACF11660BFC4"), ComVisible(True)> _
    Public Class Main
#Region " Private's "

        'Plugin properties
        Private Const mName As String = "DevTest VB.NET 2005"
        Private Const mVersion As String = "1.00"
        Private Const mAuthor As String = "Author Name"
        Private mDescription As String = "Plugin use version " & _
                                System.Reflection.Assembly.GetExecutingAssembly().ImageRuntimeVersion.ToString & " of the MS .NET Framework"
        Private Const mRelease As Date = #1/20/2008 7:15:40 PM#
        Private Const mComments As String = "Comments"
        Private Const mUseSetup As Boolean = True 'Set True if availabed plugin settings
        Private Const mUseEvents As Boolean = True 'Set True if availabed plugin events

        Private m_frmMain As frmMain

#End Region

#Region " Class Events "

        Public Sub New()
            '------------------------------------------------------------------
            'Purpose:   When the class is loaded
            '
            '
            'Params:    Events: create a ref.. for the class events
            '
            'Returns:   none
            '------------------------------------------------------------------
            MyBase.New()
            Events = New cEvents
        End Sub

        Protected Overrides Sub Finalize()
            '------------------------------------------------------------------
            'Purpose:   When the class is Terminate
            '           Clear the active processes
            '
            'Params:    none
            '
            'Returns:   none
            '------------------------------------------------------------------
            MyBase.Finalize()
        End Sub

#End Region

#Region " Plugin Properties "

        '----------------------------------------------------------------------
        'Purpose:   Get the Plugin properties
        '
        'Params:    Ex: plugin_name(array_index).Name
        '
        'Returns:   Property value
        '----------------------------------------------------------------------

        <ComVisible(True)> _
        Public ReadOnly Property Name() As String
            Get
                Return mName
            End Get
        End Property

        <ComVisible(True)> _
        Public ReadOnly Property Version() As String
            Get
                Return mVersion
            End Get
        End Property

        <ComVisible(True)> _
        Public ReadOnly Property Author() As String
            Get
                Return mAuthor
            End Get
        End Property

        <ComVisible(True)> _
        Public ReadOnly Property Description() As String
            Get
                Return mDescription
            End Get
        End Property

        <ComVisible(True)> _
        Public ReadOnly Property ReleaseDate() As Date
            Get
                Return mRelease
            End Get
        End Property

        <ComVisible(True)> _
        Public ReadOnly Property Comments() As String
            Get
                Return mComments
            End Get
        End Property

        <ComVisible(True)> _
        Public ReadOnly Property UseSetup() As Boolean
            Get
                Return mUseSetup
            End Get
        End Property

        <ComVisible(True)> _
        Public ReadOnly Property UseEvents() As Boolean
            Get
                Return mUseEvents
            End Get
        End Property

        <ComVisible(True)> _
        Public ReadOnly Property Enabled() As Boolean
            Get
                Return g_Enabled
            End Get
        End Property

#End Region

#Region " Plugin Subs "
        <ComVisible(True)> _
        Public Sub Ini(ByRef objGlobal() As Object)
            '------------------------------------------------------------------
            'Purpose:   When the plugin is loaded
            '
            'Params:
            '           objGlobal(): Array objects link to core
            '
            'Returns:   none
            '------------------------------------------------------------------

            Try
                'Set link objects for all forms
                frmHub = objGlobal(0)
                g_colForms = objGlobal(1)

                'Set link objects for all classes
                g_objChatRoom = objGlobal(2)
                g_objFunctions = objGlobal(3)
                '
                g_colIPBans = objGlobal(4)
                g_colUsers = objGlobal(5)
                g_colCommands = objGlobal(6)
                g_colMessages = objGlobal(7)
                g_colLanguages = objGlobal(8)
                g_colSWinsocks = objGlobal(9)
                g_colSVariables = objGlobal(10)
                '
                g_objRegistered = objGlobal(11)
                g_objSettings = objGlobal(12)
                g_objRegExps = objGlobal(13)
                g_objScheduler = objGlobal(14)
                g_objFileAccess = objGlobal(15)
                g_objStatus = objGlobal(16)

                g_objHighlighter = objGlobal(17)
                g_objSCI = objGlobal(18)
                g_objComDialog = objGlobal(19)
                g_objSQLite = objGlobal(20)
                g_objActiveX = objGlobal(21)
                g_colDictionary = objGlobal(22)

                G_PLGPATH = Application.StartupPath
                G_APPPATH = frmHub.GetVariable("G_APPPATH")

                'Set desabled by defaut (NOT turn True unless you know what you are doing.)
                g_Enabled = False

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "Main.Ini()")
            End Try
        End Sub

        <ComVisible(True)> _
        Public Sub Term()
            '------------------------------------------------------------------
            'Purpose:   When the plugin is terminate (used at Unload frmHub)
            '
            'Params:    none
            '
            'Returns:   none
            '------------------------------------------------------------------
            RunEvent("UnloadMain")
        End Sub

        <ComVisible(True)> _
        Public Sub LoadForm()
            '------------------------------------------------------------------
            'Purpose:   When the plugin setup is loaded (used RTB menu or
            '           Plugin Config (frmHub Tab Misc/Plugins))
            'Params:    none
            '
            'Returns:   none
            '------------------------------------------------------------------
            Try
                m_frmMain = New frmMain
                m_frmMain.Show()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "Main.LoadForm()")
            End Try
        End Sub
#End Region

#Region " Plugin Functions "

        <ComVisible(True)> _
        Public Function LoadPlug() As Short
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

        <ComVisible(True)> _
        Public Function RunEvent(ByVal ParamArray VarExpr() As Object) As Object
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
            Try

                Select Case LCase$(CStr(VarExpr(0)))
                    Case "startedserving"
                        Events.StartedServing()
                    Case "stoppedserving"
                        Events.StoppedServing()
                    Case "userconnected"
                        Events.UserConnected(VarExpr(1))
                    Case "regconnected"
                        Events.RegConnected(VarExpr(1))
                    Case "opconnected"
                        Events.OpConnected(VarExpr(1))
                    Case "userquit"
                        Events.UserQuit(VarExpr(1))
                    Case "addedpermban"
                        Events.AddedPermBan(CStr(VarExpr(1)), CStr(VarExpr(2)), CStr(VarExpr(3)), CStr(VarExpr(4)))
                    Case "addedtempban"
                        Events.AddedTempBan(CStr(VarExpr(1)), CLng(VarExpr(2)), CStr(VarExpr(3)), CStr(VarExpr(4)), CStr(VarExpr(5)))
                    Case "addedregisteredUser"
                        Events.AddedRegisteredUser(CStr(VarExpr(1)), CStr(VarExpr(2)), CInt(VarExpr(3)), CStr(VarExpr(4)), CLng(VarExpr(5)))
                    Case "removedregistereduser"
                        Events.RemovedRegisteredUser(CStr(VarExpr(1)))
                    Case "attemptedconnection"
                        Events.AttemptedConnection(CStr(VarExpr(1)))
                    Case "startedredirecting"
                        Events.StartedRedirecting()
                    Case "predataarrival"
                        RunEvent = CObj(Events.PreDataArrival(VarExpr(1), CStr(VarExpr(2))))
                    Case "dataarrival"
                        Events.DataArrival(VarExpr(1), CStr(VarExpr(2)))
                    Case "custcomarrival"
                        Events.CustComArrival(VarExpr(1), VarExpr(2), CStr(VarExpr(3)), CBool(VarExpr(4)))
                    Case "massmessage"
                        Events.MassMessage(CStr(VarExpr(1)))
                    Case "failedconf"
                        RunEvent = CObj(Events.FailedConf(VarExpr(1), CInt(VarExpr(2))))
                    Case "unloadmain"
                        Events.UnloadMain()
                    Case "submain"
                        Events.SubMain()
                    Case "reload"
                        Events.Reload()
                    Case "switch"
                        Events.Switch(CBool(VarExpr(1)))
                    Case "preconnectionrequest"
                        RunEvent = CBool(Events.PreConnectionRequest(VarExpr(1), CLng(VarExpr(2))))
                End Select

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-cEvents.RunEvent(""" & CStr(VarExpr(0)) & """)")
            End Try
            Return Nothing
        End Function
#End Region
    End Class
End Namespace