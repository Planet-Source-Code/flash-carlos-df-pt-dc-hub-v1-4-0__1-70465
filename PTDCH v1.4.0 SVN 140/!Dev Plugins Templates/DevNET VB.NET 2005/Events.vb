Public Class cEvents
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

    '--------------------------------------------------------------
    'Triggered when hub start serving
    '--------------------------------------------------------------
    Public Sub StartedServing()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.StartedServing()")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when hub stop serving
    '--------------------------------------------------------------
    Public Sub StoppedServing()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.StoppedServing()")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when a non registered user has logged in.
    '--------------------------------------------------------------
    Public Sub UserConnected(ByRef objUser As Object)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.UserConnected(" & objUser.sName & ")")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when a registered user has logged in.
    '--------------------------------------------------------------
    Public Sub RegConnected(ByRef objUser As Object)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.RegConnected(" & objUser.sName & ")")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when an Operator has logged in.
    '--------------------------------------------------------------
    Public Sub OpConnected(ByRef objUser As Object)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.OpConnected(" & objUser.sName & ")")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when a user has logged out.
    '--------------------------------------------------------------
    Public Sub UserQuit(ByRef objUser As Object)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.UserQuit(" & objUser.sName & ")")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when a new permanent ban is added to the hub's database.( colIPBans.Add )
    '--------------------------------------------------------------
    Public Sub AddedPermBan(ByRef strIP As String, ByRef strNick As String, ByRef strBannedBy As String, ByRef strReason As String)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.AddedPermBan(" & strIP & ")")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when a new user is added to the hub's database. ( colRegistered.Add )
    '--------------------------------------------------------------
    Public Sub AddedTempBan(ByRef strIP As String, ByRef lngMinutes As Long, ByRef strNick As String, ByRef strBannedBy As String, ByRef strReason As String)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.AddedTempBan(""" & strIP & """)")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when a new user is added to the hub's database. ( colRegistered.Add )
    '--------------------------------------------------------------
    Public Sub AddedRegisteredUser(ByRef strName As String, ByRef strPassword As String, ByRef intClass As Integer, ByRef strAdminName As String, ByRef lngMin As Long)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.AddedRegisteredUser(""" & strName & """)")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when a new user is removed from the hub's database. ( colRegistered.Rem )
    '--------------------------------------------------------------
    Public Sub RemovedRegisteredUser(ByRef strName As String)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.RemovedRegisteredUser(""" & strName & """)")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when there is a new connection to the listening port(s).
    '--------------------------------------------------------------
    Public Sub AttemptedConnection(ByRef strIP As String)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.AttemptedConnection(""" & strIP & """)")
        End Try
    End Sub

    Public Function PreConnectionRequest(ByRef wskShock As Object, ByRef requestID As Integer) As Boolean
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
        Try
            'Implicite default is False, but it does not hurt to specify it.
            PreConnectionRequest = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.PreConnectionRequest(""" & wskShock.RemoteHostIP & ", " & requestID & """)")
        End Try
    End Function

    Public Function PreDataArrival(ByRef objUser As Object, ByRef strCommand As String) As String
        '------------------------------------------------------------------
        'Purpose:   Allow data string to be modified before the hub see it
        '           Triggered whenever the hub receive data from an accepted connection
        '
        'Params:    objUser: The user sending the data.
        '                    strCommand: The data received. (Might have already been modify by an other plugin.)
        '
        'Comments:  PreDataArrival = Empty : Prevent any further processing by the Hub and other plugin/script.
        '           PreDataArrival = [same or modified string] : The data will be further processed by the hub and plugins/scripts

        '------------------------------------------------------------------
        Try
            'DO NOT remove this line unless you know what you are doing.
            Return strCommand
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.PreDataArrival(""" & objUser.sName & ", " & strCommand & """)")
        End Try
    End Function

    '--------------------------------------------------------------
    'Triggered whenever a connected user send data.
    '--------------------------------------------------------------
    Public Sub DataArrival(ByRef objUser As Object, ByRef strData As String)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.DataArrival(""" & objUser.sName & ", " & strData & """)")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when the hub command prefix is detected in a user's message to the hub.
    '--------------------------------------------------------------
    Public Sub CustComArrival(ByRef objUser As Object, ByRef objCommand As Object, ByRef strMessage As String, ByRef blnMainChat As Boolean)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.CustComArrival(""" & objUser.sName & ", " & objCommand & ", " & blnMainChat & """)")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when the hub command prefix is detected in a user's message to the hub.
    '--------------------------------------------------------------
    Public Sub StartedRedirecting()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.StartedRedirecting()")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when sending a Mass message. (Tru interface)
    '--------------------------------------------------------------
    Public Sub MassMessage(ByRef strMessage As String)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.MassMessage(""" & strMessage & """)")
        End Try
    End Sub

    Public Function FailedConf(ByRef objUser As Object, ByRef intType As Object) As Boolean
        '------------------------------------------------------------------
        'Purpose:   Allow to bypass the hub's rules settings
        '           Triggered whenever a user fail a hub's rules.
        '
        'Params:    objUser: The user sending the data.
        '           intType: Integer telling what rule the user failled
        '
        'Comments:  See: Enum enuAlert in the help file.
        '
        '           False = Let the hub take care of the user,
        '           message the user,  banning and/or disconnecting him.
        '------------------------------------------------------------------
        Try
            'Implicite default is False, but it does not hurt to specify it.
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.FailedConf(""" & objUser.sName & ", " & intType & """)")
        End Try
    End Function

    '--------------------------------------------------------------
    'Triggered when this plugin is stop. Do your cleanup here.
    '--------------------------------------------------------------
    Public Sub UnloadMain()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.UnloadMain()")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when this plugin is start
    '--------------------------------------------------------------
    Public Sub SubMain()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.SubMain()")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered when owners ask for this plugin to be reset.
    '--------------------------------------------------------------
    Public Sub Reload()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.Reload()")
        End Try
    End Sub

    '--------------------------------------------------------------
    'Triggered whenever hub's serving state is changed.
    '--------------------------------------------------------------
    Public Sub Switch(ByRef blnState As Boolean)
        Try

            'DO NOT remove this code unless you know what you are doing.
            'Set plugin state
            g_Enabled = blnState

            frmHub.PlgRefreshGUI()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, g_Main.Name & "-" & "cEvents.Switch(""" & blnState & """)")
        End Try
    End Sub
End Class