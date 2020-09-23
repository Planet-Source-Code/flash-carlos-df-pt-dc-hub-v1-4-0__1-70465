VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type typPlan
    User            As String
    DateTime        As Date
    Enabled         As Boolean
    Command         As String
    Parameter       As String
    Increase        As Integer
    IncreaseType    As enIncrease
    Description     As String
    Status          As String
End Type

Public Enum enIncrease
    iNone = 0
    iMinute = 1
    iHour = 2
    iDay = 3
    iMonth = 4
End Enum
    
Private Plan()                  As typPlan
Private TempPlan()              As typPlan

Private WithEvents m_Timer      As Timer
Attribute m_Timer.VB_VarHelpID = -1
Private m_frmParent             As frmParent
    
Private m_lvwPlan               As ListItems
Private m_lvwItem               As ListItem

Private Sub Class_Initialize()
1:    Set m_lvwPlan = frmHub.lvwPlan.ListItems
      'Create our dummy form to get it's Timer
3:    Set m_frmParent = New frmParent
4:    Set m_Timer = m_frmParent.tmrTimer
5:    m_Timer.Enabled = False
6:    m_Timer.Interval = 1000
7:    Call PreInitialize
End Sub

Private Sub Class_Terminate()
1:    Call PreTerminate
2:    Set m_lvwPlan = Nothing
3:    m_Timer.Enabled = False
4:    Set m_Timer = Nothing
5:    Set m_frmParent = Nothing
6:    Erase Plan
End Sub

'------------------------------------------------------------------------------
'Publics
'------------------------------------------------------------------------------
Public Sub ShowAddDialog(ByRef frmOwner As Form)
1:    Dim frm As New frmPlan
2:    On Error GoTo Err
3:    With frm
4:        .Caption = "Add Schelduler"
5:        .Show vbModal, frmOwner
6:    End With
7:    Exit Sub
8:
Err:
9:    HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.ShowAddDialog()"
End Sub

Public Sub ShowEditDialog(ByRef sUser As String, _
                          ByRef sCommand As String, _
                          ByRef sParameter As String, _
                          ByRef frmOwner As Form)
4:  Dim frm As New frmPlan
5:  Dim lngTemp As Long

7:  On Error GoTo Err
    
9:  lngTemp = IsRegistered(sUser, sCommand, sParameter)
   
11:    If lngTemp <> -1 Then
12:        With frm
13:          .Caption = "Edit Schelduler"
        
15:          .txtBox(0).Text = Plan(lngTemp).User
16:          .txtBox(1).Text = Plan(lngTemp).DateTime
17:          .txtBox(2).Text = Plan(lngTemp).Command
18:          .txtBox(3).Text = Plan(lngTemp).Parameter
19:          .txtBox(4).Text = Plan(lngTemp).Increase
20:          .txtBox(5).Text = Plan(lngTemp).Description
             '
              Select Case Plan(lngTemp).IncreaseType
                Case 0: .cmbBox.Text = "None"
                Case 1: .cmbBox.Text = "Minute(s)"
                Case 2: .cmbBox.Text = "Hour(s)"
                Case 3: .cmbBox.Text = "Day(s)"
                Case 4: .cmbBox.Text = "Month(s)"
22:          End Select
             '
24:          If Plan(lngTemp).Enabled Then _
                  .chkBox(0).Value = 1 _
             Else .chkBox(0).Value = 0
             '
28:          If Plan(lngTemp).Increase Then _
                    .chkBox(1).Value = 1 _
             Else .chkBox(1).Value = 0

32:         .Show vbModal, frmOwner
33:       End With
34:    End If

36:   Exit Sub

38:
Err:
39:   HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.ShowEditDialog()"
End Sub

Public Function AddPlan(ByRef sUser As String, _
                        ByRef dDateTime As Date, _
                        ByRef bEnabled As Boolean, _
                        ByRef sCommand As String, _
               Optional ByRef sParameter As String = "", _
               Optional ByRef iIncrease As Integer = 0, _
               Optional ByRef eIncreaseType As enIncrease = iDay, _
               Optional ByRef sDescription As String = "", _
               Optional ByRef bOverwrite As Boolean = True) As Boolean
               
10:    Dim lngIndex As Long
      
12:    On Error GoTo Err

13:    AddPlan = True
       
      'Check if is nothing.. If rem all bots.. the var is nothing
15:    If Not IsNothingPlan Then
            'Check if it has already been registered
17:          lngIndex = IsRegistered(sUser, sCommand, sParameter)
18:         If lngIndex = -1 Then
19:             ReDim Preserve Plan(LBound(Plan) To UBound(Plan) + 1) As typPlan
20:             lngIndex = UBound(Plan)
21:         Else
22:             If Not bOverwrite Then Exit Function
23:         End If
24:   Else
25:         ReDim Plan(0) As typPlan
26:         m_Timer.Enabled = False
27:         lngIndex = 0
28:   End If

30:   With Plan(lngIndex)
31:        .User = sUser
32:        .DateTime = dDateTime
           'Force desabled if date < now
34:        If .DateTime < Now Then _
                .Enabled = False _
           Else .Enabled = bEnabled
37:        .Command = sCommand
38:        .Parameter = sParameter
39:        .Increase = iIncrease
40:        .IncreaseType = eIncreaseType
41:        .Description = sDescription
42:        .Status = ""
43:   End With

45:   Call RefreshLvw
46:   Call CheckTimer

48:   Exit Function

50:   AddPlan = True
51:
Err:
53:   HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.AddPlan(""" & sUser & """)"
End Function

Public Function RemPlan(ByRef sUser As String, _
                        ByRef sCommand As String, _
                        ByRef sParameter As String) As Boolean
3:   Dim intIndex As Integer
4:   Dim lngIndex As Long
5:   Dim intAux As Integer
6:   On Error GoTo Err
    
      'Check if it is registered
9:    lngIndex = IsRegistered(sUser, sCommand, sParameter)

11:    If lngIndex = -1 Then
12:         RemPlan = False
13:    Else
           'It is only verified 1 item exists
15:        If UBound(Plan) = 0 Then
16:             Erase Plan
17:             m_Timer.Enabled = False
18:             m_lvwPlan.Clear
19:             RemPlan = True
20:             Exit Function
21:        End If

        'Clear temp array.. Note: is 'UBound(TempPlan) - 1' -->because 1 has deleted
24:        Erase TempPlan()
25:        ReDim TempPlan(LBound(Plan) To UBound(Plan) - 1) As typPlan
        
        'Copy array info to temp array and exluind the plan deleted
28:        For intIndex = LBound(Plan) To UBound(Plan)
        
30:            If Plan(intIndex).User <> sUser And _
               Plan(intIndex).Command <> sCommand And _
               Plan(intIndex).Parameter <> sParameter Then
33:               With Plan(intIndex)
34:                    TempPlan(intAux).User = .User
35:                    TempPlan(intAux).DateTime = .DateTime
36:                    TempPlan(intAux).Enabled = .Enabled
37:                    TempPlan(intAux).Command = .Command
38:                    TempPlan(intAux).Parameter = .Parameter
39:                    TempPlan(intAux).Increase = .Increase
40:                    TempPlan(intAux).IncreaseType = .IncreaseType
41:                    TempPlan(intAux).Description = .Description
42:                    TempPlan(intAux).Status = .Status
43:               End With
44:               intAux = intAux + 1
45:            End If
            
47:        Next
        
           'Clear array...
50:        Erase Plan()
51:        ReDim Plan(LBound(TempPlan) To UBound(TempPlan)) As typPlan
        
           'Copy array info from de temp array
54:        For intIndex = LBound(TempPlan) To UBound(TempPlan)
55:            With TempPlan(intIndex)
56:                 Plan(intIndex).User = .User
57:                 Plan(intIndex).DateTime = .DateTime
58:                 Plan(intIndex).Enabled = .Enabled
59:                 Plan(intIndex).Command = .Command
60:                 Plan(intIndex).Parameter = .Parameter
61:                 Plan(intIndex).Increase = .Increase
62:                 Plan(intIndex).IncreaseType = .IncreaseType
63:                 Plan(intIndex).Description = .Description
64:                    Plan(intIndex).Status = .Status
65:            End With
66:        Next
       
           'Clear temp array
69:        Erase TempPlan
        
           'Refresh ListView
72:        Call RefreshLvw

74:        RemPlan = True
75:    End If

77:   Exit Function

79:
Err:
80:   HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.RemPlan(""" & sUser & """)"
End Function
'------------------------------------------------------------------------------
'End Publics
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'Privates
'------------------------------------------------------------------------------
Private Sub RefreshLvw()
1:     Dim lngLoop     As Long
2:     On Error GoTo Err

4:     m_lvwPlan.Clear
5:     If IsNothingPlan Then Exit Sub
       
7:     For lngLoop = LBound(Plan) To UBound(Plan)
     
9:        With Plan(lngLoop)
10:            Set m_lvwItem = m_lvwPlan.Add(, lngLoop & "s", .User)
11:            m_lvwItem.SubItems(1) = .DateTime
12:            m_lvwItem.SubItems(2) = .Enabled
13:            m_lvwItem.SubItems(3) = .Command
14:            m_lvwItem.SubItems(4) = .Parameter
15:            m_lvwItem.SubItems(5) = .Increase
            
               Select Case .IncreaseType
                    Case iNone: m_lvwItem.SubItems(6) = "None"
                    Case iMinute: m_lvwItem.SubItems(6) = "Minute(s)"
                    Case iHour: m_lvwItem.SubItems(6) = "Hour(s)"
                    Case iDay: m_lvwItem.SubItems(6) = "Day(s)"
                    Case iMonth: m_lvwItem.SubItems(6) = "Month(s)"
17:            End Select
               
19:            m_lvwItem.SubItems(7) = .Description
20:            m_lvwItem.SubItems(8) = .Status
21:       End With
       
23:    Next
        
25:    Set m_lvwItem = Nothing

27:   Exit Sub
28:
Err:
29:   HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.RefreshLvw()"
End Sub

Private Function IsRegistered(ByRef sUser As String, _
                              ByRef sCommand As String, _
                              ByRef sParameter As String) As Long
3:    Dim lngLoop As Long

5:    On Error GoTo Err
    
       'Set to -1, meaning it hasn't found the plan User
8:    IsRegistered = -1

10:    If IsNothingPlan Then Exit Function
       'Note: Key is sUser + sCommand + sParameter
       'Loop through and see if the name matches any; if it does, return array index
13:    For lngLoop = LBound(Plan) To UBound(Plan)
14:          If Plan(lngLoop).User = sUser And _
             Plan(lngLoop).Command = sCommand And _
             Plan(lngLoop).Parameter = sParameter Then _
                    IsRegistered = lngLoop: Exit For
18:    Next

20:    Exit Function

22:
Err:
23:    HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.IsRegistered()"
End Function

'Check if the var type is nothing
Private Function IsNothingPlan() As Boolean
1:    Dim strTest As String

3:    IsNothingPlan = False

5:    On Error GoTo Err
6:    strTest = Plan(0).User

8:    Call CheckTimer

10:   Exit Function
11:
Err:
12:   Err.Clear
13:   IsNothingPlan = True
14:   Resume Next
End Function

'Check if exist plan enabled
Private Sub CheckTimer()
1:    Dim intIndex As Integer

3:    For intIndex = LBound(Plan) To UBound(Plan)
4:        If Plan(intIndex).Enabled Then
5:             m_Timer.Enabled = True
6:             Exit Sub
7:        End If
8:    Next

10:    m_Timer.Enabled = False

12:    Exit Sub

14:
Err:
15:   HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.CheckTimer()"
End Sub

Private Sub PreInitialize()
1:     Dim objXML          As clsXMLParser
2:     Dim objNode         As clsXMLNode
3:     Dim colNodes        As Collection
4:     Dim objSubNode      As clsXMLNode
5:     Dim colSubNodes     As Collection
6:     Dim strTemp         As String
7:     On Error GoTo Err

9:     Dim sUser As String, sDescription As String, sParameter As String, sCommand As String
10:    Dim bEnabled As Boolean
11:    Dim iIncrease As Integer, iIncreaseType As Variant
12:    Dim dDateTime As Date

14:    strTemp = G_APPPATH & "\Settings\Plan.xml"

16:    If g_objFileAccess.FileExists(strTemp) Then

18:        Set objXML = New clsXMLParser

20:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
21:        objXML.Parse
           
23:        Set colNodes = objXML.Nodes(1).Nodes
           'Just in case...
           'On Error Resume Next
26:        For Each objNode In colNodes

28:            If objNode.Name = "Plan" Then
               
30:                 Set colSubNodes = objNode.Nodes
                    
32:                 For Each objSubNode In colSubNodes
                         Select Case CStr(objSubNode.Name)
                                Case "User"
33:                                  sUser = CStr(objSubNode.Value)
                                Case "DateTime"
34:                                  dDateTime = CDate(objSubNode.Value)
                                Case "Enabled"
35:                                  bEnabled = CBool(objSubNode.Value)
                                Case "Command"
36:                                  sCommand = CStr(objSubNode.Value)
                                Case "Parameter"
37:                                  sParameter = CStr(objSubNode.Value)
                                Case "Increase"
38:                                  iIncrease = CInt(objSubNode.Value)
                                Case "IncreaseType"
39:                                  iIncreaseType = CVar(objSubNode.Value)
                                Case "Description"
40:                                  sDescription = CStr(objSubNode.Value)
41:                            End Select
42:                 Next

44:                 If sUser <> "" And sCommand <> "" Then
45:                     AddPlan sUser, dDateTime, bEnabled, sCommand, sParameter, iIncrease, CVar(iIncreaseType), sDescription
46:                 End If

48:                 sUser = "": sDescription = "": sParameter = ""
49:                 bEnabled = False
50:                 iIncrease = 0: iIncreaseType = 0 'none
                 
52:             End If

54:        Next
    
56:        On Error GoTo Err
    
58:        objXML.Clear

60:        Set objXML = Nothing
61:        Set objNode = Nothing
62:        Set colNodes = Nothing
63:        Set objSubNode = Nothing
64:        Set colSubNodes = Nothing

66:    End If

68:    Exit Sub

70:
Err:
71:    HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.PreInitialize()"
End Sub

Private Sub PreTerminate()
1:       Dim strTemp     As String
2:       Dim intFF       As Integer
3:       Dim intLoop     As Integer

5:       On Error GoTo Err

7:       strTemp = G_APPPATH & "\Settings\Plan.xml"
        
          'If the settings file exists, delete it
10:       If FileExists(strTemp) Then Kill strTemp
          
12:       If IsNothingPlan Then Exit Sub

14:       intFF = FreeFile
          'Append to Plan.xml
16:       Open strTemp For Append As intFF
        
18:        Print #intFF, "<PTDCH>"

           'Loop through
21:        For intLoop = LBound(Plan) To UBound(Plan)
22:            With Plan(intLoop)
23:                Print #intFF, vbTab & "<Plan>"
24:                Print #intFF, vbTab & vbTab & "<User>" & .User & "</User>"
25:                Print #intFF, vbTab & vbTab & "<DateTime>" & .DateTime & "</DateTime>"
26:                Print #intFF, vbTab & vbTab & "<Enabled>" & .Enabled & "</Enabled>"
27:                Print #intFF, vbTab & vbTab & "<Command>" & .Command & "</Command>"
28:                Print #intFF, vbTab & vbTab & "<Parameter>" & .Parameter & "</Parameter>"
29:                Print #intFF, vbTab & vbTab & "<Increase>" & .Increase & "</Increase>"
30:                Print #intFF, vbTab & vbTab & "<IncreaseType>" & .IncreaseType & "</IncreaseType>"
31:                Print #intFF, vbTab & vbTab & "<Description>" & .Description & "</Description>"
32:                Print #intFF, vbTab & "</Plan>"
33:            End With
34:        Next

36:        Print #intFF, "</PTDCH>"
        
38:       Close intFF

40:   Exit Sub

42:
Err:
43:   HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.PreTerminate"
End Sub

Private Sub m_Timer_Timer()
1:    Dim intLoop       As Integer
2:    Dim intLoopUser   As Integer
3:    Dim strUsers()    As String
4:    Dim strUser       As String
5:    Dim objUser       As clsUser
    
7:    On Error GoTo Err
    
9:    For intLoop = LBound(Plan) To UBound(Plan)
          'Check if is enabled the plan
11:        If Plan(intLoop).Enabled Then
               'Compare date and time
13:            If Plan(intLoop).DateTime < Now Then
                  'Check if is enabled the Scheduler
15:               If g_objSettings.EnabledScheduler Then
                     'Check if command is found in collection
17:                   If g_colCommands.Exists(Plan(intLoop).Command) Then
                         'Check if command is enabled
19:                      If g_colCommands.Item(Plan(intLoop).Command).Enabled Then
                            'Check if servering is offline
21:                         If G_SERVING Then
                               'Expand array
23:                             strUsers = Split(Plan(intLoop).User, ";")
                                'Check if user is online
25:                             For intLoopUser = LBound(strUsers) To UBound(strUsers)
26:                                 If g_colUsers.Online(strUsers(intLoopUser)) Then
27:                                    strUser = strUsers(intLoopUser)
28:                                    Exit For
29:                                 End If
30:                             Next
                                'Check strinf if is empty
32:                             If LenB(strUser) Then
                                   'Create user object
34:                                Set objUser = g_colUsers.ItemByName(CStr(strUser))
                                   'Run command
36:                                g_colCommands.Execute objUser, Plan(intLoop).Command & " " & Plan(intLoop).Parameter, True
                                   'Set status
38:                                Plan(intLoop).Status = "Command executed with success at: " & Now
39:                             Else
40:                                Plan(intLoop).Status = "Command not executed because user(s) (" & Plan(intLoop).User & ") is offline at: " & Now
41:                             End If
42:                         Else
43:                             Plan(intLoop).Status = "Command not executed because servering is offline at: " & Now
44:                         End If
45:                     Else
46:                         Plan(intLoop).Status = "Command not executed because is desabled"
47:                     End If
48:                 Else
49:                     Plan(intLoop).Status = "Command not found in collection"
50:                 End If
51:             Else
52:                 Plan(intLoop).Status = "Command not executed because Scheduler is desabled at: " & Now
53:             End If
                'Increase or desable plan
55:             If Not Plan(intLoop).Increase = 0 Then
                   Select Case Plan(intLoop).IncreaseType
                      Case iMinute
56:                       Plan(intLoop).DateTime = DateAdd("n", Plan(intLoop).Increase, Plan(intLoop).DateTime + Plan(intLoop).Increase)
                       Case iHour
57:                       Plan(intLoop).DateTime = DateAdd("h", Plan(intLoop).Increase, Plan(intLoop).DateTime + Plan(intLoop).Increase)
                       Case iDay
58:                       Plan(intLoop).DateTime = DateAdd("d", Plan(intLoop).Increase, Plan(intLoop).DateTime + Plan(intLoop).Increase)
                       Case iMonth
59:                       Plan(intLoop).DateTime = DateAdd("m", Plan(intLoop).Increase, Plan(intLoop).DateTime + Plan(intLoop).Increase)
60:                End Select
61:             Else
62:                Plan(intLoop).Enabled = False
63:                Plan(intLoop).IncreaseType = iNone
64:             End If
               'Refresh listview
66:             Call RefreshLvw
67:         End If
68:      End If
69:   Next

71:   Exit Sub
72:
Err:
73:   m_Timer.Enabled = False
74:   HandleError Err.Number, Err.Description, Erl & "|" & "clsPlan.m_Timer_Timer"
End Sub
Private Function FileExists(sFileName As String) As Boolean
    On Error GoTo Err:
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(sFileName) And vbDirectory) = 0
Err:
    ' if an error occurs, this function returns False
End Function
'------------------------------------------------------------------------------
'End Privates
'------------------------------------------------------------------------------
