VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

#If Status Then
    'Protocol '*************************
    Public NickList             As Long
    Public ConnectMe            As Long
    Public RevConnectToMe       As Long
    Public Kicks                As Long
    Public Redirects            As Long
    Public BotINFO              As Long
    Public NetInfo              As Long
    Public Errors               As Long
    Public ErrSockets           As Long
    Public Requests             As Long
    '***********************************
    
    'Trafic ****************************
    Public MsgMainChat          As Long
    Public MsgPM                As Long
    Public Searchs              As Long
    Public MyINFOs              As Long
    Public TotalSend            As Long
    Public TotalRecived         As Long
    '
    Private SpeedSend           As Long
    Private SpeedRecived        As Long
    '***********************************
    
    Public Enum enSatistics
        iProtocol = 0
        iTrafic = 1
    End Enum
    
    Public Enum enProtocol
        iNickList = 0
        iConnectMe = 1
        iRevConnectToMe = 2
        iKicks = 3
        iRedirects = 4
        iBotINFO = 5
        iNetINFO = 6
        iErrors = 7
        iErrSockets = 8
        iRequests = 9
    End Enum
    
    Public Enum enTrafic
        iMsgMainChat = 0
        iMsgPM = 1
        iSearchs = 2
        iMyINFOs = 3
        iTotalSend = 4
        iTotalRecived = 5
    End Enum
    
    Private WithEvents m_Timer           As Timer
Attribute m_Timer.VB_VarHelpID = -1
    Private m_frmParent                  As frmParent
    
    'Logs  *****************************************************************
    'Set the list box's horizontal scroll bars
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hWnd As Long, ByVal wMsg As Long, _
                ByVal wParam As Long, lParam As Any) As Long
    Private Const LB_SETHORIZONTALEXTENT = &H194
    
    #Const LIMIT_MSG = True
    
    #If LIMIT_MSG Then
        Private m_lngMsgCountMChat        As Long
        Private m_lngMsgCountPMLog        As Long
        Private m_lngMsgCountMiscLog      As Long
        Private Const c_iMLinesChat       As Integer = 350
        Private Const c_iMLinesPMLog      As Integer = 250
        Private Const c_iMLinesMiscLog    As Integer = 150
    #End If
    
    Private m_lvwUsers                    As ListItems
    Private m_lstMessages(2)              As ListBox
    Private m_lvwItem                     As ListItem
    Private m_frmHub                      As Form
    '*********************************************************************
    
    'Statistics ***********************************************************
    Private m_lblProtocol(9)              As Label
    Private m_lblTrafic(7)                As Label
    '*********************************************************************
    
    Private Sub Class_Initialize()
1:        Set m_lvwUsers = frmHub.lvwUsers.ListItems
2:        Set m_lstMessages(0) = frmHub.lstStatus(0)
3:        Set m_lstMessages(1) = frmHub.lstStatus(1)
4:        Set m_lstMessages(2) = frmHub.lstStatus(2)
5:        Set m_frmHub = frmHub
6:        Set m_lblProtocol(0) = frmHub.lblStatistics(0)
7:        Set m_lblProtocol(1) = frmHub.lblStatistics(1)
8:        Set m_lblProtocol(2) = frmHub.lblStatistics(2)
9:        Set m_lblProtocol(3) = frmHub.lblStatistics(3)
10:       Set m_lblProtocol(4) = frmHub.lblStatistics(4)
11:       Set m_lblProtocol(5) = frmHub.lblStatistics(5)
12:       Set m_lblProtocol(6) = frmHub.lblStatistics(6)
13:       Set m_lblProtocol(7) = frmHub.lblStatistics(7)
14:       Set m_lblProtocol(8) = frmHub.lblStatistics(8)
15:       Set m_lblProtocol(9) = frmHub.lblStatistics(9)
16:       Set m_lblTrafic(0) = frmHub.lblStatistics(20)
17:       Set m_lblTrafic(1) = frmHub.lblStatistics(21)
18:       Set m_lblTrafic(2) = frmHub.lblStatistics(22)
19:       Set m_lblTrafic(3) = frmHub.lblStatistics(23)
20:       Set m_lblTrafic(4) = frmHub.lblStatistics(24)
21:       Set m_lblTrafic(5) = frmHub.lblStatistics(25)
22:       Set m_lblTrafic(6) = frmHub.lblStatistics(26)
23:       Set m_lblTrafic(7) = frmHub.lblStatistics(27)
24:       IniStatistics
          'Create our dummy form to get it's Timer
26:       Set m_frmParent = New frmParent
27:       Set m_Timer = m_frmParent.tmrTimer
28:       m_Timer.Interval = 1000
29:       m_Timer.Enabled = True
    End Sub
    
    Private Sub Class_Terminate()
1:        Dim i As Integer
2:        m_Timer.Enabled = False
3:        Set m_Timer = Nothing
4:        Set m_frmParent = Nothing
5:        Set m_lvwUsers = Nothing
6:        Set m_lstMessages(0) = Nothing
7:        Set m_lstMessages(1) = Nothing
8:        Set m_lstMessages(2) = Nothing
9:        Set m_frmHub = Nothing
10:       For i = LBound(m_lblProtocol) To UBound(m_lblProtocol)
11:           Set m_lblProtocol(i) = Nothing
12:       Next
13:       For i = LBound(m_lblTrafic) To UBound(m_lblTrafic)
14:          Set m_lblTrafic(i) = Nothing
15:       Next
    End Sub
    
    Public Sub IniStatistics()
1:        On Error GoTo Err
2:        Dim i As Integer
        
          'Protocol
5:        For i = LBound(m_lblProtocol) To UBound(m_lblProtocol)
6:            m_lblProtocol(i).Caption = "0"
7:        Next
        
          'Trafic
10:       For i = LBound(m_lblTrafic) To UBound(m_lblTrafic) - 2
11:           m_lblTrafic(i).Caption = "0 Bytes"
12:       Next
13:       m_lblTrafic(6).Caption = "0 Kb/s"
14:       m_lblTrafic(7).Caption = "0 Kb/s"
        
16:       Exit Sub
17:
Err:
18:       HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.ClearStatistics()"
    End Sub
    
    Public Sub UpDateProtocol(ByRef typ As enProtocol)
1:        On Error GoTo Err
    
        Select Case typ
            Case 0: NickList = NickList + 1
            Case 1: ConnectMe = ConnectMe + 1
            Case 2: RevConnectToMe = RevConnectToMe + 1
            Case 3: Kicks = Kicks + 1
            Case 4: Redirects = Redirects + 1
            Case 5: BotINFO = BotINFO + 1
            Case 6: NetInfo = NetInfo + 1
            Case 7: Errors = Errors + 1
            Case 8: ErrSockets = ErrSockets + 1
            Case 9: Requests = Requests + 1
3:          End Select
          
5:       DoEvents
          
7:       RefreshLvw iProtocol

9:      Exit Sub
10:
Err:
11:      HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.UpDateProtocol(" & typ & ")"
    End Sub

    Public Sub UpDateTrafic(ByVal lngValue As Long, ByRef typ As enTrafic)
1:        On Error GoTo Err

        Select Case typ
            Case 0: MsgMainChat = MsgMainChat + lngValue
            Case 1: MsgPM = MsgPM + lngValue
            Case 2: Searchs = Searchs + lngValue
            Case 3: MyINFOs = MyINFOs + lngValue
            Case 4: TotalSend = TotalSend + lngValue: SpeedSend = lngValue
            Case 5: TotalRecived = TotalRecived + lngValue: SpeedRecived = lngValue
3:          End Select

5:       DoEvents

7:       RefreshLvw iTrafic

9:      Exit Sub
10:
Err:
11:      HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.UpDateTrafic(" & lngValue & ", " & typ & ")"
    End Sub
    
    Public Sub RefreshLvw(ByRef enu As enSatistics)
1:       On Error GoTo Err
    
3:        If Not EnabledStatistics Then Exit Sub
    
        Select Case enu
    
            '************************
            Case iProtocol
            '************************
            
9:            m_lblProtocol(0).Caption = NickList
10:            m_lblProtocol(1).Caption = ConnectMe
11:            m_lblProtocol(2).Caption = RevConnectToMe
12:            m_lblProtocol(3).Caption = Kicks
13:            m_lblProtocol(4).Caption = Redirects
14:            m_lblProtocol(5).Caption = BotINFO
15:            m_lblProtocol(6).Caption = NetInfo
16:            m_lblProtocol(7).Caption = Errors
17:            m_lblProtocol(8).Caption = ErrSockets
18:            m_lblProtocol(9).Caption = Requests

            '************************
            Case iTrafic
            '************************
            
23:            m_lblTrafic(0).Caption = ConverteSize(CDbl(MsgMainChat))
24:            m_lblTrafic(1).Caption = ConverteSize(CDbl(MsgPM))
25:            m_lblTrafic(2).Caption = ConverteSize(CDbl(Searchs))
26:            m_lblTrafic(3).Caption = ConverteSize(CDbl(MyINFOs))
27:            m_lblTrafic(4).Caption = ConverteSize(CDbl(TotalSend))
28:            m_lblTrafic(5).Caption = ConverteSize(CDbl(TotalRecived))

30:         End Select

32:      Exit Sub
33:
Err:
34:      HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.RefreshLvw(" & enu & ")"
    End Sub

    Private Function GetSpeed(ByVal dblBytes As Double) As String
1:     GetSpeed = FormatNumber(dblBytes / 1024, 3) & " Kb/s"
    End Function
    
    'Create a nice balance ;-)
    Private Sub m_Timer_Timer()
1:    On Error GoTo Err

3:      If SpeedSend > 0 Then
4:         m_lblTrafic(6).Caption = GetSpeed(CDbl(SpeedSend))
5:         If SpeedSend <= 10 Then m_lblTrafic(6).Caption = "0,000 Kbs"
6:      End If
        
8:      If SpeedSend > 0 Then
9:         m_lblTrafic(7).Caption = GetSpeed(CDbl(SpeedRecived))
10:        If SpeedRecived <= 10 Then m_lblTrafic(7).Caption = "0,000 Kbs"
11:     End If

13:     DoEvents

15:     If SpeedSend > 0 Then SpeedSend = SpeedSend - 10 'Bytes
16:     If SpeedRecived > 0 Then SpeedRecived = SpeedSend - 10 'Bytes

18:     If SpeedSend < 0 Then SpeedSend = 0
19:     If SpeedRecived < 0 Then SpeedRecived = 0

21:     Exit Sub
22:
Err:
23:     HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.m_Timer_Timer()"
    End Sub

    'Determines the closest size in the approriate *B
    Private Function ConverteSize(ByVal dblBytes As Double) As String
1:       On Error GoTo Err
         
         Select Case dblBytes
            Case Is > 1048575
3:                ConverteSize = FormatNumber(dblBytes / 1048576, 2) & " MiB"
            Case Is > 1023
4:                ConverteSize = FormatNumber(dblBytes / 1024, 2) & " KiB"
            Case Else
5:                ConverteSize = dblBytes & " Bytes"
6:      End Select
        
8:      Exit Function
        
10:
Err:
11:      HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.ShareSize(" & dblBytes & ")"
    End Function
    
    'Adds a user to the listview
    Public Sub UAdd(ByRef objUser As clsUser)
1:        Dim strTemp       As String
          
3:        On Error GoTo Err
        
5:        strTemp = CStr(objUser.iWinsockIndex)
          
          'Add listitem with unknown icon
8:        Set m_lvwItem = m_lvwUsers.Add(, strTemp & "i", "(No name)")
        
          'Add extra info
11:        m_lvwItem.SubItems(1) = CStr(objUser.IP)
12:        m_lvwItem.SubItems(2) = strTemp
13:        m_lvwItem.SubItems(3) = CDate(objUser.ConnectedSince)
      
15:        Set m_lvwItem = Nothing
        
17:        Exit Sub
        
19:
Err:
20:         HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.UAdd()"
    End Sub
    
    'Updates a users info
    Public Sub UUpdate(ByRef intIndex As Integer, ByRef strName As String)
1:      On Error GoTo Err
    
        'Must remember to add more stuff
        'Suggestions appreciated (ie add icons)
    
        'Update name
7:      m_lvwUsers(intIndex & "i").Text = strName

9:      Exit Sub

11:
Err:
12:     HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.UUpdate()"
    End Sub
    
    'Removes a user from the listview
    Public Sub URemove(ByRef intIndex As Integer)
1:        On Error GoTo Err

3:        m_lvwUsers.Remove intIndex & "i"

5:       Exit Sub
6:
Err:
7:      HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.UUpdate()"
    End Sub
    
    'Adds a message to the message listbox
    Public Sub MAdd(ByRef strMessage As String) ', ByRef strIP As String, ByRef lngColour As Long)

2:       On Error Resume Next

       'PM Log
5:      If Left(strMessage, 4) = "$To:" Then
            #If LIMIT_MSG Then
7:             If m_lngMsgCountPMLog > c_iMLinesPMLog Then
8:                m_lstMessages(1).RemoveItem 0 'remove first line
9:                m_lngMsgCountPMLog = m_lstMessages(1).ListCount - 1 'set m_lngMsgCountPMLog count
10:             Else
11:                m_lngMsgCountPMLog = m_lngMsgCountPMLog + 1
12:             End If
            #End If
14:          m_lstMessages(1).AddItem ("[" & Now & "]" & strMessage)
             'set selected last line
16:          m_lstMessages(1).Selected(m_lstMessages(1).ListCount - 1) = True
17:          SetListboxHSbar 1
        
        'Main Chat Log
20:     ElseIf Left(strMessage, 1) <> "$" Then
            #If LIMIT_MSG Then
22:              If m_lngMsgCountMChat > c_iMLinesMiscLog Then
23:                 m_lstMessages(0).RemoveItem 0 'remove first line
24:                 m_lngMsgCountMChat = m_lstMessages(0).ListCount - 1 'set m_lngMsgCountMChat count
25:              Else
26:                 m_lngMsgCountMChat = m_lngMsgCountMChat + 1
27:              End If
            #End If
29:          m_lstMessages(0).AddItem (strMessage)
             'set selected last line
31:          m_lstMessages(0).Selected(m_lstMessages(0).ListCount - 1) = True
32:          SetListboxHSbar 0
       
        'Misc protocol log
35:     Else

37:          If EnabledMiscLog Then
                #If LIMIT_MSG Then
39:                  If m_lngMsgCountMiscLog > c_iMLinesMiscLog Then
40:                     m_lstMessages(2).RemoveItem 0 'remove first line
41:                     m_lngMsgCountMiscLog = m_lstMessages(2).ListCount - 1 'set m_lngMsgCountMiscLog count
42:                  Else
43:                     m_lngMsgCountMiscLog = m_lngMsgCountMiscLog + 1
44:                  End If
                #End If
46:             m_lstMessages(2).AddItem (strMessage)
                'set selected last line
48:             m_lstMessages(2).Selected(m_lstMessages(2).ListCount - 1) = True
49:             SetListboxHSbar 2
50:          End If
51:     End If
        
    End Sub
    
    'Clears all messages
    Public Sub MClear(iIndex As Integer)
1:    On Error GoTo Err
2:      m_lstMessages(iIndex).Clear
        #If LIMIT_MSG Then
            Select Case iIndex
               Case 0: m_lngMsgCountMChat = 0
               Case 1: m_lngMsgCountPMLog = 0
               Case 2: m_lngMsgCountMiscLog = 0
4:            End Select
        #End If
6:     SetListboxHSbar iIndex
7:     Exit Sub
8:
Err:
9:     HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.MClear(" & iIndex & ")"
    End Sub
    
    'Clears m_lvwUsers
    Public Sub UClear()
1:        On Error GoTo Err
        
3:        m_lvwUsers.Clear
4:        Exit Sub
5:
Err:
6:        HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.UClear()"
    End Sub
    
    Private Sub SetListboxHSbar(iIndex As Integer)
           ' Set the list box's horizontal extent so it
           ' can display its longest entry. This routine
           ' assumes the form is using the same font as
           ' the list box.
5:         On Error GoTo Err

7:         Dim i As Integer
8:         Dim new_len As Long
9:         Dim max_len As Long

11:        For i = 0 To 1 'm_lstMessages(iIndex).ListCount - 1
12:                 new_len = 10 + m_frmHub.ScaleX(m_frmHub.TextWidth(m_lstMessages(iIndex).List(i)), m_frmHub.ScaleMode, vbPixels)
13:                 If max_len < new_len Then max_len = new_len
14:                 DoEvents
15:        Next i

17:        SendMessage m_lstMessages(iIndex).hWnd, _
                    LB_SETHORIZONTALEXTENT, _
                         max_len, 0
        
21:        Exit Sub
22:
Err:
23:        HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.SetListboxHSbar(" & iIndex & ")"
    End Sub
    
    'This function will be for little memory..
    'It is not necessary to generate the log, if the listbox is not visible
    'this because this log use a lot of memory..
    Private Function EnabledMiscLog() As Boolean
1:        If m_frmHub.WindowState = vbMinimized Or _
             m_lstMessages(2).Visible = False Then _
               EnabledMiscLog = False _
          Else EnabledMiscLog = True
    End Function
    
    'This function will be for little memory..
    'It is not necessary to generate the info, if the picture box is not visible
    'this because this log use a lot of memory..
    Private Function EnabledStatistics() As Boolean
1:      If m_frmHub.WindowState = vbMinimized Or _
           m_frmHub.picStatus(5).Visible = False Then _
             EnabledStatistics = False _
        Else EnabledStatistics = True
    End Function
#End If
