VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type typChatBot
    Name            As String
    MinClass        As Integer
    Enabled         As Boolean
    Operator        As Boolean
    Share           As Double
    Description     As String
    Connection      As String
    Email           As String
    Tag             As String
    Icon            As Long
    MyINFO_Hello    As String
    MyINFO_NoHello  As String
End Type

Private ChatRomBot()         As typChatBot
Private TempChatRomBot()     As typChatBot

Private m_lvwChatRom         As ListItems
Private m_lvwItem            As ListItem
Private m_objLoopUser        As clsUser

Private Sub Class_Initialize()
1:     Set m_lvwChatRom = frmHub.lvwChatRom.ListItems
2:     PreInitialize
End Sub

Private Sub Class_Terminate()
1:     PreTerminate
2:     Set m_lvwChatRom = Nothing
3:     Erase ChatRomBot
End Sub

'------------------------------------------------------------------------------
'Publics
'------------------------------------------------------------------------------
Public Sub ShowAddDialog(ByRef frmOwner As Form)
1:    Dim frm As New frmChatRoom
2:    On Error GoTo Err
3:    With frm
4:        .Caption = "Add Chat Rom"
5:        .Show vbModal, frmOwner
6:    End With
7:    Exit Sub
8:
Err:
9:    HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.ShowAddDialog()"
End Sub

Public Sub ShowEditDialog(ByRef sBotName As String, ByRef frmOwner As Form)
1:  Dim frm As New frmChatRoom
2:  Dim lngTemp As Long

4:  On Error GoTo Err
    
6:  lngTemp = IsRegisteredBotName(sBotName)
   
8:    If lngTemp <> -1 Then
9:        With frm
10:          .Caption = "Edit Chat Rom: " & sBotName
        
12:          .txtBox(0).Text = ChatRomBot(lngTemp).Name
13:          .txtBox(0).Enabled = False
14:          .txtBox(1).Text = ChatRomBot(lngTemp).Share
15:          .txtBox(2).Text = ChatRomBot(lngTemp).Description
16:          .txtBox(3).Text = ChatRomBot(lngTemp).Connection
17:          .txtBox(4).Text = ChatRomBot(lngTemp).Email
18:          .txtBox(5).Text = ChatRomBot(lngTemp).Tag
19:          .txtBox(6).Text = ChatRomBot(lngTemp).Icon
            
21:          .cmbShare.Text = "Bytes"
            
             Select Case ChatRomBot(lngTemp).MinClass
                 Case 2: .cmbClass.Text = "2 = Mentored"
                 Case 3: .cmbClass.Text = "3 = Registered"
                 Case 4: .cmbClass.Text = "4 = Invisible"
                 Case 5: .cmbClass.Text = "5 = VIP"
                 Case 6: .cmbClass.Text = "6 = Operator"
                 Case 7: .cmbClass.Text = "7 = Invisible Operator"
                 Case 8: .cmbClass.Text = "8 = Super Operator"
                 Case 9: .cmbClass.Text = "9 = Invisible Super Operator"
                 Case 10: .cmbClass.Text = "10 = Admin"
                 Case 11: .cmbClass.Text = "11 = Invisible Admin"
23:          End Select
                
25:          If ChatRomBot(lngTemp).Operator Then _
                .chkBox(0).Value = 1 _
             Else .chkBox(0).Value = 0
28:          If ChatRomBot(lngTemp).Enabled Then _
                .chkBox(1).Value = 1 _
             Else .chkBox(1).Value = 0

32:         .Show vbModal, frmOwner
33:       End With
34:    End If

36:   Exit Sub

38:
Err:
39:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.ShowEditDialog(""" & sBotName & """)"
End Sub

Public Function AddChat(ByRef sName As String, _
                         ByRef bEnabled As Boolean, _
                         ByRef iMinClass As Integer, _
                         ByRef bOperator As Boolean, _
                Optional ByRef dShare As Double = 0, _
                Optional ByRef sDescription As String = "", _
                Optional ByRef sConnection As String = "", _
                Optional ByRef sEmail As String = "", _
                Optional ByRef sTag As String = "", _
                Optional ByRef lIcon As Long = 15, _
                Optional ByRef bOverwrite As Boolean = True) As Boolean

12:    Dim lngIndex          As Long
13:    Dim strMyINFO         As String
14:    Dim strOne            As String
15:    Dim strTwo            As String
      
17:    On Error GoTo Err
    
      'Check if is nothing.. If rem all bots.. the var is nothing
20:   If Not IsNothingChatRoom Then
            'Check if it has already been registered
22:         lngIndex = IsRegisteredBotName(sName)
23:         If lngIndex = -1 Then
24:             ReDim Preserve ChatRomBot(LBound(ChatRomBot) To UBound(ChatRomBot) + 1) As typChatBot
25:             lngIndex = UBound(ChatRomBot)
26:         Else
27:             If Not bOverwrite Then Exit Function
                'It is verified before was active, and now is desabled rem bot from user list..
29:             If ChatRomBot(lngIndex).Enabled Then
30:                 If bEnabled = False Then
31:                    Call RemBotFromUserList(sName)
32:                 End If
33:             End If
34:         End If
35:   Else
36:         ReDim ChatRomBot(0) As typChatBot
37:         lngIndex = 0
38:   End If

40:   With ChatRomBot(lngIndex)
41:        .Name = sName
42:        .Enabled = bEnabled
43:        .MinClass = iMinClass
44:        .Operator = bOperator
45:        .Share = dShare
46:        .Description = sDescription
47:        .Connection = sConnection
48:        .Email = sEmail
49:        .Tag = sTag
50:        .Icon = lIcon

           'Check if found parammetres for tag
53:        If sTag <> "" Then
54:            If Left(sTag, 1) <> "<" Then sTag = "<" & sTag
55:            If Right(sTag, 1) <> "<" Then sTag = sTag & ">"
56:        End If
           
           'Generate MyINFOS --> $MyINFO $ALL fLaSh Owner PTDCH and Creator<++ V:0.674,M:P,H:1/1/3,S:4>$ $DSL$$0$
59:        strMyINFO = "$MyINFO $ALL " & sName & " " & sDescription & sTag & "$ $" & sConnection & ChrW$(lIcon) & "$" & sEmail & "$" & dShare & "$|"

61:        If bOperator Then _
                strOne = strMyINFO & "$OpList " & sName & "$$|" _
           Else strOne = strMyINFO
           '
65:        strTwo = "$Hello " & sName & "|" & strOne
           '
67:        .MyINFO_Hello = strOne
68:        .MyINFO_NoHello = strTwo
69:   End With
    
      'Refresh ListView
72:   Call RefreshLvw

      'Send bot name to user list ..
75:   If ChatRomBot(lngIndex).Enabled Then
76:       Call SendBotToUserList(sName)
77:   End If

79:   Call RegisterBotByName(sName)

81:   AddChat = True
     
83:   Exit Function

85:   AddChat = False
86:
Err:
87:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.AddChat(""" & sName & """)"
End Function

Public Function RemChat(ByRef sName As String) As Boolean
1:   Dim intIndex As Integer
2:   Dim lngIndex As Long
3:   Dim intAux As Integer
4:   On Error GoTo Err
    
      'Check if it is registered
7:    lngIndex = IsRegisteredBotName(sName)

9:    If lngIndex = -1 Then
10:        RemChat = False
11:    Else
12:        RemChat = True
           'Rem User name from user list
14:        Call RemBotFromUserList(sName)
           'Rem bot register
16:        Call UnRegisterBotByName(sName)
           'It is only verified 1 item exists
18:        If UBound(ChatRomBot) = 0 Then
19:             Erase ChatRomBot
                'Add opchat by defaut..
21:             AddChat "OpChat", False, 6, True, 0, "Chatroom for operators"
22:             Call RefreshLvw
23:             Exit Function
24:        End If

           'Clear temp array.. Note: is 'UBound(TempChatRomBot) - 1' -->because 1 has deleted
27:        Erase TempChatRomBot()
28:        ReDim TempChatRomBot(LBound(ChatRomBot) To UBound(ChatRomBot) - 1) As typChatBot
        
           'Copy array info to temp array and exluind the bot deleted
31:        For intIndex = LBound(ChatRomBot) To UBound(ChatRomBot)
32:            If ChatRomBot(intIndex).Name <> sName Then
33:               With ChatRomBot(intIndex)
34:                    TempChatRomBot(intAux).Name = .Name
35:                    TempChatRomBot(intAux).Enabled = .Enabled
36:                    TempChatRomBot(intAux).MinClass = .MinClass
37:                    TempChatRomBot(intAux).Operator = .Operator
38:                    TempChatRomBot(intAux).Share = .Share
39:                    TempChatRomBot(intAux).Description = .Description
40:                    TempChatRomBot(intAux).Email = .Email
41:                    TempChatRomBot(intAux).Tag = .Tag
42:                    TempChatRomBot(intAux).Icon = .Icon
43:                    TempChatRomBot(intAux).MyINFO_Hello = .MyINFO_Hello
44:                    TempChatRomBot(intAux).MyINFO_NoHello = .MyINFO_NoHello
45:               End With
46:               intAux = intAux + 1
47:            End If
48:        Next
        
           'Clear array...
51:        Erase ChatRomBot()
52:        ReDim ChatRomBot(LBound(TempChatRomBot) To UBound(TempChatRomBot)) As typChatBot
        
           'Copy array info from de temp array
55:        For intIndex = LBound(TempChatRomBot) To UBound(TempChatRomBot)
56:            With TempChatRomBot(intIndex)
57:                 ChatRomBot(intIndex).Name = .Name
58:                 ChatRomBot(intIndex).Enabled = .Enabled
59:                 ChatRomBot(intIndex).MinClass = .MinClass
60:                 ChatRomBot(intIndex).Operator = .Operator
61:                 ChatRomBot(intIndex).Share = .Share
62:                 ChatRomBot(intIndex).Description = .Description
63:                 ChatRomBot(intIndex).Email = .Email
64:                 ChatRomBot(intIndex).Tag = .Tag
65:                 ChatRomBot(intIndex).Icon = .Icon
66:                 ChatRomBot(intIndex).MyINFO_Hello = .MyINFO_Hello
67:                 ChatRomBot(intIndex).MyINFO_NoHello = .MyINFO_NoHello
68:            End With
69:        Next
       
           'Clear temp array
72:        Erase TempChatRomBot

           'Refresh ListView
75:        Call RefreshLvw
   
77:        RemChat = True
78:    End If

80:   Exit Function

82:
Err:
83:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.RemChat(""" & sName & """)"
End Function

Public Function SwitchChat(ByRef sName As String, _
                           ByRef blnState As Boolean) As Boolean
2:    Dim lngIndex As Long
3:    On Error GoTo Err
        
      'Check if it is registered
6:    lngIndex = IsRegisteredBotName(sName)

8:    If lngIndex = -1 Then
9:        SwitchChat = False
10:   Else
11:       If blnState Then
12:           If Not ChatRomBot(lngIndex).Enabled Then
13:               ChatRomBot(lngIndex).Enabled = True
14:               Call SendBotToUserList(ChatRomBot(lngIndex).Name)
15:           End If
16:       Else
17:           If ChatRomBot(lngIndex).Enabled Then
18:               ChatRomBot(lngIndex).Enabled = False
19:               Call RemBotFromUserList(ChatRomBot(lngIndex).Name)
20:           End If
21:       End If
22:       SwitchChat = True
23:       Call RefreshLvw
24:   End If

26:   Exit Function
27:
Err:
29:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.SwitchChat(" & sName & ")"
End Function

Public Function ProcessChat(ByRef strKey As String) As Long
1:    Dim lngLoop     As Long
2:    On Error GoTo Err

4:    ProcessChat = -1

6:    For lngLoop = LBound(ChatRomBot) To UBound(ChatRomBot)
7:          If ChatRomBot(lngLoop).Enabled Then
8:              If ChatRomBot(lngLoop).Name = strKey Then
9:                  ProcessChat = ChatRomBot(lngLoop).MinClass
10:                 Exit Function
11:             End If
12:         End If
13:   Next

15:   Exit Function
16:
Err:
17:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.ProcessChat(" & strKey & ")"
End Function

'This Sub is used by frmHub.UpDateBots..
Public Sub UpDate(ByRef curUser As clsUser)
1:    Dim lngLoop     As Long
2:    On Error GoTo Err

      'Loop through and send MyINFO strings
5:    For lngLoop = LBound(ChatRomBot) To UBound(ChatRomBot)
           'Check if is enabled
7:         If ChatRomBot(lngLoop).Enabled Then
                'Check min class to use this chat rom
9:             If curUser.Class >= ChatRomBot(lngLoop).MinClass Then
                    'Check protocol type
11:                  If curUser.NoHello Then _
                          curUser.SendData ChatRomBot(lngLoop).MyINFO_NoHello _
                     Else curUser.SendData ChatRomBot(lngLoop).MyINFO_Hello
14:             End If
15:        End If
16:   Next

18:   Exit Sub
19:
Err:
20:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.UpDate()"
End Sub

'This Sub is used at starting servering..
Public Sub RegisterBots()
1:    Dim lngLoop     As Long
2:    On Error GoTo Err

      'Loop through and send MyINFO strings
5:    For lngLoop = LBound(ChatRomBot) To UBound(ChatRomBot)
           'Check if is enabled
7:         If ChatRomBot(lngLoop).Enabled Then
8:             g_objRegistered.Add ChatRomBot(lngLoop).Name, "Auto ChatRom bot Name locking system", Locked, "PTDCH / Core", , True
9:             Exit Sub
10:        End If
11:   Next

13:   Exit Sub
14:
Err:
15:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.RegisterBots()"
End Sub

'This Sub is used at stoping servering..
Public Sub UnRegisterChat()
1:    Dim lngLoop     As Long
2:    On Error GoTo Err

4:    For lngLoop = LBound(ChatRomBot) To UBound(ChatRomBot)
           'Check if is enabled
6:         If ChatRomBot(lngLoop).Enabled Then
7:             g_objRegistered.Remove ChatRomBot(lngLoop).Name
8:             Exit Sub
9:         End If
10:   Next

12:   Exit Sub
13:
Err:
15:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.UnRegisterChat()"
End Sub
'------------------------------------------------------------------------------
'End Publics
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'Privates
'------------------------------------------------------------------------------
Private Sub RegisterBotByName(ByRef sName As String)
1:    Dim lngLoop As Long
2:    On Error GoTo Err

4:    If G_SERVING Then
5:        For lngLoop = LBound(ChatRomBot) To UBound(ChatRomBot)
               'Check bot name
7:             If ChatRomBot(lngLoop).Name = sName Then
8:                  g_objRegistered.Add ChatRomBot(lngLoop).Name, "Auto ChatRom bot Name locking system", Locked, "PTDCH / Core", , True
9:                  Exit Sub
10:            End If
11:       Next
12:   End If

14:   Exit Sub
15:
Err:
'16:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.RegisterBotByName()"
End Sub

Private Sub UnRegisterBotByName(ByRef sName As String)
1:    Dim lngLoop As Long
2:    On Error GoTo Err

4:    g_objRegistered.Remove sName
      
6:    Exit Sub
7:
Err:
8:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.UnRegisterBotByName()"
End Sub

Private Sub RefreshLvw()
1:     Dim lngLoop     As Long
2:     On Error GoTo Err

4:     m_lvwChatRom.Clear

6:     For lngLoop = LBound(ChatRomBot) To UBound(ChatRomBot)
         'Add listitem
8:        With ChatRomBot(lngLoop)
9:             Set m_lvwItem = m_lvwChatRom.Add(, lngLoop & "s", .Name)
10:            m_lvwItem.SubItems(1) = .Enabled

               Select Case .MinClass
                    Case 2: m_lvwItem.SubItems(2) = "2 = Mentored"
                    Case 3: m_lvwItem.SubItems(2) = "3 = Registered"
                    Case 4: m_lvwItem.SubItems(2) = "4 = Invisible"
                    Case 5: m_lvwItem.SubItems(2) = "5 = VIP"
                    Case 6: m_lvwItem.SubItems(2) = "6 = Operator"
                    Case 7: m_lvwItem.SubItems(2) = "7 = Invisible Operator"
                    Case 8: m_lvwItem.SubItems(2) = "8 = Super Operator"
                    Case 9: m_lvwItem.SubItems(2) = "9 = Invisible Super Operator"
                    Case 10: m_lvwItem.SubItems(2) = "10 = Admin"
                    Case 11: m_lvwItem.SubItems(2) = "11 = Invisible Admin"
12:            End Select
             
14:            m_lvwItem.SubItems(3) = .Operator
15:            m_lvwItem.SubItems(4) = g_objFunctions.ShareSize(CDbl(.Share))
16:            m_lvwItem.SubItems(5) = .Description
17:            m_lvwItem.SubItems(6) = .Connection
18:            m_lvwItem.SubItems(7) = .Email
19:            m_lvwItem.SubItems(8) = .Tag
20:            m_lvwItem.SubItems(9) = .Icon
21:       End With
22:    Next
        
24:    Set m_lvwItem = Nothing

26:   Exit Sub
27:
Err:
28:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.RefreshLvw()"
End Sub

Private Function IsRegisteredBotName(ByRef strName As String) As Long
1:    Dim lngLoop As Long
    
3:    On Error GoTo Err
    
    'Make any needed character replacements in the nickname
6:    If InStrB(1, strName, " ") Then strName = Replace(strName, " ", "_")
7:    If InStrB(1, strName, "$") Then strName = Replace(strName, "$", "_")
8:    If InStrB(1, strName, "|") Then strName = Replace(strName, "|", "_")
    
    'Set to -1, meaning it hasn't found the bot Name
11:    IsRegisteredBotName = -1

13:    If IsNothingChatRoom Then Exit Function

        'Loop through and see if the Name matches any; if it does, return array index
16:    For lngLoop = LBound(ChatRomBot) To UBound(ChatRomBot)
17:          If ChatRomBot(lngLoop).Name = strName Then IsRegisteredBotName = lngLoop: Exit For
18:    Next

20:    Exit Function

22:
Err:
23:    HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.IsRegisteredBotName(""" & strName & """)"
End Function

Private Sub PreInitialize()
1:     Dim objXML          As clsXMLParser
2:     Dim objNode         As clsXMLNode
3:     Dim colNodes        As Collection
4:     Dim objSubNode      As clsXMLNode
5:     Dim colSubNodes     As Collection
6:     Dim strTemp         As String
7:     Dim bAddedBot       As Boolean
 
9:     Dim sName As String, sDescription As String, sConnection As String, sEmail As String, sTag As String
10:    Dim bEnabled As Boolean, bOperator As Boolean
11:    Dim lMinClass As Long, lIcon As Long
12:    Dim dShare As Double

14:    strTemp = G_APPPATH & "\Settings\ChatRom.xml"

16:    If g_objFileAccess.FileExists(strTemp) Then

18:        Set objXML = New clsXMLParser
19:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
20:        objXML.Parse
           
22:        Set colNodes = objXML.Nodes(1).Nodes
           'Just in case...
           'On Error Resume Next
25:        For Each objNode In colNodes

27:            If objNode.Name = "ChatRoom" Then
               
29:                 Set colSubNodes = objNode.Nodes
                    
31:                 For Each objSubNode In colSubNodes
                         Select Case objSubNode.Name
                                Case "Name"
32:                                  sName = CStr(objSubNode.Value)
                                Case "Enabled"
33:                                  bEnabled = CBool(objSubNode.Value)
                                Case "MinClass"
34:                                  lMinClass = CLng(objSubNode.Value)
                                Case "Operator"
35:                                  bOperator = CBool(objSubNode.Value)
                                Case "Share"
36:                                  dShare = CDbl(objSubNode.Value)
                                Case "Description"
37:                                  sDescription = CStr(objSubNode.Value)
                                Case "Connection"
38:                                  sConnection = CStr(objSubNode.Value)
                                Case "Email"
39:                                  sEmail = CStr(objSubNode.Value)
                                Case "Tag"
40:                                  sTag = CStr(objSubNode.Value)
                                Case "Icon"
41:                                  lIcon = CLng(objSubNode.Value)
42:                            End Select
43:                 Next
44:                 If LenB(sName) Then
45:                    AddChat sName, bEnabled, CInt(lMinClass), CBool(bOperator), dShare, sDescription, sConnection, sEmail, sTag, lIcon
46:                    bAddedBot = True
47:                 End If
48:                 sName = "": sDescription = "": sConnection = "": sEmail = "": sTag = "": lIcon = 0
49:                 bEnabled = False: bOperator = False
50:                 lMinClass = 0: dShare = 0
51:             End If
52:        Next
    
54:        On Error GoTo Err
    
56:        objXML.Clear

58:        Set objXML = Nothing
59:        Set objNode = Nothing
60:        Set colNodes = Nothing
61:        Set objSubNode = Nothing
62:        Set colSubNodes = Nothing

64:    Else
65:        bAddedBot = False
66:    End If

       'If no bot found add opchat by defaut..
69:    If Not bAddedBot Then AddChat "OpChat", True, 6, True, 0, "Chatroom for operators"
       
71:   Exit Sub

73:
Err:
74:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.AddChat()"
End Sub

Private Sub PreTerminate()
1:        On Error GoTo Err

3:        Dim strTemp     As String
4:        Dim intFF       As Integer
5:        Dim intLoop     As Integer

7:        strTemp = G_APPPATH & "\Settings\ChatRom.xml"
        
          'If the settings file exists, delete it
10:       If FileExists(strTemp) Then
11:             Kill strTemp
12:       End If

13:       intFF = FreeFile
        
          'Append to ChatRom.xml
16:       Open strTemp For Append As intFF
        
18:        Print #intFF, "<PTDCH>"
        
           'Loop through and send MyINFO strings
21:        For intLoop = LBound(ChatRomBot) To UBound(ChatRomBot)
22:            With ChatRomBot(CInt(intLoop))
23:                Print #intFF, vbTab & "<ChatRoom>"
24:                Print #intFF, vbTab & vbTab & "<Name>" & .Name & "</Name>"
25:                Print #intFF, vbTab & vbTab & "<Enabled>" & .Enabled & "</Enabled>"
26:                Print #intFF, vbTab & vbTab & "<MinClass>" & .MinClass & "</MinClass>"
27:                Print #intFF, vbTab & vbTab & "<Operator>" & .Operator & "</Operator>"
28:                Print #intFF, vbTab & vbTab & "<Share>" & .Share & "</Share>"
29:                Print #intFF, vbTab & vbTab & "<Description>" & .Description & "</Description>"
30:                Print #intFF, vbTab & vbTab & "<Connection>" & .Connection & "</Connection>"
31:                Print #intFF, vbTab & vbTab & "<Tag>" & .Tag & "</Tag>"
32:                Print #intFF, vbTab & vbTab & "<Icon>" & .Icon & "</Icon>"
33:                Print #intFF, vbTab & vbTab & "<Email>" & .Email & "</Email>"
34:                Print #intFF, vbTab & "</ChatRoom>"
35:            End With
36:        Next

38:        Print #intFF, "</PTDCH>"
        
40:       Close intFF

42:   Exit Sub

44:
Err:
45:   HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.PreTerminate"
End Sub

'Check if the var type is nothing
Private Function IsNothingChatRoom() As Boolean
1:    Dim strTest As String
    
3:    On Error GoTo Err
    
5:    strTest = ChatRomBot(0).Name
6:    IsNothingChatRoom = False
    
8:    Exit Function
9:
Err:
10:    Err.Clear
11:    IsNothingChatRoom = True
End Function

'Send a bot name to the lists
Private Sub SendBotToUserList(ByRef strName As String)
1:   Dim intIndex       As Integer
2:   Dim lngLoop        As Long
3:   Dim lngIndex       As Long
4:   Dim objLoopUser    As clsUser
5:   On Error GoTo Err
    
7:   If G_SERVING = False Then Exit Sub
    
    'Check if it is registered
10:    lngIndex = IsRegisteredBotName(strName)

12:    If lngIndex = -1 Then
13:        Exit Sub
14:    Else
15:        For lngLoop = LBound(ChatRomBot) To UBound(ChatRomBot)
16:            If ChatRomBot(lngLoop).Name = strName Then
                'Check if is enabled the chat room
18:                If ChatRomBot(lngLoop).Enabled Then
                    'Send to MyINFO to users
20:                    For Each m_objLoopUser In g_colUsers
21:                        If m_objLoopUser.Class >= ChatRomBot(lngLoop).MinClass Then
22:                            If m_objLoopUser.NoHello Then _
                                 m_objLoopUser.SendData ChatRomBot(lngLoop).MyINFO_NoHello _
                            Else m_objLoopUser.SendData ChatRomBot(lngLoop).MyINFO_Hello
25:                        End If
26:                    Next
27:                End If
28:            End If
29:        Next
30:        Set m_objLoopUser = Nothing
31:    End If
    
33:    Exit Sub
34:
Err:

36:    HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.SendBotName()"
End Sub

'Send a bot name to the lists
Private Sub RemBotFromUserList(ByRef strName As String)
1:    Dim lngIndex     As Long

3:    On Error GoTo Err
    
5:    If G_SERVING = False Then Exit Sub
    
7:    If lngIndex = -1 Then
8:       Exit Sub
9:    Else
10:      g_colUsers.SendToAll "$Quit " & strName & "|"
11:   End If
    
13:   Exit Sub
14:
Err:
16:
17:    HandleError Err.Number, Err.Description, Erl & "|" & "clsChatRoom.SendBotName()"
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
