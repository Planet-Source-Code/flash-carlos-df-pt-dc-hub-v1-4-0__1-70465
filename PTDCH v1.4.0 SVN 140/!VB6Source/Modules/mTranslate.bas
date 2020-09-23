Attribute VB_Name = "mTranslate"
Option Explicit
  
'INTERFACE LANGUAGE
Public Sub TranslateCtrlCaption(ByVal ctrlName As String, ByVal ctrlCaption As String)
1:    On Error Resume Next
2:    Dim i
3:    Dim strName As String
4:    Dim intIndex As Integer
5:    strName = g_objFunctions.BeforeFirst(ctrlName, "(")
6:    intIndex = g_objFunctions.BeforeLast(g_objFunctions.AfterFirst(ctrlName, "("), ")")
7:    For Each i In frmHub.Controls
8:         If i.Name = strName Then
9:              If i.Index = intIndex Then
10:                i.Caption = ctrlCaption
11:                Exit Sub
12:             End If
13:        End If
14:   Next
End Sub

Public Sub TranslateTabSCaption(ByVal ctrlName As String, ByVal ctrlCaption As String)
1:   On Error Resume Next

3:   Dim i As Integer
4:   Dim X As String

6:   With frmHub
7:      For i = 1 To .tbsMenu.Tabs.count
8:          X = "tbsMenu.Tabs(" & i & ")"
9:         If ctrlName = X Then _
         .tbsMenu.Tabs(i).Caption = ctrlCaption: Exit Sub
11:     Next i
12:     For i = 1 To .tbsSecurity.Tabs.count
13:         X = "tbsSecurity.Tabs(" & i & ")"
14:        If ctrlName = X Then _
         .tbsSecurity.Tabs(i).Caption = ctrlCaption: Exit Sub
16:     Next i
17:     For i = 1 To .tbsInteractions.Tabs.count
18:         X = "tbsInteractions.Tabs(" & i & ")"
19:        If ctrlName = X Then _
                .tbsInteractions.Tabs(i).Caption = ctrlCaption: Exit Sub
21:     Next i
22:     For i = 1 To .tabAdv.Tabs.count
23:         X = "tabAdv.Tabs(" & i & ")"
24:        If ctrlName = X Then _
                .tabAdv.Tabs(i).Caption = ctrlCaption: Exit Sub
26:     Next i
27:     For i = 1 To .tbsHelp.Tabs.count
28:         X = "tbsHelp.Tabs(" & i & ")"
29:        If ctrlName = X Then _
                .tbsHelp.Tabs(i).Caption = ctrlCaption: Exit Sub
31:     Next i
32:     For i = 1 To .tbsStatus.Tabs.count
33:         X = "tbsStatus.Tabs(" & i & ")"
34:        If ctrlName = X Then _
                .tbsStatus.Tabs(i).Caption = ctrlCaption: Exit Sub
36:     Next i
37:     For i = 1 To .tbsInfo.Tabs.count
38:         X = "tbsInfo.Tabs(" & i & ")"
39:        If ctrlName = X Then _
                .tbsInfo.Tabs(i).Caption = ctrlCaption: Exit Sub
41:     Next i
42:     For i = 1 To .tbsDbManager.Tabs.count
43:         X = "tbsDbManager.Tabs(" & i & ")"
44:        If ctrlName = X Then _
                .tbsDbManager.Tabs(i).Caption = ctrlCaption: Exit Sub
46:     Next i
47:     For i = 1 To .tbsDbManager.Tabs.count
48:         X = "tabBansIPs.Tabs(" & i & ")"
49:        If ctrlName = X Then _
                .tabBansIPs.Tabs(i).Caption = ctrlCaption: Exit Sub
51:     Next i
52:  End With
End Sub

Public Sub TranslateListViewCaption(ByVal ctrlName As String, ByVal ctrlText As String)
1:   On Error Resume Next

3:   Dim i As Integer
4:   Dim X As String

6:   With frmHub
7:      For i = 1 To .lvwCommands.ColumnHeaders.count
8:          X = "lvwCommands.ColumnHeaders(" & i & ")"
9:         If ctrlName = X Then _
         .lvwCommands.ColumnHeaders(i).Text = ctrlText: Exit Sub
11:     Next i
12:     For i = 1 To .lvwUsers.ColumnHeaders.count
13:         X = "lvwUsers.ColumnHeaders(" & i & ")"
14:        If ctrlName = X Then _
         .lvwUsers.ColumnHeaders(i).Text = ctrlText: Exit Sub
16:     Next i
17:     For i = 1 To .lvwScripts.ColumnHeaders.count
18:         X = "lvwScripts.ColumnHeaders(" & i & ")"
19:        If ctrlName = X Then _
                .lvwScripts.ColumnHeaders(i).Text = ctrlText: Exit Sub
21:     Next i
22:     For i = 1 To .lvwRegistered.ColumnHeaders.count
23:         X = "lvwRegistered.ColumnHeaders(" & i & ")"
24:        If ctrlName = X Then _
                .lvwRegistered.ColumnHeaders(i).Text = ctrlText: Exit Sub
26:     Next i
27:     For i = 1 To .lvwBans.ColumnHeaders.count
28:         X = "lvwBans.ColumnHeaders(" & i & ")"
29:        If ctrlName = X Then _
                .lvwBans.ColumnHeaders(i).Text = ctrlText: Exit Sub
31:     Next i
32:     For i = 1 To .lvwPlugins.ColumnHeaders.count
33:         X = "lvwPlugins.ColumnHeaders(" & i & ")"
34:        If ctrlName = X Then _
                .lvwPlugins.ColumnHeaders(i).Text = ctrlText: Exit Sub
36:     Next i
37:     For i = 1 To .lvwChatRom.ColumnHeaders.count
38:         X = "lvwChatRom.ColumnHeaders(" & i & ")"
39:        If ctrlName = X Then _
                .lvwChatRom.ColumnHeaders(i).Text = ctrlText: Exit Sub
41:     Next i
42:     For i = 1 To .lvwPlan.ColumnHeaders.count
43:         X = "lvwPlan.ColumnHeaders(" & i & ")"
44:        If ctrlName = X Then _
                .lvwPlan.ColumnHeaders(i).Text = ctrlText: Exit Sub
46:     Next i
47:     For i = 1 To .lvwPlan.ColumnHeaders.count
48:         X = "lvwTempIPBan.ColumnHeaders(" & i & ")"
49:        If ctrlName = X Then _
                .lvwTempIPBan.ColumnHeaders(i).Text = ctrlText: Exit Sub
51:     Next i

53:     For i = 1 To .lvwPlan.ColumnHeaders.count
54:         X = "lvwPermIPBan.ColumnHeaders(" & i & ")"
55:        If ctrlName = X Then _
                .lvwPermIPBan.ColumnHeaders(i).Text = ctrlText: Exit Sub
57:     Next i

59:     For i = 1 To .lvwLanguages.ColumnHeaders.count
60:         X = "lvwLanguages.ColumnHeaders(" & i & ")"
61:        If ctrlName = X Then _
                .lvwLanguages.ColumnHeaders(i).Text = ctrlText: Exit Sub
63:     Next i
64:  End With
End Sub
Public Sub TranslateTexts(ByVal ctrlName As String, ByVal ctrlText As String)
1: On Error Resume Next
2:    Dim i
3:    Dim strName As String
4:    Dim intIndex As Integer
5:    strName = g_objFunctions.BeforeFirst(ctrlName, "( ")
6:    intIndex = g_objFunctions.BeforeLast(g_objFunctions.AfterFirst(ctrlName, "("), ")")
7:    For Each i In frmHub.Controls
8:        If i.Name = strName Then
9:            If i.Index = intIndex Then
10:                i.Text = ctrlText
11:                Exit Sub
12:            End If
13:        End If
14:    Next
End Sub

Public Sub TranslateToolBar(ByVal ctrlName As String, ByVal ctrlToolTip As String)
1: On Error Resume Next
2:    Dim i As Integer
3:    Dim X As String
4:    With frmHub
5:        For i = 1 To .tlbScript.Buttons.count
6:            X = "tlbScript.Buttons.Item(" & i & ")"
7:            If ctrlName = X Then _
         .tlbScript.Buttons.Item(i).ToolTipText = ctrlToolTip: Exit Sub
9:        Next i
10:   End With
End Sub

Public Sub ClearTranslations()
1:    Dim CTL As Control
2:    Dim i As Integer

4:    On Error GoTo Err
 
     'Clear all captions and all tooltips..
     'This is ideal to verify mistakes in the translations ;-)
8:     For Each CTL In frmHub.Controls
          Select Case TypeName(CTL)
             Case "CommandButton", "OptionButton", "CheckBox", "Label"
9:                CTL.Caption = ""
10:               CTL.ToolTipText = ""
             Case "ListView"
11:              For i = 1 To CTL.ColumnHeaders.count
12:                     CTL.ColumnHeaders(i).Text = "": Next
             Case "TabStrip"
13:              For i = 1 To CTL.Tabs.count
14:                     If Not CTL.Name = "tbsScripts" Then CTL.Tabs(i).Caption = ""
15:              Next
             Case "ToolBar"
16:              For i = 1 To CTL.Buttons.count
17:                  CTL.Buttons.Item(i).ToolTipText = "": Next
             Case "TextBox"
                     CTL.ToolTipText = ""
           End Select
19:    Next

       'Adds the caption of the objects that with are not translated...
22:    With frmHub
23:        .cmdSkin(0).Caption = "<"
24:        .cmdSkin(1).Caption = ">"
25:        .LabelsURL(0).Caption = "E-mail"
26:        .LabelsURL(1).Caption = "Home Page"
27:        .lblStatus(0).Caption = "1    2    3    4    5    6"
28:        .lblStatus(1).Caption = "------------- Send Chat -------------"
29:        .lblStatus(2).Caption = "1 = Send Chat To All"
30:        .lblStatus(3).Caption = "2 = Send Chat To Op"
31:        .lblStatus(4).Caption = "3 = Send Chat To UnRegistered"
32:        .lblStatus(5).Caption = "4 = Send PM To All"
33:        .lblStatus(6).Caption = "5 = Send PM To Op"
34:        .lblStatus(7).Caption = "6 = Send PM To UnRegistered"
35:        .LblShadowed(0).Caption = "PT Direct Connect Hub"
36:        .LblShadowed(1).Caption = "PT Direct Connect Hub"
37:        .LblShadowed(2).Caption = "Produced bY fLaSh"
38:        .LblShadowed(3).Caption = "Produced bY fLaSh"
39:        .LblShadowed(4).Caption = "Direct Connect P2P Network"
40:        .LblShadowed(5).Caption = "Direct Connect P2P Network"
41:        .lblHolder(30).Caption = "DNS 1"
42:        .lblHolder(34).Caption = "DNS 2"
43:        .lblHolder(35).Caption = "DNS 3"
44:        .lblHolder(36).Caption = "DNS 4"
45:        .lblHolder(38).Caption = "DNS 1"
46:        .lblHolder(39).Caption = "DNS 2"
47:        .lblHolder(41).Caption = "DNS 3"
48:        .lblHolder(42).Caption = "DNS 4"
49:        .lblHolder(55).Caption = "0"
50:        .lblHolder(56).Caption = "0"
51:        .lblHolder(57).Caption = "0 Bytes"
52:        .lblHolder(58).Caption = "0"
53:        .lblHolder(59).Caption = "0"
54:        .lblHolder(60).Caption = "0 Bytes"
55:        .lvwUsers.ColumnHeaders(2).Text = "IP"
56:        .lvwBans.ColumnHeaders(1).Text = "IP"
57:        .lvwPermIPBan.ColumnHeaders(1).Text = "IP"
58:        .lvwTempIPBan.ColumnHeaders(1).Text = "IP"
59:        .lvwTempIPBan.ColumnHeaders(2).Text = "Expire"
60:        .lblStatistics(0).Caption = "0"
61:        .lblStatistics(1).Caption = "0"
62:        .lblStatistics(2).Caption = "0"
63:        .lblStatistics(3).Caption = "0"
64:        .lblStatistics(4).Caption = "0"
65:        .lblStatistics(5).Caption = "0"
66:        .lblStatistics(6).Caption = "0"
67:        .lblStatistics(7).Caption = "0"
68:        .lblStatistics(8).Caption = "0"
69:        .lblStatistics(9).Caption = "0"
70:        .lblStatistics(10).Caption = "NickList:"
71:        .lblStatistics(11).Caption = "ConnectMe:"
72:        .lblStatistics(12).Caption = "RevConnectMe:"
73:        .lblStatistics(13).Caption = "Kicks:"
74:        .lblStatistics(14).Caption = "Redirects:"
75:        .lblStatistics(15).Caption = "BotINFO:"
76:        .lblStatistics(16).Caption = "NetINFO:"
77:        .lblStatistics(17).Caption = "Failed Sockets:"
78:        .lblStatistics(18).Caption = "Aborted Sockets:"
79:        .lblStatistics(19).Caption = "Requests:"
80:        .lblStatistics(20).Caption = "0 Bytes"
81:        .lblStatistics(21).Caption = "0 Bytes"
82:        .lblStatistics(22).Caption = "0 Bytes"
83:        .lblStatistics(23).Caption = "0 Bytes"
84:        .lblStatistics(24).Caption = "0 Bytes"
85:        .lblStatistics(25).Caption = "0 Bytes"
86:        .lblStatistics(26).Caption = "0 Kb/s"
87:        .lblStatistics(27).Caption = "0 Kb/s"
88:        .lblStatistics(28).Caption = "Main Chat messages:"
89:        .lblStatistics(29).Caption = "Private messages:"
90:        .lblStatistics(30).Caption = "A/P Searchs:"
91:        .lblStatistics(31).Caption = "MyINFOs:"
92:        .lblStatistics(32).Caption = "Total Send:"
93:        .lblStatistics(33).Caption = "Total Recived:"
94:        .lblStatistics(34).Caption = "Speed Send:"
95:        .lblStatistics(35).Caption = "Speed Recived:"
96:        .lblQueryDB(0).Caption = "Run"
97:        .lblQueryDB(1).Caption = "Cls"
98:    End With
  
  
102:   Exit Sub
103:
Err:
104:   HandleError Err.Number, Err.Description, Erl & "|mTranslate.ClearTranslations()"
End Sub
