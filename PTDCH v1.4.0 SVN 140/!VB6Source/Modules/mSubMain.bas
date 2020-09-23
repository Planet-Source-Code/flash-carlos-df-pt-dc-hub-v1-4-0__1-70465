Attribute VB_Name = "mSubMain"
Option Explicit

Private mXP         As clsXPTheme
Private mPlgins     As clsPlugins

Public Sub Main()
1:    On Error GoTo Err
2:    Dim i As Integer
    
4:     G_APPPATH = App.Path
5:     G_ERRORFILE = FreeFile
    
7:     frmLoading.Show
    
9:     Set g_objFileAccess = New clsFileAccess
10:    Set g_objActiveX = New clsActiveX
    
12:    Call CheckDirs
13:    Call CheckFiles
14:    Call CheckDLLs

       'Set app for XP Style
17:    Set mXP = New clsXPTheme
18:    Call mXP.InitializeXP
    
       'Inicialize Hub conf...
21:    Call Load(frmHub)
    
23:    Set mPlgins = New clsPlugins
    
25:    Call SetFlatBorder 'Set flat style in the pictureboxs and full selected row in listview..
    
27:    Call Load(frmHub)
    
       'Set caption to proper format
30:    frmHub.Caption = "PT Direct Connect Hub " & vbVersion & " - SVN " & vbSVNVersion
     
        'Skin ///////////////////////////////////////////////////////////////
        'Add themes to combobox
34:     frmHub.cmbSkin.AddItem "01-Defaut"
35:     frmHub.cmbSkin.AddItem "02-Cyan Blue"
36:     frmHub.cmbSkin.AddItem "03-Cyan Green"
37:     frmHub.cmbSkin.AddItem "04-Metallic"
38:     frmHub.cmbSkin.AddItem "05-Metallic Blue"
39:     frmHub.cmbSkin.AddItem "06-Metallic Green"
40:     frmHub.cmbSkin.AddItem "07-Metallic Navy Blue"
41:     frmHub.cmbSkin.AddItem "08-Metallic Oliver"
42:     frmHub.cmbSkin.AddItem "09-Texture Grain"
43:     frmHub.cmbSkin.AddItem "10-Texture Spater"
44:     frmHub.cmbSkin.AddItem "11-Texture Tiles"
45:     frmHub.cmbSkin.AddItem "12-Texture Toxedo"
46:     frmHub.cmbSkin.AddItem "13-Blue Berry"
47:     frmHub.cmbSkin.AddItem "14-Glace Table"
48:     frmHub.cmbSkin.AddItem "15-Pink"
49:     frmHub.cmbSkin.AddItem "16-Gun Blue"
50:     frmHub.cmbSkin.AddItem "17-Gun Metal"
    
       ' If checkbox is checked Randomize skin
53:     If g_objSettings.blSkin Then
54:        If g_objSettings.RndSkin And g_objSettings.blSkin Then
55:           Randomize
56:           g_objSettings.lngSkin = CInt((16) * Rnd + 1)
57:        End If
58:     End If
     
       'Set combobox text
       Select Case g_objSettings.lngSkin
         Case 1: frmHub.cmbSkin.Text = "01-Defaut"
         Case 2: frmHub.cmbSkin.Text = "02-Cyan Blue"
         Case 3: frmHub.cmbSkin.Text = "03-Cyan Green"
         Case 4: frmHub.cmbSkin.Text = "04-Metallic"
         Case 5: frmHub.cmbSkin.Text = "05-Metallic Blue"
         Case 6: frmHub.cmbSkin.Text = "06-Metallic Green"
         Case 7: frmHub.cmbSkin.Text = "07-Metallic Navy Blue"
         Case 8: frmHub.cmbSkin.Text = "08-Metallic Oliver"
         Case 9: frmHub.cmbSkin.Text = "09-Texture Grain"
         Case 10: frmHub.cmbSkin.Text = "10-Texture Spater"
         Case 11: frmHub.cmbSkin.Text = "11-Texture Tiles"
         Case 12: frmHub.cmbSkin.Text = "12-Texture Toxedo"
         Case 13: frmHub.cmbSkin.Text = "13-Blue Berry"
         Case 14: frmHub.cmbSkin.Text = "14-Glace Table"
         Case 15: frmHub.cmbSkin.Text = "15-Pink"
         Case 16: frmHub.cmbSkin.Text = "16-Gun Blue"
         Case 17: frmHub.cmbSkin.Text = "17-Gun Metal"
61:     End Select
        'END Skin \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        
64:     Set g_objAbout = New clsAbout
    
66:     frmHub.cmbChangeLog.Text = frmHub.cmbChangeLog.List(0)
         
        'Combo Boxs
69:     frmHub.cmbRegistered.AddItem "All Classes"
70:     frmHub.cmbRegistered.AddItem "Non-OPs only"
71:     frmHub.cmbRegistered.AddItem "OPs and above"
72:     frmHub.cmbRegistered.AddItem "Admins and above"
73:     frmHub.cmbRegistered.Text = "All Classes"
    
        'Load notepad text from the file
76:    If g_objFileAccess.FileExists(G_APPPATH & "\Settings\notepad.txt") Then
77:        frmHub.txtNotePad.Text = g_objFileAccess.ReadFile(G_APPPATH & "\Settings\notepad.txt")
78:    End If
    
        'Load motd text from the file
81:    If g_objFileAccess.FileExists(G_APPPATH & "\Settings\motd.txt") Then
82:         g_objSettings.JoinMsg = g_objFileAccess.ReadFile(G_APPPATH & "\Settings\motd.txt")
83:         frmHub.txtData(6).Text = g_objSettings.JoinMsg
84:    End If

        'Set text bot name in the tab Status
87:    frmHub.txtStForm.Text = g_objSettings.BotName
    
        'Load pictures from the resource file
90:    frmHub.picLog(0).Picture = LoadImage(102)
91:    frmHub.picLog(1).Picture = LoadImage(103)
92:    frmHub.picLog(2).Picture = LoadImage(104)
            
       'Load Plugins if then
95:    If g_objSettings.Plugins Then
96:        mPlgins.InstallPlugins
97:        frmHub.PlgXmlLoad
98:        frmHub.PlgRefreshGUI
99:        If frmHub.lvwPlugins.ListItems.Count > 1 Then _
                 frmHub.cmdPlugins(2).Enabled = True _
            Else frmHub.cmdPlugins(2).Enabled = False
102:        If frmHub.lvwPlugins.ListItems.Count > 0 Then _
                 frmHub.cmdPlugins(3).Enabled = True _
            Else frmHub.cmdPlugins(3).Enabled = False
105:        frmHub.cmdPlugins(4).Enabled = True
106:    Else
107:        frmHub.lvwPlugins.Enabled = False
108:        frmHub.cmdPlugins(3).Enabled = False
109:        frmHub.cmdPlugins(4).Enabled = False
110:    End If
    
        'Load Scripts
113:    Load frmScript
114:    frmScript.SLoadDir
115:    frmScript.XmlBooleanLoad
116:    frmScript.SReset -2, False, False

        'Set dimension
119:    frmHub.Width = g_WinMinW
120:    frmHub.Height = g_WinMinH
    
        'Festore form in the last windows position ..
123:    RestoreFormSize
    
125:    If g_objSettings.StartMinimized Then
126:        frmHub.Show
127:    End If
    
129:    frmHub.WindowState = vbMinimized

       'Check for updates if then
132:    If g_objSettings.AutoCheckUpdate Then
133:        Call frmUpDate.Notific(True)
134:    End If
    
136:    Call DelPopUpMenu
    
138:    frmHub.RefreshGUI True
    
140:    Call Unload(frmLoading)
    
142:    Set mXP = Nothing
143:    Set mPlgins = Nothing
    
145:    Set frmLoading = Nothing
    
147:    If Not g_objSettings.StartMinimized Then
148:        frmHub.WindowState = vbNormal
149:        frmHub.Visible = True
150:        frmHub.Show
151:    End If

152:    G_GUI_IS_LOADED = True
        
153:    Exit Sub
154:
Err:
155:    HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.Main()"
156:    Resume Next
End Sub
Private Sub CheckFiles()
1:    On Error GoTo Err
    
3:    If Not (g_objFileAccess.FileExists(G_APPPATH & "\DBs\userdb.mdb")) Then _
           Call LoadAndSaveEmptyDB
    
6:    If Not (g_objFileAccess.FileExists(G_APPPATH & "\Settings\VB.bin")) Then _
           Call LoadAndSaveHighlighter(enuHighlighter.VBScript)
    
9:    If Not (g_objFileAccess.FileExists(G_APPPATH & "\Settings\JScripts.bin")) Then _
           Call LoadAndSaveHighlighter(enuHighlighter.JScript)
    
12:   If Not (g_objFileAccess.FileExists(G_APPPATH & "\Settings\sql.bin")) Then _
           Call LoadAndSaveHighlighter(enuHighlighter.SQL)
    
15:   If Not g_objFileAccess.FileExists(G_APPPATH & "\Languages\English.xml") Then _
           Call LoadAndSaveXML(enuXML.EGLanguage)
    
18:   If Not g_objFileAccess.FileExists(G_APPPATH & "\Settings\UsersMessages.xml") Then _
           Call LoadAndSaveXML(enuXML.EGUsersMessages)
    
21:   If Not g_objFileAccess.FileExists(G_APPPATH & "\Settings\ScriptHelp.vbs") Then _
           Call LoadAndSaveScriptHelp
    
24:   Exit Sub
    
26:
Err:
28:   HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.CheckFiles()"
End Sub

Private Sub CheckDirs()

     'Make sure we don't loose owners previous hub dirs
3:   On Error GoTo Err

5:      Dim i As Integer
6:      Dim sPath(6) As String
      
8:       sPath(0) = G_APPPATH & "\Logs"
9:       sPath(2) = G_APPPATH & "\DBs"
10:      sPath(3) = G_APPPATH & "\Settings"
11:      sPath(4) = G_APPPATH & "\Scripts"
12:      sPath(5) = G_APPPATH & "\Plugins"
13:      sPath(6) = G_APPPATH & "\Languages"
       
15:      For i = 0 To 6
16:        If Not (g_objFileAccess.FileExists(sPath(i))) Then _
                g_objFileAccess.CreateDir sPath(i)
18:      Next i

20:   Exit Sub
21:
Err:
22:    HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.CheckDirs()"
End Sub

Private Sub CheckDLLs()
1:   On Error GoTo Err

3:   Dim strDll(4) As String
4:   Dim i As Integer
     
6:   strDll(0) = G_APPPATH & "\libbz2.dll"
7:   strDll(1) = G_APPPATH & "\MyIPTools.DLL"
8:   strDll(2) = G_APPPATH & "\zlib.dll"
9:   strDll(3) = G_APPPATH & "\SciLexer.dll"
10:  strDll(4) = G_APPPATH & "\SQLite3VB.dll"

11:  For i = 0 To 4
12:        If Not (g_objFileAccess.FileExists(strDll(i))) Then _
                MsgBox "Failed to initialize the '" & strDll(i) & "' interface." & vbNewLine & vbNewLine & _
                       "Please verify that '" & strDll(i) & "' is in the program " & _
                       "directory or the system32 directory.", vbCritical
16:  Next
     
18:  Exit Sub
19:
Err:
20:  HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.CheckDLLs()"
End Sub

Public Sub LoadDfsMessages()

2:     On Error GoTo Err
   
       'pre-define Strings LANGUAGE
5:     g_colMessages.Add "msgMissingStr", "This string is still missing in language file English.xml. Please contact DDCH-Team on http://www.shadowdc.com/forums"
6:     g_colMessages.Add "msgYourIP", "Your IP is %[ip] . Copy to clipboard?"
7:     g_colMessages.Add "msgExitPTDCH", "Press Yes to confirm exiting DDCH."
8:     g_colMessages.Add "msgUpdating", "Already in progress of downloading an update."
9:     g_colMessages.Add "msgRedirAll", "Redirect all users including operators?"
10:    g_colMessages.Add "msgGettingIP", "Already attempting to determine IP."
11:    g_colMessages.Add "msgMostRecent", "You have the most recent version of DDCH."
12:    g_colMessages.Add "msgDownload", "Do you wish to download it?"
13:    g_colMessages.Add "msgInvalidBanName", "You cannot ban names longer than 40 characters."
14:    g_colMessages.Add "msgClearPermIPs", "Press 'Yes' to confirm clearing the permanent IP ban list."
15:    g_colMessages.Add "msgAlreadyAdded", " has already been added to the list."
16:    g_colMessages.Add "msgInvalidBanLength", "The ban length must be numeric"
17:    g_colMessages.Add "msgClearTempIPs", "Press Yes to confirm clearing the temporary IP ban list."
18:    g_colMessages.Add "msgAlreadyRegged", " is already registered."
19:    g_colMessages.Add "msgInvalidRegName", "Registered names cannot be longer than 40 characters."
20:    g_colMessages.Add "msgInvalidPass", "Passwords cannot be longer than 20 characters."
21:    g_colMessages.Add "msgInvalidClass", "Invalid class."
22:    g_colMessages.Add "msgNotRegged", " is not registered."
23:    g_colMessages.Add "msgPortInUse", "Port %[port] is already in use."
24:    g_colMessages.Add "msgEnterRedirUsersAddress", "Enter the address to redirect the users to"
25:    g_colMessages.Add "msgEnterPM", "Enter private message to send to all users"
26:    g_colMessages.Add "msgEnterOpPM", "Enter private message to send to all operators"
27:    g_colMessages.Add "msgEnterBanName", "Enter the name to ban"
28:    g_colMessages.Add "msgEnterBanReason", "Enter the reason why you're banning the name (optional)"
29:    g_colMessages.Add "msgEnterReplace", "Enter the name to replace "
30:    g_colMessages.Add "msgEnterPermIP", "Enter the IP to permanently ban."
31:    g_colMessages.Add "msgEnterRemIP", "Enter the IP to remove"
32:    g_colMessages.Add "msgEnterDataToSel", "Enter the data to send to the selected users"
33:    g_colMessages.Add "msgEnterDataToAll", "Enter the data to send to all users"
34:    g_colMessages.Add "msgEnterLength", "Enter the length of the ban (in minutes)"
35:    g_colMessages.Add "msgKickReason", "Reason for kick"
36:    g_colMessages.Add "msgEnterRedirAddress", "Enter the address to redirect to"
37:    g_colMessages.Add "msgRedirReason", "Reason for redirect"
38:    g_colMessages.Add "msgBanReason", "Reason for ban"
39:    g_colMessages.Add "msgEnterTag", "Enter the tag to add"
40:    g_colMessages.Add "msgEnterTempIP", "Enter the IP to temporarily ban."
41:    g_colMessages.Add "msgEnterBanLength", "Enter the length in minutes to ban the IP."
42:    g_colMessages.Add "msgRenameBan", "Rename Ban"
43:    g_colMessages.Add "msgEnterRegName", "Enter the name you want to register"
44:    g_colMessages.Add "msgEnterPass", "Enter the password for "
45:    g_colMessages.Add "msgEnterClass", "Enter the class for "
46:    g_colMessages.Add "msgEnterNewPass", "Enter the new password for "
47:    g_colMessages.Add "msgEnterNewClass", "Enter the new class for "
48:    g_colMessages.Add "msgEnterNewName", "Enter the new name for "
49:    g_colMessages.Add "msgConfirmExit", "Confirm Exit"
50:    g_colMessages.Add "msgUpdate", "Update in progress"
51:    g_colMessages.Add "msgRedirUsers", "Redirect users"
52:    g_colMessages.Add "msgDetectIP", "Detect IP"
53:    g_colMessages.Add "msgNoUpdate", "No update avaliable"
54:    g_colMessages.Add "msgBanName", "Ban Name"
55:    g_colMessages.Add "msgConfirmClear", "Confirm clear"
56:    g_colMessages.Add "msgBanTempIP", "Ban Temporary IP"
57:    g_colMessages.Add "msgRegUser", "Register user"
58:    g_colMessages.Add "msgEditRegged", "Edit registered user"
59:    g_colMessages.Add "msgStartServing", "Start serving"
60:    g_colMessages.Add "msgMassMsg", "Mass Message"
61:    g_colMessages.Add "msgMassMsgOp", "Op Mass Message"
62:    g_colMessages.Add "msgMassMsgUnReg", "UnReg Mass Message"
63:    g_colMessages.Add "msgBanPermIP", "Ban Permanent IP"
64:    g_colMessages.Add "msgRemoveIP", "Remove IP"
65:    g_colMessages.Add "msgSendToSel", "Send data (selected)"
66:    g_colMessages.Add "msgSendToAll", "Send data (all)"
67:    g_colMessages.Add "msgKickSel", "Kick (selected)"
68:    g_colMessages.Add "msgRedirSel", "Redirect (selected)"
69:    g_colMessages.Add "msgBan", "Ban"
70:    g_colMessages.Add "msgAddTag", "Add tag"
71:    g_colMessages.Add "msgRenameUser", "Rename user"
72:    g_colMessages.Add "msgKick", "Kick"
73:    g_colMessages.Add "msgRedir", "Redirect"
74:    g_colMessages.Add "msgIPError", "An error occured while trying to retrieve your IP from www.whatismyip.org. Your IP, as can be determined locally, is %[ip]. Copy to clipboard?"
75:    g_colMessages.Add "msgDownloadError", "An error occured while downloading the update (%[number]: %[description])."
76:    g_colMessages.Add "msgUpdateError", "Update Error"
77:    g_colMessages.Add "msgIPNotValide", " IP addresse is not valide."
78:    g_colMessages.Add "msgDays", "Day(s):"
79:    g_colMessages.Add "msgHours", "Hour(s):"
80:    g_colMessages.Add "msgMinutes", "Minute(s):"
       'Defaut forms buttons -----------------------------------------------------
82:    g_colMessages.Add "msgClose", "Close"
83:    g_colMessages.Add "msgCancel", "Cancel"
84:    g_colMessages.Add "msgOK", "OK"
85:    g_colMessages.Add "msgAdd", "Add"
86:    g_colMessages.Add "msgRemame", "Rename"
87:    g_colMessages.Add "msgEdit", "Edit"
88:    g_colMessages.Add "msgClipboard", "Clipboard"
       ' Strings for frmCAccounts
90:    g_colMessages.Add "msgConvertRegs", "Convert Accounts for PTDCH database"
91:    g_colMessages.Add "msgConvRegsDBType", "Select the type of database:"
92:    g_colMessages.Add "msgConvRegsNoErr", "No Errors"
93:    g_colMessages.Add "msgConvRegsWithErr", "With Errors"
94:    g_colMessages.Add "msgConvRegsCount", "Accounts Count:"
95:    g_colMessages.Add "msgConvRegsDir", "Select directory of the "
96:    g_colMessages.Add "msgConvRegsNoXML", "XML file not found! "
97:    g_colMessages.Add "msgConvRegsBrowse", "Browse"
98:    g_colMessages.Add "msgConvRegsConv", "Convert"
99:    g_colMessages.Add "msgConvAccountN", "Account Nº"
100:   g_colMessages.Add "msgConvName", "Name"
101:   g_colMessages.Add "msgConvPassword", "Password"
102:   g_colMessages.Add "msgConvProfile", "Profile"
103:   g_colMessages.Add "msgConvErr", "Error Description"
       ' Strings for frmCommand
105:   g_colMessages.Add "msgCommand", "Edit Command"
106:   g_colMessages.Add "msgCmdEnabled", "Enabled"
107:   g_colMessages.Add "msgCmdTrigger", "Trigger"
108:   g_colMessages.Add "msgCmdMinClas", "Minimum class"
       'Strings for frmNewScript
110:   g_colMessages.Add "msgNewScript", "New Script"
111:   g_colMessages.Add "msgNewScriptName", "Enter the name of the script:"
112:   g_colMessages.Add "msgNewScriptType", "Select script type:"
113:   g_colMessages.Add "msgScriptAlready", " is a name already in use by another script."
       '
115:   g_colMessages.Add "msgRegAdd", "Register add at: "
116:   g_colMessages.Add "msgRegUpdate", "Register updated at: "
       'Strings for frmEditScintilla
118:   g_colMessages.Add "msgSCIFind", "Find"
119:   g_colMessages.Add "msgSCIReplace", "Replace"
120:   g_colMessages.Add "msgSCIGoTo", "GoTo"
121:   g_colMessages.Add "msgSCIFindNext", "Find Next"
122:   g_colMessages.Add "msgSCIReplace", "Replace"
123:   g_colMessages.Add "msgSCIReplaceAll", "Replace All"
124:   g_colMessages.Add "msgSCIFindPrev", "Find Previous"
125:   g_colMessages.Add "msgSCIGo", "Go"
126:   g_colMessages.Add "msgSCIWrap", "Wrap around"
127:   g_colMessages.Add "msgSCIWhole", "Match whole word only"
128:   g_colMessages.Add "msgSCICase", "Match case"
129:   g_colMessages.Add "msgSCIRegExp", "Regular expression"
130:   g_colMessages.Add "msgSCIFindWhat", "Find what:"
131:   g_colMessages.Add "msgSCIReplWith:", "Replace with:"
132:   g_colMessages.Add "msgSCIDestLine", "Destination Line:"
133:   g_colMessages.Add "msgSCICurrLine", "Current Line: "
134:   g_colMessages.Add "msgSCILastLine", "Last Line:"
135:   g_colMessages.Add "msgSCIColumn", "Column:"
136:   g_colMessages.Add "msgSCIReplTimes", "Replaced %[times] times"
       
138:  Exit Sub

140:
Err:
141:    HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.LoadDfsMessages()"
End Sub

Public Sub LoadDfsSettings()

2:    On Error GoTo Err

4:    g_objSettings.HubName = "PT DC Hub Demo"
5:    g_objSettings.HubDesc = "[PTDCH " & vbVersion & "]"
6:    g_objSettings.HubIP = "127.0.0.1"
7:    g_objSettings.BotName = "Security"

    'g_objSettings.BotEmail = vbNullString
    'g_objSettings.JoinMsg = vbNullstring
    'g_objSettings.RedirectIP = vbNullString
    'g_objSettings.RedirectAddress = vbNullString
    ' NEW REDIRECT ADDRESSES
    'g_objSettings.ForMinShareRedirectAddress = vbNullString
    'g_objSettings.ForMaxShareRedirectAddress = vbNullString
    'g_objSettings.ForMinSlotsRedirectAddress = vbNullString
    'g_objSettings.ForMaxSlotsRedirectAddress = vbNullString
    'g_objSettings.ForMaxHubsRedirectAddress = vbNullString
    'g_objSettings.ForSlotPerHubRedirectAddress = vbNullString
    'g_objSettings.ForNoTagHubRedirectAddress = vbNullString
    'g_objSettings.ForTooOldDcppRedirectAddress = vbNullString
    'g_objSettings.ForTooOldNMDCRedirectAddress = vbNullString
    'g_objSettings.ForBWPerSlotRedirectAddress = vbNullString
    'g_objSettings.ForFakeShareRedirectAddress = vbNullString
    'g_objSettings.ForFakeTagRedirectAddress = vbNullString
    'g_objSettings.ForPasModeRedirectAddress = vbNullString
       '
28:    g_objSettings.RegisterIP = "dcreg.mine.nu;reg.hublist.org;dcinfo.dynu.com;hubreg.1stleg.com"
29:    g_objSettings.Ports = "1411;411"
30:    g_objSettings.CSeperator = " "
31:    g_objSettings.MinShareMsg = "You have not met the minimum share."
32:    g_objSettings.DCppMinVersionMsg = "You are using an outdated DC++ client. Please goto http://dcplusplus.sourceforge.net/ and update it."
33:    g_objSettings.MinSlotsMsg = "You do not have enough slots open."
34:    g_objSettings.MaxSlotsMsg = "You have too many slots open."
35:    g_objSettings.HSRatioMsg = "You have not met the hub per slot ratio."
36:    g_objSettings.BSRatioMsg = "You have not met the bandwidth (in KB/s) per slot ratio (as measured by the limiter you are using)."
37:    g_objSettings.MaxHubsMsg = "You are connected to too many hubs. Disconnect from some and reconnect."
38:    g_objSettings.NMDCMinVersionMsg = "You are using an outdated NMDC client. Please goto http://www.neo-modus.com/ and update it. If you are using another client, please change the version setting."
39:    g_objSettings.DenyNoTagMsg = "You do not have an identification tag for your client (ie <++, <DC, etc). Please enable your tag, if possible."
40:    g_objSettings.MaxShareMsg = "You are sharing more than maximum allowed amount."
41:    g_objSettings.FakeShareMsg = "You are suspected of trying to cheat. Goodbye."
42:    g_objSettings.FakeTagMsg = "You are suspected of trying to cheat. Goodbye."
43:    g_objSettings.Socks5Msg = "Socks5 mode not allowed."
44:    g_objSettings.PassiveModeMsg = "Passive mode not allowed."
45:    g_objSettings.NoCOClientsMsg = "Chat only clients are not allowed in here."
46:    g_objSettings.HammeringRd = "a.b.c"

48:    g_objSettings.MaxUsers = 150
49:    g_objSettings.DefaultBanTime = 60
    'g_objSettings.IMinShare = 0
51:    g_objSettings.ScriptTimeout = 15000
52:    g_objSettings.DCMaxHubs = 50
53:    g_objSettings.DCOSlots = 0
54:    g_objSettings.MinSlots = 0
55:    g_objSettings.MaxSlots = 30
56:    g_objSettings.MinShareSize = 0
57:    g_objSettings.MaxShareSize = 0
58:    g_objSettings.CPrefix = 43
59:    g_objSettings.DCOSpeed = 10
60:    g_objSettings.FWInterval = 10000
61:    g_objSettings.FWBanLength = 120
62:    g_objSettings.FWMyINFO = 5
63:    g_objSettings.FWGetNickList = 5
64:    g_objSettings.FWActiveSearch = 15
65:    g_objSettings.FWPassiveSearch = 3
66:    g_objSettings.MaxPassAttempts = 3
67:    g_objSettings.DataFragmentLen = 2048
    'g_objSettings.SendJoinMsg = 0
    'svn 216
70:    g_objSettings.ConDropInterval = 250
71:    g_objSettings.FWDropMsgInterval = 300
    
73:    g_objSettings.DCSlotsPerHub = 1
74:    g_objSettings.DCBandPerSlot = 1
75:    g_objSettings.DCMinVersion = 0.181
76:    g_objSettings.NMDCMinVersion = 0

78:    g_objSettings.MinConnectCls = 1
    
80:    g_objSettings.MinClsConnectSend = True
81:    g_objSettings.MinClsSearchSend = True
    'g_objSettings.AutoCheckUpdate = False
83:    g_objSettings.AutoKickMLDC = True
    
'-----SOCKS5 CHECK--------------------------
    'g_objSettings.DenySocks5 = False
'-----SOCKS5 CHECK END----------------------
    'g_objSettings.DenyPassive = False

90:    g_objSettings.AutoRegister = True
    'g_objSettings.AutoRedirect = False
    'g_objSettings.AutoRedirectFull = False
    'g_objSettings.AutoRedirectNonReg = False
    'g_objSettings.AutoRedirectFullNonReg = False
    'g_objSettings.AutoRedirectFullNonOps = False
    'g_objSettings.AutoStart = False
97:    g_objSettings.CompactDBOnExit = True
    'g_objSettings.ConfirmExit = False
99:    g_objSettings.DCValidateTags = True
    'g_objSettings.DCIncludeOPed
101:    g_objSettings.OPBypass = True
102:    g_objSettings.PreloadWinsocks = True
103:    g_objSettings.SendMessageAFK = True
    'g_objSettings.RegOnly = False
    'g_objSettings.MentoringSystem = False
    'g_objSettings.PreventSearchBots = False
107:    g_objSettings.DescriptiveBanMsg = True

    'g_objSettings.UseVipChat = False
110:    g_objSettings.UseBotName = True
   'g_objSettings.DisablePassiveSeach = False
    
'-------------Notifications--------------------
114:    g_objSettings.PopUpNewReg = True
    'g_objSettings.PopUpOpConected = False
    'g_objSettings.PopUpOpDisconected = False
117:    g_objSettings.PopUpUserKick = True
118:    g_objSettings.PopUpUserBaned = True
    'g_objSettings.PopUpUserRedirected = False
    'g_objSettings.PopUpStartedServing = False
    'g_objSettings.PopUpStopedServing = False
'    g_objSettings.RedirectFBWPerSlot = False
'    g_objSettings.RedirectFFakeShare = False
'    g_objSettings.RedirectFFakeTag = False
'    g_objSettings.RedirectFPasMode = False
'    g_objSettings.PopUpCoreError = False
    
'------------End Here----------------------------
128:    g_objSettings.FilterCPrefix = True
129:    g_objSettings.EnabledCommands = True
    'g_objSettings.ScriptSafeMode = False
    'g_objSettings.StartMinimized = False
    'g_objSettings.SendMsgAsPrivate = False
    'g_objSettings.DenyNoTag = False
    'g_objSettings.HideFadeImg = False
135:    g_objSettings.CheckFakeShare = True
136:    g_objSettings.PreventGuessPass = True
137:    g_objSettings.EnableFloodWall = True
138:    g_objSettings.OpsCanRedirect = True
139:    g_objSettings.MinimizeTray = True
140:    g_objSettings.HideMyinfos = True
141:    g_objSettings.ACOClients = True
142:    g_objSettings.MinMyinfoFakeCls = 5

       'svn 159 , setable only in xml and scripts...
145:    g_objSettings.FWMainChat = 20
       'g_objSettings.FWGlobal = 60
147:    g_objSettings.ZLINELENGHT = 1400

        'Defaut language
150:    g_objSettings.Interface = "English"
        
        'System Priority defaut value
153:    g_objSettings.PriorityVal = 1
        'g_objSettings.PriorityBl = False
155:    frmHub.sldPriority.Enabled = False
         
        'set defaut skin -------------------------
158:    g_objSettings.blSkin = True
159:    g_objSettings.lngSkin = 1 '01-Defaut

161:    g_objSettings.Plugins = True
        
    '-------------- DATA BASE INTERFACE --------------
    '0 = MsAccess
    '1 = MySQL
166:    g_objSettings.DBType = 0
    'Variables only for MySQL connection
168:    g_objSettings.DBUserName = "Admin"
169:    g_objSettings.DBPassword = ""
170:    g_objSettings.DBServerAddresse = "localhost"
171:    g_objSettings.DBServerPort = 3306 'defaut mysql port
172:    g_objSettings.DBName = "userdb"  'defaut db name in mysql
    '-------------------------------------------------

174:    Call LoadDfsMessages
        
176:    DoEvents

178:    AddLog "Hub settings loaded."

180:  Exit Sub

182:
Err:
183:    HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.LoadDfsSettings()"
End Sub

'Start PTDCH at windows starting
Public Sub AddRegRun()
1:   On Error GoTo Err
2:   Dim Reg As Object
3:   Set Reg = CreateObject("Wscript.Shell")
4:   Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
5:   Exit Sub
6:
Err:
7:   HandleError Err.Number, Err.Description, Erl & "|mSubMain.AddRegRun()"
End Sub

Public Sub RemRegRun()
1:   On Error GoTo Err
2:   Dim Reg As Object
3:   Set Reg = CreateObject("Wscript.Shell")
4:   On Error Resume Next
5:   Reg.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName
6:   Exit Sub
7:
Err:
8:   HandleError Err.Number, Err.Description, Erl & "|mSubMain.RemRegRun()"
End Sub
