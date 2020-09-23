VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

Public Sub InstallPlugins(Optional blnForceInstall As Boolean = False)
    '------------------------------------------------------------------
    'Purpose:   This Function Scan all DLL and TBL found in folder Plugins,
    '
    'Params:    none
    '
    'Returns:   none
    '
    '   Called by Sub Main (in begin of the aplication)
    '------------------------------------------------------------------
10:    On Error GoTo Err
11:    Dim strFiles()    As String
12:    Dim strFileTitle  As String
13:    Dim strTblFile    As String
14:    Dim strPlgServ    As String
15:    Dim intLoop       As Integer
16:    Dim intIndex      As Integer
17:    Dim blnPlgError   As Boolean
18:    Static bIsLoaded  As Boolean
    
20:    Erase g_objPlugin
21:    ReDim g_objPlugin(0)
     
       'Get dll files to array
24:    strFiles = g_objFileAccess.ListFiles(G_APPPATH & "\Plugins\*.dll")
    
       'Check if is empty the array..
27:    If IsNothingArray(strFiles) Then
28:        g_PluginsFound = False: Exit Sub
29:    End If
    
       'Loop in array..
32:    For intLoop = LBound(strFiles) To UBound(strFiles)
           'Get dll file title
34:        strFileTitle = g_objFileAccess.GetFileTitle(strFiles(intLoop))
           'Plguin server is dll_filename + .Main
36:        strPlgServ = strFileTitle & ".Main"
           'Check if the plugin is already installed..
38:        If (Not AlreadyInstalled(strPlgServ)) Or (blnForceInstall) Then
               'Check if tbl file (Type Lib) is found..
               'This file is generated when creating plugins in .NET
41:            strTblFile = strFileTitle & ".tlb"
42:            If g_objFileAccess.FileExists(strTblFile) Then
43:                If Not g_objActiveX.RegisterTBL(strTblFile) Then
44:                    AddLog "There was an error loading the Plugin:" & strFiles(intLoop) & " (with tbl)"
45:                    GoTo NextLoop
46:                Else
47:                    AddLog "Plugin " & strFiles(intLoop) & " (with tbl) installed with success."
48:                End If
49:            Else
50:                If Not g_objActiveX.RegisterActiveX(strFiles(intLoop), Register) Then
51:                    AddLog "There was an error loading the Plugin:" & strFiles(intLoop)
52:                    GoTo NextLoop
53:                Else
54:                    AddLog "Plugin " & strFiles(intLoop) & " installed with success."
55:                End If
56:            End If
57:        End If
        
59:        If Not g_objPlugin(0).Object Is Nothing Then
60:            ReDim Preserve g_objPlugin(UBound(g_objPlugin) + 1)
61:        End If
        
           'Set array index
64:        intIndex = UBound(g_objPlugin)
        
          'Turn var to false the possible error at loading the dll
67:        blnPlgError = False
         
           'Just case..
70:        On Error Resume Next
           'Create object ref..
72:        Set g_objPlugin(intIndex).Object = CreateObject(strPlgServ)
        
74:        g_objPlugin(intIndex).Name = g_objPlugin(intIndex).Object.Name
75:        g_objPlugin(intIndex).Version = g_objPlugin(intIndex).Object.Version
76:        g_objPlugin(intIndex).Author = g_objPlugin(intIndex).Object.Author
77:        g_objPlugin(intIndex).Description = g_objPlugin(intIndex).Object.Description
78:        g_objPlugin(intIndex).ReleaseDate = g_objPlugin(intIndex).Object.ReleaseDate
79:        g_objPlugin(intIndex).Comments = g_objPlugin(intIndex).Object.Comments
80:        g_objPlugin(intIndex).UseSetup = g_objPlugin(intIndex).Object.UseSetup
81:        g_objPlugin(intIndex).UseEvents = g_objPlugin(intIndex).Object.UseEvents
82:        g_objPlugin(intIndex).Index = intIndex

           'Check if plugin has been loaded
85:        If Err.Number Then
86:            AddLog "There was an error creating the Plugin object: " & strFiles(intLoop)
87:            blnPlgError = True
88:            Err.Clear
89:        Else 'Set link objects
90:            If Not g_objPlugin(intIndex).Object.LoadPlug = 1 Or _
                  Not SetLinkObjects(g_objPlugin(intIndex).Object) Then
92:                    AddLog "There was an error loading the plugin: " & strFiles(intLoop)
93:                    blnPlgError = True
94:            End If
95:        End If
        
97:        On Error GoTo Err
        
99:        If blnPlgError Then
100:            If UBound(g_objPlugin) > 0 Then
101:                ReDim Preserve g_objPlugin(UBound(g_objPlugin) - 1)
102:            Else
103:                Set g_objPlugin(0).Object = Nothing
104:            End If
105:        Else
106:            AddLog "Plugin " & strFiles(intLoop) & " loaded with success."
107:            g_PluginsFound = True
108:       End If
NextLoop:
110:       DoEvents
        
112:   Next

114:   Exit Sub
115:
Err:
       'In case of mistake it cancels all plugins!!
117:   g_PluginsFound = False
118:   Erase g_objPlugin
119:   HandleError Err.Number, Err.Description, Erl & "|" & "clsPlugins.InstallPlugins()"
End Sub
Private Function SetLinkObjects(ByVal objPlg As Object) As Boolean
1:    Dim objGlobal(24) As Object
2:    Dim colForms      As New Collection
3:    On Error GoTo Err
 
5:    colForms.Add frmBanName, "frmBanName"
6:    colForms.Add frmBanPerm, "frmBanPerm"
7:    colForms.Add frmBanTemp, "frmBanTemp"
8:    colForms.Add frmCAccounts, "frmCAccounts"
9:    colForms.Add frmChatRoom, "frmChatRoom"
10:    colForms.Add frmCommand, "frmCommand"
11:    colForms.Add frmDebugSC, "frmDebugSC"
12:    colForms.Add frmEditScintilla, "frmEditScintilla"
13:    colForms.Add frmHelp, "frmHelp"
14:    colForms.Add frmMulti, "frmMulti"
15:    colForms.Add frmNewScript, "frmNewScript"
16:    colForms.Add frmParent, "frmParent"
17:    colForms.Add frmPlan, "frmPlan"
18:    colForms.Add frmProperties, "frmProperties"
19:    colForms.Add frmReg, "frmReg"
20:    colForms.Add frmScript, "frmScript"
21:    colForms.Add frmSocks, "frmSocks"
22:    colForms.Add frmUpDate, "frmUpDate"
23:    colForms.Add frmUserInfo, "frmUserInfo"
 
       'Set link objects for all forms
26:    Set objGlobal(0) = frmHub
27:    Set objGlobal(1) = colForms

       'Set link objects for all classes
30:    Set objGlobal(2) = g_objChatRoom
31:    Set objGlobal(3) = g_objFunctions
       '
33:    Set objGlobal(4) = g_colIPBans
34:    Set objGlobal(5) = g_colUsers
35:    Set objGlobal(6) = g_colCommands
36:    Set objGlobal(7) = g_colMessages
37:    Set objGlobal(8) = g_colLanguages
38:    Set objGlobal(9) = g_colSWinsocks
39:    Set objGlobal(10) = g_colSVariables
       '
41:    Set objGlobal(11) = g_objRegistered
42:    Set objGlobal(12) = g_objSettings
43:    Set objGlobal(13) = g_objRegExps
44:    Set objGlobal(14) = g_objScheduler
45:    Set objGlobal(15) = g_objFileAccess
46:    Set objGlobal(16) = g_objStatus

48:    Set objGlobal(17) = g_objHighlighter
49:    Set objGlobal(18) = g_objSCI
50:    Set objGlobal(19) = g_objComDialog
51:    Set objGlobal(20) = g_objSQLite
52:    Set objGlobal(21) = g_objActiveX
53:    Set objGlobal(22) = g_colDictionary
54:    Set objGlobal(23) = g_objTimer
55:    Set objGlobal(24) = g_objTimersCol

57:    Call objPlg.Ini(objGlobal())
        
59:    Set colForms = Nothing
        
61:    SetLinkObjects = True
        
63:    Exit Function
64:
Err:
65:    On Error Resume Next
66:    SetLinkObjects = False
67:    AddLog "Plugin " & objPlg.Name & " error when establishing the link to dll objects.."
End Function
Private Function AlreadyInstalled(ByVal strPlgServ As String) As Boolean
1:    Dim objPlugin As Object
2:    Dim strTemp     As String
3:    On Error GoTo Err
    
5:    Set objPlugin = CreateObject(strPlgServ)
    
7:    strTemp = objPlugin.Name
    
9:    Set objPlugin = Nothing
    
11:   AlreadyInstalled = True

13:   Exit Function
14:
Err:
15:   AlreadyInstalled = False
End Function

Private Function IsNothingArray(sArray() As String) As Boolean
1:    Dim strTmp As String
2:    On Error GoTo Err
3:    strTmp = sArray(0)
4:    IsNothingArray = False
5:    Exit Function
Err:
6:    IsNothingArray = True
End Function
