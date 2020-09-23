VERSION 5.00
Begin VB.Form frmScript 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleWidth      =   525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
'API calls
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

'Constants
Private Const PPC_JSCRIPT       As Integer = 1
Private Const PPC_VBSCRIPT      As Integer = 2
'Private Const PPC_PERLSCRIPT    As Integer = 3

Private Const PPC_LIBRARY       As Integer = 4 Or PPC_VBSCRIPT Or PPC_JSCRIPT 'Or PPC_PERL
Private Const PPC_INCLUDE       As Integer = 8 Or PPC_VBSCRIPT Or PPC_JSCRIPT 'Or PPC_PERL
Private Const PPC_ENDIF         As Integer = 16 Or PPC_VBSCRIPT
Private Const PPC_ELSE          As Integer = 32 Or PPC_VBSCRIPT
Private Const PPC_IF            As Integer = 64 Or PPC_VBSCRIPT
Private Const PPC_ELSEIF        As Integer = 128 Or PPC_VBSCRIPT
Private Const PPC_CONST         As Integer = 256 Or PPC_VBSCRIPT

Private Const CHR_SHARP         As Integer = 35
Private Const CHR_AT            As Integer = 64

'Types
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Public Enum enuMoveScript
    MoveUp = 0
    MoveDown = 1
End Enum

Private Const n_sDefScript = _
    "Option Explicit" & vbTwoLine & _
    "Sub Main()" & vbTwoLine & _
    vbTab & "MsgBox ""Hello World!"", , ""VBScript""" & vbTwoLine & _
    "End Sub" & vbTwoLine & _
    "Sub Error(Line)" & vbNewLine & _
    vbTab & "MsgBox Err.Description" & vbNewLine & _
    "End Sub" & vbNewLine
    
'Private variables
Private WithEvents m_objUpdate   As clsHTTPDownload
Attribute m_objUpdate.VB_VarHelpID = -1

Public Sub SLoadDir()
1:     Static blnLoaded    As Boolean

3:     Dim objSC           As ScriptControl
4:     Dim frmWS           As frmSocks
5:     Dim objSV           As clsDictionary
6:     Dim WFD             As WIN32_FIND_DATA
7:     Dim frmLoop         As Form

9:     Dim lngOne          As Long
10:    Dim lngTwo          As Long
11:    Dim intIndex        As Integer
12:    Dim i               As Integer
13:    Dim strTemp         As String
14:    Dim strLanguage     As String

16:    On Error GoTo Err

       'Check to see if it's been loaded before
       'If True, then unload forms/listitems
20:    If blnLoaded Then
           'Lock update for picSciMain. Prevents ghostly text from appearing. Always a good idea to use nonetheless ;-)
22:        LockWindowUpdate frmHub.hwnd
           'Delete controls
24:        lngTwo = frmHub.ScriptControl.UBound
25:        If lngTwo Then
26:            For lngOne = 1 To lngTwo
27:                Unload frmHub.ScriptControl(lngOne)
28:                Unload frmHub.tmrScriptTimer(lngOne)
29:            Next
30:        End If

           'Clear forms
33:        For Each frmLoop In Forms
                Select Case frmLoop.Name
                   Case "frmProperties", "frmSocks"
34:                     Unload frmLoop
35:             End Select
36:        Next

38:        Set frmLoop = Nothing

           ' Erase array
41:        Erase g_objSciLexer()

           ' Unload objects
44:        For i = 1 To frmHub.picSciMain.UBound
45:             Unload frmHub.picSciMain(i)
46:        Next i

48:        Set g_colSWinsocks = Nothing
49:        Set g_colSVariables = Nothing
51:    Else
52:        blnLoaded = True
53:    End If
    
55:    Set g_colSWinsocks = New Collection
56:    Set g_colSVariables = New Collection

58:    Call frmHub.NewTimersAPI
       
       'Clear listview and tab strip
61:    frmHub.lvwScripts.ListItems.Clear
62:    frmHub.tbsScripts.Tabs.Clear

       'Resize/clear out event array
65:    frmHub.SResizeArrEvent 1, False
    
       'If not found scripts..
68:    If Dir(G_APPPATH & "\Scripts\") = "" Then
69:        g_objFileAccess.WriteFile G_APPPATH & "\Scripts\" & g_colMessages.Item("msgNewScript") & ".vbs", n_sDefScript
70:    End If

       'Get first file handle
73:    lngOne = FindFirstFile(G_APPPATH & "\Scripts\*.*", WFD)
    
       'If it doesn't equal -1, then there are files
76:    If Not lngOne = -1 Then

78:        Do Until lngTwo = 18&

               'Can't be a directory
81:            If Not (WFD.dwFileAttributes And &H10) = vbDirectory Then

                    'Extract file name
84:                 lngTwo = InStrB(1, WFD.cFileName, vbNullChar)
                
86:                 If lngTwo Then _
                         strTemp = LeftB$(WFD.cFileName, lngTwo) _
                    Else strTemp = WFD.cFileName
                    
90:                 lngTwo = InStrRev(strTemp, ".")
                
                    'Check extension and determine language
                    Select Case Mid$(strTemp, lngTwo + 1)
                        Case "vbs", "script": strLanguage = "VBScript"
                        Case "js": strLanguage = "JScript"
'                       Case "pl": strLanguage = "PerlScript"
                        Case Else: GoTo NextLoop
94:                 End Select
                    
                    'Increment count
97:                intIndex = intIndex + 1
                    
                    'Load new code editor.. and add new item to listview
100:                Call LvwAddItem(intIndex, strTemp, strLanguage)
101:                Call AddNewCodeEditor(intIndex, strTemp, strLanguage)
                    
                   'Load objects
104:                Load frmHub.ScriptControl(intIndex)
105:                Load frmHub.tmrScriptTimer(intIndex)

                    'Get scriptcontrol
108:                Set objSC = frmHub.ScriptControl(intIndex)

                   'Load winsock collection
111:                Set frmWS = New frmSocks
112:                Set frmWS.Script = objSC

114:                frmWS.Tag = CStr(intIndex)
115:                g_colSWinsocks.Add frmWS, CStr(intIndex)
                    
                   'Load static var dictionary
118:                Set objSV = New clsDictionary
119:                g_colSVariables.Add objSV, CStr(intIndex)
    
                    'Set settings
122:                objSC.Language = strLanguage
123:                objSC.Timeout = g_objSettings.ScriptTimeout
124:                objSC.UseSafeSubset = g_objSettings.ScriptSafeMode

126:           End If
                
128:
NextLoop:
               'Get next file
130:           lngTwo = FindNextFile(lngOne, WFD)
        
               'Exit if it's zero
133:           If lngTwo = 0 Then Exit Do

135:        Loop

137:        If blnLoaded Then
138:            LockWindowUpdate 0
139:        End If

            'Redim array if needed
            Select Case intIndex
                Case 0, 1
                Case Else
142:                    frmHub.SResizeArrEvent intIndex, False
143:        End Select

145:   End If
      
147:   Exit Sub
    
149:
Err:
150:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SLoadDir()"
End Sub
Public Sub SLoadScript(ByVal strName As String)
1:     Dim objSC           As ScriptControl
2:     Dim frmWS           As frmSocks
3:     Dim objSV           As clsDictionary
4:     Dim frmLoop         As Form

6:     Dim intIndex        As Integer
7:     Dim strLanguage     As String
8:     Dim lngTwo As Long

10:    On Error GoTo Err

12:    If frmHub.ScriptControl.UBound <> 0 Then
13:       intIndex = (frmHub.ScriptControl.UBound + 1)
14:    Else
15:       intIndex = 1
16:    End If
       
       'Check extension and determine language
19:    If Right(strName, 3) = "vbs" Or Right(strName, 6) = "script" Then
20:         strLanguage = "VBScript"
21:    ElseIf Right(strName, 2) = "js" Then
22:         strLanguage = "JScript"
'19:    ElseIf Right(strName, 2) = "pl" Then
'20:         strLanguage = "PerlScript"
25:    Else
26:         Exit Sub
27:    End If

29:    Call LvwAddItem(intIndex, strName, strLanguage)
30:    Call AddNewCodeEditor(intIndex, strName, strLanguage)
                    
       'Load objects
33:    Load frmHub.ScriptControl(intIndex)
34:    Load frmHub.tmrScriptTimer(intIndex)
                    
       'Get scriptcontrol
37:    Set objSC = frmHub.ScriptControl(intIndex)

       'Load winsock collection
40:    Set frmWS = New frmSocks
41:    Set frmWS.Script = objSC

43:    frmWS.Tag = CStr(intIndex)
44:    g_colSWinsocks.Add frmWS, CStr(intIndex)
                    
       'Load static var dictionary
47:    Set objSV = New clsDictionary
48:    g_colSVariables.Add objSV, CStr(intIndex)
    
       'Set settings
51:    objSC.Language = strLanguage
52:    objSC.Timeout = g_objSettings.ScriptTimeout
53:    objSC.UseSafeSubset = g_objSettings.ScriptSafeMode

55:    If frmHub.ScriptControl.UBound <> 1 Or frmHub.ScriptControl.UBound <> 0 Then
56:         frmHub.SResizeArrEvent intIndex, False
57:    Else
58:         frmHub.SResizeArrEvent intIndex, True
59:    End If
       
61:    Exit Sub
    
63:
Err:
64:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SLoadScript(" & strName & ")"
End Sub

Public Sub SSave(Optional ByVal intIndex As Integer = 0)
1:    On Error GoTo Err

3:    Dim strTemp       As String
4:    Dim lngSelected   As Long
      
9:    With frmHub

11:       If intIndex = 0 Then
              'Save selected
13:           lngSelected = .IsListViewSelected(.lvwScripts)
14:           If lngSelected <> -1 Then
15:                strTemp = g_objSciLexer(lngSelected).Text
16:                .tbsScripts.Tabs(lngSelected).Tag = strTemp
17:                g_objFileAccess.WriteFile G_APPPATH & "\Scripts\" & .tbsScripts.Tabs(lngSelected).Key, strTemp
                   '
19:                g_objSciLexer(lngSelected).ClearUndoBuffer
20:           End If
21:       Else
              'Save by Index
23:           strTemp = g_objSciLexer(intIndex).Text
              '
25:           .tbsScripts.Tabs(intIndex).Tag = strTemp
26:           g_objFileAccess.WriteFile G_APPPATH & "\Scripts\" & .tbsScripts.Tabs(intIndex).Key, strTemp
              '
28:           g_objSciLexer(intIndex).ClearUndoBuffer
29:       End If
          
31:   End With

33:   Exit Sub
34:
Err:
35:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SSave(" & intIndex & ")"
End Sub
Public Sub SResetByName(ByVal strName As String, Optional ByVal blnUpDateCode As Boolean = True, Optional ByVal blnFirst As Boolean)
4:    Dim intIndex    As Integer
5:    Dim strTemp   As String
    
7:    On Error GoTo Err

9:    With frmHub
    
11:         For intIndex = 1 To .lvwScripts.ListItems.Count
12:            If .lvwScripts.ListItems(intIndex).Text = strName Then
13:                If blnUpDateCode Then .tbsScripts.Tabs(intIndex).Tag = g_objSciLexer(intIndex).Text
14:                If SetSReset(intIndex, .ScriptControl(intIndex), blnFirst) Then
15:                     If blnUpDateCode Then Call SSave(intIndex)
16:                End If
17:                Exit Sub
18:            End If
19:        Next
    
21:   End With
    
23:   Exit Sub
    
25:
Err:
26:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SResetByName(" & strName & ")"
End Sub

'Reset script.. only update the scripts to file, if no errors
Public Sub SReset(Optional ByVal lngSel As Long, Optional ByVal blnUpDateCode As Boolean = True, Optional ByVal blnFirst As Boolean)
4:    Dim intIndex    As Integer
5:    Dim strTemp     As String
6:    Dim lngSelected As Long
     
8:    On Error GoTo Err
    
10:    With frmHub
                        
       Select Case lngSel
        
            '*********************************************
            Case -2 'All checked scripts
            '*********************************************
            
16:                 For intIndex = 1 To .lvwScripts.ListItems.Count
17:                     If .lvwScripts.ListItems(intIndex).Checked Then
18:                            If blnUpDateCode Then .tbsScripts.Tabs(intIndex).Tag = g_objSciLexer(intIndex).Text
19:                         If SetSReset(intIndex, .ScriptControl(intIndex), blnFirst) Then
20:                                If blnUpDateCode Then Call SSave(intIndex)
21:                         End If
22:                     End If
23:                 Next

            '*********************************************
            Case -1 'All scripts
            '*********************************************

28:                 For intIndex = 1 To .lvwScripts.ListItems.Count
29:                     If blnUpDateCode Then .tbsScripts.Tabs(intIndex).Tag = g_objSciLexer(intIndex).Text
30:                     If SetSReset(intIndex, .ScriptControl(intIndex), blnFirst) Then
31:                            If blnUpDateCode Then Call SSave(intIndex)
32:                     End If
33:                 Next
            
            '*********************************************
            Case 0 'Selected script
            '*********************************************
            
38:                lngSelected = .IsListViewSelected(.lvwScripts)
39:                If lngSelected <> -1 Then
40:                    If blnUpDateCode Then .tbsScripts.Tabs(lngSelected).Tag = g_objSciLexer(lngSelected).Text
41:                    If SetSReset(lngSelected, .ScriptControl(lngSelected), blnFirst) Then
42:                         If blnUpDateCode Then Call SSave(lngSelected)
43:                    End If
44:                End If

            '*********************************************
            Case Is > 0 ' by Index
            '*********************************************

49:               If blnUpDateCode Then .tbsScripts.Tabs(lngSel).Tag = g_objSciLexer(lngSel).Text
50:               If SetSReset(CInt(lngSel), .ScriptControl(CInt(lngSel)), blnFirst) Then
51:                        If blnUpDateCode Then Call SSave(CInt(lngSel))
52:               End If
                  
54:        End Select
        
56:   End With
    
58:   Exit Sub
    
60:
Err:
61:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SReset(" & lngSel & ")"
End Sub

Private Function SetSReset(ByVal intIndex As Integer, ByVal objSC As ScriptControl, Optional ByVal blnFirst As Boolean) As Boolean
1:    Dim intChar     As Integer
2:    Dim strCode     As String
3:    Dim strPath     As String

5:     If Not blnFirst Then
          'Raise UnloadMain() event
7:        On Error Resume Next
8:        objSC.Run "UnloadMain"
9:     End If

11:    On Error GoTo Err
  
      'Reset script code/objects then readd objects
14:    objSC.Reset

      'Forms
17:    objSC.AddObject "Core", frmHub
18:    objSC.AddObject "frmHub", frmHub
19:    objSC.AddObject "frmScript", frmScript
        
      'Default VB objects
22:    objSC.AddObject "App", App
23:    objSC.AddObject "Forms", Forms

      'Default DC objects
26:    objSC.AddObject "tmrScriptTimer", frmHub.tmrScriptTimer(intIndex)
27:    objSC.AddObject "colUsers", g_colUsers

      'Extended PTDCH objects
30:    objSC.AddObject "wskScript", g_colSWinsocks(CStr(intIndex)).wskScript
31:    objSC.AddObject "colStatic", g_colSVariables(CStr(intIndex))
32:    objSC.AddObject "ScriptCtrl", objSC
33:    objSC.AddObject "Settings", g_objSettings
34:    objSC.AddObject "Functions", g_objFunctions, True
35:    objSC.AddObject "colRegistered", g_objRegistered
36:    objSC.AddObject "colIPBans", g_colIPBans
37:    objSC.AddObject "FileAccess", g_objFileAccess
38:    objSC.AddObject "colCommands", g_colCommands
39:    objSC.AddObject "RegExps", g_objRegExps
40:    objSC.AddObject "colLanguages", g_colLanguages
41:    objSC.AddObject "Status", g_objStatus
42:    objSC.AddObject "Sheduler", g_objScheduler
43:    objSC.AddObject "ChatRoom", g_objChatRoom
44:    objSC.AddObject "ActiveX", g_objActiveX
45:    objSC.AddObject "Debug", frmDebugSC
46:    objSC.AddObject "TimersAPI", frmHub.oTimersAPI

       'Get first char to identify language
49:    intChar = AscW(objSC.Language)
    
51:    If intChar = 80 Then
52:        On Error Resume Next
53:        objSC.AddCode frmHub.tbsScripts.Tabs(intIndex).Tag
54:        On Error GoTo Err
55:    Else
          'Prepare code buffer
57:        strCode = GenTempFile()
58:        g_objFileAccess.WriteFile strCode, frmHub.tbsScripts.Tabs(intIndex).Tag
        
           'Do preparsing actions if JScript/VBScript
61:        strPath = SSPrereset(objSC, strCode, vbNullString, intChar = 86)
        
           'Read code to control
64:        On Error Resume Next
65:        objSC.AddCode g_objFileAccess.ReadFile(strPath)
66:        On Error GoTo Err
67:        g_objFileAccess.DeleteFile strPath
68:        g_objFileAccess.DeleteFile strCode
69:    End If

       'Clear error text..
72:    frmHub.txtScriptError.Text = ""

       'If there was an error, then tell the user, and cancel reset
75:    If objSC.Error.Number Then

         'Report error / Add Log
78:       MsgBeep beepSystemDefault 'alert sound
79:       frmHub.txtScriptError.Text = "[" & Now & "] Script Error: " & frmHub.lvwScripts.ListItems(intIndex).Text & " (" & objSC.Error.Description & ") on line " & CInt(objSC.Error.Line)
80:       AddLog "Script Error: " & frmHub.lvwScripts.ListItems(intIndex).Text & " (" & objSC.Error.Description & ") on line " & CInt(objSC.Error.Line)
          
          'Remove code/objects again
83:       objSC.Reset
84:       frmHub.SClearEvents intIndex

          'Make sure listitem is unchecked
87:       With frmHub.lvwScripts
88:            .ListItems(intIndex).Checked = False
89:            .ListItems(intIndex).SubItems(1) = "Inactive"
90:            .ListItems(intIndex).SubItems(3) = Now
91:       End With

          ' return False
94:       SetSReset = False

96:    Else
          
          'Set events
99:       frmHub.SFindEvents intIndex
        
          'Make sure listitem is checked
102:       With frmHub.lvwScripts
103:            If .ListItems(intIndex).Checked = False Then _
                    .ListItems(intIndex).Checked = True
105:           .ListItems(intIndex).SubItems(1) = "Active"
106:           .ListItems(intIndex).SubItems(3) = Now
107:       End With
            
          'Run Main
110:      On Error Resume Next
111:      objSC.Run "Main"
112:      On Error GoTo Err
          
114:      If G_GUI_IN_UNLOAD Then
              'NOTE: it is possible if is used this code in this script:
              '     Sub Main()
              '         frmHub.APP_TERMINATE
              '     End Sub
              'This function set nothing all PTDCH objects.. however this process still continues!!
120:          End ' Hard end
121:      End If
          
123:      If objSC.Error.Number Then
              'Report error / Add Log
125:          MsgBeep beepSystemDefault 'alert sound
126:          frmHub.txtScriptError.Text = "[" & Now & "] Script Error: " & frmHub.lvwScripts.ListItems(intIndex).Text & " (" & objSC.Error.Description & ") on line " & CInt(objSC.Error.Line)
127:          AddLog "Script Error: " & frmHub.lvwScripts.ListItems(intIndex).Text & " (" & objSC.Error.Description & ") on line " & CInt(objSC.Error.Line)
128:          SetSReset = False
129:      Else
130:          SetSReset = True
131:      End If

133:   End If
   
135:   Exit Function
    
137:
Err:
138:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SetSReset(" & intIndex & ")"
End Function
Public Sub SStopByName(ByVal strName As String)
         
2:    Dim intIndex    As Integer
3:    Dim strTemp   As String
    
5:    On Error GoTo Err

7:    With frmHub
    
9:         For intIndex = 1 To .lvwScripts.ListItems.Count
10:            If .lvwScripts.ListItems(intIndex).Text = strName Then
11:                 SetSStop intIndex, .ScriptControl(intIndex)
12:                 Exit Sub
13:            End If
14:        Next
    
16:   End With
    
18:   Exit Sub
    
20:
Err:
21:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SStopByName(" & strName & ")"
End Sub

Public Sub SStop(Optional ByVal lngSel As Long)

2:    Dim intIndex As Integer
3:    Dim lngSelected As Long
    
5:    On Error GoTo Err
    
7:    With frmHub

       Select Case lngSel
        
            '*********************************************
            Case -2 'All checked scripts
            '*********************************************

13:                 For intIndex = 1 To .lvwScripts.ListItems.Count
14:                     If .lvwScripts.ListItems(intIndex).Checked Then
                            'Stop script..
16:                         SetSStop intIndex, .ScriptControl(intIndex)
17:                     End If
18:                 Next

            '*********************************************
            Case -1 'All scripts
            '*********************************************
            
23:                 For intIndex = 1 To .lvwScripts.ListItems.Count
                        'Stop script..
25:                     SetSStop intIndex, .ScriptControl(intIndex)
26:                 Next
                
            '*********************************************
            Case 0 'Selected script
            '*********************************************

31:                 lngSelected = .IsListViewSelected(.lvwScripts)
32:                 If lngSelected <> -1 Then
                        'Stop script..
34:                     SetSStop lngSelected, .ScriptControl(lngSelected)
35:                 End If

            '*********************************************
            Case Is > 0 ' by Index
            '*********************************************
                   
                   'Stop script..
41:                SetSStop lngSel, .ScriptControl(CInt(lngSel))

43:        End Select

45:   End With

47:   Exit Sub
    
49:
Err:
50:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SStop(" & lngSel & ")"
End Sub

Private Sub SetSStop(ByVal intIndex As Integer, objSC As ScriptControl)
1:    On Error Resume Next
    
      'Raise UnloadMain() event
4:    objSC.Run "UnloadMain"
    
6:    On Error GoTo Err
        
      'Reset all code/objects
9:    objSC.Reset
    
      'Set script event enabled status' to false
12:   frmHub.SClearEvents intIndex

      'Uncheck listitem
15:   With frmHub.lvwScripts
16:        If .ListItems(intIndex).Checked Then _
                .ListItems(intIndex).Checked = False
18:        .ListItems(intIndex).SubItems(1) = "Inactive"
20:        .ListItems(intIndex).SubItems(3) = Now
21:   End With
      
23:   Exit Sub

25:
Err:
26:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SetSStop(" & intIndex & ")"
End Sub

Private Sub AddNewCodeEditor(ByVal intIndex As Integer, _
                             ByVal strName As String, _
                             ByVal strLanguage As String)
3:     On Error GoTo Err

5:     Dim strTemp  As String, strTemp2 As String

7:     ReDim Preserve g_objSciLexer(intIndex)
            
9:     Set g_objSciLexer(intIndex) = New clsYScintilla

11:    strTemp = g_objFileAccess.ReadFile(G_APPPATH & "\Scripts\" & strName)
        
13:    Load frmHub.picSciMain(intIndex)

       Select Case intIndex
            Case 1: frmHub.picSciMain(intIndex).Visible = True
            Case Else: frmHub.picSciMain(intIndex).Visible = False
15:    End Select

17:    If Len(strName) > 18 Then _
            strTemp2 = Left(strName, 16) & ".." _
       Else strTemp2 = strName

21:    frmHub.tbsScripts.Tabs.Add (intIndex), strName, strTemp2
22:    frmHub.tbsScripts.Tabs(intIndex).Tag = strTemp
         
24:    g_objSciLexer(intIndex).CreateScintilla frmHub.picSciMain(intIndex)

26:    g_objSciLexer(intIndex).SetFixedFont "Courier New", 10
        
       'Give the scrollbar a nice long width to
       'handle a long line which may occur.
30:    g_objSciLexer(intIndex).ScrollWidth = 10000
       'This is absolutly an imperative line
32:    g_objSciLexer(intIndex).Attach frmHub.picSciMain(intIndex)
33:    g_objSciLexer(intIndex).Folding = True
34:    g_objSciLexer(intIndex).LineNumbers = True
35:    g_objSciLexer(intIndex).AutoIndent = True
36:    g_objSciLexer(intIndex).SetMarginWidth MarginLineNumbers, 50
37:    g_objSciLexer(intIndex).ContextMenu = True
38:    g_objSciLexer(intIndex).LineBreak = SC_EOL_CRLF


41:    Call g_objHighlighter.SetHighlighterBasedOnExt(g_objSciLexer(intIndex), strName)

43:    g_objSciLexer(intIndex).Text = strTemp
            
45:    frmHub.tbsScripts.ZOrder vbSendToBack
       
47:    frmHub.Form_Resize

49:    g_objSciLexer(intIndex).ClearUndoBuffer
       
51:    Exit Sub
52:
Err:
53:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.AddNewCodeEditor(" & intIndex & ")"
End Sub

Public Sub SProperties(ByRef strIndex As String, _
                       ByRef strName As String, _
                       ByRef lngType As Long)

4:    Dim frmProp As Form
5:    Dim Modal As Byte

7:    Dim sFile As String
8:    Dim sXML As String

10:    On Error GoTo Err
    
      'If *.xml file properties not found, create on new..
13:     sFile = G_APPPATH & "\Scripts\" & (LeftB(strName, InStrB(1, strName, ".") - 1) & ".xml")
14:     If Not g_objFileAccess.FileExists(sFile) Then
15:         sXML = _
            "<Properties>" & vbNewLine & _
            vbTab & "<Author></Author>" & vbNewLine & _
            vbTab & "<Copyright></Copyright>" & vbNewLine & _
            vbTab & "<Version></Version>" & vbNewLine & _
            vbTab & "<Website></Website>" & vbNewLine & _
            vbTab & "<Description></Description>" & vbNewLine & _
            vbTab & "<Comments></Comments>" & vbNewLine & _
            "</Properties>"
24:         g_objFileAccess.WriteFile (sFile), sXML
25:     Else 'if found
           'Loop through to find if the form exists
27:         For Each frmProp In Forms
               'Check to see if it's the right kind of form
29:            If frmProp.Name = "frmProperties" Then
30:               If frmProp.Tag = strIndex Then
                     'Set focus
32:                  frmProp.SetFocus
33:                  Set frmProp = Nothing
34:                  Exit Sub
35:               End If
36:            End If
37:         Next
38:     End If
        'We haven't found a form and must create one
40:     Set frmProp = New frmProperties

42:     frmProp.Tag = strIndex
43:     frmProp.PType = lngType
44:     frmProp.file = strName

        'Set Full Selected Rum
47:     LVFullRow frmProp.lvwProperties.hwnd
48:     frmProp.stBar.Panels(1).Text = strName
        
        ' hook window for sizing control
        ' Disable the following line if you will be debugging form.
52:     Call HookWin(frmProp.hwnd, G_PrWnd)

54:     frmProp.Show Modal, frmHub
    
56:  Exit Sub
    
58:
Err:
59:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.PSLoad()"
End Sub

Public Sub XmlBooleanLoad()
1:   Dim objXML          As clsXMLParser
2:   Dim objNode         As clsXMLNode
3:   Dim colNodes        As Collection
4:   Dim colSubNodes     As Collection

6:   Dim strTemp         As String
7:   Dim i               As Integer

9:   On Error GoTo Err

11:    Set objXML = New clsXMLParser
      
13:    strTemp = G_APPPATH & "\Settings\Scripts.xml"

15:    If g_objFileAccess.FileExists(strTemp) Then
         
17:       objXML.Data = g_objFileAccess.ReadFile(strTemp)
18:       objXML.Parse

20:       Set colNodes = objXML.Nodes(1).Nodes

22:       On Error Resume Next

24:       For Each objNode In colNodes
25:            Set colSubNodes = objNode.Attributes
26:            With frmHub.lvwScripts
27:                For i = 1 To .ListItems.Count
28:                   If .ListItems(i).Text = CStr(colSubNodes("Name").Value) Then
29:                         .ListItems(i).Checked = colSubNodes("Value").Value
30:                   End If
31:                Next
32:            End With
33:       Next

35:       On Error GoTo Err
    
37:       objXML.Clear
    
39:       Set objNode = Nothing
40:       Set colSubNodes = Nothing
41:       Set colNodes = Nothing

43:   End If

45:   Exit Sub
    
47:
Err:
48:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.XmlBooleanLoad()"
End Sub

Public Sub XmlBooleanSave()
1:    On Error GoTo Err
2:    Dim intFF       As Integer
3:    Dim strTemp     As String
4:    Dim i           As Integer

      'Save Scripts Value (Checked or UnChecked)
    
8:     strTemp = G_APPPATH & "\Settings\Scripts.xml"

10:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp
 
12:    intFF = FreeFile

14:    Open strTemp For Append As intFF
15:      Print #intFF, "<Scripts>"
16:        With frmHub.lvwScripts
17:            For i = 1 To .ListItems.Count
18:                 Print #intFF, vbTab & "<Script Name=""" & .ListItems(i).Text & """" & " Value=""" & .ListItems(i).Checked & """ />"
19:            Next
20:        End With
21:      Print #intFF, "</Scripts>";
22:    Close intFF
    
24:   Exit Sub
    
26:
Err:
27:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.XmlBooleanSave()"
End Sub

Private Function SSPrereset(ByRef objSC As ScriptControl, _
                            ByRef strRead As String, _
                            ByRef strWrite As String, _
                            ByRef blnVBScript As Boolean) As String
    '------------------------------------------------------------------
    'Purpose:   To perform preprocessing commands given by the script,
    '           which are denoted by the symbol # or @
    '
    'Params:
    '           objSC:          Reference to script's control object
    '           strRead:        Path to script input file
    '           strWrite:       Path to script output file (generated if
    '                           not given) where code with the preprocessor
    '                           instructions are interpreted
    '           blnVBScript:    Toggles if language is VBScript or JScript
    '
    'Returns:
    '           strWrite (if it was given, it returns the same as the
    '           given path, otherwise it returns the path to the
    '           temporary file generated)
    '------------------------------------------------------------------
    
22:    Dim intRead         As Integer
23:    Dim intWrite        As Integer
24:    Dim intFlag         As Integer
    
26:    On Error GoTo Err
    
    'If VBScript, we search for #
    'If JScript, we search for @
30:    If blnVBScript Then
31:        intFlag = CHR_SHARP
32:    Else
33:        intFlag = CHR_AT
34:    End If
    
    'Open script for reading
37:    intRead = FreeFile
38:    Open strRead For Binary Access Read Lock Read Write As intRead
    
    'Create temporary file for appending to
41:    intWrite = FreeFile
    
43:    If StrPtr(strWrite) Then
44:        If LenB(Dir(strWrite)) Then
45:            Kill strWrite
46:        End If
47:    Else
48:        strWrite = GenTempFile()
49:    End If
        
51:    Open strWrite For Append Lock Read Write As intWrite
    
    'Begin preprocessing
54:    Preproc intRead, intWrite, intFlag, objSC
    
    'Close file handles
57:    Close intRead
58:    Close intWrite
    
    'Return path to code
61:    SSPrereset = strWrite
    
63:    Exit Function
    
65:
Err:
66:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SSPrereset(, """ & strRead & """, """ & strWrite & """, " & blnVBScript & ")"
End Function

Private Sub ParseIf(ByVal intRead As Integer, _
                    ByVal intWrite As Integer, _
                    ByVal intChar As Integer, _
                    ByRef strExp As String, _
                    ByRef objSC As ScriptControl)
    '------------------------------------------------------------------
    'Purpose:   To interpret the preprocessor command #If, #ElseIf
    '           #Else and #End If.
    '
    '           #If/etc works just like their counterparts without
    '           the # except for one difference; the boolean expression
    '           is only ever evaluated once before starting the script.
    '           The code included in the script is for whichever
    '           statement for the #Ifs/ElseIfs evaluates to true first
    '           or the code for #Else if all are false and it is
    '           included.
    '
    '           Example:
    '               #If <expression> Then
    '                   'Include code for this exp
    '               #ElseIf <expression> Then
    '                   'If the first wasn't true, try this
    '               #Else
    '                   'Alright neither was true, include this code
    '               #End If
    '
    'Params:
    '           intRead:        File handle on input for the script code
    '           intWrite:       File handle on file to output parsed
    '                           code to
    '           intChar:        Character which denotes a preprocessor
    '                           command (#/@)
    '           strExp:         Expression which is to be evaluated
    '                           to determine if the code is to be
    '                           included or skipped (includes #If
    '                           and trailing Then)
    '           objSC:          Script's control object
    '------------------------------------------------------------------
    
39:    Dim strLine     As String
40:    Dim intCount    As Integer
41:    Dim intRet      As Integer
    
43:    On Error GoTo Err
    
    'Set count of totatal #End If to 1
46:    intCount = 1
    
    'Trim out expression to evaluate
49:    strExp = MidB$(strExp, 9, LenB(strExp) - 19)
    
    'Get a boolean out of it
52:    intRet = CBool(objSC.Eval(strExp))
    
    'Keep looping until we get a flag to exit
    'In other words, keep looping until we find an #If or #ElseIf which
    'evaluates to true, an #Else which is always true, or an #End If
57:    Do Until intRet Or EOF(intRead)
        'Make sure we don't pass the file boundaries
59:        Do Until EOF(intRead)
            'Take in a line and trim
61:            Line Input #intRead, strLine
62:            strLine = TrueTrim(strLine)
            
            'Skip null lines
65:            If LenB(strLine) Then
                'If it is a preproc command, everything is fine and dandy
67:                If AscW(strLine) = intChar Then
                    'Ignore #Include, #If, #Library, etc because this part of the
                    'code block is being ignored anyways
                    Select Case PreProcCmd(strLine, intChar)
                        Case PPC_IF
70:                            intCount = intCount + 1
                        Case PPC_ELSEIF
                            'Another expression to parse out and check
72:                            intRet = CBool(objSC.Eval(MidB$(strLine, 17, LenB(strLine) - 27)))
73:                            Exit Do
                        Case PPC_ELSE
                            '#Else means everything else failed and we must use this block
75:                            intRet = -1
76:                            Exit Do
                        Case PPC_ENDIF
77:                            intCount = intCount - 1
                            
79:                            If intCount = 0 Then
                                'Nothing left to the #If; we're done
81:                                intRet = 1
82:                                Exit Do
83:                            End If
84:                    End Select
85:                End If
86:            End If
87:        Loop
88:    Loop
    
    'Continue beyond this point only if there is a code block to include
91:    If intRet = -1 Then
        'Make sure we don't pass the file boundaries
93:        Do Until EOF(intRead)
            'Read line and trim whitespace
95:            Line Input #intRead, strLine
96:            strLine = TrueTrim(strLine)
            
            'Skip null lines
99:            If LenB(strLine) Then
                'Preproc command?
101:                If AscW(strLine) = intChar Then
                    'If so, check the type; now we must parse all of them
                    'because this is code we want to include in the script
    
                    'Another note is that we shouldn't parse any preproc commands
                    'if they are inside other blocks - hence the reason for the
                    'intRet checks
                    Select Case PreProcCmd(strLine, intChar)
                        Case PPC_LIBRARY
108:                            If intRet Then
109:                                ParseLibrary strLine, objSC
110:                            End If
                        Case PPC_INCLUDE
111:                            If intRet Then
112:                                ParseInclude intRead, intWrite, intChar, strLine, objSC
113:                            End If
                        Case PPC_CONST
114:                            If intRet Then
115:                                ParseConst intWrite, strLine, objSC
116:                            End If
                        Case PPC_IF
117:                            If intRet Then
118:                                ParseIf intRead, intWrite, intChar, strLine, objSC
119:                            End If
                        Case PPC_ELSEIF, PPC_ELSE
                            'If we've found an #ElseIf or #Else, that means we have
                            'to trim the rest of the #If/#End If block out before
                            'finishing up
123:                            intRet = 0
                        Case PPC_ENDIF
                            'Found the end of the block; exit out of the loop
125:                            Exit Do
126:                    End Select
127:                Else
                    'Only add code to script if we are in the block we want to keep
129:                    If intRet = -1 Then
130:                        Print #intWrite, strLine
131:                    End If
132:                End If
133:            Else
134:                If intRet = -1 Then
135:                    Print #intWrite, ""
136:                End If
137:            End If
138:        Loop
139:    End If

141:    Exit Sub
    
143:
Err:
144:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.ParseIf(" & intRead & ", " & intWrite & ", " & intChar & ", """ & strExp & """, )"
End Sub

Private Sub Preproc(ByVal intRead As Integer, _
                    ByVal intWrite As Integer, _
                    ByVal intChar As Integer, _
                    ByRef objSC As ScriptControl)
    '------------------------------------------------------------------
    'Purpose:   Searches for preprocessing commands in code and
    '           then calls the appropriate function to process
    '           any commands it finds
    '
    'Params:
    '           intRead:        File handle on input for the script code
    '           intWrite:       File handle on file to output parsed
    '                           code to
    '           intChar:        Character which denotes a preprocessor
    '                           command (#/@)
    '           objSC:          Script's control object
    '
    'Returns:
    '           Copy of strString without trailing/leading whitespace
    '------------------------------------------------------------------
    
21:    Dim strLine         As String
    
23:    On Error GoTo Err
    
    'Do until the end of the file is reached
26:    Do Until EOF(intRead)
        'Read line and trim whitespace
28:        Line Input #intRead, strLine
29:        strLine = TrueTrim(strLine)
        
        'Skip empty lines
32:        If LenB(strLine) Then
            'Is it a preproc command?
34:            If AscW(strLine) = intChar Then
                'Find out!
                Select Case PreProcCmd(strLine, intChar)
                    Case PPC_INCLUDE
36:                        ParseInclude intRead, intWrite, intChar, strLine, objSC
                    Case PPC_LIBRARY
37:                        ParseLibrary strLine, objSC
                    Case PPC_CONST
38:                        ParseConst intWrite, strLine, objSC
                    Case PPC_IF
39:                        ParseIf intRead, intWrite, intChar, strLine, objSC
                    Case PPC_ELSE, PPC_ELSEIF, PPC_ENDIF
                        'Orphaned statements it appears; just ignore 'em
41:                End Select
42:            Else
                'Just a regular line of code then; write to document
44:                Print #intWrite, strLine
45:            End If
46:        Else
47:            Print #intWrite, ""
48:        End If
49:    Loop
    
51:    Exit Sub
    
53:
Err:
    'Just to keep Error.txt clean because of
    '<Line Input> bug if last line of file is empty
    '"Input past end of file"
'    If (LCase(strLine) = "end sub") Or (LCase(strLine) = "end function") Or (strLine = "") Then _
'        Exit Sub
59:    If Err.Number = 62 Then Exit Sub
    
61:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.Preproc(" & intRead & ", " & intWrite & ", " & intChar & ", """ & strLine & """, )"
End Sub

Private Function PreProcCmd(ByRef strLine As String, _
                            ByVal intChar As Integer) As Integer
    '------------------------------------------------------------------
    'Purpose:   To check if the line starting with intChar is a valid
    '           preproc command, and if it is, to confirm that it has
    '           the proper form
    '
    'Params:
    '           strLine:        Line of code to evaluate to check if
    '                           it is a valid preprocessor command
    '           intChar:        Character which denotes a preproc command
    '
    'Returns:
    '           0 if it is not a valid command, otherwise it returns
    '           a unique numerical code denoting the type of command
    '           that we have
    '------------------------------------------------------------------
    
18:    Dim strTemp     As String
19:    Dim lngPos      As Long
    
21:    On Error GoTo Err
    
    'Set to all lower case
24:    strTemp = MidB$(LCase$(strLine), 3)
    
    'Pretty straight forward...#Else and #End If
    'are the only 2 commands thus far without any parameters
    'so check them first
29:    If strTemp = "end if" Then
30:        PreProcCmd = PPC_ENDIF
31:    ElseIf strTemp = "else" Then
32:        PreProcCmd = PPC_ELSE
33:    Else
        'OK, not either of those, extract first word
35:        lngPos = InStrB(1, strTemp, " ")
        
        'Make sure there is a word to extract
38:        If lngPos Then
            Select Case LeftB$(strTemp, lngPos - 1)
                Case "const"
39:                    PreProcCmd = PPC_CONST
                Case "if"
                    '#If has to end in " Then" to be valid
41:                    If RightB$(strTemp, 10) = " then" Then
42:                        PreProcCmd = PPC_IF
43:                    End If
                Case "elseif"
                    '#ElseIf has to end in " Then" to be valid
45:                    If RightB$(strTemp, 10) = " then" Then
46:                        PreProcCmd = PPC_ELSEIF
47:                    End If
                Case "include"
                    '#Include must have "" surrounding the path
49:                    If AscW(MidB$(strTemp, lngPos + 2)) = CHR_DQUOTE Then
50:                        If AscW(RightB$(strTemp, 2)) = CHR_DQUOTE Then
51:                            PreProcCmd = PPC_INCLUDE
52:                        End If
53:                    End If
                Case "library"
                    '#Library must have "" surrounding the name
55:                    If AscW(MidB$(strTemp, lngPos + 2)) = CHR_DQUOTE Then
56:                        If AscW(RightB$(strTemp, 2)) = CHR_DQUOTE Then
57:                            PreProcCmd = PPC_INCLUDE
58:                        End If
59:                    End If
60:            End Select
61:        End If
62:    End If
    
    'Confirm that the command we found
65:    If intChar = CHR_SHARP Then
66:        If (PreProcCmd And PPC_VBSCRIPT) = 0 Then
67:            PreProcCmd = 0
68:        End If
69:    Else
70:        If (PreProcCmd And PPC_JSCRIPT) = 0 Then
71:            PreProcCmd = 0
72:        End If
73:    End If
    
75:    Exit Function
    
77:
Err:
78:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.PreProcCmd(""" & strLine & """, " & intChar & ")"
End Function

Private Sub ParseConst(ByVal intWrite As Integer, _
                       ByRef strLine As String, _
                       ByRef objSC As ScriptControl)
    '------------------------------------------------------------------
    'Purpose:   To parse a preproc constant and add it to the script's
    '           control object to allow expressions to use it (#If,If,
    '           etc)
    '
    '           Format:
    '               #Const <name> = <value>
    '               #Const MYCONST = "this is a constant!!!"
    '               #Const MYNUM = 453
    '
    'Params:
    '           intWrite:       File handle on file to output parsed
    '                           code to
    '           strLine:        Code to extract constant from
    '           objSC:          Script's control object
    '------------------------------------------------------------------
    
20:    On Error GoTo Err
    
    'Add blank line to script to make line numbers more accurate
23:    Print #intWrite, ""
    
    'Create constant
26:    objSC.ExecuteStatement MidB$(strLine, 3)
    
28:    Exit Sub
    
30:
Err:
31:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.ParseConst(" & intWrite & ", """ & strLine & """, )"
End Sub

Private Sub ParseInclude(ByVal intRead As Integer, _
                         ByVal intWrite, _
                         ByVal intChar As Integer, _
                         ByRef strLine As String, _
                         ByRef objSC As ScriptControl)
    '------------------------------------------------------------------
    'Purpose:   To parse an #Include statement, which includes code
    '           from another file and inserts it into the script.
    '           Since the included code might contain preproc
    '           commands as well, we must start the process again
    '           for this as well (call to Preproc)
    '
    '           Base directory is the DDCH installation folder
    '
    '           Format:
    '               #Include "<path_to_file>"
    '               #Include "\Scripts\Includes\header.vbs"
    '
    'Params:
    '           intRead:        File handle on input for the script code
    '           intWrite:       File handle on file to output parsed
    '                           code to
    '           intChar:        Character which denotes a preproc command
    '           strLine:        Code to extract constant from
    '           objSC:          Script's control object
    '------------------------------------------------------------------
    
27:    Dim intIR       As Integer
    
29:    On Error GoTo Err
    
    'Open the file for reading
32:    intIR = FreeFile
33:    Open G_APPPATH & "\" & MidB$(strLine, 21, LenB(strLine) - 23) For Binary Access Read Lock Read Write As intIR
    
    'Begin preproc on the external code if any
36:    Preproc intIR, intWrite, intChar, objSC
    
    'Close file
39:    Close intIR
    
41:    Exit Sub
    
43:
Err:
44:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.ParseInclude(" & intRead & ", " & intWrite & ", " & intChar & ", """ & strLine & """, )"
End Sub

Private Sub ParseLibrary(ByRef strLib As String, _
                         ByRef objSC As ScriptControl)
'
End Sub

Private Sub LvwAddItem(ByVal intIndex As Integer, _
                       ByVal strName As String, _
                       ByVal strLanguage As String)
3:    On Error GoTo Err
4:    Dim lvwItem As Variant
5:    Set lvwItem = frmHub.lvwScripts.ListItems.Add(intIndex, intIndex & "s", strName)

7:    lvwItem.SubItems(1) = "Inactive"
8:    lvwItem.SubItems(2) = strLanguage
9:    lvwItem.SubItems(3) = Now

11:    Set lvwItem = Nothing
    
13:   Exit Sub
14:
Err:
15:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.LvwAddItem(" & intIndex & ", " & strName & ", " & strLanguage & ")"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmScript = Nothing
End Sub
