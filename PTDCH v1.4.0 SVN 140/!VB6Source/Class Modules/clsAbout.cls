VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum enVersion
    [All] = 0
    [0.x.x] = 1
    [1.0.x] = 2
    [1.1.x] = 3
    [1.2.x] = 4
    [1.3.x] = 5
    [1.4.x] = 6
End Enum

Private Const m_cVAll           As String = "All Versions"
Private Const m_cV0xx           As String = "PTDCH 0.x.x"
Private Const m_cV10x           As String = "PTDCH 1.0.x"
Private Const m_cV11x           As String = "PTDCH 1.1.x"
Private Const m_cV12x           As String = "PTDCH 1.2.x"
Private Const m_cV13x           As String = "PTDCH 1.3.x"
Private Const m_cV14x           As String = "PTDCH 1.4.x"

Private Const mCountVer         As Integer = 6

Private m_strVersion(mCountVer) As String
Private m_strAbout              As String

Private m_cmbChangeLog          As ComboBox
Private m_txtAbout              As TextBox
Private m_txtChangeLog          As TextBox

Private Sub Class_Initialize()
1:    Set m_txtAbout = frmHub.txtAbout
2:    Set m_txtChangeLog = frmHub.txtChangeLog
3:    Set m_cmbChangeLog = frmHub.cmbChangeLog
4:    BuildVersions
End Sub

Private Sub Class_Terminate()
1:    Set m_txtAbout = Nothing
2:    Set m_txtChangeLog = Nothing
3:    Set m_cmbChangeLog = Nothing
End Sub

Private Sub BuildVersions()
1:    On Error GoTo Err
2:    Dim strLine(100) As String
3:    Dim i As Integer, Y As Integer
4:    Const sLine As String = _
       "----------------------------------------------------"
    
7:     strLine(i) = sLine
8:     i = (i + 1): strLine(i) = "** PT Direct Connect Hub V." & vbVersion & " - SVN " & vbSVNVersion & " **"
9:     i = (i + 1): strLine(i) = sLine
10:    i = (i + 1): strLine(i) = "PTDCH is a server-software for the Direct Connect P2P Network."
11:    i = (i + 1): strLine(i) = "-Created by fLaSh"
12:    i = (i + 1): strLine(i) = "-Programmed in MS Visual Basic 6"
13:    i = (i + 1): strLine(i) = "-100% open source"
14:    i = (i + 1): strLine(i) = "-Licenced under GPL"
15:    i = (i + 1): strLine(i) = "-Based in DDCH"
16:    i = (i + 1): strLine(i) = sLine
17:    i = (i + 1): strLine(i) = "** About Developer **"
18:    i = (i + 1): strLine(i) = sLine
19:    i = (i + 1): strLine(i) = "-Name: Carlos Ferreira"
20:    i = (i + 1): strLine(i) = "-E-mail: Carlosferreiracarlos@hotmail.com"
21:    i = (i + 1): strLine(i) = "-Phone: 966 506 396"
22:    i = (i + 1): strLine(i) = "-Home: Braga, S. Victor - Portugal"
23:    i = (i + 1): strLine(i) = sLine
24:    i = (i + 1): strLine(i) = "** Thanks **"
25:    i = (i + 1): strLine(i) = sLine
26:    i = (i + 1): strLine(i) = "-And my respect to developers of SDCH/DDCH"
27:    i = (i + 1): strLine(i) = "(TasMan, The Left Hand, ButterflySoul, HaArD and Selyb)."
28:    i = (i + 1): strLine(i) = "-All people of the DDCH forum"
29:    i = (i + 1): strLine(i) = "-To TheNOP for the rendered help about NMDC protocol, various info about code of DDCH and TortoiseSVN Hosting"
30:    i = (i + 1): strLine(i) = "-Goes out to the former project admin (SDCH/DDCH) who showed much compassion for the project to continue to move forward, though under a new name."
31:    i = (i + 1): strLine(i) = sLine
32:    i = (i + 1): strLine(i) = "** Librarys Credits **"
33:    i = (i + 1): strLine(i) = sLine
34:    i = (i + 1): strLine(i) = "-libbz2.dll:"
35:    i = (i + 1): strLine(i) = "Unknown (Sorry)"
36:    i = (i + 1): strLine(i) = "-MyIPTools.dll:"
37:    i = (i + 1): strLine(i) = "by JDommi"
38:    i = (i + 1): strLine(i) = "jdommi@t-online.de"
39:    i = (i + 1): strLine(i) = "-SciLexer.dll"
40:    i = (i + 1): strLine(i) = "by Neil Hodgson"
41:    i = (i + 1): strLine(i) = "neilh@scintilla.org"
42:    i = (i + 1): strLine(i) = "-SQLite3VB.dll"
43:    i = (i + 1): strLine(i) = "by Todd Tanner"
44:    i = (i + 1): strLine(i) = "toddjt78@msn.com"
45:    i = (i + 1): strLine(i) = "-zlib.dll"
46:    i = (i + 1): strLine(i) = "by Mark Adler"
47:    i = (i + 1): strLine(i) = "madler@alumni.caltech.edu"
48:    i = (i + 1): strLine(i) = sLine
49:    i = (i + 1): strLine(i) = "Release: " & vbReleaseDate
50:    i = (i + 1): strLine(i) = "HomePage: " & "http://HublistChecker.pt.vu/"
51:    i = (i + 1): strLine(i) = "Forum: " & "http://HublistChecker.pt.vu/"
52:    i = (i + 1): strLine(i) = sLine
53:    i = (i + 1): strLine(i) = "Regards,"
54:    i = (i + 1): strLine(i) = "fLaSh - Carlos D.F."
55:    i = (i + 1): strLine(i) = "Braga (S.Victor) - Portugal"
56:    i = (i + 1): strLine(i) = sLine

58:    For Y = 0 To i
59:        m_strAbout = m_strAbout & strLine(Y) & vbNewLine
60:    Next

62:    m_strAbout = Mid(m_strAbout, 1, Len(m_strAbout) - 2)
       
64:    m_strVersion(0) = _
    "PTDCH v.0.x.x " & vbNewLine & _
    "Initial release, only for friends"
                      
68:    m_strVersion(1) = _
    "PTDCH v.1.0.0 " & vbNewLine & _
    "Initial publication of the project"
                      
72:    m_strVersion(2) = _
    "PTDCH v.1.1.0 " & vbNewLine & _
    "**Fixed some bugs in all code" & vbNewLine & _
    "**Add Data Base explorer GUI" & vbNewLine & _
    "**Add Interface for multiples chat rooms" & vbNewLine & _
    "**Add various protocol statistics" & vbNewLine & _
    "**Add new interface for Scheduler (Plan) "
                      
80:    m_strVersion(3) = _
    "PTDCH v.1.2.0 " & vbNewLine & _
    "**Add TabControl for Scripts" & vbNewLine & _
    "**Add Enabled resize various objects in Statistics" & vbNewLine & _
    "**Add new interface for Plugins with all API Hub events "

86:    m_strVersion(4) = _
    "PTDCH v.1.3.0 " & vbNewLine & _
    "**Add descriptions for Bans IP Perms and Temps (reason, nick,etc..)" & vbNewLine & _
    "**Updated Hub events Bans" & vbNewLine & _
    "**Fixed start PTDCH at windows starting" & vbNewLine & _
    "**Add load Empty XML, Highliths, userdb.mdb from Resource" & vbNewLine & _
    "**Add supports for SQLite by API (SQLiteVB.dll)"
                      
94:    m_strVersion(5) = _
    "PTDCH v.1.4.0 " & vbNewLine & _
    "**Updated Hub events Bans" & vbNewLine & _
    "**Removed some OCX (DataGrid, Ado and RicthTextBox)" & vbNewLine & _
    "**Add new about box, and change log box " & vbNewLine & _
    "**Add new Hub Event (Core Error) " & vbNewLine & _
    "**Add timer API for plugins and scripts " & vbNewLine & _
    "**Add new ToolTipText XP " & vbNewLine & _
    "**Fixed smmal bug in chat room interface " & vbNewLine & _
    "**Fixed many optimizations in all code"
    
101:    m_strVersion(6) = m_strVersion(0) & vbTwoLine & m_strVersion(1) & vbTwoLine & m_strVersion(2) & vbTwoLine & m_strVersion(3) & vbTwoLine & m_strVersion(4) & vbTwoLine & m_strVersion(5)

103:    m_txtAbout.Text = m_strAbout

105:    m_cmbChangeLog.AddItem m_cVAll
106:    m_cmbChangeLog.AddItem m_cV0xx
107:    m_cmbChangeLog.AddItem m_cV10x
108:    m_cmbChangeLog.AddItem m_cV11x
109:    m_cmbChangeLog.AddItem m_cV12x
110:    m_cmbChangeLog.AddItem m_cV13x
111:    m_cmbChangeLog.AddItem m_cV14x
 
113:    m_cmbChangeLog.Text = m_cmbChangeLog.List(0)

115:    Exit Sub
116:
Err:
117:    HandleError Err.Number, Err.Description, Erl & "|" & "clsAbout.BuildVersions()"
End Sub

Public Sub SetVersion(ByRef intType As enVersion)
1:   On Error GoTo Err

     Select Case intType
        Case [All]: m_txtChangeLog.Text = m_strVersion(6)
        Case [0.x.x]: m_txtChangeLog.Text = m_strVersion(0)
        Case [1.0.x]: m_txtChangeLog.Text = m_strVersion(1)
        Case [1.1.x]: m_txtChangeLog.Text = m_strVersion(2)
        Case [1.2.x]: m_txtChangeLog.Text = m_strVersion(3)
        Case [1.3.x]: m_txtChangeLog.Text = m_strVersion(4)
        Case [1.4.x]: m_txtChangeLog.Text = m_strVersion(5)
3:    End Select
    
5:    Exit Sub
6:
Err:
7:    HandleError Err.Number, Err.Description, Erl & "|" & "clsAbout.GetVersion(" & intType & ")"
End Sub
