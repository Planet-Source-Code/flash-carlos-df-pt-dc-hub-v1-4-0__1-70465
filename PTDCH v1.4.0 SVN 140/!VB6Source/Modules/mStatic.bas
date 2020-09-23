Attribute VB_Name = "mStatic"
Option Explicit

'Constants
Public Const vbSVNVersion       As Integer = 142
Public Const vbVersion          As String = "1.4.0"

#If SVN Then
    Public Const vbBeta         As String = "Debug"
#Else
    Public Const vbBeta         As String = vbNullString
#End If

Public Const vbLock             As String = "EXTENDEDPROTOCOLDEFDEFDEFDEFDEFDEFDEF Pk=PTDCH" & vbVersion & vbSVNVersion & "DEFDEF"
Public Const vbWelcome          As String = "This hub is running V. " & vbVersion & " of the PTDCH produced by fLaSh (UpTime: %[UpTime])|"
Public Const vbChar5            As String = ""
Public Const vbChar160          As String = " "
Public Const vbTwoLine          As String = vbNewLine & vbNewLine
Public Const vbPartialClassList As String = "2 = Mentored" & vbNewLine & "3 = Registered" & vbNewLine & "4 = Invisible" & vbNewLine & "5 = VIP" & vbNewLine & "6 = Operator" & vbNewLine & "7 = Invisible Operator" & vbNewLine & "8 = Super Operator" & vbNewLine & "9 = Invisible Super Operator" & vbNewLine & "10 = Admin" & vbNewLine & "11 = Invisible Admin"
Public Const vbScriptConst      As String = "Const vbVersion = " & vbVersion & ":Const vbBeta = """ & vbBeta & """"
Public Const vbSFC              As Long = 28
Public Const vbReleaseDate      As Date = #1/12/2008 10:47:00 PM#

Public Const CHR_CR             As Integer = 13
Public Const CHR_LF             As Integer = 10
Public Const CHR_TAB            As Integer = 9
Public Const CHR_SPACE          As Integer = 32
Public Const CHR_DQUOTE         As Integer = 34

'-- Script function constant script event boolean array identifiers
Public Const vbSMain                            As Long = 0
Public Const vbSDataArrival                     As Long = 1
Public Const vbSAttemptedConnection             As Long = 2
Public Const vbSUserConnected                   As Long = 3
Public Const vbSRegConnected                    As Long = 4
Public Const vbSOpConnected                     As Long = 5
Public Const vbSUserQuit                        As Long = 6
Public Const vbSStartedServing                  As Long = 7
Public Const vbSAddedRegisteredUser             As Long = 8
Public Const vbSwskScript_Close                 As Long = 9
Public Const vbSwskScript_Connect               As Long = 10
Public Const vbSwskScript_ConnectionRequest     As Long = 11
Public Const vbSwskScript_DataArrival           As Long = 12
Public Const vbSwskScript_Error                 As Long = 13
Public Const vbStmrScriptTimer_Timer            As Long = 14
Public Const vbSAddedPermBan                    As Long = 15
Public Const vbSAddedTempBan                    As Long = 16
Public Const vbSStartedRedirecting              As Long = 17
Public Const vbSStoppedServing                  As Long = 18
Public Const vbSMassMessage                     As Long = 19
Public Const vbSUnloadMain                      As Long = 20
Public Const vbSError                           As Long = 21
Public Const vbSTimeout                         As Long = 22
Public Const vbSRemovedRegisteredUser           As Long = 23
Public Const vbSCustComArrival                  As Long = 24
Public Const vbSPreDataArrival                  As Long = 25
Public Const vbSFailedConf                      As Long = 26
Public Const vbSCoreError                       As Long = 27
Public Const vbStmrAPI_Timer                    As Long = 28

'Flood protection constants
Public Const vbFWMyINFO         As Long = 5
Public Const vbFWGetNickList    As Long = 35
Public Const vbFWActiveSearch   As Long = 35
Public Const vbFWPassiveSearch  As Long = 3
Public Const vbFWMilliseconds   As Long = 10000

'Public variables
Public G_APPPATH                As String
Public G_ERRORFILE              As Integer

Public G_GUI_IN_UNLOAD                As Boolean
Public G_GUI_IS_LOADED                As Boolean
Public G_SERVING                As Boolean

#If SVN Then
    Public G_LOGPATH            As String
#End If

'Public objects
Public g_objSciLexer()          As clsYScintilla
Public g_objFileAccess          As clsFileAccess
Public g_objFunctions           As clsFunctions
Public g_colUsers               As clsHub
Public g_colIPBans              As clsIPBans
Public g_objRegistered          As clsRegistered
Public g_objSettings            As clsSettings
Public g_objStatus              As clsStatus
Public g_colCommands            As clsCommands
Public g_objRegExps             As clsRegExps
Public g_objScheduler           As clsPlan
Public g_objChatRoom            As clsChatRoom
Public g_objHighlighter         As clsYHighlighter
Public g_objActiveX             As clsActiveX
Public g_objAbout               As clsAbout
Public g_colToolTip             As clsDictionary
Public g_colMessages            As clsDictionary

Public g_colSWinsocks           As Collection
Public g_colSVariables          As Collection
Public g_colLanguages           As Collection

'Enums
Public Enum enInterfaceDB
    MsAccess = 0
    MySQL = 1
End Enum

Public Enum enuState
    Disconnected = -1
    Wait_Key = 0
    Wait_Validate = 1
    Wait_Pass = 2
    Wait_PassPM = 3
    Wait_Info = 4
    Logged_In = 5
End Enum

Public Enum enuClass
    Locked = -1
    Unknown = 0
    Normal = 1
    Mentored = 2
    Registered = 3
    Invisible = 4
    Vip = 5
    Op = 6
    InvisibleOp = 7
    SuperOp = 8
    InvisibleSuperOp = 9
    Admin = 10
    InvisibleAdmin = 11
End Enum

Public Enum enuAlert
    MinShare = 0
    FakeTag = 1
    MinSlots = 2
    HSRatio = 3
    BSRatio = 4
    MaxHubs = 5
    DCppversion = 6
    NMDCVersion = 7
    NoTag = 8
    FakeShare = 9
    MaxShare = 10
    MaxSlots = 11
    Socks5 = 12
    PassiveMode = 13
    NoCOClients = 14
End Enum

Public Enum enuOpenFileMode
    vbRandom = 0
    vbInput = 1
    vbOutput = 2
    vbAppend = 3
    vbBinary = 4
End Enum

'Types
Public Type typToolTips
    sMessage    As String
    sTitle      As String
    iStyle      As ToolTipStyleEnum
    iIcon       As ToolTipTypeEnum
End Type
Public g_arrToolTips()      As typToolTips

Public Type typHighlighter
  StyleBold(127)        As Long
  StyleItalic(127)      As Long
  StyleUnderline(127)   As Long
  StyleVisible(127)     As Long
  StyleEOLFilled(127)   As Long
  StyleFore(127)        As Long
  StyleBack(127)        As Long
  StyleSize(127)        As Long
  StyleFont(127)        As String
  StyleName(127)        As String
  Keywords(7)           As String
  strFilter             As String
  strComment            As String
  strName               As String
  iLang                 As Long
End Type
Public g_arrHighlighters()      As typHighlighter

Public Type typPlugin
    Object      As Object
    UseSetup    As Boolean
    UseEvents   As Boolean
    '
    Name        As String
    Version     As String
    Author      As String
    Description As String
    ReleaseDate As Date
    Comments    As String
    '
    Index       As Integer
End Type

Public g_objPlugin()            As typPlugin
Public g_PluginsFound           As Boolean

'Used by plugin interface
Public g_objSCI                 As New clsYScintilla
Public g_objComDialog           As New clsCommonDialog
Public g_objSQLite              As New clsSQLite
Public g_colDictionary          As New clsDictionary
Public g_objTimer               As New clsTimer
Public g_objTimersCol           As New clsTimersCol
