Attribute VB_Name = "mGlobal"
Option Explicit

'Compiler condition
'Debug mode - It means that it will print messages to the VB IDE debug window
'             Still is executed, even if compiled, so it is CPU friendly if
'             it is turned off when compiling
#Const DEBUG_MODE = False
'API calls
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
  
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
 
'API Move Form ///////////////////////////////////////////////////////////////////////
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
   
'API Stuff
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long


Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(32) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(32) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private m_lngTaskbarMsg         As Long
Private m_lngPrevProc           As Long

'Log messages to the user
Public Sub AddLog(ByRef strMsg As String)
1:   On Error GoTo Err
2:   frmHub.txtLog.Text = frmHub.txtLog.Text & "[" & Now & "] " & strMsg & vbNewLine
3:   frmHub.txtLog.SelStart = Len(frmHub.txtLog.Text)
4:   Exit Sub
5:
Err:
6:   HandleError Err.Number, Err.Description, Erl & "|" & "mGlobal.AddLog()"
End Sub

Public Sub PrintDebug(ByVal strData As String)
1:    On Error GoTo Err

#If DEBUG_MODE Then
4:    Debug.Print strData
#End If
    
7:    Exit Sub
8:
Err:
9:    HandleError Err.Number, Err.Description, Erl & "|mGlobal.PrintDebug()"
End Sub

Public Sub HandleError(ByRef lngNumber As Long, ByRef strDescription As String, ByRef strMethod As String, Optional ByRef lngDLLError As Long)
1:    Dim strError As String
2:    On Error GoTo Err
      'Error log format :
      'Date-Time|Method|Number|DLLError|Description|Version|Beta Version|SVN Version

      'Prevent error number '0' from being logged
7:    If lngNumber Then
         'Add beta version if it is a beta
9:       strError = UTCDate & "|" & strMethod & "|" & lngNumber & "|" & lngDLLError & "|" & strDescription & "|" & vbVersion & "|" & vbBeta & "|" & vbSVNVersion & "|"
10:      If Not G_GUI_IN_UNLOAD Then
            'Add to system log
12:          AddLog "Error: " & UTCDate & "|" & strMethod & "|" & lngNumber & "|" & lngDLLError & "|" & strDescription
             'Show pupop notification ..
14:          If g_objSettings.PopUpCoreError Then
15:                  g_objFunctions.ShowBallon "PT DC Hub v" & vbVersion & " - Core Error", _
                                               g_objSettings.HubName & vbNewLine & _
                                               "Method: " & strMethod & vbNewLine & _
                                               "Description: " & strDescription, 2, True
19:          End If
             'Raise script event
21:          frmHub.SEvent_CoreError strError
22:       End If
          'Print to error log
24:       Print #G_ERRORFILE, strError
          'Print to Debug window if in debug mode
#If DEBUG_MODE Then
27:       Debug.Print strError
#End If
29:    End If

31:    Exit Sub
32:
Err:
#If DEBUG_MODE Then
34:    Debug.Print Now & "|mGlobal.HandleError()|" & Err.Number & "|" & Err.Description & "|" & Err.LastDllError
#End If
    
37:    Err.Clear
38:    Resume Next
End Sub
Public Function UTCDate(Optional ByVal strRef As String) As Date
1:    Dim k As TIME_ZONE_INFORMATION
2:      On Error GoTo Err
      'Get time zone difference
4:    GetTimeZoneInformation k
    
      'If a date is specified, then use that one, else use current
7:    If LenB(strRef) Then _
         UTCDate = DateAdd("n", k.Bias, CDate(strRef)) _
      Else _
        UTCDate = DateAdd("n", k.Bias, Now)
11:    Exit Function
12:
Err:
13:  HandleError Err.Number, Err.Description, Erl & "|mGlobal.UTCDate()"
End Function

Public Function IIfLng(ByVal Expression As Boolean, ByRef TruePart As Long, ByRef FalsePart As Long) As Long
1:    If Expression Then IIfLng = TruePart Else IIfLng = FalsePart
End Function

Public Function XMLUnescape(ByRef strData As String) As String
1:    On Error GoTo Err
    
3:    Dim lngPos As Long
    
5:    XMLUnescape = strData
    
7:    If LenB(XMLUnescape) Then
8:        lngPos = InStrB(1, XMLUnescape, "&")
        
        'If there is a & in the string, that is where we should start searching
11:        If lngPos Then
            'Make sure there is a semi colon, telling us there may be escape sequences
13:            If InStrB(lngPos, XMLUnescape, ";") Then
                'Escape various illegal characters
15:                If InStrB(lngPos, XMLUnescape, "&lt;") Then XMLUnescape = Replace(XMLUnescape, "&lt;", "<")
16:                If InStrB(lngPos, XMLUnescape, "&gt;") Then XMLUnescape = Replace(XMLUnescape, "&gt;", ">")
17:                If InStrB(lngPos, XMLUnescape, "&quot;") Then XMLUnescape = Replace(XMLUnescape, "&quot;", """")
18:                If InStrB(lngPos, XMLUnescape, "&apos;") Then XMLUnescape = Replace(XMLUnescape, "&apos;", "'")
19:                If InStrB(lngPos, XMLUnescape, "&amp;") Then XMLUnescape = Replace(XMLUnescape, "&amp;", "&")
20:            End If
21:        End If
22:    End If
    
24:    Exit Function
    
26:
Err:
27:    HandleError Err.Number, Err.Description, Erl & "|" & "mGlobal.XMLUnescape(" & strData & ")"
End Function

Public Function XMLEscape(ByRef strData As String) As String
1:    On Error GoTo Err
    
3:    XMLEscape = strData
    
    'Check for the illegal characters
6:    If InStrB(1, XMLEscape, "&") Then XMLEscape = Replace(XMLEscape, "&", "&amp;")
7:    If InStrB(1, XMLEscape, "<") Then XMLEscape = Replace(XMLEscape, "<", "&lt;")
8:    If InStrB(1, XMLEscape, ">") Then XMLEscape = Replace(XMLEscape, ">", "&gt;")
9:    If InStrB(1, XMLEscape, """") Then XMLEscape = Replace(XMLEscape, """", "&quot;")
10:   If InStrB(1, XMLEscape, "'") Then XMLEscape = Replace(XMLEscape, "'", "&apos;")

12:    Exit Function
    
14:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "mGlobal.XMLEscape(" & strData & ")"
End Function

Public Function GetByte(ByVal lngData As Long) As Byte
1:    GetByte = CByte(lngData And 255)
End Function

Public Function DebugUser(ByRef objUser As clsUser) As String
1:    If ObjPtr(objUser) Then DebugUser = "[""" & objUser.sName & """," & objUser.bOperator & "," & objUser.iWinsockIndex & ",""" & objUser.Supports & """,""" & objUser.sMyInfoString & """]"
End Function

Public Sub SetTaskbarMsg(ByVal lngProc As Long, ByVal lngMsg As Long)
1:    m_lngTaskbarMsg = lngMsg
2:    m_lngPrevProc = lngProc
End Sub

Public Function GenTempFile() As String
1:    Do
2:        Randomize GetTickCount
3:        GenTempFile = G_APPPATH & "\T" & GetTickCount & Rnd & ".tmp"
4:    Loop While g_objFileAccess.FileExists(GenTempFile)
End Function

Public Function TrueTrim(ByRef strString As String) As String
    '------------------------------------------------------------------
    'Purpose:   To trim any kind of whitespace from the beginning and
    '           and then end of a string. Whitespace includes spaces,
    '           tabs, carriage returns and line feeds
    '
    'Params:
    '           strString:      String to remove leading and trailing
    '                           whitespace from
    '
    'Returns:
    '           Copy of strString without trailing/leading whitespace
    '------------------------------------------------------------------

14:    Dim arr_intChars()      As Integer
15:    Dim i                   As Long
16:    Dim lngStart            As Long
17:    Dim lngEnd              As Long
18:       On Error GoTo Err
    'Get length of string
20:    lngEnd = Len(strString) - 1

    'Make sure there is something to trim
23:    If lngEnd >= 0 Then
        'Set start to first character
25:        lngStart = 1

        'Open character array on string
28:        OpenChrArr arr_intChars, strString

        'Find position of first non-whitespace character
31:        For i = 0 To lngEnd
            Select Case arr_intChars(i)
                Case CHR_SPACE, CHR_TAB, CHR_LF, CHR_CR
32:                    lngStart = lngStart + 1
                Case Else
33:                    Exit For
34:            End Select
35:        Next

        'Find position of last non-whitespace character
38:        For i = lngEnd To lngStart Step -1
            Select Case arr_intChars(i)
                Case CHR_SPACE, CHR_TAB, CHR_LF, CHR_CR
39:                    lngEnd = lngEnd - 1
                Case Else
40:                    Exit For
41:            End Select
42:        Next

        'Close character array
45:        CloseChrArr arr_intChars

        'Extract trimmed string
48:        TrueTrim = Mid$(strString, lngStart, lngEnd - lngStart + 2)
49:    End If

51:    Exit Function

53:
Err:
54:    HandleError Err.Number, Err.Description, Erl & "|mGlobal.TrueTrim()"
End Function

Public Function ValidIP(ByVal strIPAddress As String) As Boolean
   'Function to check if IP is valid --> if it's not higher than 255.255.255.255
2:  Dim sArray As Variant
  
4:  On Error GoTo Err
    
6:  sArray = Split(strIPAddress, ".")
7:  If sArray(0) > 255 Or sArray(1) > 255 Or sArray(2) > 255 Or sArray(3) > 255 Then
8:     ValidIP = False
9:  Else
10:    ValidIP = True
11: End If

13: Exit Function
14:
Err:
15:  ValidIP = False
End Function

Public Function CharCount(str As String, Char As String) As Long
'Get character count in a string
2:    CharCount = UBound(Split(LCase(str), LCase(Char)))
End Function

'Formate digits ex: 12:1:7 for --> 12:01:07
Public Function StrZero(ByVal strValor As String, ByVal bytComprimento As Byte) As String
1:   If Len(strValor) <= bytComprimento Then
2:      StrZero = String(bytComprimento - Len(strValor), "0") & strValor
3:   Else
4:      StrZero = strValor
5:   End If
End Function

'Pause the app without freezing it ('Sleep' freezes the app)
Public Function Pause(HowLong As Long)
1:  Dim Start&
2:  Start = GetTickCount()
3:  Do
4:    DoEvents
5:  Loop Until Start + HowLong < GetTickCount
End Function

Public Function frmMove(frm As Form)
1:  Call ReleaseCapture
2:  Call SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Function

Public Function GetAppVersion() As String
1:   GetAppVersion = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Function

Public Function SQLDate(ByVal sConvertDate As Date) As String
1:   On Error GoTo Err
2:   SQLDate = Format(sConvertDate, "mm/dd/yyyy")
3:   Exit Function
Err:
6:   HandleError Err.Number, Err.Description, Erl & "|" & "mGlobal.SQLDate(" & sConvertDate & ")"
End Function

Public Function SQLHandleQuotes(ByVal sData As String) As String
1:    On Error GoTo Err
2:    Dim i As Integer
    
4:    i = InStr(sData, "'")
    
6:    While i <> 0
7:        sData = Left(sData, i) & "'" & Mid(sData, i + 1)
8:        i = i + 2
9:        i = InStr(i, sData, "'")
10:   Wend
    
12:   SQLHandleQuotes = sData
    
14:   Exit Function
15:
Err:
16:   HandleError Err.Number, Err.Description, Erl & "|" & "mGlobal.SQLHandleQuotes(" & sData & ")"
End Function

Public Function SQLQuotes(ByVal sData As String) As String
1:   On Error GoTo Err
2:   If InStr(1, sData, "'", vbTextCompare) Then
3:      SQLQuotes = "'" & Replace(sData, "'", "''") & "'"
4:   Else
5:      SQLQuotes = "'" & sData & "'"
6:   End If
7:   Exit Function
Err:
9:   HandleError Err.Number, Err.Description, Erl & "|" & "mGlobal.SQLQuotes(" & sData & ")"
End Function

Public Function HubUpTime() As String
1:   Dim iMonths    As Integer
2:   Dim iWeeks     As Integer
3:   Dim iDays      As Integer
4:   Dim iHours     As Integer
5:   Dim lMinutes   As Long
6:   Dim lCurrTime  As Long
7:   Dim sUpTime    As String
     
9:   On Error GoTo Err
   
11:   If Not G_SERVING Then
12:       HubUpTime = "Server off line"
13:       Exit Function
14:   End If
   
16:   lCurrTime = DateDiff("s", CDate(frmHub.ServingDate), DateTime.Now)

18:   lMinutes = (lCurrTime \ 60) Mod 60
19:   iHours = (lCurrTime \ 3600) Mod 24
20:   iDays = (lCurrTime \ 86400) Mod 7
21:   iWeeks = (lCurrTime \ 604800) Mod 4
22:   iMonths = (lCurrTime \ 604800)

24:   If iMonths > 0 Then _
           If iMonths = 1 Then _
                sUpTime = "1 month, " _
           Else sUpTime = iMonths & " months, "
      
29:   If iWeeks > 0 Then _
            If iWeeks = 1 Then _
                 sUpTime = sUpTime & "1 week, " _
            Else sUpTime = sUpTime & iWeeks & " weeks, "

34:   If iDays > 0 Then _
            If iDays = 1 Then _
                 sUpTime = sUpTime & "1 day, " _
            Else sUpTime = sUpTime & iDays & " days, "
        
39:   If iHours > 0 Then _
            If iHours = 1 Then _
                 sUpTime = sUpTime & "1 hour, " _
            Else sUpTime = sUpTime & iHours & " hours, "
        
44:   If lMinutes = 0 Or lMinutes = 1 Then _
           sUpTime = sUpTime & "1 minute" _
      Else sUpTime = sUpTime & lMinutes & " minutes"

48:   HubUpTime = sUpTime
      
50:   Exit Function
    
52:
Err:
53:    HandleError Err.Number, Err.Description, Erl & "|" & "mGlobal.HubUpTime()"
End Function
