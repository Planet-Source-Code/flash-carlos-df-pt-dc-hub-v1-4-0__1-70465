VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPBans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_colTemp   As Collection
Private m_colPerm   As Collection
Private m_objData   As clsIPBansData

Private Sub Class_Initialize()
1:    Set m_colTemp = New Collection
2:    Set m_colPerm = New Collection
End Sub

Private Sub Class_Terminate()
1:    Set m_colTemp = Nothing
2:    Set m_colPerm = Nothing
End Sub

'Ban a new IP
'lngMinutes = -1 = Perm ban
'           > 0  = Temp ban
'
'IsIniHub: used only at LoadSettings (hide popup ballon)
Public Sub Add(ByRef strIP As String, Optional ByRef lngMinutes As Long = -1, Optional strNick As String = Empty, Optional strBannedBy As String = Empty, Optional strReason As String = Empty, Optional IsIniHub As Boolean = False)
1:    On Error GoTo Err
2:    Dim strTemp As String
      
      'Select right ban type
5:    If lngMinutes = -1 Then
6:        If Not CheckPerm(strIP) Then
7:             Set m_objData = New clsIPBansData
8:             m_objData.IP = strIP
9:             m_objData.Nick = strNick
10:            m_objData.BannedBy = strBannedBy
11:            m_objData.Reason = strReason
               '
13:            m_colPerm.Add m_objData, strIP
               '
               'Run script event
16:            frmHub.SEvent_AddedPermBan strIP, strNick, strBannedBy, strReason
17:            Set m_objData = Nothing
               '
19:            If Not IsIniHub Then
20:                 strTemp = "IP Kicked Permanent" & vbNewLine & _
                              "Nick: " & strNick & vbNewLine & _
                              "IP: " & strIP & vbNewLine & _
                              "Banned by: " & strBannedBy & vbNewLine & _
                              "Reason: " & strReason
                    '
                    'Show Ballon notification
27:                 If g_objSettings.PopUpUserKick Then _
                         g_objFunctions.ShowBallon "PT DC Hub " & vbVersion, _
                                                   g_objSettings.HubName & vbNewLine & _
                                                   strTemp, 1, True
                    '
32:                 AddLog strTemp
33:           End If
34:        End If
35:    Else
           'If already banned, then replace old ban time with new
37:        If CheckTemp(strIP) Then
38:            m_colTemp(strIP).ExpDate = DateAdd("n", lngMinutes, CheckDate(m_colTemp(strIP).ExpDate))
39:        Else
40:            Set m_objData = New clsIPBansData
41:            m_objData.IP = strIP
42:            m_objData.ExpDate = DateAdd("n", lngMinutes, Now)
43:            m_objData.Nick = strNick
44:            m_objData.BannedBy = strBannedBy
45:            m_objData.Reason = strReason
               '
47:            m_colTemp.Add m_objData, strIP
               '
               'Run script event
50:            frmHub.SEvent_AddedTempBan strIP, lngMinutes, strNick, strBannedBy, strReason
               '
52:            Set m_objData = Nothing
               '
54:            If Not IsIniHub Then
55:                strTemp = "IP Kicked Temporary" & vbNewLine & _
                             "Nick: " & strNick & vbNewLine & _
                             "IP: " & strIP & vbNewLine & _
                             "Time: " & lngMinutes & " minute(s)" & vbNewLine & _
                             "Banned by: " & strBannedBy & vbNewLine & _
                             "Reason: " & strReason
                   '
                   'Show Ballon notification
63:                If g_objSettings.PopUpUserKick Then _
                        g_objFunctions.ShowBallon "PT DC Hub " & vbVersion, _
                                                  g_objSettings.HubName & vbNewLine & _
                                                  strTemp, 1, True
                   '
68:                AddLog strTemp
69:            End If
70:        End If
71:    End If
    
73:    Exit Sub
    
75:
Err:
76:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.Add(""" & strIP & """, " & lngMinutes & ")"
End Sub

'Remove an IP
'intType = 0 = Both
'        = 1 = Temp
'        = 2 = Perm
Public Sub Remove(ByRef strIP As String, Optional ByRef intType As Integer)
1:    On Error GoTo Err

    'Select which ban lists we should check/remove
      Select Case intType
        Case 0
4:            If CheckTemp(strIP) Then m_colTemp.Remove strIP
5:            If CheckPerm(strIP) Then m_colPerm.Remove strIP
        Case 1
6:            If CheckTemp(strIP) Then m_colTemp.Remove strIP
        Case 2
7:            If CheckPerm(strIP) Then m_colPerm.Remove strIP
8:    End Select

10:    Exit Sub
    
12:
Err:
13:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.Remove(""" & strIP & """)"
End Sub

'Check if the ips are banned
'   0    = Not banned
'   -1   = Perm banned
'   Else = Temp banned (length in minutes)
Public Function Check(ByRef strIP As String) As Long
1:    On Error GoTo Err
    
3:    If CheckPerm(strIP) Then
4:        Check = -1
5:    ElseIf CheckTemp(strIP) Then
        'Check if the date has expired
7:        Check = DateDiff("n", Now, m_colTemp(strIP).ExpDate)
         'If the date is earlier than the current date, the temp ip ban
         'has expired
10:        If Check < 1 Then
11:            m_colTemp.Remove strIP
12:            Check = 0
13:        End If
14:    End If
    
16:    Exit Function
    
18:
Err:
19:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.Check(""" & strIP & """)"
End Function

Private Function CheckPerm(ByRef strIP As String) As Boolean
1:    On Error GoTo Err
    
3:    m_colPerm.Item strIP
4:    CheckPerm = True

6:    Exit Function
7:
Err:
End Function

Private Function CheckTemp(ByRef strIP As String) As Boolean
1:    On Error GoTo Err
    
3:    m_colTemp.Item strIP
4:    CheckTemp = True

6:    Exit Function
7:
Err:
End Function

'Clear temp ban list
Public Sub ClearTemp()
1:    Set m_colTemp = New Collection
End Sub

'Clear perm ban list
Public Sub ClearPerm()
1:    Set m_colPerm = New Collection
End Sub

'Loop through listitems to get a list of temp banned ips (with expiration date)
Public Function TempList() As String
1:    On Error GoTo Err
    
      'Don't loop if there is nothing in the collection
4:    If m_colTemp.count Then
5:        For Each m_objData In m_colTemp
6:            TempList = TempList & "|" & m_objData.IP & "|" & m_objData.ExpDate & "|" & m_objData.Nick & "|" & m_objData.Reason & "|" & m_objData.BannedBy & "|" & vbNewLine
7:        Next
        
9:       Set m_objData = Nothing
10:  End If

12:    Exit Function
    
14:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.TempList()"
End Function

'Loop through listitems to get a list of perm banned ips
Public Function PermList() As String
1:    Dim varLoop     As Variant

3:    On Error GoTo Err
    
      'Don't loop if there is nothing in the collection
6:    If m_colPerm.count Then
7:        For Each varLoop In m_colPerm
8:            PermList = PermList & varLoop & "|"
9:        Next
10:   End If

12:   Exit Function
    
14:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.PermList()"
End Function

Public Property Get TempItems() As Collection
1:    Set TempItems = m_colTemp
End Property

Public Property Get PermItems() As Collection
1:    Set PermItems = m_colPerm
End Property

Private Function CheckDate(ByRef datCheck As Date) As Date
1:    If DateDiff("n", datCheck, Now) Then
2:        CheckDate = Now
3:    Else
4:        CheckDate = datCheck
5:    End If
End Function

