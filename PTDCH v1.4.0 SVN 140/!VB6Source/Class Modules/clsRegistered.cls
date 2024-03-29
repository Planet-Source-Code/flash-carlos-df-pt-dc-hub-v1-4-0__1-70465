VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_objConnection     As Connection
Private m_objRecordSet      As Recordset

Private Sub Class_Initialize()
1:    Set m_objConnection = frmHub.oPermaCon
End Sub

Private Sub Class_Terminate()
1:    Set m_objConnection = Nothing
2:    Set m_objRecordSet = Nothing
End Sub

'Register an user
' Return parameters:
'   0 = No error
'   1 = Registered already
'   2 = Name longer than 40 chars
'   3 = Password longer than 20 chars
'   4 = Reason longer than 100 chars
'   5 = No Name
Public Function Add(ByRef strName As String, ByRef strPassword As String, ByRef intClass As enuClass, Optional ByRef strAdminName As String = "Admin / GUI", Optional ByRef lngMin As Long, Optional blnHide As Boolean = False) As Integer
1:   On Error GoTo Err

     'Make sure the user isn't already registered
4:   If Registered(strName) Then
5:        Add = 1
6:        Exit Function
7:   End If
    
     'Make name isn't too long / empty
     Select Case LenB(strName)
     
        '**************
        Case Is > 80
        '**************
13:            Add = 2
14:            Exit Function
                   
        '**************
        Case 0
        '**************
18:            Add = 5
19:            Exit Function
          
        '**************
        Case Else
        '**************
        
            'If locked, it goes into a different table
            '*******************************************************
24:         If intClass = Locked Then
            '*******************************************************

                Select Case LenB(strPassword)
    
                    '=============
                    Case Is > 100
                    '=============
28:                     Add = 3
29:                     Exit Function
                    
                    '=============
                    Case 0
                    '=============
33:                     m_objConnection.Execute "INSERT INTO UsrClass(UserName, Class) " & _
                                                "VALUES(" & SQLQuotes(strName) & "," & intClass & ");", , 129
                    
36:                     If lngMin Then
37:                         m_objConnection.Execute "INSERT INTO BanNames(UserName, BannedBy, RefDate, Perm) " & _
                                                    "VALUES(" & SQLQuotes(strName) & "," & SQLQuotes(strAdminName) & ",'" & Format$(DateAdd("n", lngMin, Now), "yyyy-mm-dd hh:mm:ss") & "',0);", , 129
39:                     Else
40:                         m_objConnection.Execute "INSERT INTO BanNames(UserName, BannedBy, RefDate, Perm) " & _
                                                    "VALUES(" & SQLQuotes(strName) & "," & SQLQuotes(strAdminName) & ",'" & Format$(Now, "yyyy-mm-dd hh:mm:ss") & "',1);", , 129
42:                     End If

                    '=============
                    Case Else
                    '=============
46:                     m_objConnection.Execute "INSERT INTO UsrClass(UserName, Class) " & _
                                                "VALUES(" & SQLQuotes(strName) & "," & intClass & ");", , 129
                    
49:                     If lngMin Then
50:                         m_objConnection.Execute "INSERT INTO BanNames(UserName, Reason, BannedBy, RefDate, Perm) " & _
                                                    "VALUES(" & SQLQuotes(strName) & "," & SQLQuotes(strPassword) & "," & SQLQuotes(strAdminName) & ",'" & Format$(DateAdd("n", lngMin, Now), "yyyy-mm-dd hh:mm:ss") & "',0);", , 129
52:                     Else
53:                         m_objConnection.Execute "INSERT INTO BanNames(UserName, Reason, BannedBy, RefDate, Perm) " & _
                                                    "VALUES(" & SQLQuotes(strName) & "," & SQLQuotes(strPassword) & "," & SQLQuotes(strAdminName) & ",'" & Format$(Now, "yyyy-mm-dd hh:mm:ss") & "',1);", , 129
55:                     End If
                         
                       'Show pupop notification ..
58:                     If g_objSettings.PopUpUserBaned And Not blnHide Then
59:                          g_objFunctions.ShowBallon "PT DC Hub " & vbVersion, _
                                                    g_objSettings.HubName & vbNewLine & _
                                                    "User " & IIf(lngMin > 0, "Temp", "Perm") & " Banned" & vbNewLine & _
                                                    "Nick: " & strName & vbNewLine & _
                                                    "Reason: " & strPassword & vbNewLine & _
                                                    "Time: " & lngMin & " minute(s)" & vbNewLine & _
                                                    "By: " & strAdminName, 0, True
66:                     End If
                        
                        'Raise script event
69:                     frmHub.SEvent_AddedTempBan Empty, lngMin, strName, strAdminName, strPassword

71:             End Select
                
            '*******************************************************
            Else
            '*******************************************************
            
                'Make sure the password isn't blank / too long
                Select Case LenB(strPassword)
                
                    '=============
                    Case Is > 40
                    '=============
77:                        Add = 3
78:                        Exit Function
                    
                    '=============
                    Case 0
                    '=============
82:                        Add = 4
83:                        Exit Function
                        
                    '=============
                    Case Else
                    '=============
87:                      m_objConnection.Execute "INSERT INTO UsrClass(UserName, Class) " & _
                                                 "VALUES(" & SQLQuotes(strName) & "," & intClass & ");", , 129
89:                      m_objConnection.Execute "INSERT INTO UsrStatic(UserName, Pass, RegedBy, RegDate) " & _
                                                 "VALUES(" & SQLQuotes(strName) & "," & SQLQuotes(strPassword) & "," & SQLQuotes(strAdminName) & ",'" & Format$(Now, "yyyy-mm-dd hh:mm:ss") & "');", , 129
91:                      m_objConnection.Execute "INSERT INTO UsrDynamic(UserName) " & _
                                                 "VALUES(" & SQLQuotes(strName) & ");", , 129
                        
                         'Show pupop notification ..
95:                     If g_objSettings.PopUpNewReg And Not blnHide Then
96:                            g_objFunctions.ShowBallon "PT DC Hub " & vbVersion, _
                                                      g_objSettings.HubName & vbNewLine & _
                                                      "New user registed" & vbNewLine & _
                                                      "Nick: " & strName & vbNewLine & _
                                                      "Password: " & strPassword & vbNewLine & _
                                                      "Class: " & intClass & vbNewLine & _
                                                      "By: " & strAdminName, 0, True
103:                     End If
                         
                         'Raise script event
106:                     frmHub.SEvent_AddedRegisteredUser strName, strPassword, CInt(intClass), strAdminName, lngMin

108:                End Select

110:            End If

112:    End Select
    
114:    Exit Function
    
116:
Err:
117:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegistered.Add(""" & strName & """, """ & strPassword & """, " & intClass & ", """ & strAdminName & """)"
End Function

'Set language for registered
Public Function SetLanguage(ByRef strName As String, ByRef strLanguageID As String) As Integer
1:    On Error GoTo Err
  
3:    If Registered(strName) = Unknown Then SetLanguage = 1: Exit Function
4:    If Not g_objFunctions.ValidLang(strLanguageID) Then SetLanguage = 2: Exit Function
    
6:    m_objConnection.Execute "UPDATE UsrStatic SET UsrStatic.i18n='" & strLanguageID & "' WHERE UsrStatic.UserName=" & SQLQuotes(strName) & ";", , 129
    
    
9:    Exit Function
    
11:
Err:
12:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegistered.SetLanguage(" & strName & ", " & strLanguageID & ")"
End Function

'Remove a user if registered
Public Sub Remove(ByRef strName As String)
1:    On Error GoTo Err

3:    If Registered(strName) = Unknown Then Exit Sub
      
5:    m_objConnection.Execute "DELETE FROM UsrClass WHERE UserName = " & SQLQuotes(strName) & ";", , 129
6:    m_objConnection.Execute "DELETE FROM BanNames WHERE UserName = " & SQLQuotes(strName) & ";", , 129
    
     'Raise script event "RemovedRegisteredUser(strName)"
9:    frmHub.SEvent_RemovedRegisteredUser strName

11:   Exit Sub
  
13:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegistered.Remove(""" & strName & """)"
End Sub
'Check if a user is registered
Public Function Registered(ByRef strName As String) As enuClass
1:    On Error GoTo Err
  
3:    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrClass.Class FROM UsrClass WHERE UsrClass.UserName=" & SQLQuotes(strName) & ";", , 1)
    
      'If not registered, then set the class to "Unknown"
6:    If m_objRecordSet.BOF Then _
         Registered = Unknown _
      Else _
         Registered = m_objRecordSet.Collect(0)
        
11:   If Registered = Locked Then
12:        Set m_objRecordSet = m_objConnection.Execute("SELECT Count(*) FROM BanNames WHERE UserName=" & SQLQuotes(strName) & " And BanNames.Perm = 0 And BanNames.RefDate < NOW();", , 1)
        
14:        If m_objRecordSet.Collect(0) Then
15:            m_objConnection.Execute "Delete * From UsrClass Where UserName=" & SQLQuotes(strName) & ";", , 129
16:            Registered = Unknown
17:       End If
18:   End If
    
20:   Set m_objRecordSet = Nothing
  
22:   Exit Function
  
24:
Err:
25:   HandleError Err.Number, Err.Description, Erl & "|" & "clsRegistered.Registered(""" & strName & """)"
End Function

'Edit a user if registered
Public Function Edit(ByRef strName As String, ByRef strPassword As String, ByRef intClass As enuClass, Optional ByRef strAdminName As String = "Admin / GUI", Optional ByRef lngMin As Long) As Integer
1:   On Error GoTo Err
  
     'Name cannot be longer than 80 characters
4:   If LenB(strName) > 80 Then Edit = 2: Exit Function
  
     'Find out if it's registered
     Select Case Registered(strName)
        Case Unknown
7:            Edit = 1
        Case Locked
            'If the new class is locked, then update reason
            Select Case intClass
                Case Locked, Unknown
                    'Reason can't be longer than 50 characters
                    Select Case LenB(strPassword)
                        Case 0
                        Case Is > 100: Edit = 3
                        Case Else: m_objConnection.Execute "UPDATE BanNames SET Reason=" & SQLQuotes(strPassword) & " WHERE UserName=" & SQLQuotes(strName), 129
10:                 End Select
                    
                    'UPDATE ban time (0 = no change, -1 = change to perm, else = temp ban in minutes)
                    Select Case lngMin
                        Case 0
                        Case -1: m_objConnection.Execute "UPDATE BanNames SET Perm=True WHERE(UserName=" & SQLQuotes(strName) & ");", , 129
                        Case Else: m_objConnection.Execute "UPDATE BanNames SET RefDate='" & Format$(DateAdd("n", lngMin, Now), "yyyy-mm-dd hh:mm:ss") & "', Perm = 0 WHERE UserName=" & SQLQuotes(strName) & ";", , 129
13:                 End Select
                Case Else
                    'If not clean up database by deleting BanNames entry
                    'and by creating new UsrStatic/Dynamic entries
16:                    m_objConnection.Execute "Delete * From BanNames WHERE ((UsrClass.UserName)=" & SQLQuotes(strName) & ");", , 129
                    
                    'Prevent them from adding a too long / empty password
                    Select Case LenB(strPassword)
                        Case Is > 40: Edit = 3
                        Case 0: Edit = 4
                        Case Else
19:                            m_objConnection.Execute "UPDATE UsrClass SET Class=" & intClass & " WHERE UserName=" & SQLQuotes(strName), , 129
20:                            m_objConnection.Execute "INSERT INTO UsrStatic (UserName, Pass, RegedBy, RegDate) Values(" & SQLQuotes(strName) & "," & SQLQuotes(strPassword) & "," & SQLQuotes(strAdminName) & ",'" & Format$(Now, "yyyy-mm-dd hh:mm:ss") & "');", , 129
21:                            m_objConnection.Execute "INSERT INTO UsrDynamic (UserName) VALUES(" & SQLQuotes(strName) & ");", , 129
22:                 End Select
23:            End Select
        Case intClass
            'Class is same, so only password can change
        
            'Password cannot be longer than 40 characters
            'If it is blank, assume no change
            Select Case LenB(strPassword)
                Case Is > 40: Edit = 3
                Case 0
                Case Else: m_objConnection.Execute "UPDATE UsrStatic SET Pass=" & SQLQuotes(strPassword) & " WHERE UserName=" & SQLQuotes(strName), , 129
28:         End Select
        Case Else
            Select Case intClass
                Case Unknown
                    'Password cannot be longer than 40 characters
                    'If it is blank, assume no change
                    Select Case LenB(strPassword)
                        Case Is > 40: Edit = 3
                        Case 0
                        Case Else
31:                            m_objConnection.Execute "UPDATE UsrStatic SET Pass=" & SQLQuotes(strPassword) & " WHERE UserName=" & SQLQuotes(strName), , 129
32:                 End Select
                Case Locked
                    'UPDATE class
34:                    m_objConnection.Execute "UPDATE UsrClass SET Class=" & intClass & " WHERE UserName=" & SQLQuotes(strName), , 129
                    
                    'Delete old entries in regular user tables
37:                    m_objConnection.Execute "DELETE FROM UsrStatic WHERE UsrClass.UserName=" & SQLQuotes(strName) & ";", , 129
38:                    m_objConnection.Execute "DELETE FROM UsrDynamic WHERE UsrClass.UserName=" & SQLQuotes(strName) & ";", , 129
    
                    'Create new entry in BanNames table
                    
                    'Reason can't be longer than 50 characters
                    Select Case LenB(strPassword)
                        Case 0: m_objConnection.Execute "INSERT INTO BanNames (UserName, BannedBy, RefDate, Perm) VALUES(" & SQLQuotes(strName) & "," & SQLQuotes(strAdminName) & ",'" & Format$(Now, "yyyy-mm-dd hh:mm:ss") & "',1);", , 129
                        Case Is > 100: Edit = 3
                        Case Else: m_objConnection.Execute "INSERT INTO BanNames (UserName, Reason, BannedBy, RefDate, Perm) VALUES(" & SQLQuotes(strName) & "," & SQLQuotes(strPassword) & "," & SQLQuotes(strAdminName) & ",'" & Format$(Now, "yyyy-mm-dd hh:mm:ss") & "',1);", , 129
43:                    End Select
                    
                    'UPDATE ban time if applicable
46:                    If lngMin Then m_objConnection.Execute "UPDATE BanNames SET RefDate='" & Format$(DateAdd("n", lngMin, Now), "yyyy-mm-dd hh:mm:ss") & "' Perm = 0 WHERE UserName=" & SQLQuotes(strPassword) & ";", , 129
                Case Else
                    'UPDATE class
48:                    m_objConnection.Execute "UPDATE UsrClass SET Class=" & intClass & " WHERE UserName=" & SQLQuotes(strName), , 129

                    'Password cannot be longer than 40 characters
                    'If it is blank, assume no change
                    Select Case LenB(strPassword)
                        Case Is > 40: Edit = 3
                        Case 0
                        Case Else
52:                            m_objConnection.Execute "UPDATE UsrStatic SET Pass=" & SQLQuotes(strPassword) & " WHERE UserName=" & SQLQuotes(strName), , 129
53:                    End Select
54:            End Select
55:    End Select

57:    Exit Function
  
59:
Err:
60:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegistered.Edit(""" & strName & """, """ & strPassword & """, " & intClass & ", """ & strAdminName & """)"
End Function

'Rename a user if registered
Public Function Rename(ByRef strOldName As String, ByRef strNewName As String) As Integer
1:    On Error GoTo Err

    'Make sure they are registered
4:    If Registered(strOldName) = Unknown Then
5:        Rename = 1
6:    Else
        'Make sure the new name fits
        Select Case LenB(strNewName)
            Case Is > 80: Rename = 2
            Case 0: Rename = 3
            Case Else: m_objConnection.Execute "UPDATE UsrClass SET UserName=" & SQLQuotes(strNewName) & " WHERE UserName=" & SQLQuotes(strOldName), , 129
8:        End Select
9:    End If
  
11:    Exit Function
  
13:
Err:
14:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegistered.Rename(""" & strOldName & """, """ & strNewName & """)"
End Function

'Check if the password is correct - if it is, return their class, otherwise return 0
Public Function Check(ByRef strName As String, ByRef strPassword As String) As enuClass
1:    On Error GoTo Err
  
3:    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrStatic.Pass FROM UsrStatic WHERE UsrStatic.UserName=" & SQLQuotes(strName) & ";", , 1)
  
    'If not registered, then make class equal Unknown
6:    If m_objRecordSet.BOF Then
7:        Check = Unknown
8:    Else
        'If password matches, set the return value to their class
10:        If strPassword = m_objRecordSet.Collect(0) Then Check = Registered(strName) Else Check = Unknown
11:    End If
    
13:    Set m_objRecordSet = Nothing
  
15:    Exit Function
  
17:
Err:
18:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegistered.Check(""" & strName & """, """ & strPassword & """)"
End Function

'Retrieve various information on a user
Public Function GetInfo(ByRef strName As String, Optional ByRef strField As String = "all") As String
1:    Dim objField As Field
  
3:    On Error GoTo Err
  
    Select Case Registered(strName)
        Case Locked
            Select Case LCase$(strField)
                Case "all"
                    'Get all attributes
6:                    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrClass.UserName, BanNames.BannedBy, BanNames.RefDate, BanNames.Reason, BanNames.Perm FROM UsrClass INNER JOIN BanNames ON UsrClass.UserName = BanNames.UserName WHERE UsrClass.UserName = " & SQLQuotes(strName) & ";", , 1)
                
8:                    If Not m_objRecordSet.BOF Then
9:                        For Each objField In m_objRecordSet.Fields
10:                            GetInfo = GetInfo & objField.Value & "|"
11:                        Next
12:                    End If
                Case "password"
                    'Get reason for ban
14:                    Set m_objRecordSet = m_objConnection.Execute("SELECT BanNames.Reason FROM BanNames WHERE BanNames.UserName=" & SQLQuotes(strName) & ";", , 1)
                    
16:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case "classname"
                    'Class will always be locked
18:                    GetInfo = "Locked"
                Case "regdate"
                    'Get the date the user was banned
20:                    Set m_objRecordSet = m_objConnection.Execute("SELECT BanNames.RefDate FROM BanNames WHERE BanNames.UserName=" & SQLQuotes(strName) & ";", , 1)
                    
22:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case "regby", "regedby"
                    'Get who they were banned by
24:                    Set m_objRecordSet = m_objConnection.Execute("SELECT BanNames.BannedBy FROM BanNames WHERE BanNames.UserName=" & SQLQuotes(strName) & ";", , 1)
                    
26:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case "perm"
                    'Get perm status
28:                    Set m_objRecordSet = m_objConnection.Execute("SELECT BanNames.Perm FROM BanNames WHERE BanNames.UserName=" & SQLQuotes(strName) & ";", , 1)
                    
30:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case Else
31:                    GetInfo = "Invalid field."
32:            End Select
        Case Unknown
            'Return nothing b/c the user is not registered
        Case Else
            Select Case LCase$(strField)
                Case "all"
                    'Get all attributes of user
35:                    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrClass.UserName, UsrStatic.Pass, UsrClass.Class, ClassTypes.Name, UsrStatic.i18n, UsrStatic.RegedBy, UsrStatic.RegDate, UsrDynamic.LastLogin, UsrDynamic.LastIP FROM (UsrStatic INNER JOIN (ClassTypes INNER JOIN UsrClass ON ClassTypes.ID = UsrClass.Class) ON UsrStatic.UserName = UsrClass.UserName) INNER JOIN UsrDynamic ON (UsrClass.UserName = UsrDynamic.UserName) AND (UsrStatic.UserName = UsrDynamic.UserName) WHERE UsrClass.UserName=" & SQLQuotes(strName) & ";", , 1)
            
37:                    If Not m_objRecordSet.BOF Then
38:                        For Each objField In m_objRecordSet.Fields
39:                            GetInfo = GetInfo & objField.Value & "|"
40:                        Next
41:                    End If
                Case "language"
                    'Get preferred language of user
43:                    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrStatic.i18n FROM UsrStatic WHERE UsrStatic.UserName=" & SQLQuotes(strName) & ";", , 1)
            
45:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case "classname"
                    'Get name of class
47:                    Set m_objRecordSet = m_objConnection.Execute("SELECT ClassTypes.Name FROM ClassTypes INNER JOIN UsrClass ON ClassTypes.ID = UsrClass.Class WHERE UsrClass.UserName=" & SQLQuotes(strName) & ";", , 1)
            
49:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case "password"
                    'Get user's password
51:                    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrStatic.Pass FROM UsrStatic WHERE UsrStatic.UserName=" & SQLQuotes(strName) & ";", , 1)
            
53:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case "regdate"
                    'Get the date the user was registered
55:                    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrStatic.RegDate FROM UsrStatic WHERE UsrStatic.UserName=" & SQLQuotes(strName) & ";", , 1)
            
57:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case "regby", "regedby"
                    'Get the name of the user who registered this user
59:                    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrStatic.RegedBy FROM UsrStatic WHERE UsrStatic.UserName=" & SQLQuotes(strName) & ";", , 1)
            
61:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case "lastlogin"
                    'Get the date of the user's last login
63:                    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrDynamic.LastLogin FROM UsrDynamic WHERE UsrDynamic.UserName=" & SQLQuotes(strName) & ";", , 1)
            
65:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case "lastip"
                    'Get the IP that the user last logged in with
67:                    Set m_objRecordSet = m_objConnection.Execute("SELECT UsrDynamic.LastIP FROM UsrDynamic WHERE UsrDynamic.UserName=" & SQLQuotes(strName) & ";", , 1)
            
69:                    If Not m_objRecordSet.BOF Then GetInfo = m_objRecordSet.Collect(0)
                Case Else
                    'Invalid field
71:                    GetInfo = "Invalid field."
72:            End Select
73:    End Select
  
75:    Set m_objRecordSet = Nothing
  
77:    Exit Function
  
79:
Err:
    'Ignore null errors; they don't matter
81:    If Not Err.Number = 94 Then _
            HandleError Err.Number, Err.Description, Erl & "|" & "clsRegistered.GetInfo(""" & strName & """, """ & strField & """)"
End Function

'Retrieve a list of users
Public Function GetList(ByRef intSort As Integer, Optional ByRef intClass As enuClass = 0) As String
1:    On Error GoTo Err
  
    'Sort by :
    '0 = Do not sort
    '1 = Class
    '2 = Username
    '3 = Class, then Username
  
    'If intClass = 0 / Unknown, the we are only getting one class
10:    If intClass Then
        'If we are just getting one class, sorted by class = unsorted, and sorted by name = sorted by name/class
        Select Case intSort
            Case 0, 1: Set m_objRecordSet = m_objConnection.Execute("SELECT UsrClass.UserName FROM UsrClass WHERE UsrClass.Class=" & intClass & ";", , 1)
            Case 2, 3: Set m_objRecordSet = m_objConnection.Execute("SELECT UsrClass.UserName FROM UsrClass WHERE UsrClass.Class=" & intClass & " ORDER BY UsrClass.UserName;", , 1)
12:        End Select
    
        'Loop until an error occurs
15:        Do
16:            GetList = GetList & m_objRecordSet.Collect(0) & vbNewLine
17:            m_objRecordSet.MoveNext
18:        Loop
19:    Else
        'Find out in which order we should sort it in
        Select Case intSort
            Case 0: Set m_objRecordSet = m_objConnection.Execute("SELECT UsrClass.UserName, UsrClass.Class FROM UsrClass;")
            Case 1: Set m_objRecordSet = m_objConnection.Execute("SELECT UsrClass.UserName, UsrClass.Class FROM UsrClass ORDER BY UsrClass.Class;", , 1)
            Case 2: Set m_objRecordSet = m_objConnection.Execute("SELECT UsrClass.UserName, UsrClass.Class FROM UsrClass ORDER BY UsrClass.UserName;", , 1)
            Case 3: Set m_objRecordSet = m_objConnection.Execute("SELECT UsrClass.UserName, UsrClass.Class FROM UsrClass ORDER BY UsrClass.Class, UsrClass.UserName;", , 1)
21:        End Select
    
           'Loop until an error occurs
24:        Do
25:            GetList = GetList & m_objRecordSet.Collect(0) & "|" & m_objRecordSet.Collect(1) & vbNewLine
26:            m_objRecordSet.MoveNext
27:        Loop
28:    End If

30:
Err:
31:    Set m_objRecordSet = Nothing
End Function
