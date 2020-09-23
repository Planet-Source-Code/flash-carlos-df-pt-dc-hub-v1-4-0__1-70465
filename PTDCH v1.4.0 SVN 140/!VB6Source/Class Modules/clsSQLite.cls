VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'SQLite API Interface

'SQL Lite dll declarations:
Private Declare Sub SQLite3_open Lib "SQLite3VB.dll" Alias "sqlite3_open" (ByVal filename As String, ByRef Handle As Long)
Private Declare Sub sqlite3_close Lib "SQLite3VB.dll" (ByVal DB_Handle As Long)
Private Declare Function sqlite3_last_insert_rowid Lib "SQLite3VB.dll" (ByVal DB_Handle As Long) As Long
Private Declare Function sqlite3_changes Lib "SQLite3VB.dll" (ByVal DB_Handle As Long) As Long
Private Declare Function sqlite_get_table Lib "SQLite3VB.dll" (ByVal DB_Handle As Long, ByVal SQLString As String, ByRef ErrStr As String) As Variant()
Private Declare Function sqlite_libversion Lib "SQLite3VB.dll" () As String ' Now returns a BSTR
'This function returns the number of rows from the last sql statement. Use this to ensure you have a valid array
Private Declare Function number_of_rows_from_last_call Lib "SQLite3VB.dll" () As Long

'Private variables
Private mDBPath             As String
Private mIsConnected        As Boolean
Private mHandle             As Long
Private mLastError          As String

Private mSQLiteData         As clsSQLiteData

Private Sub Class_Terminate()
1:    Call CloseDB
End Sub

Public Property Get Version() As String
1:    On Error Resume Next
2:    Version = sqlite_libversion
End Property

Public Property Let DBPath(ByVal sPath As String)
1:    On Error Resume Next
2:    If g_objFileAccess.FileExists(sPath) Then
3:         mDBPath = sPath
4:    End If
End Property
Public Property Get DBPath() As String
1:    On Error Resume Next
2:    DBPath = mDBPath
End Property

Public Property Get IsConnected() As Boolean
1:    IsConnected = mIsConnected
End Property

Public Property Get Handle() As Long
1:    Handle = mHandle
End Property

Public Property Get LastCallNumRows() As Long
1:    On Error Resume Next
2:    LastCallNumRows = number_of_rows_from_last_call
End Property

Public Property Get LastInsertRowID() As Long
1:    On Error Resume Next
2:    LastInsertRowID = sqlite3_last_insert_rowid(mHandle)
End Property

Public Property Get LastError() As Variant
1:    LastError = mLastError
End Property

Public Function OpenDB() As Boolean
1:    On Error GoTo Err
      
      'Check if alredy opened
4:    If mIsConnected Then Call CloseDB

      'Open the data base API
7:    SQLite3_open mDBPath, mHandle

9:    mIsConnected = True
10:   OpenDB = True
    
12:   Exit Function
13:
Err:
14:   OpenDB = False
End Function

Public Function CloseDB() As Boolean
1:    On Error GoTo Err

       'Check if alredy opened
4:     If mIsConnected Then sqlite3_close mHandle
    
6:     mIsConnected = False
7:     CloseDB = True
    
9:     Exit Function
10:
Err:
11:    CloseDB = False
End Function

Public Function NewData() As clsSQLiteData
    '------------------------------------------------------------------
    'Purpose:   This Function set object link to class clsSQLiteData to
    '           create a similar process to Recordset
    'Params:    None
    '
    'Returns:   New clsSQLiteData object
    '------------------------------------------------------------------
    'Exemple to use in VBScripts:
    '------------------------------------------------------------------
    '    Dim objRS
    '    Dim i
    
    '    Set objRS = mSQLite.NewData
    
    '    objRS.SetData mSQLite.Execute("SELECT * FROM TableTest ")
    
    '    For i = 0 To objRS.RowsCount
    '         MsgBox objRS.Item(0) & vbNewLine & _
    '                objRS.Item(1) & vbNewLine & _
    '                objRS.Item(2) & vbNewLine & _
    '                objRS.Item(3) & vbNewLine & _
    '                objRS.Item(4) & vbNewLine & _
    '                objRS.Item(5) & vbNewLine & _
    '                objRS.Item(6), , "Item " & i
    '         objRS.MoveNext
    '    Next
    '
    '    Set objRS = Nothing
    '------------------------------------------------------------------
1:     On Error Resume Next
2:     Set NewData = New clsSQLiteData
End Function

Public Function ExecuteToData(ByVal sQueryString As String) As clsSQLiteData
    '------------------------------------------------------------------
    'Purpose:   Function that executes a query and returns the result as
    '           a multidimensional array
    '           First row of array specifies column names, each row thereafter contains data
    '           Returns a single-row, single-column array
    '
    'Params:    sQueryString: sql query string
    '           lReturnedRows: return afected rows
    '
    'Returns:   multidimensional array, afected rows
    '
    '   Called by Plugins or Scripts
    '------------------------------------------------------------------

32:    On Error GoTo Hell
       Dim varData As Variant
       
34:    If Not mIsConnected Then Exit Function
       
       'Clear last error
37:    mLastError = Empty

       'Execute command..
40:    varData = sqlite_get_table(mHandle, sQueryString, mLastError)
    
42:    If Len(mLastError) = 0 Then
43:        GoTo Hell
       Else
           Set mSQLiteData = New clsSQLiteData
           mSQLiteData.SetData varData
           ExecuteToData = mSQLiteData
44:    End If
    
50:    Exit Function
51:
Hell:
52:    'ReDim ExecuteToData(0, 0)
End Function

Public Function Execute(ByVal sQueryString As String, Optional ByRef lReturnedRows As Long = -1) As Variant()
    '------------------------------------------------------------------
    'Purpose:   Function that executes a query and returns the result as
    '           a multidimensional array
    '           First row of array specifies column names, each row thereafter contains data
    '           Returns a single-row, single-column array
    '
    'Params:    sQueryString: sql query string
    '           lReturnedRows: return afected rows
    '
    'Returns:   multidimensional array, afected rows
    '
    '   Called by Plugins or Scripts
    '------------------------------------------------------------------
    '    Exemple to use:
    '
    '    Dim varRS   As Variant
    '    Dim strTemp As String
    '    Dim i       As Integer
    '
    '    varRS = Execute("SELECT UserName, Password FROM TableTest ")
    '
    '    'Read value and the column names
    '    '
    '    '  varRS(a,b)
    '    '      a = rows
    '    '      b = columns
    '    '
    '    For i = 0 To UBound(varRS, 2)
    '         strTemp = strTemp & varRS(0, i)
    '    Next
    ''------------------------------------------------------------------
32:    On Error GoTo Hell

34:    If Not mIsConnected Then Exit Function
    
       'Clear last error
37:    mLastError = Empty

       'Execute command..
40:    Execute = sqlite_get_table(mHandle, sQueryString, mLastError)
    
42:    If Not Len(mLastError) = 0 Then
43:        GoTo Hell
44:    End If
    
46:    If Not lReturnedRows = -1 Then
47:        lReturnedRows = UBound(Execute)
48:    End If
    
50:    Exit Function
51:
Hell:
52:    ReDim Execute(0, 0)
End Function

Public Function RunAction(ByVal sQueryString As String) As Boolean
    '------------------------------------------------------------------
    'Purpose:   Subroutine that executes a query without returning a data (faster)
    '
    'Params:    sQueryString: sql query string
    '
    'Returns:   Boolean only : True = No error : False = Error
    '
    '   Called by Plugins or Scripts
    '------------------------------------------------------------------

11:   On Error GoTo Err
12:   If Not mIsConnected Then Exit Function
    
     'Clear last error
15:   mLastError = Empty

     'Execute command..
18:   sqlite_get_table mHandle, sQueryString, mLastError
    
20:   RunAction = True
21:   Exit Function
22:
Err:
24:   RunAction = False
End Function

Public Function TableExists(ByVal strTable As String) As Boolean
    '------------------------------------------------------------------
    'Purpose:   Check if one table exist
    '
    'Params:    TableExists: the table
    '
    'Returns:   Boolean only : True = No error : False = Error
    '
    '   Called by Plugins or Scripts
    '------------------------------------------------------------------
    
11:    Dim strSQL   As String
12:    Dim lngRows  As Long
13:    Dim varRS    As Variant
    
15:    On Error GoTo Err
    
17:    If Len(strTable) = 0 Then
18:        GoTo Err
19:    End If
    
21:    strSQL = "SELECT name FROM SQLITE_MASTER " & _
                "WHERE TYPE = 'table' AND NAME = '" & strTable & "'"
    
       'Clear last error
25:    mLastError = Empty
    
27:    varRS = sqlite_get_table(mHandle, strSQL, mLastError)
    
29:    If Len(mLastError) = 0 Then
30:        lngRows = number_of_rows_from_last_call()
31:    End If
    
33:    If lngRows > 0 Then
34:        TableExists = True
35:    Else
36:        GoTo Err
37:    End If
    
39:    Exit Function
40:
Err:
41:    TableExists = False
End Function

Public Function TableIsEmpty(ByVal strTable As String) As Boolean
    '------------------------------------------------------------------
    'Purpose:   Check if one table is empty
    '
    'Params:    strTable: the table
    '
    'Returns:   Boolean only : True = No error : False = Error
    '
    '   Called by Plugins or Scripts
    '------------------------------------------------------------------
    
11:    Dim strSQL   As String
12:    Dim varRS    As Variant
    
14:    On Error GoTo Err
    
16:    strSQL = "SELECT (SELECT ROWID FROM '" & strTable & "' limit 1) IS NOT NULL"
    
       'Clear last error
19:    mLastError = Empty
    
21:    varRS = sqlite_get_table(mHandle, strSQL, mLastError)
    
23:    If varRS(1, 0) = 1 Then
24:       GoTo Err
25:    Else
26:       TableIsEmpty = False
27:    End If
    
29:    Exit Function
30:
Err:
31:    TableIsEmpty = True
End Function

Public Function AlterDB(ByVal strCommand As String, Optional bShowPopUp As Boolean = False) As Variant
    '------------------------------------------------------------------
    'Purpose:   Get the data base status by command
    '
    'Params:    strCommand: sql command string : bShowPopUp: Shop PopUp Ballon
    '
    'Returns:   Data base base status
    '
    '   Called by Plugins or Scripts
    '------------------------------------------------------------------
    
11:    Dim lngChanges           As Long
12:    Dim lngLastInsertRowID   As Long
13:    Dim varRS(1 To 3)        As Variant
    
15:    On Error GoTo Hell
    
       'Clear last error
18:    mLastError = Empty
    
20:    sqlite_get_table mHandle, strCommand, mLastError
    
22:    If Len(mLastError) = 0 Then
23:        lngChanges = sqlite3_changes(mHandle)
24:        lngLastInsertRowID = sqlite3_last_insert_rowid(mHandle)
25:    End If
    
27:    varRS(1) = lngChanges
28:    varRS(2) = lngLastInsertRowID
29:    varRS(3) = mLastError
    
31:    AlterDB = varRS
    
33:    If bShowPopUp Then
34:       g_objFunctions.ShowBallon "SQLite - AlterDB", _
                    "DB Path: " & mDBPath & vbCrLf & _
                    "Error: " & mLastError & vbCrLf & vbCrLf & _
                    "Handle: " & mHandle & vbCrLf & _
                    "Changes: " & lngLastInsertRowID & _
                    "Last Error: " & mLastError, 0, True
40:    End If

42:    Exit Function
43:
Hell:
44:    On Error Resume Next
45:    varRS(1) = lngChanges
46:    varRS(2) = lngLastInsertRowID
47:    varRS(3) = mLastError

49:    AlterDB = varRS
End Function
