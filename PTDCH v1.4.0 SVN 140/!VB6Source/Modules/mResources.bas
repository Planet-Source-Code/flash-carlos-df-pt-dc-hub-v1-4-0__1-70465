Attribute VB_Name = "mResources"
Option Explicit

Public Enum enuHighlighter
    VBScript = 0
    JScript = 1
    SQL = 2
End Enum

Public Enum enuXML
    EGLanguage = 0
    EGUsersMessages = 1
End Enum

Private Sub WriteBinary(ByRef strFileName As String, ByRef bytArray() As Byte)
1:   Dim lngFile As Long
2:   On Error GoTo Err

4:   lngFile = FreeFile
    
6:   Open strFileName For Binary Access Write As lngFile 'Open file
7:       Put lngFile, , bytArray  'Write array to file
8:   Close lngFile 'Close file

10:  Exit Sub
11:
Err:
12:  HandleError Err.Number, Err.Description, Erl & "|" & "mResources.WriteBinary(" & strFileName & ")"
End Sub

Public Function LoadImage(ByRef iID As Integer) As IPictureDisp
1:   On Error GoTo Err

3:   Set LoadImage = LoadResPicture(iID, vbResBitmap)

5:   Exit Function
6:
Err:
7:   HandleError Err.Number, Err.Description, Erl & "|" & "mResources.LoadImage(" & iID & ")"
End Function

Public Sub LoadAndSaveHighlighter(ByRef intType As enuHighlighter)
1:    Dim bytArray()  As Byte 'Byte array
2:    Dim lngFile     As Long
3:    Dim strFileName As String
4:    On Error GoTo Err

     'Load the resource into the byte array:
      Select Case intType
        Case enuHighlighter.VBScript
7:            bytArray = LoadResData(101, "Highlighter")
8:            strFileName = G_APPPATH & "\Settings\VB.bin"
        Case enuHighlighter.JScript
9:            bytArray = LoadResData(102, "Highlighter")
10:            strFileName = G_APPPATH & "\Settings\JScripts.bin"
        Case enuHighlighter.SQL
11:           bytArray = LoadResData(103, "Highlighter")
12:           strFileName = G_APPPATH & "\Settings\sql.bin"
13:   End Select

15:   WriteBinary strFileName, bytArray
        
17:   AddLog "Highlighter(" & intType & ") data loaded from resource and saved successfull."
    
19:   Exit Sub
20:
Err:
21: HandleError Err.Number, Err.Description, Erl & "|" & "mResources.LoadHighlighter(" & intType & ")"
End Sub

'IN DEV TEST..
Public Sub LoadAndSaveXML(ByRef intType As enuXML)
1:    Dim bytArray()  As Byte 'Byte array
2:    Dim strFileName As String
3:    On Error GoTo Err
        
      'Load the resource into the byte array:
      Select Case intType
          Case enuXML.EGLanguage
6:            bytArray = LoadResData(101, "XML")
7:            strFileName = G_APPPATH & "\Languages\English.xml"
          Case enuXML.EGUsersMessages
8:            bytArray = LoadResData(102, "XML")
9:            strFileName = G_APPPATH & "\Settings\UsersMessages.xml"
10:   End Select

12:   WriteBinary strFileName, bytArray

14:   AddLog "XML(" & intType & ") data loaded from resource and saved successfull."
    
16:   Exit Sub
17:
Err:
18:   HandleError Err.Number, Err.Description, Erl & "|" & "mResources.LoadAndSaveXML(" & intType & ")"
End Sub

Public Sub LoadAndSaveEmptyDB()
1:    Dim bytArray()  As Byte
2:    Dim strFileName As String
3:    On Error GoTo Err
        
      'Load the resource into the byte array:
6:    bytArray = LoadResData(101, "DB")
7:    strFileName = G_APPPATH & "\DBs\userdb.mdb"

9:    WriteBinary strFileName, bytArray
      
11:   Exit Sub
12:
Err:
13:   HandleError Err.Number, Err.Description, Erl & "|" & "mResources.LoadAndSaveEmptyDB()"
End Sub

Public Sub LoadAndSaveScriptHelp()
1:    Dim bytArray()  As Byte
2:    Dim strFileName As String
3:    On Error GoTo Err
        
      'Load the resource into the byte array:
6:    bytArray = LoadResData(101, "TXT")
7:    strFileName = G_APPPATH & "\Settings\ScriptHelp.vbs"

9:    WriteBinary strFileName, bytArray
      
11:   Exit Sub
12:
Err:
13:   HandleError Err.Number, Err.Description, Erl & "|" & "mResources.LoadAndSaveEmptyDB()"
End Sub
