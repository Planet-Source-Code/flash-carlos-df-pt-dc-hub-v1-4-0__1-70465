VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsYHighlighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private hlCount As Long
Private HCount As Integer
Private CurHigh As Integer
Private Const ciIncriment As Integer = 15000
Private lOffset As Long

Private sBuffer As String

Private Sub ReInit()
1: sBuffer = ""
2: lOffset = 0
End Sub

Private Function GetString() As String
1: GetString = Left$(sBuffer, lOffset)
2: sBuffer = "" 'reset
End Function

'This function lets you assign a string to the concating buffer.
Private Sub SetString(ByRef Source As String)
1: sBuffer = Source & String$(ciIncriment, 0)
End Sub

Private Function FindHighlighter(strLangName As String) As Integer
1:  On Error GoTo Err
  
3:  Dim i As Integer
  
5:   For i = 0 To UBound(g_arrHighlighters) - 1
6:    If UCase(g_arrHighlighters(i).strName) = UCase(strLangName) Then
7:      FindHighlighter = i
8:      Exit Function
9:    End If
10:  Next i
  
12:  Exit Function

14:
Err:
15:  HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.FindHighlighter()"
End Function

Private Function GetstrExtension(sFileName As String) As String
1:    On Error GoTo Err
    
3:  Dim lPos As Long
4:  lPos = InStrRev(sFileName, ".")
    
6:  If lPos = 0 Then
7:      GetstrExtension = " "
8:  Else
9:      GetstrExtension = LCase$(Right$(sFileName, Len(sFileName) - lPos))
10: End If
    
12: Exit Function
    
14:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.GetstrExtension()"
End Function

Public Function SetHighlighter(cScintilla As clsYScintilla, _
                               strHighlighter As String)
2:  On Error GoTo Err
  
4:   Dim i As Long, X As Integer
  
6:  X = FindHighlighter(strHighlighter)
7:  cScintilla.StyleClearALL
8:  cScintilla.StartStyle
  
10:  For i = 0 To 127
11:    cScintilla.SetStyleBold i, g_arrHighlighters(X).StyleBold(i)
12:    cScintilla.SetStyleItalic i, g_arrHighlighters(X).StyleItalic(i)
13:    cScintilla.SetStyleUnderline i, g_arrHighlighters(X).StyleUnderline(i)
14:    cScintilla.SetStyleVisible i, g_arrHighlighters(X).StyleVisible(i)
15:    If g_arrHighlighters(X).StyleFont(i) <> "" Then cScintilla.SetStyleFont i, g_arrHighlighters(X).StyleFont(i)
16:    cScintilla.SetStyleFore i, g_arrHighlighters(X).StyleFore(i)
17:    cScintilla.SetStyleBack i, g_arrHighlighters(X).StyleBack(i)
18:    cScintilla.SetStyleSize i, g_arrHighlighters(X).StyleSize(i)
19:    cScintilla.SetStyleEOLFilled i, g_arrHighlighters(X).StyleEOLFilled(i)
20:  Next i
  
22:  For i = 0 To 7
23:    If g_arrHighlighters(X).Keywords(i) <> "" Then cScintilla.SetKeywords i, g_arrHighlighters(X).Keywords(i)
24:  Next i
  
26:  cScintilla.SetLexer g_arrHighlighters(X).iLang
27:  cScintilla.Colourise
28:  CurHigh = X
  
30:  Exit Function

32:
Err:
33:  HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.SetHighlighter()"
End Function

Public Sub SetHighlighterBasedOnExt(cScintilla As clsYScintilla, _
                                    strFileName As String)
2:   On Error GoTo Err
3:   Dim strExtension As String, strClrExt As String, X As Long
  
5:   strExtension = LCase$(Mid$(strFileName, InStrRev(strFileName, ".") + 1, Len(strFileName) - InStrRev(strFileName, ".")))
  
7:   For X = 0 To UBound(g_arrHighlighters)
8:      If InStr(1, g_arrHighlighters(X).strFilter, strExtension) Then
9:          On Error Resume Next
10:          SetHighlighter cScintilla, g_arrHighlighters(X).strName
11:     End If
12:  Next X

14:  Exit Sub
15:
Err:
16:  HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.SetHighlighterBasedOnExt(" & strFileName & ")"
End Sub

Private Function LoadHighlighter(strFile As String)
1:  On Error GoTo Err
  
3:   Dim fFile As Integer
4:   fFile = FreeFile
  
6:   ReDim Preserve g_arrHighlighters(0 To HCount + 1)
  
8:   Open strFile For Binary Access Read As #fFile
9:      Get #fFile, , g_arrHighlighters(HCount)
10:      g_arrHighlighters(HCount).strName = strFile
11:  Close #fFile
  
13:  FreeFile fFile
14:  HCount = HCount + 1
  
16:  Exit Function

18:
Err:
19:    HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.LoadHighlighter()"
End Function

Public Sub LoadDirectory(strDir As String)
1:  On Error GoTo Err
  
3:   Dim str As String, i As Long
4:   hlCount = 0
  
6:   If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
7:   str = Dir(strDir & "\*bin")
8:   Erase g_arrHighlighters
9:   HCount = 0
  
11:  Do Until str = ""
12:    hlCount = hlCount + 1
13:    LoadHighlighter strDir & "\" & str
14:    str = Dir
15:  Loop
  
17:  Exit Sub

19:
Err:
20:  HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.LoadDirectory()"
End Sub
