Attribute VB_Name = "mGlobal"
Option Explicit

'------------------------------------------------------------------
'Purpose:   Objects Links to Core and objects of the plugin global
'
'Params:    Depends of the objects
'
'Returns:   Depends of the objects
'------------------------------------------------------------------

'Set link objects for all forms
Public frmHub           As Form
Public g_colForms       As Object

'Set link objects for all classes
Public g_objChatRoom    As Object
Public g_objFunctions   As Object
'
Public g_colIPBans      As Object
Public g_colUsers       As Object
Public g_colCommands    As Object
Public g_colMessages    As Object
Public g_colLanguages   As Object
Public g_colSWinsocks   As Object
Public g_colSVariables  As Object
'
Public g_objRegistered  As Object
Public g_objSettings    As Object
Public g_objRegExps     As Object
Public g_objScheduler   As Object
Public g_objFileAccess  As Object
Public g_objStatus      As Object

Public g_objHighlighter As Object
Public g_objSCI         As Object
Public g_objComDialog   As Object
Public g_objSQLite      As Object
Public g_objActiveX     As Object
Public g_colDictionary  As Object
Public g_objTimer       As Object
Public g_objTimersCol   As Object

'Public internal plugin objects
Public Events           As cEvents
Public g_Enabled        As Boolean
Public G_PLGPATH        As String
Public G_APPPATH        As String
Public g_Main           As New main

Public Sub HandleError(ByRef lngNumber As Long, ByRef strDescription As String, ByRef strMethod As String, Optional ByRef lngDLLError As Long)
    '------------------------------------------------------------------
    'Purpose:   Transfers the description of the error to core
    '           (It is useful to work in debug mode)
    'Params:    Method|Number|DLLError|Description|Plugin Name|Version|Beta|
    '
    'Returns:   none
    '------------------------------------------------------------------

    On Error GoTo Err

   'Prevent error number '0' from being logged
   If Not lngNumber = 0 Then
        frmHub.RunFunction "HandleError", lngNumber, strDescription, "(Plugin:" & g_Main.Name & ")" & strMethod, lngDLLError
   End If

   Err.Clear

   Exit Sub
Err:
End Sub
