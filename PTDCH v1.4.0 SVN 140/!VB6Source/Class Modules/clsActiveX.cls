VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActiveX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

Private Declare Function LoadTypeLib Lib "oleaut32.dll" (pFileName As Byte, pptlib As Object) As Long
Private Declare Function RegisterTypeLib Lib "oleaut32.dll" (ByVal ptlib As Object, szFullPath As Byte, szHelpFile As Byte) As Long
Private Declare Function UnRegisterTypeLib Lib "oleaut32.dll" (libID As GUID, ByVal wVerMajor As Integer, ByVal wVerMinor As Integer, ByVal lCID As Long, ByVal tSysKind As eSYSKIND) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (lpsz As Byte, pclsid As GUID) As Long

Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Enum eSYSKIND
   SYS_WIN16 = 0&
   SYS_WIN32 = 1&
   SYS_MAC = 2&
End Enum

Public Enum RegOp
    Register = 1
    UnRegister = 0
End Enum

Public Function RegisterActiveX(ByRef lzAxDll As String, ByRef mRegOption As RegOp) As Boolean
1:   On Error GoTo Err

3:   Dim mLib As Long, DllProcAddress As Long
4:   Dim mThread
5:   Dim sWait As Long
6:   Dim mExitCode As Long
7:   Dim lpThreadID As Long

9:    mLib = LoadLibrary(lzAxDll)
    
11:    If mLib = 0 Then RegisterActiveX = False: Exit Function

13:    If mRegOption = Register Then
14:        DllProcAddress = GetProcAddress(mLib, "DllRegisterServer")
15:    Else
16:        DllProcAddress = GetProcAddress(mLib, "DllUnregisterServer")
17:    End If
    
19:    If DllProcAddress = 0 Then
20:        RegisterActiveX = False
21:        Exit Function
22:    Else
23:        mThread = CreateThread(ByVal 0, 0, ByVal DllProcAddress, ByVal 0, 0, lpThreadID)
24:        If mThread = 0 Then
25:            FreeLibrary mLib
26:            RegisterActiveX = False
27:            Exit Function
28:        Else
29:            sWait = WaitForSingleObject(mThread, 10000)
30:            If sWait <> 0 Then
31:                FreeLibrary mLib
32:                mExitCode = GetExitCodeThread(mThread, mExitCode)
33:                ExitThread mExitCode
34:                Exit Function
35:            Else
36:                FreeLibrary mLib
37:                CloseHandle mThread
38:            End If
39:        End If
40:    End If
    
42:    RegisterActiveX = True
    
44:   Exit Function
45:
Err:
46:    HandleError Err.Number, Err.Description, Erl & "|" & "clsActiveX.RegisterActiveX()"
End Function

Public Function RegisterTBL(ByRef sLib As String) As Boolean
1:    Dim suLib() As Byte
2:    Dim lErrOK As Long
3:    Dim tlb As Object
4:    On Error GoTo Err

      'Basic automatically translates strings to Unicode Byte arrays
      'but doesn't null-terminate, so you must do it yourself
8:    suLib = sLib & vbNullChar

      'Pass first byte of array
11:   lErrOK = LoadTypeLib(suLib(0), tlb)

12:   DoEvents

13:   If lErrOK = 0 Then
14:        lErrOK = RegisterTypeLib(tlb, suLib(0), 0)
15:   End If

17:   RegisterTBL = CBool(lErrOK = 0)
      
21:   Exit Function

23:
Err:
25:   HandleError Err.Number, Err.Description, Erl & "|" & "clsActiveX.RegisterTypeLib(" & sLib & ")"
End Function

Public Function GetWindowsDir() As String
1:    Dim strBruffer As String
2:    On Error GoTo Err

4:    strBruffer = String(200, Chr$(0))
    
6:    GetWindowsDir = CStr(Left$(strBruffer, GetWindowsDirectory(strBruffer, Len(strBruffer))))

8:    Exit Function

10:
Err:
11:   GetWindowsDir = Empty
12:   HandleError Err.Number, Err.Description, Erl & "|" & "clsActiveX.GetWindowsDir()"
End Function
