VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLiteData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------
'Purpose:   This class is used by clsSQLite to create a similar
'           process to Recordset
'           Function that executes a query and returns the result as
'           a multidimensional array
'           First row of array specifies column names, each row
'           thereafter contains data
'           Returns a single-row, single-column array
'
'    'Read value and the column names
'    '
'    '  varRS(a,b)
'    '      a = rows
'    '      b = columns
'------------------------------------------------------------------
    
Private m_varData   As Variant
Private m_lngIndex  As Long

Public Sub SetData(ByVal varData As Variant)
1:    On Error GoTo Err
2:    m_varData = varData
3:    m_lngIndex = 1
4:    Exit Sub
Err:
End Sub

Public Sub SetIndex(ByVal lngIndex As Long)
1:    On Error GoTo Err
2:    If Not m_lngIndex >= CollunsCount And Not m_lngIndex = 1 Then
3:        m_lngIndex = lngIndex
4:    End If
5:    Exit Sub
Err:
End Sub

Public Sub MoveNext()
1:    On Error GoTo Err
2:    If Not m_lngIndex >= CollunsCount Then
3:        m_lngIndex = m_lngIndex + 1
4:    End If
5:    Exit Sub
Err:
End Sub

Public Sub MovePrevious()
1:    On Error GoTo Err
2:    If Not m_lngIndex = 1 Then
3:        m_lngIndex = m_lngIndex - 1
4:    End If
5:    Exit Sub
Err:
End Sub

Public Property Get RowsCount() As Long
1:    On Error GoTo Err
2:    RowsCount = (UBound(m_varData, 1) - 1)
3:    Exit Property
Err:
5:    RowsCount = -1
End Property

Public Property Get CollunsCount() As Long
1:    On Error GoTo Err
2:    CollunsCount = UBound(m_varData, 2)
3:    Exit Property
Err:
4:    CollunsCount = -1
End Property

Public Property Get Collun(ByVal lngIndex As Long) As String
1:    On Error GoTo Err
2:    Collun = m_varData(m_lngIndex, (lngIndex + 1))
3:    Exit Property
Err:
5:    Collun = Empty
End Property

Public Property Get Collect(ByVal lngIndex As Long) As String
1:    On Error GoTo Err
2:    Collect = m_varData(m_lngIndex, lngIndex)
3:    Exit Property
Err:
5:    Collect = Empty
End Property

Public Property Get ByIndex(ByVal lngItem As Long, lngCollun As Long) As String
1:    On Error GoTo Err
2:    ByIndex = m_varData(lngItem, lngCollun)
3:    Exit Property
Err:
5:    ByIndex = Empty
End Property
