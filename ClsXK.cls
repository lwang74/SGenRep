VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsXK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Each demention stores a XueKe. One teacher stores in one item.
'ClsTeacher is structure of one item.
Dim m_all() As Collection, m_Cnt As Integer

Public Property Get XKColl(ByVal Index As Integer) As Collection
    If Index < 1 Or Index > m_Cnt Then
        Set XKColl = Nothing
    Else
        Set XKColl = m_all(Index)
    End If
End Property

Public Property Get Count() As Integer
    Count = m_Cnt
End Property

Public Sub AddXKCls(CTeachr As ClsTeacher)
Dim i As Integer, bFd As Boolean
    'Search for this kerCheng was found.
    For i = 1 To UBound(m_all)
        If CTeachr.KerChen = m_all(i)(1).KerChen Then
            m_all(i).Add CTeachr
            bFd = True
            Exit For
        End If
    Next i
    
    If Not bFd Then
        m_Cnt = m_Cnt + 1
        ReDim Preserve m_all(m_Cnt) As Collection
        Set m_all(m_Cnt) = New Collection
        m_all(m_Cnt).Add CTeachr
    End If
End Sub

Private Sub Class_Initialize()
    ReDim m_all(0)
End Sub
