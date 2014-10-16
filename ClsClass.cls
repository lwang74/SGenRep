VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_WSheet As Excel.Worksheet
Dim m_Wk_time(0 To 4, 0 To 7) As String
Dim m_StrCell As Excel.Range
Dim m_ClassNum As Integer 'zero base
Dim m_CntOneDay As Integer '每天几节课
Dim m_Begin As Integer '0 for 7节课; 1 for 8 节课

Private Sub Class_Initialize()
    m_CntOneDay = 7
    m_Begin = 0
'    ReDim Preserve m_Wk_time(0 To 4, 0 To m_CntOneDay - 1) As String
End Sub

Public Property Let CntOneDay(cnt)
    m_CntOneDay = cnt
    If 7 = m_CntOneDay Then
        m_Begin = 1
    End If
'    ReDim m_Wk_time(0 To 4, 0 To m_CntOneDay - 1) As String
End Property

Public Property Get ClassNum() As Integer
    ClassNum = m_ClassNum
End Property

Public Property Let ClassNum(Num As Integer)
    m_ClassNum = Num
End Property

Public Property Set SetSheet(RSheet As Excel.Worksheet)
    Set m_WSheet = RSheet
End Property

Public Property Set StartCell(StrCell As Excel.Range)
    Set m_StrCell = StrCell
End Property

'***** Public methods *****
Public Sub AddKC(ByVal Week As Integer, ByVal KerJie As Integer, ByVal KerChen As String)
    If Week > 4 Or KerJie > m_CntOneDay - 1 Then
        MsgBox "Week over. Week=" & Week & " KerJie=" & KerJie & " Kerchen=" & KerChen
    End If
    m_Wk_time(Week, KerJie + m_Begin) = KerChen
End Sub

Public Sub OutPut()
Dim i As Integer, j As Integer
    If m_ClassNum <> 1 Then
        m_WSheet.Range("A1:I9").Copy m_StrCell
    End If
    
    m_StrCell.Value = Left(m_StrCell.Value, 3) & m_ClassNum & Mid(m_StrCell.Value, 5)
    For i = LBound(m_Wk_time, 1) To UBound(m_Wk_time, 1)
        For j = m_Begin To UBound(m_Wk_time, 2)
            m_StrCell.Offset(i + 3).Offset(, j + 1 - m_Begin).Value = m_Wk_time(i, j)
        Next j
    Next i
End Sub

