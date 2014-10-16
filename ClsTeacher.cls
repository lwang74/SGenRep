VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_WSheet As Excel.Worksheet
Dim m_Class(0 To 4, 0 To 7) As String
Dim m_StrCell As Excel.Range
Dim m_Teacher As String, m_KerChen As String, m_Classes As String
Dim m_Seq As Integer
Dim m_CntOneDay As Integer '每天几节课 7 or 8
Dim m_Begin As Integer '0 for 7节课; 1 for 8 节课

Private Sub Class_Initialize()
    m_CntOneDay = 7
    m_Begin = 0
'    ReDim Preserve m_Class(0 To 4, 0 To m_CntOneDay - 1) As String
End Sub

Public Property Let CntOneDay(cnt)
    m_CntOneDay = cnt
    If 7 = m_CntOneDay Then
        m_Begin = 1
    End If
'    ReDim m_Class(0 To 4, 0 To m_CntOneDay - 1) As String
End Property

Public Property Set SetSheet(RSheet As Excel.Worksheet)
    Set m_WSheet = RSheet
End Property

Public Property Set StartCell(StrCell As Excel.Range)
    Set m_StrCell = StrCell
End Property

Public Property Let TeacherSeq(Seq As Integer)
    m_Seq = Seq
End Property

Public Property Let Teacher(TName As String)
    m_Teacher = TName
End Property

Public Property Let KerChen(KName As String)
    m_KerChen = KName
End Property

Public Property Get KerChen() As String
    KerChen = m_KerChen
End Property

Public Property Let Classes(CName As String)
    m_Classes = CName
End Property

'***** Public methods *****
Public Sub AddClass(ByVal Week As Integer, ByVal KerJie As Integer, ByVal Class As String)
    If Week > 4 Or KerJie > m_CntOneDay - 1 Then
        MsgBox "Week over. Week=" & Week & " KerJie=" & KerJie & " Class=" & Class
    End If
    m_Class(Week, KerJie + m_Begin) = Class
End Sub

Public Sub OutPut()
Dim i As Integer, j As Integer
    If m_Seq <> 1 Then
        m_WSheet.Range("A1:I9").Copy m_StrCell
    End If
    
    m_StrCell.Offset(, 1).Value = m_Teacher
    m_StrCell.Offset(, 4).Value = m_KerChen
    m_StrCell.Offset(, 7).Value = m_Classes
    For i = LBound(m_Class, 1) To UBound(m_Class, 1)
        For j = m_Begin To UBound(m_Class, 2)
            m_StrCell.Offset(, j + 1 - m_Begin).Offset(i + 3).Value = m_Class(i, j)
        Next j
    Next i
End Sub

Public Sub XKOutPut(StartRg As Excel.Range, ByVal RowOffset As Integer)
Dim tmRg As Excel.Range, i As Integer, j As Integer
    Set tmRg = StartRg.Offset(RowOffset)
    StartRg.Worksheet.Range(StartRg, StartRg.Offset(, 50)).Copy tmRg
    tmRg.Value = m_Classes
    tmRg.Offset(, 1).Value = m_Teacher
    For i = LBound(m_Class, 1) To UBound(m_Class, 1)
        For j = m_Begin To UBound(m_Class, 2)
            tmRg.Offset(, i * 8 + j + 2).Value = m_Class(i, j)
        Next j
    Next i
End Sub

