VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsInSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_WSheet As Excel.Worksheet
Dim m_StarCell As Excel.Range
Dim m_CntOneDay As Integer '每天几节课

Private Sub Class_Initialize()
    m_CntOneDay = 7
End Sub

Public Property Let CntOneDay(cnt)
    m_CntOneDay = cnt
End Property

Public Property Set SetSheet(RSheet As Excel.Worksheet)
    Set m_WSheet = RSheet
End Property

Public Property Set StartCell(StrCell As Excel.Range)
    Set m_StarCell = StrCell
End Property

'***** Public methods *****
'Read all input sheet, covert then into two classes.
Public Sub DoCircleForXueKe(Cxk As ClsXK)
Dim tmRg As Excel.Range
Dim tmKerChen As Excel.Range, wk As Integer, KerJie As Integer, ClassNum As Integer
Dim tmTName As Excel.Range
Dim TpCls As ClsClass, i As Integer, cnt As Integer
Dim TcherCnt As Integer, Cther As ClsTeacher

    'check this sheet weekly 7 or 8
    Set tmRg = m_StarCell.Offset(, 1)
    If tmRg.Value = 8 Then
        m_CntOneDay = 8
    Else
        m_CntOneDay = 7
    End If
    
    Set tmRg = m_StarCell.Offset(4).Offset(, 3)
    Set tmKerChen = m_StarCell.Offset(4)
    Set tmTName = tmKerChen.Offset(, 2)
    cnt = tmKerChen.MergeArea.Count
    TcherCnt = 1
    
    Do Until Trim(tmKerChen.Value) = ""
        Set Cther = New ClsTeacher
        Cther.CntOneDay = m_CntOneDay
        Cther.TeacherSeq = TcherCnt
        Cther.KerChen = tmKerChen.Value
        Cther.Classes = tmTName.Offset(, -1).Value
        Cther.Teacher = tmTName.Value
        For wk = 0 To 4
            For KerJie = 0 To m_CntOneDay - 1
                Cther.AddClass wk, KerJie, tmRg.Offset(, KerJie + wk * m_CntOneDay).Value
            Next KerJie
        Next wk
        Cxk.AddXKCls Cther
        TcherCnt = TcherCnt + 1
        Set tmRg = tmRg.Offset(1)
        Set tmTName = tmTName.Offset(1)
        cnt = cnt - 1
        If cnt = 0 Then
            Set tmKerChen = tmKerChen.Offset(1)
            cnt = tmKerChen.MergeArea.Count
        End If
    Loop
End Sub

'Read all input sheet, covert then into two classes.
Public Sub DoCircle(Cshs As Collection, CThrs As Collection)
Dim tmRg As Excel.Range
Dim tmKerChen As Excel.Range, wk As Integer, KerJie As Integer, ClassNum As Integer
Dim tmTName As Excel.Range
Dim TpCls As ClsClass, i As Integer, cnt As Integer
Dim TcherCnt As Integer, Cther As ClsTeacher

    Set tmRg = m_StarCell.Offset(4).Offset(, 3)
    Set tmKerChen = m_StarCell.Offset(4)
    Set tmTName = tmKerChen.Offset(, 2)
    cnt = tmKerChen.MergeArea.Count
    TcherCnt = 1
    
    Do Until Trim(tmKerChen.Value) = ""
        Set Cther = New ClsTeacher
        Cther.CntOneDay = m_CntOneDay
        CThrs.Add Cther, Str(TcherCnt)
        Cther.TeacherSeq = TcherCnt
        Cther.KerChen = tmKerChen.Value
        Cther.Classes = tmTName.Offset(, -1).Value
        Cther.Teacher = tmTName.Value
        For wk = 0 To 4
            For KerJie = 0 To m_CntOneDay - 1
                ClassVal Cshs, tmRg.Offset(, KerJie + wk * m_CntOneDay).Value, tmKerChen, wk, KerJie
                Cther.AddClass wk, KerJie, tmRg.Offset(, KerJie + wk * m_CntOneDay).Value
            Next KerJie
        Next wk
        TcherCnt = TcherCnt + 1
        Set tmRg = tmRg.Offset(1)
        Set tmTName = tmTName.Offset(1)
        cnt = cnt - 1
        If cnt = 0 Then
            Set tmKerChen = tmKerChen.Offset(1)
            cnt = tmKerChen.MergeArea.Count
        End If
    Loop
End Sub

Private Sub ClassVal(Cshs As Collection, ByVal ClsVal As String, KerChen As Excel.Range, ByVal wk As Integer, ByVal KerJie As Integer)
Dim Pos As Integer, Pos1 As Integer, Pos2 As Integer
Dim ClsValInt1 As Integer, ClsValInt2 As Integer, i As Integer

    Pos = InStr(1, ClsVal, "~", vbTextCompare)
    If Pos <> 0 Then
        ClsValInt1 = Val(Left(ClsVal, Pos - 1))
        ClsValInt2 = Val(Mid(ClsVal, Pos + 1))
        For i = ClsValInt1 To ClsValInt2
            SettoCls Cshs, i, KerChen, wk, KerJie
        Next i
        Exit Sub
    End If
    
    Pos = InStr(1, ClsVal, ".", vbTextCompare)
    Pos2 = InStr(1, ClsVal, ",", vbTextCompare)
    If Pos2 <> 0 And Pos2 < Pos Then Pos = Pos2
    If Pos <> 0 Then
        Do While Pos <> 0
            ClsValInt1 = Val(Mid(ClsVal, Pos1 + 1, Pos - Pos1 - 1))
            SettoCls Cshs, ClsValInt1, KerChen, wk, KerJie
            Pos1 = Pos
            Pos = InStr(Pos1 + 1, ClsVal, ".", vbTextCompare)
            Pos2 = InStr(Pos1 + 1, ClsVal, ",", vbTextCompare)
            If Pos2 <> 0 And Pos2 < Pos Then Pos = Pos2
        Loop
        ClsValInt1 = Val(Mid(ClsVal, Pos1 + 1))
        SettoCls Cshs, ClsValInt1, KerChen, wk, KerJie
        Exit Sub
    End If
    
    SettoCls Cshs, Val(ClsVal), KerChen, wk, KerJie
End Sub

Private Sub SettoCls(Cshs As Collection, ByVal ClsNum As Integer, KerChen As Excel.Range, ByVal wk As Integer, ByVal KerJie As Integer)
Dim TpCls As ClsClass, i As Integer
Dim cc As ClsClass
    
    If ClsNum > 0 Then
        If ClsNum > Cshs.Count Then
            For i = Cshs.Count + 1 To ClsNum
                Set cc = New ClsClass
                cc.CntOneDay = m_CntOneDay
                Cshs.Add cc, Str(i)
                Set TpCls = Cshs(i)
                TpCls.ClassNum = i
            Next i
        End If
        Set TpCls = Cshs(ClsNum)
        TpCls.AddKC wk, KerJie, KerChen.Value
    End If
End Sub


