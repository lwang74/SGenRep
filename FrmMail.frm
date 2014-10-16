VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Generator"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdXK 
      Caption         =   "学科"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   885
      TabIndex        =   19
      Top             =   3000
      Width           =   765
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   150
      Left            =   360
      TabIndex        =   16
      Top             =   2760
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   255
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton BtnBack 
      Caption         =   "&Back"
      Height          =   405
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Width           =   1185
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   5400
      TabIndex        =   5
      Top             =   3000
      Width           =   1185
   End
   Begin VB.CommandButton BtnNext 
      Caption         =   "&Next"
      Height          =   405
      Left            =   4080
      TabIndex        =   4
      Top             =   3000
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2535
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   6495
      Begin VB.OptionButton OptKS 
         Caption         =   "8 课时"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   32
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton OptKS 
         Caption         =   "7 课时"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   31
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox CmbTeacher 
         Height          =   315
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1485
         Width           =   2445
      End
      Begin VB.ComboBox CmbClass 
         Height          =   315
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1050
         Width           =   2445
      End
      Begin VB.ComboBox CmbIn 
         Height          =   315
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   660
         Width           =   2445
      End
      Begin VB.Label Label9 
         Caption         =   "课时数:"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Output Sheet Name(Teacher):"
         Height          =   255
         Left            =   375
         TabIndex        =   14
         Top             =   1530
         Width           =   2445
      End
      Begin VB.Label Label5 
         Caption         =   "Output Sheet Name(Clase):"
         Height          =   255
         Left            =   375
         TabIndex        =   12
         Top             =   1095
         Width           =   2445
      End
      Begin VB.Label Label4 
         Caption         =   "Input Sheet Name:"
         Height          =   255
         Left            =   375
         TabIndex        =   10
         Top             =   705
         Width           =   2445
      End
      Begin VB.Label Label3 
         Caption         =   "Select excel file full path. and click Next button."
         Height          =   945
         Left            =   225
         TabIndex        =   9
         Top             =   240
         Width           =   5685
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   75
      Width           =   6255
      Begin VB.CommandButton BtnOpen 
         Caption         =   "&Open"
         Height          =   375
         Left            =   5220
         TabIndex        =   3
         Top             =   1455
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1605
         TabIndex        =   1
         Top             =   1500
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Select excel file full path. and click Next button. Note: you must open the excel file."
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   225
         TabIndex        =   6
         Top             =   240
         Width           =   5685
      End
      Begin VB.Label Label1 
         Caption         =   "Excel File Path:"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   1515
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2415
      Index           =   3
      Left            =   105
      TabIndex        =   20
      Top             =   1020
      Width           =   6855
      Begin VB.CommandButton CmdBrw 
         Caption         =   "..."
         Height          =   315
         Left            =   5745
         TabIndex        =   29
         Top             =   405
         Width           =   420
      End
      Begin VB.CommandButton CmdDown 
         Caption         =   "Dow&n"
         Height          =   315
         Left            =   3750
         TabIndex        =   28
         Top             =   1500
         Width           =   570
      End
      Begin VB.CommandButton CmdUp 
         Caption         =   "&Up"
         Height          =   315
         Left            =   3735
         TabIndex        =   27
         Top             =   1095
         Width           =   570
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "&Del"
         Height          =   315
         Left            =   3075
         TabIndex        =   26
         Top             =   1500
         Width           =   570
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   3075
         TabIndex        =   25
         Top             =   1095
         Width           =   570
      End
      Begin VB.TextBox TxOtExFl 
         Height          =   315
         Left            =   3090
         TabIndex        =   23
         Top             =   390
         Width           =   2580
      End
      Begin VB.ListBox LstInEx 
         Height          =   1425
         Left            =   120
         TabIndex        =   21
         Top             =   375
         Width           =   2820
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Output Excel files"
         Height          =   255
         Left            =   3135
         TabIndex        =   24
         Top             =   135
         Width           =   2730
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Input Excel files"
         Height          =   255
         Left            =   105
         TabIndex        =   22
         Top             =   165
         Width           =   2730
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Index           =   2
      Left            =   165
      TabIndex        =   17
      Top             =   270
      Width           =   6255
      Begin VB.Label Label10 
         Caption         =   "The Class sheet and Teacher sheet is generated, Please check."
         Height          =   1590
         Left            =   300
         TabIndex        =   18
         Top             =   255
         Width           =   5790
      End
   End
End
Attribute VB_Name = "FrmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrPage As Integer
Dim FilSys As New FileSystemObject
Dim SchXls As Excel.Workbook
Dim InputSheet As Excel.Worksheet
Dim OutClassSheet As Excel.Worksheet
Dim OutTeacherSheet As Excel.Worksheet


'Xue Ke template workbook located in same app path
Dim TmpBook As Excel.Application ' .Workbook
'Xue Ke template workbook, "Sheet1" sheet in which the format is defined.
Dim TmpSheet As Excel.Worksheet
'Xue Ke output book
Dim XKBook As Excel.Workbook
'Xue ke class to store all content sorted by XueKe l-ubound.
Dim Cxks As New ClsXK

Dim CInSht As ClsInSht
Dim ClsGp As New Collection
Dim ClsTc As New Collection

Private Sub Page(ByVal PageInt As Integer)
Dim i As Integer

    Select Case PageInt
    Case 0
        For i = 0 To Me.Frame1.Count - 1
            Me.Frame1(i).BorderStyle = 0
            Me.Frame1(i).Left = 0
            Me.Frame1(i).Top = 0
        Next i
        Me.Frame1(0).Visible = True
        Me.Frame1(1).Visible = False
        Me.Frame1(2).Visible = False
        Me.Frame1(3).Visible = False
        Me.BtnBack.Enabled = False
        Me.BtnNext.Enabled = Me.Text1.Text <> ""
        Me.PrgBar.Visible = False
        Me.CmdXK.Visible = True
    Case 1
        Me.Frame1(0).Visible = False
        Me.Frame1(1).Visible = True
        Me.Frame1(2).Visible = False
        Me.Frame1(3).Visible = False
        Me.BtnBack.Enabled = True
        Me.BtnNext.Enabled = False
        Me.PrgBar.Visible = False
        Me.CmdXK.Visible = False
    Case 2
        Me.Frame1(0).Visible = False
        Me.Frame1(1).Visible = False
        Me.Frame1(2).Visible = True
        Me.Frame1(3).Visible = False
        Me.BtnBack.Enabled = False
        Me.BtnNext.Enabled = False
        Me.BtnExit.Enabled = False
        Me.PrgBar.Min = 0
        Me.PrgBar.Max = 100
        Me.PrgBar.Visible = True
        Me.CmdXK.Visible = False
    Case 3 '学科
        Me.Frame1(0).Visible = False
        Me.Frame1(1).Visible = False
        Me.Frame1(2).Visible = False
        Me.Frame1(3).Visible = True
        Me.BtnBack.Enabled = True
        Me.BtnNext.Enabled = False
        Me.BtnExit.Enabled = True
        Me.CmdXK.Visible = False
        Me.PrgBar.Min = 0
        Me.PrgBar.Max = 100
        Me.PrgBar.Visible = False
    Case Else
    End Select
End Sub

Private Sub BtnBack_Click()
    If CurrPage = 3 Then
        CurrPage = 0
    Else
        CurrPage = CurrPage - 1
    End If
    Page CurrPage
End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub BtnNext_Click()
Dim Sht As Excel.Worksheet
Dim Cls1 As ClsClass, Tchr1 As ClsTeacher, RowOffset As Integer, ColOffSet As Integer
Dim i As Integer, j As Integer
    
    If CurrPage >= 3 Then 'Xue Ke generate
        Me.BtnNext.Enabled = False
        Me.PrgBar.Min = 0
        Me.PrgBar.Max = Me.LstInEx.ListCount - 1
        Me.PrgBar.Visible = True
        For i = 0 To Me.LstInEx.ListCount - 1
            Set SchXls = GetObject(Me.LstInEx.List(i))
            Set InputSheet = Nothing
            For j = 1 To SchXls.Worksheets.Count
                If Left(Trim(SchXls.Worksheets(j).Range("A1").Text), 9) = "天津市第八十二中学" Then
                    Set InputSheet = SchXls.Worksheets(j)
                    Exit For
                End If
            Next j
            If Not InputSheet Is Nothing Then
                Set CInSht = New ClsInSht
                Set CInSht.SetSheet = InputSheet
                Set CInSht.StartCell = InputSheet.Range("A1")
                CInSht.DoCircleForXueKe Cxks
            End If
            Me.PrgBar.Value = i
            DoEvents
        Next i
        Me.PrgBar.Visible = False

        If Not FilSys.FileExists(App.Path & "\Template.xls") Then
            MsgBox "The Excel Template file ""Template.xls"" not found!"
            Exit Sub
        End If
        Set TmpBook = CreateObject("Excel.application")
'        Set TmpSheet = TmpBook.Worksheets("Sheet1")
        Set XKBook = TmpBook.Workbooks.Add(App.Path & "\Template.xls")
        
        Set Sht = XKBook.Worksheets(1)
        Sht.Name = Cxks.XKColl(1)(1).KerChen
        Me.PrgBar.Max = Cxks.Count
        Me.PrgBar.Min = 2
        Me.PrgBar.Visible = True
        Dim Buff As String, Buff1 As String, Buff2 As String, Pos As Integer
        For i = 2 To Cxks.Count
            Sht.Copy , XKBook.Worksheets(i - 1)
            XKBook.Worksheets(XKBook.Worksheets.Count).Name = Cxks.XKColl(i)(1).KerChen
            Buff = XKBook.Worksheets(XKBook.Worksheets.Count).Range("A1").Value
            Pos = InStr(1, Buff, "天津市第八十二中学")
            Buff1 = Left(Buff, Pos + Len("天津市第八十二中学") - 1)
            Buff2 = Mid(Buff, Pos + Len("天津市第八十二中学") + 2)
            XKBook.Worksheets(XKBook.Worksheets.Count).Range("A1").Value = Buff1 & Cxks.XKColl(i)(1).KerChen & Buff2
            DoEvents
            Me.PrgBar.Value = i
        Next i
        Me.PrgBar.Visible = False
        
        Dim CT As ClsTeacher
        Me.PrgBar.Min = 1
        Me.PrgBar.Max = Cxks.Count
        Me.PrgBar.Visible = True
        For i = 1 To Cxks.Count
            For j = 1 To Cxks.XKColl(i).Count
                Set CT = Cxks.XKColl(i)(j)
                CT.XKOutPut XKBook.Worksheets(CT.KerChen).Range("A5"), j - 1
                DoEvents
            Next j
            Me.PrgBar.Value = i
        Next i
        
        If FilSys.FileExists(Me.TxOtExFl.Text) Then
            FilSys.DeleteFile Me.TxOtExFl.Text, True
        End If
        
        XKBook.SaveAs Me.TxOtExFl.Text
        XKBook.Close
'        XKBook.Application.Quit
        Set TmpBook = Nothing
        Me.PrgBar.Visible = False
        Me.BtnNext.Enabled = True
        Exit Sub
    End If
    
    If CurrPage = 0 Then
        If Not FilSys.FileExists(Me.Text1.Text) Then
            MsgBox "The Excel file not found!", vbCritical
            Me.Text1.SelStart = 0: Me.Text1.SelLength = Len(Me.Text1.Text)
            Me.Text1.SetFocus
            Exit Sub
        End If
        Set SchXls = GetObject(Me.Text1.Text)
        Me.CmbIn.Clear
        Me.CmbClass.Clear
        Me.CmbTeacher.Clear
        For Each Sht In SchXls.Worksheets
            Me.CmbIn.AddItem Sht.Name
            Me.CmbClass.AddItem Sht.Name
            Me.CmbTeacher.AddItem Sht.Name
        Next Sht
    ElseIf CurrPage = 1 Then
        Set InputSheet = SchXls.Worksheets(Me.CmbIn.Text)
        Set CInSht = New ClsInSht
        If Me.OptKS(1).Value Then
            CInSht.CntOneDay = 8
        Else
            CInSht.CntOneDay = 7
        End If
        
        Set CInSht.SetSheet = InputSheet
        Set CInSht.StartCell = InputSheet.Range("A1")
        Set OutClassSheet = SchXls.Worksheets(Me.CmbClass.Text)
        Set OutTeacherSheet = SchXls.Worksheets(Me.CmbTeacher.Text)
        
        CurrPage = CurrPage + 1
        Me.Label10.Caption = "The Class sheet and Teacher sheet is generating now, Please wait..."
        Page CurrPage
        DoEvents
        CInSht.DoCircle ClsGp, ClsTc
        For Each Cls1 In ClsGp
            Set Cls1.SetSheet = OutClassSheet
            RowOffset = Int((Cls1.ClassNum - 1) / 2)
            ColOffSet = (Cls1.ClassNum - 1) - RowOffset * 2
            Set Cls1.StartCell = OutClassSheet.Range("A1").Offset(RowOffset * 10).Offset(, ColOffSet * 3)
            Cls1.OutPut
            If Me.PrgBar.Value >= 100 Then
                Me.PrgBar.Value = 0
            End If
            Me.PrgBar.Value = Me.PrgBar.Value + 1
        Next Cls1
        
        For Each Tchr1 In ClsTc
            Set Tchr1.SetSheet = OutTeacherSheet
            RowOffset = Int(i / 2)
            ColOffSet = i - RowOffset * 2
            Set Tchr1.StartCell = OutTeacherSheet.Range("A1").Offset(RowOffset * 10).Offset(, ColOffSet * 9)
            Tchr1.OutPut
            If Me.PrgBar.Value >= 100 Then
                Me.PrgBar.Value = 0
            End If
            Me.PrgBar.Value = Me.PrgBar.Value + 1
            i = i + 1
        Next Tchr1
        Me.Label10.Caption = "The process finished! Please open the Excel file and save the document."
        Me.PrgBar.Visible = False
        Me.BtnExit.Enabled = True
    End If
    CurrPage = CurrPage + 1
    Page CurrPage
End Sub

Private Sub BtnOpen_Click()
On Error GoTo Erl
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.Filter = "Excel File(*.xls)|*.xls|All Files(*.*)|*.*"
    Me.CommonDialog1.ShowOpen
    Me.Text1.Text = Me.CommonDialog1.FileName
Erl:
End Sub

Private Sub CmbClass_Click()
    Me.BtnNext.Enabled = Me.CmbIn.Text <> "" And Me.CmbClass.Text <> "" And Me.CmbTeacher.Text <> ""
End Sub

Private Sub CmbIn_Click()
    Me.BtnNext.Enabled = Me.CmbIn.Text <> "" And Me.CmbClass.Text <> "" And Me.CmbTeacher.Text <> ""
End Sub

Private Sub CmbTeacher_Click()
    Me.BtnNext.Enabled = Me.CmbIn.Text <> "" And Me.CmbClass.Text <> "" And Me.CmbTeacher.Text <> ""
End Sub

Private Sub CmdAdd_Click()
On Error GoTo Erl
Dim i As Integer, Fd As Boolean
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.Filter = "Excel File(*.xls)|*.xls|All Files(*.*)|*.*"
    Me.CommonDialog1.ShowOpen
    For i = 0 To Me.LstInEx.ListCount - 1
        If UCase(Me.LstInEx.List(i)) = UCase(Me.CommonDialog1.FileName) Then
            Fd = True
            Exit For
        End If
    Next i
    If Not Fd Then
        Me.LstInEx.AddItem Me.CommonDialog1.FileName
    End If
Erl:
    BtnEnable
End Sub

Private Sub CmdBrw_Click()
On Error GoTo Erl
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.Filter = "Excel File(*.xls)|*.xls|All Files(*.*)|*.*"
    Me.CommonDialog1.ShowOpen
    Me.TxOtExFl.Text = Me.CommonDialog1.FileName
Erl:
    BtnEnable
End Sub

Private Sub CmdDel_Click()
Dim i As Integer
    For i = 0 To Me.LstInEx.ListCount - 1
        If Me.LstInEx.Selected(i) Then
            Me.LstInEx.RemoveItem i
            Exit For
        End If
    Next i
    BtnEnable
End Sub

Private Sub CmdDown_Click()
Dim i As Integer, tmpstr As String
    With Me.LstInEx
    For i = 0 To .ListCount - 1
        If .Selected(i) Then
            tmpstr = .List(i)
            .List(i) = .List(i + 1)
            .List(i + 1) = tmpstr
            .Selected(i + 1) = True
            Exit For
        End If
    Next i
    End With
    BtnEnable
End Sub

Private Sub CmdUp_Click()
Dim i As Integer, tmpstr As String
    With Me.LstInEx
    For i = 0 To .ListCount - 1
        If .Selected(i) Then
            tmpstr = .List(i)
            .List(i) = .List(i - 1)
            .List(i - 1) = tmpstr
            .Selected(i - 1) = True
            Exit For
        End If
    Next i
    End With
    BtnEnable
End Sub

Private Sub CmdXK_Click()
    CurrPage = 3
    Page CurrPage
    BtnEnable
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Page CurrPage
    
    Me.OptKS(0).Value = True
            
    Show
    Me.Text1.Text = GetSetting("SGenRep", "Recent", "XleFile", "")
    BtnEnable
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "SGenRep", "Recent", "XleFile", Me.Text1.Text
End Sub

Private Sub BtnEnable()
Dim i As Integer
    If CurrPage = 3 Then
        Me.BtnNext.Enabled = Me.TxOtExFl.Text <> "" And Me.LstInEx.ListCount > 0
        For i = 0 To Me.LstInEx.ListCount - 1
            If Me.LstInEx.Selected(i) Then
                Me.CmdDel.Enabled = True
                Me.CmdUp.Enabled = i > 0
                Me.CmdDown.Enabled = i < Me.LstInEx.ListCount - 1
                Exit Sub
            End If
        Next i
        Me.CmdDel.Enabled = False
        Me.CmdUp.Enabled = False
        Me.CmdDown.Enabled = False
    End If
End Sub

Private Sub LstInEx_Click()
    BtnEnable
End Sub

Private Sub Text1_Change()
    Me.BtnNext.Enabled = Me.Text1.Text <> ""
End Sub
