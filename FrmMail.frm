VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Generator"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6540
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
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
         Caption         =   "Select excel file full path. and click Next button."
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
      Height          =   2055
      Index           =   2
      Left            =   660
      TabIndex        =   17
      Top             =   765
      Width           =   6255
      Begin VB.Label Label10 
         Caption         =   "The Class sheet and Teacher sheet is generated, Please check."
         Height          =   1590
         Left            =   225
         TabIndex        =   18
         Top             =   240
         Width           =   5790
      End
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   150
      Left            =   180
      TabIndex        =   16
      Top             =   2265
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Index           =   1
      Left            =   330
      TabIndex        =   8
      Top             =   675
      Width           =   6255
      Begin VB.ComboBox CmbTeacher 
         Height          =   300
         Left            =   3105
         TabIndex        =   15
         Top             =   1485
         Width           =   2445
      End
      Begin VB.ComboBox CmbClass 
         Height          =   300
         Left            =   3105
         TabIndex        =   13
         Top             =   1050
         Width           =   2445
      End
      Begin VB.ComboBox CmbIn 
         Height          =   300
         Left            =   3105
         TabIndex        =   11
         Top             =   660
         Width           =   2445
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   390
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton BtnBack 
      Caption         =   "&Back"
      Height          =   405
      Left            =   2220
      TabIndex        =   7
      Top             =   2610
      Width           =   1185
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   4860
      TabIndex        =   5
      Top             =   2610
      Width           =   1185
   End
   Begin VB.CommandButton BtnNext 
      Caption         =   "&Next"
      Height          =   405
      Left            =   3555
      TabIndex        =   4
      Top             =   2610
      Width           =   1185
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
        Me.BtnBack.Enabled = False
        Me.BtnNext.Enabled = True
        Me.PrgBar.Visible = False
    Case 1
        Me.Frame1(0).Visible = False
        Me.Frame1(1).Visible = True
        Me.Frame1(2).Visible = False
        Me.BtnBack.Enabled = True
        Me.BtnNext.Enabled = True
        Me.PrgBar.Visible = False
    Case 2
        Me.Frame1(0).Visible = False
        Me.Frame1(1).Visible = False
        Me.Frame1(2).Visible = True
        Me.BtnBack.Enabled = True
        Me.BtnNext.Enabled = False
        Me.PrgBar.Visible = True
    Case Else
    End Select
End Sub

Private Sub BtnBack_Click()
    CurrPage = CurrPage - 1
    Page CurrPage
End Sub

Private Sub BtnExit_Click()
    Unload Me
End Sub

Private Sub BtnNext_Click()
Dim Sht As Excel.Worksheet
    If CurrPage = 0 Then
        If Not FilSys.FileExists(Me.Text1.Text) Then
            MsgBox "The Excel file not found!", vbCritical
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
    
    End If
    CurrPage = CurrPage + 1
    Page CurrPage
End Sub

Private Sub BtnOpen_Click()
On Error Resume Next
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.Filter = "Excel File(*.xls)|*.xls|All Files(*.*)|*.*"
    Me.CommonDialog1.ShowOpen
    Me.Text1.Text = Me.CommonDialog1.FileName
End Sub

Private Sub Form_Load()
    Page CurrPage
End Sub
