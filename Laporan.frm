VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Laporan 
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   1440
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3000
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Akhir"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1250
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Awal"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Harian"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3000
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   3000
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1250
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Laporan Transaksi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
'On Error Resume Next
Call BukaDB
RSPenjualan.Open "Select Distinct Tanggal From Penjualan order By 1", Conn
RSPenjualan.Requery
Do Until RSPenjualan.EOF
    Combo1.AddItem Format(RSPenjualan!Tanggal, "DD-MMM-YYYY")
    Combo2.AddItem Format(RSPenjualan!Tanggal, "YYYY ,MM, DD")
    Combo3.AddItem Format(RSPenjualan!Tanggal, "YYYY ,MM, DD")
    RSPenjualan.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct month(Tanggal) as Bulan from Penjualan", Conn
Do While Not RSTGL.EOF
    Combo4.AddItem RSTGL!Bulan & Space(5) & MonthName(RSTGL!Bulan)
    RSTGL.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(Tanggal)  as Tahun from Penjualan", Conn
Do While Not RSTHN.EOF
    Combo5.AddItem RSTHN!Tahun
    RSTHN.MoveNext
Loop
Conn.Close

End Sub

'Lap Harian
Private Sub combo1_click()
    CR.SelectionFormula = "Totext({Penjualan.Tanggal})='" & CDate(Combo1) & "'"
    CR.ReportFileName = App.Path & "\Lap Jual Harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'Lap Mingguan (Tgl Antara)
Private Sub Combo3_Click()
    If Combo2 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    Else
        If Combo3 < Combo2 Or Combo2 > Combo3 Then
            MsgBox "Tanggal terbalik"
            Combo3.SetFocus
            Exit Sub
        ElseIf Combo3 = Combo2 Then
            MsgBox "pilih tanggal yang berbeda"
            Combo3.SetFocus
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{Penjualan.Tanggal} in date (" & Combo2 & ") to date (" & Combo3 & ")"
    CR.ReportFileName = App.Path & "\Lap Jual Mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'Lap Bulanan
Private Sub Combo5_Click()
    Call BukaDB
    RSPenjualan.Open "select * from Penjualan where month(tanggal)='" & Val(Left(Combo4, 2)) & "' and year(tanggal)='" & (Combo5) & "'", Conn
    If RSPenjualan.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    CR.SelectionFormula = "Month({Penjualan.Tanggal})=" & Val(Left(Combo4, 2)) & " and Year({Penjualan.Tanggal})=" & Val(Combo5.Text)
    CR.ReportFileName = App.Path & "\Lap Jual Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub



