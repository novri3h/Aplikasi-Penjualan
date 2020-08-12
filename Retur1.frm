VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Retur1 
   Caption         =   "Retur Penjualan"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   1800
      TabIndex        =   22
      Top             =   4080
      Width           =   800
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   960
      TabIndex        =   21
      Top             =   4080
      Width           =   800
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   800
   End
   Begin VB.TextBox Dibayar 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   4920
      TabIndex        =   11
      Top             =   4440
      Width           =   1250
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1750
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3840
      Top             =   840
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Retur1.frx":0000
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Nama Barang"
         Caption         =   "Nama Barang"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Harga Jual"
         Caption         =   "Harga Jual"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Jumlah"
         Caption         =   "Jumlah"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   345
      Left            =   2880
      TabIndex        =   19
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   3960
      TabIndex        =   18
      Top             =   4080
      Width           =   945
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   3960
      TabIndex        =   17
      Top             =   4440
      Width           =   945
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   3960
      TabIndex        =   16
      Top             =   4800
      Width           =   945
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4920
      TabIndex        =   15
      Top             =   4080
      Width           =   1245
   End
   Begin VB.Label Item 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3360
      TabIndex        =   14
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Kembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4920
      TabIndex        =   13
      Top             =   4800
      Width           =   1245
   End
   Begin VB.Label KodeKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3120
      TabIndex        =   12
      Top             =   4440
      Width           =   750
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jam"
      Height          =   345
      Left            =   3840
      TabIndex        =   10
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kasir"
      Height          =   345
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Faktur"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Tanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label NamaKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Jam 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Retur"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label NoRetur 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   1755
   End
End
Attribute VB_Name = "Retur1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call Auto
End Sub

Private Sub Form_Load()
On Error Resume Next
Call BukaDB
Combo1.Clear
RSPenjualan.Open "Select Distinct Faktur from Penjualan ", Conn
Do Until RSPenjualan.EOF
    Combo1.AddItem RSPenjualan!Faktur
    RSPenjualan.MoveNext
Loop
Conn.Close
Call Gelap
End Sub

Private Sub Auto()
Call BukaDB
RSRTPenjualan.Open "select * from rtrPenjualan Where Noretur In(Select Max(noretur)From RTRPenjualan)Order By noretur Desc", Conn
RSRTPenjualan.Requery
    Dim Urutan As String * 12
    Dim Hitung As Long
    With RSRTPenjualan
        If .EOF Then
            Urutan = "R" + Format(Date, "yymmdd") + "0001"
            NoRetur = Urutan
        Else
            If Mid(!NoRetur, 2, 6) <> Format(Date, "yymmdd") Then
                Urutan = "R" + Format(Date, "yymmdd") + "0001"
            Else
                Hitung = Right(!NoRetur, 11) + 1
                Urutan = "R" + Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
            End If
        End If
        NoRetur = Urutan
    End With
End Sub

Private Sub combo1_click()
Call BukaDB
Conn.CursorLocation = adUseClient
RSPenjualan.Open "select * from Penjualan where Faktur='" & Combo1.Text & "'", Conn
RSPenjualan.Requery

If Not RSPenjualan.EOF Then
    Tanggal = RSPenjualan!Tanggal
    Jam = RSPenjualan!Jam
    KodeKsr = RSPenjualan!KodeKsr
    Item = RSPenjualan!Item
    Total = RSPenjualan!Total
    Dibayar = RSPenjualan!Dibayar
    Kembali = RSPenjualan!Kembali
End If
    
If KodeKsr = "" Then
    NamaKsr = ""
Else
    RSkasir.Open "select * from Kasir where KodeKsr='" & KodeKsr & "'", Conn
    If Not RSkasir.EOF Then NamaKsr = RSkasir!NamaKsr
    Conn.Close
End If
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOJual.mdb"
Adodc1.RecordSource = "select NamaBrg as [Nama Barang], HargaJual as [Harga Jual],JmlJual as Jumlah, HargaJual*JmlJual as Total from Barang,detailJual,penjualan where DetailJual.kodeBrg=Barang.kodeBrg and left(detailjual.faktur,10)=penjualan.faktur and penjualan.faktur='" & Combo1 & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
'Call Total
'Call Item
End Sub

Private Sub Datagrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
    Case vbKeyEscape
'        Adodc1.Recordset!Kode = Null
        Adodc1.Recordset![Nama Barang] = Null
        Adodc1.Recordset![Harga Jual] = Null
        Adodc1.Recordset!Jumlah = Null
        Adodc1.Recordset!Total = Null
        Adodc1.Recordset.Update
'        Call TotalItem
'        Call TotalHarga
        DataGrid1.Refresh
End Select
End Sub

Private Sub combo1_keyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Function Item()
'Adodc1.Recordset.MoveFirst
'Jumlah = 0
'Do While Not Adodc1.Recordset.EOF
'    Jumlah = Jumlah + Adodc1.Recordset!Jumlah
'    Adodc1.Recordset.MoveNext
'Loop
'Text3 = Jumlah
'End Function

'Function Total()
'Adodc1.Recordset.MoveFirst
'Jumlah = 0
'Do While Not Adodc1.Recordset.EOF
'    Jumlah = Jumlah + Adodc1.Recordset!Total
'    Adodc1.Recordset.MoveNext
'Loop
'Text4 = Jumlah
'End Function

Sub Gelap()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub

