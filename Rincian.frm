VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Rincian 
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   4260
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   6720
      TabIndex        =   7
      Top             =   3360
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   350
      Left            =   6240
      TabIndex        =   6
      Top             =   3360
      Width           =   500
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   960
      TabIndex        =   4
      Top             =   3360
      Width           =   1000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   2880
      TabIndex        =   2
      Top             =   3360
      Width           =   2800
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   3840
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
         Name            =   "Century"
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
      Bindings        =   "Rincian.frx":0000
      Height          =   2655
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
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
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rincian Penjualan Barang"
      BeginProperty Font 
         Name            =   "Century"
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
      TabIndex        =   8
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kasir"
      Height          =   345
      Left            =   2040
      TabIndex        =   3
      Top             =   3360
      Width           =   855
   End
End
Attribute VB_Name = "Rincian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
Call BukaDB
List1.Clear
RSPenjualan.Open "Select Distinct Faktur from Penjualan ", Conn
Do Until RSPenjualan.EOF
    List1.AddItem RSPenjualan!Faktur
    RSPenjualan.MoveNext
Loop
Conn.Close
Call Gelap
End Sub

Private Sub list1_click()
Call BukaDB
Conn.CursorLocation = adUseClient
RSPenjualan.Open "select * from Penjualan where Faktur='" & List1.Text & "'", Conn
RSPenjualan.Requery

If Not RSPenjualan.EOF Then Text1 = RSPenjualan!Tanggal

RSkasir.Open "select * from Kasir where KodeKsr='" & RSPenjualan!KodeKsr & "'", Conn
If Not RSkasir.EOF Then Text2 = RSkasir!NamaKsr
Conn.Close

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOJual.mdb"
Adodc1.RecordSource = "select NamaBrg as [Nama Barang], HargaJual as [Harga Jual],JmlJual as Jumlah, HargaJual*JmlJual as Total from Barang,detailJual,penjualan where DetailJual.kodeBrg=Barang.kodeBrg and left(detailjual.faktur,10)=penjualan.faktur and penjualan.faktur='" & List1 & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Call Total
Call Item
End Sub

Private Sub List1_keyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Function Item()
Adodc1.Recordset.MoveFirst
Jumlah = 0
Do While Not Adodc1.Recordset.EOF
    Jumlah = Jumlah + Adodc1.Recordset!Jumlah
    Adodc1.Recordset.MoveNext
Loop
Text3 = Jumlah
End Function

Function Total()
Adodc1.Recordset.MoveFirst
Jumlah = 0
Do While Not Adodc1.Recordset.EOF
    Jumlah = Jumlah + Adodc1.Recordset!Total
    Adodc1.Recordset.MoveNext
Loop
Text4 = Jumlah
End Function

Sub Gelap()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub
