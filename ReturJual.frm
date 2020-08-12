VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ReturJual 
   Caption         =   "Retur Penjualan"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   120
      Width           =   1750
   End
   Begin VB.TextBox Dibayar 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   5760
      TabIndex        =   4
      Top             =   5040
      Width           =   1250
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   7200
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   800
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   960
      TabIndex        =   1
      Top             =   4680
      Width           =   800
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   1800
      TabIndex        =   0
      Top             =   4680
      Width           =   800
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3720
      Top             =   5040
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   1440
      Top             =   5160
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "Transaksi"
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
   Begin MSDataGridLib.DataGrid DTGrid 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   " Nomor"
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
         DataField       =   "Kode"
         Caption         =   "   Kode"
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
         DataField       =   "Nama"
         Caption         =   "Nama"
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
         DataField       =   "Harga"
         Caption         =   "       Harga"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "Total"
         Caption         =   "     Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label NoRetur 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1200
      TabIndex        =   22
      Top             =   480
      Width           =   1750
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Retur"
      Height          =   345
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label KodeKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2640
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Kembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5760
      TabIndex        =   19
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label Item 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4200
      TabIndex        =   18
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Jam 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5760
      TabIndex        =   17
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5760
      TabIndex        =   16
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label NamaKsr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5760
      TabIndex        =   15
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Tanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1200
      TabIndex        =   14
      Top             =   840
      Width           =   1750
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   4800
      TabIndex        =   13
      Top             =   5400
      Width           =   945
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Faktur"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kasir"
      Height          =   345
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   4800
      TabIndex        =   9
      Top             =   5040
      Width           =   945
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   345
      Left            =   4800
      TabIndex        =   8
      Top             =   4680
      Width           =   945
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jam"
      Height          =   345
      Left            =   4680
      TabIndex        =   7
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item"
      Height          =   345
      Left            =   3720
      TabIndex        =   6
      Top             =   4680
      Width           =   495
   End
End
Attribute VB_Name = "ReturJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
    Jam = Time$
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOJual.mdb"
Adodc1.RecordSource = "Transaksi"
Set DTGrid.DataSource = Adodc1
DTGrid.Refresh

'If KodeKsr = "" Then
'    MsgBox "Kasir tidak terdeteksi"
'    Login.Show
'    Exit Sub
'End If

Call BukaDB
RSBarang.Open "Barang", Conn
List1.Clear
Do Until RSBarang.EOF
    List1.AddItem RSBarang!NamaBrg & Space(50) & RSBarang!KodeBrg
    RSBarang.MoveNext
Loop

Call Auto
Call SiapTransaksi
Adodc1.Recordset.MoveFirst

Tanggal = Date
CmdSimpan.Enabled = False
End Sub

Private Sub Form_Load()
    Call BukaDB
    RSPenjualan.Open "select * from penjualan order by faktur", Conn
    Combo1.Clear
    Do While Not RSPenjualan.EOF
        Combo1.AddItem RSPenjualan!Faktur
        RSPenjualan.MoveNext
    Loop
    KodeKsr = Login.TxtKodeKsr
    NamaKsr = Login.TxtNamaKsr
    DTGrid.Col = 1
    CmdSimpan.Enabled = False
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

Function SiapTransaksi()
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Loop
    For i = 1 To 10
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!Nomor = i
        Adodc1.Recordset.Update
    Next i
    DTGrid.Col = 1
End Function

Private Sub DTGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        Adodc1.Recordset!Kode = Null
        Adodc1.Recordset!Nama = Null
        Adodc1.Recordset!Harga = Null
        Adodc1.Recordset!Jumlah = Null
        Adodc1.Recordset!Total = Null
        Adodc1.Recordset.Update
        Call TotalItem
        Call TotalHarga
        DTGrid.Refresh
End Select
End Sub

Private Sub DTGrid_AfterColEdit(ByVal ColIndex As Integer)
    If DTGrid.Col = 1 Then
        Call BukaDB
        RSBarang.Open "Select * from Barang where Kodebrg='" & Adodc1.Recordset!Kode & "'", Conn
        If RSBarang.EOF Then
            Pesan = MsgBox("Kode Barang Tidak Terdaftar")
            DTGrid.Col = 1
            Exit Sub
        End If
        Adodc1.Recordset!Kode = RSBarang!KodeBrg
        Adodc1.Recordset!Nama = RSBarang!NamaBrg
        Adodc1.Recordset!Harga = RSBarang!HargaJual
        DTGrid.Col = 4
        DTGrid.Refresh
        Exit Sub
    End If
    
    If DTGrid.Col = 4 Then
        Adodc1.Recordset!Jumlah = Adodc1.Recordset!Jumlah
        Adodc1.Recordset!Total = Adodc1.Recordset!Harga * Adodc1.Recordset!Jumlah
        Adodc1.Recordset.Update
        Adodc1.Recordset.MoveNext
        DTGrid.Col = 1
        Call TotalHarga
        Call TotalItem
    End If
End Sub

Function TotalItem()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Item = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Jumlah <> 0
    Item = Item + Adodc1.Recordset!Jumlah
    Adodc1.Recordset.MoveNext
    Item = Item
Loop
End Function

Function TotalHarga()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Total = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Total <> 0
    Total = Total + Adodc1.Recordset!Total
    Adodc1.Recordset.MoveNext
    Total = Format(Total, "#,###,###")
Loop
End Function

Private Sub Bersihkan()
    Item = ""
    Total = ""
    Dibayar = ""
    Kembali = ""
End Sub

Private Sub Dibayar_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If Dibayar = "" Or Val(Dibayar) < (Total) Then
            MsgBox "Jumlah Pembayaran Kurang"
            Dibayar.SetFocus
        Else
            Dibayar = Format(Dibayar, "###,###,###")
            If Dibayar = Total Then
                Kembali = Dibayar - Total
            Else
                Kembali = Format(Dibayar - Total, "###,###,###")
            End If
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub CmdSimpan_Keypress(Keyascii As Integer)
    If Keyascii = 27 Then
        CmdSimpan.Enabled = False
        Dibayar = ""
        Dibayar.SetFocus
    End If
End Sub

Private Sub CmdSimpan_Click()
   
    Dim SQLTambahJual As String
    SQLTambahJual = "Insert Into Penjualan(Faktur,Tanggal,Jam,Total,Item,Dibayar,Kembali,KodeKsr)" & _
    "values('" & Faktur & "','" & Tanggal & "','" & Jam & "','" & Total & "','" & Item & "','" & Dibayar & "','" & Kembali & "','" & KodeKsr & "')"
    Conn.Execute (SQLTambahJual)
         
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Kode <> vbNullString Then
            Dim SQLTambahDetail As String
            SQLTambahDetail = "Insert Into Detailjual(Faktur,Kodebrg,JmlJual) " & _
            "values ('" & Faktur + Adodc1.Recordset!Nomor & "','" & Adodc1.Recordset!Kode & "','" & Adodc1.Recordset!Jumlah & "')"
            Conn.Execute (SQLTambahDetail)
        End If
    Adodc1.Recordset.MoveNext
    Loop
        
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Kode <> vbNullString Then
            Call BukaDB
            RSBarang.Open "Select * from Barang where Kodebrg='" & Adodc1.Recordset!Kode & "'", Conn
            If Not RSBarang.EOF Then
                Dim Kurangi As String
                Kurangi = "update barang set jumlahbrg='" & RSBarang!JumlahBrg - Adodc1.Recordset!Jumlah & "' where kodebrg='" & Adodc1.Recordset!Kode & "'"
                Conn.Execute (Kurangi)
            End If
        End If
    Adodc1.Recordset.MoveNext
    Loop
    Bersihkan
    Form_Activate
    Call Cetak
End Sub

Private Sub CmdBatal_Click()
    Dibayar = ""
    Total = ""
    Item = ""
    Form_Activate
End Sub

Private Sub Cmadodc1utup_Click()
    Unload Me
End Sub

Function Cetak()
Call BukaDB
RSPenjualan.Open "select * from penjualan Where Faktur In(Select Max(Faktur)From penjualan)Order By Faktur Desc", Conn
Layar.Show
Dim Total, JmlJual, JmlHasil As Double
Dim MGrs As String
Layar.Font = "Courier New"
Layar.Print
Layar.Print
RSkasir.Open "select * From Kasir where KodeKsr= '" & RSPenjualan!KodeKsr & "'", Conn
Layar.Print Tab(5); "Faktur     :   "; RSPenjualan!Faktur
Layar.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!Tanggal, "DD-MMMM-YYYY")
Layar.Print Tab(5); "Jam        :   "; Format(RSPenjualan!Jam, "HH:MM:SS")
Layar.Print Tab(5); "Kasir      :   "; RSkasir!NamaKsr
MGrs = String$(33, "-")
Layar.Print Tab(5); MGrs
RSDetailJual.Open "select * from detailjual Where left(Faktur,10)='" & RSPenjualan!Faktur & "'", Conn
RSDetailJual.MoveFirst
No = 0
Do While Not RSDetailJual.EOF
    No = No + 1
    Set RSBarang = New ADODB.Recordset
    RSBarang.Open "select * From Barang where Kodebrg= '" & RSDetailJual!KodeBrg & "'", Conn
    RSBarang.Requery
    Harga = RSBarang!HargaJual
    Jumlah = RSDetailJual!JmlJual
    Hasil = Harga * Jumlah
    Layar.Print Tab(5); No; Space(2); RSBarang!NamaBrg
    Layar.Print Tab(10); RKanan(Jumlah, "##"); Space(1); "X";
    Layar.Print Tab(15); Format(Harga, "###,###,###");
    Layar.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailJual.MoveNext
Loop
Layar.Print Tab(5); MGrs
Layar.Print Tab(5); "Total      :";
Layar.Print Tab(25); RKanan(RSPenjualan!Total, "###,###,###");
Layar.Print Tab(5); "Dibayar    :";
Layar.Print Tab(25); RKanan(RSPenjualan!Dibayar, "###,###,###");
Layar.Print Tab(5); MGrs
Layar.Print Tab(5); "Kembali    :";
If RSPenjualan!Dibayar = RSPenjualan!Total Then
    Layar.Print Tab(34); RSPenjualan!Dibayar - RSPenjualan!Total
Else
    Layar.Print Tab(25); RKanan(RSPenjualan!Dibayar - RSPenjualan!Total, "###,###,###");
End If
Layar.Print Tab(5); MGrs
Layar.Print Tab(5); "Terima Kasih atas kunjungan Anda"
Layar.Print
Layar.Print
Layar.Print
Conn.Close
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Private Sub List1_keyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If DTGrid.SelText <> Right(List1, 5) Then
            DTGrid.SelText = Right(List1, 5)
            Adodc1.Recordset.Update
            Call BukaDB
            RSBarang.Open "Select * from Barang where KodeBrg='" & Right(List1, 5) & "'", Conn, adOpenDynamic, adLockOptimistic
            RSBarang.Requery
            If Not RSBarang.EOF Then
                Adodc1.Recordset!Kode = RSBarang!KodeBrg
                Adodc1.Recordset!Nama = RSBarang!NamaBrg
                Adodc1.Recordset!Harga = RSBarang!HargaJual
                Adodc1.Recordset.Update
                DTGrid.SetFocus
                DTGrid.Col = 4
            End If
        End If
    End If
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub




