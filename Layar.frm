VERSION 5.00
Begin VB.Form Layar 
   BackColor       =   &H80000009&
   Caption         =   "ESC=Tutup  ***  Enter=Cetak"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Layar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        Pesan = MsgBox("printer sudah siap", vbYesNo)
        If Pesan = vbYes Then
            Call Cetakprinter
        Else
            Unload Me
        End If
    ElseIf Keyascii = 27 Then
        Unload Me
    End If
End Sub

Function Cetakprinter()
Call BukaDB
RSPenjualan.Open "select * from penjualan Where Faktur In(Select Max(Faktur)From penjualan)Order By Faktur Desc", Conn
Dim Total, JmlJual, JmlHasil As Double
Dim MGrs As String
Printer.Font = "Courier New"
Printer.Print
Printer.Print
RSkasir.Open "select * From Kasir where KodeKsr= '" & RSPenjualan!KodeKsr & "'", Conn
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print Tab(5); "Faktur     :   "; RSPenjualan!Faktur
Printer.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!Tanggal, "DD-MMMM-YYYY")
Printer.Print Tab(5); "Jam        :   "; Format(RSPenjualan!Jam, "HH:MM:SS")
Printer.Print Tab(5); "Kasir      :   "; RSkasir!NamaKsr
MGrs = String$(33, "-")
Printer.Print Tab(5); MGrs
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
    Printer.Print Tab(5); No; Space(2); RSBarang!NamaBrg
    Printer.Print Tab(10); RKanan(Jumlah, "##"); Space(1); "X";
    Printer.Print Tab(15); Format(Harga, "###,###,###");
    Printer.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailJual.MoveNext
Loop
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Total      :";
Printer.Print Tab(25); RKanan(RSPenjualan!Total, "###,###,###");
Printer.Print Tab(5); "Dibayar    :";
Printer.Print Tab(25); RKanan(RSPenjualan!Dibayar, "###,###,###");
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Kembali    :";
If RSPenjualan!Dibayar = RSPenjualan!Total Then
    Printer.Print Tab(34); RSPenjualan!Dibayar - RSPenjualan!Total
Else
    Printer.Print Tab(25); RKanan(RSPenjualan!Dibayar - RSPenjualan!Total, "###,###,###");
End If
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Terima Kasih atas kunjungan Anda"
Printer.Print
Printer.Print
Printer.Print
Printer.EndDoc
Conn.Close
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function
