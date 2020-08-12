VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Barang 
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
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
   ScaleHeight     =   5070
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   5655
      Begin VB.CommandButton Cmdinput 
         Caption         =   "&Input"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton Cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton Cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   1800
         Width           =   1000
      End
      Begin VB.CommandButton Cmdtutup 
         Caption         =   "&Tutup"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1250
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   4260
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   1250
      End
      Begin VB.TextBox Text4 
         Height          =   350
         Left            =   4200
         TabIndex        =   7
         Top             =   960
         Width           =   1250
      End
      Begin VB.TextBox Text5 
         Height          =   350
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   1250
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Kode"
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Harga Beli"
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Harga Jual"
         Height          =   345
         Left            =   3120
         TabIndex        =   12
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Jumlah"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1005
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1845
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3254
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "KodeBrg"
         Caption         =   "Kode"
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
         DataField       =   "NamaBrg"
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
      BeginProperty Column02 
         DataField       =   "HargaBeli"
         Caption         =   "Harga Beli"
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
         DataField       =   "HargaJual"
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
      BeginProperty Column04 
         DataField       =   "JumlahBrg"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Barang"
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
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mvBookMark As Variant

Private Sub Form_Activate()
Call BukaDB
Conn.CursorLocation = adUseClient
RSBarang.Open "barang", Conn
With RSBarang
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
End With
Set DataGrid1.DataSource = RSBarang.DataSource
End Sub

Sub Form_Load()
Text1.MaxLength = 5
Text2.MaxLength = 30
Text3.MaxLength = 8
Text4.MaxLength = 8
Text5.MaxLength = 4
KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSBarang.Open "Select * From Barang where KodeBrg='" & Text1 & "'", Conn
End Function

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
End Sub

Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    CmdInput.Caption = "&Input"
    CmdEdit.Caption = "&Edit"
    CmdHapus.Caption = "&Hapus"
    CmdTutup.Caption = "&Tutup"
    CmdInput.Enabled = True
    CmdEdit.Enabled = True
    CmdHapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSBarang
        If Not RSBarang.EOF Then
            Text2 = RSBarang!NamaBrg
            Text3 = RSBarang!HargaBeli
            Text4 = RSBarang!HargaJual
            Text5 = RSBarang!jumlahbrg
        End If
    End With
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdInput.Caption = "&Simpan"
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Text1.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
            MsgBox "Data Belum Lengkap...!"
            If Text1 = "" Then
                Text1.SetFocus
            ElseIf Text2 = "" Then
                Text2.SetFocus
            ElseIf Text3 = "" Then
                Text3.SetFocus
            ElseIf Text3 = "" Then
                Text3.SetFocus
            ElseIf Text4 = "" Then
                Text4.SetFocus
            ElseIf Text5 = "" Then
                Text5.SetFocus
            End If
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Barang (KodeBrg,NamaBrg,HargaBeli,HargaJual,JumlahBrg) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "')"
            Conn.Execute SQLTambah
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    If CmdEdit.Caption = "&Edit" Then
        CmdInput.Enabled = False
        CmdEdit.Caption = "&Simpan"
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        Text1.SetFocus
    Else
        If Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Barang Set NamaBrg= '" & Text2 & "', HargaBeli='" & Text3 & "', HargaJual='" & Text4 & "',JumlahBrg='" & Text5 & "' where KodeBrg='" & Text1 & "'"
            Conn.Execute SQLEdit
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub CmdHapus_Click()
    If CmdHapus.Caption = "&Hapus" Then
        CmdInput.Enabled = False
        CmdEdit.Enabled = False
        CmdTutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        Text1.SetFocus
    End If
End Sub

Private Sub CmdTutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(Text1) < 5 Then
        MsgBox "Kode Harus 5 Digit"
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
            If Not RSBarang.EOF Then
                TampilkanData
                MsgBox "Kode Barang Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
            If Not RSBarang.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Barang Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSBarang.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Barang where kodebrg= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    KondisiAwal
                    Form_Activate
                Else
                    KondisiAwal
                    CmdHapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                Text1.SetFocus
            End If
    End If
End If
End Sub

Private Sub Text2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then Text4.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Text4_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If Val(Text4) <= Val(Text3) Then
            MsgBox "Harga jual jangan <= harga beli"
            Text4 = ""
            Text4.SetFocus
        Else
            Text5.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Text5_keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

