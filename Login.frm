VERSION 5.00
Begin VB.Form Login 
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3660
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
   ScaleHeight     =   1980
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   3345
      TabIndex        =   3
      Top             =   600
      Width           =   3400
      Begin VB.TextBox TxtNamaKsr 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1200
         TabIndex        =   0
         Top             =   120
         Width           =   2000
      End
      Begin VB.TextBox TxtPasswordKsr 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "X"
         TabIndex        =   1
         Top             =   600
         Width           =   2000
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama User"
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1005
      End
   End
   Begin VB.TextBox TxtKodeKsr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Width           =   2025
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
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
      TabIndex        =   7
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Kasir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1005
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A As Byte
Dim B As Byte

Private Sub Form_Load()
'batasi jumlah karakter
TxtNamaKsr.MaxLength = 30
TxtPasswordKsr.MaxLength = 10
'nama dan password diubah menjadi karakter X
'TxtNamaKsr.PasswordChar = "X"
TxtPasswordKsr.PasswordChar = "X"
TxtPasswordKsr.Enabled = False
TxtKodeKsr.Enabled = False
End Sub

Private Sub TxtNamaKsr_KeyPress(Keyascii As Integer)
'ubah karakter jadi besar semua
Keyascii = Asc(UCase(Chr(Keyascii)))
'jika menekan ESC form ditutup
If Keyascii = 27 Then Unload Me
'jika menekan enter setelah mengisi nama, maka..
If Keyascii = 13 Then
    'buka database
    Call BukaDB
    'cari nama kasir yang diketik
    RSkasir.Open "Select NamaKsr from Kasir where NamaKsr ='" & TxtNamaKsr & "'", Conn
    'jika tidak ditemukan, maka
    If RSkasir.EOF Then
        'batasi akses ke nama kasir 3 kali kesempatan
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            End
'            Unload Me
        End If
    Else
        'jika nama kasir benar, maka nama kasir menjadi false
        TxtNamaKsr.Enabled = False
        'password kasir menjadi true dan menjadi fokus kursor
        TxtPasswordKsr.Enabled = True
        TxtPasswordKsr.SetFocus
    End If
End If
End Sub

'coding ini sama dengan nama kasir
Private Sub txtpasswordksr_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
'Dim KodeKasir As String
'Dim NamaKasir As String
If Keyascii = 13 Then
    Call BukaDB
    RSkasir.Open "Select * from Kasir where NamaKsr ='" & TxtNamaKsr & "' and PasswordKsr='" & TxtPasswordKsr & "'", Conn
    If RSkasir.EOF Then
        B = B + 1
        If 1 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 2 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 3 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            End
'            Unload Me
        End If
    Else
        
        'jika nama dan password benar, maka...tutup form login
        'Unload Me
        TxtKodeKsr = RSkasir!KodeKsr
        Me.Visible = False
        'panggil menu utama
        Menu.Show
        'Menu.STBar.Panels(1).Text = RSkasir!kodeksr
        Menu.STBar.Panels(1).Text = Login.TxtKodeKsr
        Menu.STBar.Panels(2).Text = Login.TxtNamaKsr
        
    End If
End If
End Sub



'
'Dim A As Byte
'Dim B As Byte
'
'Private Sub Form_Load()
'    TxtNamaKsr.MaxLength = 35
'    TxtPasswordKsr.MaxLength = 15
'    TxtPasswordKsr.PasswordChar = "*"
'    TxtPasswordKsr.Enabled = False
'    TxtKodeKsr.Enabled = False
'End Sub
'
'Private Sub TxtNamaKsr_KeyPress(Keyascii As Integer)
'Keyascii = Asc(UCase(Chr(Keyascii)))
'If Keyascii = 27 Then Unload Me
'If Keyascii = 13 Then
'
'    Call BukaDB
'    RSkasir.Open "Select NamaKsr from Kasir where NamaKsr ='" & TxtNamaKsr & "'", Conn
'    If RSkasir.EOF Then
'        A = A + 1
'        If 1 - A = 0 Then
'            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
'                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
'            TxtNamaKsr = ""
'            TxtNamaKsr.SetFocus
'        ElseIf 2 - A = 0 Then
'            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
'                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
'            TxtNamaKsr = ""
'            TxtNamaKsr.SetFocus
'        ElseIf 3 - A = 0 Then
'            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
'                    "Nama '" & TxtNamaKsr & "' tidak dikenal" & Chr(13) & _
'                    "Kesempatan habis, Ulangi dari awal"
'            'End
'            Conn.Close
'            Unload Me
'        End If
'    Else
'        TxtNamaKsr.Enabled = False
'        TxtPasswordKsr.Enabled = True
'        TxtPasswordKsr.SetFocus
'        Conn.Close
'    End If
'End If
'End Sub
'
'Private Sub txtpasswordksr_KeyPress(Keyascii As Integer)
'Keyascii = Asc(UCase(Chr(Keyascii)))
'If Keyascii = 27 Then Unload Me
'If Keyascii = 13 Then
'
'    Call BukaDB
'    RSkasir.Open "Select * from Kasir where NamaKsr ='" & TxtNamaKsr & "' and PasswordKsr='" & TxtPasswordKsr & "'", Conn
'    If RSkasir.EOF Then
'        B = B + 1
'        If 1 - B = 0 Then
'            MsgBox "Kesempatan ke " & B & " Salah"
'            TxtPasswordKsr = ""
'            TxtPasswordKsr.SetFocus
'        ElseIf 2 - B = 0 Then
'            MsgBox "Kesempatan ke " & B & " Salah"
'            TxtPasswordKsr = ""
'            TxtPasswordKsr.SetFocus
'        ElseIf 3 - B = 0 Then
'            MsgBox "Kesempatan ke " & B & " Salah"
'            'End
'            Conn.Close
'            Unload Me
'        End If
'    Else
'        TxtKodeKsr = RSkasir!KodeKsr
'        Me.Visible = False
'        Menu.Show
'        Menu.STBar.Panels(1).Text = RSkasir!KodeKsr
'        Menu.STBar.Panels(2).Text = RSkasir!NamaKsr
'        Conn.Close
'    End If
'End If
'End Sub
'
