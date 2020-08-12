VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama Program Penjualan"
   ClientHeight    =   4470
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5160
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "a"
            Object.ToolTipText     =   "Barang"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "b"
            Object.ToolTipText     =   "Kasir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "c"
            Object.ToolTipText     =   "Penjualan"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "d"
            Object.ToolTipText     =   "Laporan Data Barang"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "e"
            Object.ToolTipText     =   "Laporan Penjualan"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "f"
            Object.ToolTipText     =   "Rincian Penjualan"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "g"
            Object.ToolTipText     =   "Keluar"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CR 
      Left            =   1800
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar STBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4095
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1F503
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1F81D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1FB37
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":1FE51
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2016B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":20485
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":2079F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "&File"
      Begin VB.Menu mnbarang 
         Caption         =   "&Barang"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnkasir 
         Caption         =   "&Kasir"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "&Transaksi"
      Begin VB.Menu mn1 
         Caption         =   "&Penjualan"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnctkbarang 
         Caption         =   "Data Barang"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnctkpenjualan 
         Caption         =   "Data Penjualan"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnrincian 
         Caption         =   "Rincian Penjualan"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnutility 
      Caption         =   "Utility"
      Begin VB.Menu mnganpass 
         Caption         =   "Ganti Password User"
      End
      Begin VB.Menu mnbackup 
         Caption         =   "Backup Database"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub combo1_click()
    CR.SelectionFormula = "{detailjual.faktur}[1 to 10]='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\coba.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 0

End Sub

Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
End Sub



Private Sub mn1_Click()
Penjualan.Show
End Sub

Private Sub mn2_Click()
Penjualan2.Show
End Sub

Private Sub mnbackup_Click()
BackupDatabase.Show
End Sub

Private Sub mnbarang_Click()
Barang.Show
End Sub

Private Sub mnctkbarang_Click()
    CR.ReportFileName = App.Path & "\Lap Barang.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 0
End Sub

Private Sub mnctkpenjualan_Click()
Laporan.Show
End Sub

Private Sub mnganpass_Click()
GantiPass.Show
End Sub

Private Sub mnkasir_Click()
Kasir.Show
End Sub

Private Sub mnkeluar_Click()
End
End Sub

Private Sub mnretur_Click()
ReturJual.Show
End Sub

Private Sub mnretur1_Click()
Retur1.Show
End Sub

Private Sub mnrincian_Click()
Rincian.Show
End Sub

Private Sub mnsql_Click()
UjiSQL.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "a"
        Barang.Show
    Case "b"
        Kasir.Show
    Case "c"
        Penjualan.Show
    Case "d"
        CR.ReportFileName = App.Path & "\Lap Barang.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case "e"
        Laporan.Show
    Case "f"
       Rincian.Show
    Case "g"
        End
End Select
End Sub
