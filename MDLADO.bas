Attribute VB_Name = "MdlADO"

Public Conn As New ADODB.Connection
Public RSBarang As ADODB.Recordset
Public RSkasir As ADODB.Recordset
Public RSPenjualan As ADODB.Recordset
Public RSDetailJual As ADODB.Recordset
Public RSTR1 As ADODB.Recordset
Public RSRTPenjualan As ADODB.Recordset
Public RSRTDetailJual As ADODB.Recordset
Public RSTR2 As ADODB.Recordset
Public STr As String


Public Sub BukaDB()
Set Conn = New ADODB.Connection
Set RSBarang = New ADODB.Recordset
Set RSkasir = New ADODB.Recordset
Set RSPenjualan = New ADODB.Recordset
Set RSDetailJual = New ADODB.Recordset
Set RSTR1 = New ADODB.Recordset
Set RSRTPenjualan = New ADODB.Recordset
Set RSRTDetailJual = New ADODB.Recordset

Set RSTR1 = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ADOJual.mdb"
End Sub

