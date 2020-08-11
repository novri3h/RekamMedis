Attribute VB_Name = "Module1"

Public Conn As New ADODB.Connection
Public RSObat As ADODB.Recordset
Public RSADM As ADODB.Recordset
Public RSApoteker As ADODB.Recordset
Public RSPendaftaran As ADODB.Recordset
Public RSPembayaran As ADODB.Recordset
Public RSPemakai As ADODB.Recordset
Public RSPoli As ADODB.Recordset
Public RSDokter As ADODB.Recordset
Public RSResep As ADODB.Recordset
Public RSPasien As ADODB.Recordset
Public RSDetail As ADODB.Recordset


Public Sub Koneksi()
Set Conn = New ADODB.Connection
Set RSObat = New ADODB.Recordset
Set RSADM = New ADODB.Recordset
Set RSApoteker = New ADODB.Recordset
Set RSPendaftaran = New ADODB.Recordset
Set RSPembayaran = New ADODB.Recordset
Set RSPemakai = New ADODB.Recordset
Set RSPoli = New ADODB.Recordset
Set RSDokter = New ADODB.Recordset
Set RSResep = New ADODB.Recordset
Set RSPasien = New ADODB.Recordset
Set RSDetail = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBrawatjalan.mdb"
End Sub




