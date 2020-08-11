VERSION 5.00
Begin VB.Form Layar 
   BackColor       =   &H80000009&
   Caption         =   "Enter= Cetak **** ESC = Tutup"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
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
   ScaleHeight     =   6300
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Layar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Unload Me
ElseIf Keyascii = 13 Then
    Pesan = MsgBox("Printer sudah siap pakai..?", vbYesNo)
    If Pesan = vbYes Then
        Call CetakKePrinter
    Else
        Unload Me
    End If
End If
End Sub

Sub CetakKePrinter()
On Error GoTo salah
Call Koneksi
RSResep.Open "select * from Resep Where Nomorrsp In(Select Max(Nomorrsp)From Resep)Order By Nomorrsp Desc", Conn

Dim MGrs As String
Printer.Font = "Courier New"
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print
Printer.Print
RSPasien.Open "select * From pasien where KODEPSN= '" & RSResep!KodePsn & "'", Conn
RSDokter.Open "select * From Dokter where Kodedkt= '" & RSResep!Kodedkt & "'", Conn
RSPoli.Open "select * From poli where kodepl= '" & RSResep!Kodepl & "'", Conn

Printer.Print Tab(5); "Nomorrsp   :   "; RSResep!nomorrsp
Printer.Print Tab(5); "Tanggal    :   "; Format(RSResep!TanggalRsp, "DD-MMM-YY")
Printer.Print Tab(5); "Dokter     :   "; RSDokter!Namadkt
Printer.Print Tab(5); "Pasien     :   "; RSPasien!NamaPsn
Printer.Print Tab(5); "Poli       :   "; RSPoli!Namapl
MGrs = String$(33, "-")
Printer.Print Tab(5); MGrs
RSDetail.Open "select * from Detail Where Nomorrsp='" & RSResep!nomorrsp & "'", Conn
RSDetail.MoveFirst
No = 0
Do While Not RSDetail.EOF
    No = No + 1
    Set RSObat = New ADODB.Recordset
    RSObat.Open "select * From Obat where KodeObt= '" & RSDetail!KODEOBT & "'", Conn
    RSObat.Requery
    Printer.Print Tab(5); No; Space(2); RSObat!NamaOBT
    Printer.Print Tab(10); RKanan(RSDetail!dosis, "###"); Space(1); "X";
    Printer.Print Tab(15); Format(RSObat!hargaobt, "###,###,###");
    Printer.Print Tab(25); RKanan(RSDetail!dosis * RSObat!hargaobt, "###,###,###")
    RSDetail.MoveNext
Loop
'==========================
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Total      :";
Printer.Print Tab(25); RKanan(RSResep!TotalHRG, "###,###,###");
Printer.Print Tab(5); "Dibayar    :";
Printer.Print Tab(25); RKanan(RSResep!Dibayar, "###,###,###");
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Kembali    :";
If RSResep!Dibayar = RSResep!TotalHRG Then
    Printer.Print Tab(34); RSResep!Dibayar - RSResep!TotalHRG
Else
    Printer.Print Tab(25); RKanan(RSResep!Dibayar - RSResep!TotalHRG, "###,###,###");
End If
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "Semoga Lekas Sembuh"
Printer.Print
Printer.Print
Printer.Print
Conn.Close
Printer.EndDoc

On Error GoTo 0
Exit Sub
salah:
MsgBox "Cek apakah printer sudah ON...?"
End Sub

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

