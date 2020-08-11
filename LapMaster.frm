VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LapMaster 
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
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
   ScaleHeight     =   2970
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   5160
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox List1 
      Height          =   1185
      Left            =   4440
      TabIndex        =   14
      Top             =   1440
      Width           =   2000
   End
   Begin VB.ComboBox Combo7 
      Height          =   345
      Left            =   4440
      TabIndex        =   12
      Top             =   960
      Width           =   2000
   End
   Begin VB.ComboBox Combo6 
      Height          =   345
      Left            =   2400
      TabIndex        =   11
      Top             =   2400
      Width           =   2000
   End
   Begin VB.ComboBox Combo5 
      Height          =   345
      Left            =   2400
      TabIndex        =   9
      Top             =   2040
      Width           =   2000
   End
   Begin VB.ComboBox Combo4 
      Height          =   345
      Left            =   2400
      TabIndex        =   8
      Top             =   1680
      Width           =   2000
   End
   Begin VB.ComboBox Combo3 
      Height          =   345
      Left            =   2400
      TabIndex        =   7
      Top             =   1320
      Width           =   2000
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   2400
      TabIndex        =   6
      Top             =   960
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   2000
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Laporan Data Master"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Laporan Resep Per :"
      Height          =   315
      Left            =   4440
      TabIndex        =   13
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Laporan Data Master"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   2250
   End
   Begin VB.Label Label6 
      Caption         =   "Laporan Pemakai Per Status"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2250
   End
   Begin VB.Label Label5 
      Caption         =   "Laporan Pasien Per Gender"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2250
   End
   Begin VB.Label Label4 
      Caption         =   "Laporan Obat Per Poli"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2250
   End
   Begin VB.Label Label3 
      Caption         =   "Laporan Obat Per Jenis"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2250
   End
   Begin VB.Label Label2 
      Caption         =   "Laporan Dokter (Spesialis)"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2250
   End
End
Attribute VB_Name = "LapMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Combo1.AddItem "Dokter"
Combo1.AddItem "Obat"
Combo1.AddItem "Pasien"
Combo1.AddItem "Poli"
Combo1.AddItem "Pemakai"
Combo1.AddItem "Pendaftaran"

Combo7.AddItem "Nomor"
Combo7.AddItem "Tanggal"
Combo7.AddItem "Dokter"
Combo7.AddItem "Pasien"
Combo7.AddItem "Poli"

Call Koneksi
RSDokter.Open "select distinct spesialis from dokter", Conn
Do While Not RSDokter.EOF
    Combo2.AddItem RSDokter!spesialis
    RSDokter.MoveNext
Loop

RSObat.Open "select distinct jenisobt from obat", Conn
Do While Not RSObat.EOF
    Combo3.AddItem RSObat!JenisObt
    RSObat.MoveNext
Loop
Conn.Close

Call Koneksi
RSObat.Open "select distinct katagori from obat", Conn
Do While Not RSObat.EOF
    Combo4.AddItem RSObat!katagori
    RSObat.MoveNext
Loop
Conn.Close

Call Koneksi
RSPasien.Open "select distinct genderpsn from pasien", Conn
Do While Not RSPasien.EOF
    Combo5.AddItem RSPasien!genderpsn
    RSPasien.MoveNext
Loop
Conn.Close

Call Koneksi
RSPemakai.Open "select distinct statuspmk from pemakai", Conn
Do While Not RSPemakai.EOF
    Combo6.AddItem RSPemakai!StatusPMK
    RSPemakai.MoveNext
Loop
Conn.Close

End Sub


Private Sub Combo1_Click()

If Combo1 = "Dokter" Then
    CR.ReportFileName = App.Path & "\dokter.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End If

If Combo1 = "Obat" Then
    CR.ReportFileName = App.Path & "\obat.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End If
    
If Combo1 = "Pasien" Then
     CR.ReportFileName = App.Path & "\pasien.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End If

If Combo1 = "Poli" Then
     CR.ReportFileName = App.Path & "\poli.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End If

If Combo1 = "Pemakai" Then
     CR.ReportFileName = App.Path & "\pemakai.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End If

If Combo1 = "Pendaftaran" Then
     CR.ReportFileName = App.Path & "\pendaftaran.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End If
End Sub

Private Sub Combo2_Click()
    CR.SelectionFormula = "({dokter.spesialis})='" & Combo2 & "'"
    CR.ReportFileName = App.Path & "\dokter.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Combo3_Click()
 CR.SelectionFormula = "({obat.jenisobt})='" & Combo3 & "'"
    CR.ReportFileName = App.Path & "\obat.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Combo4_Click()
 CR.SelectionFormula = "({obat.katagori})='" & Combo4 & "'"
    CR.ReportFileName = App.Path & "\obat.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Combo5_Click()
 CR.SelectionFormula = "({Pasien.genderpsn})='" & Combo5 & "'"
    CR.ReportFileName = App.Path & "\Pasien.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Combo6_Click()
 CR.SelectionFormula = "({Pemakai.statuspmk})='" & Combo6 & "'"
    CR.ReportFileName = App.Path & "\Pemakai.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Combo7_Click()
If Combo7 = "Nomor" Then
    Call Koneksi
    RSResep.Open "select distinct nomorrsp from resep", Conn
    List1.Clear
    Do While Not RSResep.EOF
        List1.AddItem RSResep!nomorrsp
        RSResep.MoveNext
    Loop
    Conn.Close
End If

If Combo7 = "Tanggal" Then
    Call Koneksi
    RSResep.Open "select distinct tanggalrsp from resep", Conn
    List1.Clear
    Do While Not RSResep.EOF
        List1.AddItem RSResep!TanggalRsp
        RSResep.MoveNext
    Loop
    Conn.Close
End If

If Combo7 = "Dokter" Then
    Call Koneksi
    RSDokter.Open "select distinct dokter.kodedkt,namadkt from Dokter,resep where dokter.kodedkt=resep.kodedkt", Conn
    List1.Clear
    Do While Not RSDokter.EOF
        List1.AddItem RSDokter!Kodedkt & Space(5) & RSDokter!Namadkt
        RSDokter.MoveNext
    Loop
    Conn.Close
End If

If Combo7 = "Pasien" Then
    Call Koneksi
    RSPasien.Open "select distinct pasien.kodepsn,namapsn from pasien,resep where pasien.kodepsn=resep.kodepsn ", Conn
    List1.Clear
    Do While Not RSPasien.EOF
        List1.AddItem RSPasien!KodePsn & Space(5) & RSPasien!NamaPsn
        RSPasien.MoveNext
    Loop
    Conn.Close
End If


If Combo7 = "Poli" Then
    Call Koneksi
    RSPoli.Open "select distinct poli.kodepl,namapl from poli,resep where poli.kodepl=resep.kodepl", Conn
    List1.Clear
    Do While Not RSPoli.EOF
        List1.AddItem RSPoli!Kodepl & Space(5) & RSPoli!Namapl
        RSPoli.MoveNext
    Loop
    Conn.Close
End If
End Sub

Private Sub List1_Click()
If Combo7 = "Nomor" Then
    CR.SelectionFormula = "({RESEP.nomorrsp})='" & List1 & "'"
    CR.ReportFileName = App.Path & "\reseppernomor.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End If

If Combo7 = "Tanggal" Then
    CR.SelectionFormula = "totext({RESEP.tanggalrsp})='" & CDate(List1) & "'"
    CR.ReportFileName = App.Path & "\reseppertanggal.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End If


If Combo7 = "Dokter" Then
    CR.SelectionFormula = "({resep.kodedkt})='" & Left(List1, 5) & "'"
    CR.ReportFileName = App.Path & "\resepperdokter.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End If

If Combo7 = "Pasien" Then
    CR.SelectionFormula = "({resep.kodepsn})='" & Left(List1, 8) & "'"
    CR.ReportFileName = App.Path & "\resepperpasien.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End If


If Combo7 = "Poli" Then
    CR.SelectionFormula = "({resep.kodepl})='" & Left(List1, 2) & "'"
    CR.ReportFileName = App.Path & "\resepperpoli.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End If


End Sub
