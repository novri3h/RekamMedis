VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Resep 
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
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
   ScaleHeight     =   5880
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   3480
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Dibayar 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   6480
      TabIndex        =   3
      Top             =   5040
      Width           =   1250
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   2280
      TabIndex        =   6
      Top             =   4680
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   1200
      TabIndex        =   5
      Top             =   4680
      Width           =   1000
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   1000
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "Resep.frx":0000
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   "Nomor"
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
         Caption         =   "Harga"
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
         DataField       =   "Dosis"
         Caption         =   "Dosis"
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
         DataField       =   "SubTotal"
         Caption         =   "SubTotal"
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
            Locked          =   -1  'True
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1244,976
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   375
      Left            =   120
      Top             =   5160
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Resep"
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
      TabIndex        =   25
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label Kembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6480
      TabIndex        =   24
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6480
      TabIndex        =   23
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kembali"
      Height          =   345
      Left            =   5160
      TabIndex        =   22
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dibayar"
      Height          =   345
      Left            =   5160
      TabIndex        =   21
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   345
      Left            =   5160
      TabIndex        =   20
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label Item 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3480
      TabIndex        =   19
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Namapl 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5520
      TabIndex        =   18
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label NamaPsn 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5520
      TabIndex        =   17
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Namadkt 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5520
      TabIndex        =   16
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Kodepl 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4440
      TabIndex        =   15
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label KodePsn 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4440
      TabIndex        =   14
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Kodedkt 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4440
      TabIndex        =   13
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Poli"
      Height          =   345
      Left            =   3120
      TabIndex        =   12
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Pasien"
      Height          =   345
      Left            =   3120
      TabIndex        =   11
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Dokter"
      Height          =   345
      Left            =   3120
      TabIndex        =   10
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label Tanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   9
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Resep"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1245
   End
End
Attribute VB_Name = "Resep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call Koneksi
ado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBRAWATJALAN.mdb"
ado.RecordSource = "Temporer"
Set dg.DataSource = ado
dg.Refresh
RSPendaftaran.Open "SELECT * FROM PENDAFTARAN where ket='0'", Conn
Combo1.Clear
Do Until RSPendaftaran.EOF
    Combo1.AddItem RSPendaftaran!NomorDft
    RSPendaftaran.MoveNext
Loop

Call Tabel_Kosong
ado.Recordset.MoveFirst
Tanggal = Format(Date, "DD-MM-YYYY")
End Sub


Function Tabel_Kosong()
    ado.Recordset.MoveFirst
    Do While Not ado.Recordset.EOF
        ado.Recordset.Delete
        ado.Recordset.MoveNext
    Loop
    For I = 1 To 1
        ado.Recordset.AddNew
        ado.Recordset!Nomor = I
        ado.Recordset.Update
    Next I
    dg.Col = 1
End Function

Private Sub Combo1_Click()
Call Koneksi
RSPendaftaran.Open "Select * from Pendaftaran where nomordft='" & Combo1 & "'", Conn
RSPendaftaran.Requery
If Not RSPendaftaran.EOF Then
    RSDokter.Open "select * from dokter where kodedkt='" & RSPendaftaran!Kodedkt & "'", Conn
    If Not RSDokter.EOF Then
        Kodedkt = RSDokter!Kodedkt
        Namadkt = RSDokter!Namadkt
    End If
    
    RSPasien.Open "select * from pasien where kodepsn='" & RSPendaftaran!KodePsn & "'", Conn
    If Not RSPasien.EOF Then
        KodePsn = RSPasien!KodePsn
        NamaPsn = RSPasien!NamaPsn
    End If
    
    RSPoli.Open "select * from poli where kodepl='" & RSPendaftaran!Kodepl & "'", Conn
    If Not RSPoli.EOF Then
        Kodepl = RSPoli!Kodepl
        Namapl = RSPoli!Namapl
    End If
    
    RSObat.Open "SELECT * FROM OBAT WHERE KATAGORI= '" & Namapl & "'", Conn
    List1.Clear
    Do While Not RSObat.EOF
        List1.AddItem RSObat!NamaOBT & Space(5) & RSObat!JUMLAHOBT & Space(50) & RSObat!KODEOBT
        RSObat.MoveNext
    Loop
    
Else
    MsgBox "nomor tidak terdaftar"
    Combo1.SetFocus
End If

End Sub

Private Sub combo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Combo1 = "" Then
        MsgBox "nomor resep harus diisi"
        Combo1.SetFocus
        Exit Sub
    Else
        Combo1_Click
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub DG_AfterColEdit(ByVal ColIndex As Integer)
    If dg.Col = 1 Then
        If Len(ado.Recordset!Kode) < 5 Then
            MsgBox "Kode Harus 5 digit"
            dg.Col = 1
            Exit Sub
        End If
    
        Call Koneksi
        RSObat.Open "Select * from Obat where KodeObt='" & ado.Recordset!Kode & "'", Conn
        If Not RSObat.EOF Then
            ado.Recordset!Kode = RSObat!KODEOBT
            ado.Recordset!Nama = RSObat!NamaOBT
            ado.Recordset!Harga = RSObat!hargaobt
            dg.Col = 4
            dg.Refresh
            Exit Sub
        End If
    End If
    
    If dg.Col = 4 Then
        If ado.Recordset!dosis > RSObat!JUMLAHOBT Then
            MsgBox "STOK OBAT KURANG"
            Exit Sub
        Else
            ado.Recordset!dosis = ado.Recordset!dosis
            ado.Recordset!subtotal = ado.Recordset!Harga * ado.Recordset!dosis
            ado.Recordset.Update
            Call Tambah_Baris
            ado.Recordset.MoveNext
            dg.Col = 1
            ado.Recordset.MoveLast
            Item = Format(Jumlah, "#,###,###")
            Total = Format(Jumlah2, "#,###,###")
        End If
    End If
End Sub

Private Sub List1_keyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If dg.SelText <> Right(List1, 5) Then
            dg.SelText = Right(List1, 5)
            ado.Recordset.Update
            Call Koneksi
            RSObat.Open "Select * from Obat where KodeObt='" & Right(List1, 5) & "'", Conn
            RSObat.Requery
            If Not RSObat.EOF Then
                ado.Recordset!Kode = RSObat!KODEOBT
                ado.Recordset!Nama = RSObat!NamaOBT
                ado.Recordset!Harga = RSObat!hargaobt
                ado.Recordset.Update
                dg.SetFocus
                dg.Col = 4
            End If
        End If
    End If
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


Function Tambah_Baris()
    For I = ado.Recordset.RecordCount To ado.Recordset.RecordCount
        ado.Recordset.AddNew
        ado.Recordset!Nomor = I + 1
        ado.Recordset.Update
    Next I
End Function

Private Sub DG_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If dg.Col = 4 Then
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack Or Keyascii = vbKeyReturn) Then Keyascii = 0
End If
End Sub


Private Sub Bersihkan()
    Combo1 = ""
    Kodedkt = ""
    Namadkt = ""
    KodePsn = ""
    NamaPsn = ""
    Kodepl = ""
    Namapl = ""
    Total = ""
    Dibayar = ""
    Kembali = ""
    Combo1 = ""
    Item = ""
    List1.Clear
End Sub

Private Sub CmdSimpan_Click()
If Combo1 = "" Or Item = "" Then
    MsgBox "Data belum lengkap"
    Exit Sub
End If

    Call Koneksi
    Dim InputResep As String
    'simpan ke tabel resep
    InputResep = "Insert Into Resep(Nomorrsp,Tanggalrsp,kodedkt,kodepsn,kodepl,kodepmk,TotalHrg,Dibayar,Kembali)" & _
    "values('" & Combo1 & "','" & Tanggal & "','" & Kodedkt & "','" & KodePsn & "','" & Kodepl & "','" & Menu.STBar.Panels(3).Text & "','" & Total & "','" & Dibayar & "','" & Kembali & "')"
    Conn.Execute (InputResep)
    
    aaa = "update pendaftaran set ket='1' where nomordft='" & Combo1 & "'"
    Conn.Execute aaa
    
    'simpan ke tabel detailresep
    ado.Recordset.MoveFirst
    Do While Not ado.Recordset.EOF
        If ado.Recordset!Kode <> vbNullString Then
            Dim InputDetail As String
            InputDetail = "Insert Into Detail(Nomorrsp,KodeObt,harga,dosis,subtotal) " & _
            "values ('" & Combo1 & "','" & ado.Recordset!Kode & "','" & ado.Recordset!Harga & "','" & ado.Recordset!dosis & "','" & ado.Recordset!subtotal & "')"
            Conn.Execute (InputDetail)
        End If
    ado.Recordset.MoveNext
    Loop
        
    'kurangi jumlah obat
    ado.Recordset.MoveFirst
    Do While Not ado.Recordset.EOF
        If ado.Recordset!Kode <> vbNullString Then
            Call Koneksi
            RSObat.Open "Select * from Obat where KodeObt='" & ado.Recordset!Kode & "'", Conn
            If Not RSObat.EOF Then
                Dim Kurangi As String
                Kurangi = "update Obat set jumlahObt='" & RSObat!JUMLAHOBT - ado.Recordset!dosis & "' where kodeObt='" & ado.Recordset!Kode & "'"
                Conn.Execute (Kurangi)
            End If
        End If
    ado.Recordset.MoveNext
    Loop
    
    simpanbyr = "insert into pembayaran(nomorbyr,kodepsn,tanggalbyr,jumlahBYR) values ('" & Combo1 & "','" & KodePsn & "','" & Tanggal & "','" & Total & "')"
    Conn.Execute simpanbyr
    
    Bersihkan
    Form_Activate
    Combo1.SetFocus
    Call Cetak
    'Call CetakCR
    
End Sub

Sub CetakCR()
    CR.ReportFileName = App.Path & "\buktiresep.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub CmdBatal_Click()
    Bersihkan
    Form_Activate
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

Function Jumlah()
    Set TTlHarga = New ADODB.Recordset
    TTlHarga.Open "select sum(dosis) as JumTotal from Temporer", Conn
    Jumlah = TTlHarga!JumTotal
End Function

Function Jumlah2()
    Set TTlHarga = New ADODB.Recordset
    TTlHarga.Open "select sum(subtotal) as JumTotal from Temporer", Conn
    Jumlah2 = TTlHarga!JumTotal
End Function

Function Cetak()
Call Koneksi
RSResep.Open "select * from Resep Where Nomorrsp In(Select Max(Nomorrsp)From Resep)Order By Nomorrsp Desc", Conn
Layar.Show
Dim MGrs As String
Layar.Font = "Courier New"
Layar.Print
Layar.Print
RSPasien.Open "select * From pasien where KODEPSN= '" & RSResep!KodePsn & "'", Conn
RSDokter.Open "select * From Dokter where Kodedkt= '" & RSResep!Kodedkt & "'", Conn
RSPoli.Open "select * From poli where kodepl= '" & RSResep!Kodepl & "'", Conn

Layar.Print Tab(5); "Nomorrsp   :   "; RSResep!nomorrsp
Layar.Print Tab(5); "Tanggal    :   "; Format(RSResep!TanggalRsp, "DD-MMM-YY")
Layar.Print Tab(5); "Dokter     :   "; RSDokter!Namadkt
Layar.Print Tab(5); "Pasien     :   "; RSPasien!NamaPsn
Layar.Print Tab(5); "Poli       :   "; RSPoli!Namapl
MGrs = String$(33, "-")
Layar.Print Tab(5); MGrs
RSDetail.Open "select * from Detail Where Nomorrsp='" & RSResep!nomorrsp & "'", Conn
RSDetail.MoveFirst
No = 0
Do While Not RSDetail.EOF
    No = No + 1
    Set RSObat = New ADODB.Recordset
    RSObat.Open "select * From Obat where KodeObt= '" & RSDetail!KODEOBT & "'", Conn
    RSObat.Requery
    Layar.Print Tab(5); No; Space(2); RSObat!NamaOBT
    Layar.Print Tab(10); RKanan(RSDetail!dosis, "###"); Space(1); "X";
    Layar.Print Tab(15); Format(RSObat!hargaobt, "###,###,###");
    Layar.Print Tab(25); RKanan(RSDetail!dosis * RSObat!hargaobt, "###,###,###")
    RSDetail.MoveNext
Loop
'==========================
Layar.Print Tab(5); MGrs
Layar.Print Tab(5); "Total      :";
Layar.Print Tab(25); RKanan(RSResep!TotalHRG, "###,###,###");
Layar.Print Tab(5); "Dibayar    :";
Layar.Print Tab(25); RKanan(RSResep!Dibayar, "###,###,###");
Layar.Print Tab(5); MGrs
Layar.Print Tab(5); "Kembali    :";
If RSResep!Dibayar = RSResep!TotalHRG Then
    Layar.Print Tab(34); RSResep!Dibayar - RSResep!TotalHRG
Else
    Layar.Print Tab(25); RKanan(RSResep!Dibayar - RSResep!TotalHRG, "###,###,###");
End If
Layar.Print Tab(5); MGrs
Layar.Print Tab(5); "Semoga Lekas Sembuh"
Layar.Print
Layar.Print
Layar.Print
Conn.Close

End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function


