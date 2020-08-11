VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pendaftaran 
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
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
   ScaleHeight     =   6075
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdPasienBaru 
      Caption         =   "Pasien Baru"
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      Top             =   3480
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1000
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   400
      Left            =   1080
      TabIndex        =   1
      Top             =   3480
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   400
      Left            =   2040
      TabIndex        =   2
      Top             =   3480
      Width           =   1000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   3000
      TabIndex        =   3
      Top             =   3480
      Width           =   1000
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   6120
      TabIndex        =   6
      Top             =   600
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   2000
   End
   Begin VB.TextBox Telepon 
      Height          =   350
      Left            =   6120
      TabIndex        =   11
      Top             =   2400
      Width           =   2000
   End
   Begin VB.TextBox Umur 
      Height          =   350
      Left            =   6120
      TabIndex        =   10
      Top             =   2040
      Width           =   2000
   End
   Begin VB.TextBox Gender 
      Height          =   350
      Left            =   6120
      TabIndex        =   9
      Top             =   1680
      Width           =   2000
   End
   Begin VB.TextBox Alamat 
      Height          =   350
      Left            =   6120
      TabIndex        =   8
      Top             =   1320
      Width           =   2000
   End
   Begin VB.TextBox Nama 
      Height          =   350
      Left            =   6120
      TabIndex        =   7
      Top             =   960
      Width           =   2000
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   1845
      Left            =   120
      TabIndex        =   27
      Top             =   4080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3254
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   330
      Left            =   120
      Top             =   6240
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Pendaftaran"
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
      TabIndex        =   29
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label Biaya 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   6120
      TabIndex        =   26
      Top             =   2880
      Width           =   1995
   End
   Begin VB.Label NomorAntri 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2040
      TabIndex        =   25
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label TanggalDft 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2040
      TabIndex        =   24
      Top             =   960
      Width           =   1995
   End
   Begin VB.Label NomorDft 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2040
      TabIndex        =   23
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biaya"
      Height          =   345
      Left            =   4200
      TabIndex        =   22
      Top             =   2880
      Width           =   1755
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Poli"
      Height          =   345
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   1755
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telepon"
      Height          =   345
      Left            =   4200
      TabIndex        =   20
      Top             =   2400
      Width           =   1755
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Umur"
      Height          =   345
      Left            =   4200
      TabIndex        =   19
      Top             =   2040
      Width           =   1755
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gender"
      Height          =   345
      Left            =   4200
      TabIndex        =   18
      Top             =   1680
      Width           =   1755
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alamat"
      Height          =   345
      Left            =   4200
      TabIndex        =   17
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Pasien"
      Height          =   345
      Left            =   4200
      TabIndex        =   16
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Pasien"
      Height          =   345
      Left            =   4200
      TabIndex        =   15
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nomor Antrian"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1755
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nomor Pendaftaran"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1755
   End
End
Attribute VB_Name = "Pendaftaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
ado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBRAWATJALAN.mdb"
ado.RecordSource = "Pendaftaran"
ado.Refresh
Set dg.DataSource = ado
dg.Refresh
Call Auto
TanggalDft = Format(Date, "dd-mm-yyyy")
Call Koneksi
RSPoli.Open "poli", Conn
Combo1.Clear
Do Until RSPoli.EOF
    Combo1.AddItem RSPoli!Kodepl & Space(5) & RSPoli!Namapl
    RSPoli.MoveNext
Loop

RSPasien.Open "pasien", Conn
Combo2.Clear
Do Until RSPasien.EOF
    Combo2.AddItem RSPasien!KodePsn & Space(5) & RSPasien!NamaPsn
    RSPasien.MoveNext
Loop
CmdPasienBaru.Enabled = False
End Sub

Private Sub Form_Load()
Call KondisiAwal
End Sub

Private Sub CmdPasienBaru_Click()
Call PsnBaru
Call KosongPsn
Call BukaPsn
Combo2.Enabled = False
Nama.SetFocus
End Sub

Private Sub Auto()
Call Koneksi
RSPendaftaran.Open "select * from Pendaftaran Where NomorDft In(Select Max(NomorDft)From Pendaftaran)Order By NomorDft Desc", Conn
RSPendaftaran.Requery
Dim Urutan As String * 10
Dim Hitung As Long
With RSPendaftaran
    If .EOF Then
        Urutan = Format(Date, "yymmdd") + "0001"
        NomorDft = Urutan
    Else
        If Left(!NomorDft, 6) <> Format(Date, "yymmdd") Then
            Urutan = Format(Date, "yymmdd") + "0001"
        Else
            Hitung = (!NomorDft) + 1
            Urutan = Format(Date, "yymmdd") + Right("0000" & Hitung, 4)
        End If
    End If
    NomorDft = Urutan
End With
End Sub

Private Sub CmdBatal_Click()
Call KondisiAwal
List1.Clear
End Sub

Private Sub CmdInput_Click()
Call Terang
CmdPasienBaru.Enabled = True
Combo1.SetFocus
End Sub

Private Sub CmdSimpan_Click()
If Combo1 = "" Or Combo2 = "" Or Nama = "" Or Alamat = "" Or Gender = "" Or Umur = "" Or Telepon = "" Then
    MsgBox "Data belum lengkap"
Else
    Call Koneksi
    RSPasien.Open "select * from pasien where kodepsn='" & Left(Combo2, 8) & "'", Conn
    If RSPasien.EOF Then
        simpanpasien = "insert into pasien (kodepsn,namapsn,alamatpsn,genderpsn,umurpsn,teleponpsn) values " & _
        "('" & Combo2 & "','" & Nama & "','" & Alamat & "','" & Gender & "','" & Umur & "','" & Telepon & "')"
        Conn.Execute simpanpasien
    End If
    
    Simpan = "insert into pendaftaran(nomordft,tanggaldft,kodedkt,kodepsn,kodepl,kodepmk,biaya,ket) values " & _
    "('" & NomorDft & "','" & TanggalDft & "','" & Left(List1, 5) & "','" & Left(Combo2, 8) & "','" & Left(Combo1, 2) & "','" & Menu.STBar.Panels(3).Text & "','" & Biaya & "',0)"
    Conn.Execute Simpan
    Form_Activate
    Call KondisiAwal
End If
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
Call Koneksi
RSDokter.Open "select * from dokter where kodepl='" & Left(Combo1, 2) & "'", Conn
If Not RSDokter.EOF Then
    List1.Clear
    Do While Not RSDokter.EOF
        List1.AddItem RSDokter!Kodedkt & Space(5) & RSDokter!Namadkt
        RSDokter.MoveNext
    Loop
Else
    MsgBox "kode poli belum terdaftar"
    Combo1.SetFocus
    List1.Clear
End If

End Sub

Private Sub Combo2_Click()
Call Koneksi
RSPasien.Open "select * from pasien where kodepsn='" & Left(Combo2, 8) & "'", Conn
If Not RSPasien.EOF Then
    Nama.Enabled = False
    Alamat.Enabled = False
    Gender.Enabled = False
    Umur.Enabled = False
    Telepon.Enabled = False
    Nama = RSPasien!NamaPsn
    Alamat = RSPasien!alamatpsn
    Gender = RSPasien!genderpsn
    Umur = RSPasien!umurpsn
    Telepon = RSPasien!teleponpsn
Else
    MsgBox "ini pasien baru"
    Nama.SetFocus
End If
End Sub



Private Sub combo2_keypress(Keyascii As Integer)
If Keyascii = 13 Then
    If Combo2 = "" Then
        MsgBox "kode pasien harus diisi"
        Combo2.SetFocus
        Exit Sub
    Else
        Call Koneksi
        RSPasien.Open "select * from pasien where kodepsn='" & Left(Combo2, 8) & "'", Conn
        If Not RSPasien.EOF Then
            Nama = RSPasien!NamaPsn
            Alamat = RSPasien!alamatpsn
            Gender = RSPasien!genderpsn
            Umur = RSPasien!alamatpsn
            Telepon = RSPasien!teleponpsn
        Else
            Call KosongPsn
            MsgBox "kode pasien tidak terdaftar"
            Combo2.SetFocus
        End If
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub


Private Sub PsnBaru()
Call Koneksi
RSPasien.Open "select * from pASIEN Where KODEPSN In(Select Max(kodepsn)From pASIEN)Order By kodepsn Desc", Conn
RSPasien.Requery
If RSPasien.EOF Then
    Urutan = Format(Date, "yymmdd") + "01"
    Combo2 = Urutan
    Text1 = Urutan
Else
    If Left(RSPasien!KodePsn, 6) <> Format(Date, "yymmdd") Then
        Combo2 = Format(Date, "yymmdd") + "01"
    Else
        Hitung = (RSPasien!KodePsn) + 1
        Urutan = Format(Date, "yymmdd") + Right("00" & Hitung, 2)
        Combo2 = Urutan
    End If
End If
End Sub

Sub BukaPsn()
Nama.Enabled = True
Alamat.Enabled = True
Gender.Enabled = True
Umur.Enabled = True
Telepon.Enabled = True
End Sub

Sub Blank()
Combo1 = ""
Combo2 = ""
NomorAntri = ""
Call KosongPsn
Biaya = ""
End Sub

Sub Gelap()
Combo1.Enabled = False
Combo2.Enabled = False
Nama.Enabled = False
Alamat.Enabled = False
Gender.Enabled = False
Umur.Enabled = False
Telepon.Enabled = False
End Sub

Sub Terang()
Combo1.Enabled = True
Combo2.Enabled = True
End Sub

Sub KondisiAwal()
Call Blank
Call Gelap
CmdPasienBaru.Enabled = False
Form_Activate
List1.Clear
End Sub

Private Sub List1_Click()
Call Koneksi
RSDokter.Open "select * from dokter where kodedkt='" & Left(List1, 5) & "'", Conn
If Not RSDokter.EOF Then
    Biaya = RSDokter!TARIF
Else
    MsgBox "kode dokter tidak terdaftar"
End If

RSPendaftaran.Open "SELECT COUNT(KODEDKT) AS ANTRI FROM pendaftaran where cdate(tanggaldft)='" & Date & "' and kodedkt='" & Left(List1, 5) & "'", Conn
RSPendaftaran.Requery
If RSPendaftaran!antri = 0 Then
    NomorAntri = "1"
Else
    Hitung = RSPendaftaran!antri + 1
    NomorAntri = Right("0" & Hitung, 1)
    Exit Sub
End If
End Sub

Sub KosongPsn()
Nama = ""
Alamat = ""
Gender = ""
Umur = ""
Telepon = ""
End Sub

Private Sub Nama_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Nama = "" Then
        MsgBox "Nama harus diisi"
        Nama.SetFocus
    Else
        Alamat.SetFocus
    End If
End If
End Sub

Private Sub alamat_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Alamat = "" Then
        MsgBox "alamat harus diisi"
        Alamat.SetFocus
    Else
        Gender.SetFocus
    End If
End If
End Sub

Private Sub gender_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Gender = "" Then
        MsgBox "gender harus diisi"
        Gender.SetFocus
    Else
        Umur.SetFocus
    End If
End If
End Sub

Private Sub umur_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Umur = "" Then
        MsgBox "umur harus diisi"
        Umur.SetFocus
    Else
        Telepon.SetFocus
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub telepon_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Telepon = "" Then
        MsgBox "telepon harus diisi"
        Telepon.SetFocus
    Else
        CmdSimpan.SetFocus
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

