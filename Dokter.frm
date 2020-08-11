VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Dokter 
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5280
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
   ScaleHeight     =   5850
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   2640
      Width           =   1500
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   4320
      TabIndex        =   5
      Top             =   3240
      Width           =   800
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   400
      Left            =   3480
      TabIndex        =   4
      Top             =   3240
      Width           =   800
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   400
      Left            =   2640
      TabIndex        =   3
      Top             =   3240
      Width           =   800
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   1800
      TabIndex        =   2
      Top             =   3240
      Width           =   800
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   960
      TabIndex        =   1
      Top             =   3240
      Width           =   800
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   800
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   1440
      TabIndex        =   10
      Top             =   2280
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "Dokter.frx":0000
      Height          =   1845
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   5055
      _ExtentX        =   8916
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "KodeDkt"
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
         DataField       =   "NamaDkt"
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
         DataField       =   "Spesialis"
         Caption         =   "Spesialis"
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
         DataField       =   "AlamatDkt"
         Caption         =   "Alamat"
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
         DataField       =   "TeleponDkt"
         Caption         =   "Telepon"
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
         DataField       =   "KodePl"
         Caption         =   "Kode Poli"
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
      BeginProperty Column06 
         DataField       =   "Tarif"
         Caption         =   "Tarif"
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
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   330
      Left            =   120
      Top             =   6000
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Dokter"
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
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label KodePoli 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3240
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tarif"
      Height          =   345
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telepon"
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alamat"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Dokter"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Dokter"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Poli"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1245
   End
End
Attribute VB_Name = "Dokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Koneksi
ado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBrawatjalan.mdb"
ado.RecordSource = "select * from Dokter"
ado.Refresh
Set dg.DataSource = ado
dg.Refresh

RSPoli.Open "select * from poli order by namapl", Conn
Combo1.Clear
Do While Not RSPoli.EOF
    Combo1.AddItem RSPoli!Namapl
    RSPoli.MoveNext
    Loop
End Sub

Sub Form_Load()
    Call Koneksi
    Text1.MaxLength = 5
    Text2.MaxLength = 30
    Text3.MaxLength = 30
    Text4.MaxLength = 15
    Text5.MaxLength = 8
    KondisiAwal
End Sub


Private Sub Combo1_Click()
Call Koneksi
RSPoli.Open "select * from poli where namapl='" & Combo1 & "'", Conn
KodePoli = RSPoli!Kodepl
Call KODEOTO
End Sub


Private Sub KODEOTO()
Call Koneksi
RSDokter.Open "SELECT count(spesialis) as ketemu FROM DOKTEr where spesialis='" & Combo1 & "'", Conn
RSDokter.Requery
If RSDokter!ketemu = 0 Then
    Text1 = KodePoli + "001"
Else
    Hitung = RSDokter!ketemu + 1
    Text1 = KodePoli + Right("000" & Hitung, 3)
    Exit Sub
End If
End Sub

Private Sub combo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Combo1 = "" Then
        Combo1.SetFocus
    Else
        Text2.SetFocus
    End If
End If
End Sub

Private Sub DG_Keypress(Keyascii As Integer)
If Keyascii = 13 Then
    If CmdEdit.Enabled = True Then
        Combo1.Enabled = False
        Text1.Enabled = False
        Combo1 = dg.Columns(5)
        Text1 = dg.Columns(0)
        Text2 = dg.Columns(1)
        Text3 = dg.Columns(2)
        Text4 = dg.Columns(3)
        Text5 = dg.Columns(4)
        Text2.SetFocus
    End If
    
    If CmdHapus.Enabled = True Then
        Combo1 = dg.Columns(5)
        Text1 = dg.Columns(0)
        Text2 = dg.Columns(1)
        Text3 = dg.Columns(2)
        Text4 = dg.Columns(3)
        Text5 = dg.Columns(4)
        Call CariData
        If Not RSDokter.EOF Then
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                hapus = "delete * from Dokter"
                Conn.Execute hapus
                Call KondisiAwal
            Else
                Call KondisiAwal
            End If
        End If
    End If
End If
End Sub

Function CariData()
    Call Koneksi
    RSDokter.Open "Select * From Dokter where KodeDkt='" & Text1 & "'", Conn
End Function

Private Sub CmdBatal_Click()
KosongkanText
TidakSiapIsi
KondisiAwal
End Sub

Private Sub CmdSimpan_Click()
If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
    MsgBox "Data Belum Lengkap...!"
    Exit Sub
Else
    If CmdInput.Enabled = True Then
        Dim SQLTambah As String
        SQLTambah = "Insert Into Dokter (KodePL,KodeDkt,NamaDkt,Spesialis,AlamatDkt,TeleponDkt,Tarif) values " & _
        "('" & KodePoli & "','" & Text1 & "','" & Text2 & "','" & Combo1 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "')"
        Conn.Execute SQLTambah
    ElseIf CmdEdit.Enabled = True Then
        Dim SQLEdit As String
        SQLEdit = "Update Dokter Set NamaDkt= '" & Text2 & "', AlamatDkt = '" & Text3 & "',TeleponDkt = '" & Text4 & "',tarif='" & Text5 & "' where KodeDkt='" & Text1 & "'"
        Conn.Execute SQLEdit
    End If
    Form_Activate
    KondisiAwal
End If
End Sub

Private Sub KosongkanText()
    Combo1 = ""
    KodePoli = ""
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
End Sub

Private Sub SiapIsi()
    Combo1.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Combo1.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
End Sub

Private Sub KondisiAwal()
KosongkanText
TidakSiapIsi
CmdInput.Enabled = True
CmdEdit.Enabled = True
CmdHapus.Enabled = True
CmdSimpan.Enabled = False
CmdBatal.Enabled = False
CmdTutup.Enabled = True
End Sub

Private Sub TampilkanData()
On Error Resume Next
Combo1 = RSDokter!spesialis
Text1 = RSDokter!Kodedkt
Text2 = RSDokter!Namadkt
Text3 = RSDokter!AlamatDkt
Text4 = RSDokter!TeleponDkt
Text5 = RSDokter!TARIF
End Sub


Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdSimpan.Enabled = True
        CmdBatal.Enabled = True
        CmdTutup.Enabled = False
        SiapIsi
        KosongkanText
        Combo1.SetFocus
    End If
End Sub

Private Sub CmdEdit_Click()
If CmdEdit.Caption = "&Edit" Then
    CmdInput.Enabled = False
    CmdHapus.Enabled = False
    CmdTutup.Enabled = False
    CmdSimpan.Enabled = True
    CmdBatal.Enabled = True
    SiapIsi
    Text1.SetFocus
End If
End Sub

Private Sub CmdHapus_Click()
If CmdHapus.Caption = "&Hapus" Then
    CmdTutup.Enabled = False
    CmdInput.Enabled = False
    CmdEdit.Enabled = False
    CmdBatal.Enabled = True
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

    If CmdInput.Enabled = True Then
        Call CariData
            If Not RSDokter.EOF Then
                TampilkanData
                MsgBox "Kode Dokter Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Enabled = True Then
        Call CariData
            If Not RSDokter.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Dokter Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSDokter.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Dokter where KodeDkt= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    Form_Activate
                    KondisiAwal
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
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then Text5.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub text5_keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdSimpan.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdSimpan.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

