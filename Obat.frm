VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Obat 
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
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
   ScaleHeight     =   4455
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   2000
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Width           =   2000
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   2280
      TabIndex        =   2
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   400
      Left            =   3360
      TabIndex        =   3
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   400
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   5520
      TabIndex        =   5
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   4920
      TabIndex        =   11
      Top             =   1320
      Width           =   2000
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   4920
      TabIndex        =   10
      Top             =   960
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Width           =   2000
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "Obat.frx":0000
      Height          =   1845
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   6855
      _ExtentX        =   12091
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "KodeObt"
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
         DataField       =   "NamaObt"
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
         DataField       =   "JenisObt"
         Caption         =   "Jenis"
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
         DataField       =   "Katagori"
         Caption         =   "Katagori"
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
         DataField       =   "HargaObt"
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
      BeginProperty Column05 
         DataField       =   "JumlahObt"
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
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   330
      Left            =   120
      Top             =   4560
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Obat"
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
      TabIndex        =   19
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah"
      Height          =   345
      Left            =   3600
      TabIndex        =   17
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label Label5 
      Caption         =   "Harga"
      Height          =   345
      Left            =   3600
      TabIndex        =   16
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Katagori"
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "Jenis"
      Height          =   345
      Left            =   3600
      TabIndex        =   14
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Obat"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Obat"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1245
   End
End
Attribute VB_Name = "Obat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Koneksi
ado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBrawatjalan.mdb"
ado.RecordSource = "select * from Obat"
ado.Refresh
Set dg.DataSource = ado
dg.Refresh

RSObat.Open "select DISTINCT JENISOBT from Obat", Conn
Combo2.Clear
Do While Not RSObat.EOF
    Combo2.AddItem RSObat!JenisObt
    RSObat.MoveNext
Loop

RSPoli.Open "select * from poli", Conn
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
    Text3.MaxLength = 8
    Text4.MaxLength = 4
    KondisiAwal
End Sub

Private Sub Combo1_Click()
Text1.Enabled = False
Text1 = Left(Combo1, 2)
Call KODEOTO
End Sub

Private Sub KODEOTO()
Call Koneksi
RSObat.Open "SELECT count(katagori) as ketemu FROM Obat where katagori='" & Combo1 & "'", Conn
RSObat.Requery
If RSObat!ketemu = 0 Then
    Text1 = Left(Combo1, 3) + "01"
Else
    Hitung = RSObat!ketemu + 1
    Text1 = Left(Combo1, 3) + Right("00" & Hitung, 2)
    Exit Sub
End If
End Sub

Private Sub DG_Keypress(Keyascii As Integer)
If Keyascii = 13 Then
    If CmdEdit.Enabled = True Then
        Combo1.Enabled = False
        Text1.Enabled = False
        Combo1 = dg.Columns(3)
        Text1 = dg.Columns(0)
        Text2 = dg.Columns(1)
        Combo2 = dg.Columns(2)
        Text3 = dg.Columns(4)
        Text4 = dg.Columns(5)
        Text2.SetFocus
    End If
    
    If CmdHapus.Enabled = True Then
         Combo1 = dg.Columns(3)
        Text1 = dg.Columns(0)
        Text2 = dg.Columns(1)
        Combo2 = dg.Columns(2)
        Text3 = dg.Columns(4)
        Text4 = dg.Columns(5)
        Call CariData
        If Not RSObat.EOF Then
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                hapus = "delete * from Obat"
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
    RSObat.Open "Select * From Obat where KodeObt='" & Text1 & "'", Conn
End Function

Private Sub CmdBatal_Click()
KosongkanText
TidakSiapIsi
KondisiAwal
End Sub

Private Sub CmdSimpan_Click()
If Text1 = "" Or Text2 = "" Or Combo2 = "" Or Combo1 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "Data Belum Lengkap...!"
    Exit Sub
Else
    If CmdInput.Enabled = True Then
        Dim SQLTambah As String
        SQLTambah = "Insert Into Obat (KodeObt,NamaObt,JenisObt,Katagori,HargaObt,JumlahObt) values " & _
        "('" & Text1 & "','" & Text2 & "','" & Combo2 & "','" & Combo1 & "','" & Text3 & "','" & Text4 & "')"
        Conn.Execute SQLTambah
    ElseIf CmdEdit.Enabled = True Then
        Dim SQLEdit As String
        SQLEdit = "Update Obat Set NamaObt= '" & Text2 & "', JenisObt = '" & Combo2 & "',Katagori = '" & Combo1 & "',hargaobt = '" & Text3 & "',jumlahobt= '" & Text4 & "' where KodeObt='" & Text1 & "'"
        Conn.Execute SQLEdit
    End If
    Form_Activate
    KondisiAwal
End If
End Sub

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Combo1 = ""
    Combo2 = ""
    Text3 = ""
    Text4 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Combo1.Enabled = True
    Combo2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Combo1.Enabled = False
    Combo2.Enabled = False
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
Text2 = RSObat!NamaOBT
Combo1 = RSObat!katagori
Combo2 = RSObat!JenisObt
Text3 = RSObat!hargaobt
Text4 = RSObat!JUMLAHOBT
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
    If Text1 = "" Then
        MsgBox "KODE OBAT HARUS DIISI"
        Text1.SetFocus
        Exit Sub
    End If
    
    If CmdInput.Enabled = True Then
        Call CariData
            If Not RSObat.EOF Then
                TampilkanData
                MsgBox "Kode OBAT  Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Enabled = True Then
        Call CariData
            If Not RSObat.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode oBAT Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSObat.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From oBAT where KodeObt= '" & Text1 & "'"
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
    If Keyascii = 13 Then Combo2.SetFocus
End Sub

Private Sub combo2_keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then
        If Combo2 = "" Then
            MsgBox "jenis harus diisi"
            Combo2.SetFocus
        Else
            Text3.SetFocus
        End If
    End If
End Sub


Private Sub combo1_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then
        If Combo1 = "" Then
            MsgBox "Katagori harus diisi"
            Combo1.SetFocus
        Else
            Call Koneksi
            RSPoli.Open "SELECT * FROM POLI where namapl='" & Combo1 & "'", Conn
            If RSPoli.EOF Then
                MsgBox "Katagori obat tidak sesuai dengan poli"
                Combo1.SetFocus
                Combo1 = ""
            Else
                Text2.SetFocus
            End If
        End If
    End If
End Sub


Private Sub Text3_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text4.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub



Private Sub Text4_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdSimpan.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdSimpan.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub


