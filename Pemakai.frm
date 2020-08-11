VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pemakai 
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
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
   ScaleHeight     =   4725
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   2000
   End
   Begin VB.TextBox KodeDasar 
      Height          =   350
      Left            =   3840
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1680
      TabIndex        =   8
      Top             =   1320
      Width           =   4260
   End
   Begin VB.TextBox Text3 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "X"
      TabIndex        =   9
      Top             =   1680
      Width           =   2000
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   1000
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   1000
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   400
      Left            =   3000
      TabIndex        =   3
      Top             =   2160
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   400
      Left            =   3960
      TabIndex        =   4
      Top             =   2160
      Width           =   1000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   4920
      TabIndex        =   5
      Top             =   2160
      Width           =   1000
   End
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "Pemakai.frx":0000
      Height          =   1845
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   5895
      _ExtentX        =   10398
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "KodePmk"
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
         DataField       =   "NamaPmk"
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
         DataField       =   "PassPmk"
         Caption         =   "Password"
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
         DataField       =   "StatusPmk"
         Caption         =   "Status"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   330
      Left            =   120
      Top             =   4800
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Data User"
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
      TabIndex        =   16
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Pemakai"
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Pemakai"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1500
   End
End
Attribute VB_Name = "Pemakai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Koneksi
ado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBrawatjalan.mdb"
ado.RecordSource = "pemakai"
ado.Refresh
Set dg.DataSource = ado
dg.Refresh
End Sub

Sub Form_Load()
    Call Koneksi
    Text1.MaxLength = 5
    Text2.MaxLength = 30
    Text3.MaxLength = 15
    KondisiAwal
    Combo1.AddItem "ADM"
    Combo1.AddItem "APOTEKER"
    Combo1.AddItem "ADMINISTRATOR"
End Sub

Private Sub Combo1_Click()
Text1.Enabled = False
 If Combo1 = "ADM" Then
        KodeDasar = "ADM"
        Call KODEOTO
    ElseIf Combo1 = "APOTEKER" Then
        KodeDasar = "APT"
        Call KODEOTO
    ElseIf Combo1 = "ADMINISTRATOR" Then
        KodeDasar = "ADS"
        Call KODEOTO
    End If
End Sub

Private Sub combo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Combo1 = "" Then
        MsgBox "Status harus diisi"
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
        Combo1 = dg.Columns(3)
        Text1 = dg.Columns(0)
        Text2 = dg.Columns(1)
        Text3 = dg.Columns(2)
        Text2.SetFocus
    End If
    
    If CmdHapus.Enabled = True Then
        Combo1 = dg.Columns(3)
        Text1 = dg.Columns(0)
        Text2 = dg.Columns(1)
        Text3 = dg.Columns(2)
        Call CariData
        If Not RSPemakai.EOF Then
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                hapus = "delete * from pemakai"
                Conn.Execute hapus
                Call KondisiAwal
            Else
                Call KondisiAwal
            End If
        End If
    End If
End If
End Sub

Private Sub KODEOTO()
Call Koneksi
RSPemakai.Open "SELECT count(statuspmk) as ketemu FROM pemakai where statuspmk='" & Combo1 & "'", Conn
RSPemakai.Requery
If RSPemakai!ketemu = 0 Then
    Text1 = KodeDasar + "01"
Else
    Hitung = RSPemakai!ketemu + 1
    Text1 = KodeDasar + Right("00" & Hitung, 2)
End If
End Sub

Function CariData()
    Call Koneksi
    RSPemakai.Open "Select * From Pemakai where KodePmk='" & Text1 & "'", Conn
End Function

Private Sub CmdBatal_Click()
KosongkanText
TidakSiapIsi
KondisiAwal
End Sub

Private Sub CmdSimpan_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
    MsgBox "Data Belum Lengkap...!"
    Exit Sub
Else
    If CmdInput.Enabled = True Then
            Dim SQLTambah1 As String
            SQLTambah1 = "Insert Into Pemakai (KodePmk,NamaPmk,PassPmk,StatusPmk) values " & _
            "('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "')"
            Conn.Execute SQLTambah1
    Else
            Dim SQLEdit As String
            SQLEdit = "Update Pemakai Set NamaPmk= '" & Text2 & "', PassPmk = '" & Text3 & "',StatusPmk = '" & Combo1 & "' where KodePmk='" & Text1 & "'"
            Conn.Execute SQLEdit
        End If
    Form_Activate
    KondisiAwal
End If
End Sub

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Combo1 = ""
    KodeDasar = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Combo1.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Combo1.Enabled = False
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
Text2 = RSPemakai!NamaPMK
Text3 = RSPemakai!PassPMK
Combo1 = RSPemakai!StatusPMK

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
    Combo1.Enabled = False
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
    Combo1.Enabled = False
    If Left(Text1, 3) <> "ADM" And Left(Text1, 3) <> "APT" And Left(Text1, 3) <> "ADS" Then
        MsgBox "Tiga Digit Pertama Harus ADM, APT atau ADS"
        Text1.SetFocus
        Exit Sub
    End If

    If CmdInput.Enabled = True Then
        Call CariData
            If Not RSPemakai.EOF Then
                TampilkanData
                MsgBox "Kode Pemakai Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Enabled = True Then
        Call CariData
            If Not RSPemakai.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Pemakai Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSPemakai.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Pemakai where KodePmk= '" & Text1 & "'"
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
Text3.PasswordChar = "X"
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdSimpan.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdSimpan.SetFocus
        End If
    End If

End Sub

