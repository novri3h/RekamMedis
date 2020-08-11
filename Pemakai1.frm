VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pemakai1 
   Caption         =   "Pemakai"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   5520
      TabIndex        =   8
      Top             =   600
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   400
      Left            =   4440
      TabIndex        =   7
      Top             =   600
      Width           =   1000
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   400
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   1000
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   400
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   1995
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3519
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Left            =   3360
      Top             =   1200
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
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Pemakai"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Pemakai"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1500
   End
End
Attribute VB_Name = "Pemakai1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Call KODEOTO
'Call Koneksi
'RSPemakai.Open "select * from pemakai where statuspmk='" & Combo1 & "'", Conn
'If RSPemakai.EOF Then
'    If Combo1 = "ADM" Then
'        Text1 = "ADM" + 1
'    ElseIf Combo1 = "APOTEKER" Then
'        Text1 = "APT" + 1
'    ElseIf Combo1 = "ADMINISTRATOR" Then
'        Text1 = "ADR" + 1
'    End If
'End If
End Sub

Private Sub KODEOTO()
'SELECT MAX(RIGHT(KODEPMK,2)) FROM PEMAKAI WHERE STATUSPMK="ADM"
Call Koneksi
RSPemakai.Open "select * FROM PEMAKAI Where KodePMK In(Select Max(RIGHT(KodePMK,2)) From Pemakai)", Conn
RSPemakai.Requery
    Dim Hitung As Long
    If RSPemakai.EOF And Combo1 = "ADM" Then
        Text1 = "ADM" + "01"
    Else
        Hitung = Right(RSPemakai!KodePMK, 2) + 1
        Text1 = "ADM" + Right("00" & Hitung, 2)
    End If
'Conn.Close
'
'Call Koneksi
'RSPemakai.Open "select * FROM PEMAKAI Where KodePMK In(Select Max(KodePMK) From Pemakai)", Conn
'RSPemakai.Requery
'
'    If RSPemakai.EOF And Combo1 = "APOTEKER" Then
'        Text1 = "APT" + "01"
'    Else
'        Hitung = Right(RSPemakai!KodePMK, 2) + 1
'        Text1 = "APT" + Right("00" & Hitung, 2)
'    End If
'
'Conn.Close
'Call Koneksi
'RSPemakai.Open "select * FROM PEMAKAI Where KodePMK In(Select Max(KodePMK) From Pemakai)", Conn
'RSPemakai.Requery
'
'    If RSPemakai.EOF And Combo1 = "ADMINISTRATOR" Then
'        Text1 = "ADS" + "01"
'    Else
'        Hitung = Right(RSPemakai!KodePMK, 2) + 1
'        Text1 = "ADS" + Right("00" & Hitung, 2)
'    End If
    
End Sub

Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBrawatjalan.mdb"
ADO.RecordSource = "select * from Pemakai"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
End Sub

Sub Form_Load()
    Call Koneksi
    Text1.MaxLength = 6
    Text2.MaxLength = 30
    
    KondisiAwal
    Combo1.AddItem "ADM"
    Combo1.AddItem "APT"
    Combo1.AddItem "ADS"
End Sub

Function CariData()
    Call Koneksi
    RSPemakai.Open "Select * From Pemakai where KodePMK='" & Text1 & "'", Conn
End Function
Private Sub cmdbatal_Click()
KosongkanText
TidakSiapIsi
KondisiAwal
End Sub

Private Sub cmdSimpan_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
    MsgBox "Data Belum Lengkap...!"
    Exit Sub
Else
    If CmdInput.Enabled = True Then
        Dim SQLTambah As String
        SQLTambah = "Insert Into Pemakai (KodePMK,NamaPMK,PassPMK,StatusPMK) values " & _
        "('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "')"
        Conn.Execute SQLTambah
    ElseIf CmdEdit.Enabled = True Then
        Dim SQLEdit As String
        SQLEdit = "Update Pemakai Set NamaPMK= '" & Text2 & "', PassPMK = '" & Text3 & "',StatusPMK = '" & Combo1 & "' where KodePMK='" & Text1 & "'"
        Conn.Execute SQLEdit
        
        Dim SQLEditpemakai As String
        SQLEditpemakai = "Update pemakai Set NamaPMK= '" & Text2 & "' where KodePMK='" & Text1 & "'"
        Conn.Execute SQLEditpemakai
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
        'Call AutoNomor
        'Text1.Enabled = False
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

Private Sub cmdtutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
'    If Len(Text1) < 4 Then
'        MsgBox "Kode Harus 4 Digit"
'        Text1.SetFocus
'    Else
'        Text2.SetFocus
'    End If

    If CmdInput.Enabled = True Then
        Call CariData
            If Not RSPemakai.EOF Then
                TampilkanData
                MsgBox "Kode ADM Sudah Ada"
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
                MsgBox "Kode ADM Tidak Ada"
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
                    SQLHapus = "Delete From ADM where KodePMK= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    
                    Dim SQLHapuspemakai As String
                    SQLHapuspemakai = "Delete From pemakai where KodePMK= '" & Text1 & "'"
                    Conn.Execute SQLHapuspemakai
                    
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

Private Sub text2_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call Koneksi
        RSPemakai.Open "select * from Pemakai where namaPMK='" & Text2 & "'", Conn
        If Not RSPemakai.EOF Then
            MsgBox "nama '" & Text2 & "' sudah ada, coba tambahkan karakter pembeda"
            Text2.SetFocus
        Else
            Text3.SetFocus
        End If
    End If
End Sub

Private Sub text3_keypress(KeyAscii As Integer)
    Text3.PasswordChar = "X"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Combo1.SetFocus
End Sub

'Private Sub COMBO1_keypress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 13 Then TxtAlamat.SetFocus
'End Sub
'
'Private Sub TxtAlamat_keypress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 13 Then TxtTelepon.SetFocus
'End Sub
'
'Private Sub Txttelepon_keypress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 13 Then TxtTglDaftar.SetFocus
'    TxtTglDaftar = Date
'    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
'End Sub

'Private Sub COMBO1_keypress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If CmdInput.Enabled = True Then
'            CmdSimpan.SetFocus
'        ElseIf CmdEdit.Enabled = True Then
'            CmdSimpan.SetFocus
'        End If
'    End If
'    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
'End Sub









