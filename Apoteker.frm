VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Apoteker 
   Caption         =   "Data Apoteker"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1440
      TabIndex        =   7
      Top             =   480
      Width           =   2500
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   1440
      TabIndex        =   8
      Top             =   840
      Width           =   4515
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   2500
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
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   400
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   400
      Left            =   3960
      TabIndex        =   4
      Top             =   1800
      Width           =   1000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   4920
      TabIndex        =   5
      Top             =   1800
      Width           =   1000
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   1850
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3254
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
      Left            =   240
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
   Begin VB.Label Label1 
      Caption         =   "Kode Apoteker"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Apoteker"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Telepon/HP"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   1245
   End
End
Attribute VB_Name = "Apoteker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBrawatjalan.mdb"
ADO.RecordSource = "select * from Apoteker"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
End Sub

Sub Form_Load()
    Call Koneksi
    Text1.MaxLength = 5
    Text2.MaxLength = 30
    Text3.MaxLength = 30
    Text4.MaxLength = 15
    KondisiAwal
End Sub

Function CariData()
    Call Koneksi
    RSApoteker.Open "Select * From Apoteker where KodeApt='" & Text1 & "'", Conn
End Function
Private Sub CmdBatal_Click()
KosongkanText
TidakSiapIsi
KondisiAwal
End Sub

Private Sub CmdSimpan_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
    MsgBox "Data Belum Lengkap...!"
    Exit Sub
Else
    If CmdInput.Enabled = True Then
        Dim SQLTambah As String
        SQLTambah = "Insert Into apoteker (KodeApt,NamaApt,AlamatApt,TeleponApt) values " & _
        "('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "')"
        Conn.Execute SQLTambah
    ElseIf CmdEdit.Enabled = True Then
        Dim SQLEdit As String
        SQLEdit = "Update apoteker Set NamaApt= '" & Text2 & "', AlamatApt = '" & Text3 & "',TeleponApt = '" & Text4 & "' where KodeApt='" & Text1 & "'"
        Conn.Execute SQLEdit
        
        Dim SQLEditpemakai As String
        SQLEditpemakai = "Update pemakai Set Namapmk= '" & Text2 & "' where KodeApt='" & Text1 & "'"
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
    Text4 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
End Sub

Private Sub TidakSiapIsi()
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
Text2 = RSApoteker!NamaApt
Text3 = RSApoteker!AlamatApt
Text4 = RSApoteker!TeleponApt
End Sub

Private Sub AutoNomor()
Call Koneksi
RSApoteker.Open ("select * from APOTEKER Where KodeApt In(Select Max(KodeApt)From APOTEKER )Order By KodeApt Desc"), Conn
RSApoteker.Requery
    Dim Urutan As String * 5
    Dim Hitung As Long
    With RSApoteker
        If .EOF Then
            Urutan = "APT" + "01"
            Text1 = Urutan
        Else
            Hitung = Right(!KodeApt, 2) + 1
            Urutan = "APT" + Right("00" & Hitung, 2)
        End If
        Text1 = Urutan
    End With
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
        Call AutoNomor
        Text1.Enabled = False
        Text2.SetFocus
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then

    If CmdInput.Enabled = True Then
        Call CariData
            If Not RSApoteker.EOF Then
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
            If Not RSApoteker.EOF Then
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
            If Not RSApoteker.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From ADM where KodeApt= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    
                    Dim SQLHapuspemakai As String
                    SQLHapuspemakai = "Delete From pemakai where KodeApt= '" & Text1 & "'"
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
        RSApoteker.Open "select * from apoteker where namaapt='" & Text2 & "'", Conn
        If Not RSApoteker.EOF Then
            MsgBox "nama '" & Text2 & "' sudah ada, coba tambahkan karakter pembeda"
            Text2.SetFocus
        Else
            Text3.SetFocus
        End If
    End If
End Sub

Private Sub text3_keypress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Text4.SetFocus
End Sub


Private Sub text4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdInput.Enabled = True Then
            CmdSimpan.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdSimpan.SetFocus
        End If
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub







