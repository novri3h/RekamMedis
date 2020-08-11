VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form InfoPasien 
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
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
   ScaleHeight     =   4365
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   3435
      Left            =   7800
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
   Begin MSDataGridLib.DataGrid dg 
      Bindings        =   "InfoPasien.frx":0000
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin MSAdodcLib.Adodc ado 
      Height          =   495
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
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
      Caption         =   "Informasi Pasien"
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
      TabIndex        =   4
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Pasien"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "InfoPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Call Koneksi
Dim RS As New ADODB.Recordset
RS.Open "select distinct namapsn from pasien,pendaftaran where pasien.kodepsn=pendaftaran.kodepsn", Conn
List1.Clear
Do While Not RS.EOF
    List1.AddItem RS!NamaPsn
    RS.MoveNext
Loop
End Sub

Private Sub List1_Click()
Call Koneksi
ado.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source =" & App.Path & "\dbrawatjalan.mdb"
ado.RecordSource = "select distinct namapsn as [Nama Pasien],tanggaldft as [Tanggal Daftar],namadkt as [Nama Dokter],namapl as [Nama Poli] from pasien,pendaftaran,dokter,poli where pasien.kodepsn=pendaftaran.kodepsn and dokter.kodedkt=pendaftaran.kodedkt and poli.kodepl=pendaftaran.kodepl and namapsn  like '%" & List1 & "%'"
ado.Refresh
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    Call Koneksi
    ado.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source =" & App.Path & "\dbrawatjalan.mdb"
    ado.RecordSource = "select distinct namapsn as [Nama Pasien],tanggaldft as [Tanggal Daftar],namadkt as [Nama Dokter],namapl as [Nama Poli] from pasien,pendaftaran,dokter,poli where pasien.kodepsn=pendaftaran.kodepsn and dokter.kodedkt=pendaftaran.kodedkt and poli.kodepl=pendaftaran.kodepl and namapsn  like '%" & Text1 & "%'"
    ado.Refresh
    If ado.Recordset.EOF Then
        MsgBox "nama pasien tidak ada"
        Text1.SetFocus
        Text1 = ""
    End If
End If
        
End Sub

