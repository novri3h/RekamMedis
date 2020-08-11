VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LapPembayaran 
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3735
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
   ScaleHeight     =   4590
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   1560
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   3375
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Harian"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   3375
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3375
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Awal"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Akhir"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1500
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Laporan Pembayaran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "LapPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call Koneksi
RSPembayaran.Open "Select Distinct TanggalByr From Pembayaran order By 1", Conn
RSPembayaran.Requery
Do Until RSPembayaran.EOF
    Combo1.AddItem Format(RSPembayaran!TanggalByr, "DD-MMM-YYYY")
    Combo2.AddItem Format(RSPembayaran!TanggalByr, "YYYY ,MM, DD")
    Combo3.AddItem Format(RSPembayaran!TanggalByr, "YYYY ,MM, DD")
    RSPembayaran.MoveNext
Loop
Conn.Close

Call Koneksi
Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct month(TanggalByr) as Bulan from Pembayaran", Conn
Do While Not RSTGL.EOF
    Combo4.AddItem RSTGL!Bulan & Space(5) & MonthName(RSTGL!Bulan)
    RSTGL.MoveNext
Loop
Conn.Close

Call Koneksi
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(TanggalByr)  as Tahun from Pembayaran", Conn
Do While Not RSTHN.EOF
    Combo5.AddItem RSTHN!Tahun
    RSTHN.MoveNext
Loop
Conn.Close

End Sub


Private Sub combo1_KeyPress(Keyascii As Integer)
If Combo1 = "" Or Keyascii = 27 Then Unload Me
End Sub

'Lap Harian
Private Sub Combo1_Click()
    CR.SelectionFormula = "Totext({Pembayaran.TanggalByr})='" & CDate(Combo1) & "'"
    CR.ReportFileName = App.Path & "\Bayar Harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub combo2_keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Mingguan (Tgl Antara)
Private Sub Combo3_Click()
    If Combo2 = "" Then
        MsgBox "TanggalByr awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    End If
    CR.SelectionFormula = "{Pembayaran.TanggalByr} in date (" & Combo2.Text & ") to date (" & Combo3.Text & ")"
    CR.ReportFileName = App.Path & "\Bayar Mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo4_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Bulanan
Private Sub Combo5_Click()
    Call Koneksi
    RSPembayaran.Open "select * from Pembayaran where month(TanggalByr)='" & Val(Combo4) & "' and year(TanggalByr)='" & (Combo5) & "'", Conn
    If RSPembayaran.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If

    CR.SelectionFormula = "Month({Pembayaran.TanggalByr})=" & Val(Combo4.Text) & " and Year({Pembayaran.TanggalByr})=" & Val(Combo5.Text)
    CR.ReportFileName = App.Path & "\Bayar Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

