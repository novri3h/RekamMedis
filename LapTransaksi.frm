VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LapTransaksi 
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
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
   ScaleHeight     =   4575
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   1560
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   3375
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   1680
         TabIndex        =   2
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Akhir"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Awal"
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
      TabIndex        =   8
      Top             =   600
      Width           =   3375
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3375
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1500
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Laporan"
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
      TabIndex        =   13
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "LapTransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call Koneksi
RSResep.Open "Select Distinct TanggalRsp From Resep order By 1", Conn
RSResep.Requery
Do Until RSResep.EOF
    Combo1.AddItem Format(RSResep!TanggalRsp, "DD-MMM-YYYY")
    Combo2.AddItem Format(RSResep!TanggalRsp, "YYYY ,MM, DD")
    Combo3.AddItem Format(RSResep!TanggalRsp, "YYYY ,MM, DD")
    RSResep.MoveNext
Loop
Conn.Close

Call Koneksi
Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct month(TanggalRsp) as Bulan from Resep", Conn
Do While Not RSTGL.EOF
    Combo4.AddItem RSTGL!Bulan & Space(5) & MonthName(RSTGL!Bulan)
    RSTGL.MoveNext
Loop
Conn.Close

Call Koneksi
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(TanggalRsp)  as Tahun from Resep", Conn
Do While Not RSTHN.EOF
    Combo5.AddItem RSTHN!Tahun
    RSTHN.MoveNext
Loop
Conn.Close

End Sub

'Private Sub Form_Load()
'Call Koneksi
'RSResep.Open "Select Distinct TanggalRsp From Resep order By 1", Conn
'RSResep.Requery
'Do Until RSResep.EOF
'    Combo1.AddItem RSResep!TanggalRsp
'    Combo2.AddItem Format(RSResep!TanggalRsp, "YYYY ,MM, DD")
'    Combo3.AddItem Format(RSResep!TanggalRsp, "YYYY ,MM, DD")
'    RSResep.MoveNext
'Loop
'
'For i = 1 To 12
'    Combo4.AddItem i
'Next i
'For i = 1 To 10
'    Combo5.AddItem 2000 + i
'Next i
'
'End Sub

Private Sub combo1_KeyPress(Keyascii As Integer)
If Combo1 = "" Or Keyascii = 27 Then Unload Me
End Sub

'Lap Harian
Private Sub Combo1_Click()
    CR.SelectionFormula = "Totext({Resep.TanggalRsp})='" & CDate(Combo1) & "'"
    CR.ReportFileName = App.Path & "\Lap Kirim Harian.rpt"
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
        MsgBox "TanggalRsp awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    End If
    CR.SelectionFormula = "{Resep.TanggalRsp} in date (" & Combo2.Text & ") to date (" & Combo3.Text & ")"
    CR.ReportFileName = App.Path & "\Lap Kirim Mingguan.rpt"
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
    RSResep.Open "select * from Resep where month(TanggalRsp)='" & Val(Combo4) & "' and year(TanggalRsp)='" & (Combo5) & "'", Conn
    If RSResep.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If

    CR.SelectionFormula = "Month({Resep.TanggalRsp})=" & Val(Combo4.Text) & " and Year({Resep.TanggalRsp})=" & Val(Combo5.Text)
    CR.ReportFileName = App.Path & "\Lap Kirim Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub




