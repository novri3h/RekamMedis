VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama Program Rekam Medis"
   ClientHeight    =   3675
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   3675
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   120
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   720
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":1FBC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20015
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20467
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":208B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20D0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2115D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":215AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":21A01
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":21E53
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":222A5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar STBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3300
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "&Administrator"
      Begin VB.Menu mnpemakai 
         Caption         =   "Pe&makai"
      End
      Begin VB.Menu mnpoli 
         Caption         =   "P&oli"
      End
      Begin VB.Menu mndokter 
         Caption         =   "&Dokter"
      End
      Begin VB.Menu mnlaporan 
         Caption         =   "&Laporan"
         Begin VB.Menu mnlapmaster 
            Caption         =   "Data Master"
         End
         Begin VB.Menu mnlaptrans 
            Caption         =   "Data Transaksi"
         End
         Begin VB.Menu mnsalin 
            Caption         =   "&Backup Database"
         End
      End
   End
   Begin VB.Menu mnadm 
      Caption         =   "&Adm"
      Begin VB.Menu mnpendaftaran 
         Caption         =   "&Pendaftaran"
      End
      Begin VB.Menu mnpasien 
         Caption         =   "&Pasien"
      End
      Begin VB.Menu mninfopasien 
         Caption         =   "&Informasi Pasien"
      End
   End
   Begin VB.Menu mnapoteker 
      Caption         =   "Apo&teker"
      Begin VB.Menu mnobat 
         Caption         =   "&Obat"
      End
      Begin VB.Menu mnresep 
         Caption         =   "&Resep"
      End
   End
   Begin VB.Menu mnutility 
      Caption         =   "&Utility"
      Begin VB.Menu mnganpass 
         Caption         =   "Ganti Password"
      End
      Begin VB.Menu mnbackup 
         Caption         =   "Backup Database"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
'If Keyascii = 13 Then UjiSQL.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnbackup_Click()
SalinFile.Show
End Sub

Private Sub mndokter_Click()
Dokter.Show
End Sub

Private Sub mndtadm_Click()
Adm.Show
End Sub

Private Sub mndtapoteker_Click()
Apoteker.Show
End Sub

Private Sub mnganpass_Click()
GantiPass.Show
End Sub

Private Sub mninfopasien_Click()
InfoPasien.Show
End Sub

Private Sub mnkeluar_Click()
'Pesan = MsgBox("Yakin akan akan keluar..?", vbYesNo)
'If Pesan = vbYes Then End
Menu.STBar.Panels(1).Text = ""
Menu.STBar.Panels(2).Text = ""
Menu.STBar.Panels(3).Text = ""
Menu.mnadm.Enabled = True
Menu.mnfile.Enabled = True
Menu.mnapoteker.Enabled = True
Menu.Visible = False
Login.Show
Login.Text1 = ""
Login.Text2 = ""
Login.Text3 = ""
Login.Text4 = ""
Login.Text1.Enabled = True
Login.Text1.SetFocus
End Sub

Private Sub mnlapmaster_Click()
LapMaster.Show
End Sub

Private Sub mnlaptrans_Click()
LapPembayaran.Show
End Sub

Private Sub mnobat_Click()
Obat.Show
End Sub

Private Sub mnpasien_Click()
Pasien.Show
End Sub

Private Sub mnpemakai_Click()
Pemakai.Show
End Sub

Private Sub mnpendaftaran_Click()
Pendaftaran.Show
End Sub

Private Sub mnpoli_Click()
Poli.Show
End Sub

Private Sub mnresep_Click()
Resep.Show
End Sub

Private Sub mnsalin_Click()
SalinFile.Show
End Sub

Private Sub TLBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "F1"
        Pendaftaran.Show
    Case "F2"
        Pasien.Show
    Case "F3"
        InfoPasien.Show
    Case "F4"
        Obat.Show
    Case "F5"
        Resep.Show
    Case "F6"
       GantiPass.Show
    Case "F7"
        Menu.STBar.Panels(1).Text = ""
        Menu.STBar.Panels(2).Text = ""
        Menu.STBar.Panels(3).Text = ""
        Menu.mnadm.Enabled = True
        Menu.mnfile.Enabled = True
        Menu.mnapoteker.Enabled = True
        Menu.Visible = False
        Login.Show
        Login.Text1 = ""
        Login.Text2 = ""
        Login.Text3 = ""
        Login.Text4 = ""
        Login.Text1.Enabled = True
        Login.Text1.SetFocus

    Case "F8"
        End
End Select
End Sub
