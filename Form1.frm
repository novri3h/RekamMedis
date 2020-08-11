VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama"
   ClientHeight    =   5130
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1148
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pemakai"
            Key             =   "F1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pasien"
            Key             =   "F2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Dokter"
            Key             =   "F3"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4755
      Width           =   3960
      _ExtentX        =   6985
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
   End
   Begin VB.Menu mnfile 
      Caption         =   "Administrator"
      Begin VB.Menu mnpoli 
         Caption         =   "Poli"
      End
      Begin VB.Menu mndokter 
         Caption         =   "Dokter"
      End
      Begin VB.Menu mnpemakai 
         Caption         =   "Pemakai"
      End
      Begin VB.Menu mnlap 
         Caption         =   "Laporan"
         Begin VB.Menu mnlapmaster 
            Caption         =   "Data Master"
         End
         Begin VB.Menu mnlapbayar 
            Caption         =   "Data Pembayaran"
         End
      End
   End
   Begin VB.Menu mnadm 
      Caption         =   "ADM"
      Begin VB.Menu mnpendaftaran 
         Caption         =   "Pendaftaran"
      End
      Begin VB.Menu mnpasien 
         Caption         =   "Pasien"
      End
      Begin VB.Menu mninfo 
         Caption         =   "Informasi Pasien"
      End
   End
   Begin VB.Menu mnapoteker 
      Caption         =   "Apoteker"
      Begin VB.Menu mnobat 
         Caption         =   "Obat"
      End
      Begin VB.Menu mnresep 
         Caption         =   "Resep"
      End
   End
   Begin VB.Menu mnutil 
      Caption         =   "Utility"
   End
   Begin VB.Menu mnlogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "F1"
        Pemakai.Show
    Case "F2"
        Pasien.Show
    Case "F3"
        Dokter.Show
End Select
End Sub


