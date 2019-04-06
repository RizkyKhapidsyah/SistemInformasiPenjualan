VERSION 5.00
Begin VB.MDIForm FormUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Sistem Informasi Penjualan Makanan & Minuman - PT. Matador Country"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   Icon            =   "FormUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MenuMaster 
      Caption         =   "Master"
      Begin VB.Menu menuDataCustomer 
         Caption         =   "Data Customer"
      End
      Begin VB.Menu menuDataMinuman 
         Caption         =   "Data Minuman"
      End
   End
   Begin VB.Menu menuTransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu menuPenjualan 
         Caption         =   "Penjualan"
      End
   End
   Begin VB.Menu menuLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu menuLaporanCustomer 
         Caption         =   "Laporan Customer"
      End
      Begin VB.Menu menuLaporanDataMinuman 
         Caption         =   "Laporan Data Minuman"
      End
      Begin VB.Menu menuLaporanPenjualan 
         Caption         =   "Laporan Penjualan"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
      Begin VB.Menu menuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu menuKeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    WindowState = vbMaximized
End Sub

Private Sub menuAbout_Click()
MsgBox "Sistem Informasi Penjualan Makanan & Minuman - PT. Matador Country" & vbCrLf & _
        "Created by Bambang Aditya", vbInformation + vbOKOnly, "About"
End Sub

Private Sub menuDataCustomer_Click()
    With FormDataCustomer
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuDataMinuman_Click()
    With FormDataMinuman
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuKeluar_Click()
    End
End Sub

Private Sub menuLaporanCustomer_Click()
    With FormLaporanDataCustomer
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuLaporanDataMinuman_Click()
    With FormLaporanDataMinuman
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuLaporanPenjualan_Click()
    With FormLaporanPenjualan
        .Show
        .SetFocus
    End With
End Sub

Private Sub menuPenjualan_Click()
    With FormDataPenjualan
        .Show
        .SetFocus
    End With
End Sub
