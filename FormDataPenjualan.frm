VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormDataPenjualan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Penjualan"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14070
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDataPenjualan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   14070
   Begin VB.TextBox textTanggalBeli 
      Height          =   390
      Left            =   1800
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox cmbJenisMinuman 
      Height          =   390
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   1080
      Width           =   2775
   End
   Begin VB.ComboBox cmbKodeMinuman 
      Height          =   390
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   120
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   330
      Left            =   10200
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   8280
      TabIndex        =   28
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   7200
      TabIndex        =   27
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   6120
      TabIndex        =   26
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   5040
      TabIndex        =   25
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmBaru 
      Caption         =   "&Baru"
      Height          =   495
      Left            =   3960
      TabIndex        =   24
      Top             =   4920
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4695
      Left            =   4680
      TabIndex        =   23
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
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
   Begin VB.TextBox textKembalian 
      Height          =   390
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox textJumlahBayar 
      Height          =   390
      Left            =   1800
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox textTotal 
      Height          =   390
      Left            =   1800
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox textJumlah 
      Height          =   390
      Left            =   1800
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox textHarga 
      Height          =   390
      Left            =   1800
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.ComboBox cmbUkuran 
      Height          =   390
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox textNamaMinuman 
      Height          =   390
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc AdodcJenisMinuman 
      Height          =   330
      Left            =   10560
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcKodeMinuman 
      Height          =   330
      Left            =   960
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Beli"
      Height          =   270
      Left            =   240
      TabIndex        =   35
      Top             =   2040
      Width           =   765
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   34
      Top             =   2040
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Minuman "
      Height          =   270
      Left            =   240
      TabIndex        =   32
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   31
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   21
      Top             =   4440
      Width           =   45
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kembalian"
      Height          =   270
      Left            =   240
      TabIndex        =   20
      Top             =   4440
      Width           =   645
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   18
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bayar"
      Height          =   270
      Left            =   240
      TabIndex        =   17
      Top             =   3960
      Width           =   915
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   15
      Top             =   3480
      Width           =   45
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   270
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   315
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   12
      Top             =   3000
      Width           =   45
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
      Height          =   270
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   465
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   9
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      Height          =   270
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   390
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ukuran"
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Minuman"
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Minuman"
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "FormDataPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AturKontrol()
    Nyambung
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TbPenjualan order by Kode_Minuman asc;"
        Set DataGrid1.DataSource = AdodcMain
        .Refresh
    End With
    With AdodcJenisMinuman
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbJenis order by Jenis_Minuman desc;"
        .Refresh
    End With
    With AdodcKodeMinuman
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbMinuman order by Kode_Minuman desc;"
        .Refresh
    End With
    cmbJenisMinuman.Clear
    Do Until AdodcJenisMinuman.Recordset.EOF
        cmbJenisMinuman.AddItem AdodcJenisMinuman.Recordset.Fields(0).Value, 0
        AdodcJenisMinuman.Recordset.MoveNext
    Loop
    AdodcJenisMinuman.Refresh
    cmbJenisMinuman.ListIndex = 0
    cmbKodeMinuman.Clear
    Do Until AdodcKodeMinuman.Recordset.EOF
        cmbKodeMinuman.AddItem AdodcKodeMinuman.Recordset.Fields(0).Value, 0
        AdodcKodeMinuman.Recordset.MoveNext
    Loop
    AdodcKodeMinuman.Refresh
    With cmbUkuran
        .Clear
        .AddItem "Tall", 0
        .AddItem "Grande", 1
        .AddItem "Bigben", 2
        .ListIndex = 0
    End With
    With Me
        .cmSimpan.Enabled = False
        .cmBatal.Enabled = False
    End With
    AturDatagrid
    textTotal.Locked = True
    textKembalian.Locked = True
End Sub
Public Sub Reset()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .BackColor = vbWhite
                .Enabled = True
            End With
        End If
    Next
    With Me
        .textNamaMinuman.MaxLength = 50
        .textHarga.MaxLength = 50
        .textJumlah.MaxLength = 50
        .textTotal.MaxLength = 50
        .textJumlah.MaxLength = 50
        .textKembalian.MaxLength = 50
    End With
End Sub
Public Sub Batal()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .BackColor = Me.BackColor
                .Enabled = False
            End With
        ElseIf TypeName(Objek) = "ComboBox" Then
            With Objek
                .BackColor = Me.BackColor
                .Enabled = False
            End With
        End If
    Next
End Sub
Public Sub Baru()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .BackColor = vbWhite
                .Enabled = True
            End With
        ElseIf TypeName(Objek) = "ComboBox" Then
            With Objek
                .BackColor = vbWhite
                .Enabled = True
            End With
        End If
    Next
    Me.cmbKodeMinuman.SetFocus
End Sub
Public Sub AturDatagrid()
    With Me
        .DataGrid1.AllowUpdate = False
        .DataGrid1.AllowDelete = False
        .DataGrid1.Columns(0).Width = 1065.26
        .DataGrid1.Columns(1).Width = 1319.811
        .DataGrid1.Columns(2).Width = 1124.787
        .DataGrid1.Columns(3).Width = 615.1182
        .DataGrid1.Columns(4).Width = 629.8583
        .DataGrid1.Columns(5).Width = 675.2126
        .DataGrid1.Columns(6).Width = 615.1182
        .DataGrid1.Columns(7).Width = 555.0236
        .DataGrid1.Columns(8).Width = 1100.047
        .DataGrid1.Columns(9).Width = 1200.047
    End With
End Sub

Private Sub cmBaru_Click()
    With Me
        .cmSimpan.Enabled = True
        .cmBatal.Enabled = True
    End With
    cmBaru.Enabled = False
    Call Baru
End Sub

Public Sub cmBatal_Click()
    With Me
        .cmSimpan.Enabled = False
        .cmBatal.Enabled = False
    End With
    cmBaru.Enabled = True
    Call Reset
    Call Batal
End Sub

Private Sub cmHapus_Click()
If AdodcMain.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang akan dihapus!", vbExclamation + vbOKOnly, ""
Else
    X = MsgBox("Anda yakin ingin menghapus data ini?", vbQuestion + vbYesNo, "Hapus?")
    If X = vbYes Then
        With AdodcMain
            .Recordset.Delete
            .Refresh
        End With
    End If
End If
End Sub

Private Sub cmKeluar_Click()
    Unload Me
End Sub

Private Sub cmSimpan_Click()
On Error GoTo HancurkanError
    If textNamaMinuman.Text = "" Then
        MsgBox "Silahkan isi nama minuman!", vbExclamation + vbOKOnly, ""
        textNamaMinuman.SetFocus
    ElseIf textHarga.Text = "" Then
        MsgBox "Silahkan isi harga !", vbExclamation + vbOKOnly, ""
        textHarga.SetFocus
    ElseIf textJumlah.Text = "" Then
        MsgBox "Silahkan isi jumlah yang akan diinput!", vbExclamation + vbOKOnly, ""
        textJumlah.SetFocus
    ElseIf textJumlahBayar.Text = "" Then
        MsgBox "Silahkan isi Jumlah Bayar!", vbExclamation + vbOKOnly, ""
        textJumlahBayar.SetFocus
    Else
        X = MsgBox("Apakah Anda yakin ingin menyimpan data ini ke database?", vbQuestion + vbYesNo, "Konfirmasi?")
        If X = vbYes Then
            With AdodcMain
                .Recordset.AddNew
                .Recordset.Fields(0).Value = cmbKodeMinuman.Text
                .Recordset.Fields(1).Value = textNamaMinuman.Text
                .Recordset.Fields(2).Value = cmbJenisMinuman.Text
                .Recordset.Fields(3).Value = cmbUkuran.Text
                .Recordset.Fields(4).Value = textTanggalBeli.Text
                .Recordset.Fields(5).Value = textHarga.Text
                .Recordset.Fields(6).Value = textJumlah.Text
                .Recordset.Fields(7).Value = textTotal.Text
                .Recordset.Fields(8).Value = textJumlahBayar.Text
                .Recordset.Fields(9).Value = textKembalian.Text
                .Recordset.Update
                .Refresh
            End With
            cmbKodeMinuman.SetFocus
            AturDatagrid
            Reset
        End If
    End If
Exit Sub
HancurkanError:
    PusatError
End Sub

Private Sub Form_Load()
    AturKontrol
    Reset
    cmBatal_Click
End Sub



Private Sub TextHarga_Change()
    Me.textTotal.Text = Val(textHarga.Text) * Val(textJumlah.Text)
End Sub

Private Sub textJumlah_Change()
    Me.textTotal.Text = Val(textHarga.Text) * Val(textJumlah.Text)
End Sub

Private Sub textJumlahBayar_Change()
    Me.textKembalian.Text = Val(Me.textJumlahBayar.Text) - Val(Me.textTotal.Text)
End Sub
