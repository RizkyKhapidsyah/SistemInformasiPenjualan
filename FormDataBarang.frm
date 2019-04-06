VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormDataMinuman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Minuman"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDataBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   10470
   Begin MSAdodcLib.Adodc AdodcJenisMinuman 
      Height          =   330
      Left            =   9120
      Top             =   3240
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
   Begin VB.ComboBox cmbJenisMinuman 
      Height          =   390
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1080
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   330
      Left            =   9240
      Top             =   3120
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
   Begin VB.CommandButton cmBaru 
      Caption         =   "&Baru"
      Height          =   495
      Left            =   2520
      TabIndex        =   20
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   3600
      TabIndex        =   19
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   4680
      TabIndex        =   18
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   6840
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
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
   Begin VB.TextBox textJumlah 
      Height          =   390
      Left            =   1560
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox TextHarga 
      Height          =   390
      Left            =   1560
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ComboBox cmbUkuran 
      Height          =   390
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox textNamaMinuman 
      Height          =   390
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox textKodeMinuman 
      Height          =   390
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Minuman"
      Height          =   270
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   22
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   13
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   10
      Top             =   2040
      Width           =   45
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   7
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ukuran"
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Minuman"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Minuman"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "FormDataMinuman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub AturKontrol()
    Nyambung
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From TbMinuman order by Kode_Minuman asc;"
        Set DataGrid1.DataSource = AdodcMain
        .Refresh
    End With
    With AdodcJenisMinuman
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbJenis order by Jenis_Minuman desc;"
        .Refresh
    End With
    cmbJenisMinuman.Clear
    Do Until AdodcJenisMinuman.Recordset.EOF
        cmbJenisMinuman.AddItem AdodcJenisMinuman.Recordset.Fields(0).Value, 0
        AdodcJenisMinuman.Recordset.MoveNext
    Loop
    AdodcJenisMinuman.Refresh
    cmbJenisMinuman.ListIndex = 0
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
        .textKodeMinuman.MaxLength = 15
        .textNamaMinuman.MaxLength = 50
        .textHarga.MaxLength = 50
        .textJumlah.MaxLength = 50
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
    Me.textKodeMinuman.SetFocus
End Sub
Sub AturDatagrid()
    With DataGrid1
        .Columns(0).Width = 989.8583
        .Columns(1).Width = 1709.858
        .Columns(2).Width = 1035.213
        .Columns(3).Width = 615.1182
        .Columns(4).Width = 689.9528
        .Columns(5).Width = 824.882
        .AllowUpdate = False
        .AllowDelete = False
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
    If textKodeMinuman.Text = "" Then
        MsgBox "Silahkan isi kode minuman!", vbExclamation + vbOKOnly, ""
        textKodeMinuman.SetFocus
    ElseIf textNamaMinuman.Text = "" Then
        MsgBox "Silahkan isi nama minuman!", vbExclamation + vbOKOnly, ""
        textNamaMinuman.SetFocus
    ElseIf textHarga.Text = "" Then
        MsgBox "Silahkan isi harga !", vbExclamation + vbOKOnly, ""
        textHarga.SetFocus
    ElseIf textJumlah.Text = "" Then
        MsgBox "Silahkan isi jumlah yang akan diinput!", vbExclamation + vbOKOnly, ""
        textJumlah.SetFocus
    Else
        X = MsgBox("Apakah Anda yakin ingin menyimpan data ini ke database?", vbQuestion + vbYesNo, "Konfirmasi?")
        If X = vbYes Then
            With AdodcMain
                .Recordset.AddNew
                .Recordset.Fields(0).Value = textKodeMinuman.Text
                .Recordset.Fields(1).Value = textNamaMinuman.Text
                .Recordset.Fields(2).Value = cmbJenisMinuman.Text
                .Recordset.Fields(3).Value = cmbUkuran.Text
                .Recordset.Fields(4).Value = textHarga.Text
                .Recordset.Fields(5).Value = textJumlah.Text
                .Recordset.Update
                .Refresh
            End With
            Reset
            textKodeMinuman.SetFocus
            AturDatagrid
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
