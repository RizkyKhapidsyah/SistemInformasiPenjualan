VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormDataCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Customer"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDataCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   10695
   Begin VB.CommandButton cmKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmBaru 
      Caption         =   "&Baru"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   4440
      TabIndex        =   12
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3201
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
   Begin VB.TextBox textTelp 
      Height          =   390
      Left            =   1800
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox textAlamat 
      Height          =   390
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox textNamaCustomer 
      Height          =   390
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox TextKodeCustomer 
      Height          =   390
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   330
      Left            =   9240
      Top             =   2280
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   10
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telp"
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   270
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Customer"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "FormDataCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub AturKontrol()
    Nyambung
    With AdodcMain
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * From tbCustomer order by Kode_Customer asc;"
        Set DataGrid1.DataSource = AdodcMain
        .Refresh
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
        .TextKodeCustomer.MaxLength = 10
        .textNamaCustomer.MaxLength = 30
        .textAlamat.MaxLength = 20
        .textTelp.MaxLength = 15
    End With
End Sub
Public Sub Batal()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
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
        End If
    Next
    Me.TextKodeCustomer.SetFocus
End Sub
Sub AturDatagrid()
    With Me
        .DataGrid1.Columns(0).Width = 1140.095
        .DataGrid1.Columns(1).Width = 1739.906
        .DataGrid1.Columns(2).Width = 1874.835
        .DataGrid1.Columns(3).Width = 975.1182
        .DataGrid1.AllowUpdate = False
        .DataGrid1.AllowDelete = False
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
    If TextKodeCustomer.Text = "" Then
        MsgBox "Silahkan isi kode customer!", vbExclamation + vbOKOnly, ""
        TextKodeCustomer.SetFocus
    ElseIf textNamaCustomer.Text = "" Then
        MsgBox "Silahkan isi nama customer!", vbExclamation + vbOKOnly, ""
        textNamaCustomer.SetFocus
    ElseIf textAlamat.Text = "" Then
        MsgBox "Silahkan isi alamat !", vbExclamation + vbOKOnly, ""
        textAlamat.SetFocus
    ElseIf textTelp.Text = "" Then
        MsgBox "Silahkan isi nomor telepon customer!", vbExclamation + vbOKOnly, ""
        textTelp.SetFocus
    Else
        X = MsgBox("Apakah Anda yakin ingin menyimpan data ini ke database?", vbQuestion + vbYesNo, "Konfirmasi?")
        If X = vbYes Then
            With AdodcMain
                .Recordset.AddNew
                .Recordset.Fields(0).Value = TextKodeCustomer.Text
                .Recordset.Fields(1).Value = textNamaCustomer.Text
                .Recordset.Fields(2).Value = textAlamat.Text
                .Recordset.Fields(3).Value = textTelp.Text
                .Recordset.Update
                .Refresh
            End With
            Reset
            TextKodeCustomer.SetFocus
        End If
        AturDatagrid
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

