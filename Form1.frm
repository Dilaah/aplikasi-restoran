VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   13155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   706
      TabCaption(0)   =   "Makanan"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(2)=   "gridMakan"
      Tab(0).Control(3)=   "pilihanMakanan"
      Tab(0).Control(4)=   "jumlah"
      Tab(0).Control(5)=   "simpan"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Minuman"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Dessert"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Struk"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "gridPesanan"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin MSDataGridLib.DataGrid gridPesanan 
         Height          =   3735
         Left            =   -74640
         TabIndex        =   13
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   6588
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myanmar Text"
            Size            =   9.75
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
      Begin VB.CommandButton simpan 
         Caption         =   "Pesan"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67200
         TabIndex        =   12
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox jumlah 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -67680
         TabIndex        =   11
         Text            =   "1"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox pilihanMakanan 
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -67680
         TabIndex        =   8
         Text            =   "Silahkan Dipilih"
         Top             =   960
         Width           =   2535
      End
      Begin MSDataGridLib.DataGrid gridMakan 
         Height          =   4575
         Left            =   -74760
         TabIndex        =   7
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   8070
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   33
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
            Name            =   "Myanmar Text"
            Size            =   9.75
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69240
         TabIndex        =   10
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pilih Pesanan"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69240
         TabIndex        =   9
         Top             =   960
         Width           =   1245
      End
   End
   Begin VB.TextBox text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   7095
   End
   Begin VB.TextBox text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4680
      TabIndex        =   15
      Top             =   240
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No Telepon"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPesanan As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim mode As String

Sub setupPilihan1()
strSQL = "SELECT * FROM makan ORDER BY id_makan ASC"
If rs.State = 1 Then rs.Close
rs.Open strSQL, con, adOpenStatic, adLockOptimistic
Do Until rs.EOF
    pilihanMakanan.AddItem rs.Fields("id_makan").Value & ". " & rs.Fields("nama").Value & ""
    rs.MoveNext
Loop
End Sub

Sub tampilMakan()
con.CursorLocation = adUseClient
strSQL = "select * from makan"
Set tabel = con.Execute(strSQL)
Set gridMakan.DataSource = tabel
End Sub

Sub displayData()
If rsPesanan.State = 1 Then rsPesanan.Close
strSQL = "SELECT * FROM Pesanan ORDER BY id ASC"
rsPesanan.Open strSQL, con, adOpenStatic, adLockOptimistic
Set gridPesanan.DataSource = rsPesanan.DataSource
End Sub

Private Sub Form_Activate()
    If con.State <> 1 Then bukakoneksi
    displayData
    setupPilihan1
    mode = "INSERT"
End Sub

Private Sub Form_Load()
bukakoneksi
tampilMakan
setupPilihan1
displayData
End Sub

Sub aturKontrol(ByVal logika As Boolean)
    id.Enabled = logika
    nama.Enabled = logika
    alamat.Enabled = logika
    telepon.Enabled = logika
    pilihanMakanan.Enabled = logika
    jumlah.Enabled = logika
    id.Text = ""
    nama.Text = ""
    alamat.Text = ""
    telepon.Text = ""
    pilihanMakanan.Text = ""
    jumlah.Text = ""
End Sub

Private Sub simpan_Click()
    If id.Text = "" Then
        MsgBox "No Masih Kosong", vbOKOnly + vbInformation, "Konfirmasi"
        id.SetFocus
        Exit Sub
    End If
    If nama.Text = "" Then
        MsgBox "nama Masih Kosong", vbOKOnly + vbInformation, "Konfirmasi"
        nama.SetFocus
        Exit Sub
    End If
    If alamat.Text = "" Then
        MsgBox "alamat Masih Kosong", vbOKOnly + vbInformation, "Konfirmasi"
        alamat.SetFocus
        Exit Sub
    End If
    If telepon.Text = "" Then
        MsgBox "telepon Masih Kosong", vbOKOnly + vbInformation, "Konfirmasi"
        telepon.SetFocus
        Exit Sub
    End If
If mode = "INSERT" Then
        strSQL = "INSERT INTO pesanan(id,nama, alamat, telepon, pilihanMakanan,pilihanMinuman,pilihanDessert,jumlah,total) " _
         & "VALUES('" & id.Text _
        & "','" & nama.Text _
        & "','" & alamat.Text _
        & "','" & telepon.Text _
        & "','" & pilihanMakanan.Text _
        & "','" & jumlah.Text & "')"
    ElseIf mode = "UPDATE" Then
        strSQL = "UPDATE pesanan " _
        & "SET no='" & id.Text _
        & "', nama='" & nama.Text _
        & "', alamat='" & alamat.Text _
        & "', telepon='" & telepon.Text _
        & "', pilihanMakanan='" & pilihanMakanan.Text _
        & "', jumlah='" & jumlah.Text _
        & "' WHERE nama='" & nama.Text & "')"
    End If
    con.Execute srtSQL
    tampilMakan
    displayData
    aturKontrol False
    mode = "INSERT"
    Exit Sub
End Sub

