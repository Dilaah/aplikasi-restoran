VERSION 5.00
Begin VB.Form FrmLogin 
   Caption         =   "Daftar"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9285
   LinkTopic       =   "Form2"
   ScaleHeight     =   5925
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1320
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label password 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label telepon 
      AutoSize        =   -1  'True
      Caption         =   "telepon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label alamat 
      AutoSize        =   -1  'True
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label nama 
      AutoSize        =   -1  'True
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   630
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

