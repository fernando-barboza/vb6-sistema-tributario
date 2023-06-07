VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2460
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4035
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4035
   Begin VB.TextBox txtSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1545
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   525
      Width           =   2280
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1470
      TabIndex        =   8
      Top             =   1005
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   210
      TabIndex        =   7
      Top             =   1005
      Width           =   1110
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   1545
      MaxLength       =   15
      TabIndex        =   6
      Top             =   150
      Width           =   2280
   End
   Begin VB.CommandButton cmdAvancado 
      Caption         =   "Avançado >>"
      Height          =   375
      Left            =   2715
      TabIndex        =   5
      Top             =   1005
      Width           =   1110
   End
   Begin VB.Frame fra_ServidorDataBase 
      Height          =   1005
      Left            =   210
      TabIndex        =   0
      Top             =   1395
      Width           =   3615
      Begin MSDataListLib.DataCombo dbc_Databases 
         Height          =   315
         Left            =   1170
         TabIndex        =   4
         Top             =   600
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtServidor 
         Height          =   285
         Left            =   1170
         MaxLength       =   20
         TabIndex        =   2
         Top             =   225
         Width           =   2220
      End
      Begin VB.Label lbl_DataBase 
         AutoSize        =   -1  'True
         Caption         =   "Database:"
         Height          =   195
         Left            =   255
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbl_Servidor 
         AutoSize        =   -1  'True
         Caption         =   "Servidor:"
         Height          =   195
         Left            =   255
         TabIndex        =   1
         Top             =   255
         Width           =   630
      End
   End
   Begin VB.Label lbl_Senha 
      AutoSize        =   -1  'True
      Caption         =   "Senha:"
      Height          =   195
      Left            =   900
      TabIndex        =   11
      Top             =   570
      Width           =   510
   End
   Begin VB.Label lbl_Usuario 
      AutoSize        =   -1  'True
      Caption         =   "Usuário:"
      Height          =   195
      Left            =   825
      TabIndex        =   10
      Top             =   180
      Width           =   585
   End
   Begin VB.Image img_Chave 
      Height          =   480
      Left            =   135
      Picture         =   "Login.frx":1042
      Top             =   210
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
