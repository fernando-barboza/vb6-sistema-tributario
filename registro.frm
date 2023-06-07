VERSION 5.00
Begin VB.Form frmRegistro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4230
   ClientLeft      =   3930
   ClientTop       =   2955
   ClientWidth     =   6255
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6255
   Begin VB.CommandButton cmd_OK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   4770
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3660
      Width           =   1215
   End
   Begin VB.TextBox txtfoco 
      Height          =   225
      Left            =   5220
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3750
      Width           =   375
   End
   Begin VB.Label lbl_Email 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "cpdsystems@cpdsystems.com.br"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1890
      MouseIcon       =   "registro.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2940
      Width           =   2340
   End
   Begin VB.Label lbl_CPD 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.cpdsystems.com.br"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2160
      MouseIcon       =   "registro.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3240
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Para obter a versão completa entre em contato com:"
      Height          =   390
      Left            =   1620
      TabIndex        =   2
      Top             =   2220
      Width           =   2250
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   2220
      Left            =   4740
      Picture         =   "registro.frx":0614
      Stretch         =   -1  'True
      Top             =   1230
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Este sistema é de demonstração, portanto, possui limitações na quantidade de registros a serem gravados."
      Height          =   585
      Left            =   1620
      TabIndex        =   1
      Top             =   1290
      Width           =   3060
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   4185
      Left            =   30
      Picture         =   "registro.frx":53AC
      Top             =   30
      Width           =   6210
   End
End
Attribute VB_Name = "frmRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_OK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
End Sub

Private Sub lbl_CPD_Click()
    On Error Resume Next
    ShellEx "http://www.cpdsystems.com.br", , , , , Me.hwnd
End Sub

Private Sub lbl_Email_Click()
    Dim result
    On Error Resume Next
    result = ShellExecute(Me.hwnd, vbNullString, "mailto:selectron@selectron.com.br", vbNullString, "c:\", 1)
End Sub
