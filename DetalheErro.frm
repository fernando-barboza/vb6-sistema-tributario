VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetalheErro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "DetalheErro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_Query 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
   End
   Begin MSComctlLib.ImageList img_Imagens 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DetalheErro.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DetalheErro.frx":2BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DetalheErro.frx":304A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DetalheErro.frx":349E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt_Mensagem 
      BackColor       =   &H00C0C0C0&
      Height          =   2115
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1650
      Width           =   5895
   End
   Begin VB.CommandButton cmd_Detalhes 
      Caption         =   "&Detalhes>>"
      Height          =   345
      Left            =   4890
      TabIndex        =   3
      Top             =   1110
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Fechar 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4890
      TabIndex        =   2
      Top             =   180
      Width           =   1155
   End
   Begin VB.Image img_Erro 
      Height          =   480
      Left            =   180
      Picture         =   "DetalheErro.frx":38F2
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbl_Mensagem1 
      Caption         =   "Ocorreu um erro interno neste programa. Para mais informações, clique em Detalhes."
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   210
      Width           =   3690
   End
   Begin VB.Label lbl_Mensagem2 
      Caption         =   "Se o problema persistir, entre em contato com o suporte técnico."
      Height          =   435
      Left            =   960
      TabIndex        =   0
      Top             =   1020
      Width           =   3675
   End
End
Attribute VB_Name = "frmDetalheErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnQuery As Boolean

Private Sub cmd_Detalhes_Click()
    cmd_Detalhes.Enabled = False
    Me.Height = 4275
End Sub

Private Sub cmd_Fechar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 5 And Not mblnQuery Then
        txt_Mensagem = txt_Mensagem & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                       "Instrução SQL: " & Chr(13) & Chr(10) & Trim(txt_Query)
        mblnQuery = True
    End If
End Sub

Private Sub Form_Load()
    Dim strMsg As String
    
    Call AlwaysOnTop(Me, False)
    
    Screen.MousePointer = 0
    Me.Height = 1965
    Me.Caption = App.FileDescription
    mblnQuery = False
    
    If Err.Number <> 0 Then
        strMsg = "Número do erro: " & Err.Number & Chr(13) & Chr(10) & _
                       "Descrição: " & Err.Description
    End If
    txt_Mensagem = strMsg
End Sub
