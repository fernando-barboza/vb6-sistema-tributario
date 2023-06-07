VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmImprimirGuia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão "
   ClientHeight    =   2070
   ClientLeft      =   3015
   ClientTop       =   3390
   ClientWidth     =   6765
   Icon            =   "frmImprimirGuia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1935
      Left            =   60
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   3413
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Impressão de Guia"
      TabPicture(0)   =   "frmImprimirGuia.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDtmVencimento"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txt_dtmVencimento"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_strRequerente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd_Imprimir"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmd_Imprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   5430
         TabIndex        =   2
         Top             =   1020
         Width           =   1035
      End
      Begin VB.TextBox txt_strRequerente 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1050
         Width           =   3825
      End
      Begin VB.TextBox txt_dtmVencimento 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Requerente"
         Height          =   195
         Left            =   630
         TabIndex        =   5
         Top             =   1110
         Width           =   840
      End
      Begin VB.Label lblDtmVencimento 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vencimento"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   660
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmImprimirGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Imprimir_Click()
    If IsDate(txt_dtmVencimento.Text) = False Then
        ExibeMensagem "Essa data não é válida."
    ElseIf Trim(txt_strRequerente) = "" Then
        ExibeMensagem "O campo Requerente é obrigatório."
    Else
        frmAtualizacaoDebitos.strProprietario = Trim(txt_strRequerente)
        frmAtualizacaoDebitos.dtmVencimento = Trim(txt_dtmVencimento)
        frmAtualizacaoDebitos.ImprimirGuia
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    txt_dtmVencimento = gstrDataDoSistema
End Sub

Private Sub txt_dtmVencimento_GotFocus()
    MarcaCampo txt_dtmVencimento
End Sub

Private Sub txt_dtmVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmVencimento
End Sub

Private Sub txt_dtmVencimento_LostFocus()
    txt_dtmVencimento = gstrDataFormatada(txt_dtmVencimento)
End Sub

Private Sub txt_strRequerente_GotFocus()
    MarcaCampo txt_strRequerente
End Sub

Private Sub txt_strRequerente_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strRequerente
End Sub



