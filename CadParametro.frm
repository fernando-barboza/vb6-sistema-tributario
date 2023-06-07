VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadParametro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   HelpContextID   =   106
   Icon            =   "CadParametro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5295
   Begin VB.CommandButton cmdAjudaParametro 
      Caption         =   "Ajuda"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3990
      TabIndex        =   3
      Top             =   2475
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancelaParametro 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2475
      Width           =   1110
   End
   Begin VB.CommandButton cmdOKParametro 
      Caption         =   "OK"
      Height          =   375
      Left            =   1515
      TabIndex        =   1
      Top             =   2475
      Width           =   1110
   End
   Begin TabDlg.SSTab sstParametro 
      Height          =   2235
      Left            =   165
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   165
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3942
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Geral"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_CorZebrado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFundoObjInacessivel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkListarCadastroAoIniciar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkRelatorioZebrado"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkConfirmaExclusao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkConfirmaGravacao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkFundo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkObjComGrade"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CheckBox chkObjComGrade 
         Caption         =   "Lista com grade"
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   1770
         UseMaskColor    =   -1  'True
         Width           =   2115
      End
      Begin VB.CheckBox chkFundo 
         Caption         =   "Fundo do obj inacessível diferente"
         Height          =   195
         Left            =   165
         TabIndex        =   9
         Top             =   1530
         UseMaskColor    =   -1  'True
         Width           =   2745
      End
      Begin VB.CheckBox chkConfirmaGravacao 
         Caption         =   "Confirma gravação"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         ToolTipText     =   "Exibir uma mensagem pedindo confirmação antes de gravar"
         Top             =   525
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox chkConfirmaExclusao 
         Caption         =   "Confirma exclusão"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         ToolTipText     =   "Exibir mensagem pedindo confirmação antes de excluir"
         Top             =   776
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox chkRelatorioZebrado 
         Caption         =   "Relatório zebrado"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   1027
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
      Begin VB.CheckBox chkListarCadastroAoIniciar 
         Caption         =   "Lista automática dos cadastros"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         ToolTipText     =   "Listar ao entrar na tela de cadastro ou alterar o banco de dados"
         Top             =   1278
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.Label lblFundoObjInacessivel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2970
         TabIndex        =   10
         ToolTipText     =   "Clic aqui para trocar a cor do zebrado"
         Top             =   1560
         Width           =   285
      End
      Begin VB.Label lbl_CorZebrado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   1755
         TabIndex        =   4
         ToolTipText     =   "Clic aqui para trocar a cor do zebrado"
         Top             =   1050
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmCadParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkRelatorioZebrado_Click()
    If chkRelatorioZebrado Then
        lbl_CorZebrado.Enabled = True
    Else
        lbl_CorZebrado.Enabled = False
    End If
End Sub

Private Sub cmdCancelaParametro_Click()
    Unload Me
End Sub

Private Sub cmdOKParametro_Click()
    gblnConfirmaGravacao = chkConfirmaGravacao
    gblnConfirmaExclusao = chkConfirmaExclusao
    gblnListagemAutomatica = chkListarCadastroAoIniciar
    gblnRelatorioZebrado = chkRelatorioZebrado
    gblnListViewComGrade = chkObjComGrade
    GravaUsuario
    cmdCancelaParametro_Click
End Sub

Private Sub Form_Load()
    chkConfirmaGravacao = Abs(gblnConfirmaGravacao)
    chkConfirmaExclusao = Abs(gblnConfirmaExclusao)
    chkListarCadastroAoIniciar = Abs(gblnListagemAutomatica)
    chkRelatorioZebrado = Abs(gblnRelatorioZebrado)
    chkObjComGrade = Abs(gblnListViewComGrade)
    lbl_CorZebrado.BackColor = Val(gvntCorZebrado)
    lblFundoObjInacessivel.BackColor = Val(gvntFundoObjInacessivel)
End Sub

Private Sub lbl_CorZebrado_Click()
    VerificaObjetoCor gvntCorZebrado, lbl_CorZebrado
End Sub

Sub VerificaObjetoCor(vntCor As Variant, lblObjeto As Object)
    MostraCaixaCores , , vntCor
    lblObjeto.BackColor = vntCor
End Sub

Private Sub lblFundoObjInacessivel_Click()
    VerificaObjetoCor gvntFundoObjInacessivel, lblFundoObjInacessivel
End Sub

