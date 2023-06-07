VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParametroUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções"
   ClientHeight    =   2835
   ClientLeft      =   5205
   ClientTop       =   4110
   ClientWidth     =   5295
   HelpContextID   =   4000
   Icon            =   "ParametroUsuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5295
   Begin TabDlg.SSTab sstParametro 
      Height          =   2490
      Left            =   165
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   165
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4392
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "ParametroUsuario.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_CorZebrado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFundoObjInacessivel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkListarCadastroAoIniciar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkRelatorioZebrado"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkConfirmaExclusao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkConfirmaGravacao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkFundoObjDiferente"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkObjComGrade"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkMostraDicaInicio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbointExercicioInicial"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.ComboBox cbointExercicioInicial 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2070
         Width           =   2235
      End
      Begin VB.CheckBox chkMostraDicaInicio 
         Caption         =   "&Mostra dicas ao iniciar"
         Height          =   195
         Left            =   165
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2550
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkObjComGrade 
         Caption         =   "Lista com grade"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   1815
         UseMaskColor    =   -1  'True
         Width           =   2115
      End
      Begin VB.CheckBox chkFundoObjDiferente 
         Caption         =   "Fundo do objeto inacessível diferente"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   1533
         UseMaskColor    =   -1  'True
         Width           =   2985
      End
      Begin VB.CheckBox chkConfirmaGravacao 
         Caption         =   "Confirma gravação"
         Height          =   195
         Left            =   165
         TabIndex        =   0
         ToolTipText     =   "Exibir uma mensagem pedindo confirmação antes de gravar"
         Top             =   405
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox chkConfirmaExclusao 
         Caption         =   "Confirma exclusão"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         ToolTipText     =   "Exibir mensagem pedindo confirmação antes de excluir"
         Top             =   687
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox chkRelatorioZebrado 
         Caption         =   "Relatório com zebrado"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   969
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CheckBox chkListarCadastroAoIniciar 
         Caption         =   "Lista automática dos cadastros"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         ToolTipText     =   "Listar ao entrar na tela de cadastro ou alterar o banco de dados"
         Top             =   1251
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.Label Label4 
         Caption         =   "Exercício"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   2130
         Width           =   705
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
         Left            =   3210
         TabIndex        =   6
         ToolTipText     =   "Clic aqui para trocar a cor do zebrado"
         Top             =   1563
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
         Left            =   2115
         TabIndex        =   3
         ToolTipText     =   "Clic aqui para trocar a cor do zebrado"
         Top             =   999
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmParametroUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim objform As Object
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
    
        If Not blnDadosOk Then
            Exit Sub
        End If
    
        gblnConfirmaGravacao = chkConfirmaGravacao
        gblnConfirmaExclusao = chkConfirmaExclusao
        gblnListagemAutomatica = chkListarCadastroAoIniciar
        gblnRelatorioZebrado = chkRelatorioZebrado
        gblnListViewComGrade = chkObjComGrade
        gbytchkFundoObjDiferente = chkFundoObjDiferente
        gblnMostraDicas = chkMostraDicaInicio
        gIntExercicioUsuario = gstrItemData(cbointExercicioInicial)
        GravaUsuario
        
        If gstrItemData(cbointExercicioInicial) <> gintExercicio Then
            
            For Each objform In Forms
                If Not (TypeName(objform) = "MDIMenu" Or TypeName(objform) = "frmParametroUsuario") Then Unload objform
            Next
            
            gintExercicio = gstrItemData(cbointExercicioInicial)
            MDIMenu.staBarraStatus.Panels(2).Text = "Exercicio Corrente: " & gintExercicio & " (" & RetornaSituacaoExercicio(gintExercicio) & ")"
            
        End If
        
        Unload Me

    ElseIf UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
End Sub

Private Function blnDadosOk() As Boolean
    If cbointExercicioInicial.ListIndex = -1 Then
        ExibeMensagem "É necessário informar em qual exercicio vai trabalhar."
        If cbointExercicioInicial.Enabled Then cbointExercicioInicial.SetFocus
        Exit Function
    End If
    
    If gstrItemData(cbointExercicioInicial) <> gintExercicio Then
        If MsgBox("Para trocar de Execício o sistema irá fechar todas as janelas abertas no momento. Deseja Continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            If cbointExercicioInicial.Enabled Then cbointExercicioInicial.SetFocus
            Exit Function
        End If
    End If

    
    blnDadosOk = True
End Function

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

Private Sub Form_Activate()
    gintCodSeguranca = 379
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrNovo, gstrImprimir, gstrDeletar, gstrLocalizar, gstrPreencherLista

End Sub

Private Sub Form_Load()

    LeDaTabelaParaObj "", cbointExercicioInicial, strQueryExercicio
    
    chkConfirmaGravacao = Abs(gblnConfirmaGravacao)
    chkConfirmaExclusao = Abs(gblnConfirmaExclusao)
    chkListarCadastroAoIniciar = Abs(gblnListagemAutomatica)
    chkRelatorioZebrado = Abs(gblnRelatorioZebrado)
    chkObjComGrade = Abs(gblnListViewComGrade)
    lbl_CorZebrado.BackColor = Val(gvntCorZebrado)
    lblFundoObjInacessivel.BackColor = Val(gvntFundoObjInacessivel)
    chkFundoObjDiferente = Abs(gbytchkFundoObjDiferente)
    chkMostraDicaInicio = Abs(gblnMostraDicas)
    cbointExercicioInicial.ListIndex = gintIndiceCBO(cbointExercicioInicial, gIntExercicioUsuario)
    Me.Icon = MDIMenu.Icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, _
                                   gstrNovo, _
                                   gstrImprimir, _
                                   gstrDeletar, _
                                   gstrLocalizar, _
                                   gstrPreencherLista
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


Private Function strQueryExercicio() As String

    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT intExercicio," & gstrCONVERT(CDT_NVARCHAR, "intExercicio") & " " & strCONCAT & " ' (' " & strCONCAT & " "
    strSql = strSql & gstrCASEWHEN("bytSituacao", _
        "0, 'Proposto', 1, 'Aberto', 2, 'Encerrado'") & strCONCAT & "')' "
    strSql = strSql & " AS strSituacao "
    strSql = strSql & "FROM " & gstrExercicio & " ORDER BY intExercicio"
    strQueryExercicio = strSql

End Function
