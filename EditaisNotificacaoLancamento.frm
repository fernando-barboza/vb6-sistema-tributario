VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEditaisNotificacaoLancamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editais / Notificações de Lançamento"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   DrawMode        =   1  'Blackness
   HelpContextID   =   693
   Icon            =   "EditaisNotificacaoLancamento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7395
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1995
      Left            =   150
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   150
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3519
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Editais / Notificações de Lançamento"
      TabPicture(0)   =   "EditaisNotificacaoLancamento.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcintInscricaoCadastral"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintContribuincaoMelhoria"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintEdital"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk_TodasInscricoes"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox chk_TodasInscricoes 
         Caption         =   "Selecionar todas as inscrições do Edital"
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   1200
         Width           =   3495
      End
      Begin MSDataListLib.DataCombo dbcintEdital 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   420
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintContribuincaoMelhoria 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   1530
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintInscricaoCadastral 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   840
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Label lbl_Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Edital"
         Height          =   195
         Left            =   735
         TabIndex        =   7
         Top             =   495
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1590
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   465
         TabIndex        =   5
         Top             =   960
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmEditaisNotificacaoLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk_TodasInscricoes_Click()
    If chk_TodasInscricoes.Value = 1 Then
        dbcintInscricaoCadastral.BoundText = ""
        dbcintInscricaoCadastral.Enabled = False
        TrocaCorObjeto dbcintInscricaoCadastral, True
    Else
        dbcintInscricaoCadastral.Enabled = True
        TrocaCorObjeto dbcintInscricaoCadastral, False
    End If
End Sub

Private Sub chk_TodasInscricoes_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintContribuincaoMelhoria_Click(Area As Integer)
    DropDownDataCombo dbcintContribuincaoMelhoria, Me, Area
End Sub

Private Sub dbcintContribuincaoMelhoria_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintContribuincaoMelhoria, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuincaoMelhoria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintEdital_Click(Area As Integer)
    DropDownDataCombo dbcintEdital, Me, Area
    If Area = 2 Then
        If dbcintEdital.MatchedWithList = True Then
            LeDaTabelaParaObj "", dbcintInscricaoCadastral, strQuerryInscricao
        End If
    End If
End Sub

Private Sub dbcintEdital_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintEdital, Me, , KeyCode, Shift
End Sub

Private Sub dbcintInscricaoCadastral_Click(Area As Integer)
    DropDownDataCombo dbcintInscricaoCadastral, Me, Area
End Sub

Private Sub dbcintInscricaoCadastral_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintInscricaoCadastral, Me, , KeyCode, Shift
End Sub

Private Sub dbcintInscricaoCadastral_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 693
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrAplicar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    LeDaTabelaParaObj gstrTabelaDeEdital, dbcintEdital, QuerryEdital
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbcintContribuincaoMelhoria, strQuerrry
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
End Sub

Private Function strQuerrry() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " PKId, strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita
    strSql = strSql & " WHERE "
    strSql = strSql & " intUtilizacao = 1 "
    strSql = strSql & " ORDER BY "
    strSql = strSql & " strDescricao "
    strQuerrry = strSql
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    On Error Resume Next
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            ImprimeRelatorio rptEditaisNotificacoes, strQuerryEditalNotificacao
        End If
    End If
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaObjetos
    End If
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
End Sub

Private Sub LimpaObjetos()
    dbcintEdital.BoundText = ""
    chk_TodasInscricoes.Value = 0
    dbcintInscricaoCadastral.BoundText = ""
    dbcintContribuincaoMelhoria.BoundText = ""
    dbcintEdital.SetFocus
End Sub

Private Function blnDadosOk() As Boolean
    If Not dbcintEdital.MatchedWithList Then
       ExibeMensagem "O Edital tem que ser selecionado."
       dbcintEdital.SetFocus
       Exit Function
    End If
    
    If chk_TodasInscricoes.Value = 0 Then
        If Not dbcintInscricaoCadastral.MatchedWithList Then
           ExibeMensagem "A Inscricao Cadastral tem que ser selecionada."
           dbcintInscricaoCadastral.SetFocus
           Exit Function
        End If
    End If
    
    If Not dbcintContribuincaoMelhoria.MatchedWithList Then
       ExibeMensagem "A Composição da Receita tem que ser selecionada."
       dbcintContribuincaoMelhoria.SetFocus
       Exit Function
    End If
    
    blnDadosOk = True
    
End Function
    
Private Sub dbcintEdital_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Function QuerryEdital() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT"
    strSql = strSql & " PKId, strNomeDoEdital "
    strSql = strSql & " FROM "
    strSql = strSql & gstrTabelaDeEdital
    strSql = strSql & " ORDER BY strNomeDoEdital "
    QuerryEdital = strSql
End Function

Private Function strQuerryInscricao() As String
    Dim strSql As String
    strSql = strSql & " SELECT DISTINCT"
    strSql = strSql & " A.PKId, A.strInscricaoAnterior "

    strSql = strSql & " FROM "
    strSql = strSql & gstrImobiliario & " A, "
    strSql = strSql & gstrContribuicaoMelhoria & " B "

    strSql = strSql & " WHERE "
    strSql = strSql & " B.intImobiliario = A.PKId "
    strSql = strSql & " AND B.intTabelaDeEdital = " & Val(dbcintEdital.BoundText)
    
    strSql = strSql & " ORDER BY "
    strSql = strSql & " A.strInscricaoAnterior "
    strQuerryInscricao = strSql
End Function

Private Function strQuerryEditalNotificacao() As String
    Dim strSql As String
    
    strSql = "SELECT DISTINCT A.PKId, D.intContribuinte, " & _
    "A.dblCustoDaParcela, C.strNome, D.strInscricaoAnterior " & _
    " FROM " & _
    gstrTabelaDeEdital & " A, " & _
    gstrContribuicaoMelhoria & " B, " & _
    gstrContribuinte & " C, " & _
    gstrImobiliario & " D " & _
    " WHERE " & _
    " B.intTabelaDeEdital = A.PKId " & _
    " AND B.intImobiliario = D.PKId " & _
    " AND D.intContribuinte = C.PKId " & _
    "AND A.PKId = " & dbcintEdital.BoundText
    If chk_TodasInscricoes.Value = 0 Then
        strSql = strSql & " AND D.PKId = " & dbcintInscricaoCadastral.BoundText
    End If
    strSql = strSql & " AND D.intComposicao = " & dbcintContribuincaoMelhoria.BoundText & _
    " ORDER BY D.strInscricaoAnterior, C.strNome"
    
    strQuerryEditalNotificacao = strSql
End Function

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub
