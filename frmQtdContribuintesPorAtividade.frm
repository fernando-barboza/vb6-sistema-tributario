VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmQtdContribuintesPorAtividade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quantidade de Contribuintes Por Atividade"
   ClientHeight    =   2220
   ClientLeft      =   2205
   ClientTop       =   2460
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5175
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   2205
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   3889
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Atividades"
      TabPicture(0)   =   "frmQtdContribuintesPorAtividade.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_CodigoDaUtilizacao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrNomeDoValor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_OpcaoConsulta"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbc_intAtividade2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbc_intAtividade1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin MSDataListLib.DataCombo dbc_intAtividade1 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   450
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intAtividade2 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   840
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Frame fra_OpcaoConsulta 
         Caption         =   "Opções Para Consulta"
         Height          =   885
         Left            =   720
         TabIndex        =   2
         Top             =   1200
         Width           =   4155
         Begin VB.CheckBox chk_AtividadesPrincipais 
            Caption         =   "Só Atividades &Principais"
            Height          =   195
            Left            =   150
            TabIndex        =   7
            Top             =   570
            Width           =   2895
         End
         Begin VB.CheckBox chk_TodasAtividades 
            Caption         =   "&Todas as Atividades"
            Height          =   195
            Left            =   150
            TabIndex        =   6
            Top             =   300
            Width           =   2895
         End
      End
      Begin VB.Label lblstrNomeDoValor 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   885
         Width           =   285
      End
      Begin VB.Label lbl_CodigoDaUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmQtdContribuintesPorAtividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSelecionou As Boolean
Dim mobjAux As Object
Dim mblnPrimeiraVez As Boolean
Dim strSql As String

Private Sub chk_TodasAtividades_Click()
    If chk_TodasAtividades.Value = 0 Then
        TrocaCorObjeto dbc_intAtividade1, False, False
        TrocaCorObjeto dbc_intAtividade2, False, False
    Else
        TrocaCorObjeto dbc_intAtividade1, True, True
        TrocaCorObjeto dbc_intAtividade2, True, True
        dbc_intAtividade1.BoundText = ""
        dbc_intAtividade2.BoundText = ""
    End If
End Sub

Private Sub dbc_intAtividade1_GotFocus()
    MarcaCampo dbc_intAtividade1
End Sub

Private Sub dbc_intAtividade1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intAtividade1
End Sub

Private Sub dbc_intAtividade2_GotFocus()
    MarcaCampo dbc_intAtividade2
End Sub

Private Sub dbc_intAtividade2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intAtividade2
End Sub

Private Sub Form_Load()
    
    'dbc_intAtividade1.BoundText = ""
    'dbc_intAtividade2.BoundText = ""
    chk_AtividadesPrincipais.Value = 0
    chk_TodasAtividades.Value = 0
    'TrocaCorObjeto dbc_intAtividade1, False, False
    'TrocaCorObjeto dbc_intAtividade2, False, False
    
     LeDaTabelaParaObj gstrAtividadeEC, dbc_intAtividade1, strQueryAtividade
     LeDaTabelaParaObj gstrAtividadeEC, dbc_intAtividade2, strQueryAtividade
    
End Sub

Private Function strQueryAtividade() As String
    strSql = ""
        strSql = strSql & "SELECT "
        strSql = strSql & "Pkid, "
        strSql = strSql & "strDescricao "
        strSql = strSql & "FROM "
        strSql = strSql & gstrAtividadeEC
        strSql = strSql & " ORDER BY "
        strSql = strSql & "strDescricao"
    strQueryAtividade = strSql
End Function

Private Function strQueryRelatorio() As String
    strSql = ""
    
        strSql = strSql & "SELECT "
        strSql = strSql & "GA.Strnomedogrupo Grupo, "
        strSql = strSql & "SGA.STRNOMEDOSUBGRUPO SubGrupo, "
        strSql = strSql & "AE.Intcodigo Codigo, "
        strSql = strSql & "AE.strDescricao Descricao , "
        strSql = strSql & "count(AEM.Intatividade) QTD "
        strSql = strSql & "FROM "
        strSql = strSql & gstrGrupoDeAtividade & " GA, "
        strSql = strSql & gstrSubGrupoDeAtividade & " SGA, "
        strSql = strSql & gstrAtividadeEC & " AE, "
        strSql = strSql & gstrAtividadeDaEmpresa & " AEM "
        strSql = strSql & "WHERE "
        If chk_TodasAtividades = 0 Then
            strSql = strSql & "AE.strDescricao "
            strSql = strSql & "BETWEEN "
            strSql = strSql & " '" & dbc_intAtividade1 & "' And "
            strSql = strSql & " '" & dbc_intAtividade2 & "' And "
        End If
        If chk_AtividadesPrincipais = 1 Then strSql = strSql & "AEM.blnPrincipal = 1 And "
            
        strSql = strSql & "AE.intgrupo = GA.pkid and "
        strSql = strSql & "AE.Intsubgrupo = SGA.Pkid and "
        strSql = strSql & "AEM.Intatividade = AE.Pkid "
        
        
        strSql = strSql & "GROUP BY "
        strSql = strSql & "GA.Strnomedogrupo, "
        strSql = strSql & "SGA.Strnomedosubgrupo, "
        strSql = strSql & "AE.Strdescricao, "
        strSql = strSql & "AE.intCodigo "
        strSql = strSql & "ORDER BY "
        strSql = strSql & "GA.strnomedogrupo, "
        strSql = strSql & "SGA.STRNOMEDOSUBGRUPO, "
        strSql = strSql & "AE.strDescricao"
        
    strQueryRelatorio = strSql
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1178
    If mblnSelecionou Then
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    If mobjAux Is Nothing Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then ImprimeRelatorio rptQtdContribuintesPorAtividade, strQueryRelatorio
    End If
End Sub

Private Sub Form_Deactivate()
HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
On Error GoTo err_blnDadosOK
If chk_TodasAtividades = 1 Then
    blnDadosOk = True
Else
    If dbc_intAtividade1.MatchedWithList = False Then
        ExibeMensagem "A Atividade Inicial não é Válida"
        dbc_intAtividade1.SetFocus
        Exit Function
    Else
        If dbc_intAtividade2.MatchedWithList = False Then
           ExibeMensagem "A Atividade Final Não é Válida"
           dbc_intAtividade2.SetFocus
           Exit Function
        End If
    End If
    blnDadosOk = True
End If
err_blnDadosOK:
End Function
