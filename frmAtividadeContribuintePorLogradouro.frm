VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmAtividadeContribuintePorLogradouro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Atividades e Contribuintes por Logradouro"
   ClientHeight    =   2880
   ClientLeft      =   3105
   ClientTop       =   1905
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6360
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Relatório por"
      TabPicture(0)   =   "frmAtividadeContribuintePorLogradouro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Composicao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_intOcorrencia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dbc_intOcorrencia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbc_logradouro"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_OpcaoConsulta"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chk_TodosLogradouro"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk_TodasOcorrencias"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox chk_TodasOcorrencias 
         Caption         =   "&Todas as Ocorrências"
         Height          =   195
         Left            =   1110
         TabIndex        =   6
         Top             =   1470
         Width           =   2895
      End
      Begin VB.CheckBox chk_TodosLogradouro 
         Caption         =   "&Todos os Logradouros"
         Height          =   195
         Left            =   1110
         TabIndex        =   3
         Top             =   810
         Width           =   2895
      End
      Begin VB.Frame fra_OpcaoConsulta 
         Caption         =   "Opções Para Consulta"
         Height          =   675
         Left            =   180
         TabIndex        =   7
         Top             =   1800
         Width           =   5775
         Begin VB.CheckBox chk_AtividadesPrincipais 
            Caption         =   "Só Atividades &Principais"
            Height          =   195
            Left            =   150
            TabIndex        =   8
            Top             =   300
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo dbc_logradouro 
         Height          =   315
         Left            =   1110
         TabIndex        =   2
         Top             =   450
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intOcorrencia 
         Height          =   315
         Left            =   1110
         TabIndex        =   5
         Top             =   1110
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Label lbl_intOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   1170
         Width           =   780
      End
      Begin VB.Label lbl_Composicao 
         AutoSize        =   -1  'True
         Caption         =   "Logradouro"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   510
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmAtividadeContribuintePorLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSql As String
Dim mobjAux As Object

Private Sub chk_TodasOcorrencias_Click()
    If chk_TodasOcorrencias.Value = 0 Then
        TrocaCorObjeto dbc_intOcorrencia, False, False
        dbc_intOcorrencia.SetFocus
    Else
        dbc_intOcorrencia.Text = Empty
        TrocaCorObjeto dbc_intOcorrencia, True, True
    End If
End Sub

Private Sub chk_TodosLogradouro_Click()
    If chk_TodosLogradouro.Value = 0 Then
        TrocaCorObjeto dbc_logradouro, False, False
        dbc_logradouro.SetFocus
    Else
        dbc_logradouro.Text = Empty
        TrocaCorObjeto dbc_logradouro, True, True
    End If
End Sub

Private Sub dbc_logradouro_GotFocus()
    MarcaCampo dbc_logradouro
End Sub

Private Sub dbc_logradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_logradouro
End Sub

Private Sub dbc_intOcorrencia_GotFocus()
    MarcaCampo dbc_intOcorrencia
End Sub

Private Sub dbc_intOcorrencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intOcorrencia
End Sub

Private Function strQueryRelatorio() As String
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "LG.Strdescricao Logradouro, "
    strSql = strSql & "AEC.Strdescricao Atividade, "
    strSql = strSql & gstrRIGHT("EC.Strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " InsCad, "
    strSql = strSql & "CO.Strnome RazaoSocial, "
    strSql = strSql & "EC.Intnumero Num, "
    strSql = strSql & "EC.Strcomplemento Cmpl, "
    strSql = strSql & "BA.Strdescricao Bairro, "
    strSql = strSql & "OC.Strdescricao Ocorrencia, "
    strSql = strSql & "LG.Intcep Cep "
    strSql = strSql & "From "
    strSql = strSql & gstrEconomico & " EC, "
    strSql = strSql & gstrLogradouro & " LG, "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrBairro & " BA, "
    strSql = strSql & gstrOcorrencia & " OC, "
    strSql = strSql & gstrAtividadeDaEmpresa & " AEM, "
    strSql = strSql & gstrAtividadeEC & " AEC "
    strSql = strSql & "WHERE "
    If chk_TodosLogradouro.Value = 0 Then strSql = strSql & "EC.intLogradouro = " & dbc_logradouro.BoundText & " And "
    If chk_TodasOcorrencias.Value = 0 Then strSql = strSql & "EC.intOcorrencia = " & dbc_intOcorrencia.BoundText & " And "
    If chk_AtividadesPrincipais.Value = 1 Then strSql = strSql & "AEM.blnPrincipal = 1 And "
    strSql = strSql & "EC.Intlogradouro " & strOUTJSQLServer & "= LG.pkid " & strOUTJOracle & " And "
    strSql = strSql & "EC.Intcontribuinte " & strOUTJSQLServer & "= CO.pkid " & strOUTJOracle & " And "
    strSql = strSql & "EC.Intbairro " & strOUTJSQLServer & "= BA.pkid " & strOUTJOracle & " And "
    strSql = strSql & "EC.IntOcorrencia " & strOUTJSQLServer & "= OC.pkid " & strOUTJOracle & " And "
    strSql = strSql & "AEM.Inteconomico = EC.pkid And "
    strSql = strSql & "AEM.Intatividade = AEC.Pkid "
    strSql = strSql & "ORDER BY "
    strSql = strSql & "EC.Intlogradouro,"
    strSql = strSql & "EC.intNumero, "
    strSql = strSql & "EC.Strinscricaocadastral "
    
    strQueryRelatorio = strSql
    
End Function

Private Function strQueryLogradouro() As String
    strSql = ""
    strSql = strSql & "SELECT "
        strSql = strSql & "LG.Pkid, "
        strSql = strSql & "LG.strDescricao "
    strSql = strSql & "FROM "
        strSql = strSql & gstrLogradouro & " LG "
    strSql = strSql & "ORDER BY "
        strSql = strSql & "LG.strDescricao"
    
    strQueryLogradouro = strSql
End Function

Private Function strQueryOcorrencia() As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrOcorrencia & " "
    strSql = strSql & "WHERE intUtilizacaoDaOcorrencia = 5 "
    strSql = strSql & "ORDER BY strDescricao"
    strQueryOcorrencia = strSql
End Function

Private Sub Form_Activate()
    gintCodSeguranca = 1181
    If mobjAux Is Nothing Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
End Sub

Private Sub Form_Load()
    dbc_logradouro.Tag = strQueryLogradouro & ";strDescricao"
    dbc_intOcorrencia.Tag = strQueryOcorrencia & ";strDescricao"
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            ImprimeRelatorio rptAtividadeContribuintePorLogradouro, strQueryRelatorio
        End If
    ElseIf UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        PreencherListaDeOpcoes Me.ActiveControl
        Exit Sub
    End If
    
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If chk_TodosLogradouro.Value = 0 Then
        If dbc_logradouro.MatchedWithList = False Then
            ExibeMensagem "Selecione um Logradouro"
            dbc_logradouro.SetFocus
            Exit Function
        End If
    End If
    
    If chk_TodasOcorrencias.Value = 0 Then
        If dbc_intOcorrencia.MatchedWithList = False Then
            ExibeMensagem "Selecione uma Ocorrência"
            dbc_intOcorrencia.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
    
End Function

