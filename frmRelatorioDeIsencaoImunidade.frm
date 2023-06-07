VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioDeIsencaoImunidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Isenção Imunidade"
   ClientHeight    =   3285
   ClientLeft      =   2805
   ClientTop       =   6810
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5295
   Begin VB.Frame fra_Mensagem1 
      Caption         =   "Opções de Consulta"
      Height          =   2355
      Left            =   300
      TabIndex        =   5
      Top             =   570
      Width           =   4710
      Begin VB.Frame fra_Inscricao 
         Height          =   585
         Left            =   240
         TabIndex        =   8
         Top             =   330
         Width           =   4215
         Begin VB.OptionButton optTipoDeInscricao 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   0
            Left            =   585
            TabIndex        =   0
            Top             =   270
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optTipoDeInscricao 
            Caption         =   "Econômico"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   1
            Left            =   2475
            TabIndex        =   1
            Top             =   270
            Width           =   1125
         End
      End
      Begin VB.CheckBox chkTodasInscricoes 
         Caption         =   "Selecionar todas as inscrições"
         Height          =   255
         Left            =   1425
         TabIndex        =   4
         Top             =   1875
         Width           =   2835
      End
      Begin MSDataListLib.DataCombo dbcintInscricaoInicial 
         Height          =   315
         Left            =   1425
         TabIndex        =   2
         Top             =   1110
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintInscricaoFinal 
         Height          =   315
         Left            =   1425
         TabIndex        =   3
         Top             =   1515
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblFinal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Final:"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   1605
         Width           =   1065
      End
      Begin VB.Label lblInicial 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Inicial:"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   1170
         Width           =   1140
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Isenção Imunidade"
      TabPicture(0)   =   "frmRelatorioDeIsencaoImunidade.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
End
Attribute VB_Name = "frmRelatorioDeIsencaoImunidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bitOpcaoAtual As Byte

Public Sub MantemForm(ByVal strModoOperacao As String)

  Select Case UCase(strModoOperacao)
  
    Case UCase(gstrImprimir)
      If blnDadosOk Then
         ImprimeRelatorio rptCadIsencaoImunidade, strQueryRelatorio(IIf(optTipoDeInscricao(0).Value = True, 0, 1))
      End If
      
    Case UCase(gstrNovo)
      dbcintInscricaoInicial.Text = ""
      Set dbcintInscricaoInicial.RowSource = Nothing
      dbcintInscricaoFinal.Text = ""
      Set dbcintInscricaoFinal.RowSource = Nothing
      chkTodasInscricoes.Value = 0
      dbcintInscricaoInicial.SetFocus
      
    Case UCase(gstrPreencherLista)
      If Left(Me.ActiveControl.Name, 3) = "dbc" Then
          PreencherListaDeOpcoes Me.ActiveControl
      End If
      
  End Select
  
End Sub

Private Function blnDadosOk() As Boolean
  blnDadosOk = False
  If dbcintInscricaoFinal.MatchedWithList = False And dbcintInscricaoInicial.MatchedWithList = False And chkTodasInscricoes.Value = 0 Then
     ExibeMensagem "A inscrição deve ser informada."
     dbcintInscricaoInicial.SetFocus
     Exit Function
  End If
  If dbcintInscricaoInicial.MatchedWithList = True And dbcintInscricaoFinal.MatchedWithList = True Then
     If Int(dbcintInscricaoFinal.Text) < Int(dbcintInscricaoInicial.Text) Then
        ExibeMensagem "A inscrição inicial não pode ser maior que a inscrição final."
        dbcintInscricaoFinal.SetFocus
        Exit Function
     End If
  End If
  
  blnDadosOk = True
End Function

Private Sub chkTodasInscricoes_Click()
  If chkTodasInscricoes.Value = 1 Then
     TrocaCorObjeto dbcintInscricaoInicial, True
     TrocaCorObjeto dbcintInscricaoFinal, True
  Else
     TrocaCorObjeto dbcintInscricaoInicial, False
     TrocaCorObjeto dbcintInscricaoFinal, False
  End If
End Sub

Private Sub dbcintInscricaoFinal_Click(Area As Integer)
  DropDownDataCombo dbcintInscricaoFinal, Me, Area
End Sub

Private Sub dbcintInscricaoFinal_GotFocus()
  If Trim(dbcintInscricaoFinal.Text) = "" Then
     dbcintInscricaoFinal.Text = Trim(dbcintInscricaoInicial.Text)
     dbcintInscricaoFinal_Click 0
     dbcintInscricaoFinal.Text = Trim(dbcintInscricaoInicial.Text)
     MarcaCampo dbcintInscricaoFinal
  End If
End Sub

Private Sub dbcintInscricaoInicial_Click(Area As Integer)
  DropDownDataCombo dbcintInscricaoInicial, Me, Area
End Sub

Private Sub dbcintInscricaoInicial_GotFocus()
  If Trim(dbcintInscricaoInicial.Text) = "" Then
     dbcintInscricaoInicial.Text = Trim(dbcintInscricaoFinal.Text)
     dbcintInscricaoInicial_Click 0
     dbcintInscricaoInicial.Text = Trim(dbcintInscricaoFinal.Text)
     MarcaCampo dbcintInscricaoInicial
  End If
End Sub

Private Sub Form_Activate()
  HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrAplicar, gstrSalvar
  HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir, gstrPreencherLista
  If dbcintInscricaoInicial.Enabled Then
     dbcintInscricaoInicial.SetFocus
  End If
End Sub

Private Function strQueryInscricao(intIndex As Integer) As String
Dim strSql As String
    
  strSql = ""
  strSql = strSql & "SELECT "
  Select Case intIndex
      Case 0
          strSql = strSql & "Pkid, " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao "
      Case 1
          strSql = strSql & "Pkid, " & gstrRIGHT("strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & "  strInscricaoCadastral "
  End Select
  
  strSql = strSql & "FROM "
  Select Case intIndex
      Case 0
          strSql = strSql & gstrImobiliario & " ORDER BY strInscricao "
      Case 1
          strSql = strSql & gstrEconomico & " ORDER BY strInscricaoCadastral "
  End Select
  
  strQueryInscricao = strSql

End Function

Private Sub Form_Load()
  dbcintInscricaoInicial.Tag = strQueryInscricao(0) & ";strInscricao "
  dbcintInscricaoFinal.Tag = strQueryInscricao(0) & ";strInscricao "
End Sub

Private Sub optTipoDeInscricao_Click(Index As Integer)
  If Index <> bitOpcaoAtual Then
     Set dbcintInscricaoInicial.RowSource = Nothing
     Set dbcintInscricaoFinal.RowSource = Nothing
     dbcintInscricaoInicial.Text = ""
     dbcintInscricaoFinal.Text = ""
     chkTodasInscricoes.Value = 0
     chkTodasInscricoes_Click
  End If
  bitOpcaoAtual = Index
  If Index = 0 Then
        dbcintInscricaoInicial.Tag = strQueryInscricao(Index) & ";strInscricao "
        dbcintInscricaoFinal.Tag = strQueryInscricao(Index) & ";strInscricao "
  Else
        dbcintInscricaoInicial.Tag = strQueryInscricao(Index) & ";strInscricaoCadastral "
        dbcintInscricaoFinal.Tag = strQueryInscricao(Index) & ";strInscricaoCadastral "
  End If
End Sub

Private Function strQueryRelatorio(Index As Byte) As String
Dim strSql As String
Dim strOpcao As String
Dim strTabela As String
Dim strCampo As String
  
  strOpcao = ""
  
  If Index = 0 Then
     strTabela = gstrImobiliario
     strCampo = "TA.strInscricao "
  Else
     strTabela = gstrEconomico
     strCampo = "TA.strInscricaoCadastral "
  End If
  
  If chkTodasInscricoes.Value = 0 Then
     If dbcintInscricaoInicial.MatchedWithList = True And dbcintInscricaoFinal.MatchedWithList = True Then
        strOpcao = strCampo & " BETWEEN '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoInicial.Text)), "0") & Trim(dbcintInscricaoInicial.Text) & "' AND '"
        strOpcao = strOpcao & String(gintLenInscricao - Len(Trim(dbcintInscricaoFinal.Text)), "0") & Trim(dbcintInscricaoFinal.Text) & "' AND "
     Else
        If dbcintInscricaoInicial.MatchedWithList = True Then
           strOpcao = strCampo & " = '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoInicial.Text)), "0") & Trim(dbcintInscricaoInicial.Text) & "' AND "
        Else
           strOpcao = strCampo & " = '" & String(gintLenInscricao - Len(Trim(dbcintInscricaoFinal.Text)), "0") & Trim(dbcintInscricaoFinal.Text) & "' AND "
        End If
     End If
  End If
  
  strSql = ""
  strSql = strSql & "SELECT "
  strSql = strSql & "II.pkID, "
 
  If Index = 0 Then
     strSql = strSql & gstrRIGHT(strCampo, gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
  Else
     strSql = strSql & gstrRIGHT(strCampo, gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao, "
  End If
  
  strSql = strSql & "CO.StrNome strContribuinte, "
  strSql = strSql & "CR.strDescricao strComposicao, "
  strSql = strSql & "CR.intUtilizacao, "
  strSql = strSql & "TII.strDescricao strTipoIsencao "
  strSql = strSql & "FROM "
  strSql = strSql & "tblIsencaoImunidade II, "
  strSql = strSql & strTabela & " TA, "
  strSql = strSql & gstrComposicaoDaReceita & " CR, "
  strSql = strSql & gstrContribuinte & " CO, "
  strSql = strSql & gstrTipoIsencaoImunidade & " TII "
  strSql = strSql & "WHERE "
  
  strSql = strSql & strOpcao
  
  strSql = strSql & "II.intIdentificacao = TA.pkID AND "
  strSql = strSql & "CR.pkID = II.intComposicaoDaReceita AND "
  strSql = strSql & "CO.pkID = TA.intContribuinte AND "
  strSql = strSql & "TII.pkID = II.intTipoIsencaoImunidade "
  
  strSql = strSql & "ORDER BY "
  strSql = strSql & " strComposicao,strinscricao "
  
  strQueryRelatorio = strSql

End Function

