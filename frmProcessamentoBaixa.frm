VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Begin VB.Form frmProcessamentoBaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Processamento de Baixa"
   ClientHeight    =   3510
   ClientLeft      =   5040
   ClientTop       =   4050
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3285
      Left            =   90
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   135
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   5794
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Processamento de Baixa"
      TabPicture(0)   =   "frmProcessamentoBaixa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Status"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pgr_Status"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtPKId"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chk_Simulado"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Principal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chk_Criticas"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CheckBox chk_Criticas 
         Caption         =   "Visualizar Críticas "
         Height          =   255
         Left            =   195
         TabIndex        =   6
         Top             =   2940
         Value           =   1  'Checked
         Width           =   2970
      End
      Begin VB.Frame fra_Principal 
         Height          =   2025
         Left            =   195
         TabIndex        =   10
         Top             =   480
         Width           =   4845
         Begin VB.CheckBox chk_Todos 
            Caption         =   "Selecionar Todos"
            Height          =   255
            Left            =   1605
            TabIndex        =   5
            Top             =   1590
            Width           =   1575
         End
         Begin VB.CheckBox chk_TodasContas 
            Caption         =   "Selecionar Todas"
            Height          =   255
            Left            =   1605
            TabIndex        =   3
            Top             =   930
            Width           =   1605
         End
         Begin VB.TextBox txtdtmDataMovimento 
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
            Left            =   1605
            MaxLength       =   10
            TabIndex        =   0
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton cmd_ContaCorrente 
            Height          =   315
            Left            =   4380
            Picture         =   "frmProcessamentoBaixa.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Tag             =   "585"
            ToolTipText     =   "Clique para cadastar uma Conta Bancária"
            Top             =   600
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintContaBancaria 
            Height          =   315
            Left            =   1605
            TabIndex        =   1
            Top             =   600
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintLote 
            Height          =   315
            Left            =   1605
            TabIndex        =   4
            Top             =   1260
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label lblDataDoMovimento 
            AutoSize        =   -1  'True
            Caption         =   "Data do Movimento"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   330
            Width           =   1395
         End
         Begin VB.Label lblLote 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            Height          =   195
            Left            =   1170
            TabIndex        =   12
            Top             =   1365
            Width           =   315
         End
         Begin VB.Label lblContaBancaria 
            AutoSize        =   -1  'True
            Caption         =   "Conta Corrente"
            Height          =   195
            Left            =   450
            TabIndex        =   11
            Top             =   705
            Width           =   1065
         End
      End
      Begin VB.CheckBox chk_Simulado 
         Caption         =   "Simulado"
         Height          =   255
         Left            =   4125
         TabIndex        =   7
         Top             =   2940
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   15
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar pgr_Status 
         Height          =   165
         Left            =   210
         TabIndex        =   14
         Top             =   2550
         Visible         =   0   'False
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lbl_Status 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2730
         Visible         =   0   'False
         Width           =   4785
      End
   End
End
Attribute VB_Name = "frmProcessamentoBaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnExisteContaCorrente As Boolean

Dim vetPrevia()            As String

'blnExisteContaCorrente - Variável para verificar se existe uma conta corrente
'                         com a data de movimento informada.

Private Sub chk_TodasContas_Click()
    If chk_TodasContas.Value = 1 Then
        TrocaCorObjeto dbcintContaBancaria, True
        TrocaCorObjeto dbcintLote, True
        chk_Todos.Value = 1
        chk_Todos.Enabled = False
    Else
        TrocaCorObjeto dbcintContaBancaria, False
        TrocaCorObjeto dbcintLote, False
        chk_Todos.Value = 0
        chk_Todos.Enabled = True
    End If
    
End Sub

Private Sub chk_Todos_Click()
    If chk_Todos.Value = 1 Then
        TrocaCorObjeto dbcintLote, True
    Else
        TrocaCorObjeto dbcintLote, False
    End If
End Sub

Private Sub cmd_ContaCorrente_Click()
    If dbcintContaBancaria.Enabled Then dbcintContaBancaria.SetFocus
    CarregaForm frmCadContasBancarias, dbcintContaBancaria
End Sub

Private Sub dbcintcontabancaria_Change()
    If dbcintContaBancaria.MatchedWithList Then
        dbcintLote.ListField = ""
        Set dbcintLote.RowSource = Nothing
    End If
End Sub

Private Sub dbcintcontabancaria_Click(Area As Integer)
    DropDownDataCombo dbcintContaBancaria, Me, Area
End Sub

Private Sub dbcintContaBancaria_GotFocus()
    MarcaCampo dbcintContaBancaria
    dbcintContaBancaria.Tag = strQueryContaCorrente & ";CB.strConta"
End Sub

Private Sub dbcintContaBancaria_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContaBancaria, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContaBancaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContaBancaria
End Sub

Private Sub dbcintLote_Click(Area As Integer)
    DropDownDataCombo dbcintLote, Me, Area
End Sub

Private Sub dbcintLote_GotFocus()
    MarcaCampo dbcintLote
    dbcintLote.Tag = strQueryLotes & ";MB.intLote"
End Sub

Private Sub dbcintLote_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLote, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintLote
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1132
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrProcessamentoBaixa
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar, gstrDeletar
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrProcessamentoBaixa
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrProcessamentoBaixa
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
        
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        Case Is = UCase(gstrNovo)
            LimpaObjeto Me
            txtdtmDataMovimento.SetFocus
        Case Is = UCase(gstrProcessamentoBaixa)
            If Not blnDadosOk Then Exit Sub
            
            If chk_Simulado.Value = 0 Then
                If gblnExclusaoGravacaoOk("", "Confirma Processamento de Baixa?", True) Then
                    If GravaLancamentoValor(False) Then
                        ExibePreviaDados
                    End If
                End If
            Else
                If gblnExclusaoGravacaoOk("", "Confirma Processamento de Baixa Simulado?", True) Then
                    'If GravaLancamentoValorSimulado Then
                    If GravaLancamentoValor(True) Then
                        ExibePreviaDados
                    End If
                End If
            End If
            
    End Select
                 
End Sub

Private Sub txtdtmDataMovimento_GotFocus()
    If txtdtmDataMovimento = "" Then txtdtmDataMovimento = gstrDataDoSistema
    MarcaCampo txtdtmDataMovimento
End Sub

Private Sub txtdtmDataMovimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataMovimento
End Sub

Private Sub txtdtmDataMovimento_LostFocus()
    txtdtmDataMovimento = gstrDataFormatada(txtdtmDataMovimento)
    dbcintContaBancaria.Text = ""
    dbcintContaBancaria.ListField = ""
    If txtdtmDataMovimento <> "" Then strQueryContaCorrente (True)
End Sub

Private Function strQueryContaCorrente(Optional blnVerificaContaCorrente As Boolean) As String
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset

    strSql = "SELECT Distinct CB.Pkid, "
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CB.strConta") & strCONCAT & "'-'" & strCONCAT & " CB.strDigitoVerificador ContaCorrente"
    strSql = strSql & " FROM " & gstrContaBancaria & " CB, "
    strSql = strSql & gstrMovimentoBancario & " MB " & strREADPAST
    strSql = strSql & " WHERE"
    strSql = strSql & " MB.intContaBancaria " & strOUTJSQLServer & "= CB.Pkid " & strOUTJOracle & " AND"
    strSql = strSql & " MB.dtmDtMovimento = " & gstrConvDtParaSql(txtdtmDataMovimento.Text)
    strSql = strSql & " ORDER BY ContaCorrente"
    
    If blnVerificaContaCorrente Then
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                blnExisteContaCorrente = True
            Else
                blnExisteContaCorrente = False
            End If
        End If
    End If
    
    strQueryContaCorrente = strSql

End Function

Private Function strQueryLotes() As String
Dim strSql As String

    strSql = "SELECT MB.intLote,"
    strSql = strSql & " MB.intLote"
    strSql = strSql & " FROM "
    strSql = strSql & gstrContaBancaria & " CB, "
    strSql = strSql & gstrMovimentoBancario & " MB " & strREADPAST
    strSql = strSql & " WHERE"
    strSql = strSql & " MB.intContaBancaria " & strOUTJSQLServer & "= CB.Pkid " & strOUTJOracle & " AND"
    strSql = strSql & " MB.dtmDtMovimento = " & gstrConvDtParaSql(txtdtmDataMovimento)
    
    If chk_TodasContas.Value = 0 Then
        If dbcintContaBancaria.MatchedWithList Then
            strSql = strSql & " AND MB.intContaBancaria = " & dbcintContaBancaria.BoundText
        End If
    End If
    
    strSql = strSql & " GROUP BY MB.intContaBancaria, MB.intLote"
    strSql = strSql & " ORDER BY MB.intLote"

    strQueryLotes = strSql
    
End Function

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False

    If txtdtmDataMovimento.Text = "" Then
        ExibeMensagem "É necessário informar a Data do Movimento."
        txtdtmDataMovimento.SetFocus
        Exit Function
    Else
        If Not gblnDataValida(txtdtmDataMovimento, True) Then Exit Function
    End If
    
    If Not blnExisteContaCorrente Then
        ExibeMensagem "Não existe nenhum Resumo Bancário com esta Data de Movimento."
        txtdtmDataMovimento.SetFocus
        Exit Function
    End If
    
    If chk_TodasContas.Value = 0 Then
        If Not dbcintContaBancaria.MatchedWithList Then
            ExibeMensagem "Selecione uma Conta Corrente válida."
            If dbcintContaBancaria.Enabled Then dbcintContaBancaria.SetFocus
            Exit Function
        End If
    End If
    
    If chk_Todos.Value = 0 Then
        If Not dbcintLote.MatchedWithList Then
                ExibeMensagem "Selecione um Lote válido."
                If dbcintLote.Enabled Then dbcintLote.SetFocus
                Exit Function
        End If
    End If

    blnDadosOk = True

End Function

Private Sub ExibePreviaDados()
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset

'    strSql = "SELECT "
'    strSql = strSql & gstrCONVERT(cdt_numeric, "RB.DBLVALOR") & " RECEBIDO, "
'    strSql = strSql & gstrCONVERT(cdt_numeric, "MO.Baixado") & " Baixado,"
'    strSql = strSql & gstrCONVERT(cdt_numeric, "(MO.Baixado - RB.dblValor)") & " Diferenca, "
'    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CB.strConta") & strCONCAT & "'-'" & strCONCAT & " CB.strDigitoVerificador Conta,"
'    strSql = strSql & " RB.intLote Lote"
'    strSql = strSql & " FROM "
'    strSql = strSql & gstrResumoBancario & " RB, "
'    strSql = strSql & gstrContaBancaria & " CB, "
'    strSql = strSql & " (SELECT MB.INTCONTABANCARIA "
'    strSql = strSql & " INTCONTABANCARIA,"
'    strSql = strSql & " MB.INTLOTE LOTE,"
'    strSql = strSql & " SUM(MB.dblPrincipal + "
'    strSql = strSql & " MB.dblMulta + "
'    strSql = strSql & " MB.dblJuros + "
'    strSql = strSql & " MB.dblCorrecao) Baixado "
'    strSql = strSql & " FROM "
'    strSql = strSql & gstrMovimentoBancario & " MB"
'    strSql = strSql & " WHERE"
'    strSql = strSql & " MB.DTMDTMOVIMENTO = " & gstrConvDtParaSql(txtdtmDataMovimento.Text)
'    If chk_TodasContas.Value = 0 Then
'        strSql = strSql & " AND MB.intContaBancaria = " & Val(dbcintContaBancaria.BoundText)
'    End If
'    If chk_Todos.Value = 0 Then
'        strSql = strSql & " AND MB.intLote = " & Val(dbcintLote.BoundText)
'    End If
'    strSql = strSql & " GROUP BY MB.intContaBancaria, MB.intLote) MO"
'    strSql = strSql & " WHERE"
'    strSql = strSql & " RB.dtmData = " & gstrConvDtParaSql(txtdtmDataMovimento.Text) & " AND"
'    strSql = strSql & " RB.intContaBancaria = Mo.intContaBancaria AND"
'    strSql = strSql & " RB.intLote = MO.Lote AND"
'    strSql = strSql & " RB.intContaBancaria " & strOUTJSQLServer & "= CB.Pkid " & strOUTJOracle
'
'    If chk_TodasContas.Value = 0 Then
'        strSql = strSql & " AND RB.intContaBancaria = " & Val(dbcintContaBancaria.BoundText)
'    End If
'    If chk_Todos.Value = 0 Then
'        strSql = strSql & " AND RB.intLote = " & Val(dbcintLote.BoundText)
'    End If
'
'    ImprimeRelatorio rptResumoProcessamentoBaixa, strSql, "Resumo do Processamento da Baixa de " & txtdtmDataMovimento.Text
    ImprimeRelatorioPorArray rptResumoProcessamentoBaixa, vetPrevia, "Resumo do Processamento da Baixa de " & txtdtmDataMovimento.Text

End Sub

Private Function GravaLancamentoValor(blnSimulado As Boolean) As Boolean
Dim strSql                   As String
Dim adoResultado             As ADODB.Recordset
Dim blnRollback              As Boolean
Dim intFor                   As Integer
Dim lngPkidPagamentoInicial  As Long
Dim lngPkidCriticasInicial   As Long

    blnRollback = False
    
    Screen.MousePointer = vbHourglass
    
    If gobjBanco.CriaADO(strMovimentosBancarios(blnSimulado), 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                
                ReDim vetPrevia(4, 0)
                
                'Vamos armazenar o Pkid inicial para nao duplicar a exibicao das criticas
                lngPkidCriticasInicial = glngPegaUltimaChave(gstrCriticaBaixa, "Pkid")
                
                Set gobjBanco = New clsBanco
                
                gobjBanco.ExecutaBeginTrans
                
                pgr_Status.Value = 0
                pgr_Status.Visible = True
                pgr_Status.Max = .RecordCount
                lbl_Status.Visible = True
                
                Do While Not .EOF

                    'gobjBanco.ExecutaBeginTrans
                    
                    If Not IsNull(!intLancamentoValor) Then
                    
                        If gblnBaixaCancelamento(gstrENulo(!intAlfa), gstrENulo(!Composicao), Year(txtdtmDataMovimento), gstrENulo(!intParcela), gstrENulo(!dtmDtPagamento), blnSimulado, blnSimulado, gstrENulo(!PkidMovBancario)) Then
                            
                            If gblnAnaliseDaReceita(gstrENulo(!intLancamentoValor), gstrENulo(!intContaBancaria), gstrENulo(!Composicao), Val(gstrConvVrParaSql(!dblPrincipal)), Val(gstrConvVrParaSql(!dblMulta)), Val(gstrConvVrParaSql(!dblJuros)), Val(gstrConvVrParaSql(!dblCorrecao)), gstrENulo(!Dtmdtmovimento), gstrENulo(!intAlfa), blnSimulado, False, adoResultado.RecordCount = adoResultado.AbsolutePosition, gstrENulo(!PkidMovBancario)) Then
                                strSql = ""
                                strSql = "INSERT INTO tblLancamentoPagamento(intLancamentoValor, "
                                strSql = strSql & " dblValorPrincipal, "
                                strSql = strSql & " dblValorMulta, "
                                strSql = strSql & " dblValorJuros, "
                                strSql = strSql & " dblValorCorrecao, "
                                strSql = strSql & " dblValorCorreto, "
                                strSql = strSql & " dtmDtPagamento, "
                                strSql = strSql & " dtmDtMovimento, "
                                strSql = strSql & " dtmDtAtualizacao, "
                                strSql = strSql & " lngCodUsr, "
                                strSql = strSql & " intCodigoBaixa, "
                                strSql = strSql & " strObservacao "
                                strSql = strSql & ") "
    
                                strSql = strSql & "VaLues( "
                                strSql = strSql & gstrENulo(!intLancamentoValor) & ", "
                                strSql = strSql & gstrConvVrParaSql(Val(gstrConvVrParaSql(!dblPrincipal)) + Val(gstrConvVrParaSql(!dblTarifa))) & ", "
                                'strsql = strsql & gstrConvVrParaSql(!dblPrincipal) & ", "
                                strSql = strSql & gstrConvVrParaSql(!dblMulta) & ", "
                                strSql = strSql & gstrConvVrParaSql(!dblJuros) & ", "
                                strSql = strSql & gstrConvVrParaSql(!dblCorrecao) & ", "
                                strSql = strSql & gstrConvVrParaSql(!dblCorreto) & ", "
                                strSql = strSql & gstrConvDtParaSql(!dtmDtPagamento) & ", "
                                strSql = strSql & gstrConvDtParaSql(!Dtmdtmovimento) & ", "
                                strSql = strSql & strGETDATE & ", "
                                strSql = strSql & glngCodUsr & ", "
                                strSql = strSql & gstrENulo(!intCodigoBaixa) & ", "
                                strSql = strSql & "'') "
    
                                Set gobjBanco = New clsBanco
                                If Not gobjBanco.Execute(strSql) Then
                                    ExibeMensagem "Ocorreu um erro na gravação dos registro em Lançamento de Pagamento, a operação não foi concluída."
                                    gobjBanco.ExecutaRollbackTrans
                                    Exit Function
                                End If
                                
                                'Vamos atribuir o flag de processado no registro
                                strSql = ""
                                strSql = "UPDATE " & gstrMovimentoBancario & " Set bitProcessado = 1 WHERE Pkid = " & gstrENulo(!PkidMovBancario)
    
                                Set gobjBanco = New clsBanco
                                If Not gobjBanco.Execute(strSql) Then
                                    ExibeMensagem "Ocorreu um erro na alteração do registro em Movimento Bancário, a operação não foi concluída."
                                    gobjBanco.ExecutaRollbackTrans
                                    Exit Function
                                End If
                                
                                'Vamos armazenar o Pkid inicial para nao acusar como critica de duplicacao
                                If adoResultado.AbsolutePosition = 1 Then
                                    lngPkidPagamentoInicial = glngRetornaPkidTabelaPai("seqtblLancamentoPagamento", gstrLancamentoPagamento)
                                End If
                                
                                'Vamos verificar se é um Acordo
                                If adoResultado("intUtilizacao").Value = TYP_ACORDO Then
                                    gQuitacaoDeAcordos adoResultado("intAlfa").Value, adoResultado("dtmDtPagamento").Value, adoResultado("dtmDtMovimento").Value
                                End If
                                
                                'Vamos carregar o array para o relatorio de Previa com os movimentos analisados
                                intFor = CarregaArrayDePrevia(!Conta, !Lote, !intContaBancaria, !dblPrincipal, !dblMulta, !dblJuros, !dblCorrecao, intFor, !intNumeroConta)
                                
                            Else
                                blnRollback = True
                            End If
                        Else
                            blnRollback = True
                        End If
                     
                     Else
                        If adoResultado.RecordCount = adoResultado.AbsolutePosition And Not aAnaliseReceita Is Nothing Then
                            If Not gblnAnaliseDaReceita(0, 0, 0, 0, 0, 0, 0, gstrENulo(!Dtmdtmovimento), 0, blnSimulado, False, adoResultado.RecordCount = adoResultado.AbsolutePosition, 0, True) Then
                                blnRollback = True
                            End If
                        End If
                        'Vamos carregar o array para o relatorio de Previa com os movimentos analisados
                        intFor = CarregaArrayDePrevia(!Conta, !Lote, !intContaBancaria, 0, 0, 0, 0, intFor, !intNumeroConta)
                        
                     End If
                     
                    DoEvents
                    lbl_Status.Caption = .AbsolutePosition & " de " & .RecordCount
                    Me.Refresh
            
                    pgr_Status.Value = .AbsolutePosition
                     
                    .MoveNext
                    
                Loop
                
                'Vamos finalizar o array de receitas que é carregado dentro da funcao gblnAnaliseDaReceita
                Set aAnaliseReceita = Nothing
                
                MontaCriticas lngPkidCriticasInicial, lngPkidPagamentoInicial
                
                If Not blnRollback And Not blnSimulado Then
                    gobjBanco.ExecutaCommitTrans
                Else
                    If Not blnSimulado Then
                        ExibeMensagem "Foi(ram) encontrada(s) crítica(s) no processamento, a operação não foi concluída."
                    End If
                    gobjBanco.ExecutaRollbackTrans
                End If
                    
                GravaLancamentoValor = True
                
            End If
        End With
    End If
    
    Screen.MousePointer = vbDefault
    
End Function

Private Sub MontaCriticas(lngPkidCriticaInicial As Long, Optional lngPkidPagamentoInicial As Long)
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset

'    If chk_Simulado.Value = 0 Then
        
        'Critica - Não Encontrada
        strSql = ""
        
        strSql = "INSERT INTO "
        strSql = strSql & gstrCriticaBaixa
        strSql = strSql & " (intMovimentoBancario,"
        strSql = strSql & " intTipoCritica,"
        strSql = strSql & " dtmDtAtualizacao,"
        strSql = strSql & " lngCodUsr)"
        
        If bytDBType = SQLServer Then
            strSql = strSql & " SELECT  MB.Pkid, 9, "
            strSql = strSql & strGETDATE & ", 1"
            strSql = strSql & " FROM tblContaBancaria CB, tblMovimentoBancario MB " & strREADPAST & " , tblLancamentoValor LV "
            strSql = strSql & " WHERE MB.DTMDTMOVIMENTO = " & gstrConvDtParaSql(txtdtmDataMovimento.Text) & " AND MB.Intcontabancaria = CB.Pkid AND "
            If chk_TodasContas.Value = 0 Then
                strSql = strSql & " MB.intContaBancaria = " & Val(dbcintContaBancaria.BoundText) & " AND"
                If chk_Todos.Value = 0 Then
                    strSql = strSql & " MB.intLote = " & Val(dbcintLote.BoundText) & " AND"
                End If
            End If
            strSql = strSql & " Not MB.intLancamentoValor Is Null AND  MB.intLancamentoValor not in (SELECT pkid FROM tbllancamentovalor) "
        Else
            strSql = strSql & " SELECT MO.Pkid,"
            strSql = strSql & " 9, "
            strSql = strSql & strGETDATE & ", "
            strSql = strSql & glngCodUsr
            strSql = strSql & " FROM ("
            strSql = strSql & " SELECT MB.Pkid,"
            strSql = strSql & " MB.intLancamentoValor, "
            strSql = strSql & gstrISNULL("LV.Pkid", "0") & " PkidLancamentoValor"
            strSql = strSql & " FROM "
            strSql = strSql & gstrContaBancaria & " CB, "
            strSql = strSql & gstrMovimentoBancario & " MB " & strREADPAST & ", "
            strSql = strSql & gstrLancamentoValor & " LV"
            strSql = strSql & " WHERE MB.DTMDTMOVIMENTO = "
            strSql = strSql & gstrConvDtParaSql(txtdtmDataMovimento.Text) & " AND"
            strSql = strSql & " MB.Intcontabancaria = CB.Pkid AND"
            strSql = strSql & " Not MB.intLancamentoValor Is Null AND "
            If chk_TodasContas.Value = 0 Then
                strSql = strSql & " MB.intContaBancaria = " & Val(dbcintContaBancaria.BoundText) & " AND"
                If chk_Todos.Value = 0 Then
                    strSql = strSql & " MB.intLote = " & Val(dbcintLote.BoundText) & " AND"
                End If
            End If
    
            strSql = strSql & " MB.intLancamentoValor " & strOUTJSQLServer & "= LV.Pkid " & strOUTJOracle & " ) MO"
            strSql = strSql & " WHERE MO.PkidLancamentoValor = 0"
        End If
        
        Set gobjBanco = New clsBanco
            
        gobjBanco.Execute strSql
    
        'Critica - Duplicada
        strSql = ""
    
        strSql = "INSERT INTO " & gstrCriticaBaixa
        strSql = strSql & " (intMovimentoBancario,"
        strSql = strSql & " intTipoCritica,"
        strSql = strSql & " dtmDtPagamentoAnterior,"
        strSql = strSql & " dtmDtAtualizacao,"
        strSql = strSql & " lngCodUsr)"
        strSql = strSql & " SELECT MAX(MO.Pkid),"
        strSql = strSql & " 2,"
        strSql = strSql & " MAX(LP.dtmDtPagamento), "
        strSql = strSql & strGETDATE & ","
        strSql = strSql & glngCodUsr
        strSql = strSql & " FROM "
        strSql = strSql & gstrResumoBancario & " RB, "
        strSql = strSql & gstrContaBancaria & " CB, "
        strSql = strSql & gstrLancamentoValor & " LV, "
        strSql = strSql & gstrLancamentoPagamento & " LP, "
        strSql = strSql & " (SELECT MB.PKid,"
        strSql = strSql & " MB.intLancamentoValor,"
        strSql = strSql & " MB.INTCONTABANCARIA  INTCONTABANCARIA,"
        strSql = strSql & " MB.INTLOTE LOTE,"
        strSql = strSql & " MB.DTMDTMOVIMENTO"
        strSql = strSql & " FROM "
        strSql = strSql & gstrContaBancaria & " CB, "
        strSql = strSql & gstrMovimentoBancario & " MB " & strREADPAST
        strSql = strSql & " WHERE  MB.DTMDTMOVIMENTO = " & gstrConvDtParaSql(txtdtmDataMovimento.Text)
        strSql = strSql & " AND MB.Intcontabancaria = CB.Pkid "
        If chk_TodasContas.Value = 0 Then
            strSql = strSql & " AND MB.intContaBancaria = " & Val(dbcintContaBancaria.BoundText)
            If chk_Todos.Value = 0 Then
                strSql = strSql & " AND MB.intLote = " & Val(dbcintLote.BoundText)
            End If
        End If
        strSql = strSql & " ) MO"
        strSql = strSql & " WHERE RB.dtmData = " & gstrConvDtParaSql(txtdtmDataMovimento.Text) & " AND"
        strSql = strSql & " RB.intContaBancaria = Mo.intContaBancaria AND"
        strSql = strSql & " RB.intLote = MO.Lote AND"
        strSql = strSql & " RB.intContaBancaria " & strOUTJSQLServer & "= CB.Pkid " & strOUTJOracle & " AND"
        strSql = strSql & " LV.Pkid = MO.intLancamentoValor AND"
        strSql = strSql & " LP.intLancamentoValor = MO.intLancamentoValor AND "
        strSql = strSql & " LP.pkid < " & lngPkidPagamentoInicial
        strSql = strSql & " GROUP BY MO.intLancamentoValor"
        
        Set gobjBanco = New clsBanco
            
        gobjBanco.Execute strSql
    
 '   Else
    
        'Critica - Não importado pelo disquete do banco
        strSql = ""
        
        strSql = "INSERT INTO "
        strSql = strSql & gstrCriticaBaixa
        strSql = strSql & " (intMovimentoBancario,"
        strSql = strSql & " intTipoCritica,"
        strSql = strSql & " dtmDtAtualizacao,"
        strSql = strSql & " lngCodUsr)"
        strSql = strSql & " SELECT MB.Pkid,"
        strSql = strSql & " 6, "
        strSql = strSql & strGETDATE & ", "
        strSql = strSql & glngCodUsr
        strSql = strSql & " FROM "
        strSql = strSql & gstrMovimentoBancario & " MB " & strREADPAST
        strSql = strSql & " WHERE MB.DTMDTMOVIMENTO = "
        strSql = strSql & gstrConvDtParaSql(txtdtmDataMovimento.Text) & " AND"
        strSql = strSql & " MB.bitGuia = 0 "
        If chk_TodasContas.Value = 0 Then
            strSql = strSql & " AND MB.intContaBancaria = " & Val(dbcintContaBancaria.BoundText)
            If chk_Todos.Value = 0 Then
                strSql = strSql & " AND MB.intLote = " & Val(dbcintLote.BoundText)
            End If
        End If

        Set gobjBanco = New clsBanco
            
        gobjBanco.Execute strSql
    
'    End If
    
    If chk_Criticas.Value = 1 Then
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strQueryCritica(lngPkidCriticaInicial), 5, adoResultado) Then

            If Not adoResultado.EOF Then
                ImprimeRelatorio rptCriticas, strQueryCritica(lngPkidCriticaInicial)
            End If
            
        End If
        
    End If
    
End Sub

Private Function GravaLancamentoValorSimulado() As Boolean
Dim strSql                   As String
Dim adoResultado             As ADODB.Recordset
Dim blnRollback              As Boolean
Dim intFor                   As Integer
Dim lngPkidCriticasInicial   As Long

    blnRollback = False
    
    Screen.MousePointer = vbHourglass

    If gobjBanco.CriaADO(strMovimentosBancarios(True), 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
            
                ReDim vetPrevia(4, 0)

                'Vamos armazenar o Pkid inicial para nao duplicar a exibicao das criticas
                lngPkidCriticasInicial = glngPegaUltimaChave(gstrCriticaBaixa, "Pkid")
                
                Set gobjBanco = New clsBanco
                
                pgr_Status.Value = 0
                pgr_Status.Visible = True
                pgr_Status.Max = .RecordCount
                lbl_Status.Visible = True
                
                Do While Not .EOF
                    
                    gobjBanco.ExecutaBeginTrans
                    
                    If Not IsNull(!intLancamentoValor) Then
                    
                        If gblnBaixaCancelamento(gstrENulo(!intAlfa), gstrENulo(!Composicao), Year(txtdtmDataMovimento), gstrENulo(!intParcela), gstrENulo(!dtmDtPagamento), True, True, gstrENulo(!PkidMovBancario)) Then
                            If gblnAnaliseDaReceita(gstrENulo(!intLancamentoValor), gstrENulo(!intContaBancaria), gstrENulo(!Composicao), Val(gstrConvVrParaSql(!dblPrincipal)), Val(gstrConvVrParaSql(!dblMulta)), Val(gstrConvVrParaSql(!dblJuros)), Val(gstrConvVrParaSql(!dblCorrecao)), gstrENulo(!dtmDtPagamento), gstrENulo(!intAlfa), True, False, adoResultado.RecordCount = adoResultado.AbsolutePosition, gstrENulo(!PkidMovBancario)) Then
                                                            
                                'Vamos carregar o array para o relatorio de Previa com os movimentos analisados
                                intFor = CarregaArrayDePrevia(!Conta, !Lote, !intContaBancaria, !dblPrincipal, !dblMulta, !dblJuros, !dblCorrecao, intFor, !intNumeroConta)
                                
                            Else
                                blnRollback = True
                            End If
                        Else
                            blnRollback = True
                        End If
                    
                    Else
                       
                       'Vamos carregar o array para o relatorio de Previa com os movimentos analisados
                       intFor = CarregaArrayDePrevia(!Conta, !Lote, !intContaBancaria, 0, 0, 0, 0, intFor, !intNumeroConta)
                    
                    End If
                    
                    DoEvents
                    lbl_Status.Caption = .AbsolutePosition & " de " & .RecordCount
                    Me.Refresh
            
                    pgr_Status.Value = .AbsolutePosition
                    
                    .MoveNext
                    
                Loop
                
                'Vamos finalizar o array de receitas que é carregado dentro da funcao gblnAnaliseDaReceita
                Set aAnaliseReceita = Nothing
                
                If blnRollback Then
                    ExibeMensagem "Foi(ram) encontrada(s) crítica(s) no processamento."
                End If
                
                gobjBanco.ExecutaRollbackTrans
                
                GravaLancamentoValorSimulado = True
                
                MontaCriticas lngPkidCriticasInicial
                
            End If

        End With
    End If
    
    Screen.MousePointer = vbDefault
    
End Function
    
Private Function strQueryCritica(lngPkidCriticasInicial As Long) As String
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "MB.Dtmdtmovimento, "
    strSql = strSql & "B.STRDESCRICAO, "
    If bytDBType = Oracle Then
        strSql = strSql & "Trim(" & gstrCONVERT(CDT_VARCHAR, "CB.strConta") & ")" & strCONCAT & "'-'" & strCONCAT & " Trim(CB.strDigitoVerificador) ContaCorrente, "
    Else
        strSql = strSql & "LTrim(RTrim(" & gstrCONVERT(CDT_VARCHAR, "CB.strConta") & "))" & strCONCAT & "'-'" & strCONCAT & " LTrim(RTrim(CB.strDigitoVerificador)) ContaCorrente, "
    End If
    strSql = strSql & "A.STRDESCRICAO AS strAgencia, "
    strSql = strSql & "MB.INTLOTE, "
    strSql = strSql & "LA.Strinscricao, "
    strSql = strSql & "LA.Strcomposicaodareceita, "
    strSql = strSql & "LA.Intexercicio, "
    strSql = strSql & "LA.intUtilizacao, "
    strSql = strSql & "LA.strNumeroAviso, "
    strSql = strSql & "TRB.strdescricao " & strCONCAT & " case when CRB.strDetalhe is not null then ': ' " & strCONCAT & "CRB.strDetalhe End AS strTipoCritica, "
    strSql = strSql & "MB.Dtmdtpagamento, "
    strSql = strSql & "MB.strCodigoDeBarras, "
    strSql = strSql & "(MB.DBLPRINCIPAL+MB.DBLMULTA+MB.DBLJUROS+MB.DBLCORRECAO) dblValor, "
    strSql = strSql & "CRB.DTMDTPAGAMENTOANTERIOR, "
    strSql = strSql & "CRB.intTipoCritica, "
    strSql = strSql & "LV.intParcela "
    
    If bytDBType = SQLServer Then
        strSql = strSql & " FROM tblCriticaBaixa CRB INNER JOIN "
        strSql = strSql & gstrMovimentoBancario & " MB " & strREADPAST & " ON CRB.INTMOVIMENTOBANCARIO = MB.Pkid INNER JOIN "
        strSql = strSql & gstrContaBancaria & " CB ON MB.intContaBancaria = CB.PKId INNER JOIN "
        strSql = strSql & gstrBanco & " B ON CB.intBanco = B.PKId INNER JOIN "
        strSql = strSql & gstrAgencia & " A ON CB.intAgencia = A.PKId INNER JOIN "
        strSql = strSql & gstrTipoCriticaBaixa & " TRB ON CRB.INTTIPOCRITICA = TRB.PKId LEFT OUTER JOIN "
        strSql = strSql & gstrLancamentoValor & " LV ON MB.intlancamentovalor = LV.PKId LEFT OUTER JOIN "
        strSql = strSql & gstrLancamentoAlfa & " LA ON LV.intLancamentoAlfa = LA.PKId "
        strSql = strSql & " WHERE MB.Dtmdtmovimento = " & gstrConvDtParaSql(txtdtmDataMovimento.Text)
        strSql = strSql & " AND CRB.Pkid > " & lngPkidCriticasInicial
    Else
        strSql = strSql & " FROM "
        strSql = strSql & gstrCriticaBaixa & " CRB, "
        strSql = strSql & gstrMovimentoBancario & " MB " & strREADPAST & ", "
        strSql = strSql & gstrLancamentoValor & " LV, "
        strSql = strSql & gstrLancamentoAlfa & " LA, "
        strSql = strSql & gstrContaBancaria & " CB, "
        strSql = strSql & gstrBanco & " B, "
        strSql = strSql & gstrAgencia & " A, "
        strSql = strSql & gstrTipoCriticaBaixa & " TRB"
        strSql = strSql & " WHERE "
        strSql = strSql & "CRB.Pkid > " & lngPkidCriticasInicial & " AND "
        strSql = strSql & "MB.Pkid = CRB.intMovimentoBancario AND "
        strSql = strSql & "LV.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " MB.Intlancamentovalor AND "
        strSql = strSql & "LA.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LV.Intlancamentoalfa AND "
        strSql = strSql & "CB.Pkid = MB.Intcontabancaria AND "
        strSql = strSql & "B.Pkid = CB.Intbanco AND "
        strSql = strSql & "A.Pkid = CB.INTAGENCIA AND "
        strSql = strSql & "TRB.pkid = CRB.INTTIPOCRITICA AND "
        strSql = strSql & "MB.Dtmdtmovimento = " & gstrConvDtParaSql(txtdtmDataMovimento.Text)
    End If
    
    If chk_TodasContas.Value = 0 Then
        strSql = strSql & " AND CB.Pkid =" & dbcintContaBancaria.BoundText
        If chk_Todos.Value = 0 Then
            strSql = strSql & " AND MB.INTLOTE =" & dbcintLote.BoundText
        End If
    End If
    
    strSql = strSql & " ORDER BY "
    strSql = strSql & "MB.Dtmdtmovimento, "
    strSql = strSql & "B.STRDESCRICAO, "
    strSql = strSql & "CB.STRCONTA, "
    strSql = strSql & "MB.INTLOTE "
    
strQueryCritica = strSql

End Function

Private Function CarregaArrayDePrevia(strConta As String, strLote As String, lngContaBancaria As Long, dblPrincipal As Double, dblMulta As Double, dblJuros As Double, dblCorrecao As Double, intFor As Integer, intNumeroConta As Integer) As Integer
Dim adoResultado    As ADODB.Recordset
Dim strSql          As String

    If strConta = vetPrevia(0, intFor) And strLote = vetPrevia(1, intFor) Then
        vetPrevia(2, intFor) = CCur(gstrConvVrDoSql(vetPrevia(2, intFor), , , True)) + dblPrincipal + dblMulta + dblJuros + dblCorrecao
    Else
        If vetPrevia(0, 0) <> "" Then
            intFor = intFor + 1
            ReDim Preserve vetPrevia(4, intFor)
        End If
        vetPrevia(0, intFor) = strConta
        vetPrevia(1, intFor) = strLote
        vetPrevia(2, intFor) = dblPrincipal + dblMulta + dblJuros + dblCorrecao
        
        'Vamos carregar o campo do array com o valor total do resumo para a verificacao de diferenca
        strSql = ""
        strSql = strSql & " SELECT "
        strSql = strSql & " dblValor dblValor "
        strSql = strSql & " FROM "
        strSql = strSql & gstrResumoBancario & " RB"
        strSql = strSql & " WHERE RB.dtmData = " & gstrConvDtParaSql(txtdtmDataMovimento)
        strSql = strSql & " AND RB.intContaBancaria = " & lngContaBancaria
        strSql = strSql & " AND RB.intLote = " & strLote
        
        If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
            With adoResultado
                If Not .EOF Then
                    vetPrevia(3, intFor) = !dblValor
                End If
            End With
        End If
        
        vetPrevia(4, intFor) = intNumeroConta
        
    End If
    
    CarregaArrayDePrevia = intFor
    
End Function

Private Function strMovimentosBancarios(blnSimulado As Boolean) As String
Dim strSql As String

    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " CR.Pkid as Composicao,"
    strSql = strSql & " LA.Pkid as IntALfa,"
    strSql = strSql & " LA.intUtilizacao,"
    strSql = strSql & " LV.Intparcela as IntParcela,"
    strSql = strSql & " MO.intLancamentoValor ,"
    strSql = strSql & " MO.dblPrincipal,"
    strSql = strSql & " MO.dblMulta,"
    strSql = strSql & " MO.dblJuros,"
    strSql = strSql & " MO.dblCorrecao,"
    strSql = strSql & " MO.dblTarifa,"
    strSql = strSql & " MO.dblCorreto,"
    strSql = strSql & " MO.dtmDtPagamento,"
    strSql = strSql & " MO.dtmDtMovimento, "
    strSql = strSql & " MO.intCodigoBaixa, "
    strSql = strSql & " MO.intContaBancaria, "
    strSql = strSql & " MO.PkidMovBancario, "
    strSql = strSql & " CB.intNumeroConta, "
    strSql = strSql & "LTRIM(RTRIM(" & gstrCONVERT(CDT_VARCHAR, "CB.strConta") & ")) " & strCONCAT & "'-'" & strCONCAT & " LTRIM(RTRIM(CB.strDigitoVerificador)) Conta, "
    strSql = strSql & " MO.Lote Lote "
    
    If bytDBType = SQLServer Then
        strSql = strSql & " FROM         tblLancamentoAlfa LA INNER JOIN "
        strSql = strSql & gstrLancamentoValor & " LV ON LA.PKId = LV.intLancamentoAlfa RIGHT OUTER JOIN "
        strSql = strSql & gstrContaBancaria & " CB RIGHT OUTER JOIN "
        strSql = strSql & " (SELECT     MB.Pkid PkidMovBancario, MB.intLancamentoValor, MB.INTCONTABANCARIA INTCONTABANCARIA, MB.INTLOTE LOTE, MB.dblPrincipal, "
        strSql = strSql & " MB.dblMulta, MB.dblJuros, MB.dblCorrecao, MB.dblCorreto, MB.dblTarifa, MB.dtmDtPagamento, MB.dtmDtMovimento, MB.bitProcessado, MB.intCodigoBaixa "
        strSql = strSql & " FROM          tblMovimentoBancario MB " & strREADPAST
        strSql = strSql & " WHERE  MB.DTMDTMOVIMENTO =" & gstrConvDtParaSql(txtdtmDataMovimento) & ") MO ON CB.PKId = MO.INTCONTABANCARIA ON "
        strSql = strSql & " LV.PKId = MO.intLancamentoValor LEFT OUTER JOIN "
        strSql = strSql & " tblComposicaoDaReceita CR ON LA.INTCOMPOSICAODARECEITA = CR.PKId "
        strSql = strSql & " WHERE MO.dtmDtMovimento = " & gstrConvDtParaSql(txtdtmDataMovimento)
        
        If chk_TodasContas.Value = 0 Then
            strSql = strSql & " AND MO.intContaBancaria = " & Val(dbcintContaBancaria.BoundText)
            If chk_Todos.Value = 0 Then
                strSql = strSql & " AND MO.Lote = " & Val(dbcintLote.BoundText)
            End If
        End If
            
    Else
    
        strSql = strSql & " FROM "
        strSql = strSql & gstrContaBancaria & " CB, "
        strSql = strSql & gstrLancamentoValor & " LV, "
        strSql = strSql & gstrLancamentoAlfa & " LA, "
        strSql = strSql & gstrComposicaoDaReceita & " CR, "
        strSql = strSql & "(SELECT MB.Pkid PkidMovBancario, MB.intLancamentoValor,"
        strSql = strSql & " MB.INTCONTABANCARIA  INTCONTABANCARIA,"
        strSql = strSql & " MB.INTLOTE LOTE,"
        strSql = strSql & " MB.dblPrincipal,"
        strSql = strSql & " MB.dblMulta,"
        strSql = strSql & " MB.dblJuros,"
        strSql = strSql & " MB.dblCorrecao,"
        strSql = strSql & " MB.dblCorreto,"
        strSql = strSql & " MB.dblTarifa,"
        strSql = strSql & " MB.dtmDtPagamento,"
        strSql = strSql & " MB.dtmDtMovimento,"
        strSql = strSql & " MB.bitProcessado,"
        strSql = strSql & " MB.intCodigoBaixa"
        strSql = strSql & " FROM "
        strSql = strSql & gstrMovimentoBancario & " MB " & strREADPAST
        strSql = strSql & " WHERE  MB.DTMDTMOVIMENTO =" & gstrConvDtParaSql(txtdtmDataMovimento) & ") MO"
        strSql = strSql & " WHERE MO.dtmDtMovimento = " & gstrConvDtParaSql(txtdtmDataMovimento) & " AND"
        
        If chk_TodasContas.Value = 0 Then
            strSql = strSql & " MO.intContaBancaria = " & Val(dbcintContaBancaria.BoundText) & " AND"
            If chk_Todos.Value = 0 Then
                strSql = strSql & " MO.Lote = " & Val(dbcintLote.BoundText) & " AND"
            End If
        End If
        
        strSql = strSql & " MO.intContaBancaria " & strOUTJSQLServer & "= CB.Pkid " & strOUTJOracle & " AND"
        strSql = strSql & " LV.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " MO.intLancamentoValor AND "
        strSql = strSql & " LA.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LV.Intlancamentoalfa AND "
        If Not blnSimulado Then
            strSql = strSql & " MO.bitProcessado = 0 AND "
        End If
        strSql = strSql & " CR.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LA.intComposicaoDaReceita "
    
    End If
    
    strSql = strSql & " ORDER BY CB.Intnumeroconta, MO.Lote"
    
    strMovimentosBancarios = strSql
    
End Function
