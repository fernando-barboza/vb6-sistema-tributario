VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioPeriodo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2805
   ClientLeft      =   4710
   ClientTop       =   1410
   ClientWidth     =   4485
   Icon            =   "RelatorioPeriodo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4485
   Begin VB.CheckBox chk_TransfBancaria 
      Caption         =   "Incluir Transferências Bancárias"
      Height          =   225
      Left            =   430
      TabIndex        =   10
      Top             =   5655
      Visible         =   0   'False
      Width           =   2895
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2570
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   4524
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Período "
      TabPicture(0)   =   "RelatorioPeriodo.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chk_detalhado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Agrupamento"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chk_ContaPorFolha"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chk_ChequesEmitidos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "opt_ChequesNaoPagos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "opt_ChequesCancelados"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "opt_ChequesPagos"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.Frame Frame1 
         Caption         =   "Conta Bancária"
         Height          =   1335
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   5295
         Begin VB.CheckBox chk_TodasAsContas 
            Caption         =   "Todas as Contas Bancárias"
            Height          =   255
            Left            =   2880
            TabIndex        =   20
            Top             =   840
            Width           =   2340
         End
         Begin MSDataListLib.DataCombo dbc_intContaBancaria 
            Height          =   315
            Left            =   1680
            TabIndex        =   19
            Top             =   480
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intNumeroBanco 
            Height          =   315
            Left            =   840
            TabIndex        =   18
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   540
            Width           =   420
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   3855
         Begin VB.TextBox txtdtmFinal 
            Height          =   315
            Left            =   2460
            MaxLength       =   10
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtdtmInicial 
            Height          =   315
            Left            =   720
            MaxLength       =   10
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblEmpenhoFinal 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   2070
            TabIndex        =   15
            Top             =   420
            Width           =   240
         End
         Begin VB.Label lblEmpenhoInicial 
            AutoSize        =   -1  'True
            Caption         =   "Início"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   420
            Width           =   405
         End
      End
      Begin VB.OptionButton opt_ChequesPagos 
         Caption         =   "Cheques Pagos"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   4410
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.OptionButton opt_ChequesCancelados 
         Caption         =   "Cheques Cancelados"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   5220
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton opt_ChequesNaoPagos 
         Caption         =   "Cheques Não Pagos"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   4800
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CheckBox chk_ChequesEmitidos 
         Caption         =   "Incluir Cheques Emitidos"
         Height          =   195
         Left            =   900
         TabIndex        =   1
         Top             =   3270
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.CheckBox chk_ContaPorFolha 
         Caption         =   "Uma conta por folha"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   2010
         Width           =   1815
      End
      Begin VB.Frame fra_Agrupamento 
         Caption         =   "Agrupamento"
         Height          =   570
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Visible         =   0   'False
         Width           =   3195
         Begin VB.OptionButton opt_Conta 
            Caption         =   "por Conta"
            Height          =   240
            Left            =   1830
            TabIndex        =   6
            Top             =   270
            Width           =   1050
         End
         Begin VB.OptionButton opt_Banco 
            Caption         =   "por Banco"
            Height          =   240
            Left            =   360
            TabIndex        =   5
            Top             =   270
            Value           =   -1  'True
            Width           =   1050
         End
      End
      Begin VB.CheckBox chk_detalhado 
         Caption         =   "Detalhado"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   1710
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRelatorioPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mstrOpcao               As String
    Dim gstrDataInicioExercicio As String
    Dim intContaBancaria   As Integer
    Dim intNumeroBanco     As Integer

    Dim intCodSeguranca         As Integer
    
Private Function blnPeriodoOk() As Boolean
    
    'Validação Especifica
    If UCase(mstrOpcao) = "CC" Or UCase(mstrOpcao) = "CE" Then
        If Len(Trim(dbc_intContaBancaria.BoundText)) = 0 _
            And chk_TodasAsContas.Value = vbUnchecked Then
            ExibeMensagem "É necessário selecionar a conta"
            If dbc_intContaBancaria.Enabled Then dbc_intContaBancaria.SetFocus
            Exit Function
        End If
    End If
    
    'Validação Geral
    If gblnDataValida(txtdtmInicial) = False Then
        ExibeMensagem "Data inicial incorreta."
        If txtdtmInicial.Enabled Then txtdtmInicial.SetFocus
    ElseIf gblnDataValida(txtdtmFinal) = False Then
        ExibeMensagem "Data final incorreta."
        If txtdtmFinal.Enabled Then txtdtmFinal.SetFocus
    ElseIf CVDate(txtdtmFinal) < CVDate(txtdtmInicial) Then
        ExibeMensagem "Data inicial não poder menor que a data final."
        If txtdtmInicial.Enabled Then txtdtmInicial.SetFocus
    ElseIf UCase(App.EXEName) <> "TRIBUTARIO" Then
        If Val(Right(txtdtmInicial, 4)) <> gintExercicio And mstrOpcao <> "CE" And mstrOpcao <> "CC" Then
            ExibeMensagem "Exercício da Data Inicial está incorreto ."
            If txtdtmInicial.Enabled Then txtdtmInicial.SetFocus
        ElseIf Val(Right(txtdtmFinal, 4)) <> gintExercicio And mstrOpcao <> "CE" And mstrOpcao <> "CC" Then
            ExibeMensagem "Exercício da Data Final está incorreto ."
            If txtdtmFinal.Enabled Then txtdtmFinal.SetFocus
        Else
            blnPeriodoOk = True
        End If
    Else
        blnPeriodoOk = True
    End If
End Function

Private Sub ImprimeSaldoBancario()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    Dim strPeriodo As String

    Dim strSQL  As String
    strSQL = ""
'    strSql = strSql & "sp_BalanceteGeralSaldoBancario "
'    strSql = strSql & "'" & gstrMascaraContaContabil & "', "
'    strSql = strSql & gstrConvDtParaSql(gstrDataInicioExercicio) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmInicial) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmFinal) & ", 1"
    strSQL = strSQL & gstrStoredProcedure("sp_BalanceteGeralSaldoBancario", _
        "'" & gstrMascaraContaContabil & "', " & _
        gstrConvDtParaSql(gstrDataInicioExercicio) & ", " & _
        gstrConvDtParaSql(txtdtmInicial) & ", " & _
        gstrConvDtParaSql(txtdtmFinal) & ", 1", True)
        
    If txtdtmInicial.Text = txtdtmFinal.Text Then
        strPeriodo = "No dia: " & txtdtmInicial.Text
    Else
        strPeriodo = "No periodo de: " & txtdtmInicial.Text & " até " & txtdtmFinal.Text
    End If
        
    ImprimeRelatorio rptSaldoBancario, strSQL, "Saldos Bancários - " & strPeriodo, 60
End Sub

Private Sub ImprimeNotaLancamento()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL                  As String
    strSQL = ""
'    strSql = strSql & "sp_NotaLancContabil "
'    strSql = strSql & gstrConvDtParaSql(txtdtmInicial) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmFinal)
    strSQL = strSQL & gstrStoredProcedure("sp_NotaLancContabil", _
        gstrConvDtParaSql(txtdtmInicial) & ", " & _
        gstrConvDtParaSql(txtdtmFinal), True)
    ImprimeRelatorio rptNotaLancContabil, strSQL, "Nota de Lançamento Contabil"
End Sub

Private Sub ImprimeRelacaoCreditoAnulacao()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
'    strSql = strSql & "sp_RelacaoCreditoAnulacao "
'    strSql = strSql & gstrConvDtParaSql(txtdtmInicial) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmFinal)
    strSQL = strSQL & gstrStoredProcedure("sp_RelacaoCreditoAnulacao", _
        gstrConvDtParaSql(txtdtmInicial) & ", " & _
        gstrConvDtParaSql(txtdtmFinal), True)
    ImprimeRelatorio rptRelacaoCreditoAnulacao, strSQL
End Sub

Private Sub ImprimeDemoMensalDespesaExtra()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
'    strSql = strSql & "sp_DemoMensalDespesaExtra "
'    strSql = strSql & gstrConvDtParaSql(txtdtmInicial) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmFinal)
    strSQL = strSQL & gstrStoredProcedure("sp_DemoMensalDespesaExtra", _
        gstrConvDtParaSql(txtdtmInicial) & ", " & _
        gstrConvDtParaSql(txtdtmFinal), True)
    ImprimeRelatorio rptDemoMensalDespesaExtra, strSQL
End Sub

Private Sub ImprimeBalanceteGeral()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
'    strSql = strSql & "sp_BalanceteGeralSaldoBancario "
'    strSql = strSql & "'" & gstrMascaraContaContabil & "', "
'    strSql = strSql & gstrConvDtParaSql(gstrDataInicioExercicio) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmInicial) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmFinal) & ", 0"
    strSQL = strSQL & gstrStoredProcedure("sp_BalanceteGeralSaldoBancario", _
        "'" & gstrMascaraContaContabil & "', " & _
        gstrConvDtParaSql(gstrDataInicioExercicio) & ", " & _
        gstrConvDtParaSql(txtdtmInicial) & ", " & _
        gstrConvDtParaSql(txtdtmFinal) & ", 0", True)
    ImprimeRelatorio rptBalanceteGeral, strSQL
End Sub

Private Sub DemoReceitaExtra()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
'    strSql = strSql & "sp_DemoMensalReceitaExtra "
'    strSql = strSql & gstrConvDtParaSql(txtdtmInicial) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmFinal)
    strSQL = strSQL & gstrStoredProcedure("sp_DemoMensalReceitaExtra", _
        gstrConvDtParaSql(txtdtmInicial) & ", " & _
        gstrConvDtParaSql(txtdtmFinal), True)
    ImprimeRelatorio rptDemoMensalReceitaExtra, strSQL
End Sub

Private Sub ImprimeLivroDiario()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
'    strSql = "sp_LivroDiario " & _
'             gstrConvDtParaSql(txtdtmInicial) & "," & _
'             gstrConvDtParaSql(txtdtmFinal)
    strSQL = gstrStoredProcedure("sp_LivroDiario", _
             gstrConvDtParaSql(txtdtmInicial) & "," & _
             gstrConvDtParaSql(txtdtmFinal), True)
    rptLivroDiario.Tag = "Período: " & txtdtmInicial & " à " & txtdtmFinal
    ImprimeRelatorio rptLivroDiario, strSQL
End Sub

Private Sub ImprimeFluxodeCaixa()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
'    strSql = strSql & "sp_FluxodeCaixa "
'    strSql = strSql & gintExercicio & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmInicial) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmFinal) & ", 1, 2"
    strSQL = strSQL & gstrStoredProcedure("sp_FluxodeCaixa", _
        gintExercicio & ", " & _
        gstrConvDtParaSql(txtdtmInicial) & ", " & _
        gstrConvDtParaSql(txtdtmFinal) & ", 1, 2", True)
    ImprimeRelatorio rptFluxodeCaixa, strSQL
End Sub

Private Sub VerificaOpcao()
    If blnPeriodoOk Then
        Select Case UCase(mstrOpcao)
        Case "BG"
            ImprimeBalanceteGeral
        Case "SB"
            ImprimeSaldoBancario
        Case "RE"
            DemoReceitaExtra
        Case "FC"
            ImprimeFluxodeCaixa
        Case "LD"
            ImprimeLivroDiario
        Case "NL"
            ImprimeNotaLancamento
        Case "MD" 'Minuta diária
            ImprimeMinutaDiaria
        Case "PR" 'Previsão de Receita e Despesa por Período
            ImprimePrevisaoReceitaDespesa
        Case "CA" 'Crédito e Anulaçõa
            ImprimeRelacaoCreditoAnulacao
        Case "DE" 'Demonstrativo mensal de despesa extra-orçamentária
            ImprimeDemoMensalDespesaExtra
        Case "RR" 'Receita Arrecadada
            ImprimeRelatorioReceitaArrecadada
        Case "RN" 'Receita Anulada
            ImprimeRelatorioReceitaAnulada
        Case "RMF"
            rptResumoMovimentacaoFinanceira.strDataInicial = txtdtmInicial
            rptResumoMovimentacaoFinanceira.strDataFinal = txtdtmFinal
            ImprimeRelatorio rptResumoMovimentacaoFinanceira, IIf(bytDBType = EDatabases.Oracle, "Select 1 from Dual", "Select 1 "), "Resumo da Movimentação Financeira"
        Case "CE"
            ImprimeRelatorioChequesEmitidos ' Usa a mesma rotina que CC
        Case "CC"
            ImprimeRelatorioChequesEmitidos ' Usa a mesma rotina que CE
        End Select
    End If
End Sub

Sub ImprimePrevisaoReceitaDespesa()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Alteração do nome dos objetos do Banco de Dados que tiveram seus nomes
'            truncados.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
'    strSql = strSql & "sp_RelacaoPrevisaoReceitaDespesa "
'    strSql = strSql & gstrConvDtParaSql(txtdtmInicial) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmFinal)
    strSQL = strSQL & gstrStoredProcedure("sp_RelacaoPrevisaoReceitaDespe", _
        gstrConvDtParaSql(txtdtmInicial) & ", " & _
        gstrConvDtParaSql(txtdtmFinal), True)
    ImprimeRelatorio rptRelacaoPrevisaoReceitaDespesaAnulacaoPeriodo, strSQL
End Sub

Sub ImprimeMinutaDiaria()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
'    strSql = strSql & "sp_MinutaDiariaPorPeriodo "
'    strSql = strSql & gstrConvDtParaSql(txtdtmInicial) & ", "
'    strSql = strSql & gstrConvDtParaSql(txtdtmFinal)
    strSQL = strSQL & gstrStoredProcedure("sp_MinutaDiariaPorPeriodo", _
        gstrConvDtParaSql(txtdtmInicial) & ", " & _
        gstrConvDtParaSql(txtdtmFinal), True)
    ImprimeRelatorio rptMinutaDiariaPorPeriodo, strSQL
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
    Case UCase(gstrImprimir)
        VerificaOpcao
    Case UCase(gstrPreencherLista)
        PreencherListaDeOpcoes Me.ActiveControl
    Case UCase(gstrNovo)
        LimpaObjeto
    End Select
End Sub

Private Sub LimpaObjeto()
    txtdtmInicial = ""
    txtdtmFinal = ""
    chk_TodasAsContas.Value = vbUnchecked
    dbc_intContaBancaria.BoundText = ""
    dbc_intNumeroBanco.BoundText = ""
    dbc_intNumeroBanco.SetFocus
End Sub

Public Sub CarregaFormulario(strOpcao As String)
    mstrOpcao = strOpcao
    txtdtmInicial = "02/01/" & gintExercicio
    gstrDataInicioExercicio = "01/01/" & gintExercicio
    txtdtmFinal = Format(gstrDataDoSistema(), "DD/MM/") & gintExercicio
    Select Case UCase(strOpcao)
    Case "BG" 'Balancete geral
    'APRESENTAÇÃO PADRÃO DO FORMULÁRIO
        Me.Caption = "Balancete Geral"
        Me.Caption = "Saldos Bancários"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "SB" 'Saldo bancário
    'APRESENTAÇÃO PADRÃO DO FORMULÁRIO
        Me.Caption = "Saldos Bancários"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "RE" 'Receita extra-orçamentária
    'APRESENTAÇÃO PADRÃO DO FORMULÁRIO
        Me.Caption = "Demonstrativo Mensal da Receita Extra-Orçamentária"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "FC" 'Fluxo de Caixa
        Me.Caption = "Fluxo de Caixa"
         chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "LD" 'Livro diário
        Me.Caption = "Livro Diário"
         fra_Agrupamento.Visible = False
    Case "NL" 'Nota de Lançamento
        Me.Caption = "Nota de Lançamento Contábil"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "MD" 'Minuta diária
    'APRESENTAÇÃO PADRÃO DO FORMULÁRIO
        Me.Caption = "Minuta Diária da Receita Arrecadada"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
        chk_ChequesEmitidos.Value = vbChecked
    Case "PR"
        Me.Caption = "Previsões de Receita e Despesa por Período"
         fra_Agrupamento.Visible = False
    Case "CA" 'Crédito e Anulação
    'APRESENTAÇÃO PADRÃO DO FORMULÁRIO
        Me.Caption = "Relaçoes de Crédito e Anulação por período"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "DE"
    'APRESENTAÇÃO PADRÃO DO FORMULÁRIO
        Me.Caption = "Demonstrativo Mensal da Despesa Extra-Orçamentária"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "RR"
        Me.Caption = "Receita Arrecadada"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        
        Me.Height = 3780
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 3170
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "RN"
    'APRESENTAÇÃO PADRÃO DO FORMULÁRIO
        Me.Caption = "Receita Anulada"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "RMF"
    'APRESENTAÇÃO PADRÃO DO FORMULÁRIO
        Me.Caption = "Resumo do Movimentação Financeira"
        chk_ChequesEmitidos.Visible = False
        
        chk_detalhado.Visible = True
        chk_ContaPorFolha.Visible = True
        chk_detalhado.Value = vbChecked
        fra_Agrupamento.Visible = False
        Me.Height = 3180
        Me.Width = 4575
        Frame1.Visible = False
        tab_3dPasta.Height = 2570
        tab_3dPasta.Width = 4185
        chk_TransfBancaria.Visible = False
        opt_ChequesCancelados.Visible = False
        opt_ChequesNaoPagos.Visible = False
        opt_ChequesPagos.Visible = False
        Frame2.Top = 600
        Frame2.Width = 3855
    Case "CE" 'Cheques Emitidos
        Me.Caption = "Listagem de Cheques Emitidos"
        chk_ContaPorFolha.Visible = False
        chk_detalhado.Visible = False
        chk_TransfBancaria.Value = vbUnchecked
        chk_TransfBancaria.Visible = True
        chk_ChequesEmitidos.Visible = False
        fra_Agrupamento.Visible = False
        tab_3dPasta.Height = 4820
        'tab_3dPasta.Height = 1200 '2400
        Frame1.Top = 600
        Frame2.Top = 2030
        Frame1.Visible = True
        
        Me.opt_ChequesPagos.Visible = True
        Me.opt_ChequesPagos.Left = 360
        Me.opt_ChequesPagos.Top = 3310
        Me.opt_ChequesPagos.Value = True
        Me.opt_ChequesNaoPagos.Visible = True
        Me.opt_ChequesNaoPagos.Left = 360
        Me.opt_ChequesNaoPagos.Top = 3710
        Me.opt_ChequesCancelados.Visible = True
        Me.opt_ChequesCancelados.Left = 360
        Me.opt_ChequesCancelados.Top = 4110
        Me.chk_TransfBancaria.Left = 440
        Me.chk_TransfBancaria.Top = 4510
        'Me.Height = 1800 '3030
        tab_3dPasta.Width = 5695
        Frame2.Width = 5295
        Me.Width = 6095
        Me.Height = 5430
    Case "CC" 'CÓPIA DE Cheques Emitidos
        Me.Caption = "Cópia de Cheques Emitidos"
        Frame1.Top = 600
        Frame2.Top = 2030
        tab_3dPasta.Width = 5695
        Frame2.Width = 5295
         Me.Width = 6095
        Frame1.Visible = True
        
        chk_ChequesEmitidos.Visible = True
        
        chk_ContaPorFolha.Visible = False
        
        chk_ChequesEmitidos.Top = 3230
        chk_ChequesEmitidos.Left = 120
        chk_detalhado.Visible = False
        chk_TransfBancaria.Visible = False
        chk_TransfBancaria.Value = vbChecked
        Me.opt_ChequesPagos.Visible = False
        fra_Agrupamento.Visible = False
        Me.opt_ChequesCancelados.Visible = False
        Me.opt_ChequesNaoPagos.Visible = False
        tab_3dPasta.Height = 3720 '2400
        Me.Height = 4330 '3030
    End Select
    CarregaForm Me
    intCodSeguranca = gintCodSeguranca
    Me.HelpContextID = intCodSeguranca
End Sub

Private Sub chk_TodasAsContas_Click()
    If chk_TodasAsContas.Value = vbUnchecked Then
        TrocaCorObjeto dbc_intContaBancaria, False
        TrocaCorObjeto dbc_intNumeroBanco, False
        If intContaBancaria <> 0 Then
            PreencherListaDeOpcoes dbc_intContaBancaria, intContaBancaria
            PreencherListaDeOpcoes dbc_intNumeroBanco, intNumeroBanco
        End If
    Else
        intContaBancaria = gstrENulo(Val(dbc_intContaBancaria.BoundText))
        intNumeroBanco = gstrENulo(Val(dbc_intNumeroBanco.BoundText))
        dbc_intContaBancaria.BoundText = ""
        dbc_intContaBancaria.Text = ""
        dbc_intNumeroBanco.BoundText = ""
        dbc_intNumeroBanco.Text = ""
        TrocaCorObjeto dbc_intContaBancaria, True
        TrocaCorObjeto dbc_intNumeroBanco, True
    End If
End Sub

Private Sub dbc_intContaBancaria_Click(Area As Integer)
    DropDownDataCombo dbc_intContaBancaria, Me, 0
    Dim adoResultado As ADODB.Recordset
    
    If dbc_intContaBancaria.MatchedWithList Then

        Set gobjBanco = New clsBanco

        If gobjBanco.CriaADO(strQueryContaBancaria, 5, adoResultado) Then
           If Not adoResultado.EOF Then
                If dbc_intNumeroBanco.BoundText = adoResultado!Pkid Then Exit Sub
                LeDaTabelaParaObj "", dbc_intNumeroBanco, strQueryContaBancaria
                dbc_intNumeroBanco.BoundText = adoResultado!Pkid
           End If
        End If
        adoResultado.Close
    Else
        dbc_intNumeroBanco.BoundText = ""
    End If
End Sub

Private Sub dbc_intNumeroBanco_Click(Area As Integer)
   Dim strSQL       As String
   Dim adoResultado As ADODB.Recordset
    
    
    DropDownDataCombo dbc_intNumeroBanco, Me, Area
    If dbc_intNumeroBanco.MatchedWithList Then
    
        Set gobjBanco = New clsBanco
        strSQL = "SELECT PKID, RTRIM(LTRIM(strDescricao)) from " & gstrPlanoConta & " WHERE intcontabancaria = " & dbc_intNumeroBanco.BoundText
        
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
           If Not adoResultado.EOF Then
                If dbc_intContaBancaria.BoundText = adoResultado!Pkid Then Exit Sub
                'LeDaTabelaParaObj "", dbc_intContaBancaria, strSql
                'dbc_intContaBancaria.BoundText = adoResultado!Pkid
                PreencherListaDeOpcoes dbc_intContaBancaria, adoResultado!Pkid
                
           End If
        End If
        adoResultado.Close
    Else
        dbc_intContaBancaria.BoundText = ""
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = intCodSeguranca
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir
End Sub

Private Sub Form_Load()
    If UCase(App.EXEName) = "TRIBUTARIO" Then
        chk_detalhado.Enabled = True
    Else
        chk_detalhado.Enabled = False
    End If
    
    chk_detalhado.Value = 1
    chk_ContaPorFolha.Value = vbUnchecked
    chk_ContaPorFolha.Enabled = False
    intContaBancaria = 0
    intNumeroBanco = 0
    dbc_intContaBancaria.Tag = "SELECT PKID, RTRIM(LTRIM(strDescricao)) strDescricao FROM " & gstrPlanoConta & " PC WHERE PC.Bytdisponibilidadedecaixa = 1; strDescricao"
    dbc_intNumeroBanco.Tag = strQueryContaBancaria(True) & ";intNumeroConta"

End Sub
Private Function strQueryContaBancaria(Optional blntag As Boolean) As String

Dim strSQL As String

    strSQL = "SELECT cb.pkid, cb.intnumeroconta "
    strSQL = strSQL & "FROM " & gstrPlanoConta & " PC, " & gstrContaBancaria & " CB "
    strSQL = strSQL & "WHERE pc.intContaBancaria = cb.Pkid AND pc.bytdisponibilidadedecaixa = 1 "
    If Not blntag Then
        strSQL = strSQL & " AND PC.pkid = " & dbc_intContaBancaria.BoundText
    End If
    strSQL = strSQL & " ORDER BY cb.intnumeroconta "
    
strQueryContaBancaria = strSQL
    
End Function
Private Sub opt_Banco_Click()
chk_ContaPorFolha.Value = vbUnchecked
chk_ContaPorFolha.Enabled = False
End Sub

Private Sub opt_Conta_Click()
    chk_ContaPorFolha.Enabled = True
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtdtmFinal_GotFocus()
    MarcaCampo txtdtmFinal
End Sub

Private Sub txtdtmFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmFinal
End Sub

Private Sub txtdtmFinal_LostFocus()
    txtdtmFinal = gstrDataFormatada(txtdtmFinal)
End Sub

Private Sub txtdtmInicial_GotFocus()
    MarcaCampo txtdtmInicial
End Sub

Private Sub txtdtmInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmInicial
End Sub

Private Sub txtdtmInicial_LostFocus()
    txtdtmInicial = gstrDataFormatada(txtdtmInicial)
End Sub

Private Sub ImprimeRelatorioReceitaArrecadada()
    Dim strSQL As String
    
    If opt_Banco.Value Then
        Screen.MousePointer = vbHourglass
        strSQL = gstrStoredProcedure("sp_ReceitaArrecadada", gstrConvDtParaSql(txtdtmInicial) & "," & _
                 gstrConvDtParaSql(txtdtmFinal) & IIf(chk_detalhado.Value = 1, ",1", ",0"), True)
        ImprimeRelatorio rptReceitaArrecadada, strSQL
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbHourglass
        
        strSQL = strReceitaArrecadadaB(gstrConvDtParaSql(txtdtmInicial), gstrConvDtParaSql(txtdtmFinal), _
                                       gstrConvDtParaSql("01/" & Month(gstrDataFormatada(txtdtmInicial.Text)) & "/" & Year(gstrDataFormatada(txtdtmInicial.Text))), _
                                       gstrConvDtParaSql("01/01/" & Year(gstrDataFormatada(txtdtmInicial.Text))))
        
        ImprimeRelatorio rptReceitaArrecadadaB, strSQL, "Arrecadação das Receitas Orçamentárias - Período: " & gstrDataFormatada(txtdtmInicial) & " à " & gstrDataFormatada(txtdtmFinal)
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Function strReceitaArrecadadaB(strDataInicial As String, strDataFinal As String, strDataSaldoMensal As String, strDataSaldoAnual As String) As String

Dim strSQL As String

    strSQL = ""
   
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "DETALHE.cbPkid, "
    strSQL = strSQL & "DETALHE.intNumeroConta, "
    strSQL = strSQL & "DETALHE.strContaBancaria, "
    strSQL = strSQL & "DETALHE.dtmData, "
    strSQL = strSQL & "DETALHE.dblValorPrevisao, "
    
    If bytDBType = EDatabases.Oracle Then
        strSQL = strSQL & "TO_CHAR(DETALHE.dtmData,'MM') MES, "
    Else
        strSQL = strSQL & "MONTH(DETALHE.dtmData) MES, "
    End If
    
    strSQL = strSQL & gstrISNULL("SUM(DETALHE.dblValorOrcamentario)", "0") & " dblValor, "
    strSQL = strSQL & gstrISNULL("SALDOMES.dblSaldoMensal", "0") & " dblSaldoMensal, "
    strSQL = strSQL & gstrISNULL("SALDOANUAL.dblSaldoAnual", "0") & " dblSaldoAnual, "
    strSQL = strSQL & gstrISNULL("ARRECADACAO.dblArrecadar", "0") & " dblArrecadar, "
    strSQL = strSQL & "'RECEITA ORÇAMENTARIA' strReceita "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "OC.Pkid cbPkid, "
    strSQL = strSQL & "OC.strCodigoOrcamentario Intnumeroconta, "
    strSQL = strSQL & "OC.strDescricao AS strContaBancaria, "
    strSQL = strSQL & "CA.dblValorOrcamentario dblValorOrcamentario, "
    strSQL = strSQL & "AR.dtmdata dtmdata, "
    strSQL = strSQL & "PR.dblValor dblValorPrevisao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA, "
    strSQL = strSQL & gstrCodigoOrcamentario & " OC, "
    strSQL = strSQL & gstrArrecadacaoReceita & " AR, "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CA.intConta = OC.PKid "
    strSQL = strSQL & "AND CA.intArrecadacao = AR.PKId "
    strSQL = strSQL & "AND AR.Intcontacontabil = PC.Pkid "
    strSQL = strSQL & "and pr.intcodigoorcamentario = oc.pkid "
    strSQL = strSQL & "AND CA.bytCancelado = 0 "
    strSQL = strSQL & "AND CA.bytTipo = 0 "
    strSQL = strSQL & "AND CA.Dblvalororcamentario <> 0 "
    strSQL = strSQL & "AND AR.dtmData BETWEEN  " & strDataInicial & " AND " & strDataFinal

    strSQL = strSQL & " UNION ALL "

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "OC.Pkid cbPkid, "
    strSQL = strSQL & "OC.strCodigoOrcamentario Intnumeroconta, "
    strSQL = strSQL & "OC.strDescricao AS strContaBancaria, "
    strSQL = strSQL & "CA.dblValorOrcamentario * -1 dblValorOrcamentario, "
    strSQL = strSQL & "AR.dtmdata dtmdata, "
    strSQL = strSQL & "PR.dblValor dblValorPrevisao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA, "
    strSQL = strSQL & gstrCodigoOrcamentario & " OC, "
    strSQL = strSQL & gstrArrecadacaoReceita & " AR, "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CA.intConta = OC.PKid "
    strSQL = strSQL & "AND CA.intArrecadacao = AR.PKId "
    strSQL = strSQL & "AND AR.Intcontacontabil = PC.Pkid "
    strSQL = strSQL & "and pr.intcodigoorcamentario = oc.pkid "
    strSQL = strSQL & "AND CA.bytCancelado = 1 "
    strSQL = strSQL & "AND CA.bytTipo = 0 "
    strSQL = strSQL & "AND CA.Dblvalororcamentario <> 0 "
    strSQL = strSQL & "AND AR.dtmData BETWEEN  " & strDataInicial & " AND " & strDataFinal & " ) DETALHE, "
    
    strSQL = strSQL & " (SELECT "
    strSQL = strSQL & "cbPkid, "
    strSQL = strSQL & "SUM(dblValorOrcamentario) dblSaldoMensal "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "OC.Pkid cbPkid, "
    strSQL = strSQL & "CA.dblValorOrcamentario dblValorOrcamentario "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA, "
    strSQL = strSQL & gstrCodigoOrcamentario & " OC, "
    strSQL = strSQL & gstrArrecadacaoReceita & " AR, "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrContaBancaria & " CB "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CA.intConta = OC.PKid "
    strSQL = strSQL & "AND CA.intArrecadacao = AR.PKId "
    strSQL = strSQL & "AND AR.Intcontacontabil = PC.Pkid "
    strSQL = strSQL & "AND PC.Intcontabancaria = CB.Pkid "
    strSQL = strSQL & "AND CA.bytCancelado = 0 "
    strSQL = strSQL & "AND CA.bytTipo = 0 "
    strSQL = strSQL & "AND CA.Dblvalororcamentario <> 0 "
    strSQL = strSQL & "AND AR.dtmData >= " & strDataSaldoMensal ' Saldo Mensal
    strSQL = strSQL & " AND AR.dtmData < " & strDataInicial
    
    strSQL = strSQL & " UNION ALL "
    
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "OC.Pkid cbPkid, "
    strSQL = strSQL & "CA.dblValorOrcamentario * -1 dblValorOrcamentario "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA, "
    strSQL = strSQL & gstrCodigoOrcamentario & " OC, "
    strSQL = strSQL & gstrArrecadacaoReceita & " AR, "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrContaBancaria & " CB "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CA.intConta = OC.PKid "
    strSQL = strSQL & "AND CA.intArrecadacao = AR.PKId "
    strSQL = strSQL & "AND AR.Intcontacontabil = PC.Pkid "
    strSQL = strSQL & "AND PC.Intcontabancaria = CB.Pkid "
    strSQL = strSQL & "AND CA.bytCancelado = 1 "
    strSQL = strSQL & "AND CA.bytTipo = 0 "
    strSQL = strSQL & "AND CA.Dblvalororcamentario <> 0 "
    strSQL = strSQL & "AND AR.dtmData >= " & strDataSaldoMensal
    strSQL = strSQL & " AND AR.dtmData < " & strDataInicial & ") TMP "  'Saldo Mensal
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "cbPkid) SALDOMES, "
    
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "cbPkid, "
    strSQL = strSQL & "SUM(dblValorOrcamentario) dblSaldoAnual "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "OC.Pkid cbPkid, "
    strSQL = strSQL & "CA.dblValorOrcamentario dblValorOrcamentario "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA, "
    strSQL = strSQL & gstrCodigoOrcamentario & " OC, "
    strSQL = strSQL & gstrArrecadacaoReceita & " AR, "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrContaBancaria & " CB "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CA.intConta = OC.PKid "
    strSQL = strSQL & "AND CA.intArrecadacao = AR.PKId "
    strSQL = strSQL & "AND AR.Intcontacontabil = PC.Pkid "
    strSQL = strSQL & "AND PC.Intcontabancaria = CB.Pkid "
    strSQL = strSQL & "AND CA.bytCancelado = 0 "
    strSQL = strSQL & "AND CA.bytTipo = 0 "
    strSQL = strSQL & "AND CA.Dblvalororcamentario <> 0 "
    strSQL = strSQL & "AND AR.dtmData >= " & strDataSaldoAnual 'Saldo Anual
    strSQL = strSQL & " AND AR.dtmData < " & strDataInicial
    
    strSQL = strSQL & " UNION ALL "
    
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "OC.Pkid cbPkid, "
    strSQL = strSQL & "CA.dblValorOrcamentario * -1 dblValorOrcamentario "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA, "
    strSQL = strSQL & gstrCodigoOrcamentario & " OC, "
    strSQL = strSQL & gstrArrecadacaoReceita & " AR, "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrContaBancaria & " CB "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CA.intConta = OC.PKid "
    strSQL = strSQL & "AND CA.intArrecadacao = AR.PKId "
    strSQL = strSQL & "AND AR.Intcontacontabil = PC.Pkid "
    strSQL = strSQL & "AND PC.Intcontabancaria = CB.Pkid "
    strSQL = strSQL & "AND CA.bytCancelado = 1 "
    strSQL = strSQL & "AND CA.bytTipo = 0 "
    strSQL = strSQL & "AND CA.Dblvalororcamentario <> 0 "
    strSQL = strSQL & "AND AR.dtmData >=  " & strDataSaldoAnual 'Saldo Anual
    strSQL = strSQL & "AND AR.dtmData < " & strDataInicial & " ) TMP "
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "cbPkid ) SALDOANUAL, "
    
    'MOVIMENTOS DE ARRECADACAO
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "cbPkid, "
    strSQL = strSQL & "SUM(dblValorOrcamentario) dblArrecadar "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "OC.Pkid cbPkid, "
    strSQL = strSQL & "CA.dblValorOrcamentario dblValorOrcamentario "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA, "
    strSQL = strSQL & gstrCodigoOrcamentario & " OC, "
    strSQL = strSQL & gstrArrecadacaoReceita & " AR, "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CA.intConta = OC.PKid "
    strSQL = strSQL & "AND CA.intArrecadacao = AR.PKId "
    strSQL = strSQL & "AND PR.intCodigoOrcamentario = OC.PKId "
    strSQL = strSQL & "AND CA.bytCancelado = 0 "
    strSQL = strSQL & "AND CA.bytTipo = 0 "
    strSQL = strSQL & "AND CA.Dblvalororcamentario <> 0 "
    strSQL = strSQL & "AND AR.dtmData BETWEEN " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND " & strDataFinal
    
    strSQL = strSQL & " UNION ALL "
    
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "OC.Pkid cbPkid, "
    strSQL = strSQL & "CA.dblValorOrcamentario * -1 dblValorOrcamentario "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA, "
    strSQL = strSQL & gstrCodigoOrcamentario & " OC, "
    strSQL = strSQL & gstrArrecadacaoReceita & " AR, "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CA.intConta = OC.PKid "
    strSQL = strSQL & "AND CA.intArrecadacao = AR.PKId "
    strSQL = strSQL & "AND PR.intCodigoOrcamentario = OC.PKId "
    strSQL = strSQL & "AND CA.bytCancelado = 1 "
    strSQL = strSQL & "AND CA.bytTipo = 0 "
    strSQL = strSQL & "AND CA.Dblvalororcamentario <> 0 "
    strSQL = strSQL & "AND AR.dtmData BETWEEN " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND " & strDataFinal & " ) TMP "
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "cbPkid ) ARRECADACAO "

    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "DETALHE.cbPkid " & strOUTJSQLServer & "= SALDOMES.cbPkid " & strOUTJOracle
    strSQL = strSQL & "AND DETALHE.cbPkid " & strOUTJSQLServer & "= SALDOANUAL.cbPkid " & strOUTJOracle
    strSQL = strSQL & "AND DETALHE.cbPkid " & strOUTJSQLServer & "= ARRECADACAO.cbPkid " & strOUTJOracle
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "DETALHE.cbPkid, "
    strSQL = strSQL & "DETALHE.intNumeroConta, "
    strSQL = strSQL & "DETALHE.strContaBancaria, "
    strSQL = strSQL & "DETALHE.dtmData, "
    strSQL = strSQL & "DETALHE.dblvalorPrevisao, "
    
    If bytDBType = EDatabases.Oracle Then
        strSQL = strSQL & " TO_CHAR(DETALHE.dtmData,'MM'), "
    Else
        strSQL = strSQL & " MONTH(DETALHE.dtmData), "
    End If
    
    strSQL = strSQL & "SALDOMES.dblSaldoMensal, "
    strSQL = strSQL & "SALDOANUAL.dblSaldoAnual, "
    strSQL = strSQL & "ARRECADACAO.dblArrecadar "

    
    'RECEITA EXTRA ORÇAMENTARIA
    
    strSQL = strSQL & "UNION ALL "

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "DETALHE.Pkid, "
    strSQL = strSQL & "DETALHE.Strcontacontabil, "
    strSQL = strSQL & "DETALHE.strReceitaExtra, "
    strSQL = strSQL & "DETALHE.dtmData, "
    strSQL = strSQL & "0 dblvalorPrevisao, "
    
    If bytDBType = EDatabases.Oracle Then
        strSQL = strSQL & "TO_CHAR(DETALHE.dtmData,'MM') MES, "
    Else
        strSQL = strSQL & "MONTH(DETALHE.dtmData) MES, "
    End If
    
    strSQL = strSQL & gstrISNULL("SUM(DETALHE.dblValorMes)", " 0") & "  dblValor, "
    strSQL = strSQL & gstrISNULL("SALDOMENSAL.dblSaldoMensal", "0") & "  dblSaldoMensal, "
    strSQL = strSQL & gstrISNULL("SALDOANUAL.dblSaldoAnual", "0") & " dblSaldoAnual, "
    strSQL = strSQL & "0  dblArrecadar, "
    strSQL = strSQL & "'RECEITA EXTRA ORÇAMENTÁRIA' strReceita "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "PC.Pkid, "
    strSQL = strSQL & "PC.Strcontacontabil, "
    strSQL = strSQL & "PC.strDescricao strReceitaExtra, "
    strSQL = strSQL & "PP.dtmData, "
    strSQL = strSQL & "SUM(LC.dblValor) dblValorMes "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrProcessoPagamento & " PP, "
    strSQL = strSQL & gstrLancamentoContabil & " LC "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "PC.Pkid = LC.Intconta "
    strSQL = strSQL & "AND PP.Pkid = LC.intProcesso "
    strSQL = strSQL & "AND PC.Blnextraorcamentaria = 1 "
    strSQL = strSQL & "AND LC.Bytnatureza = 0 "
    strSQL = strSQL & "AND PP.bytNormal = 1 "
    strSQL = strSQL & "AND PP.dtmData BETWEEN " & strDataInicial & " AND " & strDataFinal
    strSQL = strSQL & " GROUP BY "
    strSQL = strSQL & "PC.Pkid, "
    strSQL = strSQL & "PC.Strcontacontabil, "
    strSQL = strSQL & "PC.strDescricao, "
    strSQL = strSQL & "PP.dtmData ) DETALHE, "
    
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "PC.Pkid, "
    strSQL = strSQL & "SUM(LC.dblValor) dblSaldoMensal "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrProcessoPagamento & " PP, "
    strSQL = strSQL & gstrLancamentoContabil & " LC "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "PC.Pkid = LC.Intconta "
    strSQL = strSQL & "AND PP.Pkid = LC.intProcesso "
    strSQL = strSQL & "AND PC.Blnextraorcamentaria = 1 "
    strSQL = strSQL & "AND LC.Bytnatureza = 0 "
    strSQL = strSQL & "AND PP.bytNormal = 1 "
    strSQL = strSQL & "AND PP.dtmData >= " & strDataSaldoMensal 'Saldo Mensal
    strSQL = strSQL & "AND PP.dtmData < " & strDataInicial
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "PC.Pkid  ) SALDOMENSAL, "
    
    strSQL = strSQL & "(SELECT "
    strSQL = strSQL & "PC.Pkid, "
    strSQL = strSQL & "SUM(LC.dblValor) dblSaldoAnual "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrProcessoPagamento & " PP, "
    strSQL = strSQL & gstrLancamentoContabil & " LC "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "PC.Pkid = LC.Intconta "
    strSQL = strSQL & "AND PP.Pkid = LC.intProcesso "
    strSQL = strSQL & "AND PC.Blnextraorcamentaria = 1 "
    strSQL = strSQL & "AND LC.Bytnatureza = 0 "
    strSQL = strSQL & "AND PP.bytNormal = 1 "
    strSQL = strSQL & "AND PP.dtmData >= " & strDataSaldoAnual 'Saldo Anual
    strSQL = strSQL & "AND PP.dtmData < " & strDataInicial
    strSQL = strSQL & " GROUP BY "
    strSQL = strSQL & "PC.Pkid  ) SALDOANUAL "
    
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "DETALHE.Pkid " & strOUTJSQLServer & "= SALDOMENSAL.Pkid " & strOUTJOracle
    strSQL = strSQL & "AND DETALHE.Pkid " & strOUTJSQLServer & "= SALDOANUAL.Pkid " & strOUTJOracle
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "DETALHE.Pkid, "
    strSQL = strSQL & "DETALHE.Strcontacontabil, "
    strSQL = strSQL & "DETALHE.strReceitaExtra, "
    strSQL = strSQL & "DETALHE.dtmData, "
    
    If bytDBType = EDatabases.Oracle Then
        strSQL = strSQL & "TO_CHAR(DETALHE.dtmData,'MM') , "
    Else
        strSQL = strSQL & "MONTH(DETALHE.dtmData), "
    End If
    
    strSQL = strSQL & "dblSaldoMensal, "
    strSQL = strSQL & "dblSaldoAnual "
    
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "strReceita DESC, "
    strSQL = strSQL & "Intnumeroconta, "
    strSQL = strSQL & "dtmdata "

    strReceitaArrecadadaB = strSQL

End Function

Private Sub ImprimeRelatorioReceitaAnulada()

    Dim strSQL As String
    Screen.MousePointer = vbHourglass

    strSQL = gstrStoredProcedure("sp_ReceitaAnulada ", CStr(gintExercicio) _
    & "," & gstrConvDtParaSql(txtdtmInicial) & "," & _
             gstrConvDtParaSql(txtdtmFinal), True)
    
    ImprimeRelatorio rptReceitaAnulada, strSQL
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub ImprimeRelatorioChequesEmitidos()
Dim strLen As String
Dim strTrim As String
Dim strSQL As String

    If bytDBType = EDatabases.Oracle Then
        strLen = "Length"
    Else
        strLen = "Len"
    End If
    
    strTrim = "LTrim(RTrim"
    
        
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT strNome, "
    strSQL = strSQL & " intContaContabil, "
    strSQL = strSQL & " strContaContabil As strConta, "
    strSQL = strSQL & " dtmData, "
    strSQL = strSQL & " strDocumento As strCheque, "
    strSQL = strSQL & " intFicha, "
    strSQL = strSQL & " intBanco, "
    strSQL = strSQL & " strAgencia, "
    strSQL = strSQL & " strBanco, "
    strSQL = strSQL & " strHistorico, "
    strSQL = strSQL & " dblValor, "
    strSQL = strSQL & " ProcessoID, "
    strSQL = strSQL & " strLogin as strUsuario "
    strSQL = strSQL & " FROM "
    If Me.opt_ChequesNaoPagos.Value = True Or Me.opt_ChequesCancelados.Value = True Then
        strSQL = strSQL & "(SELECT CT.strNome, "
        strSQL = strSQL & "CB.PKID intContaContabil, "
        strSQL = strSQL & strTrim & "(CB.strConta)) " & strCONCAT & gstrCASEWHEN(strLen & "(" & strTrim & "(strDigitoVerificador)))", "0, ''", "'-'" & strCONCAT & " strDigitoVerificador") & " strContaContabil, "
        strSQL = strSQL & "CH.strCheque strDocumento, "
        strSQL = strSQL & "CB.intNumeroConta intFicha, "
        strSQL = strSQL & "BC.intBanco, "
        strSQL = strSQL & "AG.strAgencia, "
        strSQL = strSQL & "BC.strDescricao strBanco, "
        strSQL = strSQL & "'*CHE*' strHistorico, "
        strSQL = strSQL & "CH.dblValor, "
        
        'strSql = strSql & "CH.PKID ProcessoID, "
        strSQL = strSQL & "CH.pkid ProcessoID, "
        
        strSQL = strSQL & "CH.dtmData, "
        strSQL = strSQL & "US.strLogin "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrPlanoConta & " PC, "
        strSQL = strSQL & "tblcheque" & " CH, "
        strSQL = strSQL & "tblchequeOP" & " CHOP, "
        strSQL = strSQL & gstrOrdemPagamento & " OP, "
        strSQL = strSQL & gstrContribuinte & " CT, "
        strSQL = strSQL & gstrUsuarios & " US, "
        strSQL = strSQL & gstrContaBancaria & " CB, "
        strSQL = strSQL & gstrBanco & " BC, "
        strSQL = strSQL & gstrAgencia & " AG "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "CH.PKID = CHOP.intCheque AND "
        strSQL = strSQL & "CHOP.intOrdemPagamento = OP.PKID AND "
        strSQL = strSQL & "CH.intContaBancaria = CB.PKID AND "
        strSQL = strSQL & "PC.intContaBancaria = CH.intContaBancaria AND "
        strSQL = strSQL & "CT.PKID = OP.IntContribuinte AND "
        strSQL = strSQL & "US.PKID = CH.Lngcodusr AND "
        strSQL = strSQL & "AG.PKID = CB.intAgencia AND "
        strSQL = strSQL & "BC.Pkid = CB.intBanco AND "
        strSQL = strSQL & strLen & "(" & strTrim & "(CH.strCheque))) > 0 AND "
        If opt_ChequesNaoPagos.Value = True Then
            strSQL = strSQL & "CH.strFlag = 0 AND "
        ElseIf opt_ChequesCancelados.Value = True Then
            strSQL = strSQL & "CH.bytCancelado = 1 AND "
        End If
        strSQL = strSQL & "BC.Pkid = CB.intBanco "
    Else
        
        ' Orçamentário
        strSQL = strSQL & " (SELECT CT.strNome, "
        strSQL = strSQL & "PC.PKID intContaContabil, "
        strSQL = strSQL & strTrim & "(CB.strConta)) " & strCONCAT & gstrCASEWHEN(strLen & "(" & strTrim & "(strDigitoVerificador)))", "0, ''", "'-'" & strCONCAT & " strDigitoVerificador") & " strContaContabil, "
        'strSql = strSql & " PC.strContaContabil, "
        strSQL = strSQL & " LC.strDocumento, "
        strSQL = strSQL & " CB.intNumeroConta intFicha, "
        strSQL = strSQL & " BC.intBanco, "
        strSQL = strSQL & " AG.strAgencia, "
        strSQL = strSQL & " BC.strDescricao strBanco, "
        
        If mstrOpcao = "CC" Then
            strSQL = strSQL & " PP.strHistorico " & strCONCAT & " '; ' " & strCONCAT & gstrISNULL(gstrCONVERT(CDT_VARCHAR, "EP.strCodigo"), "''") & strCONCAT & "'/'" & strCONCAT & gstrISNULL(gstrCONVERT(CDT_VARCHAR, "EP.intExercicio"), "''") & strCONCAT & "'-'" & strCONCAT & gstrISNULL(gstrCONVERT(CDT_VARCHAR, "EP.bitDigito"), "''") & strCONCAT & " '; ' " & strCONCAT & "NF.strNotaFiscal strHistorico, "
        Else
            strSQL = strSQL & " PP.strHistorico strHistorico, "
        End If
        
        strSQL = strSQL & " LC.dblValor, pp.PKID ProcessoID, PP.dtmData, US.strLogin "
        strSQL = strSQL & " From "
        strSQL = strSQL & gstrProcessoPagamento & " PP, "
        strSQL = strSQL & gstrEvento & " EV, "
        strSQL = strSQL & gstrLancamentoContabil & " LC, "
        strSQL = strSQL & gstrPlanoConta & " PC, "
        strSQL = strSQL & gstrPagamentoEstornoEmpenho & " PE, "
        strSQL = strSQL & gstrEmpenho & " EP, "
        strSQL = strSQL & gstrSubempenho & " SE, "
        strSQL = strSQL & gstrContribuinte & " CT, "
        strSQL = strSQL & gstrUsuarios & " US, "
        strSQL = strSQL & gstrContaBancaria & " CB, "
        strSQL = strSQL & gstrBanco & " BC, "
        strSQL = strSQL & gstrAgencia & " AG, "
        strSQL = strSQL & gstrProgramaDeTrabalho & " PT, "
        strSQL = strSQL & gstrSubEmpenhoNF & " NF "
        strSQL = strSQL & " Where "
        strSQL = strSQL & " PP.intEvento = EV.PKID AND "
        strSQL = strSQL & " PC.intContaBancaria = CB.PKID AND "
        strSQL = strSQL & " LC.Intprocesso = PP.PKID AND "
        strSQL = strSQL & " PC.PKID = LC.INTCONTA AND "
        strSQL = strSQL & " PE.IntProcesso = PP.PKID AND "
        strSQL = strSQL & " PE.intParcela = SE.PKID AND "
        strSQL = strSQL & " SE.INTEMPENHO = EP.PKID AND "
        strSQL = strSQL & " CT.PKID = EP.Intcredor AND "
        strSQL = strSQL & " US.PKID = LC.Lngcodusr AND "
        strSQL = strSQL & " AG.PKID = CB.intAgencia AND "
        strSQL = strSQL & " BC.PKID = CB.intBanco AND "
        strSQL = strSQL & strLen & "(" & strTrim & "(LC.strDocumento))) > 0 "
        strSQL = strSQL & " AND PT.PKID = EP.intProgramaTrabalho "
        strSQL = strSQL & " AND SE.PKID " & strOUTJSQLServer & "= NF.intSubEmpenho " & strOUTJOracle & " "
        If chk_TodasAsContas.Value = vbUnchecked Then
            strSQL = strSQL & "CB.Pkid = " & dbc_intNumeroBanco.BoundText & " AND "
        End If
        
        'Cheques Emitidos - Tiago Moreira
        If chk_ChequesEmitidos.Value = vbChecked Then

            rptCopiaChequesEmitidos.bytTipoCopia = 1
            strSQL = strSQL & " UNION ALL "
            strSQL = strSQL & "SELECT CT.strNome, "
            strSQL = strSQL & "PC.PKID intContaContabil, "
            strSQL = strSQL & strTrim & "(CB.strConta)) " & strCONCAT & gstrCASEWHEN(strLen & "(" & strTrim & "(strDigitoVerificador)))", "0, ''", "'-'" & strCONCAT & " strDigitoVerificador") & " strContaContabil, "
            strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "CH.strCheque") & ", "
            strSQL = strSQL & "CB.intNumeroConta intFicha, "
            strSQL = strSQL & "BC.intBanco, "
            strSQL = strSQL & "AG.strAgencia, "
            strSQL = strSQL & "BC.strDescricao strBanco, "
            strSQL = strSQL & "'*CHE*' strHistorico, "
            strSQL = strSQL & "CH.dblValor, "
            strSQL = strSQL & "CH.PKID ProcessoID, "
            strSQL = strSQL & "CH.dtmData, "
            strSQL = strSQL & "US.strLogin "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & "tblPlanoConta PC, "
            strSQL = strSQL & "tblCheque CH, "
            strSQL = strSQL & "tblChequeOP CHOP, "
            strSQL = strSQL & "tblOrdemPagamento OP, "
            strSQL = strSQL & "tblContribuinte CT, "
            strSQL = strSQL & "tblUsuario US, "
            strSQL = strSQL & "tblContaBancaria CB, "
            strSQL = strSQL & "tblBanco BC, "
            strSQL = strSQL & "tblAgencia AG "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "CH.PKID = CHOP.intCheque AND "
            strSQL = strSQL & "CHOP.intOrdemPagamento = OP.PKID AND "
            strSQL = strSQL & "CH.intContaBancaria = CB.PKID AND "
            strSQL = strSQL & "PC.intContaBancaria = CH.intContaBancaria AND "
            strSQL = strSQL & "CT.PKID = OP.IntContribuinte AND "
            strSQL = strSQL & "US.PKID = CH.Lngcodusr AND "
            strSQL = strSQL & "AG.PKID = CB.intAgencia AND "
            strSQL = strSQL & "CH.strFlag = 0 AND " & "CH.bytCancelado = 0 AND "
            strSQL = strSQL & "BC.Pkid = CB.intBanco AND "
            strSQL = strSQL & strLen & "(" & strTrim & "(CH.strCheque))) > 0 "
        End If
        ' Extra-Orçamentário
        strSQL = strSQL & " Union All "
        strSQL = strSQL & " SELECT CT.strNome, "
        strSQL = strSQL & "PC.PKID intContaContabil, "
        strSQL = strSQL & strTrim & "(CB.strConta)) " & strCONCAT & gstrCASEWHEN(strLen & "(" & strTrim & "(strDigitoVerificador)))", "0, ''", "'-'" & strCONCAT & " strDigitoVerificador") & " strContaContabil, "
        'strSql = strSql & " PC.strContaContabil, "
        strSQL = strSQL & " LC.strDocumento , "
        strSQL = strSQL & " CB.intNumeroConta intFicha, "
        strSQL = strSQL & " BC.intBanco, "
        strSQL = strSQL & " AG.strAgencia, "
        strSQL = strSQL & " BC.strDescricao strBanco, "
        strSQL = strSQL & " PP.strHistorico, "
        strSQL = strSQL & " LC.dblValor, pp.pkid ProcessoID, PP.dtmData, US.strLogin "
        strSQL = strSQL & " From "
        strSQL = strSQL & gstrProcessoPagamento & " PP, "
        strSQL = strSQL & gstrEvento & " EV, "
        strSQL = strSQL & gstrLancamentoContabil & " LC, "
        strSQL = strSQL & gstrPlanoConta & " PC, "
        strSQL = strSQL & gstrPagamentoEstornoEmpenho & " PE, "
        strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DE, "
        strSQL = strSQL & gstrContribuinte & " CT, "
        strSQL = strSQL & gstrUsuarios & " US, "
        strSQL = strSQL & gstrContaBancaria & " CB, "
        strSQL = strSQL & gstrBanco & " BC, "
        strSQL = strSQL & gstrAgencia & " AG "
        strSQL = strSQL & " Where "
        strSQL = strSQL & " PP.intEvento = EV.PKID AND "
        strSQL = strSQL & " PC.intContaBancaria = CB.PKID AND "
        strSQL = strSQL & " LC.Intprocesso = PP.PKID AND "
        strSQL = strSQL & " PC.PKID = LC.INTCONTA AND "
        strSQL = strSQL & " PE.IntProcesso = PP.PKID AND "
        strSQL = strSQL & " DE.PKID = PE.Intdespesaextra AND "
        strSQL = strSQL & " CT.PKID = DE.INTCONTRIBUINTE AND "
        strSQL = strSQL & " US.PKID = LC.Lngcodusr AND "
        strSQL = strSQL & " AG.PKID = CB.intAgencia AND "
        strSQL = strSQL & " BC.PKID = CB.intBanco AND "
        If chk_TodasAsContas.Value = vbUnchecked Then
            strSQL = strSQL & "CB.Pkid = " & dbc_intNumeroBanco.BoundText & " AND "
        End If
        strSQL = strSQL & strLen & "(" & strTrim & "(LC.strDocumento))) > 0 "
        
        If Me.chk_TransfBancaria.Value = vbChecked Then
        
            strSQL = strSQL & " Union all "
            
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & " '** Transferência Bancária **' strNome, "
            strSQL = strSQL & "PC.PKID intContaContabil, "
            strSQL = strSQL & strTrim & "(CB.strConta)) " & strCONCAT & gstrCASEWHEN(strLen & "(" & strTrim & "(strDigitoVerificador)))", "0, ''", "'-'" & strCONCAT & " strDigitoVerificador") & " strContaContabil, "
            strSQL = strSQL & " CTB.strCheque, "
            strSQL = strSQL & " CB.intNumeroConta intFicha, "
            strSQL = strSQL & " BC.intBanco, "
            strSQL = strSQL & " AG.strAgencia, "
            strSQL = strSQL & " BC.strDescricao strBanco, "
            strSQL = strSQL & " TB.strHistorico, "
            strSQL = strSQL & " CTB.dblValor, "
            strSQL = strSQL & " -1, "
            strSQL = strSQL & " TB.dtmData, "
            strSQL = strSQL & " US.strLogin "
            strSQL = strSQL & " From "
            strSQL = strSQL & gstrEvento & " EV, "
            strSQL = strSQL & gstrPlanoConta & " PC, "
            strSQL = strSQL & gstrUsuarios & " US, "
            strSQL = strSQL & gstrTransferenciaBancaria & " TB, "
            strSQL = strSQL & gstrContaTransferenciaBancaria & " CTB, "
            strSQL = strSQL & gstrContaBancaria & " CB, "
            strSQL = strSQL & gstrBanco & " BC, "
            strSQL = strSQL & gstrAgencia & " AG "
            strSQL = strSQL & " Where CTB.Inttransferenciabancaria = TB.Pkid "
            strSQL = strSQL & " AND TB.Intevento = EV.PKID "
            strSQL = strSQL & " AND PC.PKID = CTB.intConta "
            strSQL = strSQL & " AND CB.PKID = PC.Intcontabancaria "
            strSQL = strSQL & " AND US.PKID = TB.Lngcodusr AND "
            strSQL = strSQL & " AG.PKID = CB.intAgencia AND "
            strSQL = strSQL & " BC.PKID = CB.intBanco AND "
            strSQL = strSQL & strLen & "(" & strTrim & "(CTB.strCheque))) > 0 "
            
        End If
        
    End If
        
    strSQL = strSQL & " ) Uniao "
    strSQL = strSQL & " WHERE dtmData BETWEEN " & gstrConvDtParaSql(txtdtmInicial)
    strSQL = strSQL & " AND " & gstrConvDtParaSql(txtdtmFinal)
    If Me.opt_ChequesNaoPagos.Value = True Or Me.opt_ChequesCancelados.Value = True Then
        strSQL = strSQL & " GROUP BY strNome,intContaContabil, strContaContabil, dtmData, "
        strSQL = strSQL & "strDocumento,intFicha,  intBanco,  strAgencia,  strBanco, "
        strSQL = strSQL & "strHistorico,  dblValor,  ProcessoID, strlogin "
    End If
    strSQL = strSQL & " ORDER BY intContaContabil,strCheque,ProcessoID "
    
    rptChequesEmitidos.lblRelatorio = "Listagem de Cheques Emitidos entre " & gstrDataFormatada(txtdtmInicial) & _
                                        " e " & gstrDataFormatada(txtdtmFinal)
    
    If mstrOpcao = "CE" Then ImprimeRelatorio rptChequesEmitidos, strSQL
    If mstrOpcao = "CC" Then ImprimeRelatorio rptCopiaChequesEmitidos, strSQL
End Sub
