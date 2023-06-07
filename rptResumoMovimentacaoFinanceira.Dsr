VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptResumoMovimentacaoFinanceira 
   Caption         =   "prjOrcamentario - rptResumoMovimentacaoFinanceira (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptResumoMovimentacaoFinanceira.dsx":0000
End
Attribute VB_Name = "rptResumoMovimentacaoFinanceira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strDataInicial        As String
Public strDataFinal          As String
Dim dblTotalOrcamentario     As Double
Dim dblTotalExtra            As Double
Dim dblTotalDespesaExtra     As Double
Dim dblTotalPagamentos       As Double
Dim dblTotalAdiantamentos    As Double
Dim dblTotalDespesaAnterior  As Double
Dim dblTotalRetiradaBancaria As Double
Dim dblTotalSaldoDisponivel  As Double
 
Private Sub ActiveReport_Activate()
HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
     
    dblTotalOrcamentario = 0
    dblTotalExtra = 0
    dblTotalDespesaExtra = 0
    dblTotalPagamentos = 0
    dblTotalAdiantamentos = 0
    dblTotalDespesaAnterior = 0
    dblTotalRetiradaBancaria = 0
    dblTotalSaldoDisponivel = 0
    
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    
    lblRelatorio = lblRelatorio & " - Período: " & strDataInicial & " a " & strDataFinal
    
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If Tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf Tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf Tool.ID = 16 Then
        Configura_Relatorio Me, True
    End If
End Sub


Private Sub Detail_Format()
   
   ImprimeSubReceitaOrcamentaria
   ImprimeSubReceitaExtraOrcamentaria
   ImprimeSubDepositoBancario
   ImprimeSubTransferenciasBancarias
   ImprimeSubPagamentos
   ImprimeSubAdiantamentos
   ImprimeSubDespesaExtraOrcamentaria
   ImprimeSubRetiradaBancaria
   ImprimeSubSaldoDisponivel
   
   CalculaReceitaAnterior
   
   CalculaSaldoExercicioAnterior
   
   CalculaDespesaAnterior
   
   txtstrTotalReceita = dblTotalOrcamentario + dblTotalExtra
   txtstrTotalReceita = gstrConvVrDoSql(txtstrTotalReceita)
   
   txtdblTotalGeralReceita = CDbl(txtstrTotalReceita) + CDbl(txtstrReceitaAnterior) + CDbl(txtdblSaldoExercicioAnterior)
   txtdblTotalGeralReceita = gstrConvVrDoSql(txtdblTotalGeralReceita)
   
   txtdblTotalDespesa = dblTotalPagamentos + dblTotalDespesaExtra
   txtdblTotalDespesa = gstrConvVrDoSql(txtdblTotalDespesa)
   
   txtdblTotalDespesaAnterior = gstrConvVrDoSql(dblTotalDespesaAnterior)
   
   txtdblTotalGeralDaDespesa = CDbl(txtdblTotalDespesaAnterior) + CDbl(txtdblTotalDespesa) - (dblTotalAdiantamentos)
   
   txtdblTotalGeralDaDespesa = gstrConvVrDoSql(txtdblTotalGeralDaDespesa)
   
   txtdblTotalGeral = CDbl(txtdblTotalGeralDaDespesa) + dblTotalSaldoDisponivel
   
   txtdblTotalGeral = gstrConvVrDoSql(txtdblTotalGeral)
   
   txtdblTotalAdiantamentos = gstrConvVrDoSql(dblTotalAdiantamentos)
   
   BuscaAssinaturas
   
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
Private Function strQueryMovimentoBanco() As String
Dim strSql  As String
            
    
    strSql = strSql & "SELECT * FROM ("
    
    
    strSql = strSql & "SELECT "
    strSql = strSql & "CB.INTNUMEROCONTA,"
    strSql = strSql & "CB.INTTIPOCONTABANCARIA GrupoConta,"
    strSql = strSql & "TC.Strdescricao StrdescricaoConta,"
    strSql = strSql & "CB.STRCONTA,"
    strSql = strSql & "CB.STRDIGITOVERIFICADOR,"
    strSql = strSql & "CB.STRDESCRICAO STRCONTABANCARIA,"
    strSql = strSql & "BA.intBanco,"
    strSql = strSql & "BA.PKID PKIDBanco,"
    strSql = strSql & "BA.STRDESCRICAO STRBANCO,"

    strSql = strSql & gstrISNULL("SUM(VL.CreditoAnterior)", "0") & " CreditoAnterior, "
    strSql = strSql & gstrISNULL("SUM(VL.DebitoAnterior)", "0") & " DebitoAnterior, "
    strSql = strSql & gstrISNULL("SUM(VL.Credito)", "0") & " Credito, "
    strSql = strSql & gstrISNULL("SUM(VL.Debito)", "0") & " Debito, "
    strSql = strSql & gstrCASEWHEN("PCS.blnnaturezadaconta", "1,PCS.DBLValor,0,PCS.DBLValor * (-1)")
    strSql = strSql & " SaldoInicial "
    
    strSql = strSql & "FROM ("
    
    strSql = strSql & "SELECT PC.strContaContabil, 0 CreditoAnterior, 0 DebitoAnterior, LC.dblValor Credito, 0 Debito, 0 Saldo, PP.DTMDATA "
    strSql = strSql & "FROM " & gstrProcessoPagamento & " PP, " & gstrPlanoConta & " PC, " & gstrLancamentoContabil & " LC "
    strSql = strSql & "WHERE  LC.intProcesso = PP.PKID AND LC.intConta = PC.PKID AND PP.bytNormal = 1 AND LC.bytNatureza = 1 AND PC.blnfinanceira = 1 AND PP.dtmData = " & gstrConvDtParaSql(strDataFinal)
    
    strSql = strSql & " UNION ALL "
    
    strSql = strSql & "SELECT PC.strContaContabil, 0 CreditoAnterior, 0 DebitoAnterior, 0 Credito, LC.dblValor Debito, 0 Saldo, PP.DTMDATA "
    strSql = strSql & "FROM " & gstrProcessoPagamento & " PP, " & gstrPlanoConta & " PC, " & gstrLancamentoContabil & " LC "
    strSql = strSql & "WHERE  LC.intProcesso = PP.PKID AND LC.intConta = PC.PKID AND PP.bytNormal = 1 AND LC.bytNatureza = 0 AND PC.blnfinanceira = 1 AND PP.dtmData = " & gstrConvDtParaSql(strDataFinal)
    
    strSql = strSql & " UNION ALL "
    
    strSql = strSql & "SELECT PC.strContaContabil, sum(lc.dblvalor) CreditoAnterior, 0 debitoanterior, 0 Credito, 0 Debito, 0 Saldo, " & gstrConvDtParaSql(strDataFinal) & " dtmdata "
    strSql = strSql & "FROM " & gstrProcessoPagamento & " PP, " & gstrPlanoConta & " PC, " & gstrLancamentoContabil & " LC "
    strSql = strSql & "WHERE  LC.intProcesso = PP.PKID AND LC.intConta = PC.PKID AND PP.bytNormal = 1 AND LC.bytNatureza = 1 AND PC.blnfinanceira = 1 AND "
    strSql = strSql & "(PP.dtmData >= " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND PP.dtmData <= " & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataFinal))) & ") GROUP BY PC.strContaContabil"
    
    strSql = strSql & " UNION ALL "
    
    strSql = strSql & "SELECT PC.strContaContabil, 0 creditoantrior, sum(lc.dblvalor) DebitoAnterior, 0 Credito, 0 Debito, 0 Saldo, " & gstrConvDtParaSql(strDataFinal) & " dtmdata "
    strSql = strSql & "FROM " & gstrProcessoPagamento & " PP, " & gstrPlanoConta & " PC, " & gstrLancamentoContabil & " LC "
    strSql = strSql & "WHERE  LC.intProcesso = PP.PKID AND LC.intConta = PC.PKID AND PP.bytNormal = 1 AND LC.bytNatureza = 0 AND PC.blnfinanceira = 1 AND "
    strSql = strSql & "(PP.dtmData >= " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND PP.dtmData <= " & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataFinal))) & ") GROUP BY PC.strContaContabil"
    
    strSql = strSql & ") VL"
    
    If bytDBType = Oracle Then
        strSql = strSql & "," & gstrPlanoConta & " PC, "
        strSql = strSql & gstrContaBancaria & " CB, "
        strSql = strSql & gstrBanco & " BA, "
        strSql = strSql & gstrTipoContaBancaria & " TC, "
        strSql = strSql & gstrPlanoContaSaldo & " PCS "
    
        strSql = strSql & "WHERE "

        strSql = strSql & "PC.Intcontabancaria " & strOUTJSQLServer & " = CB.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "CB.INTBANCO " & strOUTJSQLServer & " = BA.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "TC.PKID = CB.INTTIPOCONTABANCARIA AND "
        strSql = strSql & "VL.strContaContabil " & strOUTJOracle & " = " & strOUTJSQLServer & " PC.strContaContabil AND "
        strSql = strSql & "PCS.intPlanoConta = PC.PKID AND PCS.intExercicio = " & gintExercicio & " AND "
    ElseIf bytDBType = SQLServer Then
        strSql = strSql & "  RIGHT OUTER JOIN " _
            & gstrPlanoConta & " PC " _
            & " ON (VL.strContaContabil  = PC.strContaContabil) " _
            & " RIGHT OUTER JOIN " _
            & gstrContaBancaria & " CB " _
            & " ON (PC.Intcontabancaria = CB.PKID ) " _
            & " RIGHT OUTER JOIN " _
            & gstrBanco & " BA " _
            & " ON (CB.INTBANCO = BA.PKID ) " _
            & " FULL OUTER JOIN  " _
            & gstrTipoContaBancaria & " TC " _
            & " ON (TC.PKID = CB.INTTIPOCONTABANCARIA), " _
            & gstrPlanoContaSaldo & " PCS " _
            & " Where " _
            & "PCS.intPlanoConta = PC.PKID AND PCS.intExercicio = " & gintExercicio & " AND "
    End If
    
    strSql = strSql & " PC.blnFinanceira = 1 "
    
    strSql = strSql & "GROUP BY CB.INTTIPOCONTABANCARIA,"
      strSql = strSql & "TC.Strdescricao, "
      strSql = strSql & "CB.INTNUMEROCONTA, "
      strSql = strSql & "CB.STRCONTA, "
      strSql = strSql & "CB.STRDIGITOVERIFICADOR, "
      strSql = strSql & "CB.STRDESCRICAO, "
      strSql = strSql & "BA.intBanco, "
      strSql = strSql & "BA.PKID, "
      strSql = strSql & "BA.STRDESCRICAO, "
      strSql = strSql & "PCS.Blnnaturezadaconta, "
      strSql = strSql & "PCS.DBLValor "
    
      If bytDBType = Oracle Then
          strSql = strSql & "ORDER BY CB.INTTIPOCONTABANCARIA, BA.STRDESCRICAO, CB.STRCONTA, CB.STRDESCRICAO"
      End If
      
      strSql = strSql & " ) MV "
      
      strSql = strSql & " ORDER BY MV.GrupoConta,MV.intBanco,MV.intNumeroConta "
      
   strQueryMovimentoBanco = strSql
   
End Function
Sub ImprimeSubReceitaOrcamentaria()
       
       Dim strSql       As String
       Dim adoRelatorio As New ADODB.Recordset
       
      strSql = strSql & gstrStoredProcedure("sp_CompReceitaPrevistaArrecada", _
                        gstrConvDtParaSql(strDataInicial) & "," & gstrConvDtParaSql(strDataFinal) & ", " & _
                        "'" & gstrMascaraCodigoOrcamentario & "','" & CStr(gintExercicio) & "'", True)

       
        Set gobjBanco = New clsBanco
        
        With rptSubRMFReceitaOrcamentaria
        
        If gobjBanco.CriaADO(strSql, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strSql
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubReceitaOrcamentaria.object = rptSubRMFReceitaOrcamentaria
                     
                     
                     Me.SubReceitaOrcamentaria.Visible = True
                     
                     While Not adoRelatorio.EOF
                        If adoRelatorio!bytNivel = 1 Then
                           dblTotalOrcamentario = dblTotalOrcamentario + adoRelatorio!dblValorMes
                        End If
                        adoRelatorio.MoveNext
                     Wend
                     
                     adoRelatorio.MoveFirst
            
            End If
                  
           
        End If
    
    End With

End Sub
Sub ImprimeSubReceitaExtraOrcamentaria()
   Dim strSql As String
   Dim adoRelatorio As New ADODB.Recordset
   
   
   strSql = "SELECT "
   strSql = strSql & "PC.strContaContabil, "
   strSql = strSql & "PC.strDescricao, "
   strSql = strSql & gstrISNULL("SUM(LC.dblValor)", "0") & " dblValor "
   strSql = strSql & "FROM "
   strSql = strSql & gstrPlanoConta & " PC, "
   strSql = strSql & gstrProcessoPagamento & " PP, "
   strSql = strSql & gstrLancamentoContabil & " LC "
   strSql = strSql & "WHERE "
   strSql = strSql & "PP.PKID = LC.IntProcesso AND "
   strSql = strSql & "PC.PKID = LC.intConta AND "
   strSql = strSql & "PC.blnExtraOrcamentaria = 1 AND "
   strSql = strSql & "PP.intLancamentoContabil IS NOT NULL AND "
   strSql = strSql & "LC.bytNatureza = 0 AND "
   strSql = strSql & "PP.dtmData BETWEEN " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal)
   strSql = strSql & " GROUP BY PC.strContaContabil, PC.strDescricao "
   
   Set gobjBanco = New clsBanco
        
        With rptSubRMFReceitaExtra
        
        If gobjBanco.CriaADO(strSql, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strSql
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubReceitaExtraOrcamentaria.object = rptSubRMFReceitaExtra
                     
                     
                     Me.SubReceitaExtraOrcamentaria.Visible = True
            
                     While Not adoRelatorio.EOF
                        dblTotalExtra = dblTotalExtra + adoRelatorio!dblValor
                        adoRelatorio.MoveNext
                     Wend
                     adoRelatorio.MoveFirst
            End If
           
        End If
    
    End With


End Sub

Private Sub ImprimeSubDepositoBancario()

Dim strSql       As String
Dim adoRelatorio As New ADODB.Recordset


    strSql = strSql & "SELECT * FROM ("

    
    strSql = strSql & "SELECT "
    strSql = strSql & "CB.INTNUMEROCONTA,"
    strSql = strSql & "CB.INTTIPOCONTABANCARIA GrupoConta,"
    strSql = strSql & "TC.Strdescricao StrdescricaoConta,"
    strSql = strSql & "CB.STRCONTA,"
    strSql = strSql & "CB.STRDIGITOVERIFICADOR,"
    strSql = strSql & "CB.STRDESCRICAO STRCONTABANCARIA,"
    strSql = strSql & "BA.INTBANCO INTNUMEROBANCO,"
    strSql = strSql & "BA.PKID PKIDBanco,"
    strSql = strSql & "BA.STRDESCRICAO STRBANCO,"

    strSql = strSql & gstrISNULL("SUM(VL.Credito)", "0") & " Credito "
    
    strSql = strSql & "FROM ("
    
    strSql = strSql & "SELECT PC.strContaContabil, 0 CreditoAnterior, 0 DebitoAnterior, LC.dblValor Credito, 0 Debito, 0 Saldo, PP.DTMDATA "
    strSql = strSql & "FROM " & gstrProcessoPagamento & " PP, " & gstrPlanoConta & " PC, " & gstrLancamentoContabil & " LC "
    strSql = strSql & "WHERE  LC.intProcesso = PP.PKID AND PP.intOrigem <> 8 AND PP.intLancamentoContabil IS NOT NULL AND LC.intConta = PC.PKID AND PP.bytNormal = 1 AND LC.bytNatureza = 1 AND PC.blnfinanceira = 1 AND PP.dtmData BETWEEN " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal)
    
    strSql = strSql & ") VL"
    
    If bytDBType = Oracle Then
        strSql = strSql & "," & gstrPlanoConta & " PC, "
        strSql = strSql & gstrContaBancaria & " CB, "
        strSql = strSql & gstrBanco & " BA, "
        strSql = strSql & gstrTipoContaBancaria & " TC "
    
        strSql = strSql & "WHERE "

        strSql = strSql & "PC.Intcontabancaria " & strOUTJSQLServer & " = CB.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "CB.INTBANCO " & strOUTJSQLServer & " = BA.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "TC.PKID = CB.INTTIPOCONTABANCARIA AND "
        strSql = strSql & "VL.strContaContabil " & strOUTJOracle & " = " & strOUTJSQLServer & " PC.strContaContabil AND "
    ElseIf bytDBType = SQLServer Then
        strSql = strSql & "  RIGHT OUTER JOIN " _
            & gstrPlanoConta & " PC " _
            & " ON (VL.strContaContabil  = PC.strContaContabil) " _
            & " RIGHT OUTER JOIN " _
            & gstrContaBancaria & " CB " _
            & " ON (PC.Intcontabancaria = CB.PKID ) " _
            & " RIGHT OUTER JOIN " _
            & gstrBanco & " BA " _
            & " ON (CB.INTBANCO = BA.PKID ) " _
            & " FULL OUTER JOIN  " _
            & gstrTipoContaBancaria & " TC " _
            & " ON (TC.PKID = CB.INTTIPOCONTABANCARIA) " _
            & " Where "
    End If
    
    strSql = strSql & " PC.blnFinanceira = 1 "
    
    strSql = strSql & "GROUP BY CB.INTTIPOCONTABANCARIA,"
      strSql = strSql & "TC.Strdescricao, "
      strSql = strSql & "CB.INTNUMEROCONTA, "
      strSql = strSql & "CB.STRCONTA, "
      strSql = strSql & "CB.STRDIGITOVERIFICADOR, "
      strSql = strSql & "CB.STRDESCRICAO, "
      strSql = strSql & "BA.INTBANCO, "
      strSql = strSql & "BA.PKID, "
      strSql = strSql & "BA.STRDESCRICAO, "
      strSql = strSql & "PC.Blnnaturezadaconta, "
      strSql = strSql & "PC.DBLSALDODACONTA "
    
       
       If bytDBType = Oracle Then
            strSql = strSql & "ORDER BY CB.INTTIPOCONTABANCARIA, BA.INTBANCO, BA.STRDESCRICAO, CB.STRCONTA, CB.STRDESCRICAO"
       End If
       strSql = strSql & " ) MV "
       strSql = strSql & " WHERE NOT ( "
       
       strSql = strSql & "MV.Credito = 0 )"
       
   
       strSql = strSql & " ORDER BY "
       strSql = strSql & "MV.GrupoConta,"
       strSql = strSql & "MV.STRBANCO,"
       strSql = strSql & "MV.STRCONTA,"
       strSql = strSql & "MV.STRCONTABANCARIA"
   
   Set gobjBanco = New clsBanco
        
        With rptSubRMFDepositoBancario
        
        If gobjBanco.CriaADO(strSql, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strSql
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubRMFDepositoBancario.object = rptSubRMFDepositoBancario
                     
                     
                     Me.SubRMFDepositoBancario.Visible = True
            
            End If
           
        End If
    
    End With
 
End Sub
Sub ImprimeSubTransferenciasBancarias()
   Dim strSql As String
   Dim adoRelatorio As New ADODB.Recordset
   
    strSql = strSql & "SELECT * FROM ("

    
    strSql = strSql & "SELECT "
    strSql = strSql & "CB.INTNUMEROCONTA,"
    strSql = strSql & "CB.INTTIPOCONTABANCARIA GrupoConta,"
    strSql = strSql & "TC.Strdescricao StrdescricaoConta,"
    strSql = strSql & "CB.STRCONTA,"
    strSql = strSql & "CB.STRDIGITOVERIFICADOR,"
    strSql = strSql & "CB.STRDESCRICAO STRCONTABANCARIA,"
    strSql = strSql & "BA.PKID PKIDBanco,"
    strSql = strSql & "BA.STRDESCRICAO STRBANCO,"

    strSql = strSql & " VL.Credito, "
    strSql = strSql & " 0 Debito "
    
    strSql = strSql & "FROM ("
    
    strSql = strSql & "SELECT PC.strContaContabil, 0 CreditoAnterior, 0 DebitoAnterior, LC.dblValor Credito, 0 Debito, 0 Saldo, PP.DTMDATA "
    strSql = strSql & "FROM " & gstrProcessoPagamento & " PP, " & gstrPlanoConta & " PC, " & gstrLancamentoContabil & " LC, " & gstrEvento & " EV "
    strSql = strSql & "WHERE  LC.intProcesso = PP.PKID AND PP.intEvento = EV.PKID AND EV.intTipoEvento = 8 AND LC.intConta = PC.PKID AND PP.bytNormal = 1 AND LC.bytNatureza = 1 AND PC.blnfinanceira = 1 AND PP.dtmData BETWEEN " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal)
    
    strSql = strSql & ") VL"
    
    If bytDBType = Oracle Then
        strSql = strSql & "," & gstrPlanoConta & " PC, "
        strSql = strSql & gstrContaBancaria & " CB, "
        strSql = strSql & gstrBanco & " BA, "
        strSql = strSql & gstrTipoContaBancaria & " TC "
    
        strSql = strSql & "WHERE "

        strSql = strSql & "PC.Intcontabancaria " & strOUTJSQLServer & " = CB.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "CB.INTBANCO " & strOUTJSQLServer & " = BA.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "TC.PKID = CB.INTTIPOCONTABANCARIA AND "
        strSql = strSql & "VL.strContaContabil " & strOUTJOracle & " = " & strOUTJSQLServer & " PC.strContaContabil AND "
    ElseIf bytDBType = SQLServer Then
        strSql = strSql & "  RIGHT OUTER JOIN " _
            & gstrPlanoConta & " PC " _
            & " ON (VL.strContaContabil  = PC.strContaContabil) " _
            & " RIGHT OUTER JOIN " _
            & gstrContaBancaria & " CB " _
            & " ON (PC.Intcontabancaria = CB.PKID ) " _
            & " RIGHT OUTER JOIN " _
            & gstrBanco & " BA " _
            & " ON (CB.INTBANCO = BA.PKID ) " _
            & " FULL OUTER JOIN  " _
            & gstrTipoContaBancaria & " TC " _
            & " ON (TC.PKID = CB.INTTIPOCONTABANCARIA) " _
            & " Where "
    End If
    
    strSql = strSql & " PC.blnFinanceira = 1 "
    
'      strSQL = strSQL & "GROUP BY CB.INTTIPOCONTABANCARIA,"
'      strSQL = strSQL & "TC.Strdescricao, "
'      strSQL = strSQL & "CB.INTNUMEROCONTA, "
'      strSQL = strSQL & "CB.STRCONTA, "
'      strSQL = strSQL & "CB.STRDIGITOVERIFICADOR, "
'      strSQL = strSQL & "CB.STRDESCRICAO, "
'      strSQL = strSQL & "BA.PKID, "
'      strSQL = strSQL & "BA.STRDESCRICAO, "
'      strSQL = strSQL & "PC.Blnnaturezadaconta, "
'      strSQL = strSQL & "PC.DBLSALDODACONTA "
    
       
'       If bytDBType = Oracle Then
'            strSQL = strSQL & "ORDER BY CB.INTTIPOCONTABANCARIA, BA.STRDESCRICAO, CB.STRCONTA, CB.STRDESCRICAO"
'       End If
       
    strSql = strSql & " UNION ALL SELECT "
    strSql = strSql & "CB.INTNUMEROCONTA,"
    strSql = strSql & "CB.INTTIPOCONTABANCARIA GrupoConta,"
    strSql = strSql & "TC.Strdescricao StrdescricaoConta,"
    strSql = strSql & "CB.STRCONTA,"
    strSql = strSql & "CB.STRDIGITOVERIFICADOR,"
    strSql = strSql & "CB.STRDESCRICAO STRCONTABANCARIA,"
    strSql = strSql & "BA.PKID PKIDBanco,"
    strSql = strSql & "BA.STRDESCRICAO STRBANCO,"

    strSql = strSql & " 0 Credito, "
    strSql = strSql & " VL.Debito "
    
    strSql = strSql & "FROM ("
    
    strSql = strSql & "SELECT PC.strContaContabil, 0 CreditoAnterior, 0 DebitoAnterior, 0 Credito, LC.dblValor Debito, 0 Saldo, PP.DTMDATA "
    strSql = strSql & "FROM " & gstrProcessoPagamento & " PP, " & gstrPlanoConta & " PC, " & gstrLancamentoContabil & " LC, " & gstrEvento & " EV "
    strSql = strSql & "WHERE  LC.intProcesso = PP.PKID AND PP.intEvento = EV.PKID AND EV.intTipoEvento = 8 AND LC.intConta = PC.PKID AND PP.bytNormal = 1 AND LC.bytNatureza = 0 AND PC.blnfinanceira = 1 AND PP.dtmData BETWEEN " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal)
    
    strSql = strSql & ") VL"
    
    If bytDBType = Oracle Then
        strSql = strSql & "," & gstrPlanoConta & " PC, "
        strSql = strSql & gstrContaBancaria & " CB, "
        strSql = strSql & gstrBanco & " BA, "
        strSql = strSql & gstrTipoContaBancaria & " TC "
    
        strSql = strSql & "WHERE "

        strSql = strSql & "PC.Intcontabancaria " & strOUTJSQLServer & " = CB.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "CB.INTBANCO " & strOUTJSQLServer & " = BA.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "TC.PKID = CB.INTTIPOCONTABANCARIA AND "
        strSql = strSql & "VL.strContaContabil " & strOUTJOracle & " = " & strOUTJSQLServer & " PC.strContaContabil AND "
    ElseIf bytDBType = SQLServer Then
        strSql = strSql & "  RIGHT OUTER JOIN " _
            & gstrPlanoConta & " PC " _
            & " ON (VL.strContaContabil  = PC.strContaContabil) " _
            & " RIGHT OUTER JOIN " _
            & gstrContaBancaria & " CB " _
            & " ON (PC.Intcontabancaria = CB.PKID ) " _
            & " RIGHT OUTER JOIN " _
            & gstrBanco & " BA " _
            & " ON (CB.INTBANCO = BA.PKID ) " _
            & " FULL OUTER JOIN  " _
            & gstrTipoContaBancaria & " TC " _
            & " ON (TC.PKID = CB.INTTIPOCONTABANCARIA) " _
            & " Where "
    End If
    
    strSql = strSql & " PC.blnFinanceira = 1 "
    
'      strSQL = strSQL & "GROUP BY CB.INTTIPOCONTABANCARIA,"
'      strSQL = strSQL & "TC.Strdescricao, "
'      strSQL = strSQL & "CB.INTNUMEROCONTA, "
'      strSQL = strSQL & "CB.STRCONTA, "
'      strSQL = strSQL & "CB.STRDIGITOVERIFICADOR, "
'      strSQL = strSQL & "CB.STRDESCRICAO, "
'      strSQL = strSQL & "BA.PKID, "
'      strSQL = strSQL & "BA.STRDESCRICAO, "
'      strSQL = strSQL & "PC.Blnnaturezadaconta, "
'      strSQL = strSQL & "PC.DBLSALDODACONTA "
    
       
'       If bytDBType = Oracle Then
'            strSQL = strSQL & "ORDER BY CB.INTTIPOCONTABANCARIA, BA.STRDESCRICAO, CB.STRCONTA, CB.STRDESCRICAO"
'       End If
       
       strSql = strSql & " ) MV "
       strSql = strSql & " WHERE NOT ( "
       
       strSql = strSql & "MV.Credito = 0 AND MV.Debito = 0 )"
       
   
       strSql = strSql & " ORDER BY "
       strSql = strSql & "MV.GrupoConta,"
       strSql = strSql & "MV.STRBANCO,"
       strSql = strSql & "MV.STRCONTA,"
       strSql = strSql & "MV.STRCONTABANCARIA"
   
   Set gobjBanco = New clsBanco
        
        With rptSubRMFTransferenciaBancaria
        
        If gobjBanco.CriaADO(strSql, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strSql
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubRMFTransferenciaBancaria.object = rptSubRMFTransferenciaBancaria
                     
                     
                     Me.SubRMFTransferenciaBancaria.Visible = True
            
            End If
           
        End If
    
    End With
 
   
End Sub

Sub ImprimeSubPagamentos()
   Dim strSql As String
   Dim adoRelatorio As New ADODB.Recordset
   
   strSql = "SELECT "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao, "
   strSql = strSql & gstrISNULL("SUM(PEE.dblValor)", "0") & " dblValor "
   strSql = strSql & "FROM "
   strSql = strSql & gstrProgramaDeTrabalho & " PT, "
   strSql = strSql & gstrEmpenho & " EP, "
   strSql = strSql & gstrElementoDespesa & " ED, "
   strSql = strSql & gstrSubempenho & " SE, "
   strSql = strSql & gstrPagamentoEstornoEmpenho & " PEE "
   strSql = strSql & "WHERE "
   strSql = strSql & "EP.PKID = SE.intEmpenho AND "
   strSql = strSql & "PEE.intParcela = SE.PKID AND "
'   strSQL = strSQL & "PEE.intProcesso NOT IN (SELECT PEX.intProcesso "
'   strSQL = strSQL & "FROM tblPagamentoEstornoEmpenho PEX, tblSubEmpenho SEX, tblEmpenho EPX, tblTipoEmpenho TEX "
'   strSQL = strSQL & "WHERE "
'   strSQL = strSQL & "PEX.dblValor < 0 AND "
'   strSQL = strSQL & "SEX.PKID = SE.PKID AND "
'   strSQL = strSQL & "EPX.PKID = SEX.intEmpenho AND "
'   strSQL = strSQL & "PEX.intParcela = SEX.PKID AND "
'   strSQL = strSQL & "EPX.intTipo = TEX.PKID AND "
'   strSQL = strSQL & "TEX.bytAdiantamento IS NOT NULL AND "
'   strSQL = strSQL & "PEX.DTMDATA Between  " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal) & ") AND "
   strSql = strSql & "EP.intProgramaTrabalho = PT.PKID AND "
   strSql = strSql & "PT.intElementoDespesa = ED.PKID AND "
   strSql = strSql & "PT.intExercicio = " & gintExercicio & " AND "
   strSql = strSql & "PEE.dtmData BETWEEN " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal)
   strSql = strSql & " GROUP BY "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao "
   
   
    
   
   
   Set gobjBanco = New clsBanco
        
        With rptSubRMFPagamentos
        
        If gobjBanco.CriaADO(strSql, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strSql
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubRMFPagamentos.object = rptSubRMFPagamentos
                     
                     
                     Me.SubRMFPagamentos.Visible = True
                     
                     While Not adoRelatorio.EOF
                        dblTotalPagamentos = dblTotalPagamentos + adoRelatorio!dblValor
                        adoRelatorio.MoveNext
                     Wend
                     adoRelatorio.MoveFirst
            
            End If
           
        End If
    
    End With
   
End Sub

Sub ImprimeSubDespesaExtraOrcamentaria()
   Dim strSql As String
   Dim adoRelatorio As New ADODB.Recordset
   
   
   strSql = "SELECT "
   strSql = strSql & "PC.strContaContabil, "
   strSql = strSql & "PC.strDescricao, "
   strSql = strSql & gstrISNULL("SUM(LC.dblValor)", "0") & " DBLVALOR "
   strSql = strSql & "FROM "
   strSql = strSql & gstrPlanoConta & " PC, "
   strSql = strSql & gstrProcessoPagamento & " PP, "
   strSql = strSql & gstrLancamentoContabil & " LC "
   strSql = strSql & "WHERE "
   strSql = strSql & "PP.PKID = LC.IntProcesso AND "
   strSql = strSql & "PC.PKID = LC.intConta AND "
   strSql = strSql & "PC.blnExtraOrcamentaria = 1 AND "
   strSql = strSql & "PP.intLancamentoContabil IS NOT NULL AND "
   strSql = strSql & "LC.bytNatureza = 1 AND "
   strSql = strSql & "PP.dtmData BETWEEN " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal)
   strSql = strSql & " GROUP BY PC.strContaContabil, PC.strDescricao "
   strSql = strSql & " UNION ALL "
   strSql = strSql & " SELECT "
   strSql = strSql & "PC.strContaContabil, "
   strSql = strSql & "PC.strDescricao, "
   strSql = strSql & gstrISNULL("SUM(LC.dblValor)", "0") & " DBLVALOR "
   strSql = strSql & "FROM "
   strSql = strSql & gstrPlanoConta & " PC, "
   strSql = strSql & gstrProcessoPagamento & " PP, "
   strSql = strSql & gstrLancamentoContabil & " LC, "
   strSql = strSql & gstrEvento & " EV "
   strSql = strSql & "WHERE "
   strSql = strSql & "PP.PKID = LC.IntProcesso AND "
   strSql = strSql & "PP.IntEvento = EV.PKID AND "
   strSql = strSql & "EV.intTipoEvento = 4 AND "
   strSql = strSql & "PC.PKID = LC.intConta AND "
   strSql = strSql & "PP.intLancamentoContabil IS NOT NULL AND "
   strSql = strSql & "LC.bytNatureza = 1 AND "
   strSql = strSql & "PP.dtmData BETWEEN " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal)
   strSql = strSql & " GROUP BY PC.strContaContabil, PC.strDescricao "
   
   Set gobjBanco = New clsBanco
        
        With rptSubRMFDespesaExtra
        
        If gobjBanco.CriaADO(strSql, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strSql
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubDespesaExtra.object = rptSubRMFDespesaExtra
                     
                     
                     Me.SubDespesaExtra.Visible = True
                     
                     While Not adoRelatorio.EOF
                        dblTotalDespesaExtra = dblTotalDespesaExtra + adoRelatorio!dblValor
                        adoRelatorio.MoveNext
                     Wend
                     adoRelatorio.MoveFirst

            End If
             
        End If
    
    End With


End Sub
Private Sub ImprimeSubRetiradaBancaria()

Dim strSql       As String
Dim adoRelatorio As New ADODB.Recordset


    strSql = strSql & "SELECT * FROM ("

    
    strSql = strSql & "SELECT "
    strSql = strSql & "CB.INTNUMEROCONTA,"
    strSql = strSql & "CB.INTTIPOCONTABANCARIA GrupoConta,"
    strSql = strSql & "TC.Strdescricao StrdescricaoConta,"
    strSql = strSql & "CB.STRCONTA,"
    strSql = strSql & "CB.STRDIGITOVERIFICADOR,"
    strSql = strSql & "CB.STRDESCRICAO STRCONTABANCARIA,"
    strSql = strSql & "BA.PKID PKIDBanco,"
    strSql = strSql & "BA.STRDESCRICAO STRBANCO,"

    strSql = strSql & gstrISNULL("SUM(VL.Debito)", "0") & " Debito "
    
    strSql = strSql & "FROM ("
    
    strSql = strSql & "SELECT PC.strContaContabil, 0 CreditoAnterior, 0 DebitoAnterior, 0 Credito, LC.dblValor Debito, 0 Saldo, PP.DTMDATA "
    strSql = strSql & "FROM " & gstrProcessoPagamento & " PP, " & gstrPlanoConta & " PC, " & gstrLancamentoContabil & " LC, " & gstrEvento & " EV "
    strSql = strSql & "WHERE  LC.intProcesso = PP.PKID AND PP.intEvento = EV.PKID AND EV.intTipoEvento <> 8 AND LC.intConta = PC.PKID AND PP.bytNormal = 0 AND LC.bytNatureza = 0 AND PC.blnfinanceira = 1 AND PP.dtmData BETWEEN " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal)
    
    strSql = strSql & ") VL"
    
    If bytDBType = Oracle Then
        strSql = strSql & "," & gstrPlanoConta & " PC, "
        strSql = strSql & gstrContaBancaria & " CB, "
        strSql = strSql & gstrBanco & " BA, "
        strSql = strSql & gstrTipoContaBancaria & " TC "
    
        strSql = strSql & "WHERE "

        strSql = strSql & "PC.Intcontabancaria " & strOUTJSQLServer & " = CB.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "CB.INTBANCO " & strOUTJSQLServer & " = BA.PKID " & strOUTJOracle & " AND "
        strSql = strSql & "TC.PKID = CB.INTTIPOCONTABANCARIA AND "
        strSql = strSql & "VL.strContaContabil " & strOUTJOracle & " = " & strOUTJSQLServer & " PC.strContaContabil AND "
    ElseIf bytDBType = SQLServer Then
        strSql = strSql & "  RIGHT OUTER JOIN " _
            & gstrPlanoConta & " PC " _
            & " ON (VL.strContaContabil  = PC.strContaContabil) " _
            & " RIGHT OUTER JOIN " _
            & gstrContaBancaria & " CB " _
            & " ON (PC.Intcontabancaria = CB.PKID ) " _
            & " RIGHT OUTER JOIN " _
            & gstrBanco & " BA " _
            & " ON (CB.INTBANCO = BA.PKID ) " _
            & " FULL OUTER JOIN  " _
            & gstrTipoContaBancaria & " TC " _
            & " ON (TC.PKID = CB.INTTIPOCONTABANCARIA) " _
            & " Where "
    End If
    
    strSql = strSql & " PC.blnFinanceira = 1 "
    
    strSql = strSql & "GROUP BY CB.INTTIPOCONTABANCARIA,"
      strSql = strSql & "TC.Strdescricao, "
      strSql = strSql & "CB.INTNUMEROCONTA, "
      strSql = strSql & "CB.STRCONTA, "
      strSql = strSql & "CB.STRDIGITOVERIFICADOR, "
      strSql = strSql & "CB.STRDESCRICAO, "
      strSql = strSql & "BA.PKID, "
      strSql = strSql & "BA.STRDESCRICAO, "
      strSql = strSql & "PC.Blnnaturezadaconta, "
      strSql = strSql & "PC.DBLSALDODACONTA "
    
       
       If bytDBType = Oracle Then
            strSql = strSql & "ORDER BY CB.INTTIPOCONTABANCARIA, BA.STRDESCRICAO, CB.STRCONTA, CB.STRDESCRICAO"
       End If
       strSql = strSql & " ) MV "
       strSql = strSql & " WHERE NOT ( "
       
       strSql = strSql & "MV.Debito = 0 )"
       
   
       strSql = strSql & " ORDER BY "
       strSql = strSql & "MV.GrupoConta,"
       strSql = strSql & "MV.STRBANCO,"
       strSql = strSql & "MV.STRCONTA,"
       strSql = strSql & "MV.STRCONTABANCARIA"
   
   Set gobjBanco = New clsBanco
        
        With rptSubRMFRetiradaBancaria
        
        If gobjBanco.CriaADO(strSql, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strSql
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubRMFRetiradaBancaria.object = rptSubRMFRetiradaBancaria
                     
                     
                     Me.SubRMFRetiradaBancaria.Visible = True
                     
                     While Not adoRelatorio.EOF
                        dblTotalRetiradaBancaria = dblTotalRetiradaBancaria + adoRelatorio!Debito
                        adoRelatorio.MoveNext
                     Wend
                     adoRelatorio.MoveFirst
            
            End If
           
        End If
    
    End With
 
End Sub
Private Sub ImprimeSubSaldoDisponivel()

Dim strSql       As String
Dim adoRelatorio As New ADODB.Recordset
   
   strSql = strQueryMovimentoBanco
   
   Set gobjBanco = New clsBanco
        
        With rptSubRMFSaldoDisponivel
        
        If gobjBanco.CriaADO(strSql, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strSql
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubRMFSaldosBancarios.object = rptSubRMFSaldoDisponivel
                     
                     
                     Me.SubRMFSaldosBancarios.Visible = True
                     
                     While Not adoRelatorio.EOF
                        dblTotalSaldoDisponivel = dblTotalSaldoDisponivel + (adoRelatorio!SaldoInicial + ((adoRelatorio!CreditoAnterior + adoRelatorio!Credito) - (adoRelatorio!DebitoAnterior + adoRelatorio!Debito)))
                        adoRelatorio.MoveNext
                     Wend
                     adoRelatorio.MoveFirst
            End If
           
        End If
    
    End With
 
End Sub

Sub ImprimeSubAdiantamentos()
   Dim strSql As String
   Dim adoRelatorio As New ADODB.Recordset
   
   strSql = "SELECT "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao, "
   strSql = strSql & gstrISNULL("SUM(AD.dblValor)", "0") & " dblValor "
   strSql = strSql & "FROM "
   strSql = strSql & gstrProgramaDeTrabalho & " PT, "
   strSql = strSql & gstrEmpenho & " EP, "
   strSql = strSql & gstrElementoDespesa & " ED, "
   strSql = strSql & gstrSubempenho & " SE, "
   strSql = strSql & gstrTipoEmpenho & " TE, "
   strSql = strSql & gstrAnulacaoDespesa & " AD "
   strSql = strSql & "WHERE "
   strSql = strSql & "EP.PKID = SE.intEmpenho AND "
   strSql = strSql & "SE.PKID = AD.intParcela AND "
   strSql = strSql & "EP.intProgramaTrabalho = PT.PKID AND "
   strSql = strSql & "EP.intTipo = TE.PKID AND "
   strSql = strSql & "TE.bytAdiantamento = 1 AND "
   strSql = strSql & "PT.intElementoDespesa = ED.PKID AND "
   strSql = strSql & "PT.intExercicio = " & gintExercicio & " AND "
   strSql = strSql & "AD.dtmData BETWEEN " & gstrConvDtParaSql(strDataInicial) & " AND " & gstrConvDtParaSql(strDataFinal)
   strSql = strSql & " GROUP BY "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao "
   
   
   Set gobjBanco = New clsBanco
        
        With rptSubRMFAdiantamentos
        
        If gobjBanco.CriaADO(strSql, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strSql
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubRMFAdiantamentos.object = rptSubRMFAdiantamentos
                     
                     
                     Me.SubRMFAdiantamentos.Visible = True
                     
                     While Not adoRelatorio.EOF
                        dblTotalAdiantamentos = dblTotalAdiantamentos + adoRelatorio!dblValor
                        adoRelatorio.MoveNext
                     Wend
                     adoRelatorio.MoveFirst
            
            End If
           
        End If
    
    End With
   
End Sub

Sub CalculaReceitaAnterior()
   Dim strSql             As String
   Dim adoRelatorio       As New ADODB.Recordset
   Dim dblReceitaAnterior As Double
   
   strSql = strSql & gstrStoredProcedure("sp_CompReceitaPrevistaArrecada", _
                     gstrConvDtParaSql("01/01/" & gintExercicio) & "," & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataInicial))) & ", " & _
                     "'" & gstrMascaraCodigoOrcamentario & "','" & CStr(gintExercicio) & "'", True)

   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoRelatorio) Then
      
      With adoRelatorio
               
         While Not .EOF
            If adoRelatorio!bytNivel = 1 Then
               dblReceitaAnterior = dblReceitaAnterior + adoRelatorio!dblValorMes
            End If
            .MoveNext
         Wend
      End With
   End If
   
   
   strSql = "SELECT "
   strSql = strSql & "PC.strContaContabil, "
   strSql = strSql & "PC.strDescricao, "
   strSql = strSql & "LC.dblValor "
   strSql = strSql & "FROM "
   strSql = strSql & gstrPlanoConta & " PC, "
   strSql = strSql & gstrProcessoPagamento & " PP, "
   strSql = strSql & gstrLancamentoContabil & " LC "
   strSql = strSql & "WHERE "
   strSql = strSql & "PP.PKID = LC.IntProcesso AND "
   strSql = strSql & "PC.PKID = LC.intConta AND "
   strSql = strSql & "PC.blnExtraOrcamentaria = 1 AND "
   strSql = strSql & "PP.intLancamentoContabil IS NOT NULL AND "
   strSql = strSql & "LC.bytNatureza = 0 AND "
   strSql = strSql & "PP.dtmData BETWEEN " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND " & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataInicial)))
   
   If gobjBanco.CriaADO(strSql, 10, adoRelatorio) Then
      With adoRelatorio
         While Not .EOF
           dblReceitaAnterior = dblReceitaAnterior + !dblValor
           .MoveNext
         Wend
      End With
   End If
   
   txtstrReceitaAnterior = dblReceitaAnterior
   txtstrReceitaAnterior = gstrConvVrDoSql(txtstrReceitaAnterior)
   
End Sub
Sub CalculaSaldoExercicioAnterior()
   Dim strSql As String
   Dim adoRelatorio As ADODB.Recordset
   
'   strSql = "SELECT " & gstrISNULL("SUM(dblSaldoDaConta)", "0") & " dblValor"
'   strSql = strSql & " FROM " & gstrPlanoConta
'   strSql = strSql & " WHERE bytDisponibilidadeDeCaixa = 1"
   
   strSql = "SELECT " & gstrISNULL("SUM(PCS.dblValor)", "0") & " dblValor"
   strSql = strSql & " FROM " & gstrPlanoConta & " PC, " & gstrPlanoContaSaldo & " PCS "
   strSql = strSql & " WHERE PC.bytDisponibilidadeDeCaixa = 1 and PC.Pkid = PCS.intPlanoConta and PCS.intExercicio = " & gintExercicio
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoRelatorio) Then
      txtdblSaldoExercicioAnterior = gstrConvVrDoSql(adoRelatorio!dblValor)
   End If
   
End Sub
Sub CalculaDespesaAnterior()
   Dim strSql As String
   Dim adoRelatorio As New ADODB.Recordset
   
   strSql = "SELECT DA.strCodigoElementoDespesa, DA.strDescricao, " & gstrISNULL("SUM(DA.dblValor)", "0") & " dblValor FROM ( SELECT "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao, "
   strSql = strSql & gstrISNULL("SUM(PEE.dblValor)", "0") & " dblValor "
   strSql = strSql & "FROM "
   strSql = strSql & gstrProgramaDeTrabalho & " PT, "
   strSql = strSql & gstrEmpenho & " EP, "
   strSql = strSql & gstrElementoDespesa & " ED, "
   strSql = strSql & gstrSubempenho & " SE, "
   strSql = strSql & gstrPagamentoEstornoEmpenho & " PEE "
   strSql = strSql & "WHERE "
   strSql = strSql & "EP.PKID = SE.intEmpenho AND "
   strSql = strSql & "PEE.intParcela = SE.PKID AND "
'   strsql = strsql & "PEE.intProcesso NOT IN (SELECT PEX.intProcesso "
'   strsql = strsql & "FROM tblPagamentoEstornoEmpenho PEX, tblSubEmpenho SEX, tblEmpenho EPX, tblTipoEmpenho TEX "
'   strsql = strsql & "WHERE "
'   strsql = strsql & "PEX.dblValor < 0 AND "
'   strsql = strsql & "SEX.PKID = SE.PKID AND "
'   strsql = strsql & "EPX.PKID = SEX.intEmpenho AND "
'   strsql = strsql & "PEX.intParcela = SEX.PKID AND "
'   strsql = strsql & "EPX.intTipo = TEX.PKID AND "
'   strsql = strsql & "TEX.bytAdiantamento IS NOT NULL AND "
'   strsql = strsql & "PEX.DTMDATA BETWEEN" & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND " & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataFinal))) & ") AND "
   strSql = strSql & "EP.intProgramaTrabalho = PT.PKID AND "
   strSql = strSql & "PT.intElementoDespesa = ED.PKID AND "
   strSql = strSql & "PT.intExercicio = " & gintExercicio & " AND "
   strSql = strSql & "PEE.dtmData BETWEEN " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND " & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataInicial)))
   strSql = strSql & " GROUP BY "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao "
   
   strSql = strSql & "UNION ALL SELECT "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao, "
   strSql = strSql & gstrISNULL("SUM(AD.dblValor)", "0") & " * (-1) dblValor "
   strSql = strSql & "FROM "
   strSql = strSql & gstrProgramaDeTrabalho & " PT, "
   strSql = strSql & gstrEmpenho & " EP, "
   strSql = strSql & gstrElementoDespesa & " ED, "
   strSql = strSql & gstrSubempenho & " SE, "
   strSql = strSql & gstrTipoEmpenho & " TE, "
   strSql = strSql & gstrAnulacaoDespesa & " AD "
   strSql = strSql & "WHERE "
   strSql = strSql & "EP.PKID = SE.intEmpenho AND "
   strSql = strSql & "SE.PKID = AD.intParcela AND "
   strSql = strSql & "EP.intProgramaTrabalho = PT.PKID AND "
   strSql = strSql & "EP.intTipo = TE.PKID AND "
   strSql = strSql & "TE.bytAdiantamento = 1 AND "
   strSql = strSql & "PT.intElementoDespesa = ED.PKID AND "
   strSql = strSql & "PT.intExercicio = " & gintExercicio & " AND "
   strSql = strSql & "AD.dtmData BETWEEN " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND " & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataInicial)))
   strSql = strSql & " GROUP BY "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao ) DA "
   strSql = strSql & " GROUP BY DA.strCodigoElementoDespesa, DA.strDescricao "
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoRelatorio) Then
         If adoRelatorio.RecordCount > 0 Then
            While Not adoRelatorio.EOF
               dblTotalDespesaAnterior = dblTotalDespesaAnterior + (adoRelatorio!dblValor)
               adoRelatorio.MoveNext
            Wend
            
         End If
   End If

   strSql = "SELECT "
   strSql = strSql & "PC.strContaContabil, "
   strSql = strSql & "PC.strDescricao, "
   strSql = strSql & "LC.dblValor "
   strSql = strSql & "FROM "
   strSql = strSql & gstrPlanoConta & " PC, "
   strSql = strSql & gstrProcessoPagamento & " PP, "
   strSql = strSql & gstrLancamentoContabil & " LC "
   strSql = strSql & "WHERE "
   strSql = strSql & "PP.PKID = LC.IntProcesso AND "
   strSql = strSql & "PC.PKID = LC.intConta AND "
   strSql = strSql & "PC.blnExtraOrcamentaria = 1 AND "
   strSql = strSql & "PP.intLancamentoContabil IS NOT NULL AND "
   strSql = strSql & "LC.bytNatureza = 1 AND "
   strSql = strSql & "PP.dtmData BETWEEN " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND " & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataInicial)))
   strSql = strSql & " UNION ALL "
   strSql = strSql & " SELECT "
   strSql = strSql & "PC.strContaContabil, "
   strSql = strSql & "PC.strDescricao, "
   strSql = strSql & "LC.dblValor "
   strSql = strSql & "FROM "
   strSql = strSql & gstrPlanoConta & " PC, "
   strSql = strSql & gstrProcessoPagamento & " PP, "
   strSql = strSql & gstrLancamentoContabil & " LC, "
   strSql = strSql & gstrEvento & " EV "
   strSql = strSql & "WHERE "
   strSql = strSql & "PP.PKID = LC.IntProcesso AND "
   strSql = strSql & "PP.IntEvento = EV.PKID AND "
   strSql = strSql & "EV.intTipoEvento = 4 AND "
   strSql = strSql & "PC.PKID = LC.intConta AND "
   strSql = strSql & "PP.intLancamentoContabil IS NOT NULL AND "
   strSql = strSql & "LC.bytNatureza = 1 AND "
   strSql = strSql & "PP.dtmData BETWEEN " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND " & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataInicial)))
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoRelatorio) Then
         If adoRelatorio.RecordCount > 0 Then
            While Not adoRelatorio.EOF
               dblTotalDespesaAnterior = dblTotalDespesaAnterior + (adoRelatorio!dblValor)
               adoRelatorio.MoveNext
            Wend
            
         End If
   End If
   
   strSql = "SELECT "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao, "
   strSql = strSql & gstrISNULL("SUM(PEE.dblValor)", "0") & " *(-1) dblValor "
   strSql = strSql & "FROM "
   strSql = strSql & gstrProgramaDeTrabalho & " PT, "
   strSql = strSql & gstrEmpenho & " EP, "
   strSql = strSql & gstrElementoDespesa & " ED, "
   strSql = strSql & gstrSubempenho & " SE, "
   strSql = strSql & gstrTipoEmpenho & " TE, "
   strSql = strSql & gstrPagamentoEstornoEmpenho & " PEE "
   strSql = strSql & "WHERE "
   strSql = strSql & "EP.PKID = SE.intEmpenho AND "
   strSql = strSql & "SE.PKID = PEE.intParcela AND "
   strSql = strSql & "PEE.intProcesso IN (SELECT PE.intProcesso FROM " & gstrPagamentoEstornoEmpenho & " PE WHERE PE.dblValor < 0 AND PE.intParcela = SE.PKID) AND  "
   strSql = strSql & "EP.intProgramaTrabalho = PT.PKID AND "
   strSql = strSql & "EP.intTipo = TE.PKID AND "
   strSql = strSql & "TE.bytAdiantamento = 1 AND "
   strSql = strSql & "PT.intElementoDespesa = ED.PKID AND "
   strSql = strSql & "PT.intExercicio = " & gintExercicio & " AND "
   strSql = strSql & "PEE.dtmData BETWEEN " & gstrConvDtParaSql("01/01/" & gintExercicio) & " AND " & gstrConvDtParaSql(DateAdd("d", -1, CDate(strDataInicial)))
   strSql = strSql & " GROUP BY "
   strSql = strSql & "ED.strCodigoElementoDespesa, "
   strSql = strSql & "ED.strDescricao "
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoRelatorio) Then
         If adoRelatorio.RecordCount > 0 Then
            While Not adoRelatorio.EOF
               dblTotalDespesaAnterior = dblTotalDespesaAnterior - (adoRelatorio!dblValor)
               adoRelatorio.MoveNext
            Wend
            
         End If
   End If
   
End Sub
Sub BuscaAssinaturas()
   Dim strSql As String
   Dim adoResultado As ADODB.Recordset
   
   strSql = "SELECT ASS.strNome ,"
   strSql = strSql & "ASS.strCargo ,"
   strSql = strSql & "ASS.strDocumento "
   strSql = strSql & "FROM " & gstrAssinaturas & " ASS "
   strSql = strSql & "WHERE ASS.strCodigo = '1' "
   
   Set gobjBanco = New clsBanco
      
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      
      With adoResultado
         txtNome1 = gstrENulo(!STRNOME)
         txtCargo1 = gstrENulo(!strCargo)
         txtDocumento1 = gstrENulo(!strDocumento)
      End With
      
   End If
   
   strSql = "SELECT ASS.strNome ,"
   strSql = strSql & "ASS.strCargo ,"
   strSql = strSql & "ASS.strDocumento "
   strSql = strSql & "FROM " & gstrAssinaturas & " ASS "
   strSql = strSql & "WHERE ASS.strCodigo = '2' "
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      
      With adoResultado
         txtNome2 = gstrENulo(!STRNOME)
         txtCargo2 = gstrENulo(!strCargo)
         txtDocumento2 = gstrENulo(!strDocumento)
      End With
      
   End If
   
   strSql = "SELECT ASS.strNome ,"
   strSql = strSql & "ASS.strCargo ,"
   strSql = strSql & "ASS.strDocumento "
   strSql = strSql & "FROM " & gstrAssinaturas & " ASS "
   strSql = strSql & "WHERE ASS.strCodigo = '3' "
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      
      With adoResultado
         txtNome3 = gstrENulo(!STRNOME)
         txtCargo3 = gstrENulo(!strCargo)
         txtDocumento3 = gstrENulo(!strDocumento)
      End With
      
   End If
   
   strSql = "SELECT ASS.strNome ,"
   strSql = strSql & "ASS.strCargo ,"
   strSql = strSql & "ASS.strDocumento "
   strSql = strSql & "FROM " & gstrAssinaturas & " ASS "
   strSql = strSql & "WHERE ASS.strCodigo = '4' "
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      
      With adoResultado
         txtNome4 = gstrENulo(!STRNOME)
         txtCargo4 = gstrENulo(!strCargo)
         txtDocumento4 = gstrENulo(!strDocumento)
      End With
      
   End If
   
End Sub
