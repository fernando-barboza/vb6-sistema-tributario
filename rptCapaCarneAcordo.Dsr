VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCapaCarneAcordo 
   Caption         =   "Tributario - rptCapaCarneAcordo (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptCapaCarneAcordo.dsx":0000
End
Attribute VB_Name = "rptCapaCarneAcordo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Parcelas selecionadas vindo do form
Public strParcelasSelecionadas  As String
Public blnSemIndexador          As Boolean
Public blnParcelasAtualizadas   As Boolean
Public intExercicioAtualizadas  As Integer

Private Sub ActiveReport_Activate()
    If adoDataControl.Recordset.RecordCount = 0 Then
       ExibeMensagem "Não existem registros com os dados informados."
       Unload Me
       Exit Sub
    End If
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_ReportStart()
    PadronizaToolBarRelatorio Me
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia
    LeImagemLogotipo imgBrasao1, imgLogotipo, txtNomeFantasia1
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
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

Private Sub GroupHeader1_Format()
Dim adoRelatorio As ADODB.Recordset
Dim adoInscricoes As ADODB.Recordset
Dim intCont As Integer
Dim lngContaBancaria As Long
Dim strsql           As String
Dim StrIndexFormatado As String 'Usado para Colocar o Idexador entre Parenteses

  blnSemIndexador = False
  If IsNull(adoDataControl.Recordset!dblvlIndexador) Or IsNull(adoDataControl.Recordset!Strindexador) Then
     blnSemIndexador = True
  ElseIf Val(adoDataControl.Recordset!dblvlIndexador) = 0 Then
      blnSemIndexador = True
  End If
  
  lblTituloIndexador.Visible = Not blnSemIndexador
  dblvlIndexador.Visible = Not blnSemIndexador
  lblTituloQtde.Visible = Not blnSemIndexador
  txtQuantidadeFmp.Visible = Not blnSemIndexador
  txtTotalFmp.Visible = Not blnSemIndexador
  lblTituloTotal.Caption = IIf(blnSemIndexador, "Total R$", "Total R$/")
  
  Set gobjBanco = New clsBanco
  intCont = 0
  
  txtAcordo.Text = gstrFormataInscricao(txtAcordo.Text & txtExercicio.Text, TYP_ACORDO)
  txtAcordo1.Text = txtAcordo.Text
  
  'VALOR DO FMP
  If Trim(txtTotalFmp.Text) <> "" And (Len(txtTotalFmp.Text) - InStr(txtTotalFmp.Text, ",")) >= 4 Then
     txtTotalFmp.Text = Left(txtTotalFmp.Text, InStr(txtTotalFmp.Text, ",") + 4)
  Else
     txtTotalFmp.Text = Format$(txtTotalFmp.Text, "#,##0.0000")
  End If
  
  If Trim(txtQuantidadeFmp.Text) <> "" And (Len(txtQuantidadeFmp.Text) - InStr(txtQuantidadeFmp.Text, ",")) >= 4 Then
     txtQuantidadeFmp.Text = Left(txtQuantidadeFmp.Text, InStr(txtQuantidadeFmp.Text, ",") + 4)
  Else
     txtQuantidadeFmp.Text = Format$(txtQuantidadeFmp.Text, "#,##0.0000")
  End If
  
  If gobjBanco.CriaADO(strQueryInscricoes, 25, adoInscricoes) Then
     With adoInscricoes
     Do While Not .EOF
        Select Case intCont
          Case 0
            lblInscricao1.Visible = True
            lblComposicao1.Visible = True
            lblExercicio1.Visible = True
            txtInscricao1.Text = IIf(!intUtilizacao <> 0, gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao), !strInscricao)
            txtComposicao1.Text = Trim(!strComposicao)
            txtExercicio1.Text = !intExercicio
            txtInscMun.Text = txtInscricao1.Text
            txtInscMun1.Text = txtInscricao1.Text
          Case 1
            lblInscricao2.Visible = True
            lblComposicao2.Visible = True
            lblExercicio2.Visible = True
            Linha1.Visible = True
            txtInscricao2.Text = IIf(!intUtilizacao <> 0, gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao), !strInscricao)
            txtComposicao2.Text = Trim(!strComposicao)
            txtExercicio2.Text = !intExercicio
          Case 2
            txtInscricao3.Text = IIf(!intUtilizacao <> 0, gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao), !strInscricao)
            txtComposicao3.Text = Trim(!strComposicao)
            txtExercicio3.Text = !intExercicio
          Case 3
            txtInscricao4.Text = IIf(!intUtilizacao <> 0, gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao), !strInscricao)
            txtComposicao4.Text = Trim(!strComposicao)
            txtExercicio4.Text = !intExercicio
          Case 4
            txtInscricao5.Text = IIf(!intUtilizacao <> 0, gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao), !strInscricao)
            txtComposicao5.Text = Trim(!strComposicao)
            txtExercicio5.Text = !intExercicio
          Case 5
            txtInscricao5.Text = IIf(!intUtilizacao <> 0, gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao), !strInscricao)
            txtComposicao5.Text = Trim(!strComposicao)
            txtExercicio5.Text = !intExercicio
        End Select
        
        adoInscricoes.MoveNext
        intCont = intCont + 1
     Loop
     End With
  End If
  
  'Vamos obter a conta bancaria da composicao
  strsql = "SELECT PA.intContaBancaria "
  strsql = strsql & "FROM " & gstrParametroAtualizacao & " PA, " & gstrLancamentoAlfa & " LA "
  strsql = strsql & "WHERE PA.intComposicaoReceita  = LA.intComposicaoDaReceita And PA.intExercicio = LA.intExercicio AND "
  strsql = strsql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(adoDataControl.Recordset("strInscricao"))), "0") & Trim(adoDataControl.Recordset("strInscricao")) & "' AND "
  strsql = strsql & "LA.dtmdtCancelamento IS NULL AND LA.intUtilizacao = " & TYP_ACORDO
    
  If gobjBanco.CriaADO(strsql, 10, adoRelatorio) Then
    
      With adoRelatorio
          If Not (.BOF And .EOF) Then
              lngContaBancaria = IIf(IsNull(adoRelatorio("intContaBancaria").Value), 0, adoRelatorio("intContaBancaria").Value)
          End If
      End With
        
  End If
  
  adoRelatorio.Close: Set adoRelatorio = Nothing
  
  If Not adoDataControl.Recordset.EOF Then
     
     Set adoRelatorio = New ADODB.Recordset
      
     'Vamos verificar se é Febraban ou Ficha Compensacao
     If lngContaBancaria = 0 Then
     
        With rptCarneParcelas
          
          .blnParcelasAtualizadas = Me.blnParcelasAtualizadas
          
          If gobjBanco.CriaADO(strQueryParcelas, 10, adoRelatorio) Then
             If bytDBType = EDatabases.SQLServer Then
                .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
             Else
                .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
             End If
             Set .adoDataControl.Recordset = adoRelatorio
          End If
            
          .lblTitulo.Caption = "S.M.F. - Dívida Ativa"
          .lblTitulo1.Caption = "S.M.F. - Dívida Ativa"
          .lblTipo.Caption = "Acordo Nº.:"
          .lblTipo1.Caption = "Acordo Nº.:"
          .lblProprietario.Caption = "Contribuinte"
          .lblProprietario1.Caption = "Contribuinte"
          'Campos abaixo usados somente em acordo
          If Not blnSemIndexador Then
            .lblExpresso.Visible = True
            .lblExpresso1.Visible = True
            .txtIndexador.Visible = True
            .txtIndexador1.Visible = True
            .lblAtualizacao.Caption = "Valor à Pagar"
            .lblAtualizacao1.Caption = "Valor à Pagar"
          Else
            .lblExpresso.Visible = False
            .lblExpresso1.Visible = False
            .txtIndexador.Visible = False
            .txtIndexador1.Visible = False
            .lblAtualizacao.Caption = "Atualização Monetária"
            .lblAtualizacao1.Caption = "Atualização Monetária"
          End If
           
          If Not blnSemIndexador Then
             .txtValorParcela.OutputFormat = "#,##0.0000"
             .txtValorParcela1.OutputFormat = "#,##0.0000"
          Else
             .txtValorParcela.OutputFormat = "#,##0.00"
             .txtValorParcela1.OutputFormat = "#,##0.00"
          End If
          
          .blnPrimeira = True
          .blnValorEmReal = blnSemIndexador
          
        End With
         
        Set subParcelas.object = rptCarneParcelas
    
    Else
        
        With rptCarneAcordoBoleto
          
            .blnParcelasAtualizadas = Me.blnParcelasAtualizadas
          
            If gobjBanco.CriaADO(strQueryParcelas, 15, adoRelatorio) Then
               If bytDBType = EDatabases.SQLServer Then
                  .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
               Else
                  .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
               End If
               Set .adoDataControl.Recordset = adoRelatorio
            End If
            
            .txtintConta = lngContaBancaria
            
            .txtstrContribuinte = txtProprietario
            .txtstrContribuinte1 = txtProprietario
            .txtstrLogradouroC = txtLogradouroC
            .txtstrNumeroC = txtstrNumeroC
            .txtstrComplementoC = txtstrComplementoC
            .txtstrBairroC = txtstrBairroC
            .txtstrMunicipioC = txtstrMunicipioC
            .txtstrUFC = txtstrUFC
            .txtintCEPC = txtintCEPC
            .txtdblQuantidade = txtQuantidadeFmp
            .txtdblQuantidade1 = txtQuantidadeFmp
            .txtstrInscricao = txtInscricao1
            
            If Not blnSemIndexador Then
               .txtdblValor1.OutputFormat = "#,##0.0000"
               .txtdblValor2.OutputFormat = "#,##0.0000"
               .txtdblQuantidade.OutputFormat = "#,##0.0000"
               .txtdblQuantidade1.OutputFormat = "#,##0.0000"
            Else
               .txtdblValor1.OutputFormat = "#,##0.00"
               .txtdblValor2.OutputFormat = "#,##0.00"
               .txtdblQuantidade.OutputFormat = "#,##0.00"
               .txtdblQuantidade1.OutputFormat = "#,##0.00"
            End If
                        
            .blnPrimeira = True
            .blnValorEmReal = blnSemIndexador
            
        End With
        '*************************************************
        '*Programador:Italo                              *
        '*Adicionei os Campos Abaixo Que Montam os Titu  *
        '*los do Indice                                  *
        '*************************************************
        Set subParcelas.object = rptCarneAcordoBoleto
        lblTituloTotal.Caption = "Total R$/" & txt_IndTitulo.Text
        lblTituloQtde.Caption = "Quantidade(" & txt_IndTitulo.Text & ")"
        lblTituloIndexador.Caption = "Valor " & txt_IndTitulo.Text
    End If
    
  End If
  
End Sub

Private Function strQueryInscricoes() As String
Dim strsql As String
        
  strsql = ""
  strsql = strsql & "SELECT DISTINCT "
  strsql = strsql & "AD.strIdentificacao strInscricao, "
  strsql = strsql & "AD.strComposicaoDaReceita strComposicao, "
  strsql = strsql & "AD.intExercicio, "
  strsql = strsql & "CASE WHEN AD.intUtilizacao IS NULL THEN 0 ELSE AD.intUtilizacao END intUtilizacao "
  strsql = strsql & "FROM "
  strsql = strsql & gstrLancamentoAlfa & " LA, "
  strsql = strsql & gstrAcordo & " AC, "
  strsql = strsql & gstrAcordoDebitos & " AD "
  strsql = strsql & "WHERE "
  strsql = strsql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(adoDataControl.Recordset("strInscricao"))), "0") & Trim(adoDataControl.Recordset("strInscricao")) & "' AND "
  strsql = strsql & "LA.dtmdtCancelamento IS NULL AND "
  strsql = strsql & "AC.intLancamentoAlfa = LA.pkID AND "
  strsql = strsql & "AD.intAcordo = AC.pkID "

  strsql = strsql & "ORDER BY strInscricao "
  
  strQueryInscricoes = strsql
End Function

Private Function strQueryParcelas() As String
Dim strsql As String
Dim strBarra As String
Dim intCont As Integer
  
  strsql = ""
    
  strsql = strsql & "SELECT "
  strsql = strsql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ACORDO)) & " strInscricao, "
  strsql = strsql & "LA.intExercicio intExercicio, LA.Pkid PkidAlfa, LA.strComposicaoDaReceita, "
  strsql = strsql & "CR.strSigla strSigla, CR.intUtilizacao intUtilizacao, LA.intComposicaoDaReceita intComposicao, "
  strsql = strsql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strAviso, 9999 strGuia, "
  strsql = strsql & "LA.strNomeProprietario strProprietario, "
  
  strsql = strsql & "LA.dblvlIndexador, "
  strsql = strsql & "LA.strIndexador, "
  
  'Nº da Parcela
  strsql = strsql & "LV.intParcela intParcela, "
  strsql = strsql & "LV.pkID pkIDParcela, LV.bitParcelaValida, LV.intMoeda, "
  
  'Valor da Parcela
  If blnSemIndexador Then
     strsql = strsql & "LV.dblValor "
  Else
     If (bytDBType = EDatabases.Oracle) Then
         strsql = strsql & strSUBSTRING & "(LV.dblValor / LA.dblvlIndexador, 0," & gstrINSTR(",", "LV.dblValor / LA.dblvlIndexador", 1, 1) & " + 4) "
     Else
         strsql = strsql & " REPLACE(" & strSUBSTRING & "(" & gstrCONVERT(CDT_VARCHAR, "LV.dblValor / LA.dblvlIndexador") & ", 0," & gstrINSTR(".", "LV.dblValor / LA.dblvlIndexador", 1, 1) & " + 5), '.', ',') "
     End If
  End If
  strsql = strsql & "dblValorParcela, "
     
  'Valor a ser gravado na Guia
  strsql = strsql & "LV.dblValor dblValorReal, "
   
  'Vencimento da Primeira Parcela
  If (bytDBType = EDatabases.Oracle) Then
      strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'DD/MM/YYYY'") & " dtmdtVencimento, "
  Else
      strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 103") & " dtmdtVencimento, "
      'PROVISORIO
      'strsql = strsql & gstrCONVERT(CDT_VARCHAR, "CASE when month(LV.dtmdtVencimento) =1 then '2006-01-31' when month(LV.dtmdtVencimento) = 2 then '2006-02-28' else LV.dtmdtVencimento end, 103") & " dtmdtVencimento, "
  End If
  
  'Código de Barras===================================
  strsql = strsql & "'817'" & strCONCAT & " "

  'Valor da Parcela
  If (bytDBType = EDatabases.Oracle) Then
    If blnSemIndexador Then
       strBarra = strSUBSTRING & "(LV.dblValor,0," & gstrINSTR(",", "LV.dblValor", 1, 1) & " + 4) * 100 "
    Else
       strBarra = strSUBSTRING & "(LV.dblValor / LA.dblvlIndexador, 0, " & gstrINSTR(",", "LV.dblValor / LA.dblvlIndexador", 1, 1) & " + 4) * 10000 "
    End If
  Else
    If blnSemIndexador Then
       strBarra = " REPLACE(" & strSUBSTRING & "(" & gstrCONVERT(CDT_VARCHAR, "LV.dblValor") & ",0," & gstrINSTR(".", "LV.dblValor", 1, 1) & " + 5), '.', '') "
    Else
       strBarra = " REPLACE(" & strSUBSTRING & "(" & gstrCONVERT(CDT_VARCHAR, "LV.dblValor / LA.dblvlIndexador") & ", 0, " & gstrINSTR(".", "LV.dblValor / LA.dblvlIndexador", 1, 1) & " + 5), '.', '') "
    End If
  End If
  
  'strSql = strSql & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & " , '00000000000'") & ") " & strCONCAT & " "
  strsql = strsql & "LTRIM(" & gstrREPLICATE(strBarra, "0", 11) & ") " & strCONCAT & " "
  
  'Febraban
  strBarra = "(SELECT intFebraban FROM " & gstrEmpresa & ")"
  'strSQL = strSQL & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & ", '0000'") & ") " & strCONCAT & " "
  strsql = strsql & "LTRIM(" & gstrREPLICATE(strBarra, "0", 4) & ") " & strCONCAT & " "
  
  'Vencimento
  If (bytDBType = EDatabases.Oracle) Then
      strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'YYYYMMDD'") & strCONCAT & " "
  Else
      strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 112") & strCONCAT & " "
  End If
  
  'Conta Bancária
  strsql = strsql & "'0000' " & strCONCAT & " "
  
  'Guia (vai ser substituído no próprio rpt)
  strBarra = "9999"
  
  'strSQL = strSQL & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & ", '000000000'") & ") " & strCONCAT & " "
  strsql = strsql & "LTRIM(" & gstrREPLICATE(strBarra, "0", 9) & ") " & strCONCAT & " "
  
  'Exercicio
  strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LA.intExercicio")

  strsql = strsql & " strCodigoBarra "
  
  'FROM
  strsql = strsql & "FROM "
  strsql = strsql & gstrLancamentoAlfa & " LA, "
  strsql = strsql & gstrComposicaoDaReceita & " CR, "
  strsql = strsql & gstrLancamentoValor & " LV "
    
  'WHERE
  strsql = strsql & "WHERE "
  strsql = strsql & "CR.pkID = " & adoDataControl.Recordset("intComposicao") & " AND "
  strsql = strsql & "CR.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.intComposicaoDaReceita  AND "
  strsql = strsql & "LA.intExercicio = " & txtExercicio.Text & " AND "
  strsql = strsql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(adoDataControl.Recordset("strInscricao"))), "0") & Trim(adoDataControl.Recordset("strInscricao")) & "' AND "
  strsql = strsql & "LA.strEmissao = '" & adoDataControl.Recordset("strEmissao") & "' AND "
  strsql = strsql & "LV.intLancamentoAlfa = LA.pkID AND "
  
  'Vamos verificar se é atualizacao por exercicio
  If blnParcelasAtualizadas And intExercicioAtualizadas > 0 Then
    strsql = strsql & gstrDATEPART(strYEAR, "LV.dtmDtVencimento") & " = " & Trim(intExercicioAtualizadas) & " AND "
  Else
    strsql = strsql & "LV.intParcela IN (" & strParcelasSelecionadas & ") AND "
  End If
  
  strsql = strsql & "LA.intUtilizacao = " & TYP_ACORDO & " "
  strsql = strsql & "ORDER BY LV.bitParcelaValida, LV.intParcela "
  
  strQueryParcelas = strsql
End Function

