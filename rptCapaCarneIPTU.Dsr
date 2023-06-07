VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCapaCarneIPTU 
   Caption         =   "Tributario - rptCapaCarneIPTU (ActiveReport)"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptCapaCarneIPTU.dsx":0000
End
Attribute VB_Name = "rptCapaCarneIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strParcelasSelecionadas As String 'Parcelas selecionadas vindo do form
Public strEmpresaFebraban As String 'Febraban

Private Sub ActiveReport_Activate()
  HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_ReportStart()
  If adoDataControl.Recordset.RecordCount = 0 Then
     ExibeMensagem "Não existem registros com os dados informados."
     Unload Me
     Exit Sub
  End If
  
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

Private Function strQuery() As String
Dim strsql As String
  strsql = ""
  strsql = strsql & "SELECT "
  strsql = strsql & "LF.strDescricao, LF.dblFator "
  strsql = strsql & "FROM "
  strsql = strsql & gstrLancamentoFatores & " LF, "
  strsql = strsql & gstrLancamentoAlfa & " LA, "
  strsql = strsql & gstrLancamentoIPTU & " LI "
  strsql = strsql & "WHERE "
  strsql = strsql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(txtImovel.Text)), "0") & Trim(txtImovel.Text) & "' AND "
  strsql = strsql & "LI.intLancamentoAlfa = LA.pkID AND "
  strsql = strsql & "LF.intLancamentoIptu = LI.pkID "
  strQuery = strsql

End Function


Private Sub GroupHeader1_Format()
Dim intCount          As Integer
Dim adoRecFatores     As ADODB.Recordset
Dim adoRelatorio      As ADODB.Recordset
Dim strCodigoBarra    As String
Dim strsql            As String
Dim lngContaBancaria  As Long

On Error GoTo Problema_Na_Rotina

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strQuery, 10, adoRecFatores) Then
     
        intCount = 0
        Do While Not adoRecFatores.EOF
        
          Select Case intCount
            Case 0
              lblFator1.Caption = adoRecFatores(0)
              txtFator1.Text = gstrConvVrDoSql(adoRecFatores(1), 2)
            Case 1
              lblFator2.Caption = adoRecFatores(0)
              txtFator2.Text = gstrConvVrDoSql(adoRecFatores(1), 2)
            Case 2
              lblFator3.Caption = adoRecFatores(0)
              txtFator3.Text = gstrConvVrDoSql(adoRecFatores(1), 2)
            Case 3
              lblFator4.Caption = adoRecFatores(0)
              txtFator4.Text = gstrConvVrDoSql(adoRecFatores(1), 2)
            Case 4
              lblFator5.Caption = adoRecFatores(0)
              txtFator5.Text = gstrConvVrDoSql(adoRecFatores(1), 2)
            Case 5
              lblFator6.Caption = adoRecFatores(0)
              txtFator6.Text = gstrConvVrDoSql(adoRecFatores(1), 2)
            Case 6
              lblFator7.Caption = adoRecFatores(0)
              txtFator7.Text = gstrConvVrDoSql(adoRecFatores(1), 2)
            Case 7
              lblFator8.Caption = adoRecFatores(0)
              txtFator8.Text = gstrConvVrDoSql(adoRecFatores(1), 2)
          End Select
          
          intCount = intCount + 1
          adoRecFatores.MoveNext
       Loop
    End If
  
    txtImovel.Text = gstrFormataInscricao(txtImovel.Text, txtUtilizacao.Text)
    txtImovel1.Text = txtImovel.Text
  
    'Se for pedido somente capa, não gera código de barra e nem as parcelas
    'If strParcelasSelecionadas = "0" Then
    '   GroupHeader1.NewPage = ddNPNone
    '   Exit Sub
    'Else
    '   GroupHeader1.NewPage = ddNPAfter
    'End If
  
    'Vamos obter a conta bancaria da composicao
    strsql = "SELECT PA.intContaBancaria "
    strsql = strsql & "FROM " & gstrParametroAtualizacao & " PA, " & gstrLancamentoAlfa & " LA "
    strsql = strsql & "WHERE PA.intComposicaoReceita  = " & adoDataControl.Recordset("intComposicao") & " And PA.intExercicio = " & adoDataControl.Recordset("intExercicio") & " AND "
    strsql = strsql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(adoDataControl.Recordset("strImovel"))), "0") & Trim(adoDataControl.Recordset("strImovel")) & "' AND "
    strsql = strsql & "LA.dtmdtCancelamento IS NULL "
      
    If gobjBanco.CriaADO(strsql, 10, adoRelatorio) Then
      
        With adoRelatorio
            If Not (.BOF And .EOF) Then
                lngContaBancaria = IIf(IsNull(adoRelatorio("intContaBancaria").Value), 0, adoRelatorio("intContaBancaria").Value)
            End If
        End With
          
    End If
  
    adoRelatorio.Close: Set adoRelatorio = Nothing
      
    'Vamos carregar as parcelas
    If Not adoDataControl.Recordset.EOF Then
      
        Set adoRelatorio = New ADODB.Recordset
        
               
        'Vamos verificar se é Febraban ou Ficha Compensacao
        If lngContaBancaria = 0 Then
    
            With rptCarneParcelas
              
                If gobjBanco.CriaADO(strQueryParcelas, 5, adoRelatorio) Then
                    If bytDBType = EDatabases.SQLServer Then
                       .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                    Else
                       .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                    End If
                    Set .adoDataControl.Recordset = adoRelatorio
                End If
              
                .lblTitulo.Caption = "S.M.F. - Seção de Rendas Imobiliárias"
                .lblTitulo1.Caption = "S.M.F. - Seção de Rendas Imobiliárias"
                .lblTipo.Caption = "Imóvel:"
                .lblTipo1.Caption = "Imóvel:"
                .lblProprietario.Caption = "Proprietário"
                .lblProprietario1.Caption = "Proprietário"
                'Campos abaixo usados somente em acordo
                'Expresso em R$ fixo
                .lblExpresso.Visible = True
                .lblExpresso1.Visible = True
                .txtIndexador.Visible = True
                .txtIndexador1.Visible = True
                .txtIndexador.Text = "R$"
                .txtIndexador1.Text = "R$"
                .lblAtualizacao.Caption = "Atualização Monetária"
                .lblAtualizacao1.Caption = "Atualização Monetária"
                .txtValorParcela.OutputFormat = "#,##0.00"
                .txtValorParcela1.OutputFormat = "#,##0.00"
                .blnPrimeira = False
            
            End With
            
            Set subParcelas.object = rptCarneParcelas
         
        Else
            
            With rptCarneAcordoBoleto
              
                If gobjBanco.CriaADO(strQueryParcelas, 5, adoRelatorio) Then
                   If bytDBType = EDatabases.SQLServer Then
                      .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                   Else
                      .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                   End If
                   Set .adoDataControl.Recordset = adoRelatorio
                End If
                
                .txtintConta = lngContaBancaria
                
                .txtstrContribuinte = Field224
                .txtstrContribuinte1 = Field224
                .txtstrLogradouroC = Field226
                .txtstrNumeroC = Field259
                .txtstrComplementoC = Field260
                .txtstrBairroC = Field261
                .txtstrMunicipioC = Field227
                .txtstrUFC = Field228
                .txtintCEPC = Field229
                .txtstrInscricao = txtImovel1
                
                .txtdblValor1.OutputFormat = "#,##0.00"
                .txtdblValor2.OutputFormat = "#,##0.00"
                
                .blnPrimeira = True
                .blnValorEmReal = True
                
            If Len(txtInscrAuxiliar.Text) = 0 Or txtInscrAuxiliar.Text = "NULL" Then
                    Label281.Visible = False
            Else
                    Label281.Visible = True
            End If
  
            
            End With
            
            Set subParcelas.object = rptCarneAcordoBoleto
             
        End If
         
    End If
    
    Exit Sub
  

Problema_Na_Rotina:
   
    ExibeDetalheErro Err.Description & "- rptCapaCarne_GroupHeader1_Format"
    gobjBanco.ExecutaRollbackTrans
    
End Sub

Private Function strQueryParcelas() As String
Dim strsql As String
Dim strBarra As String
Dim intCont As Integer
  
  strsql = ""
    
  strsql = strsql & "SELECT " & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
  strsql = strsql & "LA.intExercicio intExercicio, LA.Pkid PkidAlfa, LA.strComposicaoDaReceita, "
  strsql = strsql & "CR.strSigla strSigla, CR.intUtilizacao intUtilizacao, LA.intComposicaoDaReceita intComposicao, "
  strsql = strsql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strAviso, 9999 strGuia, "
  strsql = strsql & "LA.strNomeProprietario strProprietario, LA.strIndexador, "
    
  'Nº da Parcela
  strsql = strsql & "LV.intParcela intParcela, "
  strsql = strsql & "LV.pkID pkIDParcela, LV.bitParcelaValida, "
  
  'Valor da Parcela
  strsql = strsql & "LV.dblValor dblValorParcela, "
  
  'Valor a ser gravado na Guia
  strsql = strsql & "LV.dblValor dblValorReal, "
  If bytDBType = Oracle Then
      'Vencimento da Primeira Parcela
      strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'DD/MM/YYYY'") & " dtmdtVencimento, "
      'Código de Barras===================================
      strsql = strsql & "'817' " & strCONCAT & " "
      'Valor da Parcela
      strBarra = "LV.dblValor * 100"
      strsql = strsql & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & " , '00000000000'") & ") " & strCONCAT & " "
      'Febraban
      strBarra = "(SELECT intFebraban FROM " & gstrEmpresa & ")"
      strsql = strsql & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & ", '0000'") & ") " & strCONCAT & " "
      'Vencimento
      strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'YYYYMMDD'") & strCONCAT & " "
      'Conta Bancária
      strsql = strsql & "'0000' " & strCONCAT & " "
      'Guia
      strBarra = "9999"
      strsql = strsql & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & ", '000000000'") & ") " & strCONCAT & " "
      'Exercicio
      strsql = strsql & "LA.intExercicio "
      strsql = strsql & "strCodigoBarra "
    Else
      strsql = strsql & "CONVERT (VARCHAR, LV.dtmdtVencimento, 103)  dtmdtVencimento, "
      strsql = strsql & " '817' + "
      strsql = strsql & " convert(varchar,replicate(0,11-len(REPLACE(isNull(LV.dblValor,0), '.',''))) + REPLACE(isNull(LV.dblValor,0), '.','')) + "
      strsql = strsql & " convert(varchar,replicate(0,4-len((SELECT isnull(intFebraban,0) FROM tblEmpresa))) + (SELECT isnull(intFebraban,0) FROM tblEmpresa)) + "
      strsql = strsql & " convert(varchar,REPLACE(  CONVERT (VARCHAR, LV.dtmdtVencimento, 103) , '/', '')) + "
      strsql = strsql & " '0000'  + "
      strsql = strsql & " '000009999' + "
      strsql = strsql & " convert(varchar,LA.intExercicio) strCodigoBarra "
    End If
    
  'FROM
  strsql = strsql & "FROM "
  strsql = strsql & gstrLancamentoAlfa & " LA, "
  strsql = strsql & gstrComposicaoDaReceita & " CR, "
  strsql = strsql & gstrLancamentoValor & " LV "
  
  'WHERE
  strsql = strsql & "WHERE "
  strsql = strsql & "CR.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.intComposicaoDaReceita AND "
  'strSql = strSql & "CR.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & "LA.intComposicaoDaReceita  AND "
  strsql = strsql & "LA.intComposicaoDaReceita = " & txtComposicao.Text & " AND "
  strsql = strsql & "LA.intExercicio = " & txtExercicio.Text & " AND "
  strsql = strsql & "LA.strEmissao = '" & txtEmissao.Text & "' AND "
  strsql = strsql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(txtInscricao.Text)), "0") & Trim(txtInscricao.Text) & "' AND "
  strsql = strsql & "LA.dtmdtCancelamento IS NULL AND "
  strsql = strsql & "LV.intLancamentoAlfa = LA.pkID AND "
  strsql = strsql & "LV.intParcela IN (" & strParcelasSelecionadas & ") "
  strsql = strsql & "ORDER BY LV.bitParcelaValida, "
  
  If (bytDBType = EDatabases.Oracle) Then
      strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'YYYYMMDD'")
  Else
      strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 112")
  End If

  strQueryParcelas = strsql
  
End Function

