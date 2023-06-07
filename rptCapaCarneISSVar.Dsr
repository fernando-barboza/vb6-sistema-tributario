VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCapaCarneISSVar 
   Caption         =   "rptCapaCarneISSVar (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptCapaCarneISSVar.dsx":0000
End
Attribute VB_Name = "rptCapaCarneISSVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strParcelasSelecionadas As String 'Parcelas selecionadas vindo do form
Public strEmpresaFebraban As String 'Febraban

Private Sub ActiveReport_ReportStart()
  PadronizaToolBarRelatorio Me
  LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia
  LeImagemLogotipo imgBrasao1, imgLogotipo, txtNomeFantasia1
End Sub

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

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf tool.ID = 16 Then
        Configura_Relatorio Me, True
    End If
End Sub

Private Sub GroupHeader1_Format()
Dim adoRelatorio As ADODB.Recordset
Dim adoAtividades As ADODB.Recordset
Dim intCont As Byte
Dim adoResultado As ADODB.Recordset
Dim strSql As String
Dim lngContaBancaria  As Long

  
    strSql = ""
    strSql = strSql & " SELECT EI.strTipoIss FROM "
    strSql = strSql & " tblLancamentoEconIss EI, "
    strSql = strSql & " tblLancamentoAlfa LA, "
    strSql = strSql & " tblLancamentoEconomico LE "
    strSql = strSql & " WHERE "
    strSql = strSql & " LE.PKID = EI.intLancamentoEconomico "
    strSql = strSql & " AND LA.PKID = LE.intLancamentoAlfa "
    strSql = strSql & " AND LA.PKID = " & adoDataControl.Recordset("pkIdPrincipal")
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            If UCase(adoResultado!STRTIPOISS) = "ISS AUTO-LANCAMENTO" Then
                xAutoLcto = "X"
            ElseIf UCase(adoResultado!STRTIPOISS) = "ISS ESTIMATIVA" Then
                xEstimativa = "X"
            End If
        End If
    End If
    
  
    strSql = ""
    strSql = strSql & "Select TL.Strsigla TipoLogradouro, "
    strSql = strSql & "LO.Strdescricao Logradouro, "
    strSql = strSql & "EM.Strnumero Numero, "
    strSql = strSql & "EM.strComplemento Complemento, "
    strSql = strSql & "BA.Strdescricao Bairro, "
    strSql = strSql & "MU.Strdescricao Municipio, "
    strSql = strSql & "UF.Strsigla Estado, "
    strSql = strSql & "EM.intCep Cep "
    strSql = strSql & "from "
    strSql = strSql & "tblEmpresa EM, "
    strSql = strSql & "Tblbairro BA, "
    strSql = strSql & "Tblmunicipio MU, "
    strSql = strSql & "tblUF UF, "
    strSql = strSql & "tblLogradouro LO, "
    strSql = strSql & "tbltipologradouro TL "
    strSql = strSql & "where MU.Pkid = EM.Intcidade "
    strSql = strSql & "and BA.pkid = EM.Intbairro "
    strSql = strSql & "and UF.Pkid = EM.Intuf "
    strSql = strSql & "and LO.pkid = EM.Intlogradouro "
    strSql = strSql & "and TL.pkid = EM.Inttipologradouro "
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txtstrEndereco = ""
                txtstrEndereco = txtstrEndereco & gstrENulo(!TipoLogradouro)
                txtstrEndereco = txtstrEndereco & " " & gstrENulo(!Logradouro)
                txtstrEndereco = txtstrEndereco & IIf(gstrENulo(!Numero) <> "", ", " & gstrENulo(!Numero), "")
                txtstrEndereco = txtstrEndereco & " " & gstrENulo(!Complemento)
                txtstrEndereco = txtstrEndereco & IIf(gstrENulo(!CEP) <> "", "     CEP " & gstrCEPFormatado(gstrENulo(!CEP)), "")
                txtstrEndereco = txtstrEndereco & "     " & gstrENulo(!Municipio)
                txtstrEndereco = txtstrEndereco & IIf(gstrENulo(!Estado) <> "", " - " & gstrENulo(!Estado), "")
                
                txtstrEndereco1 = txtstrEndereco
            End With
        End If
    End If
  
  
  Set gobjBanco = New clsBanco
  intCont = 0
  
  If gobjBanco.CriaADO(strQueryAtividades, 5, adoAtividades) Then
     Do While Not adoAtividades.EOF
        Select Case intCont
          Case 0
            txtstrAtividade.Text = adoAtividades(0)
          Case Else
            Exit Do
        End Select
      adoAtividades.MoveNext
        intCont = intCont + 1
     Loop
  End If

    strSql = "SELECT PA.intContaBancaria "
    strSql = strSql & "FROM " & gstrParametroAtualizacao & " PA, " & gstrLancamentoAlfa & " LA "
    strSql = strSql & "WHERE PA.intComposicaoReceita  = " & adoDataControl.Recordset("intComposicao") & " And PA.intExercicio = " & adoDataControl.Recordset("intExercicio") & " AND "
    strSql = strSql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(adoDataControl.Recordset("strInscricao"))), "0") & Trim(adoDataControl.Recordset("strInscricao")) & "' AND "
    strSql = strSql & "LA.dtmdtCancelamento IS NULL "
      
    If gobjBanco.CriaADO(strSql, 10, adoRelatorio) Then
      
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
        If Not lngContaBancaria = 0 Then

  
            Set adoRelatorio = New ADODB.Recordset
            Set gobjBanco = New clsBanco
            
            With rptBoleto
              If gobjBanco.CriaADO(strQueryParcelas, 5, adoRelatorio) Then
                 If bytDBType = EDatabases.SQLServer Then
                    .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                 Else
                    .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                 End If
                 Set .adoDataControl.Recordset = adoRelatorio
              End If
              
                    .txtintConta = lngContaBancaria
                            
                    .txtstrInscricao = txtstrInscricao
                    .txtstrAviso = txtstrAviso
                    .txtstrContribuinte = txtstrContribuinte
                    .txtstrAtividade = txtstrAtividade
                    
                    .txtstrNomeFantasia = txtNomeFantasia
                    
                    .txtstrContribuinte1 = txtstrContribuinte
                    .txtstrLogradouro = txtstrLogradouro
                    .txtstrNumero = txtintNumero
                    .txtstrComplemento = txtstrComplemento
                    .txtstrBairro = txtstrBairro
                    .txtstrMunicipio = txtstrMunicipio
                    .txtstrUf = txtstrUf
                    .txtstrCep = txtstrCep
                    .txtstrInscricao = txtstrInscricao
                    
            End With
            
            Set subParcelas.object = rptBoleto
        End If
    End If
  
End Sub

Private Function strQueryAtividades() As String
Dim strSql As String
  
  strSql = ""
  strSql = strSql & "SELECT "
  'strSql = strSql & "EC.strInscricaoCadastral, "
  strSql = strSql & "RTRIM(LTRIM(AEC.strDescricao)) "
  'strSql = strSql & "AE.blnPrincipal "
  strSql = strSql & "FROM "
  strSql = strSql & gstrAtividadeEC & " AEC, "
  strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
  strSql = strSql & gstrEconomico & " EC "
  strSql = strSql & "WHERE "
  strSql = strSql & "EC.strInscricaoCadastral = '" & String(gintLenInscricao - Len(Trim(txtInscricao.Text)), "0") & Trim(txtInscricao.Text) & "' AND "
  strSql = strSql & "AE.intEconomico = EC.pkID AND "
  strSql = strSql & "AEC.pkID = AE.intAtividade "
  strSql = strSql & "ORDER BY "
  strSql = strSql & "AE.blnPrincipal DESC "
  
  strQueryAtividades = strSql
End Function

Private Function strQueryReceitas() As String
    Dim strSql As String
  
  strSql = ""
  strSql = strSql & "SELECT "
  strSql = strSql & "RTRIM(LTRIM(RE.strSigla)) strSigla, "
  strSql = strSql & "SUM(" & gstrISNULL("LR.dblValor", "0") & ") dblValor "
  strSql = strSql & "FROM "
  strSql = strSql & "tblReceita RE, "
  strSql = strSql & "tblLancamentoReceita LR, "
  strSql = strSql & "tblLancamentoValor LV "
  strSql = strSql & "WHERE "
  strSql = strSql & "LV.intLancamentoAlfa = " & txtpkID.Text & " AND "
  strSql = strSql & "LV.bItParcelaValida = 1 AND "
  strSql = strSql & "LR.intLancamentoValor = LV.pkID AND "
  strSql = strSql & "RE.pkID = LR.intReceita "
  strSql = strSql & "GROUP BY "
  strSql = strSql & "RE.pkID, "
  strSql = strSql & "RE.strSigla, "
  strSql = strSql & "LR.dblValor "
  strSql = strSql & "ORDER BY "
  strSql = strSql & "strSigla "

  strQueryReceitas = strSql
End Function

Private Function strQueryParcelas() As String
Dim strSql As String
Dim strBarra As String
Dim intCont As Integer


  
  strSql = ""
    
  strSql = strSql & "SELECT "
  strSql = strSql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao, "
  strSql = strSql & "LA.intExercicio intExercicio, LA.Pkid PkidAlfa, LA.strComposicaoDaReceita, "
  strSql = strSql & "CR.strSigla strSigla, CR.intUtilizacao intUtilizacao, LA.intComposicaoDaReceita intComposicao, "
  strSql = strSql & gstrCONVERT(CDT_INT, "LA.strNumeroAviso") & " strAviso, 9999 strGuia, "
  strSql = strSql & "LA.strNomeProprietario strProprietario, "
    
  'Nº da Parcela
  strSql = strSql & "LV.intParcela intParcela, "
  strSql = strSql & "LV.pkID pkIDParcela, LV.bitParcelaValida,  "
  
  'Valor da Parcela
  strSql = strSql & "LV.dblValor dblValorParcela, "
  
  'Valor a ser gravado na Guia
  strSql = strSql & "LV.dblValor dblValorReal, "
   
    'Vencimento da Primeira Parcela
    If (bytDBType = EDatabases.Oracle) Then
        strSql = strSql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'DD/MM/YYYY'") & " dtmdtVencimento, "
    Else
        strSql = strSql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 103") & " dtmdtVencimento, "
    End If
      
    'Código de Barras===================================
    strSql = strSql & "'817' " & strCONCAT & " "
    
    'Valor da Parcela
    If (bytDBType = EDatabases.Oracle) Then
         strBarra = strSUBSTRING & "(LV.dblValor,0," & gstrINSTR(",", "LV.dblValor", 1, 1) & " + 4) * 100 "
    Else
         strBarra = " REPLACE(" & strSUBSTRING & "(" & gstrCONVERT(CDT_VARCHAR, "LV.dblValor") & ",0," & gstrINSTR(".", "LV.dblValor", 1, 1) & " + 5), '.', '') "
    End If
    
    strSql = strSql & "LTRIM(" & gstrREPLICATE(strBarra, "0", 11) & ") " & strCONCAT & " "
    
    'Febraban
    strBarra = "(SELECT intFebraban FROM " & gstrEmpresa & ")"
    'strSQL = strSQL & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & ", '0000'") & ") " & strCONCAT & " "
    strSql = strSql & "LTRIM(" & gstrREPLICATE(strBarra, "0", 4) & ") " & strCONCAT & " "
    
    'Vencimento
    If (bytDBType = EDatabases.Oracle) Then
        strSql = strSql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'YYYYMMDD'") & strCONCAT & " "
    Else
        strSql = strSql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 112") & strCONCAT & " "
    End If
    
    'Conta Bancária
    strSql = strSql & "'0000' " & strCONCAT & " "
    
    'Guia
    strBarra = "9999"
    
    strSql = strSql & "LTRIM(" & gstrREPLICATE(strBarra, "0", 9) & ") " & strCONCAT & " "
    
    'Exercicio
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "LA.intExercicio")
    
    strSql = strSql & " strCodigoBarra "

  
  'FROM
  strSql = strSql & "FROM "
  strSql = strSql & gstrLancamentoAlfa & " LA, "
  strSql = strSql & gstrComposicaoDaReceita & " CR, "
  strSql = strSql & gstrLancamentoValor & " LV "
  
  'WHERE
  strSql = strSql & "WHERE "
  strSql = strSql & "CR.pkID = " & txtComposicao.Text & " AND "
  strSql = strSql & "CR.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.intComposicaoDaReceita  AND "
  strSql = strSql & "LA.intExercicio = " & txtintExercicio.Text & " AND "
  strSql = strSql & "LA.strEmissao = '" & txtEmissao.Text & "' AND "
  strSql = strSql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(txtInscricao.Text)), "0") & Trim(txtInscricao.Text) & "' AND "
  strSql = strSql & "LA.dtmdtCancelamento IS NULL AND "
  strSql = strSql & "LV.intLancamentoAlfa = LA.pkID AND "
  strSql = strSql & "LV.intParcela IN (" & strParcelasSelecionadas & ") "
  strSql = strSql & "ORDER BY LV.intParcela "
  
  strQueryParcelas = strSql
End Function

