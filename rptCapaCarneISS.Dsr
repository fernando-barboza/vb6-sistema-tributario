VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCapaCarneISS 
   Caption         =   "rptCapaCarneISS (ActiveReport)"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptCapaCarneISS.dsx":0000
End
Attribute VB_Name = "rptCapaCarneISS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strParcelasSelecionadas As String 'Parcelas selecionadas vindo do form
Public strEmpresaFebraban As String 'Febraban
Public strNumeroAviso As String

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
Dim adoAtividades As ADODB.Recordset
Dim adoRelatorio As ADODB.Recordset
Dim adoReceitas As ADODB.Recordset
Dim intCont As Byte
Dim dblTotal As Double
Dim lngContaBancaria    As Long
Dim strsql As String
  
  Set gobjBanco = New clsBanco
  intCont = 0
  
  If gobjBanco.CriaADO(strQueryAtividades, 5, adoAtividades) Then
     Do While Not adoAtividades.EOF
        Select Case intCont
          Case 0
            lblAtividade1.Visible = True
            txtAtividade1.Visible = True
            txtAtividade1.Text = adoAtividades(0)
          Case 1
            lblAtividade2.Visible = True
            txtAtividade2.Visible = True
            txtAtividade2.Text = adoAtividades(0)
          Case 2
            lblAtividade3.Visible = True
            txtAtividade3.Visible = True
            txtAtividade3.Text = adoAtividades(0)
          Case 3
            lblAtividade4.Visible = True
            txtAtividade4.Visible = True
            txtAtividade4.Text = adoAtividades(0)
          Case 4
            lblAtividade5.Visible = True
            txtAtividade5.Visible = True
            txtAtividade5.Text = adoAtividades(0)
          Case 5
            lblAtividade6.Visible = True
            txtAtividade6.Visible = True
            txtAtividade6.Text = adoAtividades(0)
          Case Else
            Exit Do
        End Select
        
        adoAtividades.MoveNext
        intCont = intCont + 1
     Loop
  End If
  
  intCont = 0
  
  txtReceita1.Text = ""
  txtReceita2.Text = ""
  txtReceita3.Text = ""
  txtReceita4.Text = ""
  txtReceita5.Text = ""
  
  txtValor1.Text = ""
  txtValor2.Text = ""
  txtValor3.Text = ""
  txtValor4.Text = ""
  txtValor5.Text = ""
  
  If gobjBanco.CriaADO(strQueryReceitas, 5, adoReceitas) Then
     Do While Not adoReceitas.EOF
        Select Case intCont
          Case 0
            txtReceita1.Text = adoReceitas(0)
            txtValor1.Text = Format$(adoReceitas(1), "#,##0.00")
          Case 1
            txtReceita2.Text = adoReceitas(0)
            txtValor2.Text = Format$(adoReceitas(1), "#,##0.00")
          Case 2
            txtReceita3.Text = adoReceitas(0)
            txtValor3.Text = Format$(adoReceitas(1), "#,##0.00")
          Case 3
            txtReceita4.Text = adoReceitas(0)
            txtValor4.Text = Format$(adoReceitas(1), "#,##0.00")
          Case 4
            txtReceita5.Text = adoReceitas(0)
            txtValor5.Text = Format$(adoReceitas(1), "#,##0.00")
        End Select
        
        dblTotal = dblTotal + adoReceitas(1)
        adoReceitas.MoveNext
        intCont = intCont + 1
     Loop
     txtTotalLancado.Text = Format(dblTotal, "#,##0.00")
     
     'VALOR DO FMP
     If Trim(txtIndexador.Text) <> "" Then
        txtFmpTotal.Text = txtTotalLancado.Text / txtIndexador.Text
        If Len(txtFmpTotal.Text) - InStr(txtFmpTotal.Text, ",") >= 4 Then
           txtFmpTotal.Text = Left(txtFmpTotal.Text, InStr(txtFmpTotal.Text, ",") + 4)
        End If

        txtFmpParcela.Text = txtValorParcela.Text / txtIndexador.Text
        If Len(txtFmpParcela.Text) - InStr(txtFmpParcela.Text, ",") >= 4 Then
           txtFmpParcela.Text = Left(txtFmpParcela.Text, InStr(txtFmpParcela.Text, ",") + 4)
        End If
     End If
     
  End If
  
  txtInsc.Text = gstrFormataInscricao(txtInsc.Text, txtUtilizacao.Text)
  txtInsc1.Text = txtInsc.Text
  
  
   'Vamos obter a conta bancaria da composicao
    strsql = "SELECT PA.intContaBancaria "
    strsql = strsql & "FROM " & gstrParametroAtualizacao & " PA, " & gstrLancamentoAlfa & " LA "
    strsql = strsql & "WHERE PA.intComposicaoReceita  = " & adoDataControl.Recordset("intComposicao") & " And PA.intExercicio = " & adoDataControl.Recordset("intExercicio") & " AND "
    strsql = strsql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(adoDataControl.Recordset("strInscricao"))), "0") & Trim(adoDataControl.Recordset("strInscricao")) & "' AND "
    strsql = strsql & "LA.dtmdtCancelamento IS NULL "
      
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
     Set gobjBanco = New clsBanco
     
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
              .lblTitulo.Caption = "S.M.F. - Divisão de Rendas Mobiliárias"
              .lblTitulo1.Caption = "S.M.F. - Divisão de Rendas Mobiliárias"
              .lblTipo.Caption = "Insc. Mun.:"
              .lblTipo1.Caption = "Insc. Mun.:"
              .lblProprietario.Caption = "Contribuinte"
              .lblProprietario1.Caption = "Contribuinte"
              .blnPrimeira = True
              'Campos abaixo usados somente em acordo
              .lblExpresso.Visible = False
              .lblExpresso1.Visible = False
              .txtIndexador.Visible = False
              .txtIndexador1.Visible = False
              .lblAtualizacao.Caption = "Atualização Monetária"
              .lblAtualizacao1.Caption = "Atualização Monetária"
              .txtValorParcela.OutputFormat = "#,##0.00"
              .txtValorParcela1.OutputFormat = "#,##0.00"
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
                
                .txtintConta = lngContaBancaria 'INSERIR O NUMERO DA CONTA
                
                .txtstrContribuinte = txtProprietario
                .txtstrContribuinte1 = txtProprietario
                .txtstrLogradouroC = Field122
                .txtstrNumeroC = Field123
                .txtstrComplementoC = Field124
                .txtstrBairroC = Field125
                .txtstrMunicipioC = Field166
                .txtstrUFC = Field167
                .txtintCEPC = Field129
                .txtstrInscricao = txtInsc
                
                .txtdblValor1.OutputFormat = "#,##0.00"
                .txtdblValor2.OutputFormat = "#,##0.00"
                
                .blnPrimeira = True
                .blnValorEmReal = True
                
            End With
            
            Set subParcelas.object = rptCarneAcordoBoleto
             
     End If
  End If
  
End Sub

Private Function strQueryAtividades() As String
Dim strsql As String
  
  strsql = ""
  strsql = strsql & "SELECT "
  'strSql = strSql & "EC.strInscricaoCadastral, "
  strsql = strsql & "RTRIM(LTRIM(AEC.strDescricao)) "
  'strSql = strSql & "AE.blnPrincipal "
  strsql = strsql & "FROM "
  strsql = strsql & gstrAtividadeEC & " AEC, "
  strsql = strsql & gstrAtividadeDaEmpresa & " AE, "
  strsql = strsql & gstrEconomico & " EC "
  strsql = strsql & "WHERE "
  strsql = strsql & "EC.strInscricaoCadastral = '" & String(gintLenInscricao - Len(Trim(txtInscricao.Text)), "0") & Trim(txtInscricao.Text) & "' AND "
  strsql = strsql & "AE.intEconomico = EC.pkID AND "
  strsql = strsql & "AEC.pkID = AE.intAtividade "
  strsql = strsql & "ORDER BY "
  strsql = strsql & "AE.blnPrincipal DESC "
  
  strQueryAtividades = strsql
End Function

Private Function strQueryReceitas() As String
    Dim strsql As String
  
  strsql = ""
  strsql = strsql & "SELECT "
  strsql = strsql & "RTRIM(LTRIM(RE.strSigla)) strSigla, "
  strsql = strsql & "SUM(" & gstrISNULL("LR.dblValor", "0") & ") dblValor "
  strsql = strsql & "FROM "
  strsql = strsql & "tblReceita RE, "
  strsql = strsql & "tblLancamentoReceita LR, "
  strsql = strsql & "tblLancamentoValor LV "
  strsql = strsql & "WHERE "
  strsql = strsql & "LV.intLancamentoAlfa = " & txtpkID.Text & " AND "
  strsql = strsql & "LV.bItParcelaValida = 1 AND "
  strsql = strsql & "LR.intLancamentoValor = LV.pkID AND "
  strsql = strsql & "RE.pkID = LR.intReceita "
  strsql = strsql & "GROUP BY "
  strsql = strsql & "RE.pkID, "
  strsql = strsql & "RE.strSigla, "
  strsql = strsql & "LR.dblValor "
  strsql = strsql & "ORDER BY "
  strsql = strsql & "strSigla "

  strQueryReceitas = strsql
End Function

Private Function strQueryParcelas() As String
Dim strsql As String
Dim strBarra As String
Dim intCont As Integer
  
  strsql = ""
    
  strsql = strsql & "SELECT "
  strsql = strsql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao, "
  strsql = strsql & "LA.intExercicio intExercicio, LA.Pkid PkidAlfa, LA.strComposicaoDaReceita, "
  strsql = strsql & "CR.strSigla strSigla, CR.intUtilizacao intUtilizacao, LA.intComposicaoDaReceita intComposicao, "
  strsql = strsql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strAviso, 9999 strGuia, "
  strsql = strsql & "LA.strNomeProprietario strProprietario, "
    
  'Nº da Parcela
  strsql = strsql & "LV.intParcela intParcela, "
  strsql = strsql & "LV.pkID pkIDParcela, LV.bitParcelaValida,  "
  
  'Valor da Parcela
  strsql = strsql & "LV.dblValor dblValorParcela, "
  
  'Valor a ser gravado na Guia
  strsql = strsql & "LV.dblValor dblValorReal, "
   
    'Vencimento da Primeira Parcela
    If (bytDBType = EDatabases.Oracle) Then
        strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'DD/MM/YYYY'") & " dtmdtVencimento, "
    Else
        strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 103") & " dtmdtVencimento, "
    End If
      
    'Código de Barras===================================
    strsql = strsql & "'817' " & strCONCAT & " "
    
    'Valor da Parcela
    If (bytDBType = EDatabases.Oracle) Then
         strBarra = strSUBSTRING & "(LV.dblValor,0," & gstrINSTR(",", "LV.dblValor", 1, 1) & " + 4) * 100 "
    Else
         strBarra = " REPLACE(" & strSUBSTRING & "(" & gstrCONVERT(CDT_VARCHAR, "LV.dblValor") & ",0," & gstrINSTR(".", "LV.dblValor", 1, 1) & " + 5), '.', '') "
    End If
    
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
    
    'Guia
    strBarra = "9999"
    
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
  strsql = strsql & "CR.pkID = " & txtComposicao.Text & " AND "
  strsql = strsql & "CR.pkID " & strOUTJOracle & "= LA.intComposicaoDaReceita  AND "
  strsql = strsql & "LA.intExercicio = " & txtExercicio.Text & " AND "
  strsql = strsql & "LA.strEmissao = '" & txtEmissao.Text & "' AND "
  strsql = strsql & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(txtInscricao.Text)), "0") & Trim(txtInscricao.Text) & "' AND "
  strsql = strsql & "LA.dtmdtCancelamento IS NULL AND "
  strsql = strsql & "LV.intLancamentoAlfa = LA.pkID AND "
  strsql = strsql & "LV.intParcela IN (" & strParcelasSelecionadas & ") "
  If strNumeroAviso <> "" Then
    strsql = strsql & " AND " & gstrCONVERT(CDT_numeric, "LA.Strnumeroaviso") & " = " & Val(Trim(strNumeroAviso))
  End If
  strsql = strsql & " ORDER BY LV.intParcela "
  
  strQueryParcelas = strsql
End Function

