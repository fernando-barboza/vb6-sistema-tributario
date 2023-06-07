VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCapaCarneISSConstrucao 
   Caption         =   "Tributario - rptCapaCarneISSConstrucao (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptCapaCarneISSConstrucao.dsx":0000
End
Attribute VB_Name = "rptCapaCarneISSConstrucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strParcelasSelecionadas As String 'Parcelas selecionadas vindo do form
Dim blnSemIndexador As Boolean


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
  LeImagemLogotipo imgBrasao2, imgLogotipo, txtNomeFantasia2
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
    Dim adoResultado        As ADODB.Recordset
    Dim adoRelatorio        As ADODB.Recordset
    Dim intWhile            As Integer
    
    blnSemIndexador = False
    If IsNull(adoDataControl.Recordset!dblvlIndexador) Or IsNull(adoDataControl.Recordset!Strindexador) Then
       blnSemIndexador = True
    ElseIf Val(adoDataControl.Recordset!dblvlIndexador) = 0 Then
           blnSemIndexador = True
    End If
    
    txtImovel = gstrFormataInscricao(txtImovel, TYP_IMOBILIARIA)
    txtImovel1 = gstrFormataInscricao(txtImovel1, TYP_IMOBILIARIA)
    txtImovel2 = gstrFormataInscricao(txtImovel2, TYP_IMOBILIARIA)
    
    txtdtmdtvencimentoParcela = gstrDataFormatada(txtdtmdtvencimentoParcela)
    If Trim(txtdblVlIndexador) <> "" Then
        txtdblParcelaFMP.Text = Left(gstrConvVrDoSql(txtdbl1valor.Text, 2) / gstrConvVrDoSql(txtdblVlIndexador.Text, 6), InStr(1, gstrConvVrDoSql(txtdbl1valor.Text, 2) / gstrConvVrDoSql(txtdblVlIndexador.Text, 6), ",") + 4)
    End If
    If Trim(txtstrProcesso.Text) = "/-" Then
        txtstrProcesso.Text = ""
    End If
    intWhile = 1
    
    If gobjBanco.CriaADO(strQueryPredios, 10, adoResultado) Then
        With adoResultado
            If Not .EOF And Not .BOF Then
                Do While Not .EOF
                    If intWhile > 7 Then Exit Do
                    Select Case intWhile
                        Case 1
                            Construcao1.Text = gstrDataFormatada(gstrENulo(!Construcao))
                            TipoConstrucao1.Text = gstrENulo(!TipoConstrucao)
                            TipoAcabamento1.Text = gstrENulo(!TipoAcabamento)
                            Demolicao1.Text = gstrConvVrDoSql(gstrENulo(!Demolicao))
                            Lancada1.Text = gstrConvVrDoSql(gstrENulo(!Lancada))
                            ValorM21.Text = gstrConvVrDoSql(gstrENulo(!ValorM2))
                            Servico1.Text = gstrConvVrDoSql(gstrENulo(!Servico))
                            Aliquota1.Text = gstrConvVrDoSql(gstrENulo(!Aliquota))
                            Lancamento1.Text = gstrConvVrDoSql(gstrENulo(!Lancamento))
                            Abatido1.Text = gstrConvVrDoSql(gstrENulo(!Abatido))
                            dblPagar1.Text = gstrConvVrDoSql(gstrENulo(!dblPagar))
                        Case 2
                            Construcao2.Text = gstrDataFormatada(gstrENulo(!Construcao))
                            TipoConstrucao2.Text = gstrENulo(!TipoConstrucao)
                            TipoAcabamento2.Text = gstrENulo(!TipoAcabamento)
                            Demolicao2.Text = gstrConvVrDoSql(gstrENulo(!Demolicao))
                            Lancada2.Text = gstrConvVrDoSql(gstrENulo(!Lancada))
                            ValorM22.Text = gstrConvVrDoSql(gstrENulo(!ValorM2))
                            Servico2.Text = gstrConvVrDoSql(gstrENulo(!Servico))
                            Aliquota2.Text = gstrConvVrDoSql(gstrENulo(!Aliquota))
                            Lancamento2.Text = gstrConvVrDoSql(gstrENulo(!Lancamento))
                            Abatido2.Text = gstrConvVrDoSql(gstrENulo(!Abatido))
                            dblPagar2.Text = gstrConvVrDoSql(gstrENulo(!dblPagar))
                        Case 3
                            Construcao3.Text = gstrDataFormatada(gstrENulo(!Construcao))
                            TipoConstrucao3.Text = gstrENulo(!TipoConstrucao)
                            TipoAcabamento3.Text = gstrENulo(!TipoAcabamento)
                            Demolicao3.Text = gstrConvVrDoSql(gstrENulo(!Demolicao))
                            Lancada3.Text = gstrConvVrDoSql(gstrENulo(!Lancada))
                            ValorM23.Text = gstrConvVrDoSql(gstrENulo(!ValorM2))
                            Servico3.Text = gstrConvVrDoSql(gstrENulo(!Servico))
                            Aliquota3.Text = gstrConvVrDoSql(gstrENulo(!Aliquota))
                            Lancamento3.Text = gstrConvVrDoSql(gstrENulo(!Lancamento))
                            Abatido3.Text = gstrConvVrDoSql(gstrENulo(!Abatido))
                            dblPagar3.Text = gstrConvVrDoSql(gstrENulo(!dblPagar))
                        Case 4
                            Construcao4.Text = gstrDataFormatada(gstrENulo(!Construcao))
                            TipoConstrucao4.Text = gstrENulo(!TipoConstrucao)
                            TipoAcabamento4.Text = gstrENulo(!TipoAcabamento)
                            Demolicao4.Text = gstrConvVrDoSql(gstrENulo(!Demolicao))
                            Lancada4.Text = gstrConvVrDoSql(gstrENulo(!Lancada))
                            ValorM24.Text = gstrConvVrDoSql(gstrENulo(!ValorM2))
                            Servico4.Text = gstrConvVrDoSql(gstrENulo(!Servico))
                            Aliquota4.Text = gstrConvVrDoSql(gstrENulo(!Aliquota))
                            Lancamento4.Text = gstrConvVrDoSql(gstrENulo(!Lancamento))
                            Abatido4.Text = gstrConvVrDoSql(gstrENulo(!Abatido))
                            dblPagar4.Text = gstrConvVrDoSql(gstrENulo(!dblPagar))
                        Case 5
                            Construcao5.Text = gstrDataFormatada(gstrENulo(!Construcao))
                            TipoConstrucao5.Text = gstrENulo(!TipoConstrucao)
                            TipoAcabamento5.Text = gstrENulo(!TipoAcabamento)
                            Demolicao5.Text = gstrConvVrDoSql(gstrENulo(!Demolicao))
                            Lancada5.Text = gstrConvVrDoSql(gstrENulo(!Lancada))
                            ValorM25.Text = gstrConvVrDoSql(gstrENulo(!ValorM2))
                            Servico5.Text = gstrConvVrDoSql(gstrENulo(!Servico))
                            Aliquota5.Text = gstrConvVrDoSql(gstrENulo(!Aliquota))
                            Lancamento5.Text = gstrConvVrDoSql(gstrENulo(!Lancamento))
                            Abatido5.Text = gstrConvVrDoSql(gstrENulo(!Abatido))
                            dblPagar5.Text = gstrConvVrDoSql(gstrENulo(!dblPagar))
                        Case 6
                            Construcao6.Text = gstrDataFormatada(gstrENulo(!Construcao))
                            TipoConstrucao6.Text = gstrENulo(!TipoConstrucao)
                            TipoAcabamento6.Text = gstrENulo(!TipoAcabamento)
                            Demolicao6.Text = gstrConvVrDoSql(gstrENulo(!Demolicao))
                            Lancada6.Text = gstrConvVrDoSql(gstrENulo(!Lancada))
                            ValorM26.Text = gstrConvVrDoSql(gstrENulo(!ValorM2))
                            Servico6.Text = gstrConvVrDoSql(gstrENulo(!Servico))
                            Aliquota6.Text = gstrConvVrDoSql(gstrENulo(!Aliquota))
                            Lancamento6.Text = gstrConvVrDoSql(gstrENulo(!Lancamento))
                            Abatido6.Text = gstrConvVrDoSql(gstrENulo(!Abatido))
                            dblPagar6.Text = gstrConvVrDoSql(gstrENulo(!dblPagar))
                    End Select
                    intWhile = intWhile + 1
                    .MoveNext
                Loop
            End If
        End With
    End If
    
        If Not adoDataControl.Recordset.EOF Then
           Set adoRelatorio = New ADODB.Recordset
           With rptCarneParcelas
             If gobjBanco.CriaADO(strQueryParcelas, 5, adoRelatorio) Then
                 If bytDBType = EDatabases.SQLServer Then
                    .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                 Else
                    .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                 End If
                 Set .adoDataControl.Recordset = adoRelatorio
             End If
             .lblTitulo.Caption = "S.M.F. - Seção de Rendas Mobiliárias"
             .lblTitulo1.Caption = "S.M.F. - Seção de Rendas Mobiliárias"
             .lblTipo.Caption = "Imóvel:"
             .lblTipo1.Caption = "Imóvel:"
             .lblProprietario.Caption = "Proprietário"
             .lblProprietario1.Caption = "Proprietário"
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
            
             .blnPrimeira = False
             
           End With
           Set subParcelas.object = rptCarneParcelas
        End If
End Sub

Private Function strQueryParcelas() As String
    Dim strsql As String
    Dim strBarra As String
    Dim intCont As Integer
    
    strsql = ""
    If bytDBType = Oracle Then 'ORACLE
        strsql = strsql & "SELECT "
        strsql = strsql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
        strsql = strsql & "LA.intExercicio intExercicio, "
        strsql = strsql & "LA.strComposicaoDaReceita strComposicaoDaReceita, "
        strsql = strsql & "LV.bitParcelaValida, "
        strsql = strsql & "LA.pkid PkidAlfa, "
        strsql = strsql & "CR.strSigla strSigla, CR.intUtilizacao intUtilizacao, LA.intComposicaoDaReceita intComposicao, "
        strsql = strsql & "LA.strNumeroAviso strAviso, 9999 strGuia, "
        strsql = strsql & "LA.strNomeProprietario strProprietario, "
        'strSql = strSql & "'R$' strIndexador, "
        
        strsql = strsql & "LA.dblvlIndexador, "
        strsql = strsql & "LA.strIndexador, "
        
        'Nº da Parcela
        strsql = strsql & "LV.intParcela intParcela, "
        strsql = strsql & "LV.pkID pkIDParcela, "
        
        'Valor da Parcela
        If blnSemIndexador Then
           strsql = strsql & "LV.dblValor "
        Else
           strsql = strsql & "SUBSTR(LV.dblValor / LA.dblvlIndexador, 0,INSTR(LV.dblValor / LA.dblvlIndexador,',',1,1) + 4) "
        End If
        strsql = strsql & "dblValorParcela, "
        
        'Valor a ser gravado na Guia
        strsql = strsql & "LV.dblValor dblValorReal, "
        
        'Vencimento da Primeira Parcela
        strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'DD/MM/YYYY'") & " dtmdtVencimento, "
        
        'Código de Barras
        strsql = strsql & "'817' " & strCONCAT & " "
        
        'Valor da Parcela
        If blnSemIndexador Then
           'strBarra = "SUBSTR(LV.dblValor,0,INSTR(LV.dblValor,',',1,1) + 4) * 100 "
           strBarra = "LV.dblValor * 100 "
        Else
           strBarra = "SUBSTR(LV.dblValor / LA.dblvlIndexador, 0,INSTR(LV.dblValor / LA.dblvlIndexador,',',1,1) + 4) * 10000 "
        End If
        strsql = strsql & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & " , '00000000000'") & ") " & strCONCAT & " "
        
        'Febraban
        strBarra = "(SELECT intFebraban FROM " & gstrEmpresa & ")"
        strsql = strsql & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & ", '0000'") & ") " & strCONCAT & " "
        
        'Vencimento
        'strSQL = strSQL & "REPLACE( "
        'strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'DD/MM/YYYY'")
        'strSQL = strSQL & ", '/', '') " & strCONCAT & " "
        strsql = strsql & gstrCONVERT(CDT_VARCHAR, "LV.dtmdtVencimento, 'YYYYMMDD'") & strCONCAT & " "
        
        'Conta Bancária
        strsql = strsql & "'0000' " & strCONCAT & " "
        
        'Guia
        strBarra = "9999"
        
        strsql = strsql & "LTRIM(" & gstrCONVERT(CDT_VARCHAR, strBarra & ", '000000000'") & ") " & strCONCAT & " "
        
        'Exercicio
        strsql = strsql & "LA.intExercicio "
        
        strsql = strsql & "strCodigoBarra "
    Else 'SQL
        strsql = strsql & "SELECT "
        strsql = strsql & " LA.strInscricao strInscricao,   LA.intExercicio intExercicio,   CR.strSigla strSigla,   CR.intUtilizacao intUtilizacao, "
        strsql = strsql & " LA.intComposicaoDaReceita intComposicao,    LA.strNumeroAviso strAviso,     9999 strGuia,   LA.strNomeProprietario strProprietario, "
        strsql = strsql & "LA.dblvlIndexador, LA.strIndexador, "
        strsql = strsql & " LV.intParcela intParcela,   LV.pkID pkIDParcela, "
        
        'Valor da Parcela
        If blnSemIndexador Then
           strsql = strsql & "LV.dblValor "
        Else
           strsql = strsql & "SUBSTRING(LV.dblValor / LA.dblvlIndexador, 0,CHARINDEX(',',LV.dblValor / LA.dblvlIndexador) + 4) "
        End If
        strsql = strsql & "dblValorParcela, "
        
        'Valor a ser gravado na Guia
        strsql = strsql & " LV.dblValor dblValorReal , "
        
        'Vencimento da Primeira Parcela
        strsql = strsql & " REPLACE(CONVERT (VARCHAR, LV.dtmdtVencimento, 102),'.','') dtmdtVencimento, "
        
        'Código de Barras
        strsql = strsql & " CONVERT(varchar,'817')  + Replicate('0',11-len(REPLACE(RTrim(Ltrim(isnull(LV.dblValor,0))), '.','')) + "
        
        'Valor da Parcela
        If blnSemIndexador Then
           strBarra = "LV.dblValor * 100 "
        Else
           strBarra = "SUBSTRING(LV.dblValor / LA.dblvlIndexador, 0,CHARINDEX(',',LV.dblValor / LA.dblvlIndexador) + 4) * 10000 "
        End If
        strsql = strsql & "CONVERT(VARCHAR,Replicate('0',11-len(REPLACE(RTrim(Ltrim(isnull(" & strBarra & ",0))), '.','')))) + "
        strsql = strsql & strBarra & " + "
        
        strsql = strsql & " REPLACE(RTrim(Ltrim(isnull(" & strBarra & ",0))), '.','') + "
        
        strsql = strsql & " Replicate('0',4-len(RTrim(Ltrim((SELECT isnull(intFebraban,0) FROM tblEmpresa))))) + "
        strsql = strsql & " CONVERT(VARCHAR,RTrim(Ltrim((SELECT isnull(intFebraban,0) FROM tblEmpresa)))) + "
        strsql = strsql & " REPLACE( CONVERT (VARCHAR, LV.dtmdtVencimento, 103) , '/', '') "
        strsql = strsql & " + '0000' + "
        strsql = strsql & " '000009999' + Convert(varchar,LA.intExercicio)) strCodigoBarra "
    End If
    
    'FROM
    strsql = strsql & "FROM "
    strsql = strsql & gstrLancamentoAlfa & " LA, "
    strsql = strsql & gstrComposicaoDaReceita & " CR, "
    strsql = strsql & gstrLancamentoValor & " LV "
    
    'WHERE
    strsql = strsql & "WHERE "
    strsql = strsql & "LA.intComposicaoDaReceita = CR.Pkid AND "
    strsql = strsql & "LA.dtmdtCancelamento IS NULL AND "
    strsql = strsql & "LV.intLancamentoAlfa = LA.pkID AND "
    strsql = strsql & "LA.PkId = " & txtIntLancamentoAlfa & " "
    strsql = strsql & "ORDER BY LV.intParcela "
    
    strQueryParcelas = strsql
End Function

Private Function strQueryPredios() As String
    Dim strsql As String
    
    strsql = strsql & "Select "
    strsql = strsql & "LIP.DTMDATACONSTRUCAO as Construcao, "
    strsql = strsql & "LIP.STRTIPOCONSTRUCAO as TipoConstrucao, "
    strsql = strsql & "LIP.Strtipoacabamento as TipoAcabamento, "
    strsql = strsql & "LIP.DBLPORCDEMOLICAO as Demolicao, "
    strsql = strsql & "LIP.Dblarealancada as Lancada, "
    strsql = strsql & "LIP.Dblvalorm2 as ValorM2, "
    strsql = strsql & "LIP.Dblvalorservico as Servico, "
    strsql = strsql & "LIP.Dblaliquotaiss as Aliquota, "
    strsql = strsql & "LIP.Dblvalorlancto as Lancamento, "
    strsql = strsql & "LIP.DBLVALORABATIDO as Abatido, "
    'strSql = strSql & "LIP.dblvalorlancto  - LIP.dblvalorabatido as dblPagar "
    strsql = strsql & "((CASE WHEN LIP.dblValorLancto IS NULL THEN 0 ELSE LIP.dblValorLancto END) - "
    strsql = strsql & "(CASE WHEN LIP.dblValorAbatido IS NULL THEN 0 ELSE LIP.dblValorAbatido END)) dblPagar "
    strsql = strsql & "From "
    strsql = strsql & gstrLanctoIssConstrucaoPredios & " LIP "
    strsql = strsql & "Where "
    strsql = strsql & "Intlanctoissconstrucao = " & Val(txtintLanctISS)
    strQueryPredios = strsql
End Function

