VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptOrdemDePagamento 
   Caption         =   "prjOrcamentario - rptOrdemDePagamento (ActiveReport)"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   465
   ClientWidth     =   11400
   MDIChild        =   -1  'True
   _ExtentX        =   20108
   _ExtentY        =   19606
   SectionData     =   "rptOrdemDePagamento.dsx":0000
End
Attribute VB_Name = "rptOrdemDePagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Dim dblVarTotal As Double
Dim lngPkid                 As Long
Dim bytTipo                 As Byte
Dim blnPago                 As Boolean
Public intPagina            As Integer
Public blnAnulacaoReceita   As Boolean
Dim blnGrupoAberto          As Boolean
Public intOrigem            As Integer
Private intFolha            As Integer
Private lngNumero           As Long

Private Sub ActiveReport_Activate()
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

Private Sub ActiveReport_ReportStart()
Dim strSql As String
Dim adoEmpresa As ADODB.Recordset
Dim imgBrazao As Image
    
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia
    lngNumero = 0
    
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
    Dim strContaContabil As String
    Dim strDescricao As String
    Dim dblValor As Double
    Dim dblValorDesconto As Double
    
   If adoDataControl.Recordset.EOF Or adoDataControl.Recordset.BOF Then
        lngPkid = 0
        Exit Sub
   End If
    
'    If adoDataControl.NRecords > 0 Then
       strEmpenho = adoDataControl.Recordset("intEmpenho").Value & "/" & Format(adoDataControl.Recordset("intParcela").Value, "00") _
                                                                 & "." & Mid(adoDataControl.Recordset("intExercicio").Value, 3, 2)
       dblValorLiquidado = gstrConvVrDoSql(adoDataControl.Recordset("dblValorParcela").Value)
       
       dblVarTotal = ((dblVarTotal + adoDataControl.Recordset("dblValorParcela").Value) - adoDataControl.Recordset("dblDesconto").Value)
       
       intVinc = Right(adoDataControl.Recordset("strDotacao").Value, 4)
       intUnidade = Mid(adoDataControl.Recordset("strDotacao").Value, 4, 4)
    
        'M4R CALCULA VALOR REAL DO EMPENHO
       dblValorDoEmpenho = gstrConvVrDoSql(adoDataControl.Recordset("dblValorEmpenho") - ValorEmpenho(Val(adoDataControl.Recordset("PkidEmpenho"))), 2)
       dblValorDoEmpenho = gstrConvVrDoSql(dblValorDoEmpenho, 2)
       dblSaldoDoEmpenho = gstrConvVrDoSql(Val(gstrConvVrParaSql(dblValorDoEmpenho)) - Val(gstrConvVrParaSql(dblValorLiquidado)))
       
'    End If
    
    strEmpenho.Width = 926
    Line36.X1 = 1063
    Line36.X2 = 1063
    lblTextoExtra.Width = 2623


    If bytTipo = 0 Or bytTipo = 1 Then
        txt_Descricao.Visible = False
        lblTextoExtra.Visible = False
        
        strDotacao.Visible = True
        strEmpenho.Visible = True
        intVinc.Visible = True
        intUnidade.Visible = True
        dblValorDoEmpenho.Visible = True
        dblSaldoDoEmpenho.Visible = True

        
    ElseIf bytTipo = 2 Then
        gstrDespExtra adoDataControl.Recordset!PKIDDespExtra, strContaContabil, strDescricao, dblValor, dblValorDesconto
        strEmpenho = gvntFormatacaoEspecifica(strContaContabil, 1)
        lblTextoExtra = strDescricao
        dblValorLiquidado = gstrConvVrDoSql(dblValor) ' - dblValorDesconto)
        
        lblTextoExtra.Visible = True
        txt_Descricao.Visible = False
        strDotacao.Visible = False
        strEmpenho.Visible = True
        intVinc.Visible = False
        intUnidade.Visible = False
        dblValorDoEmpenho.Visible = False
        dblSaldoDoEmpenho.Visible = False
        strEmpenho.Width = 1120
        Line36.X1 = strEmpenho.Width + strEmpenho.Left + 30
        Line36.X2 = Line36.X1
        lblTextoExtra.Left = Line36.X1 + 20
        lblTextoExtra.Width = Line41.X1 - Line36.X1 - 30

    
    Else

        strDotacao.Visible = False
        strEmpenho.Visible = False
        intVinc.Visible = False
        intUnidade.Visible = False
        dblValorDoEmpenho.Visible = False
        dblSaldoDoEmpenho.Visible = False
        If blnAnulacaoReceita Then
            txt_Descricao.Visible = True
            lblTextoExtra.Visible = False
            'dblTotal = gstrConvVrDoSql(CDbl(dblTotal) + CDbl(txtValorLiquidado))
            dblTotal = gstrConvVrDoSql(Val(gstrConvVrParaSql(dblTotal)) + Val(gstrConvVrParaSql(txtValorLiquidado)))
        Else
            txt_Descricao.Visible = False
            lblTextoExtra.Visible = True
        End If
    End If
   
End Sub

Private Sub gstrDespExtra(ByVal PKIDDespExtra As String, _
                          ByRef strContaContabil As String, _
                          ByRef strDescricao As String, _
                          ByRef dblValor As Double, _
                          ByRef dblValorDesconto As Double)
                          
    Dim strSql As String
    Dim adoResultado  As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "SELECT DE.intnumero , "
    strSql = strSql & "PC.strcontacontabil ,"
    strSql = strSql & "PC.strdescricao, "
    strSql = strSql & "DE.dblvalor ,"
    strSql = strSql & "DE.dbldesconto "
    strSql = strSql & "FROM "
    strSql = strSql & gstrDespesaExtraOrcamentaria & " DE,"
    strSql = strSql & gstrPlanoConta & " PC "
    strSql = strSql & " WHERE De.Intcontacontabil = PC.PKID "
    strSql = strSql & " AND DE.PKID = " & PKIDDespExtra
    
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
        If Not adoResultado.EOF Then
            strContaContabil = gstrENulo(adoResultado!strContaContabil)
            strDescricao = gstrENulo(adoResultado!strDescricao)
            dblValor = Val(gstrConvVrParaSql(gstrENulo(adoResultado!dblValor)))
            dblValorDesconto = Val(gstrENulo(adoResultado!dblDesconto))
        End If
    End If

End Sub


Private Sub GroupFooter1_Format()

    Dim strSql As String
    Dim adoRelatorio As ADODB.Recordset
    Set adoRelatorio = Nothing
   
    
    ImprimeSub
    rptDescontosDeOPs.lblPagina = "Folha : " & intFolha + 1
    intPagina = 0
   
   
'    If blnPago = True Then Exit Sub

'   strSql = " SELECT "
'   strSql = strSql & "0 intEmpenhoAnulacao, "
'   strSql = strSql & "5 intTipo, " 'Fianças Bancarias'
'   strSql = strSql & "'Fianças Bancárias' strTipo, "
'   strSql = strSql & "0 intNumeroOP, "
'   strSql = strSql & "0 intArtigo, "
'   strSql = strSql & "0 bytEstorno, "
'   strSql = strSql & "CF.dtmDataSaida dtmPagamento, "
'   strSql = strSql & "CF.dblValor, "
'   strSql = strSql & "CT.CDC intCodigoContribuinte, "
'   strSql = strSql & "CT.strNome "
'   strSql = strSql & "FROM "
'   strSql = strSql & gstrCartaFianca & " CF, "
'   strSql = strSql & gstrContribuinte & " CT "
'   strSql = strSql & "WHERE "
'   strSql = strSql & "(CF.intExcluido = 0 OR CF.intExcluido IS NULL) AND "
'   strSql = strSql & "CF.intContribuinte = CT.PKID AND "
'   strSql = strSql & "CF.dtmDataSaida  = " & gstrConvDtParaSql("18/02/2004") & " "

'    With rptSubOrdemPagamento

'    If gobjBanco.CriaADO(strSql, 5, adoRelatorio) Then
'        If bytDBType = EDatabases.SQLServer Then
'            .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
'        Else
'            .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
'        End If
'        Set .adoDataControl.Recordset = adoRelatorio
'    End If

'    End With

'    Set SubOrdemPagamento.object = rptSubOrdemPagamento

End Sub

Private Sub GroupFooter2_Format()
    blnGrupoAberto = False
End Sub

Private Sub GroupHeader1_Format()
   Dim strEspaco As String
   
   strEspaco = Space$(50)
   
   If adoDataControl.Recordset.EOF Or adoDataControl.Recordset.BOF Then
        lngPkid = 0
        Exit Sub
   End If
   
   
    lblEmpenho.Caption = "Empenho"
    lblDotacao.Caption = "Dotação"
    

   blnAnulacaoReceita = False
   If adoDataControl.Recordset("bytTipo").Value = 0 Then
      lblTpOrdem = "O . P .  Orçamentária"
      bytTipo = 0
   ElseIf adoDataControl.Recordset("bytTipo") = 1 Then
      lblTpOrdem = "O . P .  de Restos a Pagar"
      bytTipo = 1
   ElseIf adoDataControl.Recordset("bytTipo") = 2 Then
      lblTpOrdem = "O . P .  Extra - Orçamentária"
      bytTipo = 2
      lblEmpenho.Caption = "Conta"
      lblDotacao.Caption = "Extra-Orçamentária"
   ElseIf adoDataControl.Recordset("bytTipo") = 3 Then
      lblTpOrdem = "O . P .  Anulação de Receita "
      blnAnulacaoReceita = True
      bytTipo = 3
   End If
   
   intPagina = intPagina + 1
   'lblPagina = "Folha : " & pageNumber
   'lblPagina = "Folha : " & intPagina
   
    If lngNumero <> adoDataControl.Recordset("intOrdem").Value Then
        intFolha = 0
    End If
   lngNumero = adoDataControl.Recordset("intOrdem").Value
   
   intFolha = intFolha + 1
   lblPagina = "Folha : " & intFolha
   
   lngPkid = adoDataControl.Recordset!Pkid
   
   lblNumeroOrdem = "Número : " & adoDataControl.Recordset("intOrdem").Value & "/" & adoDataControl.Recordset("intexercicioOP").Value
   lblData = "Data : " & gstrDataFormatada(adoDataControl.Recordset("dtmData").Value)
   txtintContribuinte = adoDataControl.Recordset("intContribuinte").Value & " - " & adoDataControl.Recordset("strNome").Value
   blnPago = IIf(adoDataControl.Recordset("strFonteRecurso").Value = 1, True, False)
   
   If Len(adoDataControl.Recordset("strCodigo").Value) > 0 Then
      txtstrProcesso = gstrENulo(adoDataControl.Recordset("strCodigo").Value) & "/" & gstrENulo(adoDataControl.Recordset("intExercicioProcesso").Value) & " - " & gstrENulo(adoDataControl.Recordset("bitDigito").Value)
   End If
   
   dtmDataVencimento = gstrDataFormatada(dtmDataVencimento)
   If adoDataControl.Recordset("bytNaturezaJuridica").Value = 1 Then
        lblCGCCPF = "CNPJ:"
        txtstrCNPJCPF = gstrCGCCPFFormatado(txtstrCNPJCPF, "PJ")
   Else
        lblCGCCPF = "CPF:"
        txtstrCNPJCPF = gstrCGCCPFFormatado(txtstrCNPJCPF, "PF")
   End If
   
   If Val(txtstrCNPJCPF) = 0 Then
        txtstrCNPJCPF = ""
   End If
   
   lblHistorico = IIf(Not IsNull(adoDataControl.Recordset("typHistorico").Value), adoDataControl.Recordset("typHistorico").Value, "")
   lblRecurso = adoDataControl.Recordset("strFonteRecurso").Value
   
    Select Case intOrigem
        Case 1
            lblDesconto = "Total de Desconto : " & gstrConvVrDoSql(adoDataControl.Recordset("dblDesconto").Value)
            dblTotal = gstrConvVrDoSql(dblTotal)
        Case 2
            'dblTotal = gstrConvVrDoSql(Val(gstrConvVrParaSql(dblTotal)) - Val(gstrConvVrParaSql(ValorDesconto)))
            dblTotal = gstrConvVrDoSql(Val(gstrConvVrParaSql(adoDataControl.Recordset("dblLiquidoTotal").Value)) - Val(gstrConvVrParaSql(ValorDesconto)))
            lblDesconto = "Total de Desconto : " & gstrConvVrDoSql(ValorDesconto)
    End Select
   
   
   
   
   
   dblDesconto = gstrConvVrDoSql(dblDesconto)
   
   
   
   
   
   If dblTotal = "" Then dblTotal = 0
   lblExtenso = "***** " & gstrExtenso(gstrConvVrDoSql(dblTotal)) & " *****"
   GroupHeader1.Repeat = ddRepeatOnPage
   
   If blnAnulacaoReceita Then
       dblTotal = gstrConvVrDoSql(adoDataControl.Recordset("dblValorTotal").Value)
   End If
End Sub

Private Sub GroupHeader2_Format()
    blnGrupoAberto = True
End Sub

Private Sub PageFooter_Format()

    'If Not adoDataControl.Recordset.EOF Then
    '    Line52.Visible = True
    'Else
    '    Line52.Visible = False
    'End If
    Line52.Y1 = 0
    Line52.Y2 = 0
    
    If blnGrupoAberto Then
        Line52.Visible = True
    Else
        Line52.Visible = False
    End If
    
     Select Case CStr(gstrRetSiglaPref)
            Case "GRJ"
     
                'Caso for Guarujá muda o rodapé
                'Campos rodapé normal
                Shape4.Visible = False
                lblEmitente.Visible = False
                Line31.Visible = False
                lblConferente.Visible = False
                Line29.Visible = False
                Line32.Visible = False
                lblOrdernador.Visible = False
                Line30.Visible = False
                lblRecibo.Visible = False
                Line33.Visible = False
                lblIdentidade.Visible = False
        
                'Campos rodapé especifico de Guarujá
    
                Shape5G.Visible = True
                lblEmitenteG.Visible = True
                lblBanco.Visible = True
                lblCheque.Visible = True
                Line60.Visible = True
                Line62.Visible = True
                lblOrdenadorG.Visible = True
                Line61.Visible = True
                lblReciboG.Visible = True
                Label1.Visible = True
                Label2.Visible = True
                Label3.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                Label6.Visible = True
                lblIdentidadeG.Visible = True
                Line64.Visible = True
                PageFooter.Height = 4770
                imgLogotipo.Top = 3964
                imgLogotipo.Left = 9354
                
            Case Else
     End Select
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
    
    
     Select Case CStr(gstrRetSiglaPref)
     Case "MAUA"
        lblsubTitulo = "SECRETARIA MUNICIPAL DE FINANÇAS - DIVISÃO DE CONTABILIDADE"
     Case "PUBT"
        lblsubTitulo = "DIRETORIA ADMINISTRATIVA FINANCEIRA - DEPARTAMENTO DE CONTABILIDADE"
     Case Else
        lblsubTitulo = ""
     End Select
     
    
    
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

Private Function ValorDesconto() As Double
    Dim adoResultado   As ADODB.Recordset
    Dim strSql         As String
    
    'strSql = "SELECT SUM(TMP.dblvalor) dblvalor FROM (" & strQueryDescontoOPs & ")TMP"
    strSql = strQueryDescontoOPs
    
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        While Not adoResultado.EOF
            ValorDesconto = ValorDesconto + Val(gstrConvVrParaSql(gstrENulo(adoResultado!dblValor)))
            adoResultado.MoveNext
        Wend
    End If
    
    ValorDesconto = gstrConvVrDoSql(ValorDesconto, 2)
    
    strQueryDescontoOPs
End Function

Private Function ValorEmpenho(lngPkidEmpenho) As Double
Dim strSql         As String
Dim strSqlAux      As String
Dim adoResultado   As ADODB.Recordset
Dim adoResultado2  As ADODB.Recordset
Dim lngPkidParcela As Long


If bytTipo = 0 Or bytTipo = 1 Then
    
        strSql = ""
        strSql = "SELECT Pkid FROM " & gstrSubempenho
        strSql = strSql & " WHERE intNumero = " & Val(adoDataControl.Recordset("intParcela"))
        strSql = strSql & " AND intEmpenho = " & lngPkidEmpenho
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSql, 5, adoResultado2) Then
           If Not adoResultado2.EOF Then
              lngPkidParcela = adoResultado2!Pkid
           End If
        End If
    
    If Val(adoDataControl.Recordset("intParcela")) = 0 Then 'Parcelas 0
    
        strSql = ""
        strSql = "SELECT SUM(dblValor) + ("
        
        strSqlAux = ""
        strSqlAux = "(SELECT SUM(dblValor) ValorEmpenho"
        strSqlAux = strSqlAux & " FROM "
        strSqlAux = strSqlAux & gstrSubempenho
        strSqlAux = strSqlAux & " WHERE intEmpenho =" & lngPkidEmpenho & " AND"
        strSqlAux = strSqlAux & " intNumero = 0 AND"
        strSqlAux = strSqlAux & " bytSituacao = 4"
        strSqlAux = strSqlAux & " GROUP BY intEmpenho)"
        
        strSql = strSql & gstrISNULL(strSqlAux, "0") & " ) ValorEmpenho"
        strSql = strSql & " FROM "
        strSql = strSql & gstrSubempenho
        strSql = strSql & " WHERE intEmpenho =" & lngPkidEmpenho & " AND"
        strSql = strSql & " intNumero BETWEEN 0 AND " & Val(adoDataControl.Recordset("intParcela") - 1)
        strSql = strSql & " AND bytSituacao <> 4 "
        strSql = strSql & " GROUP BY intEmpenho"
        
    End If
        
    If Val(adoDataControl.Recordset("intParcela")) > 0 Then 'Parcelas > 0
    
        strSql = ""
        strSql = " SELECT " & gstrISNULL("SUM(DED.VALOREMPENHO)", "0") & " ValorEmpenho FROM"
        strSql = strSql & " (SELECT "
        strSql = strSql & gstrISNULL("SUM(dblValor)", " 0") & " ValorEmpenho"
        strSql = strSql & " FROM "
        strSql = strSql & gstrSubempenho
        strSql = strSql & " WHERE"
        strSql = strSql & " intEmpenho = " & lngPkidEmpenho
        strSql = strSql & " AND intNumero = 0"
        strSql = strSql & " AND bytSituacao = 4"
        strSql = strSql & " AND Pkid < " & lngPkidParcela
        strSql = strSql & " UNION ALL"
        strSql = strSql & " SELECT "
        strSql = strSql & gstrISNULL("SUM(dblValor)", " 0") & " ValorEmpenho"
        strSql = strSql & " FROM "
        strSql = strSql & gstrSubempenho
        strSql = strSql & " WHERE"
        strSql = strSql & " intEmpenho = " & lngPkidEmpenho
        strSql = strSql & " AND intNumero > 0"
        strSql = strSql & " AND bytSituacao <> 4"
        strSql = strSql & " AND Pkid < " & lngPkidParcela & " ) DED"
        
    End If
Else
    strSql = "SELECT SUM(dblValor) ValorEmpenho"
    strSql = strSql & " FROM "
    strSql = strSql & gstrSubempenho
    strSql = strSql & " WHERE intEmpenho =" & lngPkidEmpenho & " AND"
    strSql = strSql & " intNumero BETWEEN 1 AND " & Val(adoDataControl.Recordset("intParcela") - 1)
    strSql = strSql & " AND bytSituacao <> 4 "
    strSql = strSql & " GROUP BY intEmpenho"
End If
    
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            ValorEmpenho = gstrConvVrDoSql(adoResultado!ValorEmpenho, 2)
        Else
            ValorEmpenho = gstrConvVrDoSql(0, 2)
        End If
    End If

End Function

Private Function strQueryDescontoOPs() As String

Dim strSql As String
    
  If bytTipo = 0 Or bytTipo = 1 Then
        If Not BlnDescontoOP(lngPkid) Then
            strSql = "SELECT ALL 'Orçamentário' quebra, ' ' strCodigo, " & _
                "SUM(Dbldesconto) Dblvalor, " & _
                "'Total Descontado' strDescricaoConta, " & _
                "OP.bytTipo, " & _
                "OP.INTNUMERO intOrdem, " & _
                "OP.intExercicio IntExercicioOP, " & _
                "OP.DTMDATA , " & _
                gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_VARCHAR, "CT.CDC") & strCONCAT & " ' - ' ") & strCONCAT & " CT.STRNOME intContribuinte, " & _
                "CT.strCNPJCPF, " & _
                "CT.strLogradouroC STRENDERECO, " & _
                "CT.intNumero, " & _
                "CT.STRCOMPLEMENTO STRCOMPLEMENTO, " & _
                "MP.strDescricao STRMUNICIPIO, " & _
                "UF.strsigla STRUF, " & _
                "BR.strDescricao STRBAIRRO, " & _
                "CP.INTCEP "
                strSql = strSql & "FROM " & gstrOrdemPagamento & " OP, " & _
                gstrOrdemPagamentoEmpenho & " OPE, " & _
                gstrSubempenho & " SEP, " & _
                gstrContribuinte & " CT, " & _
                gstrCidade & " MP, " & _
                gstrUF & " UF, " & _
                gstrBairro & " BR, " & _
                gstrCeps & " CP "
            
            strSql = strSql & "WHERE OPE.Intordempagamento = OP.Pkid " & _
                " AND OPE.Intparcela = SEP.pkid " & _
                " AND OP.intContribuinte = CT.Pkid " & _
                " AND MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio " & _
                " AND CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep " & _
                " AND BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro " & _
                " AND UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF " & _
                " AND OP.Pkid = " & lngPkid
             
            strSql = strSql & " Group by op.bytTipo, OP.INTNUMERO , " & gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_VARCHAR, "CT.CDC ") & strCONCAT & " ' - ' ") & strCONCAT & " CT.STRNOME , OP.DTMDATA , CT.strCNPJCPF, CT.strLogradouroC , CT.intNumero, CT.STRCOMPLEMENTO ,MP.strDescricao ," & _
             "UF.strsigla , BR.strDescricao, CP.INTCEP,OP.intExercicio"
            strSql = strSql & " HAVING " & gstrISNULL("Sum(dblDesconto)", "0") & " > 0 "
            
            strSql = strSql & " UNION  ALL  "
            
            
            strSql = strSql & "SELECT  ALL  'Orçamentário' quebra, ' ' strCodigo, " & _
                 "SUM(Dbldesconto) Dblvalor, " & _
                 "'Total Descontado' strDescricaoConta, " & _
                 "OP.bytTipo, " & _
                 "OP.INTNUMERO intOrdem, " & _
                 "OP.intExercicio IntExercicioOP, " & _
                 "OP.DTMDATA , " & _
                 gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_VARCHAR, " CT.CDC ") & strCONCAT & " ' - ' ") & strCONCAT & " CT.STRNOME intContribuinte, " & _
                 "CT.strCNPJCPF, " & _
                 "CT.strLogradouroC STRENDERECO, " & _
                 "CT.intNumero, " & _
                 "CT.STRCOMPLEMENTO STRCOMPLEMENTO, " & _
                 "MP.strDescricao STRMUNICIPIO, " & _
                 "UF.strsigla STRUF, " & _
                 "BR.strDescricao STRBAIRRO, " & _
                 "CP.INTCEP "
                strSql = strSql & "FROM " & gstrOrdemPagamento & " OP, " & _
                 gstrOrdemPagamentoResto & " OPR, " & _
                 gstrSubempenho & " SEP, " & _
                 gstrContribuinte & " CT, " & _
                 gstrCidade & " MP, " & _
                 gstrUF & " UF, " & _
                 gstrBairro & " BR, " & _
                 gstrCeps & " CP "
                strSql = strSql & "WHERE OPR.Intordempagamento = OP.Pkid " & _
                 " AND OPR.Intparcela = SEP.pkid " & _
                 " AND OP.intContribuinte = CT.Pkid " & _
                 " AND MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio " & _
                 " AND CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep " & _
                 " AND BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro " & _
                 " AND UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF " & _
                 " AND OP.Pkid = " & lngPkid
             
            strSql = strSql & " Group by op.bytTipo, OP.INTNUMERO , " & gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_VARCHAR, "CT.CDC ") & strCONCAT & " ' - ' ") & strCONCAT & " CT.STRNOME , OP.DTMDATA , CT.strCNPJCPF, CT.strLogradouroC , CT.intNumero, CT.STRCOMPLEMENTO ,MP.strDescricao ," & _
             "UF.strsigla , BR.strDescricao, CP.INTCEP,OP.intExercicio"
             
            strSql = strSql & " HAVING " & gstrISNULL("Sum(dblDesconto)", "0") & " > 0 "
        Else
            strSql = strSql & "Select quebra, strCodigo, "
            strSql = strSql & "Sum(Dblvalor) Dblvalor, "
            strSql = strSql & "strDescricaoConta, "
            strSql = strSql & "bytTipo, "
            strSql = strSql & "intOrdem, "
            strSql = strSql & "IntExercicioOP, "
            strSql = strSql & "DTMDATA, "
            strSql = strSql & "intContribuinte, "
            strSql = strSql & "strCNPJCPF, "
            strSql = strSql & "STRENDERECO, "
            strSql = strSql & "intNumero, "
            strSql = strSql & "STRCOMPLEMENTO, "
            strSql = strSql & "STRMUNICIPIO, "
            strSql = strSql & "STRUF, "
            strSql = strSql & "STRBAIRRO, "
            strSql = strSql & "INTCEP From ( "
        
            strSql = strSql & "SELECT ALL 'Orçamentário' quebra, CO.Strcodigoorcamentario StrCodigo, " & _
                gstrISNULL("SEPR.Dblvalor", 0) & " Dblvalor, " & _
                "CO.Strdescricao strDescricaoConta, " & _
                "OP.bytTipo, " & _
                "OP.INTNUMERO intOrdem, " & _
                "OP.intExercicio IntExercicioOP, " & _
                "OP.DTMDATA , " & _
                gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_VARCHAR, "CT.CDC") & strCONCAT & " ' - ' ") & strCONCAT & " CT.STRNOME intContribuinte, " & _
                "CT.strCNPJCPF, " & _
                "CT.strLogradouroC STRENDERECO, " & _
                "CT.intNumero, " & _
                "CT.STRCOMPLEMENTO STRCOMPLEMENTO, " & _
                "MP.strDescricao STRMUNICIPIO, " & _
                "UF.strsigla STRUF, " & _
                "BR.strDescricao STRBAIRRO, " & _
                "CP.INTCEP "
                strSql = strSql & "FROM " & gstrOrdemPagamento & " OP, " & _
                gstrOrdemPagamentoEmpenho & " OPE, " & _
                gstrSubempenho & " SEP, " & _
                gstrSubEmpRetencaoOrcamentaria & " SEPR, " & _
                gstrCodigoOrcamentario & " CO, " & _
                gstrContribuinte & " CT, " & _
                gstrCidade & " MP, " & _
                gstrUF & " UF, " & _
                gstrBairro & " BR, " & _
                gstrCeps & " CP "
            
            strSql = strSql & "WHERE OPE.Intordempagamento = OP.Pkid " & _
                " AND OPE.Intparcela = SEP.pkid " & _
                " AND SEP.Pkid = SEPR.Intparcela " & _
                " AND CO.Pkid = SEPR.Intcodigoorcamentario " & _
                " AND OP.intContribuinte = CT.Pkid " & _
                " AND MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio " & _
                " AND CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep " & _
                " AND BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro " & _
                " AND UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF " & _
                " AND OP.Pkid = " & lngPkid
            
            strSql = strSql & " UNION  ALL  "
            
            strSql = strSql & "SELECT  ALL  'Orçamentário' quebra, CO.Strcodigoorcamentario StrCodigo, " & _
                 gstrISNULL("SEPR.Dblvalor", 0) & " Dblvalor, " & _
                 "CO.Strdescricao strDescricaoConta, " & _
                 "OP.bytTipo, " & _
                 "OP.INTNUMERO intOrdem, " & _
                 "OP.intExercicio IntExercicioOP, " & _
                 "OP.DTMDATA , " & _
                 gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_VARCHAR, " CT.CDC ") & strCONCAT & " ' - ' ") & strCONCAT & " CT.STRNOME intContribuinte, " & _
                 "CT.strCNPJCPF, " & _
                 "CT.strLogradouroC STRENDERECO, " & _
                 "CT.intNumero, " & _
                 "CT.STRCOMPLEMENTO STRCOMPLEMENTO, " & _
                 "MP.strDescricao STRMUNICIPIO, " & _
                 "UF.strsigla STRUF, " & _
                 "BR.strDescricao STRBAIRRO, " & _
                 "CP.INTCEP "
                strSql = strSql & "FROM " & gstrOrdemPagamento & " OP, " & _
                 gstrOrdemPagamentoResto & " OPR, " & _
                 gstrSubempenho & " SEP, " & _
                 gstrSubEmpRetencaoOrcamentaria & " SEPR, " & _
                 gstrCodigoOrcamentario & " CO, " & _
                 gstrContribuinte & " CT, " & _
                 gstrCidade & " MP, " & _
                 gstrUF & " UF, " & _
                 gstrBairro & " BR, " & _
                 gstrCeps & " CP "
                strSql = strSql & "WHERE OPR.Intordempagamento = OP.Pkid " & _
                 " AND OPR.Intparcela = SEP.pkid " & _
                 " AND SEP.Pkid = SEPR.Intparcela " & _
                 " AND CO.Pkid = SEPR.Intcodigoorcamentario " & _
                 " AND OP.intContribuinte = CT.Pkid " & _
                 " AND OP.intContribuinte = CT.Pkid " & _
                 " AND MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio " & _
                 " AND CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep " & _
                 " AND BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro " & _
                 " AND UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF " & _
                 " AND OP.Pkid = " & lngPkid
                 
                strSql = strSql & " ) A "
                strSql = strSql & "Group By quebra, StrCodigo, "
                strSql = strSql & "strDescricaoConta, "
                strSql = strSql & "bytTipo, "
                strSql = strSql & "intOrdem, "
                strSql = strSql & "IntExercicioOP, "
                strSql = strSql & "DTMDATA, "
                strSql = strSql & "intContribuinte, "
                strSql = strSql & "strCNPJCPF, "
                strSql = strSql & "STRENDERECO, "
                strSql = strSql & "intNumero, "
                strSql = strSql & "STRCOMPLEMENTO, "
                strSql = strSql & "STRMUNICIPIO, "
                strSql = strSql & "STRUF, "
                strSql = strSql & "STRBAIRRO, "
                strSql = strSql & "INTCEP "
                 
        End If
        strSql = strSql & " UNION  ALL  "
      
        strSql = strSql & "Select quebra, "
        strSql = strSql & "strCodigo, "
        strSql = strSql & "Sum(Dblvalor) Dblvalor, "
        strSql = strSql & "strDescricaoConta, "
        strSql = strSql & "bytTipo, "
        strSql = strSql & "intOrdem, "
        strSql = strSql & "IntExercicioOP, "
        strSql = strSql & "DTMDATA, "
        strSql = strSql & "intContribuinte, "
        strSql = strSql & "strCNPJCPF, "
        strSql = strSql & "STRENDERECO, "
        strSql = strSql & "intNumero, "
        strSql = strSql & "STRCOMPLEMENTO, "
        strSql = strSql & "STRMUNICIPIO, "
        strSql = strSql & "STRUF, "
        strSql = strSql & "STRBAIRRO, "
        strSql = strSql & "INTCEP "
        strSql = strSql & "from ( "
      
        'If bytDBType = Oracle Then
            strSql = strSql & "SELECT  'Extra Orçamentário' quebra, PLC.Strcontacontabil strCodigo,  SPL.Dblvalor, " & _
                     gstrCASEWHEN("PLC.Intextramaua", "NULL,''", gstrCONVERT(CDT_NVARCHAR, "plc.intextramaua ")) & strCONCAT & gstrCASEWHEN("PLC.Strdescricao", "NULL,''", " ' - ' " & strCONCAT & " PLC.Strdescricao ") & " strDescricaoConta, " & _
                     "OP.bytTipo, " & _
                     "OP.INTNUMERO intOrdem, " & _
                     "OP.intExercicio IntExercicioOP, " & _
                     "OP.DTMDATA , " & _
                     gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_NVARCHAR, " CT.CDC ") & strCONCAT & " ' - ' ") & strCONCAT & " CT.STRNOME intContribuinte, " & _
                     "CT.strCNPJCPF, " & _
                     "CT.strLogradouroC STRENDERECO, " & _
                     "CT.intNumero, " & _
                     "CT.STRCOMPLEMENTO STRCOMPLEMENTO, " & _
                     "MP.strDescricao STRMUNICIPIO, " & _
                     "UF.strsigla STRUF, " & _
                     "BR.strDescricao STRBAIRRO, " & _
                     "CP.INTCEP "
        
         strSql = strSql & "FROM " & gstrOrdemPagamento & " OP, " & _
                  gstrOrdemPagamentoResto & " OPR, " & _
                  gstrSubempenho & " SEP, " & _
                  gstrSubempenhoLiquidado & " SPL, " & _
                  gstrPlanoConta & " PLC, " & _
                  gstrContribuinte & " CT, " & _
                  gstrCidade & " MP, " & _
                  gstrUF & " UF, " & _
                  gstrBairro & " BR, " & _
                  gstrCeps & " CP "
                 
         strSql = strSql & "WHERE OPR.Intordempagamento = OP.Pkid " & _
                  "AND SPL.Intconta = PLC.PKID " & strOUTJOracle & _
                  "AND OPR.Intparcela " & strOUTJSQLServer & "= sep.pkid " & _
                  "AND SPL.Intparcela " & strOUTJOracle & "= SEP.Pkid " & _
                  "AND OP.Pkid = " & lngPkid & _
                  " AND OP.intContribuinte = CT.PKID " & _
                  "AND MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio " & _
                  "AND CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep " & _
                  "AND BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro " & _
                  "AND UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF " & _
                  "AND SPL.Dblvalor IS NOT NULL"
'        Else
'            strSql = strSql _
'                & " SELECT   ALL    'Extra Orçamentário' AS quebra, PLC.Strcontacontabil strCodigo, SPL.dblValor, " _
'                & " CASE PLC.Intextramaua WHEN NULL THEN '' ELSE CONVERT(VARCHAR, plc.intextramaua)  END  +  CASE PLC.Strdescricao WHEN NULL THEN '' ELSE  ' - '  +  PLC.Strdescricao  END  strDescricaoConta, " _
'                & " OP.bytTipo, OP.intNumero AS intOrdem, OP.intExercicio AS IntExercicioOP, OP.dtmData," _
'                & " CASE CT.CDC WHEN NULL THEN '' ELSE  CONVERT(VARCHAR,CT.CDC)  +  ' - '  END  +  CT.STRNOME intContribuinte," _
'                & " CT.strCNPJCPF," _
'                & " CT.strLogradouroC AS STRENDERECO, CT.intNumero, CT.strComplemento AS STRCOMPLEMENTO, MP.strDescricao AS STRMUNICIPIO, " _
'                & " UF.strSigla AS STRUF, BR.strDescricao AS STRBAIRRO, CP.intCep " _
'                & " FROM   " & gstrSubempenho & " SEP RIGHT OUTER JOIN " _
'                    & gstrOrdemPagamento & " OP INNER JOIN " _
'                    & gstrOrdemPagamentoResto & " OPR ON OP.PKId = OPR.intOrdemPagamento INNER JOIN " _
'                    & gstrContribuinte & " CT ON OP.intContribuinte = CT.PKId ON SEP.PKId = OPR.intParcela LEFT OUTER JOIN " _
'                    & gstrPlanoConta & " PLC INNER JOIN " _
'                    & gstrSubempenhoLiquidado & " SPL ON PLC.PKId = SPL.intConta ON SEP.PKId = SPL.intParcela LEFT OUTER JOIN " _
'                    & gstrCidade & " MP ON CT.intMunicipio = MP.PKId LEFT OUTER JOIN " _
'                    & gstrCeps & " CP ON CT.intCEP = CP.PKId LEFT OUTER JOIN " _
'                    & gstrBairro & " BR ON CT.intBairro = BR.PKId LEFT OUTER JOIN " _
'                    & gstrUF & " UF ON CT.intUF = UF.PKId " _
'                & " Where (OP.Pkid = " & lngPkid & ") And (SPL.DBLVALOR Is Not Null) "
'
'        End If
          strSql = strSql & " UNION  ALL "
          
          strSql = strSql & "SELECT  ALL  'Extra Orçamentário' quebra,PLC.Strcontacontabil strCodigo , SPL.Dblvalor, " & _
                  gstrCASEWHEN("PLC.Intextramaua", "null,''", gstrCONVERT(CDT_VARCHAR, "plc.intextramaua ")) & strCONCAT & gstrCASEWHEN("PLC.Strdescricao", "NULL,''", " ' - ' " & strCONCAT & " PLC.Strdescricao ") & " strDescricaoConta, " & _
                  "OP.bytTipo, " & _
                  "OP.INTNUMERO intOrdem, " & _
                  "OP.intExercicio IntExercicioOP, " & _
                  "OP.DTMDATA , " & _
                  gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_VARCHAR, " CT.CDC ") & strCONCAT & " ' - ' ") & strCONCAT & " CT.STRNOME intContribuinte, " & _
                  "CT.strCNPJCPF, " & _
                  "CT.strLogradouroC STRENDERECO, " & _
                  "CT.intNumero, " & _
                  "CT.STRCOMPLEMENTO STRCOMPLEMENTO, " & _
                  "MP.strDescricao STRMUNICIPIO, " & _
                  "UF.strsigla STRUF, " & _
                  "BR.strDescricao STRBAIRRO, " & _
                  "CP.INTCEP "
                  
         strSql = strSql & "FROM " & gstrOrdemPagamento & " OP, " & _
                 gstrOrdemPagamentoEmpenho & " OPE, " & _
                 gstrSubempenho & " SEP, " & _
                 gstrSubempenhoLiquidado & " SPL, " & _
                 gstrPlanoConta & " PLC, " & _
                 gstrContribuinte & " CT, " & _
                 gstrCidade & " MP, " & _
                 gstrUF & " UF, " & _
                 gstrBairro & " BR, " & _
                 gstrCeps & " CP "
                 
         strSql = strSql & "WHERE OPE.Intordempagamento = OP.Pkid " & _
                 "AND SPL.Intconta " & strOUTJSQLServer & "= PLC.PKID " & strOUTJOracle & _
                 " AND OPE.Intparcela = sep.pkid " & _
                 "AND SPL.Intparcela " & strOUTJOracle & " = SEP.Pkid " & _
                 "AND OP.Pkid = " & lngPkid & _
                 " AND OP.intContribuinte = CT.PKID " & _
                 "AND MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio " & _
                 "AND CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep " & _
                 "AND BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro " & _
                 "AND UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF " & _
                 "AND SPL.Dblvalor IS NOT NULL"
                 
         strSql = strSql & " UNION  ALL "
                 
         strSql = strSql & "SELECT  ALL  'Orçamentário' quebra, ' ' strCodigo, (SELECT SUM(Dblvalordesconto) FROM " & gstrOrdemPagamentoDespesaExtra & " WHERE intOrdemPagamento = " & lngPkid & " GROUP BY intOrdemPagamento) Dblvalor, " & _
                 "'Total Descontado' strDescricaoConta, " & _
                 "OP.bytTipo, " & _
                 "OP.INTNUMERO intOrdem, " & _
                 "OP.intExercicio IntExercicioOP, " & _
                 "OP.DTMDATA , " & _
                 gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_VARCHAR, " CT.CDC ") & strCONCAT & " ' - '") & strCONCAT & " CT.STRNOME intContribuinte, " & _
                 "CT.strCNPJCPF, " & _
                 "CT.strLogradouroC STRENDERECO, " & _
                 "CT.intNumero, " & _
                 "CT.STRCOMPLEMENTO STRCOMPLEMENTO, " & _
                 "MP.strDescricao STRMUNICIPIO, " & _
                 "UF.strsigla STRUF, " & _
                 "BR.strDescricao STRBAIRRO, " & _
                 "CP.INTCEP "
         strSql = strSql & "FROM " & gstrOrdemPagamento & " OP, " & _
                 gstrContribuinte & " CT, " & _
                 gstrCidade & " MP, " & _
                 gstrUF & " UF, " & _
                 gstrBairro & " BR, " & _
                 gstrCeps & " CP "
         strSql = strSql & "WHERE OP.intContribuinte = CT.Pkid " & _
                  " AND MP.PKID " & strOUTJOracle & " =CT.intMunicipio " & _
                  " AND CP.PKID " & strOUTJOracle & " = CT.intCep " & _
                  " AND BR.PKID " & strOUTJOracle & " = CT.intBairro " & _
                  " AND UF.PKID " & strOUTJOracle & " = CT.intUF " & _
                  " AND OP.bytTipo = 2 " & _
                  " AND OP.Pkid = " & lngPkid
         
        strSql = strSql & " ) A "
        strSql = strSql & "Group By quebra, "
        strSql = strSql & "strCodigo, "
        strSql = strSql & "strDescricaoConta, "
        strSql = strSql & "bytTipo, "
        strSql = strSql & "intOrdem, "
        strSql = strSql & "IntExercicioOP, "
        strSql = strSql & "DTMDATA, "
        strSql = strSql & "intContribuinte, "
        strSql = strSql & "strCNPJCPF, "
        strSql = strSql & "STRENDERECO, "
        strSql = strSql & "intNumero, "
        strSql = strSql & "STRCOMPLEMENTO, "
        strSql = strSql & "STRMUNICIPIO, "
        strSql = strSql & "STRUF, "
        strSql = strSql & "STRBAIRRO, "
        strSql = strSql & "INTCEP "
         
         strSql = strSql & " ORDER BY quebra DESC, strDescricaoConta "
  Else
         strSql = "SELECT  ALL 'Orçamentário' quebra, ' ' strCodigo, (SELECT SUM(Dblvalordesconto) FROM " & gstrOrdemPagamentoDespesaExtra & " WHERE intOrdemPagamento = " & lngPkid & " GROUP BY intOrdemPagamento) Dblvalor, " & _
                 "'Total Descontado' strDescricaoConta, " & _
                 "OP.bytTipo, " & _
                 "OP.INTNUMERO intOrdem, " & _
                 "OP.intExercicio IntExercicioOP, " & _
                 "OP.DTMDATA , " & _
                 "(" & gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_VARCHAR, "CT.CDC") & strCONCAT & " ' - '") & ")" & strCONCAT & " CT.STRNOME intContribuinte, " & _
                 "CT.strCNPJCPF, " & _
                 "CT.strLogradouroC STRENDERECO, " & _
                 "CT.intNumero, " & _
                 "CT.STRCOMPLEMENTO STRCOMPLEMENTO, " & _
                 "MP.strDescricao STRMUNICIPIO, " & _
                 "UF.strsigla STRUF, " & _
                 "BR.strDescricao STRBAIRRO, " & _
                 "CP.INTCEP "
         strSql = strSql & "FROM " & gstrOrdemPagamento & " OP, " & _
                 gstrContribuinte & " CT, " & _
                 gstrCidade & " MP, " & _
                 gstrUF & " UF, " & _
                 gstrBairro & " BR, " & _
                 gstrCeps & " CP "
         strSql = strSql & "WHERE OP.intContribuinte = CT.Pkid " & _
                  " AND MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio " & _
                  " AND CP.PKID  " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep " & _
                  " AND BR.PKID" & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro " & _
                  " AND UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF " & _
                  " AND op.byttipo = 2 " & _
                  " AND OP.Pkid = " & lngPkid
         
         
    If bytDBType = Oracle Then
            strSql = strSql & " ORDER BY quebra DESC"
    End If
  End If
                        
    strQueryDescontoOPs = strSql
            
End Function
Private Sub ImprimeSub()
Dim strQuery        As String
Dim adoRelatorio    As ADODB.Recordset
    
    With rptDescontosDeOPs
        
        strQuery = strQueryDescontoOPs
       
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strQuery, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                If adoRelatorio!dblValor <> 0 Then
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strQuery
                 
                     Set .adoDataControl.Recordset = adoRelatorio
                 
                     Set SubReport.object = rptDescontosDeOPs
                     
                     GroupFooter1.NewPage = ddNPBefore
                     Me.SubReport.Visible = True
                     GroupFooter1.Visible = True
                     intPagina = intPagina + 1
                     rptDescontosDeOPs.bytRptOrigem = 1
                 Else
                    GroupFooter1.NewPage = ddNPNone
    
                     Me.ReportFooter.Visible = False
                     GroupFooter1.Visible = False
                 End If
            
            Else
                GroupFooter1.NewPage = ddNPNone

                 Me.ReportFooter.Visible = False
                 GroupFooter1.Visible = False
            
            End If
             
        End If
    
    End With

End Sub

Private Function BlnDescontoOP(lngPkid As Long) As Boolean
    Dim adoResultado    As ADODB.Recordset
    Dim strSql          As String
    
    strSql = "Select "
    strSql = strSql & "SER.Dblvalor "
    strSql = strSql & "From "
    strSql = strSql & gstrOrdemPagamento & " OP, "
    strSql = strSql & gstrOrdemPagamentoResto & " OPR, "
    strSql = strSql & gstrSubEmpRetencaoOrcamentaria & " SER "
    strSql = strSql & "Where "
    strSql = strSql & "OP.Pkid = opr.intordempagamento AND "
    strSql = strSql & "OPR.Intparcela = SER.Intparcela AND "
    strSql = strSql & "OP.Pkid = " & lngPkid
    
    strSql = strSql & "Union ALL "
    
    strSql = strSql & "Select "
    strSql = strSql & "SER.Dblvalor "
    strSql = strSql & "From "
    strSql = strSql & gstrOrdemPagamento & " OP, "
    strSql = strSql & gstrOrdemPagamentoEmpenho & " OPE, "
    strSql = strSql & gstrSubEmpRetencaoOrcamentaria & " SER "
    strSql = strSql & "Where "
    strSql = strSql & "OP.Pkid = OPE.intordempagamento AND "
    strSql = strSql & "OPE.Intparcela = SER.Intparcela AND "
    strSql = strSql & "OP.Pkid = " & lngPkid
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            BlnDescontoOP = True
            Exit Function
        End If
    End If
    
    BlnDescontoOP = False
End Function

