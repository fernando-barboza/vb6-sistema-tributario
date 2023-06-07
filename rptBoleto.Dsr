VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptBoleto 
   Caption         =   "Tributario - rptBoleto (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptBoleto.dsx":0000
End
Attribute VB_Name = "rptBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intContador         As Integer
Public blnPrimeira      As Boolean 'Não coloca margem superior na 1ª página
Public blnValorEmReal   As Boolean 'Identifica se o valor do boleto esta em Real

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
    PadronizaToolBarRelatorio Me
    intContador = 0
End Sub
'
Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
'
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

Private Sub AjustarDetalhe(intDistancia)
Dim objControl As Object

    For Each objControl In Detail.Controls
        If UCase(TypeName(objControl)) = "FIELD" _
            Or UCase(TypeName(objControl)) = "LABEL" _
            Or UCase(TypeName(objControl)) = "BARCODE" _
            Or UCase(TypeName(objControl)) = "IMAGE" Then
            
            objControl.Top = objControl.Top + intDistancia
            
        ElseIf UCase(TypeName(objControl)) = "LINE" Then
            If UCase(objControl.Name) <> "LNHDETALHE" And UCase(objControl.Name) <> "Line7" Then
                objControl.Y1 = objControl.Y1 + intDistancia
                objControl.Y2 = objControl.Y2 + intDistancia
            End If
            
        End If
    Next

End Sub

Private Sub Detail_Format()
Dim objControl      As Object
Dim strSql          As String
Dim adoBanco        As ADODB.Recordset
Dim strCodBarras    As String
Dim adoResultado    As ADODB.Recordset
Dim adoCommand   As ADODB.Command
Dim lngNumeroGuia   As Long
Dim intFebraban     As Integer
Dim ValorParcela

On Error GoTo Problema_Na_Rotina


    txtstrReferencia = gstrNomeDoMes(Month(txtdtmDtVencimento) - 1)
    
    If adoDataControl.NRecords > 0 Then
        strSql = strSql & "SELECT "
        strSql = strSql & "dblPorcentagemIssVar "
        strSql = strSql & "FROM "
        strSql = strSql & "Tbllancamentoeconiss LEI, "
        strSql = strSql & "Tbllancamentoeconomico LE "
        strSql = strSql & "WHERE "
        strSql = strSql & "LEI.Intlancamentoeconomico = LE.Pkid "
        strSql = strSql & "AND LE.Intlancamentoalfa = " & adoDataControl.Recordset("pkIdAlfa")
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
            If Not adoResultado.EOF Then
                txtdblAliquota = gstrConvVrDoSql(adoResultado!dblPorcentagemIssVar, 3) & " %"
            End If
        End If
    End If
        
    'Query utilizada para pegar o Codigo Febraban da tblEmpresa
    strSql = ""
    strSql = strSql & "Select * From " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 30, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            If gstrENulo(adoResultado!intFebraban) <> "" Then
                intFebraban = gstrENulo(adoResultado!intFebraban)
            Else
                ExibeMensagem "Código Febraban não encontrado."
                Exit Sub
            End If
        Else
            ExibeMensagem "Código Febraban não encontrado."
            Exit Sub
        End If
    End If

ProximoNumeroGuia:
    
    lngNumeroGuia = glngRetornaProximoNumeroGuia
    If Val(lngNumeroGuia) = 0 Then
        Exit Sub
    End If
  
    txtstrNumDoc.Text = lngNumeroGuia
'    txtstrNumDoc1.Text = lngNumeroGuia
    
    ValorParcela = adoDataControl.Recordset("dblValorParcela")
    If InStr(ValorParcela, ",") = 0 Then
        ValorParcela = gstrConvVrDoSql(ValorParcela)
    Else
        If Len(ValorParcela) - InStr(ValorParcela, ",") < 2 Then
            ValorParcela = gstrConvVrDoSql(ValorParcela)
        End If
    End If
    
    'Vamos definir o codigo de barras
    strCodBarras = gstrMontaCodigoBarras(FICHA_COMPENSACAO, txtintConta, ValorParcela, adoDataControl.Recordset("dtmdtVencimento"), intFebraban, lngNumeroGuia, True, blnValorEmReal)
    
    If Len(strCodBarras) = 0 Then Exit Sub
    'Vamos definir a linha digitavel
    lblstrCodigoDigitavel = gstrMontaLinhaDigitavel(FICHA_COMPENSACAO, strCodBarras)
    'Vamos definir o nosso numero
    txtstrNossoNumero = gstrMontaNossoNumero(txtintConta, lngNumeroGuia)
'    txtstrNossoNumero1 = txtstrNossoNumero
    
    'Pego o Valor Movimento Estimativa
    txtdblVAlMovEst = gstrConvVrDoSql(PegaValorMovimentoEstimativa, 2)
    
    txtdblQuantidade = ValorParcela
'    txtdblQuantidade1 = ValorParcela
        
    If blnValorEmReal Then
        txtdblValorIndexador = ""
    End If
    
    brcCodigoDeBarras.Caption = strCodBarras

    'Insere o Nº da tblGuia
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    strSql = ""
    strSql = strSql & "INSERT INTO " & gstrGuias & "("
    strSql = strSql & "intContaBancaria, "
    strSql = strSql & "intNumero, "
    strSql = strSql & "dtmdtEmissao, "
    strSql = strSql & "dblValor, "
    strSql = strSql & "strCodBarra, "
    strSql = strSql & "dtmdtAtualizacao, "
    strSql = strSql & "lngCodUsr, "
    strSql = strSql & "dtmdtVencimento "
    strSql = strSql & ") VALUES ("
    strSql = strSql & txtintConta & ", "
    strSql = strSql & lngNumeroGuia & ", "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSql = strSql & gstrConvVrParaSql(adoDataControl.Recordset("dblValorReal")) & ", '"
    strSql = strSql & brcCodigoDeBarras.Caption & "', "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSql = strSql & glngCodUsr & ", "
    strSql = strSql & gstrConvDtParaSql(adoDataControl.Recordset("dtmdtVencimento"))
    strSql = strSql & ")"

    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = gcncADOMain
    adoCommand.CommandText = strSql
    adoCommand.Execute strSql, , adExecuteNoRecords

    'Inserir a guia na tabela TblLancamentoGuias
    strSql = ""
    strSql = "INSERT INTO " & gstrLancamentoGuias & "("
    strSql = strSql & "intlancamentovalor, "
    strSql = strSql & "intguias, "
    strSql = strSql & "dblvalorprincipal, "
    strSql = strSql & "dblvalormulta, "
    strSql = strSql & "dblvalorjuros, "
    strSql = strSql & "dblvalorcorrecao, "
    strSql = strSql & "dblvalordesconto, "
    strSql = strSql & "dtmdtatualizacao, "
    strSql = strSql & "lngcodusr) "
    strSql = strSql & "Values ("
    strSql = strSql & adoDataControl.Recordset("PkidParcela") & ", "
    strSql = strSql & glngRetornaPkidTabelaPai("seqTblGuias", gstrGuias) & ", "
    strSql = strSql & gstrConvVrParaSql(adoDataControl.Recordset("dblValorReal")) & ", "
    strSql = strSql & "0, "
    strSql = strSql & "0, "
    strSql = strSql & "0, "
    strSql = strSql & "0, "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSql = strSql & glngCodUsr & ") "

    If gobjBanco.Execute(strSql) Then
        gobjBanco.ExecutaCommitTrans
    Else
        ExibeMensagem "Erro na gravação dos lançamentos da guia. Guia não gravada."
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If

    If blnPrimeira = True And adoDataControl.NRecords > 1 Then
       Detail.NewPage = ddNPAfter 'Adiciona nova pagina
       blnPrimeira = False
       GroupFooter1.NewPage = ddNPAfter
    Else
       Detail.NewPage = ddNPNone
    End If
    
    
    'Vamos atribuir a imagem do banco
    On Error Resume Next
    imgLogoBanco.SizeMode = ddSMZoom

    strSql = ""
    strSql = strSql & "SELECT BA.intLogoBanco, BA.intBanco, BA.intDigitoBanco, CB.strCedente, CB.strDigitoVerificador, AG.strAgencia, TC.strEspecieDoc, TC.strAceite, TC.strCarteira "
    strSql = strSql & "FROM "
    strSql = strSql & gstrBanco & " BA, " & gstrContaBancaria & " CB, " & gstrAgencia & " AG, " & gstrTipoCodigoBarra & " TC "
    strSql = strSql & "WHERE BA.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & "CB.intBanco AND " & _
                      "AG.Pkid = CB.intAgencia AND " & _
                      "TC.Pkid = CB.intTipoCodigoBarra AND " & _
                      "CB.Pkid = " & txtintConta

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoBanco) Then
        With adoBanco
            If .EOF = False Then
                
                LeImagem Val(gstrENulo(!intLogoBanco)), imgLogoBanco
'                LeImagem Val(gstrENulo(!intLogoBanco)), imgLogoBanco1
                
                txtstrCodigoBanco = !intBanco
                
                txtstrCodigoBanco = Format(txtstrCodigoBanco, "000")
                txtstrCodigoBanco = txtstrCodigoBanco & IIf(IsNull(!intDigitoBanco), "", "-" & !intDigitoBanco)
                
'                txtstrCodigoBanco1 = txtstrCodigoBanco
                
                txtstrAgencia = !strAgencia & " " & !strCedente
'                txtstrAgencia1 = !strAgencia & " " & !strCedente
                
                txtstrEspecieDoc = !strEspecieDoc
                txtstrAceite = !strAceite
                txtstrCarteira = !strCarteira
                
            End If
        End With
        adoBanco.Close: Set adoBanco = Nothing
    Else
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & "SELECT EM.strNomeFantasia "
    strSql = strSql & "FROM "
    strSql = strSql & gstrEmpresa & " EM "

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoBanco) Then
        With adoBanco
            If .EOF = False Then
                
                txtstrCedente = !strNomeFantasia
'                txtstrCedente1 = !strNomeFantasia
                
            End If
        End With
        adoBanco.Close: Set adoBanco = Nothing
    Else
        Exit Sub
    End If
    
    'Carrega as intrucoes da parcela
    txtstrInstrucoes = CarregaInstrucoesParcelas(False, adoDataControl.Recordset("strComposicaoDaReceita"), adoDataControl.Recordset("intExercicio"), adoDataControl.Recordset("intParcela"), adoDataControl.Recordset("bitParcelaValida"), adoDataControl.Recordset("PkidAlfa"))
    
    txtstrCep = gstrCEPFormatado(gstrVerificaCampoNulo(txtstrCep))
    txtdtmDocumento.Text = gstrDataDoSistema
    txtdtmProcessamento.Text = gstrDataDoSistema
    
    txtstrEspecie = IIf(IsNull(adoDataControl.Recordset("strIndexador")), "R$", adoDataControl.Recordset("strIndexador"))
'    txtstrEspecie1 = txtstrEspecie
    
'    fldMargem.Top = -332
'
'    If intContador = 1 Then
'        fldMargem = vbNewLine & vbNewLine & vbNewLine
'        fldMargem.Visible = True
'        lnhDetalhe.Visible = False
'        fldAjuste.Visible = False
'        Detail.CanShrink = False
'        AjustarDetalhe 360
'
'    ElseIf intContador = 2 Then
'        fldAjuste = "a" & vbNewLine & "a" & vbNewLine & "a" & vbNewLine & "a" & vbNewLine
'        fldMargem = ""
'        fldMargem.Visible = False
'        lnhDetalhe.Visible = True
'        Detail.CanShrink = False
'        AjustarDetalhe 850
'
'    Else
'        fldMargem = ""
'        fldMargem.Visible = False
'        lnhDetalhe.Visible = True
'        Detail.CanShrink = True
'        AjustarDetalhe -180
'    End If
'
'    If intContador = 3 Then
'        intContador = 0
'    End If
'
'    intContador = intContador + 1

    Exit Sub

Problema_Na_Rotina:

  If InStr(1, UCase(Err.Description), "UK_TBLGUIAS_INTNUMERODTEMISSAO") > 0 Then
      GoTo ProximoNumeroGuia
  Else
      ExibeDetalheErro Err.Description & "- rptCarneAcordoBoleto_Detail_Format"
      gobjBanco.ExecutaRollbackTrans
  End If

    
End Sub

Private Sub GroupHeader1_Format()
    intContador = intContador + 1
    
    If (intContador Mod 2) > 0 Then
        GroupHeader1.Height = 320
    Else
        GroupHeader1.Height = 875
    End If
End Sub

Private Function PegaValorMovimentoEstimativa() As Variant
Dim strSql As String
Dim adoResult As New ADODB.Recordset

    strSql = "SELECT LISS.dblValorEstimadoIss ValorEstimado FROM "
    strSql = strSql & gstrLancamentoEconomico & " LE"
    strSql = strSql & ", " & gstrLancamentoEconIss & " LISS"
    strSql = strSql & " WHERE LE.intLancamentoAlfa = " & Me.adoDataControl.Recordset("PkidAlfa")
    strSql = strSql & " AND LE.PKId = LISS.intLancamentoEconomico"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 30, adoResult) Then
        If Not adoResult.EOF Then
            PegaValorMovimentoEstimativa = gstrENulo(adoResult.Fields("ValorEstimado"))
        Else
            PegaValorMovimentoEstimativa = 0
        End If
    End If
    adoResult.Close
End Function


