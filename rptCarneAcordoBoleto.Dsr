VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCarneAcordoBoleto 
   Caption         =   "rptCarneAcordoBoleto (ActiveReport)"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptCarneAcordoBoleto.dsx":0000
End
Attribute VB_Name = "rptCarneAcordoBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intContador                 As Integer
Public blnPrimeira              As Boolean 'Não coloca margem superior na 1ª página
Public blnValorEmReal           As Boolean 'Identifica se o valor do boleto esta em Real
Public blnParcelasAtualizadas   As Boolean

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

Private Sub Detail_Format()
Dim objControl      As Object
Dim strsql          As String
Dim adoBanco        As ADODB.Recordset
Dim strCodBarras    As String
Dim adoResultado    As ADODB.Recordset
Dim adoCommand   As ADODB.Command
Dim lngNumeroGuia   As Long
Dim intFebraban     As Integer
Dim ValorParcela    As Variant
Dim dblValorReal    As Variant

On Error GoTo Problema_Na_Rotina

    'Query utilizada para pegar o Codigo Febraban da tblEmpresa
    strsql = ""
    strsql = strsql & "Select * From " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 30, adoResultado) Then
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
    DoEvents
    lngNumeroGuia = glngRetornaProximoNumeroGuia
    If Val(lngNumeroGuia) = 0 Then
        Exit Sub
    End If

    txtstrNumDoc.Text = lngNumeroGuia
    txtstrNumDoc1.Text = lngNumeroGuia
    
    'PROVISORIO
    If Month(adoDataControl.Recordset("Dtmdtvencimento").Value) = 1 Then
        lblAviso.Visible = True
    Else
        lblAviso.Visible = False
    End If
    
    dblValorReal = adoDataControl.Recordset("dblValorReal")
    ValorParcela = adoDataControl.Recordset("dblValorParcela")
    
'      strsql = ""
'    strsql = strsql & "SELECT BA.intLogoBanco, BA.intBanco, BA.intDigitoBanco, CB.strCedente, CB.strDigitoVerificador0, AG.strAgencia, TC.strEspecieDoc, TC.strAceite, TC.strCarteira, CB.strconta "
'    strsql = strsql & "FROM "
'    strsql = strsql & gstrBanco & " BA, " & gstrContaBancaria & " CB, " & gstrAgencia & " AG, " & gstrTipoCodigoBarra & " TC "
'    strsql = strsql & "WHERE BA.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & "CB.intBanco AND " & _
'                      "AG.Pkid = CB.intAgencia AND " & _
'                      "TC.Pkid = CB.intTipoCodigoBarra AND " & _
'                      "CB.Pkid = " & txtintConta
'
'     txtstrAgencia = !strAgencia & " / " & !strConta '!strCedente
'     txtstrAgencia1 = !strAgencia & " / " & !strConta '!strCedente

        
    'Caso seja acordo com parcelas atualizadas, vamos atualizar o valor das parcelas de acordo com o exercicio solicitado
    If blnParcelasAtualizadas And blnValorEmReal Then
      
        strsql = gstrStoredProcedure("sp_AtualizaParcela", adoDataControl.Recordset("intComposicao").Value & ", " & adoDataControl.Recordset("intExercicio").Value & ", " & adoDataControl.Recordset("intParcela").Value & ", " & gstrConvDtParaSql(adoDataControl.Recordset("Dtmdtvencimento").Value) & ", " & gstrConvDtParaSql(adoDataControl.Recordset("Dtmdtvencimento").Value) & ", " & gstrConvVrParaSql(adoDataControl.Recordset("dblValorReal").Value) & ", " & adoDataControl.Recordset("intMoeda").Value, True)
        
        Set gobjBanco = New clsBanco

        If gobjBanco.CriaADO(strsql, 80, adoResultado) Then
            dblValorReal = Space$(0) & gstrConvVrDoSql(adoResultado("dblValorPrincipal").Value)
        End If
      
        adoResultado.Close: Set adoResultado = Nothing
        
        ValorParcela = dblValorReal
        
        txtdblValor1 = dblValorReal
        txtdblValor2 = dblValorReal
      
    End If

    If InStr(ValorParcela, ",") = 0 Then
        ValorParcela = gstrConvVrDoSql(ValorParcela)
    Else
        If Len(ValorParcela) - InStr(ValorParcela, ",") < 2 Then
            ValorParcela = gstrConvVrDoSql(ValorParcela)
        End If
    End If

    'Vamos definir o codigo de barras ****Vencimento PROVISORIO*******
    'strCodBarras = gstrMontaCodigoBarras(FICHA_COMPENSACAO, txtintConta, ValorParcela, adoDataControl.Recordset("dtmdtVencimento"), intFebraban, lngNumeroGuia, True, blnValorEmReal)
    'PROVISORIO
    strCodBarras = gstrMontaCodigoBarras(FICHA_COMPENSACAO, txtintConta, ValorParcela, IIf(Month(adoDataControl.Recordset("Dtmdtvencimento").Value) = 1, "07/02/2006", adoDataControl.Recordset("dtmdtVencimento")), intFebraban, lngNumeroGuia, True, blnValorEmReal)
    If Len(strCodBarras) = 0 Then Exit Sub
    'Vamos definir a linha digitavel
    lblstrCodigoDigitavel = gstrMontaLinhaDigitavel(FICHA_COMPENSACAO, strCodBarras)
    'Vamos definir o nosso numero
    txtstrNossoNumero = gstrMontaNossoNumero(txtintConta, lngNumeroGuia)
    txtstrNossoNumero1 = txtstrNossoNumero
    
    txtdblQuantidade = ValorParcela
    txtdblQuantidade1 = ValorParcela
    
    brcCodigoDeBarras.Caption = strCodBarras

    'Insere o Nº da tblGuia
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    strsql = ""
    strsql = strsql & "INSERT INTO " & gstrGuias & "("
    strsql = strsql & "intContaBancaria, "
    strsql = strsql & "intNumero, "
    strsql = strsql & "dtmdtEmissao, "
    strsql = strsql & "dblValor, "
    strsql = strsql & "strCodBarra, "
    strsql = strsql & "dtmdtAtualizacao, "
    strsql = strsql & "lngCodUsr, "
    strsql = strsql & "dtmdtVencimento "
    strsql = strsql & ") VALUES ("
    strsql = strsql & txtintConta & ", "
    strsql = strsql & lngNumeroGuia & ", "
    strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strsql = strsql & gstrConvVrParaSql(dblValorReal) & ", '"
    strsql = strsql & brcCodigoDeBarras.Caption & "', "
    strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strsql = strsql & glngCodUsr & ", "
    strsql = strsql & gstrConvDtParaSql(adoDataControl.Recordset("dtmdtVencimento"))
    strsql = strsql & ")"

    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = gcncADOMain
    adoCommand.CommandText = strsql
    adoCommand.Execute strsql, , adExecuteNoRecords

    'Inserir a guia na tabela TblLancamentoGuias
    strsql = ""
    strsql = "INSERT INTO " & gstrLancamentoGuias & "("
    strsql = strsql & "intlancamentovalor, "
    strsql = strsql & "intguias, "
    strsql = strsql & "dblvalorprincipal, "
    strsql = strsql & "dblvalormulta, "
    strsql = strsql & "dblvalorjuros, "
    strsql = strsql & "dblvalorcorrecao, "
    strsql = strsql & "dblvalordesconto, "
    strsql = strsql & "dtmdtatualizacao, "
    strsql = strsql & "lngcodusr) "
    strsql = strsql & "Values ("
    strsql = strsql & adoDataControl.Recordset("PkidParcela") & ", "
    strsql = strsql & glngRetornaPkidTabelaPai("seqTblGuias", gstrGuias) & ", "
    strsql = strsql & gstrConvVrParaSql(dblValorReal) & ", "
    strsql = strsql & "0, "
    strsql = strsql & "0, "
    strsql = strsql & "0, "
    strsql = strsql & "0, "
    strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strsql = strsql & glngCodUsr & ") "

    If gobjBanco.Execute(strsql) Then
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
    
    'Se o valor nao for em Real, nao vamos exibir, pois ele sera calculado no pagamento
    If Not blnValorEmReal Then
        txtdblValor1 = ""
        txtdblValor2 = ""
    End If
    'Vamos atribuir a imagem do banco
    On Error Resume Next
    imgLogoBanco.SizeMode = ddSMZoom

    strsql = ""
    strsql = strsql & "SELECT BA.intLogoBanco, BA.intBanco, BA.intDigitoBanco, CB.strCedente, CB.strDigitoVerificador, AG.strAgencia, TC.strEspecieDoc, TC.strAceite, TC.strCarteira, CB.strconta "
    strsql = strsql & "FROM "
    strsql = strsql & gstrBanco & " BA, " & gstrContaBancaria & " CB, " & gstrAgencia & " AG, " & gstrTipoCodigoBarra & " TC "
    strsql = strsql & "WHERE BA.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & "CB.intBanco AND " & _
                      "AG.Pkid = CB.intAgencia AND " & _
                      "TC.Pkid = CB.intTipoCodigoBarra AND " & _
                      "CB.Pkid = " & txtintConta

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 15, adoBanco) Then
        With adoBanco
            If .EOF = False Then
                
                LeImagem Val(gstrENulo(!intLogoBanco)), imgLogoBanco
                LeImagem Val(gstrENulo(!intLogoBanco)), imgLogoBanco1
                
                txtstrCodigoBanco = !intBanco
                
                txtstrCodigoBanco = Format(txtstrCodigoBanco, "000")
                txtstrCodigoBanco = txtstrCodigoBanco & IIf(IsNull(!intDigitoBanco), "", "-" & !intDigitoBanco)
                
                txtstrCodigoBanco1 = txtstrCodigoBanco
                
                 txtstrAgencia.Visible = True
                 txtstrAgencia1.Visible = True
                 
                 txtstrAgencia = !strAgencia & " / " & !strConta '!strCedente
                 txtstrAgencia1 = !strAgencia & " / " & !strConta '!strCedente
                
                txtstrEspecieDoc = !strEspecieDoc
                txtstrAceite = !strAceite
                txtstrCarteira = !strCarteira
                
            End If
        End With
        adoBanco.Close: Set adoBanco = Nothing
    Else
        Exit Sub
    End If

    strsql = ""
    strsql = strsql & "SELECT EM.strNomeFantasia "
    strsql = strsql & "FROM "
    strsql = strsql & gstrEmpresa & " EM "

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoBanco) Then
        With adoBanco
            If .EOF = False Then
                
                txtstrCedente = !strNomeFantasia
                txtstrCedente1 = !strNomeFantasia
                
            End If
        End With
        adoBanco.Close: Set adoBanco = Nothing
    Else
        Exit Sub
    End If
    
    'Carrega as intrucoes da parcela
    txtstrInstrucoes = CarregaInstrucoesParcelas(False, adoDataControl.Recordset("strComposicaoDaReceita"), adoDataControl.Recordset("intExercicio"), adoDataControl.Recordset("intParcela"), adoDataControl.Recordset("bitParcelaValida"), adoDataControl.Recordset("PkidAlfa"))
    
    txtintCEPC = gstrCEPFormatado(gstrVerificaCampoNulo(txtintCEPC))
    txtdtmDocumento.Text = gstrDataDoSistema
    txtdtmProcessamento.Text = gstrDataDoSistema
    
    txtstrEspecie = IIf(IsNull(adoDataControl.Recordset("strIndexador")), "R$", adoDataControl.Recordset("strIndexador"))
    txtstrEspecie1 = txtstrEspecie
    
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

Private Sub GroupHeader1_Format()
  If blnPrimeira = True Then
     GroupHeader1.Height = 0
  Else
     GroupHeader1.Height = 105
  End If
End Sub
