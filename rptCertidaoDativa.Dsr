VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCertidaoDativa 
   Caption         =   "Tributario - rptCertidaoDativa (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptCertidaoDativa.dsx":0000
End
Attribute VB_Name = "rptCertidaoDativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblQuantidadeIndexador As Double


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
    PadronizaToolBarRelatorio Me
    'LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    'lblRelatorio = Me.Caption
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

Private Sub GroupFooter1_Format()
    PreencherCampo rchFim, "|dblTotalDebito|", gstrConvVrDoSql(fldSubTotal)
    PreencherCampo rchFim, "|dblTotalDebitoExtenso|", gstrExtenso(fldSubTotal)
    PreencherCampo rchFim, "|dblQuantidadeIndexador|", gstrConvVrDoSql(Str(Val(gstrConvVrParaSql(fldSubTotal)) / Val(gstrConvVrParaSql(IIf(dblQuantidadeIndexador = 0, 1, dblQuantidadeIndexador)))), 4)
End Sub

Private Sub GroupHeader1_Format()
Dim varAux
    
    rchComeco.LoadFile gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\CertidaoDA.rtf", rtfRTF
        
    If adoDataControl.NRecords > 0 Then
        With adoDataControl
            PreencherCampo rchComeco, "|intCertidao|", gstrConvVrDoSql(gstrENulo(.Recordset("intCertidao")), 0)
            PreencherCampo rchComeco, "|dtmDtInscricao|", gstrENulo(.Recordset("dtmDtInscricao"))
            PreencherCampo rchComeco, "|strTipoLogradouroC|", gstrENulo(.Recordset("strTipoLogradouroC"))
            PreencherCampo rchComeco, "|strTitLogradouroC|", gstrENulo(.Recordset("strTitLogradouroC"))
            PreencherCampo rchComeco, "|strLogradouroC|", gstrENulo(.Recordset("strLogradouroC"))
            PreencherCampo rchComeco, "|strNumeroC|", gstrENulo(.Recordset("strNumeroC"))
            PreencherCampo rchComeco, "|strComplementoC|", gstrENulo(.Recordset("strComplementoC"))
            PreencherCampo rchComeco, "|strBairroC|", gstrENulo(.Recordset("strBairroC"))
            PreencherCampo rchComeco, "|strMunicipioC|", gstrENulo(.Recordset("strMunicipioC"))
            PreencherCampo rchComeco, "|strUFC|", gstrENulo(.Recordset("strUFC"))
            PreencherCampo rchComeco, "|intCEPC|", gstrENulo(.Recordset("intCEPC"))
            PreencherCampo rchComeco, "|Strlogradouro|", gstrENulo(.Recordset("Strlogradouro"))
            PreencherCampo rchComeco, "|strNumero|", gstrENulo(.Recordset("strNumero"))
            PreencherCampo rchComeco, "|strComplemento|", gstrENulo(.Recordset("strComplemento"))
            PreencherCampo rchComeco, "|strBairro|", gstrENulo(.Recordset("strBairro"))
            PreencherCampo rchComeco, "|strMunicipio|", gstrENulo(.Recordset("strMunicipio"))
            PreencherCampo rchComeco, "|strUF|", gstrENulo(.Recordset("strUF"))
            PreencherCampo rchComeco, "|intCEP|", gstrENulo(.Recordset("intCEP"))
            PreencherCampo rchComeco, "|strFundamento|", gstrENulo(.Recordset("strFundamento"))
            PreencherCampo rchComeco, "|RG|", gstrENulo(.Recordset("RG"))
            PreencherCampo rchComeco, "|CPF|", gstrENulo(.Recordset("CPF"))
            
            
            varAux = gstrFormataInscricao(Right(gstrENulo(.Recordset("strInscricao")), gintRetornaTamanhoMascara(gstrENulo(.Recordset("intUtilizacao")))), gstrENulo(.Recordset("intUtilizacao")))
            PreencherCampo rchComeco, "|strInscricao|", CStr(varAux)
            PreencherCampo rchComeco, "|strNumeroAviso|", gstrConvVrDoSql(Val(gstrENulo(.Recordset("strNumeroAviso"))), 0)
            PreencherCampo rchComeco, "|intExercicio|", gstrENulo(.Recordset("intExercicio"))
            PreencherCampo rchComeco, "|Strcomposicaodareceita|", gstrENulo(.Recordset("Strcomposicaodareceita"))
            PreencherCampo rchComeco, "|Intparcela|", gstrENulo(.Recordset("Intparcela"))
            PreencherCampo rchComeco, "|Dtmdtvencimento|", gstrENulo(.Recordset("Dtmdtvencimento"))
            PreencherCampo rchComeco, "|Dblvloriginal|", gstrENulo(.Recordset("Dblvloriginal"))
            PreencherCampo rchComeco, "|Dblvlprincipal|", gstrENulo(.Recordset("Dblvlprincipal"))
            PreencherCampo rchComeco, "|Dblvlcorrecao|", gstrENulo(.Recordset("Dblvlcorrecao"))
            PreencherCampo rchComeco, "|Dblvlmulta|", gstrENulo(.Recordset("Dblvlmulta"))
            PreencherCampo rchComeco, "|Dblvljuros|", gstrENulo(.Recordset("Dblvljuros"))
            PreencherCampo rchComeco, "|Dblvltotal|", gstrENulo(.Recordset("Dblvltotal"))
            PreencherCampo rchComeco, "|Dblvlindexador|", gstrConvVrDoSql(.Recordset("Dblvlindexador"), 4)
            PreencherCampo rchComeco, "|Dtmdtcalculopeticao|", gstrENulo(.Recordset("Dtmdtcalculopeticao"))
            PreencherCampo rchComeco, "|Strnumdistribuidor|", gstrENulo(.Recordset("Strnumdistribuidor"))
            PreencherCampo rchComeco, "|Controle|", gstrENulo(.Recordset("Controle"))
            PreencherCampo rchComeco, "|strContribuinte|", gstrENulo(.Recordset("strContribuinte"))
            PreencherCampo rchComeco, "|dblQuantidadeIndexador|", gstrConvVrDoSql(.Recordset("dblQuantidadeIndexador"), 4)
            
            
            varAux = gstrENulo(.Recordset("Dtmdtcalculopeticao"))
            PreencherCampo rchComeco, "|dtmPorExtenso|", IIf(varAux <> "", gstrDataPorExtenso(CStr(varAux)), "")
        End With
    End If
    
    rchFim.LoadFile gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\CertidaoDAFim.rtf", rtfRTF
    
    If adoDataControl.NRecords > 0 Then
        With adoDataControl
            PreencherCampo rchFim, "|intCertidao|", gstrENulo(.Recordset("intCertidao"))
            PreencherCampo rchFim, "|dtmDtInscricao|", gstrENulo(.Recordset("dtmDtInscricao"))
            PreencherCampo rchFim, "|strTipoLogradouroC|", gstrENulo(.Recordset("strTipoLogradouroC"))
            PreencherCampo rchFim, "|strTitLogradouroC|", gstrENulo(.Recordset("strTitLogradouroC"))
            PreencherCampo rchFim, "|strLogradouroC|", gstrENulo(.Recordset("strLogradouroC"))
            PreencherCampo rchFim, "|strNumeroC|", gstrENulo(.Recordset("strNumeroC"))
            PreencherCampo rchFim, "|strComplementoC|", gstrENulo(.Recordset("strComplementoC"))
            PreencherCampo rchFim, "|strBairroC|", gstrENulo(.Recordset("strBairroC"))
            PreencherCampo rchFim, "|strMunicipioC|", gstrENulo(.Recordset("strMunicipioC"))
            PreencherCampo rchFim, "|strUFC|", gstrENulo(.Recordset("strUFC"))
            PreencherCampo rchFim, "|intCEPC|", gstrENulo(.Recordset("intCEPC"))
            PreencherCampo rchFim, "|Strlogradouro|", gstrENulo(.Recordset("Strlogradouro"))
            PreencherCampo rchFim, "|strNumero|", gstrENulo(.Recordset("strNumero"))
            PreencherCampo rchFim, "|strComplemento|", gstrENulo(.Recordset("strComplemento"))
            PreencherCampo rchFim, "|strBairro|", gstrENulo(.Recordset("strBairro"))
            PreencherCampo rchFim, "|strMunicipio|", gstrENulo(.Recordset("strMunicipio"))
            PreencherCampo rchFim, "|strUF|", gstrENulo(.Recordset("strUF"))
            PreencherCampo rchFim, "|intCEP|", gstrENulo(.Recordset("intCEP"))
            PreencherCampo rchFim, "|strFundamento|", gstrENulo(.Recordset("strFundamento"))
            
            varAux = gstrFormataInscricao(Right(gstrENulo(.Recordset("strInscricao")), gintRetornaTamanhoMascara(gstrENulo(.Recordset("intUtilizacao")))), gstrENulo(.Recordset("intUtilizacao")))
            PreencherCampo rchFim, "|strInscricao|", CStr(varAux)
            PreencherCampo rchFim, "|strNumeroAviso|", Val(gstrENulo(.Recordset("strNumeroAviso")))
            PreencherCampo rchFim, "|intExercicio|", gstrENulo(.Recordset("intExercicio"))
            PreencherCampo rchFim, "|Strcomposicaodareceita|", gstrENulo(.Recordset("Strcomposicaodareceita"))
            PreencherCampo rchFim, "|Intparcela|", gstrENulo(.Recordset("Intparcela"))
            PreencherCampo rchFim, "|Dtmdtvencimento|", gstrENulo(.Recordset("Dtmdtvencimento"))
            PreencherCampo rchFim, "|Dblvloriginal|", gstrENulo(.Recordset("Dblvloriginal"))
            PreencherCampo rchFim, "|Dblvlprincipal|", gstrENulo(.Recordset("Dblvlprincipal"))
            PreencherCampo rchFim, "|Dblvlcorrecao|", gstrENulo(.Recordset("Dblvlcorrecao"))
            PreencherCampo rchFim, "|Dblvlmulta|", gstrENulo(.Recordset("Dblvlmulta"))
            PreencherCampo rchFim, "|Dblvljuros|", gstrENulo(.Recordset("Dblvljuros"))
            PreencherCampo rchFim, "|Dblvltotal|", gstrENulo(.Recordset("Dblvltotal"))
            PreencherCampo rchFim, "|Dblvlindexador|", gstrConvVrDoSql(.Recordset("Dblvlindexador"), 4)
            PreencherCampo rchFim, "|Dtmdtcalculopeticao|", gstrENulo(.Recordset("Dtmdtcalculopeticao"))
            PreencherCampo rchFim, "|Strnumdistribuidor|", gstrENulo(.Recordset("Strnumdistribuidor"))
            PreencherCampo rchFim, "|Controle|", gstrENulo(.Recordset("Controle"))
            PreencherCampo rchFim, "|strContribuinte|", gstrENulo(.Recordset("strContribuinte"))
            'PreencherCampo rchFim, "|dblQuantidadeIndexador|", gstrConvVrDoSql(.Recordset("dblQuantidadeIndexador"), 4)
            dblQuantidadeIndexador = .Recordset("Dblvlindexador")
            varAux = gstrENulo(.Recordset("Dtmdtcalculopeticao"))
            PreencherCampo rchFim, "|dtmPorExtenso|", IIf(varAux <> "", gstrDataPorExtenso(CStr(varAux)), "")
        End With
    End If
    
End Sub

Private Sub PageFooter_Format()
'    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
'    lblDataHora = gstrDataDoSistema(True, , True)
    
    
    
End Sub

Private Sub Detail_Format()
    'TrocaCorParaZebrado lblSombra

        
    If adoDataControl.NRecords > 0 Then
        'fldInscricao = gstrFormataInscricao(Right(gstrENulo(adodatacontrol.Recordset("strInscricao").Value), gintRetornaTamanhoMascara(gstrENulo(adodatacontrol.Recordset("intUtilizacao").Value))), gstrENulo(adodatacontrol.Recordset("intUtilizacao").Value))
    End If
        
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

Private Sub PreencherCampo(RichEdit As Object, Campo As String, Valor As String)
Dim intPosicao As Integer

    intPosicao = RichEdit.Find(Campo, 0, -1, rtfWholeWord, -1)
    
    If intPosicao > -1 Then
        
        RichEdit.SelStart = intPosicao
        RichEdit.SelLength = Len(Campo)
        RichEdit.Clear
        
        RichEdit.InsertField Campo, intPosicao
        
        RichEdit.ReplaceField Campo, Valor
        
    End If
    


End Sub

