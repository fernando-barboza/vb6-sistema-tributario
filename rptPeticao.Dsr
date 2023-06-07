VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptPeticao 
   Caption         =   "Tributario - rptPeticao (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptPeticao.dsx":0000
End
Attribute VB_Name = "rptPeticao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
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
    PreencherCampo rchFim, "|Valor|", fldTotalGeral
    PreencherCampo rchFim, "|ValorExtenso|", gstrExtenso(fldTotalGeral)
End Sub

Private Sub GroupHeader1_Format()
Dim varAux

    rchComeco.LoadFile gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\peticao.rtf", rtfRTF

    If adoDataControl.NRecords > 0 Then
        With adoDataControl
             PreencherCampo rchComeco, "|CertidaoDividaAtiva|", gstrCertidaoPorExecutivo(.Recordset("PKID"))
             PreencherCampo rchComeco, "|Devedor|", gstrENulo(.Recordset("STREXECUTADONOME"))
             PreencherCampo rchComeco, "|TipoDeLogradouro|", gstrENulo(.Recordset("STREXECUTADOTPLOGNOTIF"))
             PreencherCampo rchComeco, "|TituloDeLogradouro|", gstrENulo(.Recordset("STREXECUTADOTPLOGNOTIF"))
             PreencherCampo rchComeco, "|Logradouro|", gstrENulo(.Recordset("STREXECUTADONOMELOGNOTIF"))
             PreencherCampo rchComeco, "|Numero|", gstrENulo(.Recordset("STREXECUTADONUMLOGNOTIF"))
             PreencherCampo rchComeco, "|Bairro|", gstrENulo(.Recordset("STREXECUTADOBAIRRONOTIF"))
             PreencherCampo rchComeco, "|Cidade|", gstrENulo(.Recordset("STREXECUTADOCIDNOTIF"))
             PreencherCampo rchComeco, "|CEP|", gstrENulo(.Recordset("INTEXECUTADOCEPNOTIF"))
        End With
    End If
    
    rchFim.LoadFile gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\peticaoFim.rtf", rtfRTF
    
    If adoDataControl.NRecords > 0 Then
        With adoDataControl
             PreencherCampo rchFim, "|ValorIndexador|", gstrConvVrDoSql(.Recordset("DBLQUANTINDEXADOR"), 4)
             PreencherCampo rchFim, "|Indexador|", gstrENulo(.Recordset("STRINDEXADORDESCR"))
             
             varAux = gstrENulo(.Recordset("dtmDtCalculoPeticao"))
             PreencherCampo rchFim, "|DataExtenso|", IIf(varAux <> "", gstrDataPorExtenso(CStr(varAux)), "")
            'PreencherCampo rchFim, "|DataExtenso|", gstrDataPorExtenso(gstrENulo(.Recordset("dtmDtCalculoPeticao")))
             PreencherCampo rchFim, "|Controle|", gstrENulo(.Recordset("Controle"))
        End With
    End If
    
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)

End Sub

Private Sub Detail_Format()
    'TrocaCorParaZebrado lblSombra
        
    If adoDataControl.NRecords > 0 Then
        fldInscricao = gstrFormataInscricao(Right(gstrENulo(adoDataControl.Recordset("strInscricao").Value), gintRetornaTamanhoMascara(gstrENulo(adoDataControl.Recordset("intUtilizacao").Value))), gstrENulo(adoDataControl.Recordset("intUtilizacao").Value))
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
