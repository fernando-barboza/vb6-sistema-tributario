VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptExtratoIndividualizadoDeLancamento 
   Caption         =   "Tributario - rptExtratoIndividualizadoDeLancamento (ActiveReport)"
   ClientHeight    =   6975
   ClientLeft      =   -3585
   ClientTop       =   1845
   ClientWidth     =   9660
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   17039
   _ExtentY        =   12303
   SectionData     =   "ExtratoIndividualizadoDeLancamento.dsx":0000
End
Attribute VB_Name = "rptExtratoIndividualizadoDeLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim intFlag As Integer
    Dim intContador As Integer
    Dim intContadorComposicao As Integer

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
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
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
    TrocaCorParaZebrado lblSombra
    txtdblValor = gstrConvVrDoSql(txtdblValor)
    
    If Val(txtintNumeroParcela) > 0 Then
        txtdblLancadoComposicao = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblLancadoComposicao)) + Val(gstrConvVrParaSql(txtdblValor)))
    End If
    intContador = intContador + 1
    intContadorComposicao = intContadorComposicao + 1
    
    If Trim(txtdtmDataPagamento.Text) = "" Then
        
        txtdblLancado = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblLancado)) + Val(gstrConvVrParaSql(txtdblValor)))
    Else
        txtdblPagoComposicao = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblPagoComposicao)) + Val(gstrConvVrParaSql(txtdblValor)))
        txtdblPago = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblPago)) + Val(gstrConvVrParaSql(txtdblValor)))
    End If
    txtdtmDataPagamento = gstrDataFormatada(txtdtmDataPagamento)
    txtDataVencimento = gstrDataFormatada(txtDataVencimento)
End Sub

Private Sub GroupFooter1_Format()
    txtintContador = intContador
    txtdblAberto = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblLancado)) - Val(gstrConvVrParaSql(txtdblPago)))
    TrocaCorParaZebrado lblSombra2
    TrocaCorParaZebrado lblsombra4
End Sub

Private Sub GroupFooter2_Format()
    txtintContadorComposicao = intContadorComposicao
'    txtstrComposicao2 = txtstrComposicao
    txtdblAbertoComposicao = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblLancadoComposicao)) - Val(gstrConvVrParaSql(txtdblPagoComposicao)))
    TrocaCorParaZebrado lblSombra6
    TrocaCorParaZebrado lblSombra7
    TrocaCorParaZebrado lblSombra8
End Sub

Private Sub GroupHeader1_Format()
    intContador = 0
    txtdblLancado.Text = "0,00"
    txtdblPago.Text = "0,00"
    TrocaCorParaZebrado lblSombra1
    txtstrCNPJCPF = gstrCGCCPFFormatado(txtstrCNPJCPF)
End Sub

Private Sub GroupHeader2_Format()
    TrocaCorParaZebrado lblsombra5
    txtdblLancadoComposicao.Text = "0,00"
    txtdblPagoComposicao.Text = "0,00"
    intContadorComposicao = 0
End Sub

Private Sub GroupHeader3_Format()
    TrocaCorParaZebrado lblSombra9
    TrocaCorParaZebrado lblSombra3
    txtdtmData = gstrDataFormatada(txtdtmData)
'    If txtbytLancPagamento.Text = "0" Then
'        txtstrLancamento.Width = 4230
'        txtstrLancamento.Left = 1980
'        intFlag = 0
'    Else
'        txtstrLancamento.Width = 2250
'        txtstrLancamento.Left = 3960
'        intFlag = 1
'    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
