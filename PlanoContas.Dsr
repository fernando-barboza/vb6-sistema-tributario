VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptPlanoContas 
   Caption         =   "prjOrcamentario - rptPlanoContas (ActiveReport)"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11745
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   20717
   _ExtentY        =   11007
   SectionData     =   "PlanoContas.dsx":0000
End
Attribute VB_Name = "rptPlanoContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_ReportEnd()
    Dim intInd  As Integer
    For intInd = 0 To Me.Pages.Count - 1
        Me.Pages(intInd).Orientation = ddOLandscape
    Next
End Sub

Private Sub ActiveReport_ReportStart()
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal tool As DDActiveReports2.DDTool)
    If tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf tool.ID = 16 Then
        Configura_Relatorio Me, True
    End If
End Sub

Private Sub Detail_Format()
    txtstrConta = gvntFormatacaoEspecifica(txtstrConta, 1)
    txtFinanceira = gstrSimOuNao(chkFinanceira)
    txtRetencao = gstrSimOuNao(chkRetencao)
    txtExtraOrcamentaria = gstrSimOuNao(chkExtraOrcamentaria)
    txtIntegraBalanco = gstrSimOuNao(chkIntegraBalanco)
    txtRetificadora = gstrSimOuNao(chkRetificadora)
    txtInversaoDeSaldo = gstrSimOuNao(chkInversaoDeSaldo)
    txtFinanceira = gstrSimOuNao(chkInversaoDeSaldo)
    txtAnalitica = gstrSimOuNao(chkAnalitica)
    txtEducacao = gstrSimOuNao(chkEducacao)
    txtSaude = gstrSimOuNao(chkSaude)
    txtFundef = gstrSimOuNao(chkFundef)
    txtPessoal = gstrSimOuNao(chkPessoal)
    txtDeduzEducacao = gstrSimOuNao(chkDeduzEducacao)
    txtDeduzSaude = gstrSimOuNao(chkDeduzSaude)
    txtDeduzFundef = gstrSimOuNao(chkDeduzFundef)
    txtDeduzPessoal = gstrSimOuNao(chkDeduzPessoal)
    txtSaldo = gstrConvVrDoSql(Val(gstrConvVrParaSql(flddblValorEmpenho)) - Val(gstrConvVrParaSql(flddblValor)))
    If Val(fldbytNatureza) Then
        txtNatureza = "Devedora"
    Else
        txtNatureza = "Credora"
    End If
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub GroupFooter1_Format()
    txtTotalSaldo = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtTotalValorEmpenho)) - Val(gstrConvVrParaSql(txtTotalValor)))
    TrocaCorParaZebrado lbl_Sombra2
End Sub

Private Sub GroupHeader1_Format()
    txtSaldo = gstrConvVrDoSql(flddblValor)
    TrocaCorParaZebrado lbl_Sombra1
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    txtGeralSaldo = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtGeralValorEmpenho)) - Val(gstrConvVrParaSql(txtGeralValor)))
    MostraEmissorRelatorio Me
End Sub
