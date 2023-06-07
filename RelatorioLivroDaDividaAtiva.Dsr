VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioLivroDaDividaAtiva 
   Caption         =   "Relatório do Livro da Dívida Ativa"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RelatorioLivroDaDividaAtiva.dsx":0000
End
Attribute VB_Name = "rptRelatorioLivroDaDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Contribuintes    As Integer

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
Dim TipoCadastro As String
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    Select Case giContador
    Case 0
        TipoCadastro = " - Imobiliário Rural "
    Case 1
        TipoCadastro = " - Imobiliário Urbano "
    Case 2
        TipoCadastro = " - Econômico "
    Case 3
        TipoCadastro = " - Contribuição de Melhorias "
    Case 4
        TipoCadastro = " - Receitas Diversas "
    End Select
    lblRelatorio = Me.Caption + " " + TipoCadastro
    Contribuintes = 0
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
    txtdblValorOriginal = gstrConvVrDoSql(txtdblValorOriginal)
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub GroupFooter1_Format()
    TotalContribuintes = Contribuintes
    TotalGeral = gstrConvVrDoSql(TotalGeral)
End Sub

Private Sub GroupFooter2_Format()
    If txtstrNome <> "" Then
        Contribuintes = Contribuintes + 1
    End If
    Total = gstrConvVrDoSql(Total)
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
