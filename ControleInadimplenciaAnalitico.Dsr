VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptControleInadimplenciaAnalitico 
   Caption         =   "Relatório para Controle de Inadimplência Analítico"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20479
   _ExtentY        =   15161
   SectionData     =   "ControleInadimplenciaAnalitico.dsx":0000
End
Attribute VB_Name = "rptControleInadimplenciaAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblTotalParcial            As Double
Dim dblTotalGeral              As Double
Dim intQuantidadeContribuintes As Integer

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
    dblTotalParcial = 0
    dblTotalGeral = 0
    intQuantidadeContribuintes = 0
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
    If txtdblValorParcela.Text <> "" Then
        txtdblValorParcela.Text = gstrConvVrDoSql(txtdblValorParcela.Text)
        dblTotalParcial = dblTotalParcial + CDbl(txtdblValorParcela.Text)
    End If
End Sub

Private Sub GroupFooter1_Format()
    txt_TotalParcial.Text = gstrConvVrDoSql(dblTotalParcial)
    dblTotalGeral = dblTotalGeral + dblTotalParcial
    dblTotalParcial = 0
End Sub

Private Sub GroupHeader1_Format()
    If txtstrNome.Text <> "" Then
        intQuantidadeContribuintes = intQuantidadeContribuintes + 1
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    txt_TotalDeContribuintes.Text = Val(intQuantidadeContribuintes)
    txt_TotalGeral.Text = gstrConvVrDoSql(dblTotalGeral)
    intQuantidadeContribuintes = 0
    dblTotalGeral = 0
    MostraEmissorRelatorio Me
End Sub
