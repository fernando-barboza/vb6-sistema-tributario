VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioDeQuantidadeDeLancamentoValorTipo 
   Caption         =   "Relat�rio de Quantidade de Lan�amentos, Valor e Tipo"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   18865
   _ExtentY        =   12541
   SectionData     =   "RelatorioDeQuantidadeDeLancamentoValorTipo.dsx":0000
End
Attribute VB_Name = "rptRelatorioDeQuantidadeDeLancamentoValorTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoResultado                        As ADODB.Recordset
Dim intTotalDeOcorrencias As Integer
Dim dblValorTotal         As Double

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
    intTotalDeOcorrencias = 0
    dblValorTotal = 0
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
    txtSomaPorLancamento = gstrConvVrDoSql(txtSomaPorLancamento)
    If txtSomaPorLancamento.Text <> "" Then
        intTotalDeOcorrencias = intTotalDeOcorrencias + 1
        dblValorTotal = dblValorTotal + CDbl(txtSomaPorLancamento)
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "P�gina " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    txt_TotalDeOcorrencias = intTotalDeOcorrencias
    txt_TotalDeValor = gstrConvVrDoSql(dblValorTotal)
    dblValorTotal = 0
    intTotalDeOcorrencias = 0
    MostraEmissorRelatorio Me
End Sub
