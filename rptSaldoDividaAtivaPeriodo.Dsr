VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSaldoDividaAtivaPeriodo 
   Caption         =   "Tributario - rptSaldoDividaAtivaPeriodo (ActiveReport)"
   ClientHeight    =   11235
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19817
   SectionData     =   "rptSaldoDividaAtivaPeriodo.dsx":0000
End
Attribute VB_Name = "rptSaldoDividaAtivaPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblQtdeInscrE As Double
Dim dblValorInscrE As Double
Dim dblQtdePgE As Double
Dim dblValorPgE As Double

Dim dblQtdeInscrT As Double
Dim dblValorInscrT As Double
Dim dblQtdePgT As Double
Dim dblValorPgT As Double
  
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
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    dblQtdeInscrE = 0
    dblValorInscrE = 0
    dblQtdePgE = 0
    dblValorPgE = 0
    
    dblQtdeInscrT = 0
    dblValorInscrT = 0
    dblQtdePgT = 0
    dblValorPgT = 0
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

Private Sub GroupFooter1_Format()
    TrocaCorParaZebrado lblSombra4
    
    txtQtdeInscrT = dblQtdeInscrT
    txtValorInscrT = gstrConvVrDoSql(dblValorInscrT, 2)
    txtQtdePgT = dblQtdePgT
    txtValorPgT = gstrConvVrDoSql(dblValorPgT, 2)
End Sub

Private Sub GroupFooter2_Format()
    TrocaCorParaZebrado lblSombra3
    
    txtQtdeInscrE = dblQtdeInscrE
    txtValorInscrE = gstrConvVrDoSql(dblValorInscrE, 2)
    txtQtdePgE = dblQtdePgE
    txtValorPgE = gstrConvVrDoSql(dblValorPgE, 2)
    
    dblQtdeInscrT = dblQtdeInscrT + dblQtdeInscrE
    dblValorInscrT = dblValorInscrT + CDbl(dblValorInscrE)
    dblQtdePgT = dblQtdePgT + dblQtdePgE
    dblValorPgT = dblValorPgT + CDbl(dblValorPgE)
    
    dblQtdeInscrE = 0
    dblValorInscrE = 0
    dblQtdePgE = 0
    dblValorPgE = 0
End Sub

Private Sub GroupHeader3_Format()
    TrocaCorParaZebrado lblSombra
    
    txtValorInscr.Text = gstrConvVrDoSql(txtValorInscr.Text, 2)
    txtValorPg.Text = gstrConvVrDoSql(txtValorPg.Text, 2)

    dblQtdeInscrE = dblQtdeInscrE + IIf(txtQtdeInscr = "", 0, txtQtdeInscr)
    dblValorInscrE = dblValorInscrE + CDbl(IIf(txtValorInscr = "", 0, txtValorInscr))
    dblQtdePgE = dblQtdePgE + IIf(txtQtdePg = "", 0, txtQtdePg)
    dblValorPgE = dblValorPgE + CDbl(IIf(txtValorPg = "", 0, txtValorPg))

End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
    MostraEmissorRelatorio Me
End Sub
