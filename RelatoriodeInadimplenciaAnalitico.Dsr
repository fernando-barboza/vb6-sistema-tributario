VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatoriodeInadimplenciaAnalitico 
   Caption         =   "Tributario - rptRelatoriodeInadimplenciaAnalitico (ActiveReport)"
   ClientHeight    =   8085
   ClientLeft      =   75
   ClientTop       =   1155
   ClientWidth     =   11430
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20161
   _ExtentY        =   14261
   SectionData     =   "RelatoriodeInadimplenciaAnalitico.dsx":0000
End
Attribute VB_Name = "rptRelatoriodeInadimplenciaAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblValorReceita As Double
Dim dblValorContribuinte As Double

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
    txtdtmDataVencimento = gstrDataFormatada(txtdtmDataVencimento, False)
    txtdblValorParcela = gstrConvVrDoSql(txtdblValorParcela, 2)
    
    If Trim(txtdblValorParcela) <> "" Then
        dblValorReceita = dblValorReceita + CDbl(txtdblValorParcela)
        dblValorContribuinte = dblValorContribuinte + CDbl(txtdblValorParcela)
    End If
    
End Sub

Private Sub GroupFooter1_Format()
txtdblTotalPorContribuinte = gstrConvVrDoSql(dblValorContribuinte, 2)
End Sub

Private Sub GroupFooter3_Format()
txtTotaldaReceita = gstrConvVrDoSql(dblValorReceita)
End Sub

Private Sub GroupHeader1_Format()
dblValorReceita = 0
dblValorContribuinte = 0
End Sub

Private Sub GroupHeader3_Format()
    dblValorReceita = 0
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

