VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioDeDiferencaNosPagamentos 
   Caption         =   "Relat�rio das Parcelas Lan�adas"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "RelatorioDeDiferencaNosPagamentos.dsx":0000
End
Attribute VB_Name = "rptRelatorioDeDiferencaNosPagamentos"
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

Private Sub ActiveReport_ReportEnd()
Dim i As Integer
For i = 0 To rptRelatorioDeDiferencaNosPagamentos.Pages.Count - 1
    rptRelatorioDeDiferencaNosPagamentos.Pages(i).Orientation = ddOLandscape
Next
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
    
    txtdblDevido = gstrConvVrDoSql(txtdblDevido, 2)
    txtdblTotalPago = gstrConvVrDoSql(txtdblTotalPago, 2)
    txtdblDiferenca = gstrConvVrDoSql(txtdblDiferenca, 2)
    txtdtmDataVencimento = gstrDataFormatada(txtdtmDataVencimento, False)
    txtdtmDataPagamento = gstrDataFormatada(txtdtmDataPagamento, False)

End Sub


Private Sub PageFooter_Format()
    lblPagina.Caption = "P�gina " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
