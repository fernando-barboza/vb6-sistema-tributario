VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptDiasNaoUteis 
   Caption         =   "Relatório de Dias Não Úteis"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   18733
   _ExtentY        =   14288
   SectionData     =   "rptDiasNaoUteis.dsx":0000
End
Attribute VB_Name = "rptDiasNaoUteis"
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
    If Val(txtbytTipo) = 0 Then
        lbl_tipo.Caption = "Feriado"
    ElseIf Val(txtbytTipo) = 1 Then
        lbl_tipo.Caption = "Sábado"
    ElseIf Val(txtbytTipo) = 2 Then
        lbl_tipo.Caption = "Domingo"
    End If
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub GroupHeader1_Format()
   ' TrocaCorParaZebrado lblSombra
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    lbl_Total = gstrTotalDeRegistros(gstrDiasNaoUteis, lblRelatorio)
    MostraEmissorRelatorio Me
End Sub



