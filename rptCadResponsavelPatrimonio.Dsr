VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCadResponsavelPatrimonio 
   Caption         =   "prjOrcamentario - rptCadResponsavelPatrimonio (ActiveReport)"
   ClientHeight    =   9195
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   11535
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20346
   _ExtentY        =   16219
   SectionData     =   "rptCadResponsavelPatrimonio.dsx":0000
End
Attribute VB_Name = "rptCadResponsavelPatrimonio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PkidGrp As Integer
Dim intLeft As Integer

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_Activate()
    If UCase(MDIMenu.Tag) = "OUVIDORIA" Then
        lblstrResponsavel.Caption = "Funcionário"
        lblRelatorio.Caption = "Relação de Funcionários (Ouvidoria)"
    End If
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
    'lbl_Titulo = Me.Caption
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
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
End Sub

Private Sub PageFooter_Format()
    lblPagina = "Página " & pageNumber
    MostraEmissorRelatorio Me
End Sub


