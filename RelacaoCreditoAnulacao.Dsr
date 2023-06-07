VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelacaoCreditoAnulacao 
   Caption         =   "prjOrcamentario - rptRelacaoCreditoAnulacao (ActiveReport)"
   ClientHeight    =   10740
   ClientLeft      =   390
   ClientTop       =   1920
   ClientWidth     =   15210
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   26829
   _ExtentY        =   18944
   SectionData     =   "RelacaoCreditoAnulacao.dsx":0000
End
Attribute VB_Name = "rptRelacaoCreditoAnulacao"
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

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_ReportStart()
    With frmRelatorioPeriodo
        lbl_Periodo = "Período: " & .txtdtmInicial & " à " & .txtdtmFinal
        Me.Caption = .Caption
    End With
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtstrEstado
    lblRelatorio = Me.Caption
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

Private Sub grhRelacao_Format()
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra1
End Sub

Private Sub ActiveReport_ReportEnd()
    Dim i As Integer
    For i = 0 To Me.Pages.Count - 1
        Me.Pages(i).Orientation = ddOLandscape
    Next
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
