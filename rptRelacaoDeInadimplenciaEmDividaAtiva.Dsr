VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelacaoDeInadimplenciaEmDividaAtiva 
   Caption         =   "Relação de Inadimplência em Dívida Ativa"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   MDIChild        =   -1  'True
   _ExtentX        =   19129
   _ExtentY        =   13891
   SectionData     =   "rptRelacaoDeInadimplenciaEmDividaAtiva.dsx":0000
End
Attribute VB_Name = "rptRelacaoDeInadimplenciaEmDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Contribuintes As Integer
Dim ValorTotal    As Double

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
    Contribuintes = 0
    ValorTotal = 0
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
    If txtValorDevido <> "" Then
        Contribuintes = Contribuintes + 1
        ValorTotal = ValorTotal + txtValorDevido
        txtValorDevido = gstrConvVrDoSql(txtValorDevido)
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    Total = Contribuintes
    TotalDebitos = gstrConvVrDoSql(ValorTotal)
    MostraEmissorRelatorio Me
End Sub
