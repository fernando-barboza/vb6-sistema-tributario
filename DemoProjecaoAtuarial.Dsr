VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptDemoProjecaoAtuarial 
   Caption         =   "prjOrcamentario - rptDemoProjecaoAtuarial (ActiveReport)"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   Tag             =   "341"
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   12144
   SectionData     =   "DemoProjecaoAtuarial.dsx":0000
End
Attribute VB_Name = "rptDemoProjecaoAtuarial"
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
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf tool.ID = 16 Then
        Configura_Relatorio Me, True
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "P�gina " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
   TrocaCorParaZebrado lblSombra
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
