VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptTipoDeLogradouro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tributario - rptTipoDeLogradouro (ActiveReport)"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   13845
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   24421
   _ExtentY        =   15161
   SectionData     =   "CadTipoDeLogradouro.dsx":0000
End
Attribute VB_Name = "rptTipoDeLogradouro"
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
    LeImagemLogotipo imgBrasao, imgLogoTipo, fld_NomeFantasia, fld_Estado
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
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
'    lbl_Total = gstrTotalDeRegistros(gstrTipoLogradouro, lblRelatorio)
    MostraEmissorRelatorio Me
End Sub


