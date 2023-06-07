VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRecComposicaoPago 
   Caption         =   "Tributario - rptRecComposicaoPago (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptRecComposicaoPago.dsx":0000
End
Attribute VB_Name = "rptRecComposicaoPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intUtilizacao As Byte

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
    
   txtTotalLancado = "0,00"
   txtTotalPago = "0,00"
   txtTotalLancamentos = "0"
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
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
   Dim intUtil As Integer
   Dim dblPorcentagem As Double
   TrocaCorParaZebrado lblSombra
   intUtil = intUtilizacao
   
   
   
   If Val(gstrConvVrParaSql(txtvalorPagto)) <> 0 Then
        dblPorcentagem = Val(gstrConvVrParaSql(txtValorLancado)) / Val(gstrConvVrParaSql(txtdblLancamentoValor))
        txtvalorPagto = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtvalorPagto)) * dblPorcentagem)
   End If
   
   
   'txtstrInscricao = gstrFormataInscricao(Right(Trim(txtstrInscricao), gintRetornaTamanhoMascara(intUtilizacao)), intUtil)
   txtstrInscricao = gstrFormataInscricao(Trim(txtstrInscricao), intUtil)
   txtTotalLancado = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtTotalLancado)) + Val(gstrConvVrParaSql(txtValorLancado)))
   txtdtpagto = gstrDataFormatada(txtdtpagto)
   txtTotalPago = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtTotalPago)) + Val(gstrConvVrParaSql(txtvalorPagto)))
   
   
   
   
   txtValorLancado = gstrConvVrDoSql(txtValorLancado)
   txtvalorPagto = gstrConvVrDoSql(txtvalorPagto)
   txtTotalLancamentos = gstrConvVrParaSql(Val(txtTotalLancamentos)) + 1
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

