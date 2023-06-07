VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioDeParcelasArrecadas 
   Caption         =   "Relatório das Parcelas Lançadas"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "RelatorioDeParcelasArrecadas.dsx":0000
End
Attribute VB_Name = "rptRelatorioDeParcelasArrecadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim intFlag As Integer
    Dim intContador As Integer
    Dim intContadorComposicao As Integer

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
    
    intContadorComposicao = intContadorComposicao + 1
    'txtdblLancadoComposicao = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblLancadoComposicao)) + Val(gstrConvVrParaSql(txtdblValor)))
    intContador = intContador + 1
    'txtdblLancado = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblLancado)) + Val(gstrConvVrParaSql(txtdblValor)))
    
    
    txtdblValorParcela = gstrConvVrDoSql(txtdblValorParcela, 2)
    txtdblJuros = gstrConvVrDoSql(txtdblJuros, 2)
    txtdblMulta = gstrConvVrDoSql(txtdblMulta, 2)
    txtdblDiferenca = gstrConvVrDoSql(txtdblDiferenca, 2)
    txtdblTotalPago = gstrConvVrDoSql(txtdblTotalPago, 2)
    
    If txtstrTipoPagamento = "Normal" Then
        txtdblDevido = CDbl(txtdblDevido) + CDbl(txtdblValorParcela)
        txtdblDevJuros = CDbl(txtdblDevJuros) + CDbl(txtdblJuros)
        txtdblDevMulta = CDbl(txtdblDevMulta) + CDbl(txtdblMulta)
        txtdblDevDif = CDbl(txtdblDevDif) + CDbl(txtdblDiferenca)
        If txtdblDevTotal <> "" Then
            txtdblDevTotal = CDbl(txtdblDevTotal) + CDbl(txtdblTotalPago)
        Else
            txtdblDevTotal = CDbl(txtdblTotalPago)
        End If
    Else
        txtdblIndevido = txtdblDevido + txtdblValorParcela
        txtdblIndJuros = txtdblDevJuros + txtdblJuros
        txtdblIndMulta = txtdblDevMulta + txtdblMulta
        txtdblIndDif = txtdblDevDif + txtdblDiferenca
        If txtdblIndTotal <> "" Then
            txtdblIndTotal = txtdblIndTotal + txtdblTotalPago
        Else
            txtdblIndTotal = txtdblTotalPago
        End If
    End If
    
    
    
    txtdblDevido = gstrConvVrDoSql(txtdblDevido, 2)
    txtdblDevJuros = gstrConvVrDoSql(txtdblDevJuros, 2)
    txtdblDevMulta = gstrConvVrDoSql(txtdblDevMulta, 2)
    txtdblDevDif = gstrConvVrDoSql(txtdblDevDif, 2)
    txtdblDevTotal = gstrConvVrDoSql(txtdblDevTotal, 2)
    
    txtdblIndevido = gstrConvVrDoSql(txtdblIndevido, 2)
    txtdblIndJuros = gstrConvVrDoSql(txtdblIndJuros, 2)
    txtdblIndMulta = gstrConvVrDoSql(txtdblIndMulta, 2)
    txtdblIndDif = gstrConvVrDoSql(txtdblIndDif, 2)
    txtdblIndTotal = gstrConvVrDoSql(txtdblIndTotal, 2)
    
    txtdtmDataPagamento = gstrDataFormatada(txtdtmDataPagamento, False)
    txtdtmDataVencimento = gstrDataFormatada(txtdtmDataVencimento, False)
    txtdtmLancamento = gstrDataFormatada(txtdtmLancamento, False)
End Sub

Private Sub GroupFooter1_Format()

txtdblDevContrib = gstrConvVrDoSql(txtdblDevContrib, 2)
txtdblJurosContrib = gstrConvVrDoSql(txtdblJurosContrib, 2)
txtdblMultaContrib = gstrConvVrDoSql(txtdblMultaContrib, 2)
txtdblDifContrib = gstrConvVrDoSql(txtdblDifContrib, 2)
txtdblTotalContrib = gstrConvVrDoSql(txtdblTotalContrib, 2)

End Sub

Private Sub GroupFooter2_Format()

If txtdblDevido <> "" Then
    txtdblReferente = gstrConvVrDoSql((CDbl(txtdblDevido) + CDbl(txtdblIndevido)), 2)
End If
If txtdblDevJuros <> "" Then
    txtdblRefJuros = gstrConvVrDoSql((CDbl(txtdblDevJuros) + CDbl(txtdblIndJuros)))
End If
If txtdblDevMulta <> "" Then
    txtdblRefMulta = gstrConvVrDoSql((CDbl(txtdblDevMulta) + CDbl(txtdblIndMulta)))
End If
If txtdblDevDif <> "" Then
    txtdblRefDif = gstrConvVrDoSql((CDbl(txtdblDevDif) + CDbl(txtdblIndDif)))
End If
If txtdblDevTotal <> "" Then
    txtdblRefTotal = gstrConvVrDoSql((CDbl(txtdblDevTotal) + CDbl(txtdblIndTotal)))
End If

If txtdblDevContrib <> "" Then
    txtdblDevContrib = CDbl(txtdblDevContrib) + CDbl(txtdblReferente)
Else
    txtdblDevContrib = CDbl(txtdblReferente)
End If
If txtdblJurosContrib <> "" Then
    txtdblJurosContrib = CDbl(txtdblJurosContrib) + CDbl(txtdblRefJuros)
Else
    txtdblJurosContrib = CDbl(txtdblRefJuros)
End If
If txtdblMultaContrib <> "" Then
    txtdblMultaContrib = CDbl(txtdblMultaContrib) + CDbl(txtdblRefMulta)
Else
    txtdblMultaContrib = CDbl(txtdblRefMulta)
End If
If txtdblDifContrib <> "" Then
    txtdblDifContrib = CDbl(txtdblDifContrib) + CDbl(txtdblRefDif)
Else
    txtdblDifContrib = CDbl(txtdblRefDif)
End If
If txtdblTotalContrib <> "" And txtdblRefTotal <> "" Then
    txtdblTotalContrib = CDbl(txtdblTotalContrib) + CDbl(txtdblRefTotal)
ElseIf txtdblRefTotal <> "" Then
    txtdblTotalContrib = CDbl(txtdblRefTotal)
Else
    txtdblRefTotal = "0,00"
End If



'    txtintContadorComposicao = intContadorComposicao
'    txtstrComposicao2 = txtstrComposicao
'    txtdblAbertoComposicao = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblLancadoComposicao)) - Val(gstrConvVrParaSql(txtdblPagoComposicao)))
End Sub

Private Sub GroupHeader1_Format()

txtdblDevContrib = 0
txtdblJurosContrib = 0
txtdblMultaContrib = 0
txtdblDifContrib = 0
txtdblTotalContrib = 0

End Sub

Private Sub GroupHeader2_Format()

txtdblDevido = 0
txtdblDevJuros = 0
txtdblDevMulta = 0
txtdblDevDif = 0

txtdblIndevido = 0
txtdblIndJuros = 0
txtdblIndMulta = 0
txtdblIndDif = 0
txtdblIndTotal = 0
'   txtdblLancadoComposicao.Text = "0,00"
'   txtdblPagoComposicao.Text = "0,00"
'   intContadorComposicao = 0
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
