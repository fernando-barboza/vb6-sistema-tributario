VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptFluxodeCaixa 
   Caption         =   "prjOrcamentario - rptFluxodeCaixa (ActiveReport)"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "FluxodeCaixa.dsx":0000
End
Attribute VB_Name = "rptFluxodeCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dblReceitasCorrentes    As Double
Dim dblDespesasCorrentes    As Double
Dim strDescricao            As String

Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
    With frmRelatorioPeriodo
        lblPeriodo = "Período: " & .txtdtmInicial & " a " & .txtdtmFinal
    End With
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
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtstrEstado
    lbl_Titulo = Me.Caption
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

Private Sub Detail_Format()
    txtstrDescricao.Left = Abs(Val(txtbytNivel) + 4) * 200
    TrocaCorParaZebrado lblSombra
    If Val(txtdblValor) = 0 Then
        txtdblValor.Visible = False
    Else
        txtdblValor.Visible = True
    End If
    
    If Val(txtdblTotal) = 0 Then
        txtdblTotal.Visible = False
    Else
        txtdblTotal.Visible = True
    End If
    strDescricao = UCase(txtstrDescricao)
    If strDescricao = "Receitas Correntes" Then
       dblReceitasCorrentes = Val(gstrConvVrParaSql(txtdblTotal))
    End If
    If strDescricao = "Despesas Correntes" Then
       dblDespesasCorrentes = Val(gstrConvVrParaSql(txtdblTotal))
    End If
End Sub

Private Sub grhOrgao_Format()
    TrocaCorParaZebrado lblSombraOrgao
End Sub

Private Sub grhConta_Format()
    TrocaCorParaZebrado lblSombraGrupo
    If Val(txtbytNatureza) = 0 Then
        lbl_DescricaoReceita = "RECEITAS"
    Else
        lbl_DescricaoReceita = "DESPESAS"
    End If
    
    If Val(txtbytNatureza) = 0 Then
        txtdblDespesa.Visible = False
        txtdblReceita.Visible = True
    Else
        txtdblDespesa.Visible = True
        txtdblReceita.Visible = False
    End If

End Sub

Private Sub grhSaldo_Format()
    TrocaCorParaZebrado lblSombraSaldoAnterior
    TrocaCorParaZebrado lblSombraSaldoLiquido
    TrocaCorParaZebrado lblSombraSaldoTotal
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
