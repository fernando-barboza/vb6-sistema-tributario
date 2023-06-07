VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelacaoPrevisaoReceitaDespesaAnulacaoPeriodo 
   Caption         =   "prjOrcamentario - rptRelacaoPrevisaoReceitaDespesaAnulacaoPeriodo (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "RelacaoPrevisaoReceitaDespesaAnulacaoPeriodo.dsx":0000
End
Attribute VB_Name = "rptRelacaoPrevisaoReceitaDespesaAnulacaoPeriodo"
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
    If Val(fldbytSupRed) = 1 Then
        txtTipoCredito.Text = "Redução"
    Else
        txtTipoCredito.Text = "Suplementação"
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra1
    TrocaCorParaZebrado lblSombra2
    TrocaCorParaZebrado lblsombra3
    TrocaCorParaZebrado lblSombra4
    TrocaCorParaZebrado lblSombra5
    TrocaCorParaZebrado lblSombra6
    TrocaCorParaZebrado lblSombra7
    fldCodigo = gvntFormatacaoEspecifica(fldCodigo, 1)
    fldstrcampos = fldCodigo + " " + fldstrDescCod
            
            lbl_Funcao.Visible = Not CBool(Val(fldbytTipo))
            txtCodigoReduzido.Visible = Not CBool(Val(fldbytTipo))
            lblstrCodigo.Visible = Not CBool(Val(fldbytTipo))
            txtstrCodigo.Visible = Not CBool(Val(fldbytTipo))
            lbltxtOrgao.Visible = Not CBool(Val(fldbytTipo))
            txtOrgao.Visible = Not CBool(Val(fldbytTipo))
            lbltxtUnidadeOrcamentaria.Visible = Not CBool(Val(fldbytTipo))
            txtUnidadeOrcamentaria.Visible = Not CBool(Val(fldbytTipo))
            lbltxtSubunidade.Visible = Not CBool(Val(fldbytTipo))
            txtSubunidade.Visible = Not CBool(Val(fldbytTipo))
            lblTipoCredito.Visible = Not CBool(Val(fldbytTipo))
            txtstrTipoCredito.Visible = Not CBool(Val(fldbytTipo))
            lbltxtFuncaoGoverno.Visible = Not CBool(Val(fldbytTipo))
            txtFuncaoGoverno.Visible = Not CBool(Val(fldbytTipo))
            lbltxtSubfuncao.Visible = Not CBool(Val(fldbytTipo))
            txtSubfuncao.Visible = Not CBool(Val(fldbytTipo))
            lbltxtPrograma.Visible = Not CBool(Val(fldbytTipo))
            txtPrograma.Visible = Not CBool(Val(fldbytTipo))
            lbltxtSubprograma.Visible = Not CBool(Val(fldbytTipo))
            txtSubprograma.Visible = Not CBool(Val(fldbytTipo))
            lbltxtProjetoAtividade.Visible = Not CBool(Val(fldbytTipo))
            txtProjetoAtividade.Visible = Not CBool(Val(fldbytTipo))
            lbltxtElementoDespesa.Visible = Not CBool(Val(fldbytTipo))
            txtElementoDespesa.Visible = Not CBool(Val(fldbytTipo))
            lbldblValorDespesa.Visible = Not CBool(Val(fldbytTipo))
            flddblValorDespesa.Visible = Not CBool(Val(fldbytTipo))
            Line12.Visible = Not CBool(Val(fldbytTipo))
            lblSombra2.Visible = Not CBool(Val(fldbytTipo))
            lblsombra3.Visible = Not CBool(Val(fldbytTipo))
            lblSombra4.Visible = Not CBool(Val(fldbytTipo))
            lblSombra5.Visible = Not CBool(Val(fldbytTipo))
            lblSombra6.Visible = Not CBool(Val(fldbytTipo))
            lblSombra7.Visible = Not CBool(Val(fldbytTipo))
            lblCodigo.Visible = CBool(Val(fldbytTipo))
            fldstrcampos.Visible = CBool(Val(fldbytTipo))
            lblValorReceita.Visible = CBool(Val(fldbytTipo))
            flddblValorReceita.Visible = CBool(Val(fldbytTipo))
            Line15.Visible = CBool(Val(fldbytTipo))
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
