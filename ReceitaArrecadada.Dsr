VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptReceitaArrecadada 
   Caption         =   "Tributario - rptReceitaArrecadada (ActiveReport)"
   ClientHeight    =   8490
   ClientLeft      =   285
   ClientTop       =   1140
   ClientWidth     =   12180
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   21484
   _ExtentY        =   14975
   SectionData     =   "ReceitaArrecadada.dsx":0000
End
Attribute VB_Name = "rptReceitaArrecadada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mdblTotalGeral  As Double
    Dim dblValorSubTotalConta As Double

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

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
    Me.Caption = "Receita Arrecadada"
    lbl_Titulo = Me.Caption
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtstrEstado
    With frmRelatorioPeriodo
        lblPerioro = "Perído: " & Trim(.txtdtmInicial) & " a " & Trim(.txtdtmFinal)
        If .chk_detalhado.Value <> 1 Then
            lblstrContaBancaria.Visible = False
            txtstrContaBancaria.Visible = False
            lblstrBanco.Visible = False
            txtstrBanco.Visible = False
            lblstrAgencia.Visible = False
            txtstrAgencia.Visible = False
            GroupHeader1.DataField = "bytTipo"
        Else
            txt_strExtraOrcamentaria.Visible = False
        End If
    End With
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
    txtstrCodigoOrcamentario = gvntFormatacaoEspecifica(txtstrCodigoOrcamentario)
    'TrocaCorParaZebrado lblSombra
    TrocaCorDaSecaoParaZebrado Detail
    dblValorSubTotalConta = dblValorSubTotalConta + Val(gstrConvVrParaSql(txtdblValorOrcamentario.Text))
End Sub

Private Sub GroupFooter1_Format()
    TrocaCorDaSecaoParaZebrado GroupFooter1
    txtDblValorSubTotal = gstrConvVrDoSql(dblValorSubTotalConta)
    dblValorSubTotalConta = 0
End Sub

Private Sub GroupHeader1_Format()
    If frmRelatorioPeriodo.chk_detalhado <> 1 And adoDataControl.Recordset.RecordCount > 0 Then
        If adoDataControl.Recordset("bytTipo").Value = 1 Then
            txt_strExtraOrcamentaria = "Extra-Orçametaria"
        Else
            txt_strExtraOrcamentaria = "Orçametaria"
        End If
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
    
End Sub

Private Sub ReportFooter_Format()
    txtTotalGeral = gstrConvVrDoSql(mdblTotalGeral)
    MostraEmissorRelatorio Me
End Sub

Private Sub ReportHeader_Format()
    dblValorSubTotalConta = 0
End Sub
