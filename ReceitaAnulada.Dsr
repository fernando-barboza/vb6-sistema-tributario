VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptReceitaAnulada 
   Caption         =   "prjOrcamentario - rptReceitaAnulada (ActiveReport)"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "ReceitaAnulada.dsx":0000
End
Attribute VB_Name = "rptReceitaAnulada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TotalGeral As Currency

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
    lbl_Titulo = Me.Caption
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtstrEstado
    On Error GoTo saida
    With frmRelatorioPeriodo
        lblPerioro = "Perído: " & Trim(.txtdtmInicial) & " a " & Trim(.txtdtmFinal)
    End With
saida:
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
    If Val(txtbytType) = 0 Then
        txtdtmdataCancelamento.Visible = True
        lbl_dtmdata.Visible = True
        linha.Visible = True
        TotalGeral = TotalGeral + Val(txtdblValorOrcamentario)
    Else
        txtdtmdataCancelamento.Visible = False
        lbl_dtmdata.Visible = False
        linha.Visible = False
    End If
    If Val(txtbytTabela) = 0 Then
        txtstrCodigoOrcamentario = gvntFormatacaoEspecifica(txtstrCodigoOrcamentario, 2)
    Else
        txtstrCodigoOrcamentario = gvntFormatacaoEspecifica(txtstrCodigoOrcamentario, 1)
    End If
    TrocaCorParaZebrado lblSombra
    
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    txtTotalGeral = gstrConvVrDoSql(TotalGeral)
    MostraEmissorRelatorio Me
End Sub
