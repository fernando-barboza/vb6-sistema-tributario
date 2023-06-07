VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptLivroDividaAtiva 
   Caption         =   "rptLivroDividaAtiva (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   30
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptLivroDividaAtiva.dsx":0000
End
Attribute VB_Name = "rptLivroDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim varAux As Double
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
    LeImagemLogotipo imgBrasao, "", txtstrNomeFantasia, txtstrEstado
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
    
    txtdtmInscricao = gstrDataFormatada(txtdtmInscricao)
    txtdtmVencimento = gstrDataFormatada(txtdtmVencimento)
    
    'Alterado em pendencia tri0809
    txtstrInscricao = gstrFormataInscricao(Right(txtstrInscricao, Val(gintRetornaTamanhoMascara(adoDataControl.Recordset.Fields("intUtilizacao")))), adoDataControl.Recordset.Fields("intUtilizacao"))
    
    txtdblCorrecao = gstrConvVrDoSql(txtdblCorrecao, 2, , True)
    txtdblMulta = gstrConvVrDoSql(txtdblMulta, 2, , True)
    txtdblJuros = gstrConvVrDoSql(txtdblJuros, 2, , True)
    Field20 = gstrConvVrDoSql(Field20, 2, , True)
    Field21 = gstrConvVrDoSql(Field21, 2, , True)
    Field22 = gstrConvVrDoSql(Field22, 2, , True)
    txtdblTotal = gstrConvVrDoSql(CDbl(txtdblCorrecao) + CDbl(txtdblMulta) + CDbl(txtdblJuros) + CDbl(Field22), 2)
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
