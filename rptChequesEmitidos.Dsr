VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptChequesEmitidos 
   Caption         =   "prjOrcamentario - rptChequesEmitidos (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptChequesEmitidos.dsx":0000
End
Attribute VB_Name = "rptChequesEmitidos"
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
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
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

Private Sub GroupFooter1_Format()
    'txt_dblTotal = gstrConvVrDoSql(txt_dblTotal)
 '      If txt_Count.Text = adoDataControl.NRecords Then
'      txtParcialGFooter.Text = txtValorTotPg.Text
 '     txtTotalGFooter.Text = txt_dblTotal.Text
      Field4.Visible = False
   '   txt_dblTotal.Visible = False
      lblTitParcPF.Visible = False
     ' lblTotPG.Visible = False
  ' End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
   ' txtValorTotPg.Text = Format(txtValorSoma.Text, "#,##0.00")
    txtValorSoma.Text = "0"
    
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
   TrocaCorParaZebrado lblSombra
   txtdblValor = gstrConvVrDoSql(txtdblValor)
   txt_Count.Text = Val(txt_Count.Text) + 1
   If txtValorSoma.Text = "" Then
      txtValorSoma.Text = txtdblValor.Text
   Else
      txtValorSoma.Text = CCur(txtValorSoma.Text) + CCur(txtdblValor.Text)
   End If
   If Val(txtintOrdemPgto) = -1 Then
      txtintOrdemPgto = ""
   End If

    If Not txtintOrdemPgto.Text = "" Then: txtintOrdemPgto = gstrRetornaOps(txtintOrdemPgto)
 '   txtValorTotPg.Text = txtValorSoma.Text
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

Private Sub ReportHeader_Format()
      Field4.Visible = True
      lblTitParcPF.Visible = True
End Sub
