VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptLivroDiario 
   Caption         =   "prjOrcamentario - rptLivroDiario (ActiveReport)"
   ClientHeight    =   8010
   ClientLeft      =   90
   ClientTop       =   1860
   ClientWidth     =   11040
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   19473
   _ExtentY        =   14129
   SectionData     =   "LivroDiario.dsx":0000
End
Attribute VB_Name = "rptLivroDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varAux     As Double
Public blnSemContas As Boolean
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
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtstrEstado
    lbl_Titulo = Me.Caption
    lbl_Periodo = Me.Tag
    
    txt_TotalDebito = "0"
    txt_TotalCredito = "0"
    
    txt_TotalGeralDebito = "0"
    txt_TotalGeralCredito = "0"
    
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

Private Sub ActiveReport_ReportEnd()
    Dim i As Integer
    For i = 0 To Me.Pages.Count - 1
        Me.Pages(i).Orientation = ddOLandscape
    Next
End Sub

Private Sub Detail_AfterPrint()
 fldstrHistorico.Visible = False
End Sub

Private Sub Detail_BeforePrint()
    lblSombra.Height = Detail.Height
End Sub

Private Sub Detail_Format()
    txt_dtmDataDetalhe = gstrDataFormatada(txt_dtmDataDetalhe)
    TrocaCorParaZebrado lblSombra
    fldstrConta.Text = gvntFormatacaoEspecifica(fldstrConta.Text, 1)
    
    If Len(Trim(txt_dblCredito)) = 0 Then
       txt_dblCredito = "0"
    End If
    If Len(Trim(txt_dblDebito)) = 0 Then
       txt_dblDebito = "0"
    End If
    
    
    If CDbl(gstrConvVrDoSql(txt_dblDebito)) < 0 Then
           txt_dblCredito = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(txt_dblDebito)) * (-1))
           txt_dblDebito = "0"
    ElseIf CDbl(gstrConvVrDoSql(txt_dblCredito)) < 0 Then
           txt_dblDebito = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(txt_dblCredito)) * (-1))
           txt_dblCredito = "0"
    End If
    
    txt_TotalDebito = gstrConvVrDoSql(CDbl(txt_TotalDebito) + CDbl(txt_dblDebito))
    txt_TotalCredito = gstrConvVrDoSql(CDbl(txt_TotalCredito) + CDbl(txt_dblCredito))
    
    txt_TotalGeralDebito = gstrConvVrDoSql(CDbl(txt_TotalGeralDebito) + CDbl(txt_dblDebito))
    txt_TotalGeralCredito = gstrConvVrDoSql(CDbl(txt_TotalGeralCredito) + CDbl(txt_dblCredito))
    
    If txt_dblDebito = "0" Then txt_dblDebito = ""
    If txt_dblCredito = "0" Then txt_dblCredito = ""
    
End Sub

Private Sub GroupHeader1_Format()
    txt_DtmData = gstrDataFormatada(txt_DtmData)
    
    txt_TotalDebito = "0"
    txt_TotalCredito = "0"
    
End Sub

Private Sub grpH_Lancamento_Format()
    fldstrHistorico.Visible = True
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

