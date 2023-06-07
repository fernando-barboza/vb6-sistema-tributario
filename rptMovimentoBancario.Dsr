VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptMovimentoBancario 
   Caption         =   "Tributario - rptMovimentoBancario (ActiveReport)"
   ClientHeight    =   11235
   ClientLeft      =   -45
   ClientTop       =   345
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19817
   SectionData     =   "rptMovimentoBancario.dsx":0000
End
Attribute VB_Name = "rptMovimentoBancario"
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
    Me.Caption = "Movimento Bancário"
    lblRelatorio = "Movimento bancário das baixas do dia " & gstrDataFormatada(gstrENulo(adoDataControl.Recordset!Dtmdtmovimento))
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

Private Sub GHLote_Format()
'    If adoDataControl.Recordset.RecordCount > 0 Then
'        If adoDataControl.Recordset("intTipoCritica").Value <> 6 Then
'            Label1.Caption = "Inscrição Cadastral"
'            Label2.Caption = "Composição da Receita"
'        Else
'            Label1.Caption = "Código de Barras"
'            Label2.Caption = ""
'        End If
'    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
    If adoDataControl.NRecords > 0 Then
       
       TrocaCorParaZebrado lblSombra
       txtDtmdtpagamento.Text = gstrDataFormatada(txtDtmdtpagamento)
       txtDtmdtmovimento.Text = gstrDataFormatada(txtDtmdtmovimento)
       
       txtdblValor.Text = gstrConvVrDoSql(txtdblValor, 2)
       
'       txtStrinscricao.Visible = Not adoDataControl.Recordset("intTipoCritica").Value = 6
'       txtStrcomposicaodareceita.Visible = Not adoDataControl.Recordset("intTipoCritica").Value = 6
'       txtstrCodigoDeBarras.Visible = adoDataControl.Recordset("intTipoCritica").Value = 6

'       If adoDataControl.Recordset("intTipoCritica").Value <> 6 And Len(Trim(txtStrinscricao.Text)) > 0 Then
       If Not txtStrinscricao.Text = "" Then
           txtStrinscricao.Text = gstrFormataInscricao(Right(txtStrinscricao.Text, gintRetornaTamanhoMascara(adoDataControl.Recordset("intUtilizacao"))), adoDataControl.Recordset("intUtilizacao"))
       End If
'       End If
       
    End If
    
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

