VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptBaixas 
   Caption         =   "rptBaixas (ActiveReport)"
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
   SectionData     =   "rptBaixas.dsx":0000
End
Attribute VB_Name = "rptBaixas"
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
    Me.Caption = "Baixas"
    lblRelatorio = "Baixas do dia " & gstrDataFormatada(gstrENulo(adoDataControl.Recordset!Dtmdtmovimento))
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

Private Sub PageFooter_Format()
    lblPagina.Caption = "P�gina " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra
    
    txtDblprincipal.Text = gstrConvVrDoSql(txtDblprincipal.Text)
    txtDblmulta.Text = gstrConvVrDoSql(txtDblmulta.Text)
    txtDbljuros.Text = gstrConvVrDoSql(txtDbljuros.Text)
    txtDblcorrecao.Text = gstrConvVrDoSql(txtDblcorrecao.Text)
    txtDblTotal.Text = gstrConvVrDoSql(txtDblTotal.Text)
End Sub

Private Sub ReportFooter_Format()
    
    txtDblprincipalB.Text = gstrConvVrDoSql(txtDblprincipalB.Text)
    txtDblmultaB.Text = gstrConvVrDoSql(txtDblmultaB.Text)
    txtDbljurosB.Text = gstrConvVrDoSql(txtDbljurosB.Text)
    txtDblcorrecaoB.Text = gstrConvVrDoSql(txtDblcorrecaoB.Text)
    txtDblTotalB.Text = gstrConvVrDoSql(txtDblTotalB.Text)
    MostraEmissorRelatorio Me

    
End Sub