VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelEtiquetasProcessoMod2 
   Caption         =   "Tributario - rptRelEtiquetasProcessoMod2 (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptRelEtiquetasProcessoMod2.dsx":0000
End
Attribute VB_Name = "rptRelEtiquetasProcessoMod2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ActiveReport_Activate()

    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
    
    lblData.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblData1.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblData2.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblData3.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    
    lblProcesso.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblProcesso1.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblProcesso2.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblProcesso3.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    
    lblContribuinte.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblContribuinte3.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    
    lblEndereco3.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    
    lblGrupoAssunto.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblGrupoAssunto3.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblDescricao3.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo
    lblInfo.Visible = frmRelEtiquetadeProcesso.blnImprimeTitulo

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
PadronizaToolBarRelatorio Me
Me.PageSettings.TopMargin = 500
Me.PageSettings.LeftMargin = 200
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
    
    If adoDataControl.NRecords > 0 Then
        
        txtdtmDtData = gstrDataFormatada(txtdtmDtData)
        txtDtmDtData1 = gstrDataFormatada(txtDtmDtData1)
        txtDtmDtData2 = gstrDataFormatada(txtDtmDtData2)
        txtDtmDtData3 = gstrDataFormatada(txtDtmDtData3)
        txtInfo = gstrDataFormatada(txtDtmDtData2)
                
        If Trim(adoDataControl.Recordset("strEnderecoAcao")) = "," Then
            If Trim(adoDataControl.Recordset("strEndereco")) = "," Then
                txtstrEndereco3.Text = ""
            Else
                lblEndereco3.Caption = "Endereço do Requerente"
                txtstrEndereco3.Text = Space$(0) & adoDataControl.Recordset("strEndereco")
                txtstrBairro3.Text = Space$(0) & adoDataControl.Recordset("strBairro")
                txtintCep3.Text = Space$(0) & adoDataControl.Recordset("intCep")
            End If
        Else
            lblEndereco3.Caption = "Endereço de Ação"
            txtstrEndereco3.Text = Space$(0) & adoDataControl.Recordset("strEnderecoAcao")
            txtstrBairro3.Text = Space$(0) & adoDataControl.Recordset("strBairroAcao")
            txtintCep3.Text = Space$(0) & adoDataControl.Recordset("intCepAcao")
        End If
        
    End If
    
    txtintCep3 = gstrCEPFormatado(txtintCep3)
    
    If adoDataControl.Recordset.AbsolutePosition = adoDataControl.NRecords Then
        PageBreak1.Enabled = False
    End If
        
End Sub
