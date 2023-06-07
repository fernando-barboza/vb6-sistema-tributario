VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptListagemCredores 
   Caption         =   "rptListagemCredores (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptListagemCredores.dsx":0000
End
Attribute VB_Name = "rptListagemCredores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intPagina       As Integer

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
    
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtEstado
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

Private Sub GrhCredor_Format()
    
    If Not adoDataControl.Recordset.EOF And Not adoDataControl.Recordset.BOF Then
        If adoDataControl.Recordset("bytNaturezaJuridica").Value = "Jurídica" Then
             lblCNPJ = "CNPJ:"
             txtstrCNPJCPF = gstrCGCCPFFormatado(txtstrCNPJCPF, "PJ")
             lblblnResidenteNoMunicipio = "Estabelecido no Município:"
             AjustaPessoaJuridica True
             lblblnResidenteNoMunicipio.Left = lbldtmDataNascimento.Left
             
             lbldtmDataCadastro.Left = lbldtmDataNascimento.Left
             chkblnResidenteNoMunicipio.Left = lblblnResidenteNoMunicipio.Left + lblblnResidenteNoMunicipio.Width + 80
             dtpdtmDataCadastro.Left = lbldtmDataCadastro.Left + lbldtmDataCadastro.Width + 80
        Else
            lblCNPJ = "CPF:"
            txtstrCNPJCPF = gstrCGCCPFFormatado(txtstrCNPJCPF, "PF")
            lblblnResidenteNoMunicipio = "Residente no Município:"
            AjustaPessoaJuridica False
        End If
        
        dbcstrLogradouroC = IIf(adoDataControl.Recordset("intTipoLogradouro").Value <> "", adoDataControl.Recordset("intTipoLogradouro").Value & " ", "") & IIf(adoDataControl.Recordset("intTituloLogradouro").Value <> "", adoDataControl.Recordset("intTituloLogradouro").Value & " ", "") & dbcstrLogradouroC
        strLogradouroD = IIf(adoDataControl.Recordset("intTipoLogradouroD").Value <> "", adoDataControl.Recordset("intTipoLogradouroD").Value & " ", "") & IIf(adoDataControl.Recordset("intTituloLogradouroD").Value <> "", adoDataControl.Recordset("intTituloLogradouroD").Value & " ", "") & strLogradouroD
        
        dtpdtmDataCadastro = gstrDataFormatada(dtpdtmDataCadastro)
        
        INTCEP = gstrCEPFormatado(INTCEP)
        txtintCEPC = gstrCEPFormatado(txtintCEPC)
        txtintCepD = gstrCEPFormatado(txtintCepD)
    End If
    
End Sub

Private Sub AjustaPessoaJuridica(ByVal blnPessoaJuridica As Boolean)
    lblstrNomeFantasia.Visible = blnPessoaJuridica
    dbcstrNomeFantasia.Visible = blnPessoaJuridica
    'lblblnResidenteNoMunicipio.Visible = blnPessoaJuridica
    'chkblnResidenteNoMunicipio.Visible = blnPessoaJuridica
    lblstrInscricaoEstadual.Visible = blnPessoaJuridica
    txtstrInscricaoEstadual.Visible = blnPessoaJuridica
    'lbldtmDataCadastro.Visible = blnPessoaJuridica
    'dtpdtmDataCadastro.Visible = blnPessoaJuridica
    'lblblnInativo.Visible = blnPessoaJuridica
    'chkblnInativo.Visible = blnPessoaJuridica
    
    lblstrIdentidade.Visible = Not blnPessoaJuridica
    txtstrIdentidade.Visible = Not blnPessoaJuridica
    lblstrTituloEleitoral.Visible = Not blnPessoaJuridica
    txtstrTituloEleitoral.Visible = Not blnPessoaJuridica
    lbldtmDataNascimento.Visible = Not blnPessoaJuridica
    lblstrCarteiraTrabalho.Visible = Not blnPessoaJuridica
    txtdtmDataNascimento.Visible = Not blnPessoaJuridica
    txtstrCarteiraTrabalho.Visible = Not blnPessoaJuridica
End Sub

Private Sub PageFooter_Format()

   lblPagina.Caption = pageNumber
   
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
    Me.lblRelatorio = Me.Caption
End Sub


Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

