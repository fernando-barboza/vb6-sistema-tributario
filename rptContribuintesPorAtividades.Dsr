VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptContribuintesPorAtividades 
   Caption         =   "Tributario - rptContribuintesPorAtividades (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptContribuintesPorAtividades.dsx":0000
End
Attribute VB_Name = "rptContribuintesPorAtividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim intAtivo As Integer
    Dim intInativo As Integer

Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
    
    If adoDataControl.Recordset.EOF Then
       ExibeMensagem "Não existem nenhum contribuinte na(s) atividade(s) selecionadas."
       Unload Me
    End If
    
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_ReportStart()
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption & " - " & adoDataControl.Recordset!strOcorrencia
    lblContagemOcorrencia.Caption = adoDataControl.Recordset!strOcorrencia
    
    If frmContribuintesPorAtividades.chkReduzido.Value = 1 Then
        grpstrAtividadeRelRed.Visible = True
        Detail.Height = 255
        lblSombra.Height = 255
        Label6.Visible = False
        txtInscricao.Visible = False
        Label5.Visible = False
        txtRazaoSocial.Visible = False
        txtPS.Visible = False
        txtblnInativo.Visible = False
        Label7.Visible = False
        txtEndereco.Visible = False
        Label8.Visible = False
        txtNumero.Visible = False
        Label10.Visible = False
        txtComplemento.Visible = False
        Label9.Visible = False
        txtBairro.Visible = False
        Label16.Visible = True
        txtInscricaoR.Visible = True
        Label18.Visible = True
        txtPSR.Visible = True
        Label17.Visible = True
        txtRazaoSocialR.Visible = True
        Label19.Visible = True
        txtblnInativoR.Visible = True
        Label20.Visible = True
        txtEnderecoR.Visible = True
        Label21.Visible = True
        txtCep.Visible = True
    End If


    
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
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
    
    TrocaCorParaZebrado lblSombra
    txtInscricao.Text = gstrFormataInscricao(txtInscricao.Text, 2)
    txtInscricaoR.Text = gstrFormataInscricao(txtInscricaoR.Text, 2)
    txtCep.Text = gstrCEPFormatado(txtCep.Text)
    

    With adoDataControl.Recordset
        If Not IsNull(!strLogradouro) Then
            txtEnderecoR.Text = Trim(!strLogradouro) & ", "
        End If
        If Not IsNull(!INTNUMERO) Then
            txtEnderecoR.Text = txtEnderecoR.Text & Trim(!INTNUMERO) & ", "
        End If
        If Not IsNull(!STRCOMPLEMENTO) Then
            txtEnderecoR.Text = txtEnderecoR.Text & Trim(!STRCOMPLEMENTO) & ", "
        End If
        If Not IsNull(!strBairro) Then
            txtEnderecoR.Text = txtEnderecoR.Text & Trim(!strBairro)
        End If
    End With
    
    'If txtblnInativo.Text = "Ativo" Then
    '    intAtivo = intAtivo + 1
    'ElseIf txtblnInativo.Text = "Inativo" Then
        intInativo = intInativo + 1
    'End If
    
    'If adoDataControl.Recordset!strDescricao = "Ativo" Then
    '    txtblnInativoR = "A"
    'Else
    '    txtblnInativoR = "I"
    'End If
    
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
    txtTotalInativo = intInativo
    txtTotalAtivo = intAtivo
    txtTotalS = intAtivo + intInativo
End Sub


