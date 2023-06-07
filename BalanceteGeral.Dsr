VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptBalanceteGeral 
   Caption         =   "prjOrcamentario - rptBalanceteGeral (ActiveReport)"
   ClientHeight    =   10695
   ClientLeft      =   0
   ClientTop       =   315
   ClientWidth     =   15360
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   18865
   SectionData     =   "BalanceteGeral.dsx":0000
End
Attribute VB_Name = "rptBalanceteGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportEnd()
    Dim intInd As Integer
    For intInd = 0 To Me.Pages.Count - 1
        Me.Pages(intInd).Orientation = ddOLandscape
    Next
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
    With frmRelatorioPeriodo
        lblPeriodo = "Período: " & .txtdtmInicial & " a " & .txtdtmFinal
    End With
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub


Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub
Private Sub ActiveReport_ReportStart()
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtstrEstado
    lbl_Titulo = Me.Caption
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
    TrocaCorParaZebrado lblSombra
    txtstrDescricao.Left = (Val(txtbytNivel) - 1) * 100
    
    If txtdblSaldoInicial > 0 Then
        fldSIDC = "C"
    ElseIf txtdblSaldoInicial < 0 Then
        fldSIDC = "D"
    Else
        fldSIDC = IIf(adoDataControl.Recordset("blnNaturezaDaConta") = 0, "C", "D")
    End If
    
    If txtdblSaldoFinal > 0 Then
        fldSFDC = "C"
    ElseIf txtdblSaldoFinal < 0 Then
        fldSFDC = "D"
    Else
        fldSFDC = IIf(adoDataControl.Recordset("blnNaturezaDaConta") = 0, "C", "D")
    End If
    
    txtdblSaldoInicial = gstrConvVrDoSql(Abs(txtdblSaldoInicial))
    txtdblSaldoFinal = gstrConvVrDoSql(Abs(txtdblSaldoFinal))
End Sub

Private Sub grhOrgao_Format()
    TrocaCorParaZebrado lblSombraOrgao
End Sub

Private Sub grhOrcamentaraExtra_orcamentaria_Format()
    TrocaCorParaZebrado lblSombraOrcamentaria
End Sub

Private Sub grhReceitaDespesa_Format()
   TrocaCorParaZebrado lblSombraOrgao
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    Dim strSql  As String
    Dim adoResultado    As ADODB.Recordset
    strSql = strSql & "SELECT * FROM " & gstrConfiguracaoGeral
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txtstrNomePrefeito = Trim(!strNomePrefeito)
                txtstrNomeTesoureiro = Trim(!strNomeTesoureiro)
                txtstrNomeContador = Trim(!strNomeContador)
            End If
        End With
    End If
    MostraEmissorRelatorio Me
End Sub
