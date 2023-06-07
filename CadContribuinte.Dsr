VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCadContribuinte 
   Caption         =   "prjOrcamentario - rptCadContribuinte (ActiveReport)"
   ClientHeight    =   5580
   ClientLeft      =   1155
   ClientTop       =   1395
   ClientWidth     =   11910
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   21008
   _ExtentY        =   9843
   SectionData     =   "CadContribuinte.dsx":0000
End
Attribute VB_Name = "rptCadContribuinte"
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

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogoTipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
    
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
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
    TrocaCorDaSecaoParaZebrado Detail
     
    If Not adoDataControl.Recordset.EOF And Not adoDataControl.Recordset.BOF Then
        If adoDataControl.Recordset("bytNaturezaJuridica").Value = 1 Then
             txtstrCNPJ = gstrCGCCPFFormatado(txtstrCNPJ, "PJ")
        Else
             txtstrCNPJ = gstrCGCCPFFormatado(txtstrCNPJ, "PF")
        End If
    End If

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



