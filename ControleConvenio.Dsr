VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptControleConvenio 
   Caption         =   "prjOrcamentario - rptControleConvenio (ActiveReport)"
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
   SectionData     =   "ControleConvenio.dsx":0000
End
Attribute VB_Name = "rptControleConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_ReportEnd()
    Dim i As Integer
    For i = 0 To Me.Pages.Count - 1
        Me.Pages(i).Orientation = ddOLandscape
    Next
End Sub

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
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
    fldstrDescricao.Left = fldstrDescricao.Left + Abs(Val(fldbytNivel)) * 300
    If Val(fldbytNivel.Text) = 0 Then
        txtArrecadacaoSaldo = gstrConvVrDoSql(Val(gstrConvVrParaSql(flddblValor)) - Val(gstrConvVrParaSql(flddblValorArrecadacao)))
        txtEmpenhoSaldo = gstrConvVrDoSql(Val(gstrConvVrParaSql(flddblValor)) - (Val(gstrConvVrParaSql(flddblEmpenhoEmpenhado)) - Val(gstrConvVrParaSql(flddblEmpenhoAnulado))))
    Else
        txtArrecadacaoSaldo = ""
        txtEmpenhoSaldo = ""
    End If
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora.Caption = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
