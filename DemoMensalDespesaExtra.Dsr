VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptDemoMensalDespesaExtra 
   Caption         =   "prjOrcamentario - rptDemoMensalDespesaExtra (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "DemoMensalDespesaExtra.dsx":0000
End
Attribute VB_Name = "rptDemoMensalDespesaExtra"
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
    With frmRelatorioPeriodo
        lbl_Periodo = "Per�odo: " & .txtdtmInicial & " � " & .txtdtmFinal
        Me.Caption = .Caption
    End With
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtstrEstado
    lbl_Titulo = Me.Caption
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

Private Sub grhRelacao_Format()
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub GroupHeader1_Format()
    If Not adoDataControl.Recordset.EOF Then
        txtstrPeriodo = gstrNomeDoMes(Left(adoDataControl.Recordset("strPeriodo"), 2))
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
    txtCodigoOrcamentario = gvntFormatacaoEspecifica(txtCodigoOrcamentario, 1)
    flddblTotalPeriodo = gstrConvVrDoSql(Val(gstrConvVrParaSql(flddblTotalPeriodo)) + _
                    Val(gstrConvVrParaSql(flddblPeriodo)))
End Sub

Private Sub ReportFooter_Format()
    TrocaCorParaZebrado lblSombra1
    MostraEmissorRelatorio Me
End Sub
