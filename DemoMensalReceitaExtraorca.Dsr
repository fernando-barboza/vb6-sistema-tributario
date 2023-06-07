VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptDemoMensalReceitaExtra 
   Caption         =   "rptDemoMensalReceitaExtra (ActiveReport)"
   ClientHeight    =   10755
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   11970
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   21114
   _ExtentY        =   18971
   SectionData     =   "DemoMensalReceitaExtraorca.dsx":0000
End
Attribute VB_Name = "rptDemoMensalReceitaExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intQuantidade As Integer

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
    
    Me.Caption = "Demonstrativo Mensal da Receita Extra-Orçamentária"
    lbl_Titulo = Me.Caption
    
    With frmRelatorioPeriodo
        lblPerioro = "Perído: " & Trim(.txtdtmInicial) & " a " & Trim(.txtdtmFinal)
    End With
    If Not gblnRestartRelatorio Then
        intQuantidade = adoDataControl.NRecords
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

Private Sub grhRelacao_Format()
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub GroupHeader1_Format()
    If intQuantidade > 0 Then
        txtstrPeriodo = gstrNomeDoMes(adoDataControl.Recordset("MES"))
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra
    txtCodigoOrcamentario = gvntFormatacaoEspecifica(txtCodigoOrcamentario, 1)
    'flddblTotalAtePeriodo = gstrConvVrDoSql(Val(gstrConvVrParaSql(flddblTotalAtePeriodo)) + _
                    Val(gstrConvVrParaSql(flddblPeriodo)))
    
    flddblTotalAtePeriodo = gstrConvVrDoSql(Val(gstrConvVrParaSql(flddblTotalAtePeriodo)) + _
                    Val(gstrConvVrParaSql(flddblAtePeriodo)))
                    
    flddblTotalPeriodo = gstrConvVrDoSql(Val(gstrConvVrParaSql(flddblTotalPeriodo)) + _
                    Val(gstrConvVrParaSql(flddblPeriodo)))
End Sub

Private Sub ReportFooter_Format()
    TrocaCorParaZebrado lblSombra1
    MostraEmissorRelatorio Me
End Sub
