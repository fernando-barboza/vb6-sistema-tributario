VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSaldoBancario 
   Caption         =   "prjOrcamentario - rptSaldoBancario (ActiveReport)"
   ClientHeight    =   9675
   ClientLeft      =   255
   ClientTop       =   1455
   ClientWidth     =   14100
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   24871
   _ExtentY        =   17066
   SectionData     =   "SaldoBancario.dsx":0000
End
Attribute VB_Name = "rptSaldoBancario"
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
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtstrEstado
    lblRelatorio = Me.Caption
'    With frmRelatorioPeriodo
'        lblPeriodo = "Período: " & Trim(.txtdtmInicial) & " a " & Trim(.txtdtmFinal)
'    End With
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

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra
    fldstrDescricao.Left = (Val(txtbytNivel) - 1) * 100
    
    If Not adoDataControl.Recordset.EOF And Not adoDataControl.Recordset.BOF Then
        If adoDataControl.Recordset("intContaReduzida").Value <> "" Then
            fldstrDescricao.Text = adoDataControl.Recordset("intContaReduzida").Value & " - " & fldstrDescricao.Text
        End If
    End If
    
    
    flddblSaldoAtual = IIf(flddblSaldoAtual = "", 0, flddblSaldoAtual)
    flddblSaldoAnterior = IIf(flddblSaldoAnterior = "", 0, flddblSaldoAnterior)
    
    If flddblSaldoAtual > 0 Then
        fldstrDC = "C"
    Else
        fldstrDC = "D"
    End If

    If flddblSaldoAnterior > 0 Then
        fldstrDCSI = "C"
    Else
        fldstrDCSI = "D"
    End If
    flddblSaldoAtual = gstrConvVrDoSql(Abs(flddblSaldoAtual))
    flddblSaldoAnterior = gstrConvVrDoSql(Abs(flddblSaldoAnterior))
End Sub

Private Sub ActiveReport_ReportEnd()
    Dim i As Integer
    For i = 0 To Me.Pages.Count - 1
        Me.Pages(i).Orientation = ddOLandscape
    Next
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

