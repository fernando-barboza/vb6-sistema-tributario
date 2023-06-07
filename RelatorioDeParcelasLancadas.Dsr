VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioDeParcelasLancadas 
   Caption         =   "Relação das Parcelas Lançadas"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "RelatorioDeParcelasLancadas.dsx":0000
End
Attribute VB_Name = "rptRelatorioDeParcelasLancadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnMostrar As Boolean
Dim dblTotalParcela As Double
Dim dblTotalDesconto As Double

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
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
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
Dim strSql As String
Dim adoRec As ADODB.Recordset

txtdblValorParcela.Text = gstrConvVrDoSql(txtdblValorParcela.Text, 2)

If Trim(txtdblValorParcela.Text) <> "" Then
    dblTotalParcela = dblTotalParcela + txtdblValorParcela.Text
    txtTotalParcela.Text = gstrConvVrDoSql(dblTotalParcela, 2)
Else
    txtTotalParcela.Text = ""
End If


If blnMostrar Then
   txtParcela.Text = ""
   txtdtmDataVencimento.Text = ""
   txtdblDesconto.Text = ""

Else
   blnMostrar = True
   If txtParcela = "0" Then
        strSql = ""
        strSql = strSql & " SELECT dblValorDesconto FROM " & gstrParcelaReceita
        strSql = strSql & " WHERE intLancamentoCalculo = " & txtLancamentoCalculo
        strSql = strSql & " AND intNumeroParcela = 0"
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 10, adoRec) Then
            If Not (adoRec.BOF And adoRec.EOF) Then
                txtdblDesconto.Text = gstrConvVrDoSql(adoRec!dblValorDesconto, 2)
            Else
                txtdblDesconto.Text = ""
            End If
        Else
            txtdblDesconto.Text = ""
        End If
        If Trim(txtdblDesconto.Text) <> "" Then
            dblTotalDesconto = dblTotalDesconto + txtdblDesconto.Text
            txtTotalDesconto.Text = gstrConvVrDoSql(dblTotalDesconto)
        End If
   End If
End If
End Sub

Private Sub grpParcela_Format()
TrocaCorParaZebrado lblSombra
blnMostrar = False
dblTotalDesconto = 0
dblTotalParcela = 0
txtTotalDesconto.Text = ""
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
