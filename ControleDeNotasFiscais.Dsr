VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptControleDeNotasFiscais 
   Caption         =   "Controle de Notas Fiscais"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   16298
   _ExtentY        =   13679
   SectionData     =   "ControleDeNotasFiscais.dsx":0000
End
Attribute VB_Name = "rptControleDeNotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoResultado As ADODB.Recordset

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
    'TrocaCorParaZebrado lblSombra
    If txtdtmDtExercicio.Text <> "" Then
        txtdtmDtExercicio.Text = Year(txtdtmDtExercicio.Text)
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

'Private Sub MostraValor()
'Dim strSql As String
'    strSql = ""
'    strSql = strSql & " SELECT "
'    strSql = strSql & " SUM(dblValorParcela) AS Valor "
'    strSql = strSql & " FROM "
'    strSql = strSql & gstrParcelaReceita
'    strSql = strSql & " WHERE "
'    strSql = strSql & " intComposicaoDaReceita = " & Val(txtintComposicaoDaReceita.Text)
'    strSql = strSql & " AND intLancamentoCalculo = " & Val(txtPKIdLancamentoCalculo.Text)
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        With adoResultado
'            Do While Not .EOF
'                txt_Valor.Text = gstrConvVrDoSql(!Valor)
'                .MoveNext
'            Loop
'        End With
'    End If
'End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
