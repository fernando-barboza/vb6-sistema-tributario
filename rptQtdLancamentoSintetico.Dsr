VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptQtdLancamentoSintetico 
   Caption         =   "Relatório Sintético de Qtde. de Lançamentos, Valor e Tipo "
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "rptQtdLancamentoSintetico.dsx":0000
End
Attribute VB_Name = "rptQtdLancamentoSintetico"
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
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
    
    
    
    
    
    
    
    
'    Set gobjBanco = New clsBanco
'    txt_Registros.Text = ""
'    If gobjBanco.CriaADO(strQuerryAnalitico2, 5, adoResultado) Then
'        With adoResultado
'            Do While Not .EOF
'                txt_Registros.Text = txt_Registros.Text & !DocNome & " :"
'                txt_Registros.Text = txt_Registros.Text & "     " & !TotalDocs
'                txt_Registros.Text = txt_Registros.Text & " " & Chr(13)
'                .MoveNext
'            Loop
'        End With
'    End If
    
End Sub

Private Function strQuerryAnalitico2() As String
' Traz as Ocorrencias e a quantidade de cada
' NomeOcorrencia = qtde
Dim strSql As String
Dim dtInicial  As Date
Dim dtFinal    As Date
dtInicial = CVDate(frmCadqtdeLancamentoValorTipo.txt_Inicial)
dtFinal = CVDate(frmCadqtdeLancamentoValorTipo.txt_Ate)

    strSql = ""
    strSql = strSql & " SELECT COUNT(*) as TotalDocs , DV.intDocumentosEmitidos Inteiro, DE.strDescricao DocNome "
    strSql = strSql & " FROM " & gstrDevolucao & " DV, "
    strSql = strSql & gstrDocumentoEmitido & " DE "
    strSql = strSql & " WHERE DV.intDocumentosEmitidos = DE.PKId "
    strSql = strSql & " AND DV.dtmDevolucao BETWEEN " & gstrConvDtParaSql(dtInicial) & " AND " & gstrConvDtParaSql(dtFinal)
    strSql = strSql & " GROUP BY DV.intDocumentosEmitidos, DE.strDescricao "

strQuerryAnalitico2 = strSql
End Function

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
    'vet(codrecetita,)=
End Sub

Private Sub GroupHeader1_Format()
    Beep
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
'    lbl_Total = gstrTotalDeRegistros(gstrSocio, lblRelatorio)
    MostraEmissorRelatorio Me
End Sub



