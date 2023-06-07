VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptInconsistenciaImobiliario 
   Caption         =   "Relatório de Inconsistências Imobiliárias"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   15743
   _ExtentY        =   12409
   SectionData     =   "rptInconsistenciaImobiliario.dsx":0000
End
Attribute VB_Name = "rptInconsistenciaImobiliario"
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

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Function strQuerryAnalitico2() As String
Dim strSql As String
Dim dtInicial  As Date
Dim dtFinal    As Date
Dim codInicial As Integer
Dim codFinal   As Integer
dtInicial = CVDate(frmCadRelacaoDeDocumentosDevolvidos.txt_Devolucao)
dtFinal = CVDate(frmCadRelacaoDeDocumentosDevolvidos.txt_Ate)
codInicial = Val(frmCadRelacaoDeDocumentosDevolvidos.txt_Inicial)
codFinal = Val(frmCadRelacaoDeDocumentosDevolvidos.txt_Final)

    strSql = ""
    strSql = strSql & " SELECT COUNT(*) as TotalDocs , DV.intDocumentosEmitidos Inteiro, DE.strDescricao DocNome "
    strSql = strSql & " FROM " & gstrDevolucao & " DV, "
    strSql = strSql & gstrDocumentoEmitido & " DE "
    strSql = strSql & " WHERE DV.intDocumentosEmitidos = DE.PKId "
    strSql = strSql & " AND DV.dtmDevolucao BETWEEN " & gstrConvDtParaSql(dtInicial) & " AND " & gstrConvDtParaSql(dtFinal)
    strSql = strSql & " AND DV.intContribuinte BETWEEN " & codInicial & " AND " & codFinal
    strSql = strSql & " GROUP BY DV.intDocumentosEmitidos, DE.strDescricao "

strQuerryAnalitico2 = strSql
End Function

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
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

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
