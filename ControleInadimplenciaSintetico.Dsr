VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptControleInadimplenciaSintetico 
   Caption         =   "Relatório para Controle de Inadimplência Sintético"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "ControleInadimplenciaSintetico.dsx":0000
End
Attribute VB_Name = "rptControleInadimplenciaSintetico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoResultado    As ADODB.Recordset
Dim dblTotalGeral   As Double
Dim intReceitas     As Integer

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
    intReceitas = 0
    dblTotalGeral = 0
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
    TrocaCorParaZebrado lblSombra
    txtdblValor = gstrConvVrDoSql(txtdblValor, 2)
End Sub
 

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub MostraValorTotal()
Dim strSql As String
'    strSql = ""
'    strSql = strSql & " SELECT "
'    strSql = strSql & " SUM(dblValorParcela) AS ValorTotal "
'    strSql = strSql & " FROM "
'    strSql = strSql & gstrParcelaReceita
'    strSql = strSql & " WHERE "
'    strSql = strSql & " intComposicaoDaReceita = " & Val(txtintComposicaoDaReceita.Text)
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        If Not adoResultado.EOF Then
'            With adoResultado
'                txt_Valor.Text = gstrConvVrDoSql(gstrVerificaCampoNulo(!ValorTotal))
'                'dblTotalGeral = dblTotalGeral + CDbl(txt_Valor.Text)
'
'            End With
'        End If
'    End If
End Sub

Private Sub ReportFooter_Format()
    intReceitas = 0
    dblTotalGeral = 0
    MostraEmissorRelatorio Me
End Sub


