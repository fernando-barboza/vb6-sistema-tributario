VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCadDespesaExtraOrcamentaria 
   Caption         =   "prjOrcamentario - rptCadDespesaExtraOrcamentaria (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "CadDespesaExtraOrcamentaria.dsx":0000
End
Attribute VB_Name = "rptCadDespesaExtraOrcamentaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoResultado As ADODB.Recordset
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
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
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

Private Sub GroupHeader1_Format()
    If txtintContribuinte <> "" Then
        txt_Total = gstrConvVrDoSql(dblTotal)
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
    txtstrContaContabil.Text = gvntFormatacaoEspecifica(txtstrContaContabil.Text, 1)
    If txtdblValor.Text <> "" Then
        txtdblValor.Text = gstrConvVrDoSql(txtdblValor.Text)
    End If
End Sub

Private Function dblTotal() As Double
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "SELECT SUM(dblValor) AS TOTAL FROM "
    strSQL = strSQL & gstrDespesaExtraOrcamentaria & " "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " intContribuinte = " & Val(txtintContribuinte.Text)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                dblTotal = gstrConvVrDoSql(!Total)
                .MoveNext
            Loop
        End With
    End If
End Function

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
