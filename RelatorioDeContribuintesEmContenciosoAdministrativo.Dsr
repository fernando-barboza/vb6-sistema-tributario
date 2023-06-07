VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioDeContribuintesEmContenciosoAdministrativo 
   Caption         =   "Relatório de Contribuintes em Contencioso Administrativo"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   17383
   _ExtentY        =   11509
   SectionData     =   "RelatorioDeContribuintesEmContenciosoAdministrativo.dsx":0000
End
Attribute VB_Name = "rptRelatorioDeContribuintesEmContenciosoAdministrativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoResultado            As ADODB.Recordset
Dim intTotalDeContribuintes As Integer
Dim dblTotalGeralDeDebitos  As Double

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
    intTotalDeContribuintes = 0
    dblTotalGeralDeDebitos = 0
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
    If txtPKIdLancamentoCalculo.Text <> "" Then
        MostraTotal
    End If
End Sub

Private Sub GroupHeader1_Format()
    If txtstrNome.Text <> "" Then
        intTotalDeContribuintes = intTotalDeContribuintes + 1
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub MostraTotal()
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " SUM(dblValorParcela) AS Valor "
    
    strSql = strSql & " FROM "
    strSql = strSql & gstrParcelaReceita
    
    strSql = strSql & " WHERE "
    strSql = strSql & " intLancamentoCalculo = " & Val(txtPKIdLancamentoCalculo)
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, AdoResultado) Then
        If Not AdoResultado.EOF Then
            With AdoResultado
                txt_Valor.Text = gstrConvVrDoSql(!Valor)
                dblTotalGeralDeDebitos = dblTotalGeralDeDebitos + CDbl(txt_Valor)
            End With
        Else
            txt_Valor.Text = ""
        End If
    End If
End Sub

Private Sub ReportFooter_Format()
    txt_TotalDeContribuintes = intTotalDeContribuintes
    txt_TotalGeralDeDebitos = gstrConvVrDoSql(dblTotalGeralDeDebitos)
    intTotalDeContribuintes = 0
    dblTotalGeralDeDebitos = 0
    MostraEmissorRelatorio Me
End Sub

