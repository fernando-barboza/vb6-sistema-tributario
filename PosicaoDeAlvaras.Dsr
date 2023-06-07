VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptPosicaoDeAlvaras 
   Caption         =   "Posição de Alvarás"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   19394
   _ExtentY        =   13547
   SectionData     =   "PosicaoDeAlvaras.dsx":0000
End
Attribute VB_Name = "rptPosicaoDeAlvaras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoResultado               As ADODB.Recordset
Dim intTotalAtividade          As Integer
Dim intTotalDeContribuintes    As Integer

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
    intTotalAtividade = 0
    intTotalDeContribuintes = 0
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
    If txtstrNome.Text <> "" Then
        intTotalDeContribuintes = intTotalDeContribuintes + 1
    End If
    If txtstrVigenciaAlvara.Text <> "" Then
        If CVDate(txtstrVigenciaAlvara.Text) > CVDate(gstrDataDoSistema) Then
            txt_Condicao.Text = "Vigente"
        Else
            txt_Condicao.Text = "Vencido"
        End If
    End If
End Sub

Private Sub GroupHeader1_Format()
    If txtAtividade.Text <> "" Then
        intTotalAtividade = intTotalAtividade + 1
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    txt_TotalDeAtividades = intTotalAtividade
    txt_TotalDeContribuintes = intTotalDeContribuintes
    intTotalAtividade = 0
    intTotalDeContribuintes = 0
    MostraEmissorRelatorio Me
End Sub


