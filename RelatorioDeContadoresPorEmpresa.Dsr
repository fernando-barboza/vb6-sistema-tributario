VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioDeContadoresPorEmpresa 
   Caption         =   "Relatório de Empresas por Contador"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   17066
   _ExtentY        =   11245
   SectionData     =   "RelatorioDeContadoresPorEmpresa.dsx":0000
End
Attribute VB_Name = "rptRelatorioDeContadoresPorEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoResultado As ADODB.Recordset
Dim intEmpresa As Integer
Dim intTodasEmpresas As Integer
Dim intTotalDeContadores As Integer

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
    intEmpresa = 0
    intTodasEmpresas = 0
    intTotalDeContadores = 0
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
    txtstrCNPJCPF = gstrCGCCPFFormatado(txtstrCNPJCPF.Text)
    intEmpresa = intEmpresa + 1
End Sub

Private Sub GroupFooter1_Format()
    txt_TotalDeEmpresasPorContribuinte = intEmpresa
    intEmpresa = 0
    intTodasEmpresas = intTodasEmpresas + Val(txt_TotalDeEmpresasPorContribuinte)
End Sub

Private Sub GroupHeader1_Format()
    If Val(txtintContador.Text) > 0 Then
        MostraContadorCRC
    End If
    intTotalDeContadores = intTotalDeContadores + 1
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub MostraContadorCRC()
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " B.strNome, A.strCRC "
    
    strSql = strSql & " FROM "
    strSql = strSql & gstrContador & " A, "
    strSql = strSql & gstrContribuinte & " B "
    
    strSql = strSql & " WHERE "
    strSql = strSql & " A.intContribuinte = B.PKId "
    strSql = strSql & " AND A.PKId = " & Val(txtintContador)
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, AdoResultado) Then
        If Not AdoResultado.EOF Then
            With AdoResultado
                txt_Contador.Text = (!strNome)
                txt_CRC.Text = (!strCRC)
            End With
        Else
            txt_Contador.Text = ""
            txt_CRC.Text = ""
        End If
    End If
End Sub

Private Sub ReportFooter_Format()
    txt_TotalDeEmpresasRelacionadas = intTodasEmpresas
    txt_TotalDeContadores = intTotalDeContadores
    intTodasEmpresas = 0
    intTotalDeContadores = 0
    MostraEmissorRelatorio Me
End Sub
