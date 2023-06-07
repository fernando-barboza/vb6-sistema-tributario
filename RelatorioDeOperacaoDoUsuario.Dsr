VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioDeOperacaoDoUsuario 
   Caption         =   "Relatório de Operações do Usuário"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10860
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   19156
   _ExtentY        =   14843
   SectionData     =   "RelatorioDeOperacaoDoUsuario.dsx":0000
End
Attribute VB_Name = "rptRelatorioDeOperacaoDoUsuario"
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
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub





'NAO DELETA
'Private Sub MostraContadorCRC()
'Dim strSql As String
'    strSql = ""
'    strSql = strSql & " SELECT "
'    strSql = strSql & " B.strNome, A.strCRC "
'
'    strSql = strSql & " FROM "
'    strSql = strSql & gstrContador & " A, "
'    strSql = strSql & gstrContribuinte & " B "
'
'    strSql = strSql & " WHERE "
'    strSql = strSql & " A.intContribuinte = B.PKId "
'    strSql = strSql & " AND A.PKId = " & Val(txtintContador)
'
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        If Not adoResultado.EOF Then
'            With adoResultado
'                txt_Contador.Text = (!strNome)
'                txt_CRC.Text = (!strCRC)
'            End With
'        Else
'            txt_Contador.Text = ""
'            txt_CRC.Text = ""
'        End If
'    End If
'End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
