VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptEditaisNotificacoes 
   Caption         =   "Editais / Notificações de Lançamentos"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   19764
   _ExtentY        =   14526
   SectionData     =   "EditaisNotificacoes.dsx":0000
End
Attribute VB_Name = "rptEditaisNotificacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTextoDoEdital As String

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
    If Not adoDataControl.Recordset.EOF Then
        strTextoDoEdital = BuscaTextoDoEdital(adoDataControl.Recordset!Pkid)
    End If
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
    If Not adoDataControl.Recordset.EOF Then
        txt_strTextoDoEdital.Text = strTextoDoEdital
    End If
End Sub

Private Function BuscaTextoDoEdital(PKIdEdital) As String
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT strTextoDoEdital " & _
    "FROM " & gstrTabelaDeEdital & _
    " WHERE PKId = " & PKIdEdital
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 15, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            BuscaTextoDoEdital = adoResultado!strTextoDoEdital
        End If
    End If
    Set gobjBanco = Nothing
End Function

Private Sub GroupFooter1_Format()
    txtdblCustoDaParcela.Text = gstrConvVrDoSql(txtdblCustoDaParcela.Text)
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
