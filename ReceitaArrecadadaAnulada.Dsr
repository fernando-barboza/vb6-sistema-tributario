VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptReceitaArrecadadaAnulada 
   Caption         =   "prjOrcamentario - rptReceitaArrecadadaAnulada (ActiveReport)"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   12000
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   21167
   _ExtentY        =   14631
   SectionData     =   "ReceitaArrecadadaAnulada.dsx":0000
End
Attribute VB_Name = "rptReceitaArrecadadaAnulada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoResultado    As ADODB.Recordset

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

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra
    If txtValorDetail.Text <> "" Then
        txtValorDetail.Text = gstrConvVrDoSql(txtValorDetail.Text)
    End If
    If txtTIPO.Text <> "" Then
        If txtTIPO.Text = "Sim" Then
            MostraContaCodigoOrcamentario
        ElseIf txtTIPO.Text = "Não" Then
            MostraContaPlanoDeConta
        End If
    End If
End Sub

Private Sub GroupHeader1_Format()
    If txtNumeroBanco.Text <> "" Then
        MostraValorTotal
        txtNumeroBanco.Text = gvntFormatacaoEspecifica(txtNumeroBanco.Text, 1)
    End If
End Sub

Private Sub GroupHeader2_Format()
    If txtNumeroDaGuia.Text <> "" Then
        txtNumeroDaGuia.Text = "0000" & txtNumeroDaGuia.Text
    End If
    If txtValorConvenio.Text <> "" Then
        txtValorConvenio.Text = gstrConvVrDoSql(txtValorConvenio.Text)
    End If
    
    If Len(Trim(txtConvenio)) = 0 And Len(Trim(txtValorConvenio.Text)) = 0 Then
       GroupHeader2.Height = 10
    Else
       GroupHeader2.Height = 540
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub MostraValorTotal()
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT SUM(A.dblValorOrcamentario) AS ValorTotal "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " A, "
    strSQL = strSQL & gstrArrecadacaoReceita & " B "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " A.intArrecadacao = B.PKId "
    strSQL = strSQL & " AND B.intContaContabil = " & Val(txtINTCONTA)
    strSQL = strSQL & " AND A.bytCancelado = 0 "
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                If !ValorTotal <> "Null" Then
                    txt_ValorBanco.Text = gstrConvVrDoSql(!ValorTotal)
                End If
            End If
        End With
    End If
    adoResultado.Close
End Sub


Private Sub MostraContaCodigoOrcamentario()
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT strCodigoOrcamentario, strDescricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrCodigoOrcamentario
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PKId = " & Val(txtCONTADETAIL)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txt_NConta.Text = gvntFormatacaoEspecifica(!strCodigoOrcamentario, 2)
                txt_Conta.Text = !strDescricao
            End If
        End With
    End If
    adoResultado.Close
End Sub

Private Sub MostraContaPlanoDeConta()
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT strContaContabil, strDescricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrPlanoConta
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PKId = " & Val(txtCONTADETAIL)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txt_NConta.Text = gvntFormatacaoEspecifica(!strContaContabil, 1)
                txt_Conta.Text = !strDescricao
            End If
        End With
    End If
    adoResultado.Close
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
