VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form frmRelSaldoDividaAtivaPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo da Dívida Ativa por Período de Inscrição"
   ClientHeight    =   1830
   ClientLeft      =   3675
   ClientTop       =   5175
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5055
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1725
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   3043
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Saldo da Dívida Ativa por Período de Inscrição"
      TabPicture(0)   =   "frmRelSaldoDividaAtivaPeriodo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra_Pagamentos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Fra_Pagamentos 
         Height          =   1245
         Left            =   90
         TabIndex        =   1
         Top             =   360
         Width           =   4755
         Begin VB.TextBox txt_dtmDtFinal 
            Height          =   285
            Left            =   1950
            MaxLength       =   10
            TabIndex        =   3
            Top             =   750
            Width           =   1155
         End
         Begin VB.TextBox txt_dtmDtInicial 
            Height          =   285
            Left            =   1950
            MaxLength       =   10
            TabIndex        =   2
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label lbl_dtmDtInicial 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data Inicial"
            Height          =   195
            Left            =   960
            TabIndex        =   5
            Top             =   390
            Width           =   915
         End
         Begin VB.Label lbl_dtmDtFinal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data Final"
            Height          =   195
            Left            =   1155
            TabIndex        =   4
            Top             =   750
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmRelSaldoDividaAtivaPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = 1428
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar, gstrSalvar, gstrAplicar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrImprimir

End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub txt_dtmDtInicial_GotFocus()
    MarcaCampo txt_dtmDtInicial
End Sub

Private Sub txt_dtmDtInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDtInicial
End Sub

Private Sub txt_dtmDtInicial_LostFocus()
    txt_dtmDtInicial = gstrDataFormatada(txt_dtmDtInicial)
End Sub

Private Sub txt_dtmDtFinal_GotFocus()
    MarcaCampo txt_dtmDtFinal
End Sub

Private Sub txt_dtmDtFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDtFinal
End Sub

Private Sub txt_dtmDtFinal_LostFocus()
    txt_dtmDtFinal = gstrDataFormatada(txt_dtmDtFinal)
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case UCase(strModoOperacao)
    Case UCase(gstrImprimir)
        If blnDadosOK Then
            rptSaldoDividaAtivaPeriodo.lblRelatorio = rptSaldoDividaAtivaPeriodo.lblRelatorio & txt_dtmDtInicial & " à " & txt_dtmDtFinal
            ImprimeRelatorio rptSaldoDividaAtivaPeriodo, strQueryRelatorio, "Saldo da Dívida Ativa", 300
        End If
    
    Case UCase(gstrNovo)
        txt_dtmDtInicial.Text = ""
        txt_dtmDtFinal.Text = ""
    
    Case UCase(gstrFechar)
        Unload Me

    End Select
    
End Sub

Private Function blnDadosOK() As Boolean
    blnDadosOK = False
    If Len(txt_dtmDtInicial) > 0 Or Len(txt_dtmDtFinal) > 0 Then
        If gblnDataValida(txt_dtmDtInicial) = False Then
            ExibeMensagem "Data inicial incorreta."
            txt_dtmDtInicial.SetFocus
            Exit Function
        ElseIf gblnDataValida(txt_dtmDtFinal) = False Then
            ExibeMensagem "Data final incorreta."
            txt_dtmDtFinal.SetFocus
            Exit Function
        ElseIf CVDate(txt_dtmDtFinal) < CVDate(txt_dtmDtInicial) Then
            ExibeMensagem "Data inicial não poder menor que a data final."
            txt_dtmDtInicial.SetFocus
            Exit Function
        End If
    Else
        ExibeMensagem "É necessário informar o intervalo de datas."
        txt_dtmDtInicial.SetFocus
        Exit Function
    End If
    blnDadosOK = True
End Function

Private Function strQueryRelatorio() As String
    Dim strSQL As String
    
    strSQL = ""

    If bytDBType = EDatabases.Oracle Then
        strSQL = "SELECT "
        strSQL = strSQL & " NVL(COUNT(DA.PKID),0) AS QtdInscrita, "
        strSQL = strSQL & " NVL(COUNT(LP.PKID),0) AS QtdPago, "
        strSQL = strSQL & " NVL(SUM(CASE WHEN LP.PKID IS NULL THEN 0 ELSE LV.dblValor END),0) AS dblValorPago, "
        strSQL = strSQL & " NVL(SUM(LV.dblValor),0) AS dblValor, "
        strSQL = strSQL & " CR.intUtilizacao, "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA, "
        strSQL = strSQL & " LA.intExercicio, "
        strSQL = strSQL & " CR.strDescricao AS strComposicaoDaReceita "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrLancamentoAlfa & " LA, "
        strSQL = strSQL & gstrDativa & " DA, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
        strSQL = strSQL & gstrComposicaoDaReceita & " CR "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & " LV.intLancamentoAlfa = LA.PKId  AND "
        strSQL = strSQL & " LA.PKID = DA.intLancamentoAlfa  AND "
        strSQL = strSQL & " LV.PKId = LP.INTLANCAMENTOVALOR (+) AND "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA = CR.PKId (+) AND "
        strSQL = strSQL & " (LV.bitParcelaValida = 1) AND "
        strSQL = strSQL & " (LV.intLancamentoAlfaAcordo IS NULL) AND "
        strSQL = strSQL & " (CR.intUtilizacao = 4) AND "
        strSQL = strSQL & " DA.dtmDtInscricao BETWEEN " & gstrConvDtParaSql(txt_dtmDtInicial) & " AND " & gstrConvDtParaSql(txt_dtmDtFinal) & " "
        strSQL = strSQL & "GROUP BY "
        strSQL = strSQL & " LA.intExercicio, "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA, "
        strSQL = strSQL & " CR.strDescricao, "
        strSQL = strSQL & " CR.intUtilizacao "
        strSQL = strSQL & "UNION ALL "
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & " NVL(COUNT(DA.PKID),0) AS QtdInscrita, "
        strSQL = strSQL & " NVL(COUNT(LP.PKID),0) AS QtdPago, "
        strSQL = strSQL & " NVL(SUM(CASE WHEN LP.PKID IS NULL THEN 0 ELSE LV.dblValor END),0) AS dblValorPago, "
        strSQL = strSQL & " NVL(SUM(LV.dblValor),0) AS dblValor, "
        strSQL = strSQL & " CR.intUtilizacao, "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA, "
        strSQL = strSQL & " LA.intExercicio, "
        strSQL = strSQL & " CR.strDescricao AS strComposicaoDaReceita "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrLancamentoAlfa & " LA, "
        strSQL = strSQL & gstrDativa & " DA, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
        strSQL = strSQL & gstrComposicaoDaReceita & " CR "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & " LV.intLancamentoAlfa = LA.PKId AND "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA = CR.PKId (+) AND "
        strSQL = strSQL & " LV.PKId = LP.INTLANCAMENTOVALOR (+) AND "
        strSQL = strSQL & " LA.PKID = DA.intLancamentoAlfa AND "
        strSQL = strSQL & " (LV.bitParcelaValida = 1) AND "
        strSQL = strSQL & " (LV.intLancamentoAlfaAcordo IS NULL) AND "
        strSQL = strSQL & " (CR.intUtilizacao <> 4) AND "
        strSQL = strSQL & " (LA.intExercicio < " & Year(gstrDataDoSistema) & ") AND "
        strSQL = strSQL & " (LA.BYTNAOINSCREVEDA = 0) AND "
        strSQL = strSQL & " DA.dtmDtInscricao BETWEEN " & gstrConvDtParaSql(txt_dtmDtInicial) & " AND " & gstrConvDtParaSql(txt_dtmDtFinal) & " "
        strSQL = strSQL & "GROUP BY "
        strSQL = strSQL & " LA.intExercicio, "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA, "
        strSQL = strSQL & " CR.strDescricao, "
        strSQL = strSQL & " CR.intUtilizacao "
        strSQL = strSQL & "ORDER BY "
        strSQL = strSQL & " intExercicio, "
        strSQL = strSQL & " strComposicaoDaReceita "
    ElseIf bytDBType = EDatabases.SQLServer Then
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & " ISNULL(COUNT(DA.PKID),0) AS QtdInscrita, "
        strSQL = strSQL & " ISNULL(COUNT(LP.PKID),0) AS QtdPago, "
        strSQL = strSQL & " ISNULL(SUM(CASE WHEN LP.PKID IS NULL THEN 0 ELSE LV.dblValor END),0) AS dblValorPago, "
        strSQL = strSQL & " ISNULL(SUM(LV.dblValor),0) AS dblValor, "
        strSQL = strSQL & " CR.intUtilizacao, "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA, "
        strSQL = strSQL & " LA.intExercicio, "
        strSQL = strSQL & " CR.strDescricao AS strComposicaoDaReceita "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrLancamentoValor & " LV  WITH (READPAST) "
        strSQL = strSQL & " INNER JOIN " & gstrLancamentoAlfa & " LA  WITH (READPAST)  ON LV.intLancamentoAlfa = LA.PKId "
        strSQL = strSQL & " INNER JOIN " & gstrDativa & " DA WITH (READPAST) ON LA.PKID = DA.intLancamentoAlfa "
        strSQL = strSQL & " LEFT OUTER JOIN " & gstrLancamentoPagamento & " LP  WITH (READPAST)  ON LV.PKId = LP.INTLANCAMENTOVALOR "
        strSQL = strSQL & " LEFT OUTER JOIN " & gstrComposicaoDaReceita & " CR  WITH (READPAST)  ON LA.INTCOMPOSICAODARECEITA = CR.PKId "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & " (LV.bitParcelaValida = 1) AND "
        strSQL = strSQL & " (LV.intLancamentoAlfaAcordo IS NULL) AND "
        strSQL = strSQL & " (CR.intUtilizacao = 4) AND "
        strSQL = strSQL & " DA.dtmDtInscricao BETWEEN " & gstrConvDtParaSql(txt_dtmDtInicial) & " AND " & gstrConvDtParaSql(txt_dtmDtFinal) & " "
        strSQL = strSQL & "GROUP BY "
        strSQL = strSQL & " LA.intExercicio, "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA, "
        strSQL = strSQL & " CR.strDescricao, "
        strSQL = strSQL & " CR.intUtilizacao "
        strSQL = strSQL & "UNION ALL "
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & " ISNULL(COUNT(DA.PKID),0) AS QtdInscrita, "
        strSQL = strSQL & " ISNULL(COUNT(LP.PKID),0) AS QtdPago, "
        strSQL = strSQL & " ISNULL(SUM(CASE WHEN LP.PKID IS NULL THEN 0 ELSE LV.dblValor END),0) AS dblValorPago, "
        strSQL = strSQL & " ISNULL(SUM(LV.dblValor),0) AS dblValor, "
        strSQL = strSQL & " CR.intUtilizacao, "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA, "
        strSQL = strSQL & " LA.intExercicio, "
        strSQL = strSQL & " CR.strDescricao AS strComposicaoDaReceita "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrLancamentoValor & " LV  WITH (READPAST) "
        strSQL = strSQL & " INNER JOIN " & gstrLancamentoAlfa & " LA  WITH (READPAST)  ON LV.intLancamentoAlfa = LA.PKId "
        strSQL = strSQL & " INNER JOIN " & gstrDativa & " DA WITH (READPAST) ON LA.PKID = DA.intLancamentoAlfa "
        strSQL = strSQL & " LEFT OUTER JOIN " & gstrLancamentoPagamento & " LP  WITH (READPAST)  ON LV.PKId = LP.INTLANCAMENTOVALOR "
        strSQL = strSQL & " LEFT OUTER JOIN " & gstrComposicaoDaReceita & " CR  WITH (READPAST)  ON LA.INTCOMPOSICAODARECEITA = CR.PKId "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & " (LV.bitParcelaValida = 1) AND "
        strSQL = strSQL & " (LV.intLancamentoAlfaAcordo IS NULL) AND "
        strSQL = strSQL & " (CR.intUtilizacao <> 4) AND "
        strSQL = strSQL & " (LA.intExercicio < " & Year(gstrDataDoSistema) & ") AND "
        strSQL = strSQL & " (LA.BYTNAOINSCREVEDA = 0) AND "
        strSQL = strSQL & " DA.dtmDtInscricao BETWEEN " & gstrConvDtParaSql(txt_dtmDtInicial) & " AND " & gstrConvDtParaSql(txt_dtmDtFinal) & " "
        strSQL = strSQL & "GROUP BY "
        strSQL = strSQL & " LA.intExercicio, "
        strSQL = strSQL & " LA.INTCOMPOSICAODARECEITA, "
        strSQL = strSQL & " CR.strDescricao, "
        strSQL = strSQL & " CR.intUtilizacao "
        strSQL = strSQL & "ORDER BY "
        strSQL = strSQL & " LA.intExercicio, "
        strSQL = strSQL & " CR.strDescricao "
    End If
    
    strQueryRelatorio = strSQL
End Function
