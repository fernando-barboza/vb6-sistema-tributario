VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioRolLanctoEconomico 
   Caption         =   "Tributario - rptRelatorioRolLanctoEconomico (ActiveReport)"
   ClientHeight    =   10620
   ClientLeft      =   0
   ClientTop       =   390
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   18733
   SectionData     =   "rptRelatorioRolLanctoEconomico.dsx":0000
End
Attribute VB_Name = "rptRelatorioRolLanctoEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    
    If adoDataControl.Recordset.EOF Then
        ExibeMensagem "Não existem registros com estas especificações."
        Unload Me
        Exit Sub
    End If
    
    TrocaCorParaZebrado lblSombra
    
    PreencheValores adoDataControl.Recordset("PkidAlfa").Value
    PreencheParcelas adoDataControl.Recordset("PkidAlfa").Value
    
    txtstrInscricao.Text = gstrFormataInscricao(Right(txtstrInscricao.Text, gintRetornaTamanhoMascara(TYP_ECONOMICA)), TYP_ECONOMICA)
    txtstrNumeroAviso.Text = Val(txtstrNumeroAviso.Text)
    
End Sub

Private Sub GroupFooter1_Format()
    PreencheTotalizacao
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

Private Sub PreencheValores(lngPKIdAlfa As Long)
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "R.strSigla, SUM(LR.dblValor) dblTotalReceita "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR, "
    strSql = strSql & gstrReceita & " R "
    strSql = strSql & "WHERE "
    strSql = strSql & "R.Pkid = LR.intReceita AND "
    strSql = strSql & "LR.intLancamentoValor = LV.Pkid AND "
    strSql = strSql & "LV.intLancamentoAlfa = " & lngPKIdAlfa & " "
    strSql = strSql & " GROUP BY R.strSigla "
    strSql = strSql & " ORDER BY R.strSigla "
    
    Set gobjBanco = New clsBanco
    Set adoResultado = New ADODB.Recordset
  
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
  
        With rptRelatorioRolLanctoEcoSubVal
            If bytDBType = EDatabases.SQLServer Then
                .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
            Else
                .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
            End If
            Set .adoDataControl.Recordset = adoResultado
        End With
        
            Set SubValores.object = rptRelatorioRolLanctoEcoSubVal
        
        End If
    End If
    
End Sub

Private Sub PreencheParcelas(lngPKIdAlfa As Long)
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "LV.dtmDtVencimento, LV.intParcela "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoValor & " LV "
    strSql = strSql & "WHERE "
    strSql = strSql & "LV.intLancamentoAlfa = " & lngPKIdAlfa & " "
    strSql = strSql & " ORDER BY LV.intParcela "
    
    Set gobjBanco = New clsBanco
    Set adoResultado = New ADODB.Recordset
  
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
  
        With rptRelatorioRolLanctoEcSubParc
            If bytDBType = EDatabases.SQLServer Then
                .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
            Else
                .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
            End If
            Set .adoDataControl.Recordset = adoResultado
        End With
        
            Set SubParcelas.object = rptRelatorioRolLanctoEcSubParc
        
        End If
    End If
End Sub

Private Sub PreencheTotalizacao()
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset

    strSql = ""
    strSql = strSql & " SELECT  R.strSigla, " & _
            " SUM(LR.dblValor) dblValor, " & _
            " Y.dblQuantidade "
    strSql = strSql & " FROM " & gstrLancamentoAlfa & " LA, " & _
            gstrLancamentoValor & " LV, " & _
            gstrLancamentoReceita & " LR, " & _
            gstrLancamentoEconomico & " LE, " & _
            gstrLctEconomicoAtividade & " LEA, " & _
            gstrReceita & " R, " & _
            " (SELECT Count(X.dblQuantidade) dblQuantidade, X.intReceita " & _
            " FROM (SELECT LA1.Pkid dblQuantidade, LR1.intReceita " & _
                   " FROM " & gstrLancamentoAlfa & " LA1, " & _
                            gstrLancamentoValor & " LV1, " & _
                            gstrLancamentoReceita & " LR1 " & _
                   " WHERE LA1.intComposicaoDaReceita = " & Val(txtintComposicao.Text) & " AND LA1.intExercicio = " & Val(txtintExercicio.Text) & " AND " & _
                           " LV1.intlancamentoalfa = LA1.pkid and " & _
                           " LR1.intLancamentoValor = LV1.Pkid " & _
                   " GROUP BY LA1.Pkid, LR1.intReceita) X " & _
            " GROUP BY X.intReceita) Y "
    strSql = strSql & " WHERE"
    strSql = strSql & " LA.intComposicaoDaReceita = " & Val(txtintComposicao.Text) & " AND LA.intExercicio = " & Val(txtintExercicio.Text) & " AND " & _
            " LV.intLancamentoAlfa = LA.Pkid  AND " & _
            " LR.intLancamentoValor = LV.Pkid  AND " & _
            " R.Pkid = LR.intReceita  AND " & _
            " Y.intReceita = LR.intReceita AND " & _
            " LE.intLancamentoAlfa = LA.Pkid  AND " & _
            " LEA.intLancamentoEconomico = LE.Pkid "
    strSql = strSql & " GROUP BY R.strSigla, LR.intReceita, Y.dblQuantidade "
   
    Set gobjBanco = New clsBanco
    Set adoResultado = New ADODB.Recordset
  
    If gobjBanco.CriaADO(strSql, 30, adoResultado) Then
        If Not adoResultado.EOF Then
  
        With rptRelatorioRolLanctoEcoSubTot
            If bytDBType = EDatabases.SQLServer Then
                .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
            Else
                .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
            End If
            Set .adoDataControl.Recordset = adoResultado
        End With
        
            Set SubTotalizacao.object = rptRelatorioRolLanctoEcoSubTot
        
        End If
    End If
    
End Sub

