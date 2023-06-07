VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptNotaDeCancelamento 
   Caption         =   "rptNotaDeCancelamento (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptNotaDeCancelamento.dsx":0000
End
Attribute VB_Name = "rptNotaDeCancelamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dblVarTotal As Double
Public blnAnulacaoReceita As Boolean

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
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia
    'lblRelatorio = Me.Caption
    
End Sub
Private Sub ActiveReport_ToolbarClick(ByVal tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf tool.ID = 16 Then
        Configura_Relatorio Me, True
    End If
End Sub

Private Sub Detail_Format()
If adoDataControl.NRecords > 0 Then
   
   txtstrEmpenho = adoDataControl.Recordset("intEmpenho").Value & "/" & Mid(adoDataControl.Recordset("intExercicio").Value, 3, 2)
   txtdblValorCancelado = gstrConvVrDoSql(adoDataControl.Recordset("dblValorParcela").Value)
   
   txtintVinculo = Right(adoDataControl.Recordset("strDotacao").Value, 4)
   txtintUnidade = Mid(adoDataControl.Recordset("strDotacao").Value, 4, 4)
   
   txtdblValorEmpenho = gstrConvVrDoSql(adoDataControl.Recordset("dblValorEmpenho") - ValorEmpenho(Val(adoDataControl.Recordset("PkidEmpenho"))), 2)
    
   txtdblSaldoDoEmpenho = gstrConvVrDoSql(ValorSaldoEmpenho(Val(adoDataControl.Recordset("PkidEmpenho"))), 2)
   'txtdblSaldoDoEmpenho = gstrConvVrDoSql(CDbl(txtdblValorEmpenho) - CDbl(txtdblValorCancelado))
    
End If
End Sub

Private Sub GroupHeader1_Format()
   
   lblPagina = "Folha : " & pageNumber
    
   
   lblData = "Data : " & adoDataControl.Recordset("dtmData").Value
   txtintContribuinte = adoDataControl.Recordset("intContribuinte").Value & " - " & adoDataControl.Recordset("strNome").Value
   
   If Len(adoDataControl.Recordset("strCodigo").Value) > 0 Then
      txtstrProcesso = adoDataControl.Recordset("strCodigo").Value & "/" & adoDataControl.Recordset("intExercicioProcesso").Value & " - " & adoDataControl.Recordset("bitDigito").Value
   End If
   
   lblHistorico = IIf(Not IsNull(adoDataControl.Recordset("strHistorico").Value), adoDataControl.Recordset("strHistorico").Value, "")
   
   lblExtenso = "***** " & gstrExtenso(gstrConvVrDoSql(adoDataControl.Recordset("dblValorParcela").Value)) & " *****"
   GroupHeader1.Repeat = ddRepeatOnPage
   
   
End Sub

Private Sub PageFooter_Format()
    If Not adoDataControl.Recordset.EOF Then
        Line52.Visible = True
    Else
        Line52.Visible = False
    End If
    
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

Private Function ValorEmpenho(lngPkidEmpenho) As Double
Dim strSql      As String
Dim adoResultado As ADODB.Recordset



    
    strSql = "SELECT " & gstrISNULL("SUM(dblValor)", "0") & " ValorEmpenho"
    strSql = strSql & " FROM "
    strSql = strSql & gstrSubempenho
    strSql = strSql & " WHERE intEmpenho =" & lngPkidEmpenho & " AND"
    strSql = strSql & " PKID < " & Val(adoDataControl.Recordset("PKIDParcela"))
    strSql = strSql & "AND ( (bytSituacao = 4 AND intNumero = 0 ) OR bytSituacao = 3 OR bytSituacao = 2 OR (bytSituacao=1 And intNumero <> 0))"
        
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            ValorEmpenho = gstrConvVrDoSql(adoResultado!ValorEmpenho, 2)
        Else
            ValorEmpenho = gstrConvVrDoSql(0, 2)
        End If
    End If
End Function


Private Function ValorSaldoEmpenho(lngPkidEmpenho) As Double
Dim strSql      As String
Dim adoResultado As ADODB.Recordset

    strSql = "SELECT E.DBLVALOR - SUM(SE.dblvalor) dblvalor FROM "
    strSql = strSql & gstrSubempenho & " SE,  " & gstrEmpenho & " E WHERE SE.intempenho = " & lngPkidEmpenho
    strSql = strSql & " AND ((SE.bytsituacao IN(2,3) AND SE.dtmData <= " & gstrConvDtParaSql(adoDataControl.Recordset("dtmData")) & ") OR "
    strSql = strSql & "(SE.bytsituacao = 4 and SE.intNumero = 0 and "
    strSql = strSql & "SE.dtmData <= " & gstrConvDtParaSql(adoDataControl.Recordset("dtmData")) & ") "
    strSql = strSql & "OR (SE.bytSituacao = 1 AND SE.intNumero <> 0 And SE.dtmData <= " & gstrConvDtParaSql(adoDataControl.Recordset("dtmData")) & "))"
    strSql = strSql & " AND E.PKID = SE.intempenho"
    strSql = strSql & " GROUP BY E.DBLVALOR"
        
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            ValorSaldoEmpenho = gstrConvVrDoSql(adoResultado!dblValor, 2)
        Else
            ValorSaldoEmpenho = gstrConvVrDoSql(0, 2)
        End If
    End If
End Function

