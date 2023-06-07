VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptNotaDeEmpenho 
   Caption         =   "prjOrcamentario - rptNotaDeEmpenho (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptNotaDeEmpenho.dsx":0000
End
Attribute VB_Name = "rptNotaDeEmpenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFornecedor   As String
Dim strProcesso     As String
Dim strAlmoxarifado As String
Dim strCompras      As String
Dim strTesouraria   As String
Dim strVrEmpenho    As String
Dim blnPag          As Boolean
Dim intPagina       As Integer

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
    
    LeImagemLogotipo imgBrasao, imgLogoTipo, txtstrNomeFantasia
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
    txtdblPrecoTotItem = gstrConvVrDoSql(txtdblPrecoTotItem)
    txtdblPrecoItem = gstrConvVrDoSql(txtdblPrecoItem)


End Sub

Private Sub GroupHeader3_Format()
'    If adoDataControl.NRecords > 0 Then
'        If adoDataControl.Recordset("bytSituacao").Value <> 4 Then
'            INTEMPENHOANULACAO = adoDataControl.Recordset("INTEMPENHO").Value
'            INTEMPENHO = ""
'            lblEstorno.Caption = ""
'        Else
'            INTEMPENHOANULACAO = adoDataControl.Recordset("INTEMPENHOANULACAO").Value
'            INTEMPENHO = adoDataControl.Recordset("INTEMPENHO").Value
'            lblEstorno = "Estorno"
'        End If
'        If Len(adoDataControl.Recordset("strCodigo").Value) > 0 Then
'            txtstrProcesso = adoDataControl.Recordset("strCodigo").Value & "/" & adoDataControl.Recordset("intExercicio").Value & " - " & adoDataControl.Recordset("bitDigito").Value
'        End If
'     End If

End Sub

Private Sub GroupHeader4_Format()
    Dim adoResultado As ADODB.Recordset
    Dim strSql As String
    Dim strFormaComunicacao As String
    blnPag = False
    DBLEMPENHADOATEDATA = gstrConvVrDoSql(DBLEMPENHADOATEDATA)
    DBLSALDOATUAL = gstrConvVrDoSql(DBLSALDOATUAL)
    lblVRDotacaoAtual = gstrConvVrDoSql(CDbl(IIf(Len(Trim(DBLEMPENHADOATEDATA)) = 0, "0,00", DBLEMPENHADOATEDATA)) + CDbl(IIf(Len(Trim(DBLSALDOATUAL)) = 0, "0,00", DBLSALDOATUAL)))
    If adoDataControl.Recordset.EOF Or adoDataControl.Recordset.BOF Then Exit Sub
    txtstrDestino = adoDataControl.Recordset("strDestino").Value
    

    If adoDataControl.NRecords > 0 Then
        If adoDataControl.Recordset("bytSituacao").Value <> 4 Then
            INTEMPENHOANULACAO = adoDataControl.Recordset("INTEMPENHO").Value
            INTEMPENHO = ""
            lblEstorno.Caption = "Original"
        Else
            INTEMPENHOANULACAO = adoDataControl.Recordset("INTEMPENHOANULACAO").Value
            INTEMPENHO = adoDataControl.Recordset("INTEMPENHO").Value
            '09/03/05 - Tipo Anulação de Despesa
            If adoDataControl.Recordset("byttipo").Value = 3 Then
                lblEstorno = "Anul.Despesa"
            Else
                lblEstorno = "Estorno"
            End If
        End If
        
        If Len(adoDataControl.Recordset("strCodigo").Value) > 0 Then
            txtstrProcesso = adoDataControl.Recordset("strCodigo").Value & "/" & adoDataControl.Recordset("intExercicio").Value & " - " & adoDataControl.Recordset("bitDigito").Value
        End If
        
        If Len(adoDataControl.Recordset("strModalidade").Value) > 0 Then
            txtstrModalidade = adoDataControl.Recordset("strModalidadeLicitacao").Value & " " & adoDataControl.Recordset("strModalidade").Value
        End If

        If Len(adoDataControl.Recordset("strSolicitacao").Value) > 0 Then
            txtstrsolicitacao = adoDataControl.Recordset("strSolicitacao").Value
        End If
        
        If Len(adoDataControl.Recordset("strContrato").Value) > 0 Then
            txtstrContrato = adoDataControl.Recordset("strContrato").Value
        End If
        
        If adoDataControl.Recordset("intParcela").Value = 0 And adoDataControl.Recordset("bytSituacao").Value <> 4 Then
            strVrEmpenho = gstrConvVrDoSql(adoDataControl.Recordset("dblValorEmpenho").Value)
        Else
            strVrEmpenho = gstrConvVrDoSql(adoDataControl.Recordset("dblValor").Value)
        End If
        
        If adoDataControl.Recordset("bytSituacao").Value = 4 Then
            DBLVALOR = " - " & strVrEmpenho
        Else
            DBLVALOR = strVrEmpenho
        End If
        
        
'        strSql = ""
'        strSql = strSql & "Select "
'        strSql = strSql & "TC.STRDESCRICAO" & strCONCAT & "': '" & strCONCAT & "FC.STRCONTEUDO AS strTicoComunicacao "
'        strSql = strSql & "From "
'        strSql = strSql & gstrContribuinte & " CB, "
'        strSql = strSql & gstrFormaDeComunicacao & " FC, "
'        strSql = strSql & gstrTipoDeComunicacao & " TC "
'        strSql = strSql & "Where "
'        strSql = strSql & "CB.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " FC.Intcontribuinte And "
'        strSql = strSql & "TC.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " FC.INTTIPODECOMUNICACAO AND "
'        strSql = strSql & "CB.Pkid = " & adoDataControl.Recordset("intCodigoContribuinte").Value
'        Set gobjBanco = New clsBanco
'        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'            With adoResultado
'                If .RecordCount >= 1 Then
'                    .MoveFirst
'                    Do While Not .EOF
'                        strFormaComunicacao = strFormaComunicacao & "     " & !strTicoComunicacao
'                        .MoveNext
'                    Loop
'                    txtstrTicoComunicacao = strFormaComunicacao
'                    strFormaComunicacao = ""
'                End If
'            End With
'        End If
        
        
    End If
        
        
        
End Sub


Private Sub GroupHeader5_Format()
Dim strSql As String
Dim adoResultado As ADODB.Recordset
Dim strFormaComunicacao As String

    If adoDataControl.NRecords > 0 Then
        strSql = ""
        strSql = strSql & "Select "
        strSql = strSql & "TC.STRDESCRICAO" & strCONCAT & "': '" & strCONCAT & "FC.STRCONTEUDO AS strTicoComunicacao "
        strSql = strSql & "From "
        strSql = strSql & gstrContribuinte & " CB, "
        strSql = strSql & gstrFormaDeComunicacao & " FC, "
        strSql = strSql & gstrTipoDeComunicacao & " TC "
        strSql = strSql & "Where "
        strSql = strSql & "CB.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " FC.Intcontribuinte And "
        strSql = strSql & "TC.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " FC.INTTIPODECOMUNICACAO AND "
        strSql = strSql & "CB.Pkid = " & adoDataControl.Recordset("intCodigoContribuinte").Value
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                If .RecordCount >= 1 Then
                    .MoveFirst
                    Do While Not .EOF
                        strFormaComunicacao = strFormaComunicacao & !strTicoComunicacao & vbNewLine
                        .MoveNext
                    Loop
                    txtstrTicoComunicacao = strFormaComunicacao
                    strFormaComunicacao = ""
                End If
            End With
        End If
    End If
End Sub

Private Sub PageFooter_Format()
If blnPag = False Then
    intPagina = 1
    blnPag = True
Else
    intPagina = intPagina + 1
End If
   lblPagina.Caption = intPagina
   
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub


Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

