VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptFichaCadastroEconomico 
   Caption         =   "Tributario - rptFichaCadastroEconomico (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptFichaCadastroEconomico.dsx":0000
End
Attribute VB_Name = "rptFichaCadastroEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub grfFeiras_Format()
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset

    strSQL = strSQL & "Select ECF.Intfeira, F.Strdescricao as strFeira, ECF.Inttipofeira, TF.Strdescricao as strTipoFeira, ECF.DBLAREA, ECF.STRNRBOX "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrEconomicoFeira & " ECF, "
    strSQL = strSQL & gstrEconomico & " EC, "
    strSQL = strSQL & gstrFeira & " F, "
    strSQL = strSQL & gstrTipoFeira & " TF "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "EC.Pkid = ECF.Inteconomico AND "
    strSQL = strSQL & "F.Pkid = ECF.Intfeira AND "
    strSQL = strSQL & "TF.Pkid = ECF.Inttipofeira AND "
    strSQL = strSQL & "EC.Pkid = " & txtPkidEconomico
    strSQL = strSQL & " Order By strFeira"
    
    If adoDataControl.Recordset.RecordCount > 0 Then
       adoDataControl.Recordset.MoveFirst
    End If
    
    If Not adoDataControl.Recordset.EOF Then
  
        Set adoResultado = New ADODB.Recordset
        Set gobjBanco = New clsBanco
      
        With rptSubFichaEconomicoFeiras
       
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            
                If adoResultado.RecordCount > 0 Then
            
                    If bytDBType = EDatabases.SQLServer Then
                        .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                    Else
                        .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                    End If
                    
                    Set .adoDataControl.Recordset = adoResultado
                    
                    grfFeiras.Visible = True
                    SubReport2.Visible = True
                    
                    Set SubReport2.object = rptSubFichaEconomicoFeiras
                Else
                    grfFeiras.Visible = False
                    SubReport2.Visible = False
                End If
            End If
    
        End With
     
    End If

End Sub

Private Sub grfISS_Format()
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset
    
    strSQL = "Select "
    strSQL = strSQL & "IE.Pkid intISSEmpresa, "
    strSQL = strSQL & "TI.Pkid intTipoISS, "
    strSQL = strSQL & "Rtrim(Ltrim(TI.Strdescricao)) strTipoISS, "
    strSQL = strSQL & "LS.Pkid intListaServico, "
    
    If bytDBType = EDatabases.SQLServer Then
       strSQL = strSQL & "REPLICATE('0',5 - " & strLen & "(LS.strCodigo))" & strCONCAT & "RTRIM(LTRIM(strCodigo)) " & _
                         strCONCAT & "' - '" & strCONCAT & _
                         " RTRIM(LTRIM(LS.strDescricao)) strListaServico, "
    Else
       strSQL = strSQL & "RTRIM(LTRIM( " & gstrCONVERT(CDT_VARCHAR, "LS.strCodigo,'00000'") & ")) " & _
                         strCONCAT & "' - '" & strCONCAT & _
                         " RTRIM(LTRIM(LS.strDescricao)) strListaServico, "
    End If
    
    strSQL = strSQL & "IE.Dtmissinicio, "
    strSQL = strSQL & "IE.Dtmissfim, "
    strSQL = strSQL & "IE.Intquantidadeiss "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrEconomico & " EC, "
    strSQL = strSQL & "tblissempresa IE, "
    strSQL = strSQL & gstrTipoIss & " TI, "
    strSQL = strSQL & gstrListaServico & " LS "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "Ec.Pkid = IE.Inteconomico AND "
    strSQL = strSQL & "TI.Pkid = IE.Inttipoiss AND "
    strSQL = strSQL & "LS.Pkid = IE.Intlistaservico AND "
    strSQL = strSQL & "EC.Pkid = " & txtPkidEconomico
    
    If adoDataControl.Recordset.RecordCount > 0 Then
       adoDataControl.Recordset.MoveFirst
    End If
    
    If Not adoDataControl.Recordset.EOF Then
  
        Set adoResultado = New ADODB.Recordset
        Set gobjBanco = New clsBanco
      
        With rptSubFichaEconomicoISS
       
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            
                If adoResultado.RecordCount > 0 Then
            
                    If bytDBType = EDatabases.SQLServer Then
                        .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                    Else
                        .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                    End If
                    
                    Set .adoDataControl.Recordset = adoResultado
                    
                    grfFeiras.Visible = True
                    SubReport4.Visible = True
                    
                    Set SubReport4.object = rptSubFichaEconomicoISS
                Else
                    grfFeiras.Visible = False
                    SubReport4.Visible = False
                End If
            End If
    
        End With
     
    End If

End Sub

Private Sub grfPublicidades_Format()
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset

    strSQL = "SELECT HP.intQuantidade, HP.dblArea, HP.strObservacao, TR.strDescricao strTipo "
              
    strSQL = strSQL & "FROM " & _
                      gstrHistoricoPublicidades & " HP, " & _
                      gstrTributo & " TR "

    strSQL = strSQL & "WHERE" & _
                      " HP.intEconomico  = " & txtPkidEconomico & _
                      " and HP.intTributo  = TR.Pkid "
    
    If adoDataControl.Recordset.RecordCount > 0 Then
       adoDataControl.Recordset.MoveFirst
    End If
    
    If Not adoDataControl.Recordset.EOF Then
  
        Set adoResultado = New ADODB.Recordset
        Set gobjBanco = New clsBanco
      
        With rptSubFichaEconomicoPublicidades
       
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            
                If adoResultado.RecordCount > 0 Then
            
                    If bytDBType = EDatabases.SQLServer Then
                        .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                    Else
                        .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                    End If
                    
                    Set .adoDataControl.Recordset = adoResultado
                    
                    grfPublicidades.Visible = True
                    SubReport1.Visible = True
                    
                    Set SubReport1.object = rptSubFichaEconomicoPublicidades
                Else
                    grfPublicidades.Visible = False
                    SubReport1.Visible = False
                End If
            End If
    
        End With
     
    End If

End Sub

Private Sub grfSocios_Format()
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset

    strSQL = "SELECT "
    strSQL = strSQL & "CO.strNome, CO.PKId PKidContribuinte, "
    strSQL = strSQL & "CO.StrCnpjCpf, "
    strSQL = strSQL & gstrISNULL("SE.intNumeroDeCotas", "0") & " Cotas, "
    strSQL = strSQL & "CO.strIdentidade, FC.strConteudo, TC.strDescricao strTipoComunicacao, "
    strSQL = strSQL & "CO.strLogradouroC, CO.intNumeroC, CO.strBairroC, UF.strSigla, CO.intCepC "
    If bytDBType = Oracle Then
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrSocioEconomico & " SE, "
        strSQL = strSQL & gstrSocio & " SO, "
        strSQL = strSQL & gstrContribuinte & " CO, "
        strSQL = strSQL & gstrFormaDeComunicacao & " FC, "
        strSQL = strSQL & gstrTipoDeComunicacao & " TC, "
        strSQL = strSQL & gstrUF & " UF "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "SE.intSocio = SO.PKId AND "
        strSQL = strSQL & "SO.intContribuinte = CO.PKID AND "
        strSQL = strSQL & "FC.intContribuinte " & strOUTJOracle & " =" & strOUTJSQLServer & " CO.PKID AND "
        strSQL = strSQL & "TC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " FC.intTipoDeComunicacao AND "
        strSQL = strSQL & "UF.Pkid = CO.intUFc AND "
    Else
        strSQL = strSQL & " FROM tblSocioEconomico SE INNER JOIN " & _
                      " tblSocio SO ON SE.intSocio = SO.PKId INNER JOIN " & _
                      " tblContribuinte CO ON SO.intContribuinte = CO.PKId INNER JOIN " & _
                      " tblUF UF ON CO.intUFC = UF.PKId LEFT OUTER JOIN " & _
                      " tblFormaDeComunicacao FC ON CO.PKId = FC.intContribuinte LEFT OUTER JOIN " & _
                      " tblTipoDeComunicacao TC ON FC.intTipoDeComunicacao = TC.PKId " & _
                      " WHERE "
    End If
    
    strSQL = strSQL & "SE.intCodEconomico = " & txtPkidEconomico
    
    If adoDataControl.Recordset.RecordCount > 0 Then
       adoDataControl.Recordset.MoveFirst
    End If
    
    If Not adoDataControl.Recordset.EOF Then
  
        Set adoResultado = New ADODB.Recordset
        Set gobjBanco = New clsBanco
      
        With rptSubFichaEconomicoSocios
       
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            
                If adoResultado.RecordCount > 0 Then
            
                    If bytDBType = EDatabases.SQLServer Then
                        .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                    Else
                        .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                    End If
                    
                    Set .adoDataControl.Recordset = adoResultado
                    
                    grfSocios.Visible = True
                    SubReport3.Visible = True
                    
                    Set SubReport3.object = rptSubFichaEconomicoSocios
                Else
                    grfSocios.Visible = False
                    SubReport3.Visible = False
                End If
            End If
    
        End With
     
    End If

End Sub

Private Sub grpInscricao_BeforePrint()
    LineAux1.Y1 = (Frame1.Height + LineAux1.Y1) - 760
    LineAux2.Y1 = LineAux1.Y1
End Sub

Private Sub grpInscricao_Format()

    txtStrinscricao = gstrFormataInscricao(Replace(txtStrinscricao, ".", ""), TYP_ECONOMICA)
    txtstrInscricaoImob = gstrFormataInscricao(Replace(txtstrInscricaoImob, ".", ""), TYP_IMOBILIARIA)
    txtintCep = gstrCEPFormatado(txtintCep)
    txtCNPJCPF = gstrCGCCPFFormatado(txtCNPJCPF)
    
    Select Case txtbytNatureza
        Case Is = 0
            txtbytNatureza = "Física"
        Case Is = 1
            txtbytNatureza = "Jurídica"
        Case Is = 2
            txtbytNatureza = "SC"
        Case Is = 3
            txtbytNatureza = "Outros"
    End Select
    
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

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
    If adoDataControl.NRecords = 0 Then
       ExibeMensagem "Não existe registro para a solicitação."
       Unload Me
       Exit Sub
    End If
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

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
