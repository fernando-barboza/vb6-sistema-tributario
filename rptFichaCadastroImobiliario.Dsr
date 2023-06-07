VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptFichaCadastroImobiliario 
   Caption         =   "Tributario - rptFichaCadastroImobiliario (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptFichaCadastroImobiliario.dsx":0000
End
Attribute VB_Name = "rptFichaCadastroImobiliario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Detail_Format()
    txtstrCnpjCpfPM = gstrCGCCPFFormatado(txtstrCnpjCpfPM)
    txtstrCnpjCpfPP = gstrCGCCPFFormatado(txtstrCnpjCpfPP)
    txtintCep = gstrCEPFormatado(gstrVerificaCampoNulo(txtintCep))
    txtintCepC = gstrCEPFormatado(gstrVerificaCampoNulo(txtintCepC))
End Sub
Private Sub gphGleba_Format()
End Sub

Private Sub gphTerreno_Format()
    CriaDetalheTerreno
End Sub

Private Sub GroupHeader2_Format()
    Dim adoResultado As ADODB.Recordset
    Dim strSql As String

    txtintCep = gstrCEPFormatado(txtintCep)
    txtintCepC = gstrCEPFormatado(txtintCepC)
    
    'Se o logradouro de notificação do imobiliário estiver vazio, buscar do Contribuinte
    If Trim(txtstrLogradouroC.Text) = "" Then
       strSql = ""
       strSql = strSql & "SELECT "
       strSql = strSql & "CO.intNumeroC, "
       strSql = strSql & "CO.strLogradouroC , "
       strSql = strSql & "CO.strComplementoC, "
       strSql = strSql & "CO.intNumeroc, "
       strSql = strSql & "CO.intCepC, "
       strSql = strSql & "CO.strBairroC, "
       strSql = strSql & "MU.strDescricao strMunicipioC , "
       strSql = strSql & "UF.strSigla strUFC "
       strSql = strSql & "FROM "
       strSql = strSql & gstrContribuinte & " CO, "
       strSql = strSql & gstrCidade & " MU, "
       strSql = strSql & gstrUF & " UF "
       strSql = strSql & "WHERE "
       strSql = strSql & "CO.pkID = " & adoDataControl.Recordset("intContribuinte") & " AND "
       strSql = strSql & "MU.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " CO.intMunicipioC AND "
       strSql = strSql & "UF.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " CO.intUFC "
       
       Set adoResultado = New ADODB.Recordset
       If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
          If Not adoResultado.EOF Then
             With adoResultado
               txtstrLogradouroC.Text = gstrENulo(!strlogradouroc)
               txtintNumeroC.Text = gstrENulo(!intNumeroC)
               txtstrComplementoC.Text = gstrENulo(!strComplementoC)
               txtintCepC.Text = gstrCEPFormatado(gstrENulo(!intcepc))
               txtstrBairroC.Text = gstrENulo(!strBairroC)
               txtstrMunicipioC.Text = gstrENulo(!strMunicipioC)
               txtstrUfC.Text = gstrENulo(!strUFC)
             End With
          End If
       End If
    End If
End Sub
Private Sub GroupHeader3_Format()
    Field35 = gstrFormataInscricao(Replace(Field35, ".", ""), TYP_IMOBILIARIA)
    If Not txt_DataCancel.Text = "" Then
       lbl_TituloDataCancel.Visible = True
       txt_DataCancel.Visible = True
    End If
End Sub

Private Sub GroupHeader4_Format()
    CriaFolhaPredio
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub
Private Sub PageHeader_Format()
    
    lblDataHora = gstrDataDoSistema(True, , True)
    
    If Val(txtbytEdificado.Text) = 1 Then
        lblEdificado.Visible = True
    Else
        lblEdificado.Visible = False
    End If
    
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
    'If adoDataControl.NRecords = 0 Then
    '   ExibeMensagem "Não existe registro para a solicitação."
    '   Unload Me
    '   Exit Sub
    'End If
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


Private Sub CriaFolhaPredio()

    Dim strSql As String
    Dim adoResultado As New ADODB.Recordset
   
    strSql = "SELECT CC.pkid PKId_Contrucao, CG.strNomeDaCaracteristica As Caracteristica, " & _
              "DC.STRNOMEDODETALHE As Detalhe, " & _
              "TV.DBLVALOR As Valor, " & _
              "AI.Pkid, " & _
              "AI.intMedidaDaArea, " & _
              "AI.dblFracaoIdeal, " & _
              "CC.strDescricao strCategoriaConstrucao, " & _
              "AI.dtmUltimaReforma, " & _
              "IU.pkid IDImobiliario "
              
    strSql = strSql & "FROM " & _
                      gstrCaracteristicaGeral & " CG, " & _
                      gstrImobiliario & " IU, " & _
                      gstrDetalheDaCaracteristica & " DC, " & _
                      gstrUtilizacaoDaTabelaDeValor & " UTV, " & _
                      gstrCaracteristicaDoImovel & " CI, " & _
                      gstrTabelaDeValor & " TV, " & _
                      gstrAreaImobiliario & " AI, " & _
                      gstrCategoriaConstrucao & " CC "
    
    strSql = strSql & "WHERE" & _
                      " CG.Pkid " & strOUTJOracle & " = CI.Intcodigocaracteristicageral  " & _
                      " AND IU.PKId  = AI.intImobiliario  " & _
                      " AND DC.pkid " & strOUTJOracle & " = CI.Intcodigodetalhedacaracteristi  " & _
                      " AND UTV.PKId " & strOUTJOracle & " =" & strOUTJSQLServer & " CG.intUtilizacaoDaCaracteristica " & _
                      " AND TV.Pkid " & strOUTJOracle & " = DC.Inttabeladevalores  " & _
                      " AND CG.intUtilizacaoDaCaracteristica = 3  " & _
                      " AND CI.Intarea = AI.pkid " & _
                      " AND IU.Pkid = " & txt_IDImob & _
                      " AND AI.intImobiliario = IU.Pkid " & _
                      " AND CC.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " AI.intCategoriaConstrucao "
    
    strSql = strSql & "ORDER BY AI.Pkid,CG.pkid, CG.strNomeDaCaracteristica"
    
    If Not adoDataControl.Recordset.EOF Then
  
        Set adoResultado = New ADODB.Recordset
        Set gobjBanco = New clsBanco
      
        With rptSubFichaImobiliarioPredio
       
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            
                If adoResultado.RecordCount > 0 Then
            
                    If bytDBType = EDatabases.SQLServer Then
                        .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                    Else
                        .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                    End If
                    
                    Set .adoDataControl.Recordset = adoResultado
                    
                    SubReport1.Visible = True
                    
                    Set SubReport1.object = rptSubFichaImobiliarioPredio
                Else
                    SubReport1.Visible = False
                End If
            End If
    
        End With
    Else
'        SubReport1.Visible = False
'        Set rptSubFichaImobiliarioPredio.adoDataControl.Recordset = Nothing
    End If
    
End Sub

Private Sub CriaDetalheTerreno()
    Dim strSql As String
    Dim adoResultado As New ADODB.Recordset
    
    
    strSql = "Select IM.Pkid intImobiliario, "
    strSql = strSql & "IM.strmatricula, "
    strSql = strSql & "IM.strcartorio, "
    strSql = strSql & "IM.intfolha, "
    strSql = strSql & "IM.dtmdtmatricula, "
    strSql = strSql & "IM.dtmdtescritura, "
    strSql = strSql & "IM.intLivro, "
    strSql = strSql & "IM.Dblarea, "
    strSql = strSql & "VMT.Dblvalor dblValorMetroTerreno, "
    strSql = strSql & "TT.Strnomedatestada strTestada, "
    strSql = strSql & "TI.strMedidaDaTestada strTestadaValor , "
    strSql = strSql & "CG.Intcodigodacaracteristica, "
    strSql = strSql & "CG.strNomeDaCaracteristica As Caracteristica, "
    strSql = strSql & "DC.STRNOMEDODETALHE As Detalhe, "
    strSql = strSql & "TV.DBLVALOR As Valor "

    
    strSql = strSql & "From " & gstrImobiliario & "               IM, "
    strSql = strSql & gstrTestadaImobiliario & "        TI, "
    strSql = strSql & gstrTipoDeTestada & "             TT, "
    strSql = strSql & gstrCaracteristicaDoImovel & "   CI, "
    strSql = strSql & gstrCaracteristicaGeral & "       CG, "
    strSql = strSql & gstrUtilizacaoDaTabelaDeValor & " UTV, "
    strSql = strSql & gstrDetalheDaCaracteristica & "   DC, "
    strSql = strSql & gstrTabelaDeValor & "             TV, "
    
    strSql = strSql & gstrValorMetroTerreno & "         VMT, "
    strSql = strSql & gstrHistoricoFaceDeQuadra & "     HFQ, "
    strSql = strSql & gstrFaceDeQuadra & "              FQ, "
    strSql = strSql & gstrLogradouro & "                LO "
    
    strSql = strSql & "Where IM.Pkid = Ti.Intimobiliario "
    strSql = strSql & "AND IM.Pkid = CI.INTCODIGOIMOBILIARIO "
    strSql = strSql & "AND CG.Pkid = CI.Intcodigocaracteristicageral "
    strSql = strSql & "AND UTV.Pkid = CG.Intutilizacaodacaracteristica "
    strSql = strSql & "AND DC.Pkid = CI.Intcodigodetalhedacaracteristi "
    strSql = strSql & "AND TV.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & "DC.Inttabeladevalores "
    strSql = strSql & "AND TT.Pkid = TI.Inttipodetestada "
    
    strSql = strSql & "AND FQ.Pkid = TI.Intfacedequadra "
    strSql = strSql & "AND LO.Pkid = FQ.Intlogradouro "
    strSql = strSql & "AND FQ.Pkid = HFQ.Intfacedequadra "
    strSql = strSql & "AND VMT.Pkid = HFQ.INTVALORMETROTERRENO "
    strSql = strSql & "AND VMT.Intexercicio = " & Year(gstrDataDoSistema)
    
    strSql = strSql & " AND UTV.Pkid = 2 "
    strSql = strSql & "AND TT.Bytprincipal = 1 "
    strSql = strSql & "AND IM.Pkid = " & txt_IDImob
   
    
'    If adoDataControl.Recordset.RecordCount > 0 Then
'       adoDataControl.Recordset.MoveFirst
'    End If
    
    If Not adoDataControl.Recordset.EOF Then
  
        Set adoResultado = New ADODB.Recordset
        Set gobjBanco = New clsBanco
      
        With rptSubFichaImobiliarioTerreno
       
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            
                If adoResultado.RecordCount > 0 Then
            
                    If bytDBType = EDatabases.SQLServer Then
                        .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
                    Else
                        .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
                    End If
                    
                    Set .adoDataControl.Recordset = adoResultado
                    
                    SubReport2.Visible = True
                    
                    Set SubReport2.object = rptSubFichaImobiliarioTerreno
                    
                Else
                    SubReport2.Visible = False
                End If
            End If
    
        End With
     
    End If



End Sub

