VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptDespesaExtra 
   Caption         =   "rptDespesaExtra (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptDespesaExtra.dsx":0000
End
Attribute VB_Name = "rptDespesaExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dblVarTotal             As Double
Dim lngPkid                 As Long
Public intPagina            As Integer
Public blnAnulacaoReceita   As Boolean

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

    Dim strContaContabil As String
    Dim strDescricao As String
    Dim dblValor As Double
    Dim dblValorDesconto As Double

    If adoDataControl.NRecords > 0 Then
       txtValorLiquidado = gstrConvVrDoSql(txtValorLiquidado)
       lngPkid = adoDataControl.Recordset!Pkid
    
        lblEmpenho.Caption = "Empenho"
        lblDotacao.Caption = "Dotação"

        If blnAnulacaoReceita Then
            txt_Descricao.Visible = True
            lblTextoExtra.Visible = False
            If txt_ContaContabil.Text = "Field1" Then txt_ContaContabil.Text = ""
            'dblTotal = gstrConvVrDoSql(CDbl(dblTotal) + CDbl(txtValorLiquidado))
            dblTotal = gstrConvVrDoSql(Val(gstrConvVrParaSql(dblTotal)) + Val(gstrConvVrParaSql(txtValorLiquidado)))
        Else
            gstrDespExtra adoDataControl.Recordset!PKIDDespExtra, strContaContabil, strDescricao, dblValor, dblValorDesconto
            txt_ContaContabil = gvntFormatacaoEspecifica(strContaContabil, 1)
            lblTextoExtra = strDescricao
            txtValorLiquidado = gstrConvVrDoSql(dblValor)  '- dblValorDesconto)
            txt_Descricao.Visible = False
            lblTextoExtra.Visible = True
            lblEmpenho.Caption = "Conta"
            lblDotacao.Caption = "Extra-Orçamentária"
 
        End If
   
   End If
End Sub

Private Sub gstrDespExtra(ByVal PKIDDespExtra As String, _
                          ByRef strContaContabil As String, _
                          ByRef strDescricao As String, _
                          ByRef dblValor As Double, _
                          ByRef dblValorDesconto As Double)
                          
    Dim strSql As String
    Dim adoResultado  As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "SELECT DE.intnumero , "
    strSql = strSql & "PC.strcontacontabil ,"
    strSql = strSql & "PC.strdescricao, "
    strSql = strSql & "DE.dblvalor ,"
    strSql = strSql & "DE.dbldesconto "
    strSql = strSql & "FROM "
    strSql = strSql & gstrDespesaExtraOrcamentaria & " DE,"
    strSql = strSql & gstrPlanoConta & " PC "
    strSql = strSql & " WHERE De.Intcontacontabil = PC.PKID "
    strSql = strSql & " AND DE.PKID = " & PKIDDespExtra
    
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
        If Not adoResultado.EOF Then
            strContaContabil = gstrENulo(adoResultado!strContaContabil)
            strDescricao = gstrENulo(adoResultado!strDescricao)
            dblValor = Val(gstrConvVrParaSql(gstrENulo(adoResultado!dblValor)))
            dblValorDesconto = Val(gstrENulo(adoResultado!dblDesconto))
        End If
    End If

End Sub



Private Sub GroupFooter1_Format()
    ImprimeSub
    rptDescontosDeOPs.lblPagina = "Folha : " & intPagina + 1
    intPagina = 0
End Sub

Private Sub GroupHeader1_Format()
   Dim strEspaco As String
   
   If adoDataControl.Recordset.EOF Or adoDataControl.Recordset.BOF Then
        Exit Sub
   End If
   
   strEspaco = Space$(50)
   
    lblEmpenho.Caption = "Conta Contábil"
    lblDotacao.Caption = "Descrição"
   
    lblTpOrdem = "Despesa Extra - Orçamentária "
   
   BuscaDataVencimento
   
   intPagina = intPagina + 1
   lblPagina = "Folha : " & intPagina
   
   lblNumeroOrdem = "Número : " & adoDataControl.Recordset("intOrdem").Value & "/" & adoDataControl.Recordset("intexercicioOP").Value
   lblData = "Data : " & adoDataControl.Recordset("dtmData").Value
   txtintContribuinte = adoDataControl.Recordset("intContribuinte").Value & " - " & adoDataControl.Recordset("strNome").Value
   
   If Len(adoDataControl.Recordset("strCodigo").Value) > 0 Then
      txtstrProcesso = adoDataControl.Recordset("strCodigo").Value & "/" & adoDataControl.Recordset("intExercicioProcesso").Value & " - " & adoDataControl.Recordset("bitDigito").Value
   End If
   
   If adoDataControl.Recordset("bytNaturezaJuridica").Value = 1 Then
        lblCGCCPF = "CNPJ:"
        txtstrCNPJCPF = gstrCGCCPFFormatado(txtstrCNPJCPF, "PJ")
   Else
        lblCGCCPF = "CPF:"
        txtstrCNPJCPF = gstrCGCCPFFormatado(txtstrCNPJCPF, "PF")
   End If
   
   lblHistorico = IIf(Not IsNull(adoDataControl.Recordset("typHistorico").Value), adoDataControl.Recordset("typHistorico").Value, "")
   lblDesconto = "Total de Desconto : " & gstrConvVrDoSql(adoDataControl.Recordset("dblDesconto").Value)
   dblDesconto = gstrConvVrDoSql(dblDesconto)
   dblTotal = gstrConvVrDoSql(dblTotal)
   'dblTotal.Text = gstrConvVrDoSql(dblVarTotal)
   lblExtenso = "***** " & gstrExtenso(IIf(Len(Trim(gstrConvVrDoSql(dblTotal))) = 0, 0, gstrConvVrDoSql(dblTotal))) & " *****"
   GroupHeader1.Repeat = ddRepeatOnPage
   
   If blnAnulacaoReceita Then
       dblTotal = gstrConvVrDoSql(adoDataControl.Recordset("dblValorTotal").Value)
   End If
   
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
    
    
     Select Case CStr(gstrRetSiglaPref)
     Case "MAUA"
        lblsubtitulo = "SECRETARIA MUNICIPAL DE FINANÇAS - DIVISÃO DE CONTABILIDADE"
     Case "PUBT"
        lblsubtitulo = "DIRETORIA ADMINISTRATIVA FINANCEIRA - DEPARTAMENTO DE CONTABILIDADE"
     Case Else
        lblsubtitulo = ""
     End Select
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

Private Function ValorEmpenho(lngPkidEmpenho) As Double
Dim strSql      As String
Dim adoResultado As ADODB.Recordset

    strSql = "SELECT SUM(dblValor) ValorEmpenho"
    strSql = strSql & " FROM "
    strSql = strSql & gstrSubempenho
    strSql = strSql & " WHERE intEmpenho =" & lngPkidEmpenho & " AND"
    strSql = strSql & " intNumero BETWEEN 1 AND " & Val(adoDataControl.Recordset("intParcela") - 1)
    strSql = strSql & " AND bytSituacao <> 4 "
    strSql = strSql & " GROUP BY intEmpenho"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            ValorEmpenho = gstrConvVrDoSql(adoResultado!ValorEmpenho, 2)
        Else
            ValorEmpenho = gstrConvVrDoSql(0, 2)
        End If
    End If
        

End Function

Private Function strQueryDescontoOPs() As String

Dim strSql As String

''"DECODE(CT.CDC,NULL,'', CT.CDC || ' - ') || CT.STRNOME intContribuinte, " & _

         strSql = "SELECT 'Orçamentário' quebra, (SELECT SUM(Dblvalordesconto) FROM " & gstrOrdemPagamentoDespesaExtra & " WHERE intOrdemPagamento = " & lngPkid & " GROUP BY intOrdemPagamento) Dblvalor, " & _
                 "'Total Descontado' strDescricaoConta, " & _
                 "OP.bytTipo, " & _
                 "OP.INTNUMERO intOrdem, " & _
                 "OP.intExercicio IntExercicioOP, " & _
                 "OP.DTMDATA , " & _
                 gstrCASEWHEN("CT.CDC", "NULL,''", gstrCONVERT(CDT_NVARCHAR, "CT.CDC") & strCONCAT & "' - '") & strCONCAT & " CT.STRNOME  intContribuinte, " & _
                 "CT.strCNPJCPF, " & _
                 "CT.strLogradouroC STRENDERECO, " & _
                 "CT.intNumero, " & _
                 "CT.STRCOMPLEMENTO STRCOMPLEMENTO, " & _
                 "MP.strDescricao STRMUNICIPIO, " & _
                 "UF.strsigla STRUF, " & _
                 "BR.strDescricao STRBAIRRO, " & _
                 "CP.INTCEP "
         strSql = strSql & "FROM " & gstrOrdemPagamento & " OP, " & _
                 gstrContribuinte & " CT, " & _
                 gstrCidade & " MP, " & _
                 gstrUF & " UF, " & _
                 gstrBairro & " BR, " & _
                 gstrCeps & " CP "
         strSql = strSql & "WHERE OP.intContribuinte = CT.Pkid " & _
                  " AND MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio " & _
                  " AND CP.PKID  " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep " & _
                  " AND BR.PKID" & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro " & _
                  " AND UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF " & _
                  " AND op.byttipo = 2 " & _
                  " AND OP.Pkid = " & lngPkid
         
         strSql = strSql & " ORDER BY quebra DESC"
                        
    strQueryDescontoOPs = strSql
            
End Function
Private Sub ImprimeSub()
Dim strQuery        As String
Dim adoRelatorio    As ADODB.Recordset
    
    With rptDescontosDeOPs
        
        strQuery = strQueryDescontoOPs
       
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strQuery, 30, adoRelatorio) Then
        
            
            If (adoRelatorio.RecordCount > 0) Then
                
                If adoRelatorio!dblValor <> 0 Then
            
                    .adoDataControl.ConnectionString = gcncADOMain.ConnectionString
    
                    .adoDataControl.Source = strQuery
                 
                    Set .adoDataControl.Recordset = adoRelatorio
                 
                    Set SubReport.object = rptDescontosDeOPs
                     
                    SubReport.Visible = True
                    GroupFooter1.Visible = True
                
                Else
                    
                    SubReport.Visible = False
                    GroupFooter1.Visible = False
                    
                End If
            Else
            
                 SubReport.Visible = False
                 GroupFooter1.Visible = False
            
            End If
             
        End If
    
    End With

End Sub

Private Sub BuscaDataVencimento()
   
   Dim strSql       As String
   Dim adoResultado As ADODB.Recordset
   
   
   strSql = "SELECT dtmDataVencimento FROM " & gstrOrdemPagamento & " OP, "
   strSql = strSql & gstrOrdemPagamentoDespesaExtra & " OPDE "
   strSql = strSql & " WHERE OP.PKID = OPDE.intOrdemPagamento AND "
   strSql = strSql & " OPDE.intDespesaExtraOrcamentaria = " & adoDataControl.Recordset!Pkid
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         dtmDataVencimento = gstrENulo(adoResultado!dtmDataVencimento)
      End If
   End If
   

End Sub
