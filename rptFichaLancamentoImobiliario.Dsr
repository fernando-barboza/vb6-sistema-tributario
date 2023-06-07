VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptFichaLancamentoImobiliario 
   Caption         =   "Tributario - rptFichaLancamentoImobiliario (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptFichaLancamentoImobiliario.dsx":0000
End
Attribute VB_Name = "rptFichaLancamentoImobiliario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_Format()
    txtdblFator = gstrConvVrDoSql(gstrENulo(txtdblFator), 2)
End Sub

Private Sub GroupFooter3_Format()
    Dim strSql      As String
    Dim adoResultado As ADODB.Recordset
    
    ' QUERY PARA CARREGAR TIPO / PADRÃO DA CONSTRUÇÃO
    strSql = "SELECT " & _
             "LPI.strNomePadrao, " & _
             "LPI.strNomeUso, " & _
             "LPI.Strnomedaarea, " & _
             "LPI.Strnomeuso, " & _
             "LPI.Dblmedidadaarea, " & _
             "LPI.Dblvalormetro, " & _
             "LPI.Dblfatorobsolescencia "
    strSql = strSql & "FROM " & gstrLancamentoAlfa & " LA, " & _
                                gstrLancamentoIPTU & " LI, " & _
                                gstrLancamentoPredioIPTU & " LPI "
           
    strSql = strSql & "Where LA.Pkid = " & txtIDLancamentoContabil & _
                      " AND li.intlancamentoalfa = la.pkid " & _
                      " AND lpi.intlancamentoiptu = li.pkid "

    'If Not adoDataControl.Recordset.BOF Then
    '    adoDataControl.Recordset.MoveFirst
    'End If

    Set adoResultado = New ADODB.Recordset
    Set gobjBanco = New clsBanco
  
    With rptSubFichaLancamentoImobiliario
   
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            
            If adoResultado.EOF Then
                SubReport1.Visible = False
                Exit Sub
            Else
                SubReport1.Visible = True
            End If
            
            If bytDBType = EDatabases.SQLServer Then
                .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
            Else
                .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
            End If
            Set .adoDataControl.Recordset = adoResultado
            Set SubReport1.object = rptSubFichaLancamentoImobiliario
            
            rptSubFichaLancamentoImobiliario.lngIDLancamentoContabil = txtIDLancamentoContabil.Text
            
        End If

    End With

End Sub

Private Sub GroupHeader1_Format()
    txt_strIdentificacao = gstrFormataInscricao(txt_strIdentificacao)
End Sub

Private Sub GroupHeader2_Format()
    FormataValores
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
    If adoDataControl.Recordset.RecordCount = 0 Then
       ExibeMensagem "Não existe(m) registro(s) com os dados requisitados."
       Unload Me
       Exit Sub
    End If
    
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogoTipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
End Sub
Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
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

Private Sub FormataValores()

    txtintCep = gstrCEPFormatado(txtintCep)
    txtintCepC = gstrCEPFormatado(txtintCepC)
    txtDblareaterreno = gstrConvVrDoSql(TrocaNuloPorZero(txtDblareaterreno), 2)
    txtdblvalorvenalterreno = gstrConvVrDoSql(TrocaNuloPorZero(txtdblvalorvenalterreno), 2)
    txtdblImpostoTerreno = gstrConvVrDoSql(TrocaNuloPorZero(txtdblImpostoTerreno), 2)
    txtdblAreaExcedente = gstrConvVrDoSql(TrocaNuloPorZero(txtdblAreaExcedente), 2)
    txtdblImpostoExcedente = gstrConvVrDoSql(TrocaNuloPorZero(txtdblImpostoExcedente), 2)
    txtdblValorVenalExedente = gstrConvVrDoSql(TrocaNuloPorZero(txtdblValorVenalExedente), 2)
    txtDblareaterreno = gstrConvVrDoSql(TrocaNuloPorZero(txtDblareaterreno), 2)
    txtDblvalorvenalpredio = gstrConvVrDoSql(TrocaNuloPorZero(txtDblvalorvenalpredio), 2)
    txtdblImpostoPredio = gstrConvVrDoSql(TrocaNuloPorZero(txtdblImpostoPredio), 2)
    
    txtdblAreaTerrenoTotal = gstrConvVrDoSql(CDbl(txtDblareaterreno) + CDbl(txtdblAreaExcedente), 2)
    txtdblValorVenalTerrenoTotal = gstrConvVrDoSql(CDbl(txtdblvalorvenalterreno) + CDbl(txtdblValorVenalExedente), 2)
    txtdblImpostoTerrenoTotal = gstrConvVrDoSql(CDbl(txtdblImpostoTerreno) + CDbl(txtdblImpostoExcedente), 2)
    
    txtdblTotalImposto = gstrConvVrDoSql(CDbl(txtdblImpostoTerrenoTotal) + CDbl(txtdblImpostoPredio), 2)
    txtdblTotalValorVenal = gstrConvVrDoSql(CDbl(txtdblValorVenalTerrenoTotal) + CDbl(txtDblvalorvenalpredio), 2)

    
    'txtdblTotalArea = gstrConvVrDoSql(TrocaNuloPorZero(txtdblTotalArea), 2)
    'txtdblTotalImposto = gstrConvVrDoSql(TrocaNuloPorZero(txtdblTotalImposto), 2)
    'txtdblTotalVenal = gstrConvVrDoSql(TrocaNuloPorZero(txtdblTotalVenal), 2)
    
    'txtdblTotalVenal = gstrConvVrDoSql(CDbl(txtdblValorVenalTerreno) + CDbl(txtdblValorVenalExedente) + CDbl(txtdblvalorvenalpredio), 2)
    'txtdblTotalImposto = gstrConvVrDoSql(CDbl(txtdblImpostoPredio) + CDbl(txtdblImpostoExcedente) + CDbl(txtdblImpostoPredio), 2)
    
    
End Sub

Private Function TrocaNuloPorZero(Valor As Variant) As Variant
    If IsNull(Valor) Or Valor = "" Then
        TrocaNuloPorZero = 0
    Else
        TrocaNuloPorZero = Valor
    End If
End Function
