VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptTotaisIPTU 
   Caption         =   "Tributario - rptTotaisIPTU (ActiveReport)"
   ClientHeight    =   8670
   ClientLeft      =   1710
   ClientTop       =   690
   ClientWidth     =   11610
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   20479
   _ExtentY        =   15293
   SectionData     =   "rptTotaisIPTU.dsx":0000
End
Attribute VB_Name = "rptTotaisIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strComposicao    As String
Public strEmissao       As String
Public strExercicio     As String
Dim blnDetail           As Boolean


Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
    blnDetail = False
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
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio.Caption = Me.Caption
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

Private Sub GroupHeader1_Format()
        txt_Tipo.Text = IIf(adoDataControl.Recordset("Tipo") = 1, "Lancamentos:", "Isentos / Imunes:")
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()

    If Not blnDetail Then
        blnDetail = True
        txt_QtdeTerreno = dblQueryQdte(1)
        txt_QtdeExcedente = dblQueryQdte(2)
        txt_QtdeTotTerreno = CDbl(txt_QtdeTerreno) + CDbl(txt_QtdeExcedente)
        txt_QtdeTotPredios = dblQueryQdte(3)
        txt_QtdeTotGeral = CDbl(txt_QtdeTotTerreno) + CDbl(txt_QtdeTotPredios)
    Else
        txt_QtdeTerreno = dblQueryQdte(1)
        txt_QtdeExcedente = dblQueryQdte(2)
        txt_QtdeTotTerreno = CDbl(txt_QtdeTerreno) + CDbl(txt_QtdeExcedente)
        txt_QtdeTotPredios = dblQueryQdte(3)
        txt_QtdeTotGeral = CDbl(txt_QtdeTotTerreno) + CDbl(txt_QtdeTotPredios)
    End If


    txt_Dblareaterreno = gstrConvVrDoSql(txt_Dblareaterreno, , , True)
    txt_Dblvalorvenalterreno = gstrConvVrDoSql(txt_Dblvalorvenalterreno, , , True)
    txt_Dblimpostoterreno = gstrConvVrDoSql(txt_Dblimpostoterreno, , , True)
    
    txt_Dblareaexcedente = gstrConvVrDoSql(txt_Dblareaexcedente, , , True)
    txt_Dblvalorterrenoexcedente = gstrConvVrDoSql(txt_Dblvalorterrenoexcedente, , , True)
    txt_Dblimpostoexcedente = gstrConvVrDoSql(txt_Dblimpostoexcedente, , , True)
    
    'Total de Terreno
    txt_TotTerrenoQuantidade = gstrConvVrDoSql((CDbl(txt_Dblareaterreno) + CDbl(txt_Dblareaexcedente)), , , True)
    txt_TotTerrenoValorVenal = gstrConvVrDoSql((CDbl(txt_Dblvalorvenalterreno) + CDbl(txt_Dblvalorterrenoexcedente)), , , True)
    txt_TotTerrenoImposto = gstrConvVrDoSql((CDbl(txt_Dblimpostoterreno) + CDbl(txt_Dblimpostoexcedente)), , , True)
    
    'Total de Prédios
    txt_TotQuantidade = gstrConvVrDoSql(txt_TotQuantidade, , , True)
    txt_TotValorVenal = gstrConvVrDoSql(txt_TotValorVenal, , , True)
    txt_TotImposto = gstrConvVrDoSql(txt_TotImposto, , , True)
    
    'Total Geral
    txt_TotGQuantidade = gstrConvVrDoSql((CDbl(txt_TotTerrenoQuantidade) + CDbl(txt_TotQuantidade)), , , True)
    txt_TotGValorVenal = gstrConvVrDoSql((CDbl(txt_TotTerrenoValorVenal) + CDbl(txt_TotValorVenal)), , , True)
    txt_TotGImposto = gstrConvVrDoSql((CDbl(txt_TotTerrenoImposto) + CDbl(txt_TotImposto)), , , True)

End Sub

Private Sub ReportFooter_Format()
    Dim adoRelatorio        As ADODB.Recordset
    
    MostraEmissorRelatorio Me
    
    With rptTotaisTPTUparcela
       If gobjBanco.CriaADO(strQueryParcelas, 5, adoRelatorio) Then
           If bytDBType = EDatabases.SQLServer Then
              .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
           Else
              .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
           End If
           Set .adoDataControl.Recordset = adoRelatorio
       End If
     End With
     Set subParcelas.object = rptTotaisTPTUparcela
End Sub

Private Function strQueryParcelas() As String
    Dim strSql          As String
    
    strSql = strSql & "Select "
    strSql = strSql & "PIP.Bytparcelado Parcelado, "
    strSql = strSql & "FPV.INTPARCELA as Parcela, "
    strSql = strSql & "FPV.DTMDTVENCIMENTO as Vencimento "
    strSql = strSql & "From "
    strSql = strSql & gstrParametroIPTU & " PI, "
    strSql = strSql & gstrParametroIPTUPagto & " PIP, "
    strSql = strSql & gstrFormaPagtoVencimentos & " FPV "
    strSql = strSql & "Where "
    strSql = strSql & "PI.Pkid = PIP.Intparametroiptu And "
    strSql = strSql & "PIP.Pkid = FPV.INTFORMAPAGTO And "
    strSql = strSql & "PI.Intcomposicaodareceita = " & Trim(strComposicao) & " And "
    strSql = strSql & "PI.Stremissao = " & strEmissao & " And "
    strSql = strSql & "PI.Intexercicio = " & strExercicio & " "
    strSql = strSql & "Order By "
    strSql = strSql & "PIP.Bytparcelado, "
    strSql = strSql & "FPV.INTPARCELA, "
    strSql = strSql & "FPV.DTMDTVENCIMENTO "
    
    strQueryParcelas = strSql
End Function

Private Function dblQueryQdte(bytTipo As Byte) As Double
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    If bytTipo = 1 Then
        strSql = strSql & "Select "
        strSql = strSql & "count(La.Pkid) as Total "
        strSql = strSql & "From "
        strSql = strSql & gstrComposicaoDaReceita & " CR, "
        strSql = strSql & gstrLancamentoAlfa & " LA, "
        strSql = strSql & gstrLancamentoIPTU & " LI "
        strSql = strSql & "Where "
        strSql = strSql & "CR.Pkid = LA.Intcomposicaodareceita AND "
        strSql = strSql & "LA.Pkid = LI.Intlancamentoalfa AND "
        strSql = strSql & "CR.Pkid = " & strComposicao & " And "
        strSql = strSql & "LA.strEmissao = " & strEmissao & " And "
        strSql = strSql & "LA.Intexercicio = " & strExercicio & " AND "
        strSql = strSql & "Not LI.Dblareaterreno is null AND "
        strSql = strSql & "lI.dblareaterreno > 0 "
    ElseIf bytTipo = 2 Then
        strSql = strSql & "Select "
        strSql = strSql & "count(La.Pkid) as Total "
        strSql = strSql & "From "
        strSql = strSql & gstrComposicaoDaReceita & " CR, "
        strSql = strSql & gstrLancamentoAlfa & " LA, "
        strSql = strSql & gstrLancamentoIPTU & " LI "
        strSql = strSql & "Where "
        strSql = strSql & "CR.Pkid = LA.Intcomposicaodareceita AND "
        strSql = strSql & "LA.Pkid = LI.Intlancamentoalfa AND "
        strSql = strSql & "CR.Pkid = " & strComposicao & " And "
        strSql = strSql & "LA.strEmissao = " & strEmissao & " And "
        strSql = strSql & "LA.Intexercicio = " & strExercicio & " AND "
        strSql = strSql & "Not LI.Dblareaexcedente is null AND "
        strSql = strSql & "lI.Dblareaexcedente > 0 "
    Else
        strSql = strSql & "Select count(*) Total,  "
        strSql = strSql & "Sum(" & gstrISNULL("TT.TotAreaPredio", "0") & ") TotAreaPredio, "
        strSql = strSql & "Sum(" & gstrISNULL("TT.TotAreaPredioExcedente", "0") & ") TotAreaPredioExcedente, "
        strSql = strSql & "Sum(" & gstrISNULL("TT.TOtValorImpostoPredio", "0") & ") TOtValorImpostoPredio "
        strSql = strSql & "From "
        strSql = strSql & "(Select "
        strSql = strSql & "Count(LP.Intlancamentoiptu) Tot, "
        strSql = strSql & "Sum(" & gstrISNULL("LP.Dblmedidadaarea", "0") & ") TotAreaPredio, "
        strSql = strSql & "Sum(" & gstrISNULL("LP.DBLVALORVENALPREDIO", "0") & ") TotAreaPredioExcedente, "
        strSql = strSql & "Sum(" & gstrISNULL("LP.Dblimposto", "0") & ") TOtValorImpostoPredio "
        strSql = strSql & "From "
        strSql = strSql & gstrComposicaoDaReceita & " CR, "
        strSql = strSql & gstrLancamentoAlfa & " LA, "
        strSql = strSql & gstrLancamentoIPTU & " LI, "
        strSql = strSql & gstrLancamentoPredioIPTU & " LP "
        strSql = strSql & "Where "
        strSql = strSql & "CR.Pkid = LA.Intcomposicaodareceita and "
        strSql = strSql & "LA.Pkid = LI.Intlancamentoalfa      and "
        strSql = strSql & "LI.Pkid = LP.Intlancamentoiptu      and "
        strSql = strSql & "CR.Pkid = " & strComposicao & " and "
        strSql = strSql & "LA.strEmissao = " & strEmissao & " and "
        strSql = strSql & "LA.Intexercicio = " & strExercicio & " "
        strSql = strSql & "Group by "
        strSql = strSql & "Intlancamentoiptu) TT "
    End If
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                dblQueryQdte = CDbl(gstrENulo(gstrConvVrDoSql(gstrENulo(!Total), , , True)))
                If bytTipo = 3 Then
                    'Total de Prédios
                    txt_TotQuantidade = gstrConvVrDoSql(!TotAreaPredio, , , True)
                    txt_TotValorVenal = gstrConvVrDoSql(!TotAreaPredioExcedente, , , True)
                    txt_TotImposto = gstrConvVrDoSql(!TOtValorImpostoPredio, , , True)
                End If
            End If
        End With
    End If
    
End Function
