VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCertidaoDividaAtiva 
   Caption         =   "Tributario - rptCertidaoDividaAtiva (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptCertidaoDividaAtiva.dsx":0000
End
Attribute VB_Name = "rptCertidaoDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ArrayADevolver   As XArrayDB
Dim iRow                 As Integer
Dim blnConfigRel         As Boolean

Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_DataInitialize()
    If Not blnConfigRel Then
        Fields.Add "Pkid"
        Fields.Add "IntLivro"
        Fields.Add "intFolha"
        Fields.Add "intCertidao"
        Fields.Add "dtmData"
        Fields.Add "strComposicao"
        Fields.Add "intExercicio"
        Fields.Add "strLogradouro"
        Fields.Add "strLote"
        Fields.Add "strQuadra"
        Fields.Add "strLoteamento"
        Fields.Add "strBairro"
        Fields.Add "INTCEP"
        Fields.Add "strInscricao"
        Fields.Add "strContribuinte"
        Fields.Add "strRG"
        Fields.Add "strCPF"
        Fields.Add "strNotificacao"
        Fields.Add "strNotificacao1"
        Fields.Add "strNotificacao2"
        Fields.Add "intParcela"
        Fields.Add "dtmVencimento"
        Fields.Add "Dblvalor"
        Fields.Add "dblCorrecao"
        Fields.Add "dblCorrigido"
        Fields.Add "dblMulta"
        Fields.Add "dblPorcMulta"
        Fields.Add "dblJuros"
        Fields.Add "dblPorcJuros"
        Fields.Add "dblTotal"
        Fields.Add "dblTotalPrincipal"
        Fields.Add "dblTotalDivida"
        Fields.Add "dblIndexadorBase"
        Fields.Add "dblIndexador"
        Fields.Add "strIndexador"
        Fields.Add "dtmDataBase"
        Fields.Add "strCadastro"
    Else
        Fields("dblTotalDivida") = 0
        Fields("dblIndexadorBase") = 0
    End If
    iRow = ArrayADevolver.LowerBound(1)
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
    If iRow > ArrayADevolver.UpperBound(1) Then
        EOF = True
        Exit Sub
    End If
    
    Fields("Pkid") = ArrayADevolver(iRow, 0)
    Fields("IntLivro") = ArrayADevolver(iRow, 1)
    Fields("intFolha") = ArrayADevolver(iRow, 2)
    Fields("intCertidao") = ArrayADevolver(iRow, 26)
    Fields("dtmData") = gstrDataFormatada(ArrayADevolver(iRow, 3))
    Fields("strComposicao") = ArrayADevolver(iRow, 4)
    Fields("intExercicio") = ArrayADevolver(iRow, 5)
    Fields("strLogradouro") = ArrayADevolver(iRow, 6) & " " & ArrayADevolver(iRow, 7) & " - " & ArrayADevolver(iRow, 8)
    'Fields("strLote") = ArrayADevolver(iRow, 0)
    'Fields("strQuadra") = ArrayADevolver(iRow, 0)
    'Fields("strLoteamento") = ArrayADevolver(iRow, 0)
    Fields("strBairro") = ArrayADevolver(iRow, 9) & " " & ArrayADevolver(iRow, 10) & " " & ArrayADevolver(iRow, 11)
    Fields("INTCEP") = Format$(ArrayADevolver(iRow, 12), "00000-000")
    Fields("strInscricao") = gstrFormataInscricao(Replace(ArrayADevolver(iRow, 13), ".", ""), ArrayADevolver(iRow, 35))
    Fields("strContribuinte") = ArrayADevolver(iRow, 14)
    Fields("strRG") = ArrayADevolver(iRow, 16)
    lblCPFCNPJ.Caption = IIf(Len(ArrayADevolver(iRow, 15)) > 11, "C.N.P.J.:", "C.P.F.:") 'CNPJ ou CPF
    Fields("strCPF") = gstrCGCCPFFormatado(ArrayADevolver(iRow, 15))
    Fields("strNotificacao") = ArrayADevolver(iRow, 17) & " " & ArrayADevolver(iRow, 18) & " " & ArrayADevolver(iRow, 19)
    Fields("strNotificacao1") = ArrayADevolver(iRow, 20) & " " & ArrayADevolver(iRow, 21) & " " & ArrayADevolver(iRow, 22)
    Fields("strNotificacao2") = gstrCEPFormatado(ArrayADevolver(iRow, 23))
    Fields("intParcela") = ArrayADevolver(iRow, 27)
    Fields("dtmVencimento") = ArrayADevolver(iRow, 28)
    Fields("Dblvalor") = gstrConvVrDoSql(ArrayADevolver(iRow, 29), 2, , True)
    Fields("dblCorrecao") = gstrConvVrDoSql(ArrayADevolver(iRow, 32), 2, , True)
    Fields("dblCorrigido") = gstrConvVrDoSql(ArrayADevolver(iRow, 33), 2, , True)
    Fields("dblMulta") = gstrConvVrDoSql(ArrayADevolver(iRow, 30), 2, , True)
    Fields("dblPorcMulta") = Format$((gstrConvVrDoSql(ArrayADevolver(iRow, 30), 2, , True) * 100) / gstrConvVrDoSql(ArrayADevolver(iRow, 29), 2, , True), "0.00")
    Fields("dblJuros") = gstrConvVrDoSql(ArrayADevolver(iRow, 31), 2, , True)
    Fields("dblPorcJuros") = Format$((gstrConvVrDoSql(ArrayADevolver(iRow, 31), 2, , True) * 100) / gstrConvVrDoSql(ArrayADevolver(iRow, 29), 2, , True), "0.00")
    Fields("dblTotal") = gstrConvVrDoSql(ArrayADevolver(iRow, 34), 2, , True)
    Fields("dblTotalPrincipal") = gstrConvVrDoSql(CDbl(Fields("dblTotalPrincipal")) + ArrayADevolver(iRow, 29), 2, , True)
    Fields("dblTotalDivida") = gstrConvVrDoSql(CDbl(Fields("dblTotalDivida")) + ArrayADevolver(iRow, 34), 2, , True)
    Fields("dblIndexadorBase") = gstrConvVrDoSql(Fields("dblTotalDivida") / ArrayADevolver(iRow, 25), 4, , True)
    Fields("dblIndexador") = gstrConvVrDoSql(ArrayADevolver(iRow, 25), 4, , True)
    Fields("strIndexador") = "(" & ArrayADevolver(iRow, 24) & ")"
    Fields("dtmDataBase") = gstrDataFormatada(ArrayADevolver(iRow, 3))
    Fields("strCadastro") = ArrayADevolver(iRow, 36)
    EOF = False
    iRow = iRow + 1
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_ReportStart()
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
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
        blnConfigRel = True
    End If
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

Public Sub InicializaArray(ArrayCampos As XArrayDB)
    Set ArrayADevolver = ArrayCampos
End Sub
