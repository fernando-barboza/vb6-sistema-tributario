VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptGuiaPrecoPublico 
   Caption         =   "Tributario - rptGuiaPrecoPublico (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptGuiaPrecoPublico.dsx":0000
End
Attribute VB_Name = "rptGuiaPrecoPublico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ArrayGuia()      As String
Dim iRow                 As Integer

Private Sub ActiveReport_DataInitialize()

  On Error GoTo Erro
    
    Fields.Add "strNumGuia"
    Fields.Add "dtmRecolher"
    Fields.Add "strContribuinte"
    Fields.Add "strLogradouro"
    Fields.Add "strBairro"
    Fields.Add "strMunicipio"
    Fields.Add "strUf"
    Fields.Add "strQuadra"
    Fields.Add "strLote"
    Fields.Add "strInscricao"
    Fields.Add "strAviso"
    Fields.Add "strReceitas"
    Fields.Add "strHistorico"
    Fields.Add "dblValor"
    Fields.Add "dblCorrecao"
    Fields.Add "dblMulta"
    Fields.Add "dblJuros"
    Fields.Add "dblTotal"
    Fields.Add "dtmEmissao"
    Fields.Add "strFuncionario"
    Fields.Add "dtmVencimento"
    Fields.Add "strLinhaDig"
    Fields.Add "strCodBarras"
    Fields.Add "strProcesso"
    Fields.Add "intCep"
    
    iRow = LBound(ArrayGuia, 2)
    
    Exit Sub

Erro:
    Resume Next
    
End Sub

Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If iRow > UBound(ArrayGuia, 2) Then
        EOF = True
        Exit Sub
    End If
    
    Fields("strNumGuia") = ArrayGuia(0, iRow)
    Fields("dtmRecolher") = ArrayGuia(1, iRow)
    Fields("strContribuinte") = ArrayGuia(2, iRow)
    Fields("strLogradouro") = ArrayGuia(3, iRow)
    Fields("strBairro") = ArrayGuia(4, iRow)
    Fields("intCep") = ArrayGuia(24, iRow)
    Fields("strMunicipio") = ArrayGuia(21, iRow)
    Fields("strUf") = ArrayGuia(22, iRow)
    Fields("strQuadra") = ArrayGuia(5, iRow)
    Fields("strLote") = ArrayGuia(6, iRow)
    Fields("strInscricao") = ArrayGuia(7, iRow)
    Fields("strAviso") = ArrayGuia(8, iRow)
    Fields("strReceitas") = ArrayGuia(9, iRow)
    Fields("strHistorico") = ArrayGuia(10, iRow)
    Fields("dblValor") = ArrayGuia(11, iRow)
    Fields("dblCorrecao") = ArrayGuia(12, iRow)
    Fields("dblMulta") = ArrayGuia(13, iRow)
    Fields("dblJuros") = ArrayGuia(14, iRow)
    Fields("dblTotal") = ArrayGuia(15, iRow)
    Fields("dtmEmissao") = ArrayGuia(16, iRow)
    Fields("strFuncionario") = ArrayGuia(17, iRow)
    Fields("dtmVencimento") = ArrayGuia(18, iRow)
    Fields("strLinhaDig") = ArrayGuia(19, iRow)
    Fields("strCodBarras") = ArrayGuia(20, iRow)
    Fields("strProcesso") = ArrayGuia(23, iRow)
    
    EOF = False
    iRow = iRow + 1

End Sub

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    LeImagemLogotipo imgBrasao2, imgLogotipo2, txtNomeFantasia2, txtEstado2
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

Public Sub InicializaArray(ArrayCampos() As String)
    ArrayGuia = ArrayCampos
End Sub

Private Sub Detail_Format()
    txtintCep1 = gstrCEPFormatado(gstrVerificaCampoNulo(txtintCep1))
    txtintCep2 = gstrCEPFormatado(gstrVerificaCampoNulo(txtintCep2))
End Sub
