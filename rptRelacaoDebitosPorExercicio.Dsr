VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelacaoDebitosPorExercicio 
   Caption         =   "Tributario - rptRelacaoDebitosPorExercicio (ActiveReport)"
   ClientHeight    =   9210
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   11550
   MDIChild        =   -1  'True
   _ExtentX        =   20373
   _ExtentY        =   16245
   SectionData     =   "rptRelacaoDebitosPorExercicio.dsx":0000
End
Attribute VB_Name = "rptRelacaoDebitosPorExercicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ArrayExercicios() As String
Dim iRow                  As Integer
'Dim dblPrincipal        As Double
'Dim dblMulta            As Double
'Dim dblJuros            As Double
'Dim dblCorrecao         As Double
'Dim dblTotal            As Double
Dim blnConfig           As Boolean


Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_DataInitialize()
    If Not blnConfig Then
        Fields.Add "strInscricao"
        Fields.Add "intExercicio"
        Fields.Add "strComposicaoDaReceita"
        Fields.Add "strNomeProprietario"
        Fields.Add "strEndereco"
        Fields.Add "dblPrincipal"
        Fields.Add "dblMulta"
        Fields.Add "dblJuros"
        Fields.Add "dblCorrecao"
        Fields.Add "dblTotal"
        Fields.Add "strAcordo"
    End If
    iRow = LBound(ArrayExercicios, 2)
    
'    dblPrincipal = 0
'    dblMulta = 0
'    dblJuros = 0
'    dblCorrecao = 0
'    dblTotal = 0

End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
Dim adoAux As New ADODB.Recordset

    If iRow > UBound(ArrayExercicios, 2) Then
        EOF = True
        Exit Sub
    End If
    
    'Vamos obter campos que nao existem no array
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT LA.Intutilizacao , LA.strInscricao, LA.intExercicio, LA.strComposicaoDaReceita, LA.strNomeProprietario, LA.strLogradouro " & strCONCAT & "', '" & strCONCAT & " strnumero " & strCONCAT & "' '" & strCONCAT & " strcomplemento " & strCONCAT & "' '" & strCONCAT & " strBairro " & strCONCAT & "' - '" & strCONCAT & "' '" & strCONCAT & " strMunicipio strEndereco FROM " & gstrLancamentoAlfa & " LA WHERE LA.Pkid = " & ArrayExercicios(1, iRow), 5, adoAux) Then
        If Not adoAux.EOF Then
            Fields("strInscricao") = gstrFormataInscricao(Right(adoAux("strInscricao").Value, gintRetornaTamanhoMascara(adoAux("intUtilizacao").Value)), adoAux("intUtilizacao").Value)
            Fields("intExercicio") = adoAux("intExercicio")
            Fields("strComposicaoDaReceita") = adoAux("strComposicaoDaReceita")
            Fields("strNomeProprietario") = adoAux("strNomeProprietario")
            Fields("strEndereco") = adoAux("strEndereco")
            Fields("dblPrincipal") = ArrayExercicios(4, iRow)
            Fields("dblMulta") = ArrayExercicios(5, iRow)
            Fields("dblJuros") = ArrayExercicios(6, iRow)
            Fields("dblCorrecao") = ArrayExercicios(7, iRow)
            Fields("dblTotal") = ArrayExercicios(8, iRow)
            Fields("strAcordo") = ArrayExercicios(9, iRow)
            
'            dblPrincipal = dblPrincipal + CDbl(ArrayExercicios(4, iRow))
'            dblMulta = dblMulta + CDbl(ArrayExercicios(5, iRow))
'            dblJuros = dblJuros + CDbl(ArrayExercicios(6, iRow))
'            dblCorrecao = dblCorrecao + CDbl(ArrayExercicios(7, iRow))
'            dblTotal = dblTotal + CDbl(ArrayExercicios(8, iRow))
        End If
    End If
    Set gobjBanco = Nothing
    
    EOF = False
    iRow = iRow + 1

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

Private Sub ActiveReport_ToolbarClick(ByVal tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf tool.ID = 16 Then
        Configura_Relatorio Me, True
        blnConfig = True
    End If
End Sub

Private Sub Detail_Format()
    'TrocaCorParaZebrado lblSombra
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
'    txt_dblPrincipal = gstrConvVrDoSql(dblPrincipal)
'    txt_dblMulta = gstrConvVrDoSql(dblMulta)
'    txt_dblJuros = gstrConvVrDoSql(dblJuros)
'    txt_dblCorrecao = gstrConvVrDoSql(dblCorrecao)
'    txt_dblTotal = gstrConvVrDoSql(dblTotal)

End Sub

Public Sub InicializaArray(ArrayCampos() As String)
    ArrayExercicios = ArrayCampos
End Sub

