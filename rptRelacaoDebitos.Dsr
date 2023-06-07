VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelacaoDebitos 
   Caption         =   "rptRelacaoDebitos (ActiveReport)"
   ClientHeight    =   9045
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   11385
   MDIChild        =   -1  'True
   _ExtentX        =   20082
   _ExtentY        =   15954
   SectionData     =   "rptRelacaoDebitos.dsx":0000
End
Attribute VB_Name = "rptRelacaoDebitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnConfig           As Boolean
Private ArrayParcelas() As String
Dim iRow                As Integer
Dim dblPrincipal        As Double
Dim dblMulta            As Double
Dim dblJuros            As Double
Dim dblCorrecao         As Double
Dim dblTotal            As Double 'Não soma quem está em acordo

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
        Fields.Add "dtmVencimento"
        Fields.Add "intParcela"
        Fields.Add "dblPrincipal"
        Fields.Add "dblMulta"
        Fields.Add "dblJuros"
        Fields.Add "dblCorrecao"
        Fields.Add "dblTotal"
        Fields.Add "strAcordo"
    End If
    iRow = LBound(ArrayParcelas, 2)
    
    dblPrincipal = 0
    dblMulta = 0
    dblJuros = 0
    dblCorrecao = 0
    dblTotal = 0

End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
Dim adoAux As New ADODB.Recordset
Dim strSQL As String

    If iRow > UBound(ArrayParcelas, 2) Then
        EOF = True
        Exit Sub
    End If
    
    strSQL = "SELECT LA.strInscricao, " & _
             "LA.intExercicio, " & _
             "LA.strComposicaoDaReceita, " & _
             "LA.strNomeProprietario, " & _
             "LA.strLogradouro " & strCONCAT & "', '" & strCONCAT & " LA.strnumero " & strCONCAT & "' '" & strCONCAT & " LA.strcomplemento " & strCONCAT & "' '" & strCONCAT & " LA.strBairro " & strCONCAT & "' - '" & strCONCAT & "' '" & strCONCAT & " LA.strMunicipio strEndereco, " & _
             "LA.intUtilizacao"
    strSQL = strSQL & " FROM " & gstrLancamentoAlfa & " LA "
    strSQL = strSQL & " WHERE LA.Pkid = " & ArrayParcelas(1, iRow)
    
    'Vamos obter campos que nao existem no array
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoAux) Then
        If Not adoAux.EOF Then
            Fields("strInscricao") = gstrFormataInscricao(Right(adoAux("strInscricao").Value, gintRetornaTamanhoMascara(adoAux("intUtilizacao").Value)), adoAux("intUtilizacao").Value)
            Fields("intExercicio") = adoAux("intExercicio")
            Fields("strComposicaoDaReceita") = adoAux("strComposicaoDaReceita")
            Fields("strNomeProprietario") = adoAux("strNomeProprietario")
            Fields("strEndereco") = adoAux("strEndereco")
            Fields("dtmVencimento") = ArrayParcelas(12, iRow)
            Fields("intParcela") = ArrayParcelas(2, iRow)
            Fields("dblPrincipal") = ArrayParcelas(4, iRow)
            Fields("dblMulta") = ArrayParcelas(5, iRow)
            Fields("dblJuros") = ArrayParcelas(6, iRow)
            Fields("dblCorrecao") = ArrayParcelas(7, iRow)
            Fields("dblTotal") = ArrayParcelas(8, iRow)
            Fields("strAcordo") = ArrayParcelas(9, iRow)
            
            If Len(Trim(ArrayParcelas(9, iRow))) <= 1 Then
               dblPrincipal = dblPrincipal + CDbl(ArrayParcelas(4, iRow))
               dblMulta = dblMulta + CDbl(ArrayParcelas(5, iRow))
               dblJuros = dblJuros + CDbl(ArrayParcelas(6, iRow))
               dblCorrecao = dblCorrecao + CDbl(ArrayParcelas(7, iRow))
               dblTotal = dblTotal + CDbl(ArrayParcelas(8, iRow))
            End If
            
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
    txt_dblPrincipal = gstrConvVrDoSql(dblPrincipal)
    txt_dblMulta = gstrConvVrDoSql(dblMulta)
    txt_dblJuros = gstrConvVrDoSql(dblJuros)
    txt_dblCorrecao = gstrConvVrDoSql(dblCorrecao)
    txt_dblTotal = gstrConvVrDoSql(dblTotal)

End Sub

Public Sub InicializaArray(ArrayCampos() As String)
    ArrayParcelas = ArrayCampos
End Sub
