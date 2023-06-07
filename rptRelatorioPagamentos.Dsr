VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelatorioPagamentos 
   Caption         =   "Tributario - rptRelatorioPagamentos (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptRelatorioPagamentos.dsx":0000
End
Attribute VB_Name = "rptRelatorioPagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ArrayPagamentos() As String
Dim iRow                  As Integer

Private Sub ActiveReport_DataInitialize()

    Fields.Add "strInscricao"
    Fields.Add "intUtilizacao"
    Fields.Add "strComposicao"
    Fields.Add "intExercicio"
    Fields.Add "strEmissao"
    Fields.Add "strContribuinte"
    Fields.Add "intParcela"
    Fields.Add "dtmVcto"
    Fields.Add "dtmMov"
    Fields.Add "dtmPag"
    Fields.Add "dblValor"
    Fields.Add "strObservacao"
    Fields.Add "strAviso"
    
    iRow = LBound(ArrayPagamentos, 2)
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If iRow > UBound(ArrayPagamentos, 2) Then
        EOF = True
        Exit Sub
    End If
    
    'Vamos obter campos que nao existem no array
    Fields("strInscricao") = ArrayPagamentos(0, iRow)
    Fields("intUtilizacao") = ArrayPagamentos(1, iRow)
    Fields("strComposicao") = ArrayPagamentos(2, iRow)
    Fields("intExercicio") = ArrayPagamentos(3, iRow)
    Fields("strEmissao") = ArrayPagamentos(4, iRow)
    Fields("strContribuinte") = ArrayPagamentos(5, iRow)
    Fields("intParcela") = ArrayPagamentos(6, iRow)
    Fields("dtmVcto") = ArrayPagamentos(7, iRow)
    Fields("dtmMov") = ArrayPagamentos(8, iRow)
    Fields("dtmPag") = ArrayPagamentos(9, iRow)
    Fields("dblValor") = gstrConvVrDoSql(ArrayPagamentos(10, iRow), , , True)
    Fields("strObservacao") = ArrayPagamentos(11, iRow)
    Fields("strAviso") = ArrayPagamentos(12, iRow)
    
    EOF = False
    iRow = iRow + 1

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
    End If
End Sub

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub grp_Utilizacao_Format()
    
    txtstrInscricao = gstrFormataInscricao(Right(Fields.Item("strInscricao").Value, gintRetornaTamanhoMascara(Fields.Item("intUtilizacao").Value)), Fields.Item("intUtilizacao").Value)
    
    Select Case Fields.Item("intUtilizacao").Value
        Case Is = TYP_IMOBILIARIA
            txtstrUtilizacao = "Imobiliário"
        Case Is = TYP_ECONOMICA = 2
            txtstrUtilizacao = "Mobiliário"
        Case Is = TYP_DIVIDA_ATIVA = 3
            txtstrUtilizacao = "Dívida Ativa"
        Case Is = TYP_ACORDO = 4
            txtstrUtilizacao = "Acordo"
        Case Is = TYP_PRECO_PUBLICO = 5
            txtstrUtilizacao = "Preço Público"
        Case Is = TYP_ISS_CONSTRUCAO = 6
            txtstrUtilizacao = "ISS Contrução"
    End Select

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

Public Sub InicializaArray(ArrayCampos() As String)
    ArrayPagamentos = ArrayCampos
End Sub
