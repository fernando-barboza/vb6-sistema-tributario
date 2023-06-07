VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptLancamentosCompCancel 
   Caption         =   "Tributario - rptLancamentosCompCancel (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptLancamentosCompCancel.dsx":0000
End
Attribute VB_Name = "rptLancamentosCompCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ArrayADevolver   As XArrayDB
Dim iRow                 As Integer

Dim dblTotal             As Double
Dim dblTotalCancelamento As Double

Private Sub ActiveReport_DataInitialize()

    Fields.Add "strInscricao"
    Fields.Add "intExercicio"
    Fields.Add "strComposicaoDaReceita"
    Fields.Add "intComposicaoDaReceita"
    Fields.Add "dblValorDevolver"
    Fields.Add "strNomeProprietario"
    
    iRow = ArrayADevolver.LowerBound(1)
    
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

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
Dim adoAux As New ADODB.Recordset

    If iRow > ArrayADevolver.UpperBound(1) Then
        EOF = True
        Exit Sub
    End If
    
    'Vamos obter campos que nao existem no array
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT strInscricao, intExercicio, strComposicaoDaReceita, intComposicaoDaReceita, strNomeProprietario FROM " & gstrLancamentoAlfa & " LA WHERE LA.Pkid = " & ArrayADevolver(iRow, 0), 5, adoAux) Then
        If Not adoAux.EOF Then
            Fields("strInscricao") = adoAux("strInscricao")
            Fields("intExercicio") = adoAux("intExercicio")
            Fields("strComposicaoDaReceita") = adoAux("strComposicaoDaReceita")
            Fields("intComposicaoDaReceita") = adoAux("intComposicaoDaReceita")
            Fields("strNomeProprietario") = adoAux("strNomeProprietario")
            Fields("dblValorDevolver") = ArrayADevolver(iRow, 1)
            
            dblTotal = dblTotal + CCur(ArrayADevolver(iRow, 1) & 0)
            dblTotalCancelamento = dblTotalCancelamento + CCur(ArrayADevolver(iRow, 2) & 0)
        End If
    End If
    Set gobjBanco = Nothing
    
    EOF = False
    iRow = iRow + 1

End Sub

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    'lblRelatorio = Me.Caption
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

Private Sub GroupFooter1_Format()
    txtdblTotalCancelamento.text = gstrConvVrDoSql(dblTotalCancelamento, 2)
End Sub

Private Sub GroupFooter2_Format()
    txtdblTotal.text = gstrConvVrDoSql(dblTotal, 2)
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
    MostraEmissorRelatorio Me
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
   TrocaCorParaZebrado lblSombra
End Sub

Public Sub InicializaArray(ArrayCampos As XArrayDB)
    Set ArrayADevolver = ArrayCampos
End Sub

