VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptResumoProcessamentoBaixa 
   Caption         =   "Tributario - rptResumoProcessamentoBaixa (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptResumoProcessamentoBaixa.dsx":0000
End
Attribute VB_Name = "rptResumoProcessamentoBaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ArrayADevolver() As String
Dim iRow                 As Integer

Private Sub ActiveReport_DataInitialize()

    Fields.Add "Conta"
    Fields.Add "NumeroConta"
    Fields.Add "Lote"
    Fields.Add "Baixado"
    Fields.Add "Recebido"
    Fields.Add "Diferenca"
    
    iRow = LBound(ArrayADevolver, 2)
    
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

    If iRow > UBound(ArrayADevolver, 2) Then
        EOF = True
        Exit Sub
    End If
    
    'Vamos obter campos que nao existem no array
    Fields("Conta") = ArrayADevolver(0, iRow)
    Fields("NumeroConta") = ArrayADevolver(4, iRow) & "  -  "
    Fields("Lote") = ArrayADevolver(1, iRow)
    Fields("Recebido") = ArrayADevolver(2, iRow)
    Fields("Baixado") = ArrayADevolver(3, iRow)
    Fields("Diferenca") = CCur(gstrConvVrDoSql(ArrayADevolver(2, iRow), , , True)) - CCur(gstrConvVrDoSql(ArrayADevolver(3, iRow), , , True))
    
    EOF = False
    iRow = iRow + 1

End Sub

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
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

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
   TrocaCorParaZebrado lblSombra
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

Public Sub InicializaArray(ArrayCampos() As String)
    ArrayADevolver = ArrayCampos
End Sub

