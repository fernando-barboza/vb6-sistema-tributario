VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelTotalizacaoExecutivo 
   Caption         =   "Tributario - rptRelTotalizacaoExecutivo (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptRelTotalizacaoExecutivo.dsx":0000
End
Attribute VB_Name = "rptRelTotalizacaoExecutivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ArrayParcelas     As XArrayDB
Dim iRow                  As Integer

Private Sub ActiveReport_DataInitialize()

    Fields.Add "intComposicao"
    Fields.Add "intExercicio"
    Fields.Add "dblValorTotal"
    Fields.Add "strComposicao"
    
    iRow = ArrayParcelas.LowerBound(1)
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)

    If iRow > ArrayParcelas.UpperBound(1) Then
        EOF = True
        Exit Sub
    End If
    
    'Vamos obter campos que nao existem no array
    Fields("intComposicao") = ArrayParcelas(iRow, 0)
    Fields("intExercicio") = ArrayParcelas(iRow, 1)
    Fields("dblValorTotal") = ArrayParcelas(iRow, 2)
    Fields("strComposicao") = ArrayParcelas(iRow, 3)
    
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

Private Sub GroupFooter3_Format()
    TrocaCorParaZebrado lblSombra
End Sub

Private Sub GroupHeader1_Format()
    lblDataProcessado = Date
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
    Set ArrayParcelas = ArrayCampos
End Sub

