VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptRelSaldoDividaAtiva 
   Caption         =   "Tributario - rptRelSaldoDividaAtiva (ActiveReport)"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "rptRelSaldoDividaAtiva.dsx":0000
End
Attribute VB_Name = "rptRelSaldoDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strComposicao As String

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

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
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
   TrocaCorParaZebrado lblSombra3
End Sub

Private Sub GroupFooter3_Format()
   TrocaCorParaZebrado lblSombra4
End Sub

Private Sub GroupHeader1_Format()
    TrocaCorParaZebrado lblSombra
   
    If strComposicao = Trim$(txtintExercicio) Then
        lblComposicao.Visible = False
    Else
        lblComposicao.Visible = True
    End If
    
    strComposicao = Trim$(txtintExercicio)
    
End Sub

Private Sub GroupHeader2_Format()
   TrocaCorParaZebrado lblSombra2
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
    txtData = gstrDataDoSistema(False, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub

