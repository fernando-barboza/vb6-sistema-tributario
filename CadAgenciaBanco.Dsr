VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptAgenciaBanco 
   Caption         =   "Tributario - rptAgenciaBanco (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   -120
   ClientTop       =   1185
   ClientWidth     =   14175
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   25003
   _ExtentY        =   19076
   SectionData     =   "CadAgenciaBanco.dsx":0000
End
Attribute VB_Name = "rptAgenciaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_Activate()
HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
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

Private Sub PageFooter_Format()
    lblPagina.Caption = "P�gina " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
