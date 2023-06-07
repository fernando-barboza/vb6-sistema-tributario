VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubRelatorioGuiaDeArrecadacao1 
   Caption         =   "SubRelatorio1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   MDIChild        =   -1  'True
   _ExtentX        =   17727
   _ExtentY        =   12832
   SectionData     =   "SubRelatorioGuiaDeArrecadacao1.dsx":0000
End
Attribute VB_Name = "rptSubRelatorioGuiaDeArrecadacao1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_ReportStart()
    'PadronizaToolBarRelatorio Me
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
'    rptGuiaDeArrecadacaoMunicipal.Pages(0).Orientation = 2
'    and B.intParcelaReceita = 1
End Sub

'ESTE é um sub relatorio de guia de arrecadacao.

Private Sub Detail_Format()
    'txtdblValorParcela = gstrConvVrDoSql(rptGuiaDeArrecadacaoMunicipal.txt_Total1)
End Sub
