VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCadContaBancaria 
   Caption         =   "rptCadContaBancaria (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptCadContaBancaria.dsx":0000
End
Attribute VB_Name = "rptCadContaBancaria"
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


Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_ReportStart()
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    Me.Caption = "Relat�rio de Contas Banc�rias"
    lblRelatorio = Me.Caption
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

Private Sub PageFooter_Format()
    lblPagina.Caption = "P�gina " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra
    With adoDataControl
        If gstrENulo(.Recordset!strDigitoVerificador) <> "" Then
            If gstrENulo(.Recordset!strDigitoVerificador) <> "-" Then
                txtstrConta.Text = gstrENulo(.Recordset!strConta) & "-" & gstrENulo(.Recordset!strDigitoVerificador)
            Else
                txtstrConta.Text = gstrENulo(.Recordset!strConta)
            End If
        End If
    End With
   
   
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
