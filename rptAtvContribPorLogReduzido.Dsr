VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptAtvContribPorLogReduzido 
   Caption         =   "Project1 - rptAtvContribPorLogReduzido (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptAtvContribPorLogReduzido.dsx":0000
End
Attribute VB_Name = "rptAtvContribPorLogReduzido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intAtivos   As Integer
Dim intInativos As Integer

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
    intAtivos = 0
    intInativos = 0
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
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra
    txtstrInscricao = gstrFormataInscricao(Trim(txtstrInscricao), TYP_ECONOMICA)
    fldEndereco = Trim(fldLogradouro) & ", " & Trim(fldNumero) & ", " & Trim(fldComplemento) & ", " & Trim(fldBairro)
    If adoDataControl.NRecords > 0 Then
        fldPrimaria = IIf(adoDataControl.Recordset("blnPrincipal") = 1, "(P)", "(S)")
        fldSituacao = IIf(adoDataControl.Recordset("intCodigo") = 11, "A", "I")
        
        If adoDataControl.Recordset("intCodigo") = 11 Then
            intAtivos = intAtivos + 1
        Else
            intInativos = intInativos + 1
        End If
        
        fldCep = gstrCEPFormatado(adoDataControl.Recordset("CEP"))
    End If
End Sub

Private Sub ReportFooter_Format()
    
    fldAtivos = intAtivos
    fldInativos = intInativos
    fldTotal = CLng(intAtivos) + CLng(intInativos)
    
    
    MostraEmissorRelatorio Me
End Sub

