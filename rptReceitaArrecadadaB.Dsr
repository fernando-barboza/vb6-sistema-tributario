VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptReceitaArrecadadaB 
   Caption         =   "rptReceitaArrecadadaB (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptReceitaArrecadadaB.dsx":0000
End
Attribute VB_Name = "rptReceitaArrecadadaB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnGrupoConta As Boolean
Dim dblValorSoma As Boolean
Dim strContaAnterior As String
Dim intContaRecord As Integer

Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
    If frmRelatorioPeriodo.chk_ContaPorFolha = vbChecked Then
        grpH_Receita.NewPage = 3
        grpH_Receita.Repeat = ddRepeatAll
        grpH_Conta.NewPage = ddNPBefore
        grpH_Conta.Repeat = ddRepeatOnPageIncludeNoDetail
    End If
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
    lblRelatorio = Me.Caption
    LeImagemLogotipo imgBrasao, imgLogotipo, txtstrNomeFantasia, txtstrEstado
    blnGrupoConta = False
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

Private Sub Detail_Format()
    TrocaCorParaZebrado lblSombra1
    

    txt_dblValorMes = Val(gstrConvVrParaSql(txt_dblValorMes)) + Val(gstrConvVrParaSql(txt_dblValor))
    txt_dblValorAnual = Val(gstrConvVrParaSql(txt_dblValorAnual)) + Val(gstrConvVrParaSql(txt_dblValor))
    
    'If Not adoDataControl.Recordset.EOF And Not adoDataControl.Recordset.BOF Then
    '    txt_dblValorAnual = Val(gstrConvVrParaSql(txt_dblValorMes)) + Val(gstrConvVrParaSql(adoDataControl.Recordset!dblSaldoAnual))
    'End If
    
    txt_dblValor = gstrConvVrDoSql(txt_dblValor)
    txt_dblValorMes = gstrConvVrDoSql(txt_dblValorMes)
    txt_dblValorAnual = gstrConvVrDoSql(txt_dblValorAnual)
    
    txt_Data = gstrDataFormatada(txt_Data)
End Sub

Private Sub grhOrgao_Format()
    TrocaCorParaZebrado lblSombra1
End Sub

Private Sub grpH_Conta_Format()

    If Not adoDataControl.Recordset.EOF And Not adoDataControl.Recordset.BOF Then
        txt_NumeroContaBancaria = gvntFormatacaoEspecifica(Replace(txt_NumeroContaBancaria, ".", ""), IIf(adoDataControl.Recordset!strReceita = "RECEITA ORÇAMENTARIA", 2, 1))
    End If
    
    
    If strContaAnterior <> txt_NumeroContaBancaria Then
        strContaAnterior = txt_NumeroContaBancaria
        txt_dblValorMes = adoDataControl.Recordset!dblSaldoMensal
        txt_dblValorAnual = adoDataControl.Recordset!dblSaldoAnual
        txtdblArrecadar = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblValorPrevisao)) - adoDataControl.Recordset!dblArrecadar)
    End If
    
    If lblSombra1.BackColor = vbWhite Then
          lblSombra1.BackColor = gvntCorZebrado
    End If
    
        
End Sub

Private Sub grpH_Mes_Format()
    
    txt_dblValorMes = gstrENulo(adoDataControl.Recordset!dblSaldoMensal)
    
    If lblSombra1.BackColor = vbWhite Then
          lblSombra1.BackColor = gvntCorZebrado
    End If
    
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
