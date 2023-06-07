VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptBic2 
   Caption         =   "Formulário BIC"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "rptBic2.dsx":0000
End
Attribute VB_Name = "rptBic2"
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

Private Sub ActiveReport_Initialize()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    Me.Printer.PaperSize = 256
    Me.Printer.PaperHeight = 20000
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
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

    If Val(txtbytRedeEsgoto) = 0 Then
        chkRedeEsgoto.Value = False
    Else
        chkRedeEsgoto.Value = True
    End If
    If Val(txtbytRedeEletrica) = 0 Then
        chkRedeEletrica.Value = False
    Else
        chkRedeEletrica.Value = True
    End If
    If Val(txtbytRedeAgua) = 0 Then
        chkRedeAgua.Value = False
    Else
        chkRedeAgua.Value = True
    End If
    If Val(txtbytRedeTelefonica) = 0 Then
        chkRedeTelefonica.Value = False
    Else
        chkRedeTelefonica.Value = True
    End If
    If Val(txtbytIluminacaoPublica) = 0 Then
        chkIluminacao.Value = False
    Else
        chkIluminacao.Value = True
    End If
    If Val(txtbytColetaDeLixo) = 0 Then
        chkColeta.Value = False
    Else
        chkColeta.Value = True
    End If
    If Val(txtbytGaleriaPluvial) = 0 Then
        chkGaleriaPluvial.Value = False
    Else
        chkGaleriaPluvial.Value = True
    End If
    If Val(txtbytPavimentacao) = 0 Then
        chkPavimentacao.Value = False
    Else
        chkPavimentacao.Value = True
    End If
    If Val(txtbytMeioFio) = 0 Then
        chkMeioFio.Value = False
    Else
        chkMeioFio.Value = True
    End If
    If Val(txtbytVarrecao) = 0 Then
        chkVarrecao.Value = False
    Else
        chkVarrecao.Value = True
    End If
    If Val(txtbytArborizacao) = 0 Then
        chkArborizacao.Value = False
    Else
        chkArborizacao.Value = True
    End If
    If Val(txtbytTrasporte) = 0 Then
        chkTransporte.Value = False
    Else
        chkTransporte.Value = True
    End If

    txtbytPropriedade = Val(txtbytPropriedade) + 1
    txtbytLocalizacao = Val(txtbytLocalizacao) + 1
    txtbytSituacaoJuridica = Val(txtbytSituacaoJuridica) + 1
    txtbytCaracteristicas = Val(txtbytCaracteristicas) + 1
    txtbytOcupacao = Val(txtbytOcupacao) + 1
    txtbytOutros = Val(txtbytOutros) + 1
    txtbytUtilizacao = Val(txtbytUtilizacao) + 1
    txtbytTipo = Val(txtbytTipo) + 1
    txtbytUso = Val(txtbytUso) + 1
    txtbytAgua = Val(txtbytAgua) + 1
    txtbytEsgoto = Val(txtbytEsgoto) + 1
    txtbytPiso = Val(txtbytPiso) + 1
    txtbytEstrutura = Val(txtbytEstrutura) + 1
    txtbytEsquadriaJanela = Val(txtbytEsquadriaJanela) + 1
    txtbytRevestimentoInterno = Val(txtbytRevestimentoInterno) + 1
    txtbytRevestimentoExterno = Val(txtbytRevestimentoExterno) + 1
    txtbytForro = Val(txtbytForro) + 1
    txtbytInstalacaoEletrica = Val(txtbytInstalacaoEletrica) + 1
    txtbytInstalacaoSanitaria = Val(txtbytInstalacaoSanitaria) + 1
    txtbytCobertura = Val(txtbytCobertura) + 1
    txtbytConservacao = Val(txtbytConservacao) + 1
    lblTotalIdade = (Val(txtintAte3Anos) + Val(txtintDe3a7Anos) + Val(txtintDe7a14Anos) + Val(txtintDe14a21Anos) + Val(txtintAcimaDe21Anos))
End Sub

Private Sub PageHeader_Format()
    lblPrefeitura = Trim(gstrCidadeEmpresa)
End Sub

