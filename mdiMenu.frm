VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000009&
   ClientHeight    =   7455
   ClientLeft      =   1890
   ClientTop       =   2295
   ClientWidth     =   8880
   Icon            =   "mdiMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 actBarra 
      Align           =   1  'Align Top
      Height          =   7155
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8880
      _LayoutVersion  =   1
      _ExtentX        =   15663
      _ExtentY        =   12621
      _DataPath       =   ""
      Bands           =   "mdiMenu.frx":1042
      Begin VB.Timer time1 
         Interval        =   60000
         Left            =   1695
         Top             =   2505
      End
      Begin MSComctlLib.ImageList img_ListaIconesEspecificos 
         Left            =   1050
         Top             =   270
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":1E30
               Key             =   "CALCULARREAJUSTE"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":21CC
               Key             =   "LERARQUIVO"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":2964
               Key             =   "INCLUIRITEM"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":2C56
               Key             =   "EXCLUIRITEM"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":2FA8
               Key             =   "PROCESSAMENTOBAIXA"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":3102
               Key             =   "IMPRIMIRGUIA"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":325E
               Key             =   "GUIADEACORDO"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":33BA
               Key             =   "GUIACERTIDAONEGATIVA"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":3516
               Key             =   "GUIACERTIDAOPOSITIVA"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":3672
               Key             =   "GUIARELACAODEDEBITOS"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":37CE
               Key             =   "GUIACERTIDAODIVIDAATIVA"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":392E
               Key             =   "PARCELAMENTODEBITOATUALIZADO"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":3A8A
               Key             =   "CANCELARREATIVAR"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMenu.frx":3BE6
               Key             =   "GUIACERTIDAOPOSITIVAEFEITONEGATIVO"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picProgressao 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   8820
      TabIndex        =   0
      Top             =   6450
      Visible         =   0   'False
      Width           =   8880
      Begin VB.Frame fraProgressao 
         Caption         =   " Processando "
         Height          =   540
         Left            =   -60
         TabIndex        =   1
         Top             =   120
         Width           =   12240
         Begin VB.CommandButton cmd_Cancela 
            Cancel          =   -1  'True
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   10815
            MousePointer    =   1  'Arrow
            TabIndex        =   2
            Top             =   135
            Visible         =   0   'False
            Width           =   1110
         End
         Begin MSComctlLib.ProgressBar prgProgressao 
            Height          =   255
            Left            =   570
            TabIndex        =   3
            Top             =   225
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lblPorCento 
            AutoSize        =   -1  'True
            Caption         =   "100%"
            Height          =   195
            Index           =   1
            Left            =   10335
            TabIndex        =   5
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblPorCento 
            AutoSize        =   -1  'True
            Caption         =   "0%"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   4
            Top             =   225
            Width           =   210
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgConfigura 
      Left            =   600
      Top             =   1110
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      DialogTitle     =   "Informe o nome do arquivo "
      Filter          =   "Arquivo do tipo texto|*.txt|Arquivo de todos os tipo|*.*"
      FilterIndex     =   1
   End
   Begin MSComctlLib.StatusBar staBarraStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   7155
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   16845
            MinWidth        =   16845
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":3D42
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Picture         =   "mdiMenu.frx":3E9E
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img_ListaIconesGeral 
      Left            =   210
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":3FFA
            Key             =   "NOVO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":4156
            Key             =   "SALVAR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":42B2
            Key             =   "IMPRIMIR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":440E
            Key             =   "DELETAR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":456A
            Key             =   "APLICAR"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":46CA
            Key             =   "LOCALIZAR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":4826
            Key             =   "PREENCHERLISTA"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":4BC2
            Key             =   "FECHAR"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":4D1E
            Key             =   "PASTAFECHADA"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":4E7A
            Key             =   "PASTAABERTA"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":5422
            Key             =   "RECORTAR"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":5582
            Key             =   "COPIAR"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":56E6
            Key             =   "COLAR"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":5846
            Key             =   "AJUDA"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":59A6
            Key             =   "SUPORTE"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":5CFE
            Key             =   "NEGRITO"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":5E5A
            Key             =   "ITALICO"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":5FB6
            Key             =   "SUBLINHADO"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":6112
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":64AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMenu.frx":67D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Visible         =   0   'False
      Begin VB.Menu mnuBarra 
         Caption         =   "Exibir barra de ferramentas"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuImprimir 
      Caption         =   "Imprimir"
      Visible         =   0   'False
      Begin VB.Menu itmImprimir 
         Caption         =   "Imprimir Conteúdo do Grid"
      End
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private moMRU As New cMRUFileList ' Most Recently Used Files
Public cdlg As New GCommonDialog  ' Common dialogs class

' Help constants and function declares
Private Const HELP_CONTEXT = &H1        ' 1
Private Const HELP_QUIT = &H2           ' 2
Private Const HELP_INDEX = &H3          ' 3
Private Const HELP_CONTENTS = &H3       ' 3
Private Const HELP_HELPONHELP = &H4     ' 4
Private Const HELP_SETINDEX = &H5       ' 5
Private Const HELP_SETCONTENTS = &H5    ' 5
Private Const HELP_CONTEXTPOPUP = &H8   ' 8
Private Const HELP_FORCEFILE = &H9      ' 9
Private Const HELP_KEY = &H101          ' 257
Private Const HELP_COMMAND = &H102      ' 258
Private Const HELP_PARTIALKEY = &H105   ' 261

Private Declare Function WinHelp Lib "USER32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Private iDoc As Integer

Private m_bFontsLoaded As Boolean

'Nino
Private WithEvents obfWordEditor As cWordWrapper
Attribute obfWordEditor.VB_VarHelpID = -1

Public blnTexto As Boolean

Private Sub actBarra_BandClose(ByVal Band As ActiveBar2LibraryCtl.Band)
    On Error Resume Next
    actBarra.Bands("mnuView").Tools("miVToolbar").Checked = actBarra.Bands("bndFormulario").Visible
End Sub

Private Sub actBarra_BandOpen(ByVal Band As ActiveBar2LibraryCtl.Band, ByVal Cancel As ActiveBar2LibraryCtl.ReturnBool)
    On Error Resume Next
    actBarra.Bands("mnuView").Tools("miVToolbar").Checked = actBarra.Bands("bndFormulario").Visible
End Sub

Private Sub actBarra_ChildBandChange(ByVal Band As ActiveBar2LibraryCtl.Band)
    If Band.Name = "chdTextos" Then
        blnTexto = True
    Else
        blnTexto = False
    End If
End Sub

Private Sub CarregaFormTabelasGerais(intIndice As Integer)
Select Case intIndice
        Case 578
            CarregaForm frmCadFormulaDeCalculos
        Case 2

        'Case 397
        '    CarregaForm frmCadIndexadorEconomico
        Case 588
            CarregaForm frmCadOcorrencia
        Case 589
           ' ConfiguraListBarSubMenuAgentes
        Case 592
           ' ConfiguraListBarSubMenuPlanta
        Case 595
           ' ConfiguraListBarSubMenuReceitas
        Case 598
            CarregaForm frmCadDetalheDaCaracteristica
        Case 599
            CarregaForm frmCadDiasNaoUteis
        'Case 600
        '    CarregaForm frmCadVencimentos
        'ALteraçao feita por hugo
            Case 601
            CarregaForm frmCadTipoDeComunicacao
        Case 602
            CarregaForm frmCadCampoDeInscricao
        Case 603
        
        Case 604
           'ConfiguraListBarSubMenuTextos
        Case 8
            CarregaForm frmCadUnidadeMedida
        Case 758
            CarregaForm frmCadDocumentoEmitido
        Case 607
            CarregaForm frmCadFiscais
        Case 6
            CarregaForm frmCadDocumentos
        Case 1117
            CarregaForm frmIndexadorEconomico
        Case 1119
            CarregaForm frmCadMoedas
        Case 1120
            CarregaForm frmCodigosDeBaixa
        Case 1148
            CarregaForm frmTipoIsencaoImunidade
        Case 1299
            CarregaForm frmDescontosProvisorios
        Case 1393
            CarregaForm frmCadExecutivosAdvogados
        Case Else
            MsgBox "Item não configurado."
   End Select

End Sub

Private Sub CarreFormTabelasImobiliariasUrbanas(intIndice As Integer)
Select Case intIndice
        Case 609
            CarregaForm frmCadTiposDeArea
        Case 610
             CarregaForm frmCadTiposDeTestada
        Case 611
            CarregaForm frmCadMelhoramentosPublicos
        'Case 612
        '    CarregaForm frmCadSecoesLogradouro
        'Case 613 '(strCodItem = JBBE) - RETIRADO 26/07/04 Rafael
        '    CarregaForm frmCadFatoresCorrecao '(Tabela/Imobiliário Urbano/Fatores de Correção)
        Case 1025
            CarregaForm frmCadValorMetroTerreno
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormTabelasEconomicas(intIndice As Integer)
    Select Case intIndice
        Case 615
            CarregaForm frmCadAtividadeEconomica
        Case 1068
            CarregaForm frmCadAtividadesBasicas
        Case 1157
            CarregaForm frmCadTributos
        Case 1158
            CarregaForm frmCadServicos
        Case 1159
            CarregaForm frmCadTipoFeira
        Case 1160
            CarregaForm frmCadFeira
        Case 1173
            CarregaForm frmCadTipoTributo
        'Case 1294
        '    CarregaForm frmCadTipoOcorrenciaProcesso
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormTabelasContribuicaoDeMelhorias(intIndice As Integer)
Select Case intIndice
        'Case 617
            'CarregaForm frmCadTabelaDeEditais
       
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormCadastros(intIndice As Integer)
Select Case intIndice
        Case 15
             gintCodSeguranca = 15
             CarregaForm frmCadContribuinte
             frmCadContribuinte.Tag = "Contribuinte"
             MDIMenu.Tag = "Tributário"
        Case 735
             CarregaForm frmCadImobiliario
        'Case 736 JCC
        '     CarregaForm frmCadImobiliarioRural
        Case 737
             CarregaForm frmCadEconomico
        Case 738
            CarregaForm frmCadContribuicaoMelhorias
        Case 626
            CarregaForm frmCadContador
        Case 627
            CarregaForm frmCadSocio
        Case 628
            CarregaForm frmCadContasBancarias
        Case 629
            CarregaForm frmCadIsencaoImunidade
        Case 452
            CarregaForm frmCadProtocolizacaoProcesso
        Case 1246
            CarregaForm frmParametroLancamento
        Case 1247
            CarregaForm frmParametrosDividaAtiva
        Case 1076
            CarregaForm frmCadLancamentoIPTU
        Case 1242
            CarregaForm frmCadLancamentoISS
        Case 1115
            CarregaForm frmCadGuias
        Case 450
            CarregaForm frmCadCatalogoAssunto
        Case 1151
            CarregaForm frmCadAtualizaValores
        Case 1190
            CarregaForm frmCadLancamentoEconomico
        Case 1207
            CarregaForm frmCadDividaAtiva
        Case 1241
            CarregaForm frmCadPrecoPublico
        Case 1387
            gintCodSeguranca = 1387
            CarregaForm frmCadExecutivosFiscais
        Case 1377
            CarregaForm frmArquivoDistribuidor
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormExpedienteAdministracao(intIndice As Integer)
 Select Case intIndice
        Case 632
             CarregaForm frmCadEmissaoValidadeDeDocumentos
        Case 633
             CarregaForm frmCadDevolucaoDeDocumentos
        Case 634
            CarregaForm frmCadReavaliacaoDeValores
        Case 790
        
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormExpedienteFinanceiro(intIndice As Integer)
Select Case intIndice
        'Case 636 '(strCodItem = JDBA) - RETIRADO 26/07/04 Rafael
        '     CarregaForm frmCadDebito '(Expediente/Conta Corrente Fiscal/Lançamentos em Conta Corrente)
        'Case 646 '(strCodItem = JDBF) - RETIRADO 27/07/04 Rafael
        '     CarregaForm frmCalculoAcrescimosLegais '(Expediente/Conta Corrente Fiscal/Cálculo de Acréscimos Legais)
        Case 3
            ' CarregaForm
        Case 4
            ' CarregaForm
        Case 5
            ' CarregaForm
        Case 6
            ' CarregaForm
        Case 7
            ' CarregaForm
        Case 8
            ' CarregaForm

        Case Else
            MsgBox "Item não configurado."
   End Select

End Sub

Private Sub CarreFormExpedienteFicalizacao(intIndice As Integer)
    Select Case intIndice
        'Case 648 '(strCodItem = JDCA) - RETIRADO 27/07/04 Rafael
        '     CarregaForm frmContNotasFiscais '(Expediente/Fiscalização/Controle de Notas Fiscais)
        'Case 649 '(strCodItem = JDCB) - RETIRADO 27/07/04 Rafael
        '     CarregaForm frmCadOSdeFiscalizacao '(Expediente/Fiscalização/Ordens de Serviço)
        'Case 650 '(strCodItem = JDCC) - RETIRADO 27/07/04 Rafael
        '     CarregaForm frmCadMapaDeAcaoFiscal '(Expediente/Fiscalizacao/Mapa de Ação Fiscal)
        'Case 651 '(strCodItem = JDCD) - RETIRADO 27/07/04 Rafael
        '    CarregaForm frmCadAutoDeInfracao '(Expediente/Fiscalizacao/Autos de Infração)
        Case 5
            ' CarregaForm
        'Case 653 '(strCodItem = JDCF) - RETIRADO 27/07/04 Rafael
        '    CarregaForm frmCadISSQNVariavel '(Expediente/Fiscalizacao/Controle de Declaração de ISSQN Váriavel)
        Case Else
            MsgBox "Item não configurado."
   End Select

End Sub

Private Sub CarreFormTransferenciaParaDividaAtiva(intIndice As Integer)
    Select Case intIndice
        'Case 640 '(srtCodItem = JDBDA) - RETIRADO 26/07/04 Rafael
        '    CarregaForm frmCadTransferenciaParaDividaAtivaPeloSistema '(Expediente/Conta Corrente Fiscal/Transferências para Dívida Ativa/Débitos Gerados pelo Sistema)
        'Case 641 '(srtCodItem = JDBDB) - RETIRADO 26/07/04 Rafael
        '     CarregaForm frmCadTransferenciaParaDividaAtivaManualmente '(Expediente/Conta Corrente Fiscal/Transferências para Dívida Ativa/Débitos Gerados Manualmente)
        Case Else
            MsgBox "Item não configurado."
   End Select

End Sub

Private Sub CarreFormExpedienteContencioso(intIndice As Integer)
  Select Case intIndice
        Case 1
        Case 2
        Case 3

        Case Else
            MsgBox "Item não configurado."
   End Select

End Sub

Private Sub CarreFormExpedienteContenciosoAdministrativo(intIndice As Integer)
  Select Case intIndice
        'Case 656 '(srtCodItem = JDDAA) - RETIRADO 28/07/04 Rafael
        '    CarregaForm frmCadSuspensaoDeExigencia '(Expediente/Cobranca/Menu Administrativo/Suspensão de Exigências)
        'Case 657 '(srtCodItem = JDDAB) - RETIRADO 28/07/04 Rafael
        '    CarregaForm frmCadPrescricaoDeDebitos '(Expediente/Cobrança/Menu Administrativo/Prescrição de Débitos)
        'Case 658 '(srtCodItem = JDDAC) - RETIRADO 28/07/04 Rafael
        '    CarregaForm frmCadCancelamentoDeDebitos '(Expediente/Cobrança/Menu Administrativo/Cancelamento de Débitos)
        'Case 659 '(srtCodItem = JDDAD) - RETIRADO 28/07/04 Rafael
        '    CarregaForm frmCadRemissaoDeDebitos '(Expediente/Cobrança/Menu Administrativo/Remissão de Débitos)
        'Case 660 '(srtCodItem = JDDAE) - RETIRADO 28/07/04 Rafael
        '    CarregaForm frmCadCobrancaExtraJudicial '(Expediente/Cobrança/Menu Administrativo/Cobrança Extra-Judicial)
        Case Else
            MsgBox "Item não configurado."
   End Select

End Sub

Private Sub CarreFormExpedienteContenciosoJudicial(intIndice As Integer)
  
  Select Case intIndice
        'Case 662 '(srtCodItem = JDDBA) - RETIRADO 28/07/04 Rafael
        '    CarregaForm frmCadExecucaoFiscal '(Expediente/Cobrança/Judicial/Execução Fiscal)
   End Select

End Sub

Private Sub CarreFormExpediente(intIndice As Integer)
  Select Case intIndice
        Case 1
             'ConfiguraListBarSubMenuAdministrativo
        Case 2
             'ConfiguraListBarSubMenuFinanceiro
        Case 3
             'ConfiguraListBarSubMenuFiscalizacao
        Case 4
             'ConfiguraListBarSubMenuContencioso
        Case 5
            'CarregaForm
        Case 6
             'ConfiguraListBarSubMenuCalculos
        Case 1161
            CarregaForm frmAtualizacaoDebitos
        Case 1223
            'CarregaForm frmCadDividaAtivaManual
        Case 1261
            CarregaForm frmAlteracaoEndImobiliario
        Case 1262
            CarregaForm frmCadAlteracaoEndContribuinte
        Case Else
            MsgBox "Item não configurado."
   End Select

End Sub
Private Sub CarreFormExpedienteCalculos(intIndice As Integer)
Select Case intIndice
        Case 664
            CarregaForm frmCadCalculoIPTU
        Case 1370
            CarregaForm frmLancamentoExecutivosFiscais
        Case 1204
            CarregaForm frmCadLancamentoPrecoPublico
        'Case 665 '(srtCodItem = JDEB) - RETIRADO 28/07/04 Rafael
        '    CarregaForm frmCadCalculoISSQNFixoAnual '(Expediente/Lançamentos/ISSQN Fixo ou Anual)
        Case 3
            ' CarregaForm
        Case 4
            ' CarregaForm
        'Case 670 '(srtCodItem = JDED) - RETIRADO 29/07/04 Rafael
        '    CarregaForm frmCalculoContribuicaoMelhoria '(Expediente/Lançamentos/Contribuição de Melhorias)
        'Case 674 '(srtCodItem = JDEF) - RETIRADO 29/07/04 Rafael
        '    CarregaForm frmCadReceitasDiversas '(Expediente/Lançamentos/Receita Diversas)
        Case 7
        
        Case 1437
            CarregaForm frmLanCobAmig
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormExpedienteInscricaoDA(intIndice As Integer)
Select Case intIndice
        Case 1295
            CarregaForm frmCadDividaAtivaManual
        Case 1296
            CarregaForm frmCadDividaAtivaComposicao
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormExpedienteLancamentosISSQN(intIndice As Integer)
Select Case intIndice
        'Case 667 '(srtCodItem = JDECA) - RETIRADO 29/07/04 Rafael
        '     CarregaForm frmCadCalculoISSQNHomologadoMensal '(Expediente/Lançamentos/ISSQN Variável/ISSQN Mensal ou Homologado)
        'Case 668 '(srtCodItem = JDECB) - RETIRADO 28/07/04 Rafael
        '     CarregaForm frmCadCalculoIssqnArbitrado '(Expediente/Lançamentos/ISSQN Variável/ISSQN Arbitrado)
        'Case 669 '(srtCodItem = JDECC) - RETIRADO 28/07/04 Rafael
        '     CarregaForm frmcadISSQNEstimado '(Expediente/Lançamentos/ISSQN Variável/ISSQN Estimado)
        Case 4
            ' CarregaForm
        Case 5
            ' CarregaForm
        Case 6
            ' CarregaForm
        Case 7
            ' CarregaForm
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormParcelamentos(intIndice As Integer)
    Select Case intIndice
        'Case 643 '(srtCodItem = JDBEA) - RETIRADO 27/07/04 Rafael
        '    CarregaForm frmCadParcelamentoDividaAtiva '(Expediente/Conta Corrente Fiscal/Parcelamentos/Parcelamentos de Débitos)
        Case 644
            'CarregaForm frmCadTrocaProprietarioImoveisITBIUrbanoRural
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormLancamentosITBIUrbanoRural(intIndice As Integer)
Select Case intIndice
        'Case 672 '(srtCodItem = JDEEA) - RETIRADO 29/07/04 Rafael
        '    CarregaForm frmCadCalculoITBIUrbanoRural '(Expediente/Lancamentos/ITBI Urbano e Rural/Cálculo)
        'Case 673 '(srtCodItem = JDEEB) - RETIRADO 29/07/04 Rafael
        '    CarregaForm frmCadTrocaProprietarioImoveisITBIUrbanoRural '(Expediente/Lancamentos/ITBI Urbano e Rural/Troca de Proprietário de Imóveis)
        Case 3
            ' CarregaForm
        Case 4
            ' CarregaForm
        Case 5
            ' CarregaForm
        Case 6
            ' CarregaForm
        Case 7
            ' CarregaForm
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub


Private Sub CarreFormExpedienteBaixas(intIndice As Integer)
Select Case intIndice
    'Case 680 '(srtCodItem = JDFC) - RETIRADO 29/07/04 Rafael
    '    CarregaForm frmCadPagamentos '(Expediente/Controle de Arrecadação/Arrecadação Manual)
    Case 285
        CarregaForm frmArrecadacaoReceita
    Case 1108
        CarregaForm frmResumoBancario
    Case 1110
        CarregaForm frmMovimentoBancario
    Case 1132
        CarregaForm frmProcessamentoBaixa
    Case 1144
        CarregaForm frmBaixaManual
    Case 1154
        CarregaForm frmRecebeMovBancario
    Case 1419
        CarregaForm frmGeraDebitoAutomatico
    Case Else
            MsgBox "Item não configurado."
End Select
End Sub

Private Sub CarreFormExpedienteAtendimentoAoCidadao(intIndice As Integer)
Select Case intIndice
        Case 1
            ' CarregaForm
        Case 2
            ' CarregaForm
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarreFormFerramentas(intIndice As Integer)
    Select Case intIndice
    Case 1
         CarregaForm frmParametroUsuario
    Case 2
         CarregaForm frmCadSenha
    Case 3
         CarregaForm frmAutoNumeracao
    Case 1291
         CarregaForm frmRecebeLanctoExterno
    Case 1310
         CarregaForm frmGeracaoSpool
    Case 1417
         CarregaForm frmGeracaoSpoolNet
    Case Else
         MsgBox "Item não configurado."
    End Select
End Sub

Private Sub CarregaFormTabelasLogradouros(intIndice As Integer)
Select Case intIndice
    Case 53
        CarregaForm frmCadCidade
    Case 581
        CarregaForm frmCadBairro
    Case 582
        CarregaForm frmCadTipoLogradouro
    Case 583
        CarregaForm frmCadTituloLogradouro
    Case 584
        CarregaForm frmCadLogradouro
    Case 585
        CarregaForm frmCadDistritoFiscal
    Case 586
        CarregaForm frmCadSetorFiscal
    Case 587
        CarregaForm frmCadLoteamentos
    Case 1026
        CarregaForm frmCadFaceDeQuadra
    Case 1058
        CarregaForm frmCadTiposDeVias
    Case Else
        MsgBox "Item não configurado."
End Select
End Sub
Private Sub CarregaFormGeraisAgentesArrecadadores(intIndice As Integer)
Select Case intIndice
        Case 590
           CarregaForm frmCadBanco
        Case 591
           CarregaForm frmCadAgenciaBanco
        Case Else
           MsgBox "Item não configurado."
   End Select
End Sub

Private Sub CarregaFormGeraisTabelaDeValores(intIndice As Integer)
Select Case intIndice
        Case 593
            CarregaForm frmCadTabelaDeValores
        Case 594
            CarregaForm frmCadValoresDasFaixas
        Case Else
            MsgBox "Item não configurado."
   End Select

End Sub

Private Sub CarregaFormGeraisReceitaDoMunicipio(intIndice As Integer)
    Select Case intIndice
        Case 444
            CarregaForm frmCadReceita
        Case 445
            CarregaForm frmCadComposicaoDaReceita
        Case Else
            MsgBox "Item não configurado."
   End Select
End Sub
Private Sub CarreFormTabelaGeraisTextos(intIndice As Integer)
Select Case intIndice
        Case 605
            CarregaForm frmCadMensagem
        Case 606
             CarregaForm frmCadTextoLivre
        Case Else
            MsgBox "Item não configurado."
   End Select

End Sub

Private Sub CarregaFormRelCadastroTecnico(intIndex As Integer)
    Select Case intIndex
    Case 686
        frmRelatorioDeOperacaoDoUsuario.Caption = "Conformidade com Inclusão / Alteração / Exclusão"
        CarregaForm frmRelatorioDeOperacaoDoUsuario
    'Case 687 '(srtCodItem = JEDB) - RETIRADO 29/07/04 Rafael
    '    CarregaForm frmRelatorioDeIsencoesImunidadesNaoIncidente '(Relatórios/Cadastro Técnico Municipal/"Beneficiados com Imunidade/Isenção/Não Incidência")
        
    'Case 688 '(srtCodItem = JEDC) - RETIRADO 29/07/04 Rafael
    '    CarregaForm frmCadInconsistenciaImobiliaria '(Relatórios/Cadastro Técnico Municipal/Inconsistências Imobiliárias)
        
    'Case 689 '(srtCodItem = JEDD) - RETIRADO 29/07/04 Rafael
    '    CarregaForm frmRelatorioDeContadoresPorEmpresa '(Relatórios/Cadastro Técnico Municipal/Relação de Contadores por Empresa)
        
    'Case 690 '(srtCodItem = JEDE) - RETIRADO 29/07/04 Rafael
    '    CarregaForm frmRelatorioDeContadoresArrecadacaoPeriodo '(Relatórios/Cadastro Técnico Municipal/Contadores e Arrecadação no Período)
        
    'Case 691 '(srtCodItem = JEDF) - RETIRADO 29/07/04 Rafael
    '    CarregaForm frmRelatorioDeInscritosAtivosInativosBaixados '(Relatórios/Cadastro Técnico Municipal/"Inscritos Ativos/Inativos/Baixados")
        
    'Case 692 '(srtCodItem = JEDG) - RETIRADO 29/07/04 Rafael
    '    CarregaForm frmRelatorioDeContribuintesEmContenciosoAdministrativo '(Relatórios/Cadastro Técnico Municipal/Contribuintes em Contencioso Administrativo)
        
    'Case 693 '(srtCodItem = JEDH) - RETIRADO 30/07/04 Rafael
    '    CarregaForm frmEditaisNotificacaoLancamento '(Relatórios/Cadastro Técnico Municipal/"Editais / Notificações de Lançamento")
            
    'Case 695 '(srtCodItem = JEDJ) - RETIRADO 30/07/04 Rafael
    '    CarregaForm frmRelatorioDemoDeArrecadacaoDeISSQNPorAtividadeEconomica '(Relatórios/Cadastro Técnico Municipal/Demonstrativo de Arrecadação de ISSQN por Atividade Econômica)
        
    'Case 696 '(srtCodItem = JEDK) - RETIRADO 30/07/04 Rafael
    '    CarregaForm frmRelatorioDeQuantidadeDeLancamentosValorTipo '(Relatórios/Cadastro Técnico Municipal/Quantidade de Lançamento, Valor e Tipo)
        
    'Case 698 '(srtCodItem = JEDM) - RETIRADO 30/07/04 Rafael
    '    CarregaForm frmExtratoIndividualizadoDeLancamento '(Relatórios/Cadastro Técnico Municipal/Extrato Individualizado de Lançamento)
    
    'Case 699 'Não existe no banco - RETIRADO 30/07/04 Rafael
    '    CarregaForm frmRelatorioDeParcelasLancadas '(Relatórios/Cadastro Técnico Municipal/Relação das Parcelas Lançadas)
        
    'Case 697 'Não existe no banco - RETIRADO 30/07/04 Rafael
    '    CarregaForm frmRelatorioParcelasArrecadadas '(Relatórios/Cadastro Técnico Municipal/Relação das Parcelas Arrecadadas)
    
    'Case 700 '(srtCodItem = JEDO) - RETIRADO 30/07/04 Rafael
    '    gintCodSeguranca = 700
    '    CarregaForm frmDocumentosDiversos
    
    'Case 682 'Não existe no banco - RETIRADO 30/07/04 Rafael
    '    CarregaForm frmRelatorioDeDiferencaNosPagamentos '(Relatórios/Cadastro Técnico Municipal/Relação de Diferença nos Pagamentos)
        
     End Select
End Sub
 
Private Sub CarregaFormRelCobranca(intIndex As Integer)
    Select Case intIndex
        'Case 774 '(srtCodItem = JEGA) - RETIRADO 03/08/04 Rafael
        '    gintCodSeguranca = 774
        '     CarregaForm frmCadRelacaoDeDocumentosDevolvidos '(Relatórios/Cobrança/Relação de Documentos Devolvidos)
        'Case 775 '(srtCodItem = JEGB) - RETIRADO 03/08/04 Rafael
        '    gintCodSeguranca = 775
        '    CarregaForm frmDocumentosDiversos '(Relatórios/Cobrança/Documentos Diversos)
        Case 787
             CarregaForm frmCadRelacaoDeLancamentosDevolvidos
        Case Else
            MsgBox "Item não configurado."
    End Select
End Sub

Private Sub CarreFormRelBaixaAnaliseReceita(intIndice As Integer)
    Select Case intIndice
        Case 1131
            CarregaForm frmRelDivergencias
        Case 1133
            CarregaForm frmRelBaixas
        Case 1135
            CarregaForm frmRelCriticas
        Case 304
            'RelatorioReceitaArrecadada
            frmRelatorioPeriodo.CarregaFormulario "RR"
        Case 1418
            CarregaForm frmRelMovimentoBancario
        Case Else
            MsgBox "Item não configurado."
    End Select
End Sub


Private Sub actBarra_CustomizeEnd(ByVal bModified As Boolean)
    If bModified Then
        If Dir$(App.Path & "\Data\", vbDirectory) = "" Then MkDir App.Path & "\Data\"
        actBarra.SaveLayoutChanges App.Path & "\Data\" & App.EXEName & "_" & gstrLoginUser & ".chg", ddSOFile
    End If
End Sub

Public Sub actBarra_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 02/05/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim doc As IMDIDocument
    Dim strSql As String
    
    On Error Resume Next
    
    Select Case UCase(Tool.Category)
        Case "SUBCADASTROACORDO"
            Select Case Tool.ID
                Case 1147
                    CarregaForm frmCadAcordos
                Case 1406
                    CarregaForm frmCadCancelamentoAcordo
            End Select
        Case "SUBBANDFICHASCADASTRAIS"
            Select Case Tool.ID
                Case 1152
                    CarregaForm frmFichaCadastroImobiliario
                Case 1358
                    CarregaForm frmFichaCadastroEconomico
            End Select
        Case "SUBBANDFICHASLANCAMENTOS"
            Select Case Tool.ID
                Case 1222
                   CarregaForm frmFichaLancamentoImobiliario
                Case 1302
                    CarregaForm frmFichaLancamentoISSConstrucao
            End Select
        Case "SUBBANDISS"
            Select Case Tool.ID
                Case 1382
                    CarregaForm frmRelNotasFiscais
            End Select
        Case "SUBBANDPAGAMENTOS"
            Select Case Tool.ID
                Case 1267
                   CarregaForm frmRelPagtoPorAviso
                Case 1259
                   CarregaForm frmRelatorioPagamentos
            End Select
        Case "SUBSALDODIVIDAATIVA"
            Select Case Tool.ID
                Case 1427
                    If bytDBType = EDatabases.Oracle Then
                        strSql = "SELECT Sum(LV.dblvalor) dblValor, CR.intUtilizacao,"
                        strSql = strSql & "LA.intcomposicaodareceita, "
                        strSql = strSql & "LA.Intexercicio, "
                        strSql = strSql & "CR.strDescricao strComposicaoDaReceita "
                        strSql = strSql & "FROM " & gstrLancamentoValor & " LV, "
                        strSql = strSql & gstrLancamentoPagamento & " LP, "
                        strSql = strSql & gstrLancamentoAlfa & " LA, "
                        strSql = strSql & gstrDativa & " DA, "
                        strSql = strSql & gstrComposicaoDaReceita & " CR "
                        strSql = strSql & "WHERE LV.bitParcelaValida = 1 "
                        strSql = strSql & "AND (LV.pkid = LP.INTLANCAMENTOVALOR " & strOUTJOracle & " "
                        strSql = strSql & "AND LP.PKID Is Null) "
                        strSql = strSql & "AND LA.pkid = LV.INTLANCAMENTOALFA "
                        strSql = strSql & "AND LA.pkid = DA.INTLANCAMENTOALFA "
                        strSql = strSql & "AND LV.intLancamentoAlfaAcordo Is Null "
                        strSql = strSql & "AND CR.pkid " & strOUTJOracle & " = LA.intComposicaoDaReceita "
                        strSql = strSql & "AND CR.INTUTILIZACAO = 4 "
                        strSql = strSql & "GROUP BY LA.Intexercicio, LA.intComposicaoDaReceita, CR.strdescricao, CR.intUtilizacao "
                        strSql = strSql & "UNION ALL "
                        strSql = strSql & "SELECT Sum(LV.dblvalor) dblValor, CR.intUtilizacao, "
                        strSql = strSql & "LA.intcomposicaodareceita, "
                        strSql = strSql & "LA.Intexercicio, "
                        strSql = strSql & "CR.strDescricao strComposicaoDaReceita "
                        strSql = strSql & "FROM " & gstrLancamentoValor & " LV, "
                        strSql = strSql & gstrLancamentoPagamento & " LP, "
                        strSql = strSql & gstrLancamentoAlfa & " LA, "
                        strSql = strSql & gstrDativa & " DA, "
                        strSql = strSql & gstrComposicaoDaReceita & " CR "
                        strSql = strSql & "WHERE LV.bitParcelaValida = 1 "
                        strSql = strSql & "AND (LV.pkid = LP.INTLANCAMENTOVALOR " & strOUTJOracle & " "
                        strSql = strSql & "AND LP.PKID Is Null) "
                        strSql = strSql & "AND LA.pkid = LV.INTLANCAMENTOALFA "
                        strSql = strSql & "AND LA.pkid = DA.INTLANCAMENTOALFA "
                        strSql = strSql & "AND LV.intLancamentoAlfaAcordo Is Null "
                        strSql = strSql & "AND CR.pkid " & strOUTJOracle & " = LA.intComposicaoDaReceita "
                        strSql = strSql & "AND CR.INTUTILIZACAO <> 4 "
                        strSql = strSql & "AND LA.intExercicio < " & Year(gstrDataDoSistema) & " "
                        strSql = strSql & "AND LA.bytNaoInscreveda = 0 "
                        strSql = strSql & "GROUP BY LA.Intexercicio, LA.intComposicaoDaReceita, CR.strdescricao, CR.intUtilizacao "
                    ElseIf bytDBType = EDatabases.SQLServer Then
                        strSql = "SELECT SUM(LV.dblValor) AS dblValor, "
                        strSql = strSql & "CR.intUtilizacao, "
                        strSql = strSql & "LA.INTCOMPOSICAODARECEITA, "
                        strSql = strSql & "LA.intExercicio, "
                        strSql = strSql & "CR.strDescricao AS strComposicaoDaReceita "
                        strSql = strSql & "FROM " & gstrLancamentoValor & " LV " & strREADPAST
                        strSql = strSql & "INNER JOIN " & gstrLancamentoAlfa & " LA " & strREADPAST & " ON LV.intLancamentoAlfa = LA.PKId "
                        strSql = strSql & "INNER JOIN " & gstrDativa & " DA " & strREADPAST & " ON DA.intLancamentoAlfa = LA.PKId "
                        strSql = strSql & "LEFT OUTER JOIN " & gstrLancamentoPagamento & " LP " & strREADPAST & " ON LV.PKId = LP.INTLANCAMENTOVALOR "
                        strSql = strSql & "LEFT OUTER JOIN " & gstrComposicaoDaReceita & " CR " & strREADPAST & " ON LA.INTCOMPOSICAODARECEITA = CR.PKId "
                        strSql = strSql & "WHERE (LV.bitParcelaValida = 1) "
                        strSql = strSql & "AND (LP.PKID IS NULL) "
                        strSql = strSql & "AND (LV.intLancamentoAlfaAcordo IS NULL) "
                        strSql = strSql & "AND (CR.intUtilizacao = 4) "
                        strSql = strSql & "GROUP BY LA.intExercicio, "
                        strSql = strSql & "LA.INTCOMPOSICAODARECEITA, "
                        strSql = strSql & "CR.strDescricao, "
                        strSql = strSql & "CR.intUtilizacao "
                        strSql = strSql & "UNION ALL "
                        strSql = strSql & "SELECT SUM(LV.dblValor) AS dblValor, "
                        strSql = strSql & "CR.intUtilizacao, "
                        strSql = strSql & "LA.INTCOMPOSICAODARECEITA, "
                        strSql = strSql & "LA.intExercicio, "
                        strSql = strSql & "CR.strDescricao AS strComposicaoDaReceita "
                        strSql = strSql & "FROM " & gstrLancamentoValor & " LV " & strREADPAST
                        strSql = strSql & "INNER JOIN " & gstrLancamentoAlfa & " LA " & strREADPAST & " ON LV.intLancamentoAlfa = LA.PKId "
                        strSql = strSql & "INNER JOIN " & gstrDativa & " DA " & strREADPAST & " ON DA.intLancamentoAlfa = LA.PKId "
                        strSql = strSql & "LEFT OUTER JOIN " & gstrLancamentoPagamento & " LP " & strREADPAST & " ON LV.PKId = LP.INTLANCAMENTOVALOR "
                        strSql = strSql & "LEFT OUTER JOIN " & gstrComposicaoDaReceita & " CR " & strREADPAST & " ON LA.INTCOMPOSICAODARECEITA = CR.PKId "
                        strSql = strSql & "WHERE (LV.bitParcelaValida = 1) "
                        strSql = strSql & "AND (LP.PKID IS NULL) "
                        strSql = strSql & "AND (LV.intLancamentoAlfaAcordo IS NULL) "
                        strSql = strSql & "AND (CR.intUtilizacao <> 4) "
                        strSql = strSql & "AND (LA.intExercicio < " & Year(gstrDataDoSistema) & ") "
                        strSql = strSql & "AND (LA.BYTNAOINSCREVEDA = 0) "
                        strSql = strSql & "GROUP BY LA.intExercicio, "
                        strSql = strSql & "LA.INTCOMPOSICAODARECEITA, "
                        strSql = strSql & "CR.strDescricao, "
                        strSql = strSql & "CR.intUtilizacao "
                        strSql = strSql & "ORDER BY LA.intExercicio, "
                        strSql = strSql & "CR.strDescricao "
                    End If
                    ImprimeRelatorio rptRelSaldoDividaAtiva, strSql, "Saldo de Dívida Ativa", 300
                
                Case 1428
                    CarregaForm frmRelSaldoDividaAtivaPeriodo
                    
            End Select
        Case "PRINCIPAL"
            FormsPrincipal Tool.ID
        'Nino
        Case "DOCUMENTOS"
            FormsDocumentos Tool.ID
        Case "WORDTEMPLATE"
            OpenWordTemplate Tool.TagVariant
        
        Case "SUBCONSULTASCALCULO"
            'CarregaForm frmGeradorRelatorio
        Case "SUBTABELASGERAIS"
            CarregaFormTabelasGerais Tool.ID
        Case "SUBTABELAGERAISAGENTESARRECADADORES"
            CarregaFormGeraisAgentesArrecadadores Tool.ID
        Case "SUBTABELAGERAISRECEITADOMUNICIPIO"
            CarregaFormGeraisReceitaDoMunicipio Tool.ID
        Case "SUBTABELAGERAISPLANTASDEVALORES"
            CarregaFormGeraisTabelaDeValores Tool.ID
        Case "SUBTABELASLOGRADOUROS"
            CarregaFormTabelasLogradouros Tool.ID
        Case "SUBTABELASIMOBILIÁRIASURBANAS"
            CarreFormTabelasImobiliariasUrbanas Tool.ID
        Case "SUBTABELASECONOMICAS"
            CarreFormTabelasEconomicas Tool.ID
        Case "SUBTABELASCONTRIBUICAODEMELHORIAS"
            CarreFormTabelasContribuicaoDeMelhorias Tool.ID
        Case "SUBTABELAGERAISTEXTOS"
            CarreFormTabelaGeraisTextos Tool.ID
        Case "RELCADASTROTECNICO"
            CarregaFormRelCadastroTecnico Tool.ID
        Case "RELCOBRANCA"
            CarregaFormRelCobranca Tool.ID
        Case "CADASTROS"
            CarreFormCadastros Tool.ID
        Case "SUBCADASTROLANCAMENTOS"
            CarreFormCadastros Tool.ID
        Case "SUBCADASTROPARAMETROS"
              CarreFormCadastros Tool.ID
        Case "RELATORIOS"
            Select Case Tool.ID
                Case 10
                    strSql = ""
                    strSql = strSql & " SELECT C.PKId, C.strNome, C.bytNaturezaJuridica AS Natureza, C.dtmDataCadastro,"
                    'strSql = strSql & " RTRIM(CONVERT(NVARCHAR, L.strCodigo)) + ' - ' +"
                    strSql = strSql & " RTRIM(" & gstrCONVERT(CDT_NVARCHAR, "L.strCodigo") & ") " & strCONCAT & " ' - ' " & strCONCAT
                    'strSql = strSql & " RTRIM(LTRIM(ISNULL(TL.strSigla, '') + ' ' + ISNULL(U.strDescricao,'') + ' ' +"
                    strSql = strSql & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & strCONCAT & " ' ' " & strCONCAT & gstrISNULL("U.strDescricao", "''") & strCONCAT & " ' ' " & strCONCAT
                    'strSql = strSql & " L.strDescricao)) + ', ' + CONVERT(NVARCHAR,ISNULL(intNumero,0)) +"
                    strSql = strSql & " L.strDescricao)) " & strCONCAT & " ', ' " & strCONCAT & gstrCONVERT(CDT_NVARCHAR, gstrISNULL("intNumero", "0")) & strCONCAT
                    'strSql = strSql & " CASE CONVERT(NVARCHAR,ISNULL(strComplemento,0)) WHEN '0' THEN ' - Bairro: ' + strBairroC ELSE '/' +"
                    'strSQL = strSQL & " strComplemento + ' - Bairro: ' + strBairroC end AS Logradouro"
                    strSql = strSql & gstrCASEWHEN(gstrCONVERT(CDT_NVARCHAR, gstrISNULL("strComplemento", "0")), "'0', ' - Bairro: ' " & strCONCAT & " strBairroC", "'/'" & strCONCAT & _
                                    " strComplemento " & strCONCAT & " ' - Bairro: ' " & strCONCAT & " strBairroC") & " AS Logradouro"
                    strSql = strSql & " FROM tblLogradouro L, tblTituloLogradouro U, tblTipoLogradouro TL, tblContribuinte C"
                    strSql = strSql & " WHERE "
                    'strSql = strSql & " L.intTituloLogradouro *= U.PKId"
                    strSql = strSql & " L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId" & strOUTJOracle
                    'strSql = strSql & " AND L.intTipoLogradouro *= TL.PKId"
                    strSql = strSql & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId" & strOUTJOracle
                    strSql = strSql & " AND C.intLogradouro = L.PKId"
                    'strSql = strSql & " AND C.intTipoLogradouro *= TL.PKId"
                    strSql = strSql & " AND C.intTipoLogradouro " & strOUTJOracle & strOUTJSQLServer & "= TL.PKId"
                    strSql = strSql & " ORDER BY C.strNome"
                    ImprimeRelatorio rptCadContribuinteDuplicados, strSql
                Case 1169
                    CarregaForm frmRelInscricaoQuadraSetor
                Case 1174
                    CarregaForm frmRelatorioRolLogradouro
                Case 1177
                    ' RESPONSÁVEL LEANDRO
                    strSql = ""
                    strSql = strSql & "SELECT "
                    strSql = strSql & "GA.Strnomedogrupo Grupo, "
                    strSql = strSql & "SGA.STRNOMEDOSUBGRUPO SubGrupo, "
                    strSql = strSql & "AE.strDescricao Descricao "
                    strSql = strSql & "FROM "
                    strSql = strSql & gstrAtividadeEC & " AE, "
                    strSql = strSql & gstrGrupoDeAtividade & " GA, "
                    strSql = strSql & gstrSubGrupoDeAtividade & " SGA "
                    strSql = strSql & "WHERE "
                    strSql = strSql & "AE.intgrupo = GA.pkid and "
                    strSql = strSql & "AE.Intsubgrupo = SGA.Pkid "
                    strSql = strSql & "ORDER BY "
                    strSql = strSql & "GA.strnomedogrupo, "
                    strSql = strSql & "SGA.STRNOMEDOSUBGRUPO, "
                    strSql = strSql & "AE.strDescricao"
                    ImprimeRelatorio rptRolAtividades, strSql
                Case 1178
                    CarregaForm frmQtdContribuintesPorAtividade
                Case 1181
                    CarregaForm frmAtividadeContribuintePorLogradouro
                Case 1195
                    ' RESPONSÁVEL LEANDRO   02/07/2004
                   strSql = ""
                   strSql = strSql & "SELECT"
                   strSql = strSql & " TT.PKID ID,"
                   strSql = strSql & " TT.Strdescricao DescricaoTipo,"
                   strSql = strSql & " TE.INTEXERCICIO Exercicio,"
                   strSql = strSql & " TR.Intcodigotributo CodigoTributo, "
                   strSql = strSql & " TR.STRDESCRICAO DescricaoTributo,"
                   strSql = strSql & " TE.dblValor Valor"
                   strSql = strSql & " FROM "
                   strSql = strSql & gstrTributoTipo & " TT, "
                   strSql = strSql & gstrTributo & " TR, "
                   strSql = strSql & gstrTributoExercicio & " TE"
                   strSql = strSql & " WHERE "
                   strSql = strSql & " TT.pkid = TR.Inttributotipo AND"
                   strSql = strSql & " TR.Pkid = TE.intTributo"
                   strSql = strSql & " ORDER BY"
                   strSql = strSql & " TT.strDescricao, TR.Intcodigotributo"
                   ImprimeRelatorio rptTaxasDeLicenca, strSql
                Case 1196
                   CarregaForm frmContribuintesPorAtividades
                Case 1228
                   CarregaForm frmLivroDividaAtiva
                Case 1392
                   CarregaForm frmRelPosicaoLancamentos
                Case 1395
                   CarregaForm frmRelRecComposicaoPago
                Case 1239
''                    strsql = "SELECT sum(LV.dblvalor) dblValor, LA.intcomposicaodareceita, LA.Intexercicio, CR.strdescricao strComposicaoDaReceita " & _
''                             "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoPagamento & " LP, " & gstrLancamentoAlfa & " LA, " & gstrComposicaoDaReceita & " CR " & _
''                             "WHERE LV.bitParcelaValida = 1 AND " & _
''                             "(LV.pkid " & strOUTJSQLServer & "= LP.INTLANCAMENTOVALOR " & strOUTJOracle & " AND LP.PKID Is Null) AND LA.pkid = LV.INTLANCAMENTOALFA AND " & _
''                             "LV.intLancamentoAlfaAcordo Is Null AND CR.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LA.intComposicaoDaReceita AND " & _
''                             "(CR.INTUTILIZACAO = 4 OR (CR.INTUTILIZACAO <> 4 AND LA.intExercicio < " & Year(gstrDataDoSistema) & " AND LA.bytNaoInscreveda = 0)) " & _
''                             "GROUP BY LA.intComposicaoDaReceita, LA.Intexercicio, CR.strdescricao"
                    
'                    If bytDBType = EDatabases.Oracle Then
'                        strSQL = "SELECT Sum(LV.dblvalor) dblValor, CR.intUtilizacao,"
'                        strSQL = strSQL & "LA.intcomposicaodareceita, "
'                        strSQL = strSQL & "LA.Intexercicio, "
'                        strSQL = strSQL & "CR.strDescricao strComposicaoDaReceita "
'                        strSQL = strSQL & "FROM " & gstrLancamentoValor & " LV, "
'                        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
'                        strSQL = strSQL & gstrLancamentoAlfa & " LA, "
'                        strSQL = strSQL & gstrComposicaoDaReceita & " CR "
'                        strSQL = strSQL & "WHERE LV.bitParcelaValida = 1 "
'                        strSQL = strSQL & "AND (LV.pkid = LP.INTLANCAMENTOVALOR " & strOUTJOracle & " "
'                        strSQL = strSQL & "AND LP.PKID Is Null) "
'                        strSQL = strSQL & "AND LA.pkid = LV.INTLANCAMENTOALFA "
'                        strSQL = strSQL & "AND LV.intLancamentoAlfaAcordo Is Null "
'                        strSQL = strSQL & "AND CR.pkid " & strOUTJOracle & " = LA.intComposicaoDaReceita "
'                        strSQL = strSQL & "AND CR.INTUTILIZACAO = 4 "
'                        strSQL = strSQL & "GROUP BY LA.Intexercicio, LA.intComposicaoDaReceita, CR.strdescricao, CR.intUtilizacao "
'                        strSQL = strSQL & "UNION ALL "
'                        strSQL = strSQL & "SELECT Sum(LV.dblvalor) dblValor, CR.intUtilizacao, "
'                        strSQL = strSQL & "LA.intcomposicaodareceita, "
'                        strSQL = strSQL & "LA.Intexercicio, "
'                        strSQL = strSQL & "CR.strDescricao strComposicaoDaReceita "
'                        strSQL = strSQL & "FROM " & gstrLancamentoValor & " LV, "
'                        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
'                        strSQL = strSQL & gstrLancamentoAlfa & " LA, "
'                        strSQL = strSQL & gstrComposicaoDaReceita & " CR "
'                        strSQL = strSQL & "WHERE LV.bitParcelaValida = 1 "
'                        strSQL = strSQL & "AND (LV.pkid = LP.INTLANCAMENTOVALOR " & strOUTJOracle & " "
'                        strSQL = strSQL & "AND LP.PKID Is Null) "
'                        strSQL = strSQL & "AND LA.pkid = LV.INTLANCAMENTOALFA "
'                        strSQL = strSQL & "AND LV.intLancamentoAlfaAcordo Is Null "
'                        strSQL = strSQL & "AND CR.pkid " & strOUTJOracle & " = LA.intComposicaoDaReceita "
'                        strSQL = strSQL & "AND CR.INTUTILIZACAO <> 4 "
'                        strSQL = strSQL & "AND LA.intExercicio < " & Year(gstrDataDoSistema) & " "
'                        strSQL = strSQL & "AND LA.bytNaoInscreveda = 0 "
'                        strSQL = strSQL & "GROUP BY LA.Intexercicio, LA.intComposicaoDaReceita, CR.strdescricao, CR.intUtilizacao "
'                    ElseIf bytDBType = EDatabases.SQLServer Then
'                        strSQL = "SELECT SUM(LV.dblValor) AS dblValor, "
'                        strSQL = strSQL & "CR.intUtilizacao, "
'                        strSQL = strSQL & "LA.INTCOMPOSICAODARECEITA, "
'                        strSQL = strSQL & "LA.intExercicio, "
'                        strSQL = strSQL & "CR.strDescricao AS strComposicaoDaReceita "
'                        strSQL = strSQL & "FROM " & gstrLancamentoValor & " LV " & strREADPAST
'                        strSQL = strSQL & "INNER JOIN " & gstrLancamentoAlfa & " LA " & strREADPAST & " ON LV.intLancamentoAlfa = LA.PKId "
'                        strSQL = strSQL & "LEFT OUTER JOIN " & gstrLancamentoPagamento & " LP " & strREADPAST & " ON LV.PKId = LP.INTLANCAMENTOVALOR "
'                        strSQL = strSQL & "LEFT OUTER JOIN " & gstrComposicaoDaReceita & " CR " & strREADPAST & " ON LA.INTCOMPOSICAODARECEITA = CR.PKId "
'                        strSQL = strSQL & "WHERE (LV.bitParcelaValida = 1) "
'                        strSQL = strSQL & "AND (LP.PKID IS NULL) "
'                        strSQL = strSQL & "AND (LV.intLancamentoAlfaAcordo IS NULL) "
'                        strSQL = strSQL & "AND (CR.intUtilizacao = 4) "
'                        strSQL = strSQL & "GROUP BY LA.intExercicio, "
'                        strSQL = strSQL & "LA.INTCOMPOSICAODARECEITA, "
'                        strSQL = strSQL & "CR.strDescricao, "
'                        strSQL = strSQL & "CR.intUtilizacao "
'                        strSQL = strSQL & "UNION ALL "
'                        strSQL = strSQL & "SELECT SUM(LV.dblValor) AS dblValor, "
'                        strSQL = strSQL & "CR.intUtilizacao, "
'                        strSQL = strSQL & "LA.INTCOMPOSICAODARECEITA, "
'                        strSQL = strSQL & "LA.intExercicio, "
'                        strSQL = strSQL & "CR.strDescricao AS strComposicaoDaReceita "
'                        strSQL = strSQL & "FROM " & gstrLancamentoValor & " LV " & strREADPAST
'                        strSQL = strSQL & "INNER JOIN " & gstrLancamentoAlfa & " LA " & strREADPAST & " ON LV.intLancamentoAlfa = LA.PKId "
'                        strSQL = strSQL & "LEFT OUTER JOIN " & gstrLancamentoPagamento & " LP " & strREADPAST & " ON LV.PKId = LP.INTLANCAMENTOVALOR "
'                        strSQL = strSQL & "LEFT OUTER JOIN " & gstrComposicaoDaReceita & " CR " & strREADPAST & " ON LA.INTCOMPOSICAODARECEITA = CR.PKId "
'                        strSQL = strSQL & "WHERE (LV.bitParcelaValida = 1) "
'                        strSQL = strSQL & "AND (LP.PKID IS NULL) "
'                        strSQL = strSQL & "AND (LV.intLancamentoAlfaAcordo IS NULL) "
'                        strSQL = strSQL & "AND (CR.intUtilizacao <> 4) "
'                        strSQL = strSQL & "AND (LA.intExercicio < " & Year(gstrDataDoSistema) & ") "
'                        strSQL = strSQL & "AND (LA.BYTNAOINSCREVEDA = 0) "
'                        strSQL = strSQL & "GROUP BY LA.intExercicio, "
'                        strSQL = strSQL & "LA.INTCOMPOSICAODARECEITA, "
'                        strSQL = strSQL & "CR.strDescricao, "
'                        strSQL = strSQL & "CR.intUtilizacao "
'                        strSQL = strSQL & "ORDER BY LA.intExercicio, "
'                        strSQL = strSQL & "CR.strDescricao "
'                    End If
'
'
'                    ImprimeRelatorio rptRelSaldoDividaAtiva, strSQL, "Saldo de Dívida Ativa", 300
                   ''CarregaForm frmRelSaldoDividaAtiva
                Case 1274
                   CarregaForm frmRelatorioDeIsencaoImunidade
                Case 1441
                   CarregaForm frmRelDevedoresFaixaValores
                Case 1359
                    CarregaForm frmOcorrenciasDoEconomico
                Case 1361
                    CarregaForm frmRelAlteracaoEconomica
                Case Else
                   ExibeMensagem "Item não configurado."
            End Select
        Case "SUBSEGUNDASVIAS"
            Select Case Tool.ID
                Case 1203
                   CarregaForm frmAcordoCarneSegundaVia
                Case 1408
                   CarregaForm frmAcordoCarne2ViaAtualizada
                Case 1243
                   CarregaForm frmCarneISSConstrucao
                Case 1197
                   CarregaForm frmIPTUCarneSegundaVia
                Case 1240
                   CarregaForm frmISSCarneSegundaVia
                Case 1268
                   CarregaForm frmCarnePrecoPublico
                Case 1369
                    CarregaForm frmISSVarCarneSegundaVia
                Case Else
                   ExibeMensagem "Item não configurado."
            End Select
        Case "SUBEXECUTIVOSFISCAIS"
            Select Case Tool.ID
                Case 1389
                    frmDocPeticao.strOpcao = "PET"
                    frmDocPeticao.Caption = "Petição"
                    frmDocPeticao.tab_3dPasta.Caption = "Petição"
                    CarregaForm frmDocPeticao
                Case 1394
                    frmDocPeticao.strOpcao = "CDA"
                    frmDocPeticao.Caption = "Certidão de Dívida Ativa"
                    frmDocPeticao.tab_3dPasta.Caption = "Certidão de Dívida Ativa"
                    CarregaForm frmDocPeticao
                Case Else
                   ExibeMensagem "Item não configurado."
            End Select
        Case "SUBRELLANCAMENTOS"
            Select Case Tool.ID
                Case 1266
                    CarregaForm frmRelTotaisTPTU
                Case 1284
                    CarregaForm frmRelComparativoIPTU
                Case Else
                   ExibeMensagem "Item não configurado."
            End Select
        Case "SUBROLLANCAMENTOS"
            Select Case Tool.ID
                Case 1431
                    CarregaForm frmRelatorioRolLanctoEconomico
            End Select
        Case "SUBEXPEDIENTEADMINISTRACAO"
            CarreFormExpedienteAdministracao Tool.ID
        Case "SUBEXPEDIENTEFINANCEIRO"
            CarreFormExpedienteFinanceiro Tool.ID
        Case "SUBEXPEDIENTEFINANCEIROTRANSFERENCIASPARADIVIDAATIVA"
            CarreFormTransferenciaParaDividaAtiva Tool.ID
        Case "SUBEXPEDIENTEFISCALIZACAO"
            CarreFormExpedienteFicalizacao Tool.ID
        Case "SUBEXPEDIENTECONTENCIOSOJUDICIAL"
            CarreFormExpedienteContenciosoJudicial Tool.ID
        Case "SUBEXPEDIENTECONTENCIOSOADMINISTRATIVO"
            CarreFormExpedienteContenciosoAdministrativo Tool.ID
        Case "SUBEXPEDIENTECONTENCIOSO"
            CarreFormExpedienteContencioso Tool.ID
        Case "EXPEDIENTE"
            CarreFormExpediente Tool.ID
        Case "SUBEXPEDIENTEINSCRICAODA"
            CarreFormExpedienteInscricaoDA Tool.ID
        Case "SUBEXPEDIENTECALCULOS"
            CarreFormExpedienteCalculos Tool.ID
        Case "SUBEXPEDIENTEBAIXAS"
            CarreFormExpedienteBaixas Tool.ID
        Case "SUBEXPEDIENTEATENDIMENTOAOCIDADAO"
            CarreFormExpedienteAtendimentoAoCidadao Tool.ID
        Case "SUBEXPEDIENTECALCULOSISSQNVARIAVEL"
            CarreFormExpedienteLancamentosISSQN Tool.ID
        Case "SUBEXPEDIENTECALCULOSITBIURBANOER fURAL"
            CarreFormLancamentosITBIUrbanoRural Tool.ID
        Case "SUBEXPEDIENTEFINANCEIROPARCELAMENTOS"
            CarreFormParcelamentos Tool.ID
        Case "SUBALTERACAOENDNOTIFICACAO"
            Select Case Tool.ID
                Case 1261
                    CarregaForm frmAlteracaoEndImobiliario
                Case 1262
                    CarregaForm frmCadAlteracaoEndContribuinte
                Case Else
                    ExibeMensagem "Item não configurado."
            End Select
        Case "CONBRANCABANCARIA"
            CarregaFormCobrancaBancaria Tool.ID
        Case "SUBFERRAMENTAS"
            CarreFormFerramentas Tool.ID
        Case "SUBARQUIVOSGRAFICA"
            Select Case Tool.ID
                Case 1439
                    CarregaForm frmGeracaoSpoolIPTU
                Case 1440
                    CarregaForm frmGeracaoSpoolCobAmigavel
            End Select
        Case UCase("RelFiscalizacaoContencioso")
            CarreFormRelFiscalizacaoContencioso Tool.ID
        Case UCase("RelContaCorrenteFiscal")
            CarreFormRelContaCorrenteFiscal Tool.ID
        Case UCase("RelControleArrecadacao")
            CarreFormRelControleArrecadacao Tool.ID
        Case UCase("RelDividaAtiva")
            CarreFormRelDividaAtiva Tool.ID
        Case UCase("SUBEXPEDIENTECALCULOSITBIURBANOERURAL")
            CarreFormcadCalculoITBIUrbanoRural Tool.ID
        Case UCase("SUBRELBAIXAANALISERECEITA")
            CarreFormRelBaixaAnaliseReceita Tool.ID
        Case UCase("SUBCADASTROFISCALIZACAOISS")
            CarregaForm frmIssNotaFiscal
        Case "TABELAS"
            Select Case Tool.ID
                Case 1
                    ' FLAG
                CarregaForm frmCadCampoDeInscricao
            End Select
            
        Case "CADASTROS"
            Select Case Tool.ID
                Case 1
            End Select
    
        Case "TEXTOS"
            Texto_ToolClick Tool
            
        Case gstrMnuArquivo, gstrBtnArquivo
            If Tool.Name = gstrSair Then
                FinalizaSistema
                Exit Sub
            End If
            If Not ActiveForm Is Nothing Then
                If Tool.Name = gstrFechar Then
                    Unload ActiveForm
                Else
                    ActiveForm.MantemForm Tool.Name
                End If
            End If
            
        Case "1EDIT", "EDIT"
            If ActiveForm Is Nothing Then
                Exit Sub
            ElseIf blnTexto And TypeOf ActiveForm Is IMDIDocument Then
                Set doc = New IMDIDocument
                doc.CommandHandler Tool
                frmMDIDoc.IMDIDocument_CommandHandler Tool
                Exit Sub
            End If
        
            Select Case Tool.ID
                Case 4013
                    SendKeys "^{X}"
                Case 4014
                    SendKeys "^{C}"
                Case 4015
                    SendKeys "^{V}"
            End Select
        Case "VIEW", "1VIEW"
            Select Case Tool.Name
                Case "1miVToolbar", "miVToolbar"
                    Select Case Tool.Checked
                        Case True
                            Tool.Checked = False
                            actBarra.Bands("bndFormulario").Visible = False
                            actBarra.RecalcLayout
                        Case False
                            Tool.Checked = True
                            actBarra.Bands("bndFormulario").Visible = True
                            actBarra.RecalcLayout
                    End Select

                Case "1miVStatusBar", "miVStatusBar"
                    
                    Select Case Tool.Checked
                        Case True
                            Tool.Checked = False
                            staBarraStatus.Visible = False
                            actBarra.RecalcLayout
                        Case False
                            Tool.Checked = True
                            staBarraStatus.Visible = True
                            actBarra.RecalcLayout
                    End Select
            End Select
        Case "1HELP", "HELP"
            Select Case Tool.ID
                Case 4019
                    If ActiveForm Is Nothing Then
                        Call_HtmlHelp 1
                        Exit Sub
                    Else
                        If TypeOf ActiveForm Is Form Then
                            Call_HtmlHelp ActiveForm.HelpContextID
                        ElseIf TypeOf ActiveForm Is ActiveReport Then
                            Call_HtmlHelp ActiveForm.Tag
                        End If
                    End If
                    
                Case 4020
                    ShellEx "http://www.cpdsystems.com.br", , , , , Me.hWnd
            End Select
        Case "JANELAS"
            Select Case Tool.Name
                Case "miVertical"
                    MDIMenu.Arrange vbTileVertical
                Case "miHorizontal"
                    MDIMenu.Arrange vbTileHorizontal
                Case "miCascata"
                    MDIMenu.Arrange vbCascade
            End Select

    End Select
End Sub


Private Sub CarreFormRelDividaAtiva(intIndice As Integer)
    Select Case intIndice
        Case 782
        '    CarregaForm
        'Case 783 '(srtCodItem = JEIB) - RETIRADO 03/08/04 Rafael
        '    CarregaForm frmRelatorioLivroDaDividaAtiva '(Relatórios/Dívida Ativa/Livro da Dívida Ativa)
        'Case 784 '(srtCodItem = JEIC) - RETIRADO 04/08/04 Rafael
        '    CarregaForm frmRelatorioRelacaoDeAdimplenciaEmDividaAtiva '(Relatórios/Dívida Ativa/Relação de Adimplência em Dívida Ativa)
        'Case 785 '(srtCodItem = JEID) - RETIRADO 04/08/04 Rafael
        '    CarregaForm frmRelatorioRelacaoDeInadimplenciaEmDividaAtiva '(Relatórios/Dívida Ativa/Relação de Inadimplência em Dívida Ativa)
    End Select
End Sub
Private Sub CarreFormcadCalculoITBIUrbanoRural(intIndice As Integer)
    Select Case intIndice
        'Case 672 '(srtCodItem = JDEEA) - RETIRADO 29/07/04 Rafael
        '    CarregaForm frmCadCalculoITBIUrbanoRural '(Expediente/Lancamentos/ITBI Urbano e Rural/Cálculo)
    End Select
End Sub


Private Sub CarreFormRelControleArrecadacao(intIndice As Integer)
    Select Case intIndice
        'Case 765 '(srtCodItem = JEED) - RETIRADO 30/07/04 Rafael
        '    gintCodSeguranca = 765
        '    CarregaForm frmDocumentosDiversos '(Relatórios/Controle de Arrecadação/Documentos Diversos)
    End Select
End Sub

Private Sub CarreFormRelContaCorrenteFiscal(intIndice As Integer)
    Select Case intIndice
        'Case 768 'Não existe no banco - RETIRADO 03/08/04 Rafael
        '    CarregaForm frmRelatoriodeInadimplenciaAnalitico '(Relatórios/Conta Corrente Fiscal/Relação de Inadimplência Analítico)
        'Case 769 'Não existe no banco - RETIRADO 03/08/04 Rafael
        '    CarregaForm frmRelatorioParaControleDeInadimplencia
        'Case 771 '(srtCodItem = JEFE) - RETIRADO 03/08/04 Rafael
        '    CarregaForm frmRelatorioDeDebitosNaoInscritosDV
        'Case 772 '(srtCodItem = JEFF) - RETIRADO 03/08/04 Rafael
        '    gintCodSeguranca = 772
        '    CarregaForm frmDocumentosDiversos '(Relatórios/Conta Corrente Fiscal/Documentos Diversos)
            
    End Select
End Sub

Private Sub CarreFormRelFiscalizacaoContencioso(intIndice As Integer)
    Select Case intIndice
        'Case 780 '(srtCodItem = JEHD) - RETIRADO 03/08/04 Rafael
        '    gintCodSeguranca = 780
        '    CarregaForm frmDocumentosDiversos '(Relatórios/"Fiscalização / Contencioso"/Documentos Diversos)
        'Case 777 '(srtCodItem = JEHA) - RETIRADO 03/08/04 Rafael
        '    CarregaForm frmPosicaoDeAlvaras '(Relatórios/"Fiscalização / Contencioso"/Posição de Alvarás)
        
    End Select
End Sub

Private Sub CarregaFormCobrancaBancaria(intIndice As Integer)
Select Case intIndice
    'Case 677 '(srtCodItem = JDFAA) - RETIRADO 29/07/04 Rafael
    '    CarregaForm frmCadLayOutBaixa '(Expediente/Controle de Arrecadação/Cobrança Bancária/Confecção de Layout)
    'Case 678 '(srtCodItem = JDFAB) - RETIRADO 29/07/04 Rafael
    '    CarregaForm frmBaixaAutomatica '(Expediente/Controle de Arrecadação/Cobrança Bancária/Baixa Automática)
End Select
End Sub

Private Sub itmImprimir_Click()
    GridDeImpressao.PrintInfo.PrintPreview
End Sub

Private Sub MDIForm_Activate()
    On Error Resume Next
    If gblnFlagDicas Then
        If gblnMostraDicas = True Then
            frmDicas.Show 1
        End If
    End If
    gblnFlagDicas = False
    staBarraStatus.Panels(1).Width = Me.Width - staBarraStatus.Panels(2).Width - staBarraStatus.Panels(3).Width - staBarraStatus.Panels(4).Width
End Sub

Private Sub MDIForm_Load()
    Me.Caption = App.Title
    gblnFlagDicas = True
    
    Dim cR As New cRegistry
    cR.ClassKey = HKEY_CURRENT_USER
    cR.SectionKey = "Software\Data Dynamics\ActiveBar\2.0\DDWordPad\MDI\MRU"
    moMRU.Load cR
    moMRU.MaxFileCount = 4
    
    'Set actBarra.Bands("barFind").Tools("frmFind").Custom = frmDockFind
    ' FillFontCombos
    UpdateToolbar actBarra
    CreateToolsLateral actBarra
    CarregaIconeEspecial actBarra, img_ListaIconesEspecificos
    
    On Error Resume Next
    If actBarra.Version = "2.5.0.22" Then
        If Dir(App.Path & "\Data\" & App.EXEName & "_" & gstrLoginUser & "_usageMenu.adg", vbArchive) <> "" Then
            actBarra.LoadMenuUsageData App.Path & "\Data\" & App.EXEName & "_" & gstrLoginUser & "_usageMenu.adg", ddSOFile
        End If
    End If
    staBarraStatus.Panels(2).Text = gstrDataDoSistema
    staBarraStatus.Panels(3).Text = gstrDataFormatada(gstrDataDoSistema(True), , True, True)
    staBarraStatus.Panels(4).Text = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    
    CarregaCategoriaConstrucao
    CarregaTamanhoMascaras
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not gblnTrocaUsuario Then
        If MsgBox("Tem certeza de que deseja finalizar o sistema?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Cancel = 1
    End If
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    'frmMenu.Move frmMenu.Left, frmMenu.Top, Me.ScaleWidth, Me.ScaleHeight
    'frmMenu.DefineTamanhoListBar
    On Error GoTo 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    DoEvents
    conMainADO.FechaBancoADO
    
    On Error Resume Next
    If actBarra.Version = "2.5.0.22" Then
        If Dir(App.Path & "\Data\", vbDirectory) = "" Then MkDir App.Path & "\Data\"
        actBarra.SaveMenuUsageData App.Path & "\Data\" & App.EXEName & "_" & gstrLoginUser & "_usageMenu.adg", ddSOFile
    End If
    
    'Nino
    If Not obfWordEditor Is Nothing Then obfWordEditor.Quit
   
End Sub



'============================================================================================
'============================================================================================
'============================================================================================
'============================================================================================
'============================================================================================
'============================================================================================
'============================================================================================
'============================================================================================
'============================================================================================

Private Sub actBarra_ComboDrop(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    If Tool.Name = "miFoFontName" Or Tool.Name = "miFoFontSize" Then
        If Not m_bFontsLoaded Then
            FillFontCombos
            m_bFontsLoaded = True
        End If
    End If
End Sub

Private Sub actBarra_ComboSelChange(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    If TypeOf ActiveForm Is IMDIDocument Then
        If Tool.Name = "miFoFontName" Then
            ActiveForm.rtf.SelFontName = Tool.Text
        ElseIf Tool.Name = "miFoFontSize" Then
            ActiveForm.rtf.SelFontSize = Val(Tool.Text)
        End If
        
    End If
End Sub

Public Sub Texto_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
Dim doc As IMDIDocument
Dim strAux As String

'If Not actBarra.Bands("mnuMain").Visible Then
'    Exit Sub
'End If
strAux = Tool.Name
    'If ActiveForm Is Nothing Then
'        If Not TypeOf ActiveForm Is IMDIDocument Then
            frmMDIDoc.Show
            Set doc = ActiveForm
            If ActiveForm.Tag = "Novo" And Tool.Name = "miFSave" Then
                Tool.Name = "miFSaveAs"
            End If
        If Tool.Name = "miFNew" Then
            ActiveForm.Tag = "Novo"
        ElseIf Not TypeOf ActiveForm Is IMDIDocument Then
            GoTo Sair
        End If
        
        If doc.CommandHandler(Tool) Then
            Tool.Name = strAux
            Exit Sub
        End If
    'End If
    
    Tool.Name = strAux
    
    Select Case Tool.Name
    ' File
    Case "miFNew"
        FileNew
    Case "miFOpen": FileOpen
    Case "miFMRU1", "miFMRU2", "miFMRU3", "miFMRU4":
        FileMRU moMRU.file(Tool.TagVariant)
    Case "miFExit":
        FileExit
        Exit Sub
    
    ' View
    Case "miVStandardToolbar": ViewStandard
    Case "miVFormatToolbar": ViewFormat
    Case "miVStatusBar": ViewStatusBar
    Case "miVOptions": ViewOptions
    
    ' Window
    Case "miWNew": FileNew
    Case "miWTileH": WindowTileH
    Case "miWTileV": WindowTileH
    Case "miWCascade": WindowCascade
    Case "miWArrangeIcons": WindowArrangeIcons
    
    ' Help
    Case "miHContents": HelpContents
    Case "miHWhatsThis": HelpWhatsThis Tool
    Case "miHAbout": HelpAbout
    End Select
'    UpdateToolbar True, actBarra

Sair:
'    UpdateToolbar True, actBarra
End Sub

'Private Sub MDIForm_Load()
'End Sub

Public Sub FileNew()
Dim F As frmMDIDoc
    Set F = New frmMDIDoc
    Dim doc As IMDIDocument
    Set doc = F
    iDoc = iDoc + 1
    ' This initializes the document and shows the form
    doc.InitDoc actBarra, "Documento" & CStr(iDoc), True
    ActiveForm.Tag = "Novo"
End Sub

Private Sub FileOpen()
Dim sFile As String
    On Error GoTo ehFileOpen 'set error trap
    If cdlg.VBGetOpenFileName(sFile, "RichEdit Documento", True, False, False, False, "RichText Files (*.rtf)|*.rtf", , App.Path, "Abrir Documento...", "RTF", Me.hWnd) Then
        If UCase(Right(sFile, 4)) <> ".RTF" Then
            ' possibly not RTF file, prompt
            If MsgBox("Você gostaria de abrir este arquivo?", _
                vbYesNo + vbQuestion, "Ele não tem o formato 'Rich Text File'") _
                = vbNo Then
                Exit Sub
            End If
        End If
        Dim F As New frmMDIDoc
        Dim doc As IMDIDocument
        Set doc = F
        ' Loads the file and shows the form
        doc.InitDoc actBarra, sFile, False
        moMRU.AddFile sFile
        DisplayMRU
        ' let the OS catch up
        DoEvents
    End If
    
ehFileOpen:
    Exit Sub
End Sub

Private Function FileMRU(sfileName As String)
    ' open file
    ActiveForm.rtf.LoadFile sfileName
    moMRU.AddFile sfileName
    DisplayMRU
    ActiveForm.rtf.DataChanged = False
End Function


Private Sub FileExit()
    UpdateToolbar actBarra
    'frmMenu.Visible = True
    Unload frmMDIDoc
    Set frmMDIDoc = Nothing
    
'    Unload Me
End Sub


Private Sub ViewStandard()
    'actBarra.Bands("barStandard").Visible = Not actBarra.Bands("barStandard").Visible
    actBarra.RecalcLayout
End Sub

Private Sub ViewFormat()
    'actBarra.Bands("barFormat").Visible = Not actBarra.Bands("barFormat").Visible
    actBarra.RecalcLayout
End Sub

Private Sub ViewStatusBar()
    'actBarra.Bands("sb").Visible = Not actBarra.Bands("sb").Visible
    actBarra.RecalcLayout
End Sub

Private Sub ViewOptions()

End Sub

Private Sub WindowTileH()
    Arrange vbTileHorizontal
End Sub

Private Sub WindowTileV()
    Arrange vbTileVertical
End Sub

Private Sub WindowCascade()
    Arrange vbCascade
End Sub

Private Sub WindowArrangeIcons()
    Arrange vbArrangeIcons
End Sub

Private Sub HelpContents()
Dim hr As Long
    hr = WinHelp(Me.hWnd, App.HelpFile, HELP_CONTENTS, 0&)
End Sub

Private Sub HelpWhatsThis(Tool As ActiveBar2LibraryCtl.Tool)
    Tool.Checked = True
    actBarra.WhatsThisHelpMode = True
End Sub

Private Sub HelpAbout()
    'frmAbout.Show vbModal
End Sub


Private Sub DisplayMRU()
Dim iFile As Long
    For iFile = 1 To moMRU.FileCount
        If (moMRU.FileExists(iFile)) Then
            With actBarra.Bands("mnuFile").Tools("miFMRU" & Trim$(Str(iFile)))
                If iFile = 1 Then .Checked = True
                '.Visible = True
                .Caption = moMRU.MenuCaption(iFile)
                .TagVariant = CStr(iFile)
            End With
        End If
    Next iFile
    ' Debug.Print (moMRU.FileCount > 0)
    'actBarra.Bands("mnuFile").Tools("miFMRUSep").Visible = (moMRU.FileCount > 0)
End Sub

Private Sub FillFontCombos()
Dim i As Integer
    With actBarra.Tools("miFoFontName")
        For i = 1 To Screen.FontCount
            .CBAddItem Screen.Fonts(i)
        Next
    End With
    With actBarra.Tools("miFoFontSize")
        .CBAddItem " 8"
        .CBAddItem " 9"
        .CBAddItem "10"
        .CBAddItem "11"
        .CBAddItem "12"
        .CBAddItem "14"
        .CBAddItem "16"
        .CBAddItem "18"
        .CBAddItem "20"
        .CBAddItem "22"
        .CBAddItem "24"
        .CBAddItem "26"
        .CBAddItem "28"
        .CBAddItem "36"
        .CBAddItem "48"
        .CBAddItem "72"
    End With
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        If actBarra.Bands("bndFormulario").Visible Then
            mnuBarra.Checked = True
        Else
            mnuBarra.Checked = False
        End If
        
        PopupMenu mnuArquivo
    End If
End Sub

Private Sub mnuBarra_Click()
    On Error Resume Next
    Select Case mnuBarra.Checked
        Case True
            mnuBarra.Checked = False
            actBarra.Bands("bndFormulario").Visible = False
            actBarra.RecalcLayout
        Case False
            mnuBarra.Checked = True
            actBarra.Bands("bndFormulario").Visible = True
            actBarra.RecalcLayout
    End Select
End Sub

Private Sub FormsPrincipal(intIndice As Integer)
    Select Case intIndice
        Case 573
            CarregaForm frmCadEmpresa
        Case 3000
            EfetuaLogoff
        Case 3001
            FinalizaSistema
    End Select
End Sub

Sub FinalizaSistema()
    'If MsgBox("Tem certeza de que deseja finalizar o sistema?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    '    conMainADO.FechaBancoADO
    '    End
    'End If
    Unload Me
End Sub

Sub EfetuaLogoff()
    If MsgBox("Tem certeza de que deseja efetuar" & vbCr & " logoff ?", vbYesNo + vbQuestion + vbDefaultButton2, "Efetuar logoff no sistema") = vbYes Then
        gblnTrocaUsuario = True
        Unload Me
        frmSplash.Show
    End If
End Sub

Private Sub time1_Timer()

 staBarraStatus.Panels(2).Text = gstrDataDoSistema
 staBarraStatus.Panels(3).Text = gstrDataFormatada(gstrDataDoSistema(True), , True, True)
End Sub

Private Sub FormsDocumentos(intIndice As Integer, Optional ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Dim stpDocumentPath As String
    Dim Obj             As Object
    Select Case intIndice
        Case 1355
            
        Case 1020
            frmSelDocWordWrapper.Show vbModal
            stpDocumentPath = frmSelDocWordWrapper.DocumentoSelecionado
            Unload frmSelDocWordWrapper: Set frmSelDocWordWrapper = Nothing
            OpenWordDocument stpDocumentPath
        Case 1156
            CarregaForm frmDocCertidaoValorVenal
        Case 1180
            CarregaForm frmTermoDeAcordo
        Case 1186
            CarregaForm frmDocCadastroMobiliario
        Case 1191
            For Each Obj In Forms
                If UCase$(Obj.Name) = "FRMDOCALVARAFUNCIONAMENTO" Then
                    Unload frmDocAlvaraFuncionamento
                End If
            Next
            gintCodSeguranca = 1191
            CarregaForm frmDocAlvaraFuncionamento
            
        Case 1364
            For Each Obj In Forms
                If UCase$(Obj.Name) = "FRMDOCALVARAFUNCIONAMENTO" Then
                    Unload frmDocAlvaraFuncionamento
                End If
            Next
            gintCodSeguranca = 1364
            CarregaForm frmDocAlvaraFuncionamento
            frmDocAlvaraFuncionamento.Caption = "Cadastro Mobiliário "
        Case 1210
            CarregaForm frmGuiaPrecoPublico
        Case Else
            MsgBox "Item não configurado."
    End Select
    
End Sub

Private Sub OpenWordTemplate(ByVal stpDocumentPath As String)
Dim objFileSystem As Scripting.FileSystemObject

   If stpDocumentPath <> Space$(0) Then
   
      Set objFileSystem = New Scripting.FileSystemObject
   
      If objFileSystem.FileExists(stpDocumentPath) Then
   
         If Not obfWordEditor Is Nothing Then obfWordEditor.Quit
            
         Set obfWordEditor = New cWordWrapper
         
         If obfWordEditor.IsInstalled Then
      
            obfWordEditor.GetContainer
            obfWordEditor.DocumentPath = stpDocumentPath
            obfWordEditor.DocumentFormat = WORDOPENFORMATTEMPLATE
            obfWordEditor.DocumentOpen
      
         Else
            MsgBox "O Microsoft Word não está instalado nesta máquina. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
         End If
   
      Else
         MsgBox "O documento do Microsoft Word selecionado não foi localizado. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
      End If
      
      Set objFileSystem = Nothing
      
   End If

End Sub

Private Sub OpenWordDocument(ByVal stpDocumentPath As String)
Dim objFileSystem As Scripting.FileSystemObject

   If stpDocumentPath <> Space$(0) Then
   
      Set objFileSystem = New Scripting.FileSystemObject
   
      If objFileSystem.FileExists(stpDocumentPath) Then
      
         If Not obfWordEditor Is Nothing Then obfWordEditor.Quit
         
         Set obfWordEditor = New cWordWrapper
         
         If obfWordEditor.IsInstalled Then
      
            obfWordEditor.GetContainer
            obfWordEditor.DocumentPath = stpDocumentPath
            obfWordEditor.DocumentFormat = WORDOPENFORMATDOCUMENT
            obfWordEditor.DocumentOpen
      
         Else
            MsgBox "O Microsoft Word não está instalado nesta máquina. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
         End If
   
      Else
         MsgBox "O documento do Microsoft Word selecionado não foi localizado. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
      End If
   
      Set objFileSystem = Nothing
   
   End If

End Sub

Private Sub obfWordEditor_Quit()
    Set obfWordEditor = Nothing
End Sub

Public Sub CarregaCategoriaConstrucao()
Dim strSql As String
Dim adoRec As New ADODB.Recordset

    strSql = "SELECT PkId FROM " & gstrCategoriaConstrucao
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 30, adoRec) Then
        
        With adoRec
            
            If Not .EOF Then
                
                    vetCategoriaConstrucao.ResidencialHorizontal = !Pkid
                    .MoveNext
                    vetCategoriaConstrucao.ResidencialVertical = !Pkid
                    .MoveNext
                    vetCategoriaConstrucao.ComercialHorizontal = !Pkid
                    .MoveNext
                    vetCategoriaConstrucao.ComercialVertical = !Pkid
                    .MoveNext
                    vetCategoriaConstrucao.Industrial = !Pkid
                    .MoveNext
                    vetCategoriaConstrucao.ImobiliarioGeral = !Pkid
                    .MoveNext
                    vetCategoriaConstrucao.ImobiliarioTerreno = !Pkid
                    .MoveNext
                    vetCategoriaConstrucao.EconomicoGeral = !Pkid

            End If
            
        End With
        
    End If
    
End Sub
