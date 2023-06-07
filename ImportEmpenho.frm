VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Begin VB.Form frmImportEmpenho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importacão"
   ClientHeight    =   7635
   ClientLeft      =   1920
   ClientTop       =   2745
   ClientWidth     =   8745
   HelpContextID   =   44
   Icon            =   "ImportEmpenho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   8745
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5055
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   150
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Importacao de Empenho"
      TabPicture(0)   =   "ImportEmpenho.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_DataEmpenho"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Tipo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_ItemDespesa"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_Arquivo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_ArquivoSelecionado"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lb_status"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProcesso"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbo_Historico"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dbcintItemDespesa"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dbcintTipo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtdtmData"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmd_Tipo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmd_ItemDespesa"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fra_CodEventoLiqContabil"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmd_Historico"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "fra_Historico"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmd_SelecionarArquivo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmd_Importar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboProgramaTrabalho"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt_TotalDotado"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_SaldoDotacao"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_tmp"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboCodigoReduzido"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmd_Imprimir"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtbitDigito"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtintExercicio"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtstrCodigo"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "Compras"
      TabPicture(1)   =   "ImportEmpenho.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_importCompras"
      Tab(1).Control(1)=   "fra_ImportCompra"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtstrCodigo 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         HideSelection   =   0   'False
         Left            =   1185
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1350
         Width           =   825
      End
      Begin VB.TextBox txtintExercicio 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2010
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1350
         Width           =   465
      End
      Begin VB.TextBox txtbitDigito 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2490
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1350
         Width           =   285
      End
      Begin VB.CommandButton cmd_Imprimir 
         Caption         =   "Gerar Relatório"
         Height          =   315
         Left            =   6135
         TabIndex        =   10
         Top             =   1260
         Width           =   1260
      End
      Begin VB.CommandButton cmd_importCompras 
         Caption         =   "Importar"
         Height          =   375
         Left            =   -73110
         TabIndex        =   37
         Top             =   1530
         Width           =   1125
      End
      Begin VB.Frame fra_ImportCompra 
         Caption         =   "Importação de Pedidos  de Empenho"
         Height          =   915
         Left            =   -74850
         TabIndex        =   31
         Top             =   510
         Width           =   2865
         Begin MSDataListLib.DataCombo cbo_intAutorizacaoDeCompra 
            Height          =   315
            Left            =   390
            TabIndex        =   32
            Top             =   450
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
      End
      Begin VB.ComboBox cboCodigoReduzido 
         Height          =   315
         ItemData        =   "ImportEmpenho.frx":107A
         Left            =   3930
         List            =   "ImportEmpenho.frx":107C
         TabIndex        =   30
         ToolTipText     =   "Código do programa de trabalho"
         Top             =   -300
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txt_tmp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txt_SaldoDotacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6390
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txt_TotalDotado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5340
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.ComboBox cboProgramaTrabalho 
         Height          =   315
         Left            =   4380
         Sorted          =   -1  'True
         TabIndex        =   25
         ToolTipText     =   "Código do programa de trabalho"
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmd_Importar 
         Caption         =   "Importar"
         Height          =   315
         Left            =   6135
         TabIndex        =   6
         Top             =   855
         Width           =   1260
      End
      Begin MSComDlg.CommonDialog cdl_SelecionarArquivo 
         Left            =   5970
         Top             =   -180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Importação de Empenho"
         Filter          =   "Arquivos de Empenho|*.txt|Todos os Arquivos|*.*"
      End
      Begin VB.CommandButton cmd_SelecionarArquivo 
         Caption         =   "Arquivo"
         Height          =   315
         Left            =   6135
         TabIndex        =   3
         Top             =   450
         Width           =   1260
      End
      Begin VB.Frame fra_Historico 
         Caption         =   " Histórico "
         Height          =   1575
         Left            =   180
         TabIndex        =   22
         Top             =   2505
         Width           =   7185
         Begin VB.TextBox txtstrHistorico 
            Height          =   1365
            Left            =   0
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   180
            Width           =   7125
         End
      End
      Begin VB.CommandButton cmd_Historico 
         Height          =   300
         Left            =   7050
         Picture         =   "ImportEmpenho.frx":107E
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "248"
         ToolTipText     =   "Clique para cadastar histórico"
         Top             =   4110
         Width           =   330
      End
      Begin VB.Frame fra_CodEventoLiqContabil 
         Caption         =   " Evento Contábil Para Liquidação"
         Height          =   585
         Left            =   180
         TabIndex        =   21
         Top             =   1785
         Width           =   7185
         Begin VB.TextBox txt_codEventoLiquidacao 
            Height          =   315
            Left            =   150
            MaxLength       =   15
            TabIndex        =   11
            Top             =   210
            Width           =   1185
         End
         Begin VB.ComboBox cbo_intEventoLiquidacao 
            Height          =   315
            Left            =   1350
            TabIndex        =   12
            Top             =   210
            Width           =   5400
         End
         Begin VB.CommandButton cmd_EventoLiquidacao 
            Height          =   300
            Left            =   6780
            Picture         =   "ImportEmpenho.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Tag             =   "247"
            ToolTipText     =   "Clique para cadastar convênio"
            Top             =   225
            Width           =   330
         End
      End
      Begin VB.CommandButton cmd_ItemDespesa 
         Height          =   300
         Left            =   5730
         Picture         =   "ImportEmpenho.frx":1792
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "244"
         ToolTipText     =   "Clique para cadastar itens de despesa"
         Top             =   900
         Width           =   330
      End
      Begin VB.CommandButton cmd_Tipo 
         Height          =   300
         Left            =   5730
         Picture         =   "ImportEmpenho.frx":1B1C
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Clique para cadastar tipo"
         Top             =   465
         Width           =   330
      End
      Begin VB.TextBox txtdtmData 
         Height          =   285
         Left            =   585
         TabIndex        =   0
         Top             =   480
         Width           =   1005
      End
      Begin MSDataListLib.DataCombo dbcintTipo 
         Height          =   315
         Left            =   2265
         TabIndex        =   1
         Top             =   450
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintItemDespesa 
         Height          =   315
         Left            =   1185
         TabIndex        =   4
         Top             =   900
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cbo_Historico 
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   4095
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblProcesso 
         AutoSize        =   -1  'True
         Caption         =   "Processo"
         Height          =   195
         Left            =   450
         TabIndex        =   38
         ToolTipText     =   "Item de despesa"
         Top             =   1410
         Width           =   660
      End
      Begin VB.Label lb_status 
         AutoSize        =   -1  'True
         Caption         =   "lb_status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   4785
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lbl_ArquivoSelecionado 
         AutoSize        =   -1  'True
         Caption         =   "lbl_ArquivoSelecionado"
         Height          =   195
         Left            =   1770
         TabIndex        =   24
         Top             =   4545
         Width           =   1665
      End
      Begin VB.Label lbl_Arquivo 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo Selecionado:"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   4545
         Width           =   1515
      End
      Begin VB.Label lbl_ItemDespesa 
         AutoSize        =   -1  'True
         Caption         =   "I.de Despesa"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         ToolTipText     =   "Item de despesa"
         Top             =   990
         Width           =   945
      End
      Begin VB.Label lbl_Tipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   1920
         TabIndex        =   19
         Top             =   570
         Width           =   315
      End
      Begin VB.Label lbl_DataEmpenho 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   525
         Width           =   345
      End
   End
   Begin TabDlg.SSTab tab_Inicial 
      Height          =   1530
      Left            =   120
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   150
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   2699
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Opções"
      TabPicture(0)   =   "ImportEmpenho.frx":1EA6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "optPedidoDeEmpenho"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "opt_FolhaDePagamento"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.OptionButton opt_FolhaDePagamento 
         Caption         =   "Folha de Pagamento"
         Height          =   315
         Left            =   90
         TabIndex        =   35
         Top             =   570
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optPedidoDeEmpenho 
         Caption         =   "Pedido de Empenho ""Compras"""
         Height          =   315
         Left            =   1920
         TabIndex        =   34
         Top             =   570
         Width           =   2565
      End
   End
   Begin VB.CommandButton cmd_Ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3630
      TabIndex        =   36
      Top             =   1740
      Width           =   1110
   End
End
Attribute VB_Name = "frmImportEmpenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim mblnAlterando           As Boolean
Dim mobjAux                 As Object
Dim mblnClickOk             As Boolean
Dim intArquivoCorreto       As Integer
Dim blnItemDespesa          As Boolean
Dim strListaProgTrabalho    As String
Dim mblnRollBack            As Boolean
Dim linhasArquivo()         As linhaEmpenho
Dim blnImportacao           As Boolean
Dim blnCancelado            As Boolean
Dim intTotalEmpenhos        As Integer
Dim blnImportacaoSucesso    As Boolean
Dim mobjLista               As Object
Dim dblVlTotEmpenhos        As Double
Dim strEmpenhoIF(0 To 1)    As String
Dim intContador             As Integer
    
Private Type linhaEmpenho
    strAno         As String
    strProjAtv     As String
    strElemento    As String
    strCredorPKID  As String
    strValor       As Double
    strFntRecursos As String
    STRTIPO        As String
    strEvento      As String
    strCLC      As String
End Type

Private Sub cmd_Importar_Click()
    
    If cmd_Importar.Caption = "Cancelar" Then
        If MsgBox("Deseja realmente cancelar a importação dos dados?", vbYesNo + vbDefaultButton2, "Orçamentário") = vbNo Then Exit Sub
        cmd_Importar.Caption = "Importar"
        cmd_SelecionarArquivo.Enabled = True
        blnCancelado = True
        HabilitaDesabilitaControles False
        Exit Sub
    End If
    
    If blnDadosOK Then
        
        cmd_Importar.Caption = "Cancelar"
        HabilitaDesabilitaControles True
        cmd_SelecionarArquivo.Enabled = False
        lb_status.Visible = True
        ImportaEmpenho
        If blnImportacaoSucesso = True Then
            lb_status.Visible = False
            blnImportacaoSucesso = False
        End If
        cmd_SelecionarArquivo.Enabled = True
        DoEvents
    End If
    
End Sub

Private Sub HabilitaDesabilitaControles(ByVal blnHabilitar As Boolean)

    TrocaCorObjeto txtdtmData, blnHabilitar
    TrocaCorObjeto dbcintTipo, blnHabilitar
    TrocaCorObjeto cmd_Tipo, blnHabilitar
    'TrocaCorObjeto cmd_SelecionarArquivo, blnHabilitar
    TrocaCorObjeto dbcintItemDespesa, blnHabilitar
    TrocaCorObjeto cmd_ItemDespesa, blnHabilitar
    'TrocaCorObjeto cmd_Importar, blnHabilitar
    TrocaCorObjeto txt_codEventoLiquidacao, blnHabilitar
    TrocaCorObjeto cbo_intEventoLiquidacao, blnHabilitar
    TrocaCorObjeto cmd_EventoLiquidacao, blnHabilitar
    TrocaCorObjeto txtstrHistorico, blnHabilitar
    TrocaCorObjeto cbo_Historico, blnHabilitar
    TrocaCorObjeto cmd_Historico, blnHabilitar
    TrocaCorObjeto txtstrCodigo, blnHabilitar
    TrocaCorObjeto txtbitDigito, blnHabilitar
    TrocaCorObjeto txtintExercicio, blnHabilitar
    
End Sub

Private Sub cmd_Imprimir_Click()
    Dim strSql As String
    strSql = "SELECT "
    strSql = strSql & " IM.*,"
    strSql = strSql & " CO.strNome"
    strSql = strSql & " FROM "
    strSql = strSql & gstrImpressaoFolha & " IM, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE "
    strSql = strSql & " IM.intCredor " & strOUTJSQLServer & "= CO.cdc " & strOUTJOracle
    strSql = strSql & " AND IM.blnImpressao = 1 "
    strSql = strSql & " ORDER BY intCredor, intProjAtividade"

    ImprimeRelatorio rptImpressaoFolha, strSql, "Impressão do Arquivo de Importação"
End Sub

Private Sub cmd_ok_Click()
    
    tab_3dPasta.Visible = True
    tab_Inicial.Visible = False
    cmd_Ok.Visible = False
    If opt_FolhaDePagamento.Value = True Then
       
        blnImportacao = True
        tab_3dPasta.TabVisible(0) = True
        Me.Height = 5730
        Me.Width = 7800
        tab_3dPasta.Height = 5115
        tab_3dPasta.Width = 7560
    Else
        tab_3dPasta.TabVisible(1) = True
        tab_3dPasta.Tab = 1
        
        Me.Height = 2850
        Me.Width = 4000
        tab_3dPasta.Height = 2235
        tab_3dPasta.Width = 3500
        
    End If

End Sub

Private Sub cmd_SelecionarArquivo_Click()
Dim strArquivo As String
Dim intRetornoDialogo As Integer
    
On Error GoTo saida
    
    cmd_SelecionarArquivo.MousePointer = vbCustom
    
    
    If cmd_SelecionarArquivo.Caption = "Cancelar" Then
        If MsgBox("Deseja realmente cancelar a leitura do arquivo?", vbYesNo + vbDefaultButton2, "Orçamentário") = vbNo Then Exit Sub
        blnCancelado = True
        cmd_SelecionarArquivo.Caption = "Arquivo"
        Exit Sub
    End If
    
    cdl_SelecionarArquivo.ShowOpen
    If cdl_SelecionarArquivo.CancelError <> 0 Then
        strArquivo = cdl_SelecionarArquivo.Filename
        lbl_ArquivoSelecionado = strArquivo
        cmd_SelecionarArquivo.Caption = "Cancelar"
        cmd_Importar.Enabled = False
        cmd_Imprimir.Enabled = False
        intArquivoCorreto = LeEstruturaArquivo(strArquivo)
        intTotalEmpenhos = intArquivoCorreto
        If intArquivoCorreto <> 0 Then
            lbl_ArquivoSelecionado = strArquivo
        Else
            lbl_ArquivoSelecionado = ""
            lb_status.Caption = ""
            lb_status.Visible = False
            cmd_Importar.Enabled = True
            cmd_Imprimir.Enabled = True
            cmd_SelecionarArquivo.Caption = "Arquivo"
            Exit Sub
        End If
    End If
    
    cmd_SelecionarArquivo.Caption = "Arquivo"
    cmd_Importar.Enabled = True
    cmd_Imprimir.Enabled = True
    lb_status.Visible = True
    lb_status.Caption = intArquivoCorreto & " Empenho(s) para importar no Valor de R$ " & gstrConvVrDoSql(dblVlTotEmpenhos)
    dblVlTotEmpenhos = 0
    cmd_SelecionarArquivo.MousePointer = vbDefault
    Exit Sub

saida:
    cmd_Importar.Enabled = True
    cmd_Imprimir.Enabled = True
    cmd_SelecionarArquivo.Caption = "Arquivo"
    If Err.Number <> 32755 Then
        ExibeMensagem "Ocorreu o seguinte erro:" & vbNewLine & Err.Description
    End If
    cmd_SelecionarArquivo.MousePointer = vbDefault
    
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 241
End Sub

Private Sub Form_Load()
Dim strSql As String
    
    dbcintTipo.Tag = "SELECT PKID, strDescricao FROM " & gstrTipoEmpenho & " ORDER BY strDescricao;strDescricao"
    dbcintItemDespesa.Tag = "SELECT PKID, strDescricao FROM " & gstrItemDespesa & " ORDER BY strDescricao;strDescricao"
    cbo_Historico.Tag = "SELECT PKID, strDescricao FROM " & gstrHistorico & " ORDER BY strDescricao;strDescricao"
    cbo_intAutorizacaoDeCompra.Tag = strQueryEmpenho & ";EC.INTPEDIDOEMPENHO"
    lbl_ArquivoSelecionado = ""
    tab_3dPasta.Visible = False
    tab_3dPasta.TabVisible(0) = False
    tab_3dPasta.TabVisible(1) = False
    
    Me.Height = 2655
    Me.Width = 4980
    
    TrocaCorObjeto cmd_Imprimir, True
    
    strSql = "SELECT DISTINCT RC.intPedidoEmpenho,"
    strSql = strSql & " RC.intPedidoEmpenho"
    strSql = strSql & " FROM "
    strSql = strSql & gstrRequisicaoCompras & " RC, "
    strSql = strSql & gstrEmpenhoContrato & " EC"
    strSql = strSql & " WHERE NOT RC.intPedidoEmpenho IS NULL AND"
    strSql = strSql & " RC.Pkid = EC.intRequisicaoDeCompra AND "
    strSql = strSql & "RC.strnumeroempenho IS Null"
    strSql = strSql & " ORDER BY RC.intPedidoEmpenho"
    
    'LeDaTabelaParaObj "",
    cbo_intAutorizacaoDeCompra.Tag = strSql & ";RC.intPedidoEmpenho"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If blnImportacao Then
        If strModoOperacao = gstrPreencherLista Then
            
            If Me.ActiveControl.Name = cbo_intEventoLiquidacao.Name Then
                preencheCboeventoLiq
            End If
            
            ToolBarGeral strModoOperacao, gstrEmpenho, False, , Me
            
        End If
        
        If strModoOperacao = gstrNovo Then
            LimpaDados
        End If
    Else
        If strModoOperacao = gstrPreencherLista Then
            PreencherListaDeOpcoes Me.ActiveControl
        End If
    End If

End Sub


Private Sub preencheCboeventoLiq()
    LeDaTabelaParaObj gstrEvento, cbo_intEventoLiquidacao, "SELECT PKID, strDescricao FROM " & gstrEvento & " WHERE intTipoEvento = 7"
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra , 6-alteracoes orcamentarias
    '              7-Liquidação
End Sub

Private Sub SSTab1_DblClick()

End Sub


Private Sub txt_codEventoLiquidacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_codEventoLiquidacao_LostFocus()
    PreencheEventobyCodigo txt_codEventoLiquidacao, cbo_intEventoLiquidacao, "7"
    cbo_intEventoLiquidacao_LostFocus
End Sub


Private Sub cbo_intEventoLiquidacao_Click()
   leCodigoEvento txt_codEventoLiquidacao, cbo_intEventoLiquidacao
End Sub

Private Sub cbo_intEventoLiquidacao_GotFocus()
    If cbo_intEventoLiquidacao.Text = "" Then txt_codEventoLiquidacao.Text = ""
End Sub

Private Sub cbo_intEventoLiquidacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_intEventoLiquidacao_LostFocus()
    If cbo_intEventoLiquidacao.Text = "" Then txt_codEventoLiquidacao.Text = ""
End Sub


Private Sub cmd_EventoLiquidacao_Click()
    CarregaForm frmCadEvento, cbo_intEventoLiquidacao, strQueryAplicarEventoLiq
End Sub

Private Sub LimpaDados()
    ReDim linhasArquivo(0) As linhaEmpenho
    strEmpenhoIF(0) = ""
    strEmpenhoIF(1) = ""
    cboCodigoReduzido.Clear
    cboProgramaTrabalho.Clear
    txt_TotalDotado.Text = ""
    txt_SaldoDotacao.Text = ""
    txt_tmp.Text = ""
    txtdtmData.Text = ""
    dbcintTipo.Text = ""
    dbcintItemDespesa.Text = ""
    txtstrCodigo.Text = ""
    txtintExercicio.Text = ""
    txtbitDigito.Text = ""

    txt_codEventoLiquidacao.Text = ""
    cbo_intEventoLiquidacao.ListIndex = -1
    txtstrHistorico.Text = ""
    cbo_Historico.Text = ""
    lbl_ArquivoSelecionado.Caption = ""
    mblnRollBack = False
    strListaProgTrabalho = ""
    lb_status.Caption = ""
    lb_status.Visible = False
    dblVlTotEmpenhos = 0
End Sub

Private Function strQueryAplicarEvento() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrEvento & " "
    strSql = strSql & "WHERE intTipoEvento = 2 "
    strQueryAplicarEvento = strSql
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra
End Function

Private Function strQueryAplicarEventoLiq() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrEvento & " "
    strSql = strSql & "WHERE intTipoEvento = 7 "
    strQueryAplicarEventoLiq = strSql
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra, , 6-alteracoes orcamentarias
    '              7-Liquidação
End Function

Private Sub txtbitDigito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigito
End Sub

Private Sub txtdtmData_GotFocus()
    MarcaCampo txtdtmData
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmData
End Sub

Private Sub txtdtmData_LostFocus()
    txtdtmData = gstrDataFormatada(txtdtmData)
End Sub


Private Sub dbcintTipo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub


Private Sub cmd_Tipo_Click()
    CarregaForm frmCadTipoEmpenho, dbcintTipo
End Sub


Private Sub dbcintItemdespesa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cmd_ItemDespesa_Click()
    CarregaForm frmCadItemDespesa, dbcintItemDespesa
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintExercicio_LostFocus()
    txtintExercicio = gstrAnoFormatado(txtintExercicio)
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub

Private Sub txtstrHistorico_GotFocus()
    MarcaCampo txtstrHistorico
End Sub

Private Sub txtstrHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrHistorico
End Sub

Private Sub cbo_Historico_Change()
    txtstrHistorico = Trim(cbo_Historico.Text)
End Sub

Private Sub cbo_Historico_Click(Area As Integer)
    DropDownDataCombo cbo_Historico, Me, Area
End Sub


Private Sub cbo_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cmd_Historico_Click()
    CarregaForm frmCadHistorico, cbo_Historico
End Sub

Private Function gCLCSemLigacao() As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset

    strSql = ""
    strSql = strSql & "SELECT CC.intCLC FROM "
    strSql = strSql & gstrCruzamentoContaExtra & " CC "
    strSql = strSql & " WHERE "
    strSql = strSql & "CC.blnSemLigacao = 1 "

   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            gCLCSemLigacao = Format(!intCLC, "0000")
         End If
      End With
   End If
End Function


Private Function LeEstruturaArquivo(ByVal strArquivo As String) As Integer
Dim strSql As String
Dim strSqlAux As String
Dim linha  As String
Dim cont   As Integer
Dim totalImportar As Integer
Dim arq
Dim strContaCruzada As String
Dim intContadorInterno As Integer
Dim intContadorEmpenhos As Integer
Dim intIndiceEmpenho As Integer
Dim dblValorSomaEmpenho As Double
Dim blnPrimeiraPassagem As Boolean
Dim strCodAnterior      As String
Dim strCodAtual         As String
Dim strCLCSemLigacao    As String
Dim adoResultado        As New ADODB.Recordset
Dim strSQLOrdenado      As String

    lb_status.Visible = True
    blnPrimeiraPassagem = False
    
    arq = FreeFile

    Open strArquivo For Input As arq 'so abre para qualquer coisa
    
    If EOF(arq) Then
        ReDim linhasArquivo(0) As linhaEmpenho
        LeEstruturaArquivo = 0
        ExibeMensagem "O Arquivo '" & strArquivo & "' está vazio."
        Close arq
        Exit Function
    End If
    
    Do While Not EOF(arq)
        cont = cont + 1
        Line Input #arq, linha
    Loop
    
    totalImportar = cont
    
    strCLCSemLigacao = gCLCSemLigacao
    
    Close arq
    
    Open strArquivo For Input As arq 'so abre para qualquer coisa
    
    cont = 0
    dblVlTotEmpenhos = 0
    
    'Vamos apagar a tabela de impressao do arquivo
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    gobjBanco.Execute "DELETE FROM " & gstrImpressaoFolha
    
    'grava na tabela apenas para ordenar, e depois montar os empenhos (registros excluidos depois)
    Do While Not EOF(arq)
        cont = cont + 1
        Line Input #arq, linha
        If Len(linha) = 49 Or Len(linha) = 45 Then
        
            If Len(linha) = 45 Then
                linha = Mid(linha, 1, 24) & strCLCSemLigacao & Mid(linha, 25)
            End If
                        
            If blnCancelado Then GoTo Cancelar
        
            ReDim Preserve linhasArquivo(cont) As linhaEmpenho
            linhasArquivo(cont).strAno = Mid(linha, 1, 4)
            linhasArquivo(cont).strProjAtv = CStr(Val(Mid(linha, 5, 4)))
            linhasArquivo(cont).strElemento = CStr(Val(Mid(linha, 9, 8)))
            linhasArquivo(cont).strCredorPKID = verificaCredor(CStr(Val(Mid(linha, 17, 4))))
            linhasArquivo(cont).strFntRecursos = Mid(linha, 21, 4)
            linhasArquivo(cont).strCLC = Mid(linha, 25, 4)
            linhasArquivo(cont).strValor = Mid(linha, 29, 20)
            linhasArquivo(cont).strEvento = RetornaEvento(linhasArquivo(cont).strElemento)
            linhasArquivo(cont).STRTIPO = Mid(linha, 49, 1)

            
            DoEvents
            'lb_status.Caption = "Ordenando " & CStr(cont) & "° Registro do arquivo do Total de: " & CStr(totalImportar)
            lb_status.Caption = "Ordenando o arquivo de Importação: Total " & gstrConvVrDoSql(((Val(cont) / Val(totalImportar)) * 100), 0) & "%"
            strSql = ""
            strSql = strSql & "INSERT INTO " & gstrImpressaoFolha & " ("
            strSql = strSql & "intID,intAno, intProjAtividade, intCategoriaEcon, intCredor, "
            strSql = strSql & "strFonteRecurso,intFonteRecurso,dblValor, intEvento, intCLC, strTipo, blnImpressao"
            strSql = strSql & ") "
            strSql = strSql & " VALUES ("
            strSql = strSql & " 0,"
            strSql = strSql & linhasArquivo(cont).strAno & ", "
            strSql = strSql & linhasArquivo(cont).strProjAtv & ", "
            strSql = strSql & linhasArquivo(cont).strElemento & ", "
            strSql = strSql & Val(Mid(linha, 17, 4)) & ", "
            strSql = strSql & "'" & linhasArquivo(cont).strFntRecursos & "', "
            strSql = strSql & linhasArquivo(cont).strFntRecursos & ", "
            strSql = strSql & gstrConvVrParaSql(linhasArquivo(cont).strValor) & ","
            strSql = strSql & linhasArquivo(cont).strEvento & ","
            strSql = strSql & linhasArquivo(cont).strCLC & ","
            strSql = strSql & "'" & linhasArquivo(cont).STRTIPO & "',0"
            strSql = strSql & " )"


            If Not gobjBanco.Execute(strSql, True) Then
                blnCancelado = False
                cmd_SelecionarArquivo.Enabled = True
                lb_status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                ExibeMensagem "Ocorreu um problema ao gerar spool de impressão, nenhum registro foi gravado."
                LeEstruturaArquivo = 0
                cmd_Importar.Enabled = True
                cmd_Imprimir.Enabled = False
            End If
        Else
            ReDim linhasArquivo(0) As linhaEmpenho
            LeEstruturaArquivo = 0
            Close arq
            gobjBanco.ExecutaRollbackTrans
            ExibeMensagem "A Estrutura do arquivo está incorreta para a importação de Empenho."
            Exit Function
        End If
    Loop
    
    strSQLOrdenado = "SELECT * FROM " & gstrImpressaoFolha & " ORDER BY intano,intprojatividade , intcategoriaecon, intcredor, intclc "
    If Not gobjBanco.CriaADO(strSQLOrdenado, 60, adoResultado) Then
        blnCancelado = False
        cmd_SelecionarArquivo.Enabled = True
        lb_status.Visible = False
        gobjBanco.ExecutaRollbackTrans
        ExibeMensagem "Ocorreu um problema ao gerar spool de impressão, nenhum registro foi gravado."
        LeEstruturaArquivo = 0
        cmd_Importar.Enabled = True
        cmd_Imprimir.Enabled = False
    End If
    
    ReDim linhasArquivo(0) As linhaEmpenho
    cont = 0
    
    Do While Not adoResultado.EOF
        cont = cont + 1

            DoEvents
            lb_status.Caption = "Lendo " & CStr(cont) & "° Registro do arquivo do Total de: " & CStr(totalImportar)
            
            If blnCancelado Then GoTo Cancelar
        
            ReDim Preserve linhasArquivo(cont) As linhaEmpenho
            linhasArquivo(cont).strAno = adoResultado!intAno
            linhasArquivo(cont).strProjAtv = adoResultado!INTPROJATIVIDADE
            linhasArquivo(cont).strElemento = adoResultado!INTCATEGORIAECON
            linhasArquivo(cont).strCredorPKID = verificaCredor(adoResultado!intCredor)
            linhasArquivo(cont).strFntRecursos = adoResultado!STRFONTERECURSO
            linhasArquivo(cont).strCLC = adoResultado!intCLC
            linhasArquivo(cont).strValor = adoResultado!dblValor
            linhasArquivo(cont).strEvento = adoResultado!intEvento
            dblVlTotEmpenhos = dblVlTotEmpenhos + Val(gstrConvVrParaSql(adoResultado!dblValor))
            linhasArquivo(cont).STRTIPO = adoResultado!STRTIPO
            
            If linhasArquivo(cont).strCredorPKID = "" Then
                ReDim linhasArquivo(0) As linhaEmpenho
                LeEstruturaArquivo = 0
                Close arq
                gobjBanco.ExecutaRollbackTrans
                ExibeMensagem "Um dos Credores não faz parte do Orçamentário." & vbNewLine & "Este arquivo não pode ser importado"
                Exit Function
            End If
            
            strSql = ""
'            strSql = strSql & "INSERT INTO " & gstrImpressaoFolha & " ("
'            strSql = strSql & "intID,intAno, intProjAtividade, intCategoriaEcon, intCredor, "
'            strSql = strSql & "strFonteRecurso,intFonteRecurso,dblValor, intEvento, intCLC, strTipo, blnImpressao"
'            strSql = strSql & ") "
'            strSql = strSql & " VALUES ("
'            strSql = strSql & " (SELECT " & gstrISNULL("MAX(intID)", "0") & " + 1 FROM " & gstrImpressaoFolha & "),"
'            strSql = strSql & linhasArquivo(cont).strAno & ", "
'            strSql = strSql & linhasArquivo(cont).strProjAtv & ", "
'            strSql = strSql & linhasArquivo(cont).strElemento & ", "
'            strSql = strSql & Val(adoResultado!INTCREDOR) & ", "
'            strSql = strSql & "'" & linhasArquivo(cont).strFntRecursos & "', "
'            strSql = strSql & linhasArquivo(cont).strFntRecursos & ", "
'            strSql = strSql & gstrConvVrParaSql(linhasArquivo(cont).strValor) & ","
'            strSql = strSql & linhasArquivo(cont).strEvento & ","
'            strSql = strSql & linhasArquivo(cont).strCLC & ","
'            strSql = strSql & "'" & linhasArquivo(cont).STRTIPO & "',0"
'            strSql = strSql & " )"
            
            strSql = strSql & " UPDATE " & gstrImpressaoFolha & " SET  intID = " & cont & " WHERE PKID =" & adoResultado!Pkid
            
            
             strContaCruzada = retornaContaCruzada(linhasArquivo(cont).strCLC)
             
            If strContaCruzada = "0" Then
               ExibeMensagem "Problemas durante a Leitura dos Registros." & _
               vbNewLine & "O CLC " & linhasArquivo(cont).strCLC & " ainda não foi cruzado e a importação não pode continuar sem esta informação."

               GoTo Cancelar
            End If
             
             
             strCodAtual = linhasArquivo(cont).strAno & linhasArquivo(cont).strProjAtv & linhasArquivo(cont).strElemento & linhasArquivo(cont).strCredorPKID & linhasArquivo(cont).strFntRecursos
            
             'If strContaCruzada = "-1" And strCodAtual <> strCodAnterior Then
             If strCodAtual <> strCodAnterior Then
repeteparaUltima:
                 intContadorInterno = intContadorInterno + 1
                 If intContadorInterno = 2 Then
                    intContadorInterno = 1
                    intContadorEmpenhos = intContadorEmpenhos + 1
                     
                    strSqlAux = ""
                    strSqlAux = strSqlAux & "INSERT INTO " & gstrImpressaoFolha & " ("
                    strSqlAux = strSqlAux & "intAno, intProjAtividade, intCategoriaEcon, intCredor, "
                    strSqlAux = strSqlAux & "strFonteRecurso, intFonteRecurso,dblValor, intEvento, intCLC, strTipo, blnImpressao"
                    strSqlAux = strSqlAux & ") "
                    strSqlAux = strSqlAux & " VALUES ("
                    strSqlAux = strSqlAux & linhasArquivo(intIndiceEmpenho).strAno & ", "
                    strSqlAux = strSqlAux & linhasArquivo(intIndiceEmpenho).strProjAtv & ", "
                    strSqlAux = strSqlAux & linhasArquivo(intIndiceEmpenho).strElemento & ", "
                    strSqlAux = strSqlAux & LeCDCCredor(linhasArquivo(intIndiceEmpenho).strCredorPKID) & ", "
                    strSqlAux = strSqlAux & "'" & linhasArquivo(intIndiceEmpenho).strFntRecursos & "', "
                    strSqlAux = strSqlAux & linhasArquivo(intIndiceEmpenho).strFntRecursos & ", "
                    strSqlAux = strSqlAux & gstrConvVrParaSql(dblValorSomaEmpenho) & ","
                    strSqlAux = strSqlAux & linhasArquivo(intIndiceEmpenho).strEvento & ","
                    strSqlAux = strSqlAux & linhasArquivo(intIndiceEmpenho).strCLC & ","
                    strSqlAux = strSqlAux & "'" & linhasArquivo(intIndiceEmpenho).STRTIPO & "',1"
                    strSqlAux = strSqlAux & " )"
                     
                    If Not gobjBanco.Execute(strSqlAux, True) Then
                        blnCancelado = False
                        cmd_SelecionarArquivo.Enabled = True
                        lb_status.Visible = False
                        gobjBanco.ExecutaRollbackTrans
                        ExibeMensagem "Ocorreu um problema ao gerar spool de impressão, nenhum registro foi gravado."
                        LeEstruturaArquivo = 0
                        cmd_Importar.Enabled = True
                        cmd_Imprimir.Enabled = False
                    End If
                     
                    dblValorSomaEmpenho = 0
                 End If
                 intIndiceEmpenho = cont
             End If
             strCodAnterior = linhasArquivo(cont).strAno & linhasArquivo(cont).strProjAtv & linhasArquivo(cont).strElemento & linhasArquivo(cont).strCredorPKID & linhasArquivo(cont).strFntRecursos
             dblValorSomaEmpenho = dblValorSomaEmpenho + Val(gstrConvVrParaSql(linhasArquivo(cont).strValor))
            
            If adoResultado.RecordCount = cont And blnPrimeiraPassagem = True Then
                Exit Do
            End If
            
            If Not gobjBanco.Execute(strSql, True) Then
                blnCancelado = False
                cmd_SelecionarArquivo.Enabled = True
                lb_status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                ExibeMensagem "Ocorreu um problema ao gerar spool de impressão, nenhum registro foi gravado."
                LeEstruturaArquivo = 0
                cmd_Importar.Enabled = True
                cmd_Imprimir.Enabled = False
            End If
            
            If adoResultado.RecordCount = cont And blnPrimeiraPassagem = False Then
                blnPrimeiraPassagem = True
            
                GoTo repeteparaUltima:
                
            End If

        adoResultado.MoveNext
    Loop
        
    gobjBanco.Execute "DELETE FROM " & gstrImpressaoFolha & " WHERE blnImpressao = 3 "
    
    gobjBanco.ExecutaCommitTrans
    
    Close arq
    LeEstruturaArquivo = intContadorEmpenhos
    blnCancelado = False
    cmd_Importar.Enabled = True
    cmd_Imprimir.Enabled = True
    Exit Function
Cancelar:
    blnCancelado = False
    cmd_SelecionarArquivo.Enabled = True
    lb_status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    ExibeMensagem "A leitura do arquivo de importação de empenhos foi cancelada, nenhum registro foi gravado."
    LeEstruturaArquivo = 0
    cmd_Importar.Enabled = True
    cmd_Imprimir.Enabled = False
End Function

Private Function verificaCredor(ByVal strCredorPKID As String) As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset

    strSql = ""
    strSql = strSql & "SELECT A.PKid , A.strNome FROM "
    strSql = strSql & gstrContribuinte & " A, "
    strSql = strSql & gstrModuloContribuinte & " B, "
    strSql = strSql & gstrItens & " C WHERE "
    strSql = strSql & "B.intContribuinte = A.PKid AND "
    strSql = strSql & "B.intItem = C.PKid AND "
    strSql = strSql & "B.intItem = " & gintModulo
    strSql = strSql & " AND A.CDC = " & strCredorPKID

   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            verificaCredor = !Pkid
         End If
      End With
   End If
End Function





Private Function blnDadosOK() As Boolean
    Dim dtmDtEncerramento As Date
    Dim adoCampo          As ADODB.Field
    Dim adoResultado      As ADODB.Recordset
    Dim strSql            As String

    If gblnDataValida(txtdtmData) = False Then
        ExibeMensagem "A data do Empenho tem que ser informada corretamente."
        If txtdtmData.Enabled Then
            If txtdtmData.Enabled Then txtdtmData.SetFocus
        End If
        Exit Function
    End If
    
    If (Year(txtdtmData) <> gintExercicio) Then
        ExibeMensagem "A data do empenho tem que estar dentro do ano de " & gintExercicio & "."
        If txtdtmData.Enabled Then txtdtmData.SetFocus
        Exit Function
    End If
    
    
    dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
        
    If dtmDtEncerramento = Empty Then
       Exit Function
    Else
       If CDate(txtdtmData) <= dtmDtEncerramento Then
          ExibeMensagem "A data do Empenho deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
          If txtdtmData.Enabled Then txtdtmData.SetFocus
          Exit Function
       End If
    End If
    
 
    
    
    If dbcintTipo.Text = "" Then
        ExibeMensagem "O Tipo tem que ser informado."
        If dbcintTipo.Enabled Then dbcintTipo.SetFocus
        Exit Function
    End If
    
    'verifica se o item de despesa deverá ser obrigatorio
    Set gobjBanco = New clsBanco
    
    strSql = ""
    strSql = strSql & "Select bytItemDespesaObrigatorio from " & gstrConfiguracaoGeral & " Where Pkid = 1 "
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                blnItemDespesa = IIf(Not IsNull(!bytItemDespesaObrigatorio), .Fields("bytItemDespesaObrigatorio").Value, False)
            End If
        End With
    End If
    If blnItemDespesa = True Then
        If dbcintItemDespesa.Text = "" Then
            ExibeMensagem "O Item de despesa tem que ser informado."
            If dbcintItemDespesa.Enabled Then dbcintItemDespesa.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtstrCodigo) = "" Or Trim(txtintExercicio) = "" Or (txtbitDigito) = "" Then
        ExibeMensagem "Os dados do processo devem ser preenchidos corretamente."
        If txtstrCodigo.Enabled Then txtstrCodigo.SetFocus
        Exit Function
    End If
    
    'If blnValidarProcesso Then
        strSql = "SELECT * FROM " & gstrProtocolizacaoProcesso & " WHERE strCodigo = " & txtstrCodigo & " AND bitDigito = " & txtbitDigito & " AND intExercicio = " & txtintExercicio
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
            If adoResultado.EOF And adoResultado.BOF Then
                ExibeMensagem "Este processo não está cadastrado."
                If txtstrCodigo.Enabled Then txtstrCodigo.SetFocus
                Exit Function
            End If
        End If
    'End If
        
    If cbo_intEventoLiquidacao.ListIndex = -1 Then
        ExibeMensagem "O Evento Contabil da Liquidacao tem que ser informado."
        If cbo_intEventoLiquidacao.Enabled Then cbo_intEventoLiquidacao.SetFocus
        Exit Function
    End If

    If dbcintTipo.Text = "" Then
        ExibeMensagem "O Tipo tem que ser informado."
        If dbcintTipo.Enabled Then dbcintTipo.SetFocus
        Exit Function
    End If
    
    If intArquivoCorreto = 0 Then
        ExibeMensagem "O Arquivo de empenhos tem que ser informado."
        If cmd_SelecionarArquivo.Enabled Then cmd_SelecionarArquivo.SetFocus
        Exit Function
    End If
    
    
    If lbl_ArquivoSelecionado = "" Then
        ExibeMensagem "O Arquivo de empenhos tem que ser informado."
        If cmd_SelecionarArquivo.Enabled Then cmd_SelecionarArquivo.SetFocus
        Exit Function
    End If
    
    
    blnDadosOK = True
End Function

Private Sub ImportaEmpenho()
    Dim linha
    Dim i As Integer
    Dim strProgTrabalhoPKID As String
    Dim blnEsgotouSaldo As Boolean
    Dim blnExibeMSgEsgotouSaldo As Boolean
    Dim blnDotacaoNaoEncontrada As Boolean
    Dim blnEventoNaoCompativel As Boolean
    Dim strErro As String
    Dim adoResultado As New ADODB.Recordset
    Dim strContaCruzada As String
    Dim intContadorInterno As Integer
    Dim strCodAnterior      As String
    Dim strCodAtual         As String

    
    blnEsgotouSaldo = False
    intContador = 1

        Set gobjBanco = New clsBanco
        mblnRollBack = False
        gobjBanco.ExecutaBeginTrans

        If gobjBanco.CriaADO(strQueryAgrupaFolha, 20, adoResultado) Then
            While Not adoResultado.EOF
        
                linhasArquivo(0).strAno = adoResultado!intAno
                linhasArquivo(0).strCredorPKID = verificaCredor(adoResultado!intCredor)
                linhasArquivo(0).strElemento = adoResultado!INTCATEGORIAECON
                linhasArquivo(0).strEvento = adoResultado!intEvento
                linhasArquivo(0).strFntRecursos = adoResultado!STRFONTERECURSO
                linhasArquivo(0).strProjAtv = adoResultado!INTPROJATIVIDADE
                linhasArquivo(0).STRTIPO = adoResultado!STRTIPO
                linhasArquivo(0).strValor = "0"

                Do While Not adoResultado.EOF
                    strContaCruzada = retornaContaCruzada(adoResultado!intCLC)
                    
                    strCodAtual = adoResultado!intAno & adoResultado!INTPROJATIVIDADE & adoResultado!INTCATEGORIAECON & adoResultado!intCredor & adoResultado!STRFONTERECURSO
                   
                    'If strContaCruzada = "-1" And strCodAtual <> strCodAnterior Then
                    If strCodAtual <> strCodAnterior Then
                        intContadorInterno = intContadorInterno + 1
                        If intContadorInterno = 2 Then
                            intContadorInterno = 0
                            Exit Do
                        End If
                    End If
                    
                    strCodAnterior = adoResultado!intAno & adoResultado!INTPROJATIVIDADE & adoResultado!INTCATEGORIAECON & adoResultado!intCredor & adoResultado!STRFONTERECURSO
                    
                    linhasArquivo(0).strValor = CStr(Val(gstrConvVrParaSql(linhasArquivo(0).strValor)) + Val(gstrConvVrParaSql(adoResultado!dblValor)))
                    adoResultado.MoveNext
                Loop
                i = i + 1
                If blnCancelado Then GoTo Cancelar
                With linhasArquivo(0)
                    DoEvents
                    lb_status.Caption = "Verificando " & CStr(i) & "° Registro " & "  -  Total (" & CStr(intTotalEmpenhos) & ")"
                    strProgTrabalhoPKID = gstrProgTrabalhoPkid(.strProjAtv, .strElemento, .strFntRecursos)
                    
                    If strProgTrabalhoPKID = "" Then
                        ExibeMensagem "A Dotação correspondente a:" & _
                        vbNewLine & "Projeto Atividade: " & .strProjAtv & _
                        vbNewLine & "Elemento de Despesa: " & .strElemento & _
                        vbNewLine & "Fonte de Resursos: " & .strFntRecursos & _
                        vbNewLine & "Não está cadastrada no sistema."
                        
                        blnDotacaoNaoEncontrada = True
                    End If
                    
                    If blnCancelado Then GoTo Cancelar
                    
                    If strProgTrabalhoPKID <> "" Then
                        cboProgramaTrabalho.Clear
                        cboProgramaTrabalho.AddItem (.strProjAtv)
                        cboProgramaTrabalho.ItemData(0) = strProgTrabalhoPKID
                        cboProgramaTrabalho.ListIndex = 0
                        
                        cboCodigoReduzido.Clear
                        cboCodigoReduzido.AddItem (.strProjAtv)
                        cboCodigoReduzido.ItemData(0) = strProgTrabalhoPKID
                        cboCodigoReduzido.ListIndex = 0
                                  
                        If blnCancelado Then GoTo Cancelar
                        
                        LeProgramaTrabalho cboProgramaTrabalho, cboCodigoReduzido, _
                                           txt_tmp, txt_tmp, txt_tmp, txt_tmp, _
                                           txt_tmp, txt_tmp, _
                                           txt_tmp, txt_tmp, txt_tmp, _
                                           txt_tmp, txt_SaldoDotacao, _
                                           txt_TotalDotado, txt_tmp, , , , , , , , , , txtdtmData
                                       
                        If blnCancelado Then GoTo Cancelar
                        If CDbl(gstrConvVrDoSql(.strValor)) > CDbl(txt_SaldoDotacao) Then
                            blnExibeMSgEsgotouSaldo = True
                            blnEsgotouSaldo = True
                        End If
                                       
                        DoEvents
                        
                        If blnCancelado Then GoTo Cancelar
                        If VerificaEventoProgramaTrabalho(.strEvento, strProgTrabalhoPKID) = False Then
                            ExibeMensagem "A dotação : " & gstrProgramaTrabalhoDescricao(gstrItemData(cboProgramaTrabalho)) & _
                            vbNewLine & "não é compativel com o evento selecionado através Elemento de Despesa."
                            blnEventoNaoCompativel = True
                        End If
                    
                        DoEvents
                        If blnCancelado Then GoTo Cancelar
                        If blnExibeMSgEsgotouSaldo = True Then
                              ExibeMensagem "A dotação :" & gstrProgramaTrabalhoDescricao(gstrItemData(cboProgramaTrabalho)) & " - " & " não possui saldo para este empenho." & _
                                            vbNewLine & "Valor Empenho: " & gstrConvVrDoSql(.strValor) & "       Saldo :" & gstrConvVrDoSql(txt_SaldoDotacao) & "       Total a Suplementar :" & gstrConvVrDoSql(CDbl(.strValor) - CDbl(txt_SaldoDotacao))
                              blnExibeMSgEsgotouSaldo = False
                        End If
                                       
                        DoEvents
                        lb_status.Caption = "Importando " & CStr(i) & "° Registro " & "  -  Total (" & CStr(intTotalEmpenhos) & ")"
                        
                            IncluiEmpenho .strAno, strProgTrabalhoPKID, .strCredorPKID, _
                                          .strValor, .STRTIPO, txt_SaldoDotacao, txt_TotalDotado, .strEvento, .strProjAtv, .strElemento, .strFntRecursos
                    End If
                    
                    If mblnRollBack Then
                        GoTo voltacampos:
                        strEmpenhoIF(0) = ""
                        strEmpenhoIF(1) = ""
                    End If
                    
                    If blnCancelado Then GoTo Cancelar
                End With
            Wend
        End If
        
        If blnEsgotouSaldo = True Or blnEventoNaoCompativel = True Or blnDotacaoNaoEncontrada = True Then
            If blnEsgotouSaldo = True Then strErro = vbNewLine & "Falta de saldo em uma ou mais Dotações."
            If blnEventoNaoCompativel = True Then strErro = strErro & vbNewLine & "Um ou mais Empenhos não é compatível com o evento selecionado."
            If blnDotacaoNaoEncontrada = True Then strErro = strErro & vbNewLine & "Uma ou mais Dotações não foram encontradas."
            ExibeMensagem "A importação foi cancelada pelos seguintes problemas:" & strErro & vbNewLine & "Nenhum registro foi gravado."
            
            gobjBanco.ExecutaRollbackTrans
            GoTo voltacampos:
        End If
        
        gobjBanco.ExecutaCommitTrans
        intArquivoCorreto = 0
        ExibeMensagem "Foram importados " & CStr(i) & " Empenhos com sucesso. " & vbNewLine & "Empenho Inicial: " & strEmpenhoIF(0) & vbNewLine & "Empenho Final: " & strEmpenhoIF(1)
        blnImportacaoSucesso = True
        LimpaDados
        limpaTabelaRelatorio
        GoTo voltacampos:
    
    Exit Sub
Cancelar:
    gobjBanco.ExecutaRollbackTrans
    ExibeMensagem "A importação foi cancelada pelo Usuário, nenhum registro foi gravado."
    
voltacampos:
    cmd_Importar.Caption = "Importar"
    cmd_SelecionarArquivo.Enabled = True
    blnCancelado = False
    HabilitaDesabilitaControles False
    lb_status.Visible = True
    lb_status.Caption = CStr(intTotalEmpenhos) & " Empenho(s) para importar"
End Sub

Private Sub limpaTabelaRelatorio()
    Set gobjBanco = New clsBanco
    gobjBanco.Execute "DELETE FROM " & gstrImpressaoFolha
End Sub

Private Function RetornaEvento(ByVal strElementoDespesa As String) As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
   
            
    strSql = "SELECT EVD.intEvento from "
    strSql = strSql & gstrEventoContaContabilDebito & " EVD, "
    strSql = strSql & gstrPlanoConta & " PC, "
    strSql = strSql & gstrEvento & " EV "
    strSql = strSql & " WHERE "
    strSql = strSql & " PC.Pkid = EVD.intContaContabil "
    strSql = strSql & " AND EV.pkid = evd.intevento "
    strSql = strSql & " AND EV.INTTIPOEVENTO = 2 "
    strSql = strSql & " AND " & strSUBSTRING & "(pc.strcontacontabil,1,3) =  '" & gstrDigitoDespesa & Mid(strElementoDespesa, 1, 2) & "'"
    
      
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                RetornaEvento = !intEvento
            Else
                RetornaEvento = "0"
            End If
        End With
    End If

End Function



Private Function VerificaEventoProgramaTrabalho(ByVal intEvento As String, ByVal intProgramadeTrabalho As String) As Boolean
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
   
      strListaProgTrabalho = ""
      
      If intEvento = "0" Then
            VerificaEventoProgramaTrabalho = False
            Exit Function
      End If
        
      strSql = ""
      strSql = strSql & "SELECT PT.PKId, PT.intCodigoReduzido, PT.strCodigo, "
      strSql = strSql & " ED.strCodigoElementoDespesa "
      strSql = strSql & " FROM " & gstrProgramaDeTrabalho & " PT, "
      strSql = strSql & gstrElementoDespesa & " ED "
      strSql = strSql & " WHERE PT.intElementoDespesa = ED.PKID AND "
      strSql = strSql & strSUBSTRING & "(ED.strCodigoElementoDespesa,1," & Len(BuscaCodigosPeloEvento(CInt(intEvento), gstrDigitoDespesa, "D", 2)) & ") = '" & _
                                                                      BuscaCodigosPeloEvento(CInt(intEvento), gstrDigitoDespesa, "D", 2) & "'"
      strSql = strSql & " AND PT.intExercicio = " & gintExercicio
      strSql = strSql & " AND PT.pkid = " & intProgramadeTrabalho
      
      Set gobjBanco = New clsBanco
      If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
          With adoResultado
              If Not .EOF Then
                VerificaEventoProgramaTrabalho = True
              Else
                VerificaEventoProgramaTrabalho = False
              End If
          End With
      End If
      
   
End Function

Private Function gstrProgramaTrabalhoDescricao(ByVal strPKId As String) As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
      strSql = ""
      strSql = strSql & "SELECT PT.PKId, PT.intCodigoReduzido, PT.strCodigo "
      strSql = strSql & " FROM " & gstrProgramaDeTrabalho & " PT "
      strSql = strSql & " WHERE PT.pkid = " & strPKId
      
      Set gobjBanco = New clsBanco
      If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
          With adoResultado
              If Not .EOF Then
                    gstrProgramaTrabalhoDescricao = !intCodigoReduzido & " - " & !strCodigo
              End If
          End With
      End If
      
   
End Function


Private Function gstrProgTrabalhoPkid(ByVal strProjAtv As String, _
                                      ByVal strElemento As String, _
                                      ByVal strFntRecursos As String) As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset

    strSql = ""
    strSql = strSql & "SELECT PT.PKID FROM "
    strSql = strSql & gstrProgramaDeTrabalho & " PT,"
    strSql = strSql & gstrElementoDespesa & " ED,"
    strSql = strSql & gstrFonteRecurso & " FR,"
    strSql = strSql & gstrProjeto & " PJ"
    strSql = strSql & " WHERE "
    strSql = strSql & strSUBSTRING & "(ED.Strcodigoelementodespesa, 1," & CStr(Len(strElemento)) & ") =" & strElemento
    strSql = strSql & " AND FR.STRCODIGO ='" & strFntRecursos & "'"
    strSql = strSql & " AND PJ.STRCODIGO = " & strProjAtv
    strSql = strSql & " AND PT.Intprojetoatividade = PJ.PKID"
    strSql = strSql & " AND PT.Intelementodespesa = ED.PKID"
    strSql = strSql & " AND PT.Intfonterecurso = FR.pkid"
    strSql = strSql & " AND PT.intExercicio =" & gintExercicio
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                gstrProgTrabalhoPkid = !Pkid
            End If
        End With
    End If
End Function


Private Sub IncluiEmpenho(strAno As String, _
                          strProgTrabalhoPKID As String, _
                          strCredorPKID As String, _
                          strValor As Double, _
                          STRTIPO As String, _
                          strSaldoDotacao As String, _
                          TotalDotado As String, _
                          intEvento As String, _
                          INTProjAtiv As String, _
                          intCatEconom As String, _
                          intFonteRec As String)

    Dim strSql            As String
    Dim adoResultado      As ADODB.Recordset
    Dim strEmpenhoNumero  As String
    Dim strEmpenhoPKID    As String
    Dim cont              As Integer

tentaGravar:
    
    strEmpenhoNumero = CStr(GeraProximoDeEmpenho)
    If strEmpenhoIF(0) = "" Then
        strEmpenhoIF(0) = strEmpenhoNumero
    End If
    strEmpenhoIF(1) = strEmpenhoNumero
    strSql = ""
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    strSql = strSql & "INSERT INTO " & gstrEmpenho & " ("
    strSql = strSql & "intNumero, dtmData, intProgramaTrabalho, dblValor, "
    strSql = strSql & "intTipo,intItemDespesa, intCredor, "
    strSql = strSql & "strHistorico, dtmDtAtualizacao, intevento ,lngCodUsr, "
    strSql = strSql & "strCodigo, bitDigito, intExercicioEmpenho) "
    strSql = strSql & " VALUES (" & strEmpenhoNumero & ", "
    strSql = strSql & gstrConvDtParaSql(txtdtmData) & ", "
    strSql = strSql & strProgTrabalhoPKID & ", "
    strSql = strSql & gstrConvVrParaSql(strValor) & ", "
    strSql = strSql & gstrItemData(dbcintTipo) & ", "
   
    strSql = strSql & IIf(Trim(dbcintItemDespesa.Text) <> "", gstrItemData(dbcintItemDespesa, True), "NULL") & ", "
    
    strSql = strSql & strCredorPKID & ", "
    strSql = strSql & "'" & txtstrHistorico & "', "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSql = strSql & intEvento & ", "
    strSql = strSql & glngCodUsr & ", '"
    strSql = strSql & txtstrCodigo & "', " & txtbitDigito & ", " & txtintExercicio & " );"
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
    If Not gobjBanco.Execute(strSql, True) Then
        cont = cont + 1
        If cont = 30 Then
            ExibeMensagem "Ocorreu um erro ao gravar um empenho." & vbNewLine & "A importação foi cancelada, nenhum registro foi gravado."
            gobjBanco.ExecutaRollbackTrans
            mblnRollBack = True
            Exit Sub
        Else
            GoTo tentaGravar
        End If
        Exit Sub
    Else
        strEmpenhoPKID = gstrEmpenhoPkid(strEmpenhoNumero)
    End If
    
    strSql = ""
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    strSql = strSql & " UPDATE " & gstrSubempenho
    strSql = strSql & " SET dblEmpenhadoAteData = " & gstrConvVrParaSql(CDbl(TotalDotado) + strValor)
    strSql = strSql & ", dblSaldoAtual = " & gstrConvVrParaSql(CDbl(strSaldoDotacao) - strValor)
    strSql = strSql & ",dtmLiquidacao = "
    strSql = strSql & gstrConvDtParaSql(txtdtmData) & ", "
    strSql = strSql & "bytSituacao = 2, " '2= LIQUIDADA
    strSql = strSql & "strHistorico = '"
    strSql = strSql & txtstrHistorico & "', "
    strSql = strSql & "dtmDtAtualizacao = "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
    strSql = strSql & "intevento = " & gstrItemData(cbo_intEventoLiquidacao) & ", "
    strSql = strSql & "lngCodUsr = " & glngCodUsr & " "
    strSql = strSql & " WHERE intNumero = 0 AND "
    strSql = strSql & " intEmpenho = (SELECT PKID FROM " & gstrEmpenho & " WHERE intNumero=" & strEmpenhoNumero & " AND " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " );"
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
        
    Set gobjBanco = New clsBanco
    If Not gobjBanco.Execute(strSql) Then
        ExibeMensagem "Ocorreu um erro ao gravar um empenho." & vbNewLine & "A importação foi cancelada, nenhum registro foi gravado."
        gobjBanco.ExecutaRollbackTrans
        mblnRollBack = True
        Exit Sub
    Else
        If GravaNotasFiscais(CStr(strValor), strEmpenhoPKID) Then
            'empenho
            If Not GeraMovimentosByEvento(CInt(intEvento), txtdtmData, Str(CDbl(strValor)), "", strEmpenhoNumero, "3", , , , , True) Then
               ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil." & _
                             vbNewLine & "A importação foi cancelada, nenhum registro foi gravado."
               gobjBanco.ExecutaRollbackTrans
               mblnRollBack = True
               Exit Sub
            Else
                'subempenhoLiquidado
                If Not GeraMovimentosByEvento(gstrItemData(cbo_intEventoLiquidacao), txtdtmData, Str(CDbl(strValor)), "", strEmpenhoNumero, "3", , , , , True) Then
                      ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil." & _
                                    vbNewLine & "A importação foi cancelada, nenhum registro foi gravado."
                      gobjBanco.ExecutaRollbackTrans
                      mblnRollBack = True
                      Exit Sub
                End If
                
                GravaExtras strEmpenhoPKID, strAno, INTProjAtiv, intCatEconom, strCredorPKID, intFonteRec, intEvento, STRTIPO
                
                If Not blnGravaMovLiq(strEmpenhoPKID, RetPkidSubEmpenho(strEmpenhoPKID), strProgTrabalhoPKID, txtdtmData, strValor, txtstrHistorico) Then
                   ExibeMensagem "Ocorreram erros durante a gravação dos Movimentos de Liquidação."
                   gobjBanco.ExecutaRollbackTrans
                   mblnRollBack = True
                End If
                
                
            End If
        End If
    End If
    
End Sub

Private Function RetPkidSubEmpenho(ByVal strEmpenhoPKID As String) As String
    Dim strSql As String
    Dim adoResultado As Recordset
    
    strSql = "SELECT MAX(PKID) PkidSubEmpenho FROM "
    strSql = strSql & gstrSubempenho
    strSql = strSql & " WHERE intEmpenho=" & strEmpenhoPKID

   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                RetPkidSubEmpenho = !PkidSubEmpenho
            End If
        End With
   End If

End Function


Private Function blnGravaMovLiq(ByVal PKIDEmpenho As String, ByVal pkidParcela As String, ByVal pkidProgramaDeTrabalho As String, ByVal DTMDATA As String, ByVal dblValor As String, ByVal STRHISTORICO As String) As Boolean
   Dim strSql As String
   
   strSql = "INSERT INTO " & gstrmovliq
   strSql = strSql & " ( intEmpenho, intParcela, intProgramaTrabalho, dtmData, dblValor, strHistorico, dtmDtAtualizacao, lngCodUsr) VALUES "
   strSql = strSql & "(" & PKIDEmpenho & ", " & pkidParcela & ", " & pkidProgramaDeTrabalho & ", " & gstrConvDtParaSql(DTMDATA) & ", "
   strSql = strSql & gstrConvVrParaSql(dblValor) & ",'" & STRHISTORICO & "', " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & ")"
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.Execute(strSql) Then
        blnGravaMovLiq = True
   Else
        blnGravaMovLiq = False
   End If

   
End Function

Private Function GeraProximoDeEmpenho() As Long
   Dim strSql       As String
   Dim adoResultado As New ADODB.Recordset
   
   strSql = "SELECT " & gstrISNULL("MAX(intNumero)", "0") & " AS Codigo FROM " & gstrEmpenho & " "
   strSql = strSql & " WHERE " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio
   'strSql = strSql & " and dtmData < " & gstrConvDtParaSql("20/01/2004") & " "
   strSql = strSql & " UNION SELECT " & gstrISNULL("MAX(intEmpenhoAnulacao)", "0") & " AS Codigo FROM " & gstrSubempenho & " "
   strSql = strSql & " WHERE " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " "
   'strSql = strSql & " and dtmData < " & gstrConvDtParaSql("20/01/2004") & " "
   strSql = strSql & " ORDER BY Codigo DESC "
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            GeraProximoDeEmpenho = !Codigo
         End If
      End With
   End If
   
   GeraProximoDeEmpenho = GeraProximoDeEmpenho + 1
   
End Function


Private Function gstrEmpenhoPkid(ByVal numeroEmpenho As String) As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset

    strSql = ""
    strSql = strSql & "SELECT PKID FROM " & gstrEmpenho
    strSql = strSql & " WHERE intNumero =" & numeroEmpenho
    strSql = strSql & " AND " & gstrDATEPART(strYEAR, "dtmdata") & " = " & gintExercicio

   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            gstrEmpenhoPkid = CStr(!Pkid)
         End If
      End With
   End If '

End Function

 Private Function GravaNotasFiscais(ByVal strValor As String, ByVal strEmpenhoPKID As String) As Boolean
    Dim strSql  As String
    Dim intInd  As Integer
    
'    strSql = ""
'    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
'    strSql = strSql & "INSERT INTO " & gstrSubEmpenhoNF & " ("
'    strSql = strSql & "intSubEmpenho, dtmData, dblValorNF, "
'    strSql = strSql & "strNotaFiscal, dtmDtAtualizacao, lngCodUsr) VALUES "
'    strSql = strSql & "((SELECT MAX(PKID) FROM " & gstrSubempenho & " WHERE intEmpenho=" & strEmpenhoPKID & ") , "
'    strSql = strSql & gstrConvDtParaSql(txtdtmData.Text) & ", "
'    strSql = strSql & gstrConvVrParaSql(strValor) & ", '"
'    strSql = strSql & "SEM NOTA', "
'    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & " "
'    strSql = strSql & "); "
'    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
    strSql = ""
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    strSql = strSql & "INSERT INTO " & gstrSubEmpenhoNF & " ("
    strSql = strSql & "intSubEmpenho, dtmData, dblValorNF, "
    strSql = strSql & "strNotaFiscal, dtmDtAtualizacao, lngCodUsr)"
    strSql = strSql & "(SELECT MAX(PKID),"
    strSql = strSql & gstrConvDtParaSql(txtdtmData.Text) & ", "
    strSql = strSql & gstrConvVrParaSql(strValor) & ", '"
    strSql = strSql & "SEM NOTA', "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & " "
    strSql = strSql & " FROM " & gstrSubempenho & " WHERE intEmpenho=" & strEmpenhoPKID
    strSql = strSql & "); "
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
    If strSql <> "" Then
        Set gobjBanco = New clsBanco
        If Not gobjBanco.Execute(strSql) Then
           ExibeMensagem "Problemas durante a gravação das Notas Fiscais.Entre em contato com o fornecedor." & _
           vbNewLine & "A importação foi cancelada, nenhum registro foi gravado."
           gobjBanco.ExecutaRollbackTrans
           mblnRollBack = True
           Exit Function
        End If
    End If
    GravaNotasFiscais = True
End Function


 Private Function GravaExtras(ByVal strEmpenhoPKID As String, _
                              ByVal intAno As String, _
                              ByVal INTPROJATIVIDADE As String, _
                              ByVal INTCATEGORIAECON As String, _
                              ByVal intCredor As String, _
                              ByVal INTFONTERECURSO As String, _
                              ByVal intEvento As String, _
                              ByVal STRTIPO As String) As Boolean
    

    Dim strSql  As String
    Dim adoResultado  As ADODB.Recordset
    Dim strContaCruzada As String
    Dim intContadorInterno As Integer
    Dim strCodAtual As String
    Dim strCodAntigo As String
    Dim strCodCorrente As String
    
    strCodCorrente = intAno & INTPROJATIVIDADE & INTCATEGORIAECON & LeCDCCredor(intCredor) & INTFONTERECURSO
    
    intContadorInterno = 0
    strSql = ""
    strSql = strSql & "SELECT IPF.* "
    strSql = strSql & "FROM " & gstrImpressaoFolha & " IPF  Where blnImpressao = 0  ORDER BY intID "
    
    
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
        
        Do While Not adoResultado.EOF
            If intContador = adoResultado!intID Then
                Exit Do
            End If
            adoResultado.MoveNext
        Loop
        
        While Not adoResultado.EOF
            
            strCodAtual = adoResultado!intAno & adoResultado!INTPROJATIVIDADE & adoResultado!INTCATEGORIAECON & adoResultado!intCredor & adoResultado!STRFONTERECURSO
            strContaCruzada = retornaContaCruzada(adoResultado!intCLC)
                    
            If strContaCruzada = "0" Then
               ExibeMensagem "Problemas durante a gravação de Descontos Extra-Orçamentários." & _
               vbNewLine & "O CLC " & adoResultado!intCLC & " ainda não foi cruzado e a gravação não pode continuar sem esta informação." & _
               vbNewLine & "Nenhum registro foi gravado."
               gobjBanco.ExecutaRollbackTrans
               mblnRollBack = True
               Exit Function
'            ElseIf strContaCruzada = "-1" Then
'                intContadorInterno = intContadorInterno + 1
'                If intContadorInterno > 1 And strCodAntigo <> strCodAtual Then
'                    intContadorInterno = 0
'                    Exit Function
'                End If
'                GoTo proximo
            End If
                    
            If strCodCorrente = strCodAtual Then
                If strContaCruzada = "-1" Then
                    GoTo proximo
                End If
            Else
    '            intContador = adoResultado!intID + 1
                Exit Function
            End If
                    
            If strContaCruzada = "-2" Then
                strSql = "UPDATE  " & gstrSubempenho
                strSql = strSql & " SET DblDesconto = dblDesconto +  " & gstrConvVrParaSql(adoResultado!dblValor)
                strSql = strSql & " WHERE PKID = (SELECT MAX(PKID) FROM " & gstrSubempenho & " WHERE intEmpenho=" & strEmpenhoPKID & ")"
            Else
                strSql = "INSERT INTO " & gstrSubempenhoLiquidado & " ("
                strSql = strSql & "intParcela, intConta, dblValor, bytTipo, "
                strSql = strSql & "dtmDtAtualizacao, lngCodUsr "
                strSql = strSql & ")"
                strSql = strSql & "(SELECT MAX(PKID),"
                strSql = strSql & strContaCruzada & ", "
                strSql = strSql & gstrConvVrParaSql(adoResultado!dblValor) & ", "
                strSql = strSql & "1 , "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
                strSql = strSql & glngCodUsr
                strSql = strSql & " FROM " & gstrSubempenho & " WHERE intEmpenho=" & strEmpenhoPKID & ")"
            End If
              
            Set gobjBanco = New clsBanco
            If Not gobjBanco.Execute(strSql) Then
               ExibeMensagem "Problemas durante a gravação de Descontos Extra-Orçamentários.Entre em contato com o fornecedor." & _
               vbNewLine & "A importação foi cancelada, nenhum registro foi gravado."
               gobjBanco.ExecutaRollbackTrans
               mblnRollBack = True
               Exit Function
            End If
proximo:
            intContador = adoResultado!intID + 1
            strCodAntigo = adoResultado!intAno & adoResultado!INTPROJATIVIDADE & adoResultado!INTCATEGORIAECON & adoResultado!intCredor & adoResultado!STRFONTERECURSO
            adoResultado.MoveNext
        Wend
    End If
    
    GravaExtras = True
End Function


 Private Function strQueryEmpenho() As String
    Dim strSql  As String
    Dim adoResultado As New ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "EC.Pkid, "
    strSql = strSql & "EC.INTPEDIDOEMPENHO "
    strSql = strSql & "From "
    strSql = strSql & gstrEmpenhoContrato & " EC, "
    strSql = strSql & gstrRequisicaoCompras & " RC "
    strSql = strSql & " Where "
    strSql = strSql & "EC.intrequisicaodecompra = RC.Pkid AND "
    strSql = strSql & "RC.Strnumeroempenho is null "
    strSql = strSql & "Order by EC.INTPEDIDOEMPENHO"
    strQueryEmpenho = strSql
End Function

Private Sub cmd_importCompras_Click()
Dim strSql As String
Dim adoResultado    As ADODB.Recordset
Dim intInd          As Integer
Dim intContElemento As Integer
    
    If Not cbo_intAutorizacaoDeCompra.MatchedWithList Then
        ExibeMensagem "É necessário selecionar um Pedido de Empenho válido."
        cbo_intAutorizacaoDeCompra.SetFocus
        Exit Sub
    End If
    
    
    
    frmCadEmpenho.tab_3dPasta.Tab = 0
    
    frmCadEmpenho.MantemForm (gstrNovo)
   
    'strSql = "SELECT EC.intRequisicaoDeCompra,"
    'strSql = strSql & " (SELECT SUM(dblValor) FROM "
    'strSql = strSql & gstrEmpenhoContrato & " WHERE"
    'strSql = strSql & " intAutorizacaoDecompra = EC.Intautorizacaodecompra ) dblValor,"
    strSql = "SELECT "
    strSql = strSql & " MAX(RC.dblValorParaEmpenho) as dblValor, "
    strSql = strSql & " EC.intReserva,"
    strSql = strSql & " EC.intFornecedor,"
    strSql = strSql & " EC.intModalidade,"
    strSql = strSql & " EC.intNumeroMod,"
    strSql = strSql & " EC.intAnoMod,"
    strSql = strSql & " EC.strNumeroProcesso,"
    strSql = strSql & " EC.strObjetoAutorizacao,"
    strSql = strSql & " EC.dtmDtHomologacao,"
    strSql = strSql & " EC.intAutorizacaoDeCompra,"
    strSql = strSql & " EC.intPedidoEmpenho,"
    strSql = strSql & " EC.bitDigitoProcesso,"
    strSql = strSql & " EC.intExercicioProcesso,"
    strSql = strSql & " LE.strDescricao strLocalEntrega,"
    strSql = strSql & " EC.strprazoentrega,"
    strSql = strSql & " EC.strcondicaopagamento"
    strSql = strSql & " FROM "
    strSql = strSql & gstrEmpenhoContrato & " EC, "
    strSql = strSql & gstrRequisicaoCompras & " RC, "
    strSql = strSql & gstrLocalEntrega & " LE"
    
    'strSQL = strSQL & " WHERE  RC.intPedidoEmpenho ='" & cbo_intAutorizacaoDeCompra.BoundText & "' AND "
    strSql = strSql & " WHERE  EC.intPedidoEmpenho ='" & cbo_intAutorizacaoDeCompra.BoundText & "' AND " '-Linha adicionada
    strSql = strSql & "  EC.intRequisicaodeCompra = RC.PKID AND " '-Linha adicionada
    strSql = strSql & "RC.Intlocalentrega = LE.Pkid AND "
    strSql = strSql & " RC.Pkid = EC.intRequisicaoDeCompra   "
    strSql = strSql & "Group By "
    strSql = strSql & "EC.intReserva,"
    strSql = strSql & "EC.intFornecedor,"
    strSql = strSql & "EC.intModalidade,"
    strSql = strSql & "EC.intNumeroMod,"
    strSql = strSql & "EC.intAnoMod,"
    strSql = strSql & "EC.strNumeroProcesso,"
    strSql = strSql & "EC.strObjetoAutorizacao,"
    strSql = strSql & "EC.dtmDtHomologacao,"
    strSql = strSql & "EC.intAutorizacaoDeCompra,"
    strSql = strSql & "EC.intPedidoEmpenho,"
    strSql = strSql & "EC.bitDigitoProcesso,"
    strSql = strSql & "EC.intExercicioProcesso ,"
    strSql = strSql & "LE.strDescricao, "
    strSql = strSql & "EC.strprazoentrega, "
    strSql = strSql & "EC.strcondicaopagamento"
    
    frmCadEmpenho.intNumPedidoEmpenho = cbo_intAutorizacaoDeCompra.BoundText
    frmCadEmpenho.cbointReservaDotacao.Clear
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                DoEvents
                frmCadEmpenho.txtdtmData = gstrENulo(!dtmDtHomologacao)
                If Not IsNull(!intReserva) Then
                
                    frmCadEmpenho.cbo_intEvento.Clear
'                    Do While frmCadEmpenho.cbo_intEvento.ListCount = 0 And intContElemento < 15
'
'
'                       strSql = "SELECT E.Pkid,"
'                       strSql = strSql & " E.strDescricao"
'                       strSql = strSql & " FROM "
'                       strSql = strSql & gstrEvento & " E"
'                       strSql = strSql & " WHERE"
'                       strSql = strSql & " E.Pkid in (SELECT"
'                       strSql = strSql & " ECD.intEvento"
'                       strSql = strSql & " FROM "
'                       strSql = strSql & gstrEventoContaContabilDebito & " ECD "
'                       strSql = strSql & " WHERE"
'                       strSql = strSql & " ECD.intContaContabil = (SELECT"
'                       strSql = strSql & " PC.Pkid"
'                       strSql = strSql & " FROM "
'                       strSql = strSql & gstrPlanoConta & " PC"
'                       strSql = strSql & " WHERE"
'                       strSql = strSql & "(" & gstrTOPnOracle("SELECT " & gstrTOPnSQLServer(1) & "'" & gstrDigitoDespesa & "'" & _
'                               strCONCAT & strSUBSTRING & "(ED.strCodigoElementoDespesa,1," & CStr(15 - intContElemento) & ") strCodigo FROM " & _
'                               gstrReservaDotacao & " RD, " & _
'                               gstrProgramaDeTrabalho & " PT, " & _
'                               gstrElementoDespesa & " ED, " & _
'                               gstrEmpenhoContrato & " EC" & _
'                               " WHERE RD.PKID = EC.intReserva AND PT.pkid = RD.intProgramaTrabalho AND" & _
'                               " ED.Pkid = PT.intElementoDespesa AND EC.intReserva = " & gstrENulo(!intReserva) & _
'                               " ORDER BY ED.strCodigoElementoDespesa", 1)
'                       strSql = strSql & ") "
'                       strSql = strSql & " = "
'                       strSql = strSql & strSUBSTRING & "(PC.strContaContabil,1," & Len(gstrDigitoDespesa) + 15 - intContElemento & ")"
'                       strSql = strSql & " )"
'                       strSql = strSql & " )"
'                       intContElemento = intContElemento + 1
'
'                       'Preenche Evento Contabil
'
'                       LeDaTabelaParaObj "", frmCadEmpenho.cbo_intEvento, strSql
'
'                       If frmCadEmpenho.cbo_intEvento.ListCount <> 0 Then Exit Do
'                    Loop
'
'                   If frmCadEmpenho.cbo_intEvento.ListCount = 0 Then
'                       ExibeMensagem " Esse pedido de Empenho não tem Evento Contabil."
'                       frmCadEmpenho.MantemForm (gstrNovo)
'                       Unload Me
'                       Exit Sub
'                   End If
'
'                   frmCadEmpenho.cbo_intEvento.ListIndex = 0
                   frmCadEmpenho.cbo_intEvento.Clear
                   'Desabilita controles Evento Contabil
                   TrocaCorObjeto frmCadEmpenho.cmd_Evento, True
                   TrocaCorObjeto frmCadEmpenho.cbo_intEvento, True
                   TrocaCorObjeto frmCadEmpenho.txt_codEvento, True
                   
                   'Preenche Reseva
                   strSql = "SELECT Pkid,"
                   strSql = strSql & " intNumero"
                   strSql = strSql & " FROM "
                   strSql = strSql & gstrReservaDotacao
                   strSql = strSql & " WHERE "
                   strSql = strSql & " Pkid = " & gstrENulo(!intReserva)
                   LeDaTabelaParaObj "", frmCadEmpenho.cbointReservaDotacao, strSql
                   frmCadEmpenho.cbointReservaDotacao.ListIndex = 0
                   'Desabilita controles Dotação
                   TrocaCorObjeto frmCadEmpenho.cboCodigoReduzido, True
                   TrocaCorObjeto frmCadEmpenho.cboProgramaTrabalho, True
                   TrocaCorObjeto frmCadEmpenho.cmd_ProgramaTrabalho, True
                Else
                    ExibeMensagem " Esse pedido de Empenho não tem Reserva de Dotação"
                End If
                
                frmCadEmpenho.txtstrsolicitacao = cbo_intAutorizacaoDeCompra.BoundText
                TrocaCorObjeto frmCadEmpenho.txtstrsolicitacao, True
                
                'Preenche e desabilita o Valor do Empenho
                frmCadEmpenho.txtdblValor = gstrConvVrDoSql(!dblValor)
                TrocaCorObjeto frmCadEmpenho.txtdblValor, True
                
                'Desabilita controles Reserva
                TrocaCorObjeto frmCadEmpenho.cbointReservaDotacao, True
                
                TrocaCorObjeto frmCadEmpenho.cmd_Reserva, True
                
                frmCadEmpenho.txtdtmHomologacao = gstrENulo(!dtmDtHomologacao)
                TrocaCorObjeto frmCadEmpenho.txtdtmHomologacao, True
                                
                'Preenche e desabilita o Processo
                frmCadEmpenho.txtstrCodigo = gstrENulo(!strNumeroProcesso)
                TrocaCorObjeto frmCadEmpenho.txtstrCodigo, True
                frmCadEmpenho.txtintExercicio = gstrENulo(!intExercicioProcesso)
                TrocaCorObjeto frmCadEmpenho.txtintExercicio, True
                frmCadEmpenho.txtbitDigito = gstrENulo(!bitDigitoProcesso)
                TrocaCorObjeto frmCadEmpenho.txtbitDigito, True
                
                'Preenche e desabilita o Credor
                strSql = "SELECT Pkid, strNome"
                strSql = strSql & " FROM "
                strSql = strSql & gstrContribuinte
                strSql = strSql & " WHERE Pkid = '" & gstrENulo(!intFornecedor) & "'"
                LeDaTabelaParaObj "", frmCadEmpenho.dbcintCredor, strSql
                'frmCadEmpenho.dbcintCredor = 0
                frmCadEmpenho.dbcintCredor.BoundText = gstrENulo(!intFornecedor)
                
                'frmCadEmpenho.txt_intNContribuinte.Text = gstrENulo(!intFornecedor)
                
                TrocaCorObjeto frmCadEmpenho.dbcintCredor, True
                TrocaCorObjeto frmCadEmpenho.txt_intNContribuinte, True
                TrocaCorObjeto frmCadEmpenho.cmd_Credor, True
                frmCadEmpenho.txt_intNContribuinte = LeCDCCredor(gstrENulo(!intFornecedor))
                
                'Preenche e Historico
                frmCadEmpenho.txtstrHistorico = gstrENulo(!strObjetoAutorizacao)
                
                'Preenche e desabilita o Condições de pgto
                frmCadEmpenho.txtStrcondpagto = gstrENulo(!strcondicaopagamento)
                TrocaCorObjeto frmCadEmpenho.txtStrcondpagto, True
                frmCadEmpenho.txtStrprazoentrega = gstrENulo(!Strprazoentrega)
                TrocaCorObjeto frmCadEmpenho.txtStrprazoentrega, True
                frmCadEmpenho.txtStrlocentrega = gstrENulo(!strLocalEntrega)
                TrocaCorObjeto frmCadEmpenho.txtStrlocentrega, True

                'Preenche e desabilita Modalidade
                strSql = "SELECT Pkid, strCodigo"
                strSql = strSql & " FROM "
                strSql = strSql & gstrComprasLicitacao
                strSql = strSql & " WHERE Pkid = '" & gstrENulo(!intModalidade) & "'"
                LeDaTabelaParaObj "", frmCadEmpenho.dbcintModalidade, strSql
                frmCadEmpenho.dbcintModalidade.BoundText = gstrENulo(!intModalidade)
                TrocaCorObjeto frmCadEmpenho.dbcintModalidade, True
                frmCadEmpenho.txtstrModalidade = gstrENulo(!intNumeroMod) & "/" & gstrENulo(!intAnoMod)
                TrocaCorObjeto frmCadEmpenho.txtstrModalidade, True
                
          End If
       End With
     End If
     
'Select para preenchimento do grid de itens da tela de Empenho
    
    'Primeiro vamos verificar se o Valor para Empenho é o mesmo que o total de itens, senão os itens não são importados
    
    strSql = "Select Sum(dblValorUnitario * dblQuantidade) dblValorTotal "
    strSql = strSql & " From "
    strSql = strSql & gstrEmpenhoContratoItens & " ECI, "
    strSql = strSql & gstrEmpenhoContrato & " EC"
    strSql = strSql & " WHERE "
    strSql = strSql & " ECI.intEmpenhoContrato = EC.Pkid  AND"
    strSql = strSql & " EC.intPedidoEmpenho = " & cbo_intAutorizacaoDeCompra.BoundText
    strSql = strSql & " GROUP BY intPedidoEmpenho "
    
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then

        If Format(adoResultado!dblValorTotal, "#,##0.00") <> Format(CDbl(gstrConvVrDoSql(frmCadEmpenho.txtdblValor)), "#,##0.00") Then
            ExibeMensagem "O Valor Total de itens do Pedido de Empenho difere do Valor para Empenho cadastrado. A importação foi concluída sem a importação dos itens."
            frmCadEmpenho.blnImportadoPedidoEmpenho = True
        Else
                
            strSql = "SELECT CM.Pkid,"
            strSql = strSql & " CM.intCodigo Codigo,"
            strSql = strSql & " CM.strDescricao Descricao,"
            strSql = strSql & " MA.strMarca Marca,"
            strSql = strSql & " ECI.dblQuantidade Quantidade,"
            strSql = strSql & " ECI.dblValorUnitario ValorUnitario,"
            strSql = strSql & " UM.strDescricao UnidMedida,"
            strSql = strSql & " ECI.strObsCompra Observacao,"
            strSql = strSql & " CM.strDescricaoDetalhada DescDetalhada"
            strSql = strSql & " FROM "
            strSql = strSql & gstrEmpenhoContrato & " EC, "
            strSql = strSql & gstrEmpenhoContratoItens & " ECI, "
            'strSQL = strSQL & gstrRequisicaoCompras & " RC, "
            strSql = strSql & gstrCatalogoMaterialServico & " CM, "
            strSql = strSql & gstrMarcas & " MA, "
            strSql = strSql & gstrUnidadeMedida & " UM"
            strSql = strSql & " WHERE"
            strSql = strSql & " CM.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " ECI.intCatalogoMaterialServico AND"
            strSql = strSql & " MA.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " ECI.intMarca AND"
            strSql = strSql & " UM.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " ECI.intUnidadeDeMedida AND"
            strSql = strSql & " ECI.intEmpenhoContrato = EC.Pkid  AND"
            'strSQL = strSQL & " EC.intRequisicaoDeCompra = RC.Pkid AND"
            'strSQL = strSQL & " RC.intPedidoEmpenho ='" & cbo_intAutorizacaoDeCompra.BoundText & "'"
            strSql = strSql & " EC.intPedidoEmpenho = " & cbo_intAutorizacaoDeCompra.BoundText  '-Linha adicionada
             
            frmCadEmpenho.lvw_Itens.ListItems.Clear
            
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
                With adoResultado
                    If Not .EOF Then
                        .MoveFirst
                        Do While Not .EOF
                            Set mobjLista = frmCadEmpenho.lvw_Itens.ListItems.Add(, , !Pkid)
                            mobjLista.SubItems(1) = gstrENulo(!Codigo) '= txt_intCodigo.Text
                            mobjLista.SubItems(2) = gstrENulo(!descricao) '= txt_intCatalogoMaterialServico.Text
                            mobjLista.SubItems(3) = gstrENulo(!Marca) '= dbc_intStrMarca
                            mobjLista.SubItems(4) = gstrENulo(!Quantidade) '= txt_dblQuantidade.Text
                            mobjLista.SubItems(5) = gstrENulo(!ValorUnitario) '= gstrConvVrDoSql(txt_dblValorEstimado.Text, 2)
                            mobjLista.SubItems(6) = gstrENulo(!UnidMedida) '= txt_intUnidadedeMedida.Text
                            'mobjLista.SubItems(7) = gstrENulo(!Observacao) '= txt_strObsItem.Text
                            mobjLista.SubItems(8) = gstrENulo(!Observacao) & vbNewLine & gstrENulo(!DescDetalhada) '= txt_strdescricaodetalhada.Text
                            .MoveNext
                        Loop
                        TrocaCorObjeto frmCadEmpenho.txt_intCodigo, True
                        TrocaCorObjeto frmCadEmpenho.txt_intCatalogoMaterialServico, True
                        TrocaCorObjeto frmCadEmpenho.txt_dblQuantidade, True
                        TrocaCorObjeto frmCadEmpenho.txt_dblValorEstimado, True
                        TrocaCorObjeto frmCadEmpenho.txt_intUnidadedeMedida, True
                        TrocaCorObjeto frmCadEmpenho.txt_strObsItem, True
                        TrocaCorObjeto frmCadEmpenho.txt_strdescricaodetalhada, True
                        TrocaCorObjeto frmCadEmpenho.dbc_intStrMarca, True
                        
                        frmCadEmpenho.blnImportadoPedidoEmpenho = True
                        
                    Else
                        ExibeMensagem "Não foi encontrado itens nessa Solicitação de Empenho"
                        frmCadEmpenho.MantemForm (gstrNovo)
                    End If
                End With
            End If
        End If
    Else
        ExibeMensagem "Ocorreram erros na importação dos itens do Pedido de Empenho. Os itens não foram importados."
    End If
     
    Unload Me
     
End Sub

Private Function strQueryAgrupaFolha() As String
    Dim strSql As String
    
'    strSQL = "SELECT "
'    strSQL = strSQL & " intAno,"
'    strSQL = strSQL & " INTPROJATIVIDADE,"
'    strSQL = strSQL & " INTCATEGORIAECON,"
'    strSQL = strSQL & " INTCREDOR,"
'    strSQL = strSQL & " INTFONTERECURSO,"
'    strSQL = strSQL & " SUM(DBLVALOR) DBLVALOR,"
'    strSQL = strSQL & " intEvento"
'    strSQL = strSQL & "  FROM "
'    strSQL = strSQL & gstrImpressaoFolha
'    strSQL = strSQL & " GROUP BY"
'    strSQL = strSQL & " intAno,"
'    strSQL = strSQL & " INTPROJATIVIDADE,"
'    strSQL = strSQL & " INTCATEGORIAECON,"
'    strSQL = strSQL & " INTCREDOR,"
'    strSQL = strSQL & " INTFONTERECURSO,"
'    strSQL = strSQL & " intEvento"


    strSql = "SELECT * FROM " & gstrImpressaoFolha & " WHERE blnImpressao = 0  ORDER BY intID "
    strQueryAgrupaFolha = strSql
End Function


Private Function retornaContaCruzada(ByVal strCLC As String) As String
    Dim strSql As String
    Dim adoResultado  As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "SELECT CC.*"
    strSql = strSql & "FROM " & gstrCruzamentoContaExtra & " CC "
    strSql = strSql & " WHERE CC.intCLC = " & strCLC
    
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
        If Not adoResultado.EOF Then
            If adoResultado!blnSemLigacao = 1 Then
                retornaContaCruzada = "-1"
                Exit Function
            End If
            If adoResultado!blnDescontoExtra = 1 Then
                retornaContaCruzada = "-2"
                Exit Function
            End If
            retornaContaCruzada = CStr(adoResultado!intPlanoConta)
        Else
            retornaContaCruzada = "0"
        End If
    End If
End Function
