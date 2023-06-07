VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmLanCobAmig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Cobrança Amigável"
   ClientHeight    =   4890
   ClientLeft      =   2790
   ClientTop       =   3075
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9540
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cobrança Amigável"
      TabPicture(0)   =   "frmLanCobAmig.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLote"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblVencimento"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblIndexador(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblValIndexador"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_Status"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "prg_Status"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbc_intIndexador"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_intLote"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_dblValorMoeda"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtVencimento"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chk_ParcelasAcordo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chk_Simulado"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Fra_CompReceita"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.Frame Fra_CompReceita 
         Caption         =   "Composição da Receita"
         Height          =   2565
         Left            =   210
         TabIndex        =   15
         Top             =   1800
         Width           =   8955
         Begin VB.CommandButton Command1 
            Height          =   300
            Left            =   4980
            Picture         =   "frmLanCobAmig.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro de Composição da Receita"
            Top             =   240
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbc_intComposicao 
            Height          =   315
            Left            =   1170
            TabIndex        =   7
            Top             =   240
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Composicoes 
            Height          =   1725
            Left            =   210
            TabIndex        =   9
            Top             =   690
            Width           =   8550
            _ExtentX        =   15081
            _ExtentY        =   3043
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "intComposicao"
            Columns(0).DataField=   "intComposicao"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Composição da Receita"
            Columns(1).DataField=   "strComposicao"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Exercício"
            Columns(2).DataField=   "intExercicio"
            Columns(2).NumberFormat=   "FormatText Event"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Posição"
            Columns(3).DataField=   "intPosicao"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=13097"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=13018"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=1799"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=1720"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=8194"
            Splits(0)._ColumnProps(20)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(22)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(3).AllowSizing=0"
            Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DataMode        =   4
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTips        =   1
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.wraptext=-1,.locked=0,.bold=0"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000002&"
            _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000014&"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.bgcolor=&H80000005&"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.bgcolor=&HC0C0C0&,.locked=-1"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17,.bgcolor=&H80000016&"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HC0C0C0&"
            _StyleDefs(46)  =   ":id=46,.locked=-1"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(54)  =   "Named:id=33:Normal"
            _StyleDefs(55)  =   ":id=33,.parent=0,.transparentBmp=0"
            _StyleDefs(56)  =   "Named:id=34:Heading"
            _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(58)  =   ":id=34,.wraptext=-1"
            _StyleDefs(59)  =   "Named:id=35:Footing"
            _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(61)  =   "Named:id=36:Selected"
            _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(63)  =   "Named:id=37:Caption"
            _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(65)  =   "Named:id=38:HighlightRow"
            _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(67)  =   "Named:id=39:EvenRow"
            _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(69)  =   "Named:id=40:OddRow"
            _StyleDefs(70)  =   ":id=40,.parent=33"
            _StyleDefs(71)  =   "Named:id=41:RecordSelector"
            _StyleDefs(72)  =   ":id=41,.parent=34"
            _StyleDefs(73)  =   "Named:id=42:FilterBar"
            _StyleDefs(74)  =   ":id=42,.parent=33"
         End
         Begin MSDataListLib.DataCombo dbc_intExercicio 
            Height          =   315
            Left            =   7755
            TabIndex        =   8
            Top             =   240
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lblExercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   6960
            TabIndex        =   17
            Top             =   285
            Width           =   675
         End
         Begin VB.Label lbl_Composicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   210
            TabIndex        =   16
            Top             =   285
            Width           =   870
         End
      End
      Begin VB.CheckBox chk_Simulado 
         Caption         =   "Simulado"
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   1350
         Width           =   1095
      End
      Begin VB.CheckBox chk_ParcelasAcordo 
         Caption         =   "Parcelas em Acordo"
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   1020
         Width           =   1785
      End
      Begin VB.TextBox txtVencimento 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   570
         Width           =   960
      End
      Begin VB.TextBox txt_dblValorMoeda 
         Height          =   315
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   4
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txt_intLote 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   630
         MaxLength       =   10
         TabIndex        =   1
         Top             =   570
         Width           =   870
      End
      Begin MSDataListLib.DataCombo dbc_intIndexador 
         Height          =   315
         Left            =   5070
         TabIndex        =   3
         Top             =   570
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComctlLib.ProgressBar prg_Status 
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   4560
         Visible         =   0   'False
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lbl_Status 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   4350
         Visible         =   0   'False
         Width           =   9165
      End
      Begin VB.Label lblValIndexador 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Indexador"
         Height          =   195
         Left            =   6690
         TabIndex        =   14
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lblIndexador 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   195
         Index           =   0
         Left            =   4290
         TabIndex        =   13
         Top             =   630
         Width           =   705
      End
      Begin VB.Label lblVencimento 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento"
         Height          =   195
         Left            =   2040
         TabIndex        =   12
         Top             =   630
         Width           =   840
      End
      Begin VB.Label lblLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   630
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmLanCobAmig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xadbComposicoes      As XArrayDB
Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
            
        Case Is = UCase(gstrIncluirItem)
            IncluiComposicaoNoGrid
            
        Case Is = UCase(gstrExcluirItem)
            ExcluiComposicaoNoGrid
                
        Case Is = UCase(gstrNovo)
            Limpa_Controles Me, True, True, False, True, False
            LimpaGrid
            RetornaLote False
            
        Case Is = UCase(gstrCalcularReajuste)
            If blnDadosOk Then
                RealizaCalculoCobrancaAmigavel
            End If
            
    End Select
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1370
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir, gstrSalvar, gstrDeletar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub

Private Sub Form_Load()
    dbc_intIndexador.Tag = "SELECT Pkid, strAbreviatura FROM tblIndexadorEconomico ORDER BY strAbreviatura;strAbreviatura"
    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"
    Set xadbComposicoes = New XArrayDB
    xadbComposicoes.Clear
    xadbComposicoes.ReDim 0, 0, 0, 3
    
    TrocaCorObjeto txt_intLote, True
    RetornaLote False
    
    'Vamos auto numerar
    'If Len(Trim(txt_intSeqInicial)) = 0 Then
    '   ProximaSequenciaLote False
    'End If
End Sub
Private Sub dbc_intComposicao_Change()
    
    LimpaDataCombo dbc_intExercicio
    
    If dbc_intComposicao.MatchedWithList Then
        dbc_intExercicio.Tag = strQueryExercicio & ";intExercicio"
        PreencherListaDeOpcoes dbc_intExercicio
    Else
        dbc_intExercicio.Tag = ""
    End If
    
End Sub

Private Sub dbc_intComposicao_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbc_intComposicao, Me, Area
End Sub

Private Sub dbc_intComposicao_GotFocus()
    MarcaCampo dbc_intComposicao
End Sub

Private Sub dbc_intComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicao
End Sub
Private Sub dbc_intIndexador_Click(Area As Integer)
    DropDownDataCombo dbc_intIndexador, Me, Area
End Sub

Private Sub dbc_intIndexador_GotFocus()
    MarcaCampo dbc_intIndexador
End Sub

Private Sub dbc_intIndexador_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intIndexador, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intIndexador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intIndexador
End Sub
Private Sub txt_intLote_GotFocus()
    MarcaCampo txt_intLote
End Sub

Private Sub txt_intLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intLote
End Sub
Private Function strQueryExercicio() As String
Dim strSql As String

    strSql = "SELECT Pkid, intExercicio "
    strSql = strSql & " FROM "
    strSql = strSql & gstrParametroAtualizacao
    strSql = strSql & " WHERE"
    strSql = strSql & " intComposicaoReceita = " & dbc_intComposicao.BoundText
    strSql = strSql & " ORDER BY intExercicio"

    strQueryExercicio = strSql

End Function
Private Function strQueryComposicao() As String
Dim strSql As String
    strSql = ""
    strSql = "SELECT CO.Pkid,"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CO.intCodigo") & strCONCAT & "' - '" & strCONCAT & " CO.strDescricao Descricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrParametroAtualizacao & " PA, " & gstrComposicaoDaReceita & " CO "
    strSql = strSql & " WHERE"
    strSql = strSql & " PA.intComposicaoReceita = CO.Pkid AND "
    strSql = strSql & " CO.bytDividaAtiva =  1 "
    strSql = strSql & " GROUP BY CO.Pkid, CO.intCodigo, CO.strDescricao"
    strSql = strSql & " ORDER BY CO.intCodigo"

    strQueryComposicao = strSql

End Function
Private Sub LimpaDataCombo(dbcAux As DataCombo)

    dbcAux.Tag = ""
    dbcAux.Text = ""
    dbcAux.ListField = ""
    
End Sub

Private Sub ExcluiComposicaoNoGrid()
Dim varAux As Variant
Dim intFor As Integer

    If tdb_Composicoes.EOF Then
        ExibeMensagem "É preciso selecionar alguma Composição de Receita da lista."
        Exit Sub
    End If
            
    For intFor = 0 To xadbComposicoes.UpperBound(1)
        
        If xadbComposicoes(intFor, 3) = tdb_Composicoes.Columns("intPosicao") Then
            
            xadbComposicoes.DeleteRows intFor
            
            Exit For
            
        End If
        
    Next
    
    Set tdb_Composicoes.Array = xadbComposicoes
    tdb_Composicoes.ReBind
    tdb_Composicoes.Refresh
    
End Sub
Private Sub IncluiComposicaoNoGrid()
Dim varAux            As Variant
Dim intPosicao        As Integer
    
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "É preciso selecionar alguma Composição de Receita."
        Exit Sub
    End If
            
    If Not dbc_intExercicio.MatchedWithList Then
        ExibeMensagem "É preciso selecionar algum Exercício."
        Exit Sub
    End If
    
    If blnVerificaComposicoes(True) = True Then
        ExibeMensagem "Esta Composição de Receita já se encontra selecionada."
        Exit Sub
    End If
    
    If xadbComposicoes.UpperBound(1) > -1 Then
        'caso ja exista uma linha em branco nao vamos criar outra
        If Len(Trim(xadbComposicoes(xadbComposicoes.UpperBound(1), 3))) = 0 Then
            intPosicao = 0
        Else
            intPosicao = Val(xadbComposicoes(xadbComposicoes.UpperBound(1), 3)) + 1
            xadbComposicoes.AppendRows 1
        End If
    Else
        intPosicao = 0
        xadbComposicoes.AppendRows 1
    End If
    
    varAux = Space$(0) & dbc_intComposicao.BoundText
    xadbComposicoes(xadbComposicoes.UpperBound(1), 0) = varAux                 'IntComposicao
    
    varAux = Space$(0) & dbc_intComposicao.Text
    xadbComposicoes(xadbComposicoes.UpperBound(1), 1) = varAux                 'strComposicao
    
    varAux = Space$(0) & dbc_intExercicio.Text
    xadbComposicoes(xadbComposicoes.UpperBound(1), 2) = varAux                 'intExercicio
    
    varAux = intPosicao
    xadbComposicoes(xadbComposicoes.UpperBound(1), 3) = varAux                  'Posição
    
    Set tdb_Composicoes.Array = xadbComposicoes
    tdb_Composicoes.ReBind
    tdb_Composicoes.Refresh
    
    dbc_intComposicao.BoundText = Space$(0)
    dbc_intComposicao.SetFocus
    
End Sub
Private Sub LimpaGrid()
    
    Set xadbComposicoes = New XArrayDB
    xadbComposicoes.Clear
    xadbComposicoes.ReDim -1, -1, 0, 3
            
    Set tdb_Composicoes.Array = xadbComposicoes
    tdb_Composicoes.ReBind
    tdb_Composicoes.Refresh

End Sub
Private Function blnVerificaComposicoes(blnDuplicada As Boolean) As Boolean
Dim intFor As Integer
    
    blnVerificaComposicoes = False
    
    If blnDuplicada Then 'Verifica se ja existe a mesma composicao no grid
        For intFor = 0 To xadbComposicoes.UpperBound(1)
            If Val(xadbComposicoes(intFor, 0)) = dbc_intComposicao.BoundText And Val(xadbComposicoes(intFor, 2)) = dbc_intExercicio.Text Then
                blnVerificaComposicoes = True
                Exit Function
            End If
        Next
    Else 'Verifica se existe composicao no grid
        For intFor = 0 To xadbComposicoes.UpperBound(1)
            If Val(xadbComposicoes(intFor, 0)) > 0 Then
                blnVerificaComposicoes = True
                Exit Function
            End If
        Next
    End If
    
End Function

Private Sub txtVencimento_GotFocus()
    MarcaCampo txtVencimento
End Sub

Private Sub txtVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtVencimento
End Sub

Private Sub txtVencimento_LostFocus()
    txtVencimento = gstrDataFormatada(txtVencimento)
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If txtVencimento = Space$(0) Then
       MsgBox "A Data de Vencimento deve ser preenchida.", vbOKOnly, "Mensagem ao Usuário"
       txtVencimento.SetFocus
       Exit Function
    End If
    
    If gblnDataValida(txtVencimento.Text) = False Then
       MsgBox "Data inválida.", vbOKOnly, "Mensagem ao Usuário"
       txtVencimento.SetFocus
       Exit Function
    End If
    
    If txt_intLote = Space$(0) Then
       MsgBox "O Lote deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
       txt_intLote.SetFocus
       Exit Function
    End If
    
    If Not dbc_intIndexador.MatchedWithList Then
        MsgBox "O Indexador deve ser selecionado.", vbOKOnly, "Mensagem ao Usuário"
        dbc_intIndexador.SetFocus
        Exit Function
    End If
    
    If txt_dblValorMoeda = Space$(0) Then
       MsgBox "O Valor deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
       txt_dblValorMoeda.SetFocus
       Exit Function
    End If
    
    If blnVerificaComposicoes(False) = False Then
        ExibeMensagem "Não há nenhuma composição para gerar o calculo de executivo fiscal."
        Exit Function
    End If
    
    blnDadosOk = True
    
End Function

Private Sub RetornaLote(blnReservarNoBanco As Boolean)
Dim adoResultado As New ADODB.Recordset
Dim strSql       As String

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT " & gstrISNULL("intLoteCobAmigavel", "0") & " intLote FROM " & gstrParametrosTributario, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_intLote = adoResultado("intLote").Value + 1
        Else
            ExibeMensagem "Não existem parâmetros de Tributário."
        End If
    Else
        ExibeMensagem "Não existem parâmetros de Tributário."
    End If
    
    'Vamos atualizar o valor das sequencias no banco
    If blnReservarNoBanco Then
        gobjBanco.Execute "UPDATE " & gstrParametrosTributario & " SET intLoteCobAmigavel = " & txt_intLote
    End If

End Sub

Private Sub RealizaCalculoCobrancaAmigavel()
Dim adoResultado      As New ADODB.Recordset
Dim adoParcelas       As New ADODB.Recordset
Dim adoDativa         As New ADODB.Recordset

Dim strInscricaoAtual As String
Dim lngInscricaoAtual As Long
Dim lngComposicaoAtual As Long
Dim intExercicioAtual As Long

Dim intNumSequencial  As Integer

Dim intFor            As Long
Dim strSql            As String
Dim strSqlSimulado    As String
Dim strExercicios     As String
Dim strComposicoes    As String

'Dim dblTotalOriginal  As Double
Dim dblTotalPrincipal As Double
Dim dblTotalMulta     As Double
Dim dblTotalJuros     As Double
Dim dblTotalCorrecao  As Double
Dim dblTotalGeral     As Double
Dim dblTotalGrupo     As Double
Dim strAlfaPorInscr   As String

Dim lngAlfaAmigavel   As Long
Dim intParcelas       As Integer

Dim blnFimDeArquivo   As Boolean

Dim xadbParcelas      As XArrayDB
Dim xadbParcelas2     As XArrayDB
Dim intPosition       As Integer
Dim intNumSequencial2 As Long

Dim blnUltimoAlfa     As Boolean

On Error GoTo Problema_Na_Rotina

    Set gobjBanco = New clsBanco
    Screen.MousePointer = vbHourglass
    
    prg_Status.Value = 0
    prg_Status.Visible = True
    lbl_Status.Visible = True

    'Vamos obter os sequencias mais atuais, e reserva-los no banco
    RetornaLote Not chk_Simulado.Value = vbChecked
    
    strSql = ""
    
    'Vamos fazer a busca com parametros informados
    For intFor = 0 To xadbComposicoes.UpperBound(1)
        'Caso exista mais de 1 composicao
        If Len(Trim(strSql)) > 0 Then strSql = strSql & " UNION "
        
        strSql = strSql & " SELECT LA.strInscricao, LA.pkid, LA.intComposicaoDaREceita, LA.strComposicaoDaReceita, LA.intUtilizacao, LA.intExercicio, LV.Intparcela, LV.dtmDtVencimento, LV.dblValor ValorOrig,  LV.intMoeda, "
        If bytDBType = EDatabases.SQLServer Then
            strSql = strSql & "CASE WHEN IM.strLogradouroC IS NULL THEN CO.strLogradouroC WHEN IM.strLogradouroC = '' THEN CO.strLogradouroC ELSE IM.strLogradouroC END strLogradouroC, " & _
                              "CASE WHEN IM.strLogradouroC IS NULL THEN CO.intNumeroC WHEN IM.strLogradouroC = '' THEN CO.intNumeroC ELSE IM.intNumeroC END intNumeroC, " & _
                              "CASE WHEN IM.strLogradouroC IS NULL THEN CO.strComplementoC WHEN IM.strLogradouroC = '' THEN CO.strComplementoC ELSE IM.strComplementoC END strComplementoC, " & _
                              "CASE WHEN IM.strLogradouroC IS NULL THEN CO.strBairroC WHEN IM.strLogradouroC = '' THEN CO.strBairroC ELSE IM.strBairroC END strBairroC, " & _
                              "CASE WHEN IM.strLogradouroC IS NULL THEN MU2.strDescricao WHEN IM.strLogradouroC = '' THEN MU2.strDescricao ELSE MU.strDescricao END strMunicipioC, " & _
                              "CASE WHEN IM.strLogradouroC IS NULL THEN UF2.strSigla WHEN IM.strLogradouroC = '' THEN UF2.strSigla ELSE UF.strSigla END strufc, " & _
                              "CASE WHEN IM.strLogradouroC IS NULL THEN CO.intcepc WHEN IM.strLogradouroC = '' THEN CO.intcepc ELSE IM.intcepc END intcepc, "
        Else
            strSql = strSql & gstrCASEWHEN("IM.strLogradouroC", "NULL, CO.strLogradouroC", "IM.strLogradouroC") & "strLogradouroC, " & _
                              gstrCASEWHEN("IM.strlogradouroc", "NULL, CO.intNumeroC", "IM.intNumeroC") & " intNumeroC, " & gstrCASEWHEN("IM.strlogradouroc", "NULL, CO.strComplementoC", "IM.strComplementoC") & " strComplementoC, " & gstrCASEWHEN("IM.strlogradouroc", "NULL, CO.strBairroC", "IM.strBairroC") & " strBairroC, " & gstrCASEWHEN("IM.strlogradouroc", "NULL, MU2.strDescricao", "MU.strDescricao") & " strMunicipioC, " & gstrCASEWHEN("IM.strlogradouroc", "NULL, UF2.strSigla", "UF.strSigla") & " strufc, " & gstrCASEWHEN("IM.strlogradouroc", "NULL, CO.intcepc", "IM.intcepc") & " intcepc, "
        End If
        strSql = strSql & " CO.strNome strNomeProprietario, CO.strIdentidade, CO.strCnpjCpf, LO.strDescricao strLogradouro, IM.intNumero, IM.strComplemento, BA.strDescricao strBairro, LA.strMunicipio, LA.strUF, IM.intCep, LV.Pkid intLancamentoValor "
        strSql = strSql & " FROM " & gstrLancamentoAlfa & " LA, " & _
                          gstrLancamentoValor & " LV, " & _
                          gstrImobiliario & " IM, " & _
                          gstrContribuinte & " CO, " & _
                          gstrLogradouro & " LO, " & _
                          gstrBairro & " BA, " & _
                          gstrCidade & " MU, " & _
                          gstrUF & " UF, " & _
                          gstrCidade & " MU2, " & _
                          gstrUF & " UF2 "
        strSql = strSql & " WHERE LA.intComposicaoDaReceita = " & xadbComposicoes(intFor, 0) & " AND " & _
                          " LA.intExercicio = " & xadbComposicoes(intFor, 2) & " AND " & _
                          " LV.Intlancamentoalfa = LA.Pkid AND " & _
                          " IM.strInscricao = LA.strInscricao AND " & _
                          " CO.Pkid = IM.intContribuinte AND " & _
                          " LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " IM.intLogradouro AND " & _
                          " BA.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " IM.intBairro AND " & _
                          " MU.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " IM.intMunicipioC AND " & _
                          " UF.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " IM.intUfC AND " & _
                          " MU2.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " CO.intMunicipioC AND " & _
                          " UF2.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " CO.intUfC AND " & _
                          " " & gstrISNULL("LV.dblValor", "0") & " <> 0 AND " & _
                          " LV.bitParcelaValida = 1 AND " & _
                          " LV.Pkid Not in(SELECT intLancamentoValor FROM " & gstrLancamentoPagamento & ") AND "
        'Nao vamos considerar as parcelas em acordo
        If Not chk_ParcelasAcordo.Value Then
            strSql = strSql & " LV.intLancamentoAlfaAcordo is Null "
        End If
        
        'strSql = strSql & " AND LA.strinscricao between '00000000000001001001' and '00000000000003999999' "
    
    Next
    
    strSql = strSql & " Group By LA.strInscricao, LA.pkid, LA.intComposicaoDaREceita, LA.strComposicaoDaReceita, LA.intUtilizacao, LA.intExercicio, LV.Intparcela, LV.dtmDtVencimento, LV.dblValor,  LV.intMoeda, " & _
                               " IM.strlogradouroc, IM.intNumeroC, IM.strComplementoC, IM.strBairroC, MU.strDescricao, UF.strSigla, IM.intcepc, CO.strNome, CO.strIdentidade, CO.strCnpjCpf, LO.strDescricao, IM.intNumero, IM.strComplemento, BA.strDescricao, LA.strMunicipio, LA.strUF, IM.intCep, LV.Pkid, " & _
                               " CO.Strlogradouroc, CO.intNumeroC, CO.strComplementoC, CO.strBairroC, CO.intcepc, MU2.strDescricao,UF2.strSigla "
    If bytDBType = Oracle Then
        strSql = strSql & " Order By strInscricao, strComposicaoDaReceita , intExercicio, intParcela "
    Else
        strSql = strSql & " Order By LA.strInscricao, LA.strComposicaoDaReceita, LA.intExercicio , LV.intParcela "
    End If
    
    lbl_Status.Caption = "Consultando registros..."
    Me.Refresh
    
    If gobjBanco.CriaADO(strSql, 300, adoResultado) Then
        If Not adoResultado.EOF Then
            
            lbl_Status.Caption = "Gerando Cobrança Amigável..."
            prg_Status.Max = adoResultado.RecordCount
            Me.Refresh
            
            With adoResultado
            
                'Vamos calcular os valores de cada parcela
                For intFor = 0 To adoResultado.RecordCount - 1
                    
                    lngComposicaoAtual = adoResultado("intComposicaoDaReceita").Value
                    intExercicioAtual = adoResultado("intExercicio").Value
                    
                    gobjBanco.ExecutaBeginTrans
                    
RealizaGravacao:

                    strSql = gstrStoredProcedure("sp_AtualizaParcela", !intComposicaoDaReceita & ", " & !intExercicio & ", " & !intParcela & ", " & gstrConvDtParaSql(!Dtmdtvencimento) & ", " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & gstrConvVrParaSql(!ValorOrig) & ", " & !intMoeda, True)

                    Set gobjBanco = New clsBanco
                    
                    If gobjBanco.CriaADO(strSql, 80, adoParcelas) Then
                        
                        If (adoResultado("strInscricao").Value <> strInscricaoAtual) Then
                            
                            intNumSequencial = intNumSequencial + 1
                            
                            'Vamos preencher a tblLancamentoCobAmigavelAlfa
                            strSql = "INSERT INTO tblLancamentoCobAmigavelAlfa " & _
                                     "(intLote, intNumeroCA, strNome, " & _
                                     " strLogradouro, strNumero, strComplemento, " & _
                                     " strBairro, strMunicipio, strUF, " & _
                                     " intCEP, strLogradouroC, strNumeroC, " & _
                                     " strComplementoC, strBairroC, strMunicipioC, " & _
                                     " strUFC, intCepC, dtmDtVencimentoCoAAmigavel, strIndexador, " & _
                                     " dblVlIndexador, " & _
                                     " intGuias, dtmDtAtualizacao, lngCodUsr)"
                            strSql = strSql & " VALUES " & _
                                     "(" & txt_intLote & ", " & intNumSequencial & ", '" & Trim(Replace(adoResultado("strNomeProprietario") & Space$(0), "'", Chr(207))) & _
                                     "', '" & Trim(Replace(adoResultado("strLogradouro") & Space$(0), "'", Chr(207))) & "', '" & IIf(Trim(adoResultado("intNumero")) = "", 0, Trim(adoResultado("intNumero"))) & "', '" & Trim(Replace(adoResultado("strComplemento") & Space$(0), "'", Chr(207))) & _
                                     "', '" & Trim(Replace(adoResultado("strBairro") & Space$(0), "'", Chr(207))) & "', '" & Trim(adoResultado("strMunicipio")) & "', '" & Trim(adoResultado("strUF")) & _
                                     "', " & IIf(IsNull(adoResultado("intCep")), "NULL", adoResultado("intCep")) & ", '" & Trim(Replace(adoResultado("strLogradouroC") & Space$(0), "'", Chr(207))) & "', '" & IIf(Trim(adoResultado("intNumeroC")) = "", 0, Trim(adoResultado("intNumeroC"))) & _
                                     "', '" & Trim(Replace(adoResultado("strComplementoC") & Space$(0), "'", Chr(207))) & "', '" & Trim(Replace(adoResultado("strBairroC") & Space$(0), "'", Chr(207))) & "', '" & Trim(adoResultado("strMunicipioC")) & _
                                     "', '" & Trim(adoResultado("strUFC")) & "', " & gstrENulo(adoResultado("intCepC"), , True) & ", " & gstrConvDtParaSql(txtVencimento.Text) & ", '" & dbc_intIndexador.Text & _
                                     "', " & gstrConvVrParaSql(txt_dblValorMoeda) & _
                                     ",NULL , " & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                                     
                            If Not gobjBanco.Execute(strSql, True) Then
                                gobjBanco.ExecutaRollbackTrans
                                Screen.MousePointer = vbDefault
                                ExibeMensagem "Não foi possível Lançamento Alfa de Cobrança Amigável para a Inscrição " & !strInscricao & " (" & !strComposicaoDaReceita & " " & !intExercicio & "), Parcela " & !intParcela & ". A operação foi cancelada." & Chr(13) & Err.Description
                                GoTo ProximoRegistro
                                'Exit Sub
                            End If
                                     
                            lngAlfaAmigavel = glngRetornaPkidTabelaPai("SEQtblLancCobAmigavelAlfa", "tblLancamentoCobAmigavelAlfa")
                            
                            'Vamos zerar os valores
                            'dblTotalOriginal = 0
                            dblTotalPrincipal = 0
                            dblTotalMulta = 0
                            dblTotalJuros = 0
                            dblTotalCorrecao = 0
                            dblTotalGrupo = 0
                            
                        End If
                        
                        'Vamos totalizar os valores por inscricao
                        'dblTotalOriginal = dblTotalOriginal + CCur(gstrConvVrDoSql(adoResultado("ValorOrig").Value))
                        dblTotalPrincipal = dblTotalPrincipal + CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                        dblTotalMulta = dblTotalMulta + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value))
                        dblTotalJuros = dblTotalJuros + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value))
                        dblTotalCorrecao = dblTotalCorrecao + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                        dblTotalGeral = CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                        dblTotalGrupo = dblTotalGrupo + CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))

                        strInscricaoAtual = adoResultado("strInscricao").Value
                        lngInscricaoAtual = adoResultado("Pkid").Value
 
                        'Vamos preencher a tblLancamentoCobAmigavelValor
                        strSql = "INSERT INTO tblLancamentoCobAmigavelValor " & _
                                 "(dblVlPrincipal, dblVlMulta, dblVlJuros, " & _
                                 " dblVlCorrecao, dblVlTotal, intCobAmigavel, " & _
                                 " intLancamentoValor, dtmDtAtualizacao, lngCodUsr)"
                        strSql = strSql & " VALUES " & _
                                 "(" & gstrConvVrParaSql(adoParcelas("dblValorPrincipal").Value) & ", " & gstrConvVrParaSql(adoParcelas("dblValorMulta").Value) & ", " & gstrConvVrParaSql(adoParcelas("dblValorJuros").Value) & _
                                 ", " & gstrConvVrParaSql(adoParcelas("dblValorCorrecao").Value) & ", " & gstrConvVrParaSql(dblTotalGeral) & ", " & lngAlfaAmigavel & _
                                 ", " & gstrConvVrParaSql(adoResultado("intLancamentoValor").Value) & ", " & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                                 
                        If Not gobjBanco.Execute(strSql, True) Then
                            gobjBanco.ExecutaRollbackTrans
                            Screen.MousePointer = vbDefault
                            ExibeMensagem "Não foi possível gravar Lançamento Valor de Cobrança Amigável para a Inscrição " & !strInscricao & " (" & !strComposicaoDaReceita & " " & !intExercicio & "), Parcela " & !intParcela & ". A operação foi cancelada." & Chr(13) & Err.Description
                            GoTo ProximoRegistro
                            'Exit Sub
                        End If
                        
                        adoResultado.MoveNext
                        
                        If Not adoResultado.EOF Then
                            'Vamos verificar se sera agrupado o array
                            blnUltimoAlfa = adoResultado("Pkid").Value <> lngInscricaoAtual
                        End If
                        
                        'Vamos atualizar o Alfa com os valores somados
                        If blnUltimoAlfa Or adoResultado.EOF Then
                            gobjBanco.Execute " UPDATE tblLancamentoCobAmigavelAlfa SET " & _
                                              " dblVlTotalPrincipal = " & gstrConvVrParaSql(dblTotalPrincipal) & _
                                              ", dblVlTotalMulta = " & gstrConvVrParaSql(dblTotalMulta) & _
                                              ", dblVlTotalJuros = " & gstrConvVrParaSql(dblTotalJuros) & _
                                              ", dblVlTotalCorrecao = " & gstrConvVrParaSql(dblTotalCorrecao) & _
                                              ", dblVlTotal = " & gstrConvVrParaSql(dblTotalGrupo) & _
                                              " WHERE intLote = " & txt_intLote & " And intNumeroCA = " & intNumSequencial
                            
                        End If

                        adoResultado.MovePrevious
                        
                    Else
                        'Vamos para a proxima composicao e exercicio
                        Do While lngComposicaoAtual = adoResultado("intComposicaoDaReceita").Value And intExercicioAtual = adoResultado("intExercicio").Value
                            adoResultado.MoveNext
                            intFor = intFor + 1
                            'Caso chegue ao final do arquivo
                            If adoResultado.EOF Then
                               'Se ja existir registro da inscricao a ser salvo
                               If dblTotalGeral > 0 Then
                                    adoResultado.MovePrevious
                                   GoTo RealizaGravacao
                               Else
                                   GoTo FinalizaOperacao
                               End If
                            End If
                        Loop
                    End If
                    
                    If Not chk_Simulado.Value = vbChecked Then
                        gobjBanco.ExecutaCommitTrans
                    End If
                    
ProximoRegistro:
                    
                    adoResultado.MoveNext
                    
                    prg_Status.Value = prg_Status.Value + 1
                    
                    Me.Refresh
                    
                 Next
                    
            End With
            
        Else
            gobjBanco.ExecutaRollbackTrans
            ExibeMensagem "Não foi(ram) encontrado(s) lançamento(s)."
            Screen.MousePointer = vbDefault
            prg_Status.Visible = False
            lbl_Status.Visible = False
            Exit Sub
        End If
    End If
    
FinalizaOperacao:
    
    If chk_Simulado.Value Then
        gobjBanco.ExecutaRollbackTrans
    Else
        gobjBanco.ExecutaCommitTrans
    End If
    
    Screen.MousePointer = vbDefault
    
    prg_Status.Visible = False
    lbl_Status.Visible = False
    
    Exit Sub
    
Problema_Na_Rotina:
    ExibeDetalheErro Err.Description
    Resume
    Exit Sub
    
End Sub

