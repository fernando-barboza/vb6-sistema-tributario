VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "Todg7.ocx"
Begin VB.Form frmLancamentoExecutivosFiscais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Executivos Fiscais"
   ClientHeight    =   4875
   ClientLeft      =   2550
   ClientTop       =   3690
   ClientWidth     =   9690
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   9690
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4770
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   8414
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Executivos Fiscais - Calculo"
      TabPicture(0)   =   "frmLancamentoExecutivosFiscais.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_intSeqInicial"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_intLote"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_intIndexador"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_dblValorMoeda"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_Status"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_ContInicial"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_ContFinal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbc_intIndexador"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_intSeqInicial"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_intLote"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_dblValorMoeda"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmd_intIndexador"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fra_ComposicaoDaReceita"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chk_ExecutarDebitos"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chk_DistribuicaoEletronica"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "prg_Status"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chk_Simulado"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.CheckBox chk_Simulado 
         Caption         =   "Simulado"
         Height          =   375
         Left            =   4800
         TabIndex        =   21
         Top             =   1140
         Width           =   2475
      End
      Begin MSComctlLib.ProgressBar prg_Status 
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   4470
         Visible         =   0   'False
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CheckBox chk_DistribuicaoEletronica 
         Caption         =   "Para Distribuição Eletrônica"
         Height          =   375
         Left            =   1470
         TabIndex        =   11
         Top             =   1140
         Width           =   2475
      End
      Begin VB.CheckBox chk_ExecutarDebitos 
         Caption         =   "Executar Débitos Parcelados"
         Height          =   375
         Left            =   1470
         TabIndex        =   10
         Top             =   810
         Width           =   2475
      End
      Begin VB.Frame fra_ComposicaoDaReceita 
         Caption         =   "Composição da Receita"
         Height          =   2700
         Left            =   180
         TabIndex        =   12
         Top             =   1560
         Width           =   9195
         Begin VB.CommandButton cmd_Composicao 
            Height          =   300
            Left            =   6630
            Picture         =   "frmLancamentoExecutivosFiscais.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro de Composição da Receita"
            Top             =   315
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbc_intComposicao 
            Height          =   315
            Left            =   1230
            TabIndex        =   14
            Top             =   315
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intExercicio 
            Height          =   315
            Left            =   7980
            TabIndex        =   17
            Top             =   330
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Composicoes 
            Height          =   1725
            Left            =   210
            TabIndex        =   18
            Top             =   810
            Width           =   8790
            _ExtentX        =   15505
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
         Begin VB.Label lbl_Composicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   270
            TabIndex        =   13
            Top             =   390
            Width           =   870
         End
         Begin VB.Label lbl_Exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   7230
            TabIndex        =   16
            Top             =   390
            Width           =   765
         End
      End
      Begin VB.CommandButton cmd_intIndexador 
         Height          =   300
         Left            =   6150
         Picture         =   "frmLancamentoExecutivosFiscais.frx":013A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Indexador Econônico."
         Top             =   450
         Width           =   360
      End
      Begin VB.TextBox txt_dblValorMoeda 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8070
         TabIndex        =   9
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox txt_intLote 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3150
         MaxLength       =   5
         TabIndex        =   4
         Top             =   450
         Width           =   705
      End
      Begin VB.TextBox txt_intSeqInicial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   2
         Top             =   450
         Width           =   1110
      End
      Begin MSDataListLib.DataCombo dbc_intIndexador 
         Height          =   315
         Left            =   4770
         TabIndex        =   6
         Top             =   450
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Label lbl_ContFinal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9285
         TabIndex        =   23
         Top             =   4260
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lbl_ContInicial 
         AutoSize        =   -1  'True
         Caption         =   "1"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   4260
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lbl_Status 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   4260
         Visible         =   0   'False
         Width           =   9165
      End
      Begin VB.Label lbl_dblValorMoeda 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Indexador"
         Height          =   195
         Left            =   6660
         TabIndex        =   8
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label lbl_intIndexador 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   195
         Left            =   3990
         TabIndex        =   5
         Top             =   525
         Width           =   1035
      End
      Begin VB.Label lbl_intLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   525
         Width           =   315
      End
      Begin VB.Label lbl_intSeqInicial 
         AutoSize        =   -1  'True
         Caption         =   "Executivo Inicial"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   525
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmLancamentoExecutivosFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xadbComposicoes      As XArrayDB

Private Sub cmd_intIndexador_Click()
    ChamaFormCadastro frmIndexadorEconomico, dbc_intIndexador
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
                    
            txt_intSeqInicial.SetFocus
        
        Case Is = UCase(gstrCalcularReajuste)
            If blnDadosOK Then
                RealizaCalculoExecutivo
            End If
            
    End Select

End Sub

Private Sub Form_Load()
    dbc_intIndexador.Tag = strQueryIndexador & ";strAbreviatura"
    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"
    
    Set xadbComposicoes = New XArrayDB
    xadbComposicoes.Clear
    xadbComposicoes.ReDim 0, 0, 0, 3
    
    TrocaCorObjeto txt_intSeqInicial, True
    TrocaCorObjeto txt_intLote, True
    
    'Vamos auto numerar
    If Len(Trim(txt_intSeqInicial)) = 0 Then
        ProximaSequenciaLote False
    End If

End Sub

Private Sub txt_dblValorMoeda_GotFocus()
    MarcaCampo txt_dblValorMoeda
End Sub

Private Sub txt_dblValorMoeda_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorMoeda
End Sub

Private Sub txt_dblValorMoeda_LostFocus()
    txt_dblValorMoeda.Text = gstrConvVrDoSql(txt_dblValorMoeda.Text, 5)
End Sub

Private Sub txt_intLote_GotFocus()
    MarcaCampo txt_intLote
End Sub

Private Sub txt_intLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intLote
End Sub

Private Sub txt_intSeqInicial_GotFocus()
    MarcaCampo txt_intSeqInicial
End Sub

Private Sub txt_intSeqInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intSeqInicial
End Sub

Private Sub LimpaDataCombo(dbcAux As DataCombo)

    dbcAux.Tag = ""
    dbcAux.Text = ""
    dbcAux.ListField = ""
    
End Sub

Private Sub LimpaGrid()
    
    Set xadbComposicoes = New XArrayDB
    xadbComposicoes.Clear
    xadbComposicoes.ReDim -1, -1, 0, 3
            
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

Private Sub RealizaCalculoExecutivo()
Dim adoResultado      As New ADODB.Recordset
Dim adoParcelas       As New ADODB.Recordset
Dim adoDativa         As New ADODB.Recordset

Dim strInscricaoAtual As String
Dim lngInscricaoAtual As Long
Dim lngComposicaoAtual As Long
Dim intExercicioAtual As Long

Dim intNumSequencial  As Long

Dim intFor            As Long
Dim strSQL            As String
Dim strSqlSimulado    As String
Dim strExercicios     As String
Dim strComposicoes    As String

Dim dblTotalOriginal  As Double
Dim dblTotalPrincipal As Double
Dim dblTotalMulta     As Double
Dim dblTotalJuros     As Double
Dim dblTotalCorrecao  As Double
Dim dblTotalGeral     As Double
Dim dblTotalGrupo     As Double
Dim strAlfaPorInscr   As String

Dim lngExecutivo      As Long
Dim intParcelas       As Integer

Dim blnFimDeArquivo   As Boolean

Dim xadbParcelas      As XArrayDB
Dim xadbParcelas2     As XArrayDB
Dim intPosition       As Long
Dim intNumSequencial2 As Long

Dim blnUltimoAlfa     As Boolean

Dim aCritica          As XArrayDB
Dim varCritica        As Variant

On Error GoTo Problema_Na_Rotina

    Set gobjBanco = New clsBanco
    Screen.MousePointer = vbHourglass
    
    prg_Status.Value = 0
    prg_Status.Visible = True
    lbl_Status.Visible = True
    
    Set xadbParcelas = New XArrayDB
    xadbParcelas.ReDim 0, 0, 0, 3
    xadbParcelas.Clear

    Set aCritica = New XArrayDB
    aCritica.ReDim 0, 0, 0, 3
    aCritica.Clear
    
    'Vamos obter os sequencias mais atuais, e reserva-los no banco
    ProximaSequenciaLote True
    
    intNumSequencial = txt_intSeqInicial
    strAlfaPorInscr = ""
    
    
    strSQL = ""
    
    'Vamos fazer a busca com parametros informados
    For intFor = 0 To xadbComposicoes.UpperBound(1)
        'Caso exista mais de 1 composicao
        If Len(Trim(strSQL)) > 0 Then strSQL = strSQL & " UNION "
        
'        strSQL = strSQL & " SELECT LA.strInscricao, LA.pkid, Da.PkId Dativa, LA.intComposicaoDaREceita, LA.strComposicaoDaReceita, LA.intUtilizacao, LA.intExercicio, LV.Intparcela, LV.dtmDtVencimento, LV.dblValor ValorOrig,  LV.intMoeda, " & _
                          " LA.strlogradouroc, LA.strNumeroC, LA.strComplementoC, LA.strBairroC, LA.strMunicipioC, LA.strufc, LA.intcepc, LA.strNomeProprietario, LA.strIdentidade, LA.strCnpjCpf, LA.strNomeProprietario " & _
                          " FROM " & gstrLancamentoAlfa & " LA, " & _
                          gstrDativa & " DA, " & _
                          gstrLancamentoValor & " LV " & _
                          " WHERE intComposicaoDaReceita = " & xadbComposicoes(intFor, 0) & " AND " & _
                          " intExercicio = " & xadbComposicoes(intFor, 2) & " AND " & _
                          " LV.Intlancamentoalfa = LA.Pkid AND " & _
                          " DA.intLancamentoAlfa = LA.Pkid AND " & _
                          " DA.intExecutivo IS /*NOT*/ NULL  AND " & _
                          " " & gstrISNULL("LV.dblValor", "0") & " <> 0 AND " & _
                          " LV.Pkid not in(Select Intlancamentovalor From tblLancamentoPagamento) AND "
        
        strSQL = strSQL & " SELECT LA.strInscricao, LA.pkid, Da.PkId Dativa, LA.intComposicaoDaREceita, LA.strComposicaoDaReceita, LA.intUtilizacao, LA.intExercicio, LV.Intparcela, LV.dtmDtVencimento, LV.dblValor ValorOrig,  LV.intMoeda, " & _
                          " LA.strlogradouroc, LA.strNumeroC, LA.strComplementoC, LA.strBairroC, LA.strMunicipioC, LA.strufc, LA.intcepc, LA.strNomeProprietario, LA.strIdentidade, LA.strCnpjCpf, LA.strNomeProprietario, LA.strNumeroAviso "
                          
        If bytDBType = SQLServer Then
             strSQL = strSQL & "FROM " & gstrLancamentoAlfa & " LA "
             strSQL = strSQL & "INNER JOIN " & gstrLancamentoValor & " LV ON LA.PKId = LV.intLancamentoAlfa "
             strSQL = strSQL & "INNER JOIN " & gstrDativa & " DA ON LA.PKId = DA.INTLANCAMENTOALFA "
             strSQL = strSQL & "LEFT OUTER JOIN " & gstrLancamentoPagamento & " LP ON LV.PKId = LP.INTLANCAMENTOVALOR "
             strSQL = strSQL & "WHERE "
        Else
             strSQL = strSQL & " FROM " & gstrLancamentoAlfa & " LA, " & _
             gstrDativa & " DA, " & _
             gstrLancamentoValor & " LV, " & _
             gstrLancamentoPagamento & " LP " & _
             " WHERE LV.Intlancamentoalfa = LA.Pkid AND " & _
             " DA.intLancamentoAlfa = LA.Pkid AND " & _
             " LV.Pkid = LP.Intlancamentovalor " & strOUTJOracle & " AND "
        End If
             
        strSQL = strSQL & " DA.intExecutivo IS /*NOT*/ NULL  AND " & _
             " " & gstrISNULL("LV.dblValor", "0") & " <> 0 AND " & _
             " intComposicaoDaReceita = " & xadbComposicoes(intFor, 0) & " AND " & _
             " intExercicio = " & xadbComposicoes(intFor, 2) & " AND " & _
             " LP.Intlancamentovalor IS Null AND "
        strSQL = strSQL & " LV.bitParcelaValida = 1 AND "
            
        'Nao vamos considerar as parcelas em acordo
        If Not chk_ExecutarDebitos.Value = vbChecked Then
            strSQL = strSQL & " LV.Intlancamentoalfaacordo is null "
        End If
        
        'strSql = strSql & " AND LA.strinscricao between '00000000000002034017' and '00000000000002054028' "
    
    Next
    
    strSQL = strSQL & " Group By LA.strInscricao, LA.pkid, DA.PkId, LA.intComposicaoDaREceita, LA.strComposicaoDaReceita, LA.intUtilizacao, LA.intExercicio, LV.Intparcela, LV.dtmDtVencimento, LV.dblValor,  LV.intMoeda, " & _
                               " LA.strlogradouroc, LA.strNumeroC, LA.strComplementoC, LA.strBairroC, LA.strMunicipioC, LA.strufc, LA.intcepc, LA.strNomeProprietario, LA.strIdentidade, LA.strCnpjCpf, LA.strNomeProprietario, LA.strNumeroAviso "
    If bytDBType = Oracle Then
        strSQL = strSQL & " Order By strInscricao, Pkid , intParcela "
    Else
        strSQL = strSQL & " Order By LA.strInscricao, LA.Pkid , LV.intParcela "
    End If
    
    strSqlSimulado = Replace(strSQL, "/*NOT*/", "NOT")
    
    lbl_Status.Caption = "Consultando registros..."
    Me.Refresh
    
    If gobjBanco.CriaADO(strSQL, 500, adoResultado) Then
        If Not adoResultado.EOF Then
            
            lbl_Status.Caption = "Gerando Executivo Fiscal..."
            prg_Status.Max = adoResultado.RecordCount
            
            lbl_ContFinal.Caption = adoResultado.RecordCount
            lbl_ContInicial.Visible = True
            lbl_ContFinal.Visible = True
            
            Me.Refresh
            
            With adoResultado
            
                intParcelas = 0
                
                'Vamos calcular os valores de cada parcela
                For intFor = 0 To adoResultado.RecordCount - 1
                    gobjBanco.ExecutaBeginTrans
                    
                    strSQL = gstrStoredProcedure("sp_AtualizaParcela", !intComposicaoDaReceita & ", " & !intExercicio & ", " & !intParcela & ", " & gstrConvDtParaSql(!Dtmdtvencimento) & ", " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & gstrConvVrParaSql(!ValorOrig) & ", " & !intMoeda, True)

                    Set gobjBanco = New clsBanco
                    
                    If gobjBanco.CriaADO(strSQL, 80, adoParcelas) Then
                    
                        intParcelas = intParcelas + 1
                        
                        strInscricaoAtual = adoResultado("strInscricao").Value
                        lngInscricaoAtual = adoResultado("Pkid").Value
                        lngComposicaoAtual = adoResultado("intComposicaoDaReceita").Value
                        intExercicioAtual = adoResultado("intExercicio").Value

                        'Vamos totalizar os valores por inscricao
                        dblTotalOriginal = dblTotalOriginal + CCur(gstrConvVrDoSql(adoResultado("ValorOrig").Value))
                        dblTotalPrincipal = dblTotalPrincipal + CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                        dblTotalMulta = dblTotalMulta + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value))
                        dblTotalJuros = dblTotalJuros + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value))
                        dblTotalCorrecao = dblTotalCorrecao + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                        dblTotalGeral = dblTotalPrincipal + dblTotalMulta + dblTotalJuros + dblTotalCorrecao
                        dblTotalGrupo = dblTotalGrupo + CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                        
                        'Vamos obter o Pkid da tblDativa referente à parcela
                        strSQL = " SELECT DA.pkid " & _
                                " FROM " & gstrDativa & " DA, " & _
                                gstrDaParcel & " DAP " & _
                                " WHERE intLancamentoAlfa = " & adoResultado("Pkid").Value & " AND " & _
                                " DAP.IntDativa = DA.Pkid AND " & _
                                " DA.PkId = " & adoResultado("Dativa").Value & " AND " & _
                                " DAP.intParcela = " & adoResultado("intParcela").Value & _
                                " ORDER BY DA.Pkid "
                                
                          
                          
                        If gobjBanco.CriaADO(strSQL, 10, adoDativa) Then
                            If Not adoDativa.EOF Then
                                'Vamos preencher a tblExecutivoParcela
                                strSQL = "INSERT INTO tblExecutivoParcela " & _
                                         "(intDativa, dtmDtVencimento, intMoeda, " & _
                                         " dblVlOriginal, dblVlPrincipal, dblVlCorrecao, " & _
                                         " dblVlMulta, dblVlJuros, dblVlTotal, " & _
                                         " intParcela, dtmDtAtualizacao, lngCodUsr)"
                                strSQL = strSQL & " VALUES " & _
                                         "(" & adoDativa("Pkid").Value & ", " & gstrConvDtParaSql(adoResultado("Dtmdtvencimento")) & ", " & adoResultado("intMoeda") & _
                                         ", " & gstrConvVrParaSql(adoResultado("ValorOrig").Value) & ", " & gstrConvVrParaSql(adoParcelas("dblValorPrincipal").Value) & ", " & gstrConvVrParaSql(adoParcelas("dblValorCorrecao").Value) & _
                                         ", " & gstrConvVrParaSql(adoParcelas("dblValorMulta").Value) & ", " & gstrConvVrParaSql(adoParcelas("dblValorJuros").Value) & ", " & gstrConvVrParaSql(CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))) & _
                                         ", " & adoResultado("intParcela").Value & ", " & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                                         
                                If Not gobjBanco.Execute(strSQL, False) Then
                                    gobjBanco.ExecutaRollbackTrans
                                    Screen.MousePointer = vbDefault
                                   'ExibeMensagem "Não foi possível criar o Parcelas do Executivo. A operação foi cancelada."
                                   
                                    aCritica.AppendRows
                                    
                                    varCritica = gstrFormataInscricao(Right$(adoResultado("strInscricao"), gintRetornaTamanhoMascara(TYP_IMOBILIARIA))): aCritica(aCritica.UpperBound(1), 0) = varCritica
                                    varCritica = adoResultado("strComposicaoDaReceita"): aCritica(aCritica.UpperBound(1), 1) = varCritica
                                    varCritica = adoResultado("intExercicio"): aCritica(aCritica.UpperBound(1), 2) = varCritica
                                    varCritica = adoResultado("strNumeroAviso"): aCritica(aCritica.UpperBound(1), 3) = varCritica
                                   
                                    'prg_Status.Visible = False
                                    'lbl_Status.Visible = False
                                    GoTo ProximoRegistro
                                    'Exit Sub
                                End If
                            Else
                                gobjBanco.ExecutaRollbackTrans
                                Screen.MousePointer = vbDefault
                               'ExibeMensagem "Não foi possível localizar referência de Divida Ativa da parcela. A operação foi cancelada"
                               
                                aCritica.AppendRows
                                
                                varCritica = gstrFormataInscricao(Right$(adoResultado("strInscricao"), gintRetornaTamanhoMascara(TYP_IMOBILIARIA))): aCritica(aCritica.UpperBound(1), 0) = varCritica
                                varCritica = adoResultado("strComposicaoDaReceita"): aCritica(aCritica.UpperBound(1), 1) = varCritica
                                varCritica = adoResultado("intExercicio"): aCritica(aCritica.UpperBound(1), 2) = varCritica
                                varCritica = adoResultado("strNumeroAviso"): aCritica(aCritica.UpperBound(1), 3) = varCritica
                               
                                'prg_Status.Visible = False
                                'lbl_Status.Visible = False
                                GoTo ProximoRegistro
                                'Exit Sub
                            End If
                            
                            adoDativa.Close: Set adoDativa = Nothing
                            
                        Else
                            gobjBanco.ExecutaRollbackTrans
                            Screen.MousePointer = vbDefault
                           'ExibeMensagem "Não foi possível localizar referência de Divida Ativa da parcela. A operação foi cancelada"
                           
                            aCritica.AppendRows
                            
                            varCritica = gstrFormataInscricao(Right$(adoResultado("strInscricao"), gintRetornaTamanhoMascara(TYP_IMOBILIARIA))): aCritica(aCritica.UpperBound(1), 0) = varCritica
                            varCritica = adoResultado("strComposicaoDaReceita"): aCritica(aCritica.UpperBound(1), 1) = varCritica
                            varCritica = adoResultado("intExercicio"): aCritica(aCritica.UpperBound(1), 2) = varCritica
                            varCritica = adoResultado("strNumeroAviso"): aCritica(aCritica.UpperBound(1), 3) = varCritica
                           
                            'prg_Status.Visible = False
                            'lbl_Status.Visible = False
                            GoTo ProximoRegistro
                            'Exit Sub
                        End If
                        
                        'Caso seja o ultimo registro da inscricao
                        adoResultado.MoveNext
                        
                        If adoResultado.EOF Then GoTo RealizaGravacao
                        
                        If (adoResultado("strInscricao").Value <> strInscricaoAtual) Then
RealizaGravacao:
                            'Vamos verificar se sera agrupado o array
                            blnUltimoAlfa = True
                            
                            adoResultado.MovePrevious
                            
                            'Vamos armazenar os pkids dos LancamentosAlfa a serem atualizados na tblDativa
                            If blnUltimoAlfa Then
                                strAlfaPorInscr = strAlfaPorInscr & adoResultado("Pkid").Value & ","
                            End If
                            
                            'Vamos preencher a tblExecutivo
                            strSQL = "INSERT INTO " & gstrExecutivo & " " & _
                                     "(intNumeroProtocolo, intSerieProtocolo, bitDistribuicaoEletronica, " & _
                                     " bitDistribuido, intNumSeq, intExecutadoCepNotif, " & _
                                     " strExecutadoCidNotif, strExecutadoComplNotif, strExecutadoNumLogNotif, " & _
                                     " strExecutadoNomeLogNotif, strExecutadoTitLogNotif, strExecutadoTpLogNotif, " & _
                                     " strExecutadoIdentidade, strExecutadoCnpjCpf, strExecutadoNome, " & _
                                     " dblQuantIndexador, dblVlIndexador, strIndexadorDescr, " & _
                                     " DBLVLTOTTOTAL, DBLVLTOTJUROS, DBLVLTOTMULTA, " & _
                                     " DBLVLTOTCORRECAO, DBLVLTOTPRINCIPAL, DBLVLTOTORIGINAL, " & _
                                     " DBLVLTOTTAXAS, DBLVLTOTIMPOSTOS, DTMDTCALCULOPETICAO, " & _
                                     " INTVARA, INTFOLHASOFICIO, INTLIVROOFICIO, " & _
                                     " INTNUMOFICIO, INTFOLHASDISTRIBUIDOR, INTLIVRODISTRIBUIDOR, " & _
                                     " DTMDTDISTRIBUIDOR, STRSERIEDISTRIBUIDOR, STRNUMDISTRIBUIDOR, " & _
                                     " STRPROTOCOLO, INTLOTEEXECUTIVO, STREXECUTADOBAIRRONOTIF, " & _
                                     " STREXECUTADOUFNOTIF, dtmDtAtualizacao, lngCodUsr)"
                            strSQL = strSQL & " VALUES " & _
                                     "(NULL, NULL, " & chk_DistribuicaoEletronica.Value & _
                                     ", NULL, " & intNumSequencial & ", " & gstrENulo(adoResultado("intcepc").Value, , True) & _
                                     ", '" & Replace(gstrENulo(adoResultado("strMunicipioC").Value), "'", "") & "', '" & gstrTrataApostrofe(gstrENulo(adoResultado("strComplementoC").Value)) & "', '" & adoResultado("strNumeroC").Value & "'" & _
                                     ", '" & Replace(gstrENulo(adoResultado("strlogradouroc").Value), "'", "") & "', '', ''" & _
                                     ", '" & Replace(gstrENulo(adoResultado("strIdentidade").Value), "'", "") & "', '" & gstrENulo(adoResultado("strCnpjCpf").Value) & "', '" & Replace(gstrENulo(adoResultado("strNomeProprietario").Value), "'", "") & "'" & _
                                     ", " & gstrConvVrParaSql(gstrConvVrDoSql(dblTotalGeral / CDbl(txt_dblValorMoeda), 2)) & ", " & gstrConvVrParaSql(txt_dblValorMoeda) & ", '" & dbc_intIndexador.Text & "'" & _
                                     ", " & gstrConvVrParaSql(dblTotalGeral) & ", " & gstrConvVrParaSql(dblTotalJuros) & ", " & gstrConvVrParaSql(dblTotalMulta) & _
                                     ", " & gstrConvVrParaSql(dblTotalCorrecao) & ", " & gstrConvVrParaSql(dblTotalPrincipal) & ", " & gstrConvVrParaSql(dblTotalOriginal) & _
                                     ", NULL, NULL, " & gstrConvDtParaSql(gstrDataDoSistema) & _
                                     ", NULL, NULL, NULL" & _
                                     ", NULL, NULL, NULL" & _
                                     ", NULL, '', ''" & _
                                     ", '', " & txt_intLote & ", '" & gstrTrataApostrofe(gstrENulo(adoResultado("strBairroC").Value)) & "'" & _
                                     ", '" & gstrENulo(adoResultado("strufc").Value) & "', " & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                                     
                            intNumSequencial2 = intNumSequencial
                                     
                            If Not gobjBanco.Execute(strSQL, False) Then
                                gobjBanco.ExecutaRollbackTrans
                                Screen.MousePointer = vbDefault
                               'ExibeMensagem "Não foi possível criar o Executivo. A operação foi cancelada."
                               
                                aCritica.AppendRows
                                
                                varCritica = gstrFormataInscricao(Right$(adoResultado("strInscricao"), gintRetornaTamanhoMascara(TYP_IMOBILIARIA))): aCritica(aCritica.UpperBound(1), 0) = varCritica
                                varCritica = adoResultado("strComposicaoDaReceita"): aCritica(aCritica.UpperBound(1), 1) = varCritica
                                varCritica = adoResultado("intExercicio"): aCritica(aCritica.UpperBound(1), 2) = varCritica
                                varCritica = adoResultado("strNumeroAviso"): aCritica(aCritica.UpperBound(1), 3) = varCritica
                               
                                'prg_Status.Visible = False
                                'lbl_Status.Visible = False
                                
                                GoTo ProximoRegistro
                                'Exit Sub
                            End If
                            
                            intNumSequencial = intNumSequencial + 1
                            
                            lngExecutivo = glngRetornaPkidTabelaPai("seqtblExecutivo", gstrExecutivo)
                            
                            'Vamos atualizar a tblDativa
                            strSQL = "UPDATE " & gstrDativa & " SET intExecutivo = " & lngExecutivo & " WHERE intLancamentoAlfa in (" & Mid(strAlfaPorInscr, 1, Len(strAlfaPorInscr) - 1) & ")"
                            If Not gobjBanco.Execute(strSQL, False) Then
                                gobjBanco.ExecutaRollbackTrans
                                Screen.MousePointer = vbDefault
                                'ExibeMensagem "Não foi possível atualizar a Divida Ativa. A operação foi cancelada."
                                
                                'prg_Status.Visible = False
                                'lbl_Status.Visible = False
                                GoTo ProximoRegistro
                                'Exit Sub
                            End If
                            
                            'Vamos zerar os valores
                            dblTotalOriginal = 0
                            dblTotalPrincipal = 0
                            dblTotalMulta = 0
                            dblTotalJuros = 0
                            dblTotalCorrecao = 0
                            
                            strAlfaPorInscr = ""
                            
                        Else
                            'Vamos verificar se sera agrupado o array
                            blnUltimoAlfa = adoResultado("Pkid").Value <> lngInscricaoAtual
                            
                            adoResultado.MovePrevious
                            
                            'Vamos armazenar os pkids dos LancamentosAlfa a serem atualizados na tblDativa
                            If blnUltimoAlfa Then
                                strAlfaPorInscr = strAlfaPorInscr & adoResultado("Pkid").Value & ","
                            End If
                            
                            intNumSequencial2 = intNumSequencial
                            
                        End If
                        
                        'Vamos alimentar o array na troca de composicao e exercicio
                        If blnUltimoAlfa Then
                            
                            'Vamos carregar o array das totalizacoes
                            xadbParcelas.ReDim 0, intPosition, 0, 7

                            xadbParcelas(intPosition, 0) = adoResultado("intComposicaoDaReceita").Value
                            xadbParcelas(intPosition, 1) = adoResultado("intExercicio").Value
                            xadbParcelas(intPosition, 2) = dblTotalGrupo
                            xadbParcelas(intPosition, 3) = adoResultado("strComposicaoDaReceita").Value
                            xadbParcelas(intPosition, 4) = intParcelas
                            xadbParcelas(intPosition, 5) = intNumSequencial2
                            xadbParcelas(intPosition, 6) = adoResultado("strInscricao").Value
                            xadbParcelas(intPosition, 7) = adoResultado("intUtilizacao").Value
                            
                            intParcelas = 0
                            'Vamos zerar os valores
                            dblTotalGrupo = 0
                            
                            blnUltimoAlfa = False
                            
                            intPosition = intPosition + 1
                            
                        End If
                        
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
                    lbl_ContInicial.Caption = prg_Status.Value
                    
                    Me.Refresh
                    DoEvents
                    
                 Next
                    
            End With
            
        Else
            gobjBanco.ExecutaRollbackTrans
            ExibeMensagem "Não foi(ram) encontrado(s) registro(s) com esta Inscrição em Lançamento Alfa, ou a Composição de Receita não inscreve em Divida Ativa."
            Screen.MousePointer = vbDefault
            prg_Status.Visible = False
            lbl_Status.Visible = False
            lbl_ContInicial.Visible = False
            lbl_ContFinal.Visible = False
            
            Exit Sub
        End If
    End If
    
FinalizaOperacao:
    
    If chk_Simulado.Value = vbChecked Then
        gobjBanco.ExecutaRollbackTrans
    Else
        gobjBanco.ExecutaCommitTrans
    End If
    
    strComposicoes = ""
    strExercicios = ""

    For intFor = 0 To xadbComposicoes.UpperBound(1)
        
        strComposicoes = strComposicoes & xadbComposicoes(intFor, 1) & vbNewLine
        strExercicios = strExercicios & xadbComposicoes(intFor, 2) & vbNewLine
    
    Next
    
    rptExecutivosFiscaisSimul.fldComposicoes = strComposicoes
    rptExecutivosFiscaisSimul.fldExercicios = strExercicios
    
    Set xadbParcelas2 = New XArrayDB
    
    xadbParcelas2.ReDim 0, xadbParcelas.UpperBound(1) - 1, 0, 7
    
    For intFor = 0 To xadbParcelas.UpperBound(1) - 1
        
        xadbParcelas2(intFor, 0) = xadbParcelas(intFor, 0)
        xadbParcelas2(intFor, 1) = xadbParcelas(intFor, 1)
        xadbParcelas2(intFor, 2) = xadbParcelas(intFor, 2)
        xadbParcelas2(intFor, 3) = xadbParcelas(intFor, 3)
        xadbParcelas2(intFor, 4) = xadbParcelas(intFor, 4)
        xadbParcelas2(intFor, 5) = xadbParcelas(intFor, 5)
        xadbParcelas2(intFor, 6) = xadbParcelas(intFor, 6)
        xadbParcelas2(intFor, 7) = xadbParcelas(intFor, 7)
        
    Next
    
    xadbParcelas2.QuickSort 0, xadbParcelas2.UpperBound(1), 5, XORDER_ASCEND, XTYPE_INTEGER, 1, XORDER_ASCEND, XTYPE_INTEGER
    
    ImprimeRelatorioPorArray rptExecutivosFiscaisSimul, , "Executivos Fiscais - Cálculo ", , xadbParcelas2, True

    xadbParcelas.QuickSort 0, xadbParcelas.UpperBound(1), 0, XORDER_ASCEND, XTYPE_INTEGER, 1, XORDER_ASCEND, XTYPE_INTEGER
    
    ImprimeRelatorioPorArray rptRelTotalizacaoExecutivo, , "Executivos Fiscais - Cálculo ", , xadbParcelas, True
    
    If aCritica.UpperBound(1) > 0 Then aCritica.DeleteRows 0
    aCritica.QuickSort 0, aCritica.UpperBound(1), 0, XORDER_ASCEND, XTYPE_STRING, 1, XORDER_ASCEND, XTYPE_STRING, 2, XORDER_ASCEND, XTYPE_INTEGER, 3, XORDER_ASCEND, XTYPE_LONG
    
    ImprimeRelatorioPorArray rptRelCriticaExecutivoNaoCalculado, , "Relatório de Executivos Fiscais não calculados", , aCritica, True
    
    Screen.MousePointer = vbDefault
    
    prg_Status.Visible = False
    lbl_Status.Visible = False
    lbl_ContInicial.Visible = False
    lbl_ContFinal.Visible = False
    
    Exit Sub
    
Problema_Na_Rotina:
    ExibeDetalheErro Err.Description
    Resume
    Exit Sub
    
End Sub

Private Function gstrTrataApostrofe(ByVal strValor As String) As String
    gstrTrataApostrofe = Replace(strValor, "'", "''")
End Function




Private Function strQueryIndexador() As String
Dim strSQL As String
    
    strSQL = "SELECT Pkid, strAbreviatura"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrIndexadorEconomico
    strSQL = strSQL & " ORDER BY strAbreviatura"
    
    strQueryIndexador = strSQL

End Function

Private Function strQueryComposicao() As String
Dim strSQL As String

    strSQL = "SELECT CO.Pkid,"
    strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "CO.intCodigo") & strCONCAT & "' - '" & strCONCAT & " CO.strDescricao Descricao"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParametroAtualizacao & " PA, " & gstrComposicaoDaReceita & " CO "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " PA.intComposicaoReceita = CO.Pkid AND "
    strSQL = strSQL & " CO.bytDividaAtiva =  1 "
    strSQL = strSQL & " GROUP BY CO.Pkid, CO.intCodigo, CO.strDescricao"
    strSQL = strSQL & " ORDER BY CO.intCodigo"

    strQueryComposicao = strSQL

End Function

Private Function strQueryExercicio() As String
Dim strSQL As String

    strSQL = "SELECT Pkid, intExercicio "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParametroAtualizacao
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " intComposicaoReceita = " & dbc_intComposicao.BoundText
    strSQL = strSQL & " ORDER BY intExercicio"

    strQueryExercicio = strSQL

End Function

Private Function blnDadosOK() As Boolean
    
    blnDadosOK = False
    
    If txt_intSeqInicial = Space$(0) Then
       MsgBox "O Executivo Inicial deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
       txt_intSeqInicial.SetFocus
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
    
    blnDadosOK = True
       
End Function

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

Private Sub ProximaSequenciaLote(blnReservarNoBanco As Boolean)
Dim adoResultado As New ADODB.Recordset
Dim intFor       As Long
Dim strSQL       As String

    Screen.MousePointer = vbArrowHourglass
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO("SELECT intExecutivo, intLoteExecutivo FROM " & gstrParametrosTributario, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            If Not IsNull(adoResultado("intExecutivo").Value) Then
                txt_intSeqInicial = adoResultado("intExecutivo").Value + 1
            Else
                txt_intSeqInicial = "1"
            End If
            If Not IsNull(adoResultado("intLoteExecutivo").Value) Then
                txt_intLote = adoResultado("intLoteExecutivo").Value + 1
            Else
                txt_intLote = "1"
            End If

        Else
            txt_intSeqInicial = "1"
            txt_intLote = "1"
        End If
    Else
        txt_intSeqInicial = "1"
        txt_intLote = "1"
    End If
    
    'Vamos atualizar o valor das sequencias no banco
    If blnReservarNoBanco Then
        'Vamos fazer a busca com parametros informados para agrupar e obter o numero de inscricoes e serao gravadas
        For intFor = 0 To xadbComposicoes.UpperBound(1)
            'Caso exista mais de 1 composicao
            If Len(Trim(strSQL)) > 0 Then strSQL = strSQL & " UNION "
            
'            strSQL = strSQL & " SELECT LA.strInscricao " & _
                              " FROM " & gstrLancamentoAlfa & " LA, " & _
                              gstrDativa & " DA, " & _
                              gstrLancamentoValor & " LV " & _
                              " WHERE intComposicaoDaReceita = " & xadbComposicoes(intFor, 0) & " AND " & _
                              " intExercicio = " & xadbComposicoes(intFor, 2) & " AND " & _
                              " LV.Intlancamentoalfa = LA.Pkid AND " & _
                              " DA.intLancamentoAlfa = LA.Pkid AND " & _
                              " DA.intExecutivo IS NULL  AND " & _
                              " LV.Pkid not in(Select Intlancamentovalor From tblLancamentoPagamento) AND "
            
            strSQL = strSQL & " SELECT LA.strInscricao "
            
            If bytDBType = SQLServer Then
                 strSQL = strSQL & "FROM " & gstrLancamentoAlfa & " LA "
                 strSQL = strSQL & "INNER JOIN " & gstrLancamentoValor & " LV ON LA.PKId = LV.intLancamentoAlfa "
                 strSQL = strSQL & "INNER JOIN " & gstrDativa & " DA ON LA.PKId = DA.INTLANCAMENTOALFA "
                 strSQL = strSQL & "LEFT OUTER JOIN " & gstrLancamentoPagamento & " LP ON LV.PKId = LP.INTLANCAMENTOVALOR "
                 strSQL = strSQL & "WHERE "
            Else
                 strSQL = strSQL & " FROM " & gstrLancamentoAlfa & " LA, " & _
                 gstrDativa & " DA, " & _
                 gstrLancamentoValor & " LV, " & _
                 gstrLancamentoPagamento & " LP " & _
                 " WHERE LV.Intlancamentoalfa = LA.Pkid AND " & _
                 " DA.intLancamentoAlfa = LA.Pkid AND " & _
                 " LV.Pkid = LP.Intlancamentovalor " & strOUTJOracle & " AND "
            End If
                 
            strSQL = strSQL & " DA.intExecutivo IS /*NOT*/ NULL  AND " & _
                 " " & gstrISNULL("LV.dblValor", "0") & " <> 0 AND " & _
                 " intComposicaoDaReceita = " & xadbComposicoes(intFor, 0) & " AND " & _
                 " intExercicio = " & xadbComposicoes(intFor, 2) & " AND " & _
                 " LP.Intlancamentovalor IS Null AND "
            strSQL = strSQL & " LV.bitParcelaValida = 1 AND "
            
            'Nao vamos considerar as parcelas em acordo
            If Not chk_ExecutarDebitos.Value = vbChecked Then
                strSQL = strSQL & " LV.Intlancamentoalfaacordo is null "
            End If
            'strSql = strSql & " AND LA.strinscricao in ('00000000000002034017' , '00000000000002054028') "
        Next
        
        strSQL = strSQL & " Group By LA.strInscricao "
    
        If gobjBanco.CriaADO(strSQL, 500, adoResultado) Then
            gobjBanco.Execute "UPDATE " & gstrParametrosTributario & " SET intExecutivo = " & txt_intSeqInicial + adoResultado.RecordCount - 1 & ", intLoteExecutivo = " & txt_intLote
        End If
        
    End If
    
    Screen.MousePointer = vbDefault

End Sub



