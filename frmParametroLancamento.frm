VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmParametroLancamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parâmetros para Lançamento"
   ClientHeight    =   7020
   ClientLeft      =   1785
   ClientTop       =   3075
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   8880
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5280
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   9313
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros para Lançamento"
      TabPicture(0)   =   "frmParametroLancamento.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Exercicio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Emissao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_intNumAvisoInicial"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_intNumAvisoFinal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tab_PagamentoParcelas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdb_Pagamentos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_ComposicaoDaReceita"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_intExercicio"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_strEmissao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_intNumAvisoInicial"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_intNumAvisoFinal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPkidParametros"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPkidPagamentos"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPkidParcelas"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.TextBox txtPkidParcelas 
         Height          =   315
         Left            =   5250
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtPkidPagamentos 
         Height          =   315
         Left            =   3780
         TabIndex        =   52
         Top             =   -15
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtPkidParametros 
         Height          =   315
         Left            =   2490
         TabIndex        =   51
         Top             =   -15
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txt_intNumAvisoFinal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox txt_intNumAvisoInicial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4485
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1260
         Width           =   1260
      End
      Begin VB.TextBox txt_strEmissao 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2535
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1260
         Width           =   675
      End
      Begin VB.TextBox txt_intExercicio 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1260
         Width           =   705
      End
      Begin VB.Frame fra_ComposicaoDaReceita 
         Caption         =   "Composição da Receita"
         Height          =   780
         Left            =   60
         TabIndex        =   1
         Top             =   405
         Width           =   8595
         Begin VB.CommandButton cmd_Composicao 
            Height          =   300
            Left            =   5850
            Picture         =   "frmParametroLancamento.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro de Composição da Receita"
            Top             =   315
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbc_intComposicao 
            Height          =   315
            Left            =   1020
            TabIndex        =   3
            Top             =   315
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_Composicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   90
            TabIndex        =   2
            Top             =   360
            Width           =   870
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Pagamentos 
         Height          =   1230
         Left            =   120
         TabIndex        =   49
         Top             =   3930
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   2170
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKID"
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Título"
         Columns(1).DataField=   "Titulo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Desconto"
         Columns(2).DataField=   "Desconto"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Desconto Taxas"
         Columns(3).DataField=   "intDesctoTaxas"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Cód. Moeda"
         Columns(4).DataField=   "CodMoeda"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Valor Moeda"
         Columns(5).DataField=   "ValorMoeda"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   16
         Columns(6)._MaxComboItems=   5
         Columns(6).ValueItems(0)._DefaultItem=   0
         Columns(6).ValueItems(0).Value=   "0"
         Columns(6).ValueItems(0).Value.vt=   8
         Columns(6).ValueItems(0).DisplayValue=   "Não"
         Columns(6).ValueItems(0).DisplayValue.vt=   8
         Columns(6).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems(1)._DefaultItem=   0
         Columns(6).ValueItems(1).Value=   "1"
         Columns(6).ValueItems(1).Value.vt=   8
         Columns(6).ValueItems(1).DisplayValue=   "Sim"
         Columns(6).ValueItems(1).DisplayValue.vt=   8
         Columns(6).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems.Count=   2
         Columns(6).Caption=   "Parcela Principal"
         Columns(6).DataField=   "bytParcelado"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=3201"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3122"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2090"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2011"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2461"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2381"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2011"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1931"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=2487"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2408"
         Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=2328"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2249"
         Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
         _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
         _StyleDefs(31)  =   ":id=18,.fgcolor=&H8000000E&"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
         _StyleDefs(34)  =   ":id=19,.fgcolor=&H8000000E&"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(67)  =   "Named:id=33:Normal"
         _StyleDefs(68)  =   ":id=33,.parent=0"
         _StyleDefs(69)  =   "Named:id=34:Heading"
         _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   ":id=34,.wraptext=-1"
         _StyleDefs(72)  =   "Named:id=35:Footing"
         _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(74)  =   "Named:id=36:Selected"
         _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(76)  =   "Named:id=37:Caption"
         _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(78)  =   "Named:id=38:HighlightRow"
         _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   "Named:id=39:EvenRow"
         _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(82)  =   "Named:id=40:OddRow"
         _StyleDefs(83)  =   ":id=40,.parent=33"
         _StyleDefs(84)  =   "Named:id=41:RecordSelector"
         _StyleDefs(85)  =   ":id=41,.parent=34"
         _StyleDefs(86)  =   "Named:id=42:FilterBar"
         _StyleDefs(87)  =   ":id=42,.parent=33"
      End
      Begin TabDlg.SSTab tab_PagamentoParcelas 
         Height          =   2205
         Left            =   105
         TabIndex        =   13
         Top             =   1650
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   3889
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Formas de Pagamento"
         TabPicture(0)   =   "frmParametroLancamento.frx":013A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_FormaDePagamento"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Parcelas"
         TabPicture(1)   =   "frmParametroLancamento.frx":0156
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Parcelas a cancelar"
         TabPicture(2)   =   "frmParametroLancamento.frx":0172
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame2 
            Height          =   1770
            Left            =   -74940
            TabIndex        =   36
            Top             =   315
            Width           =   8415
            Begin VB.CommandButton cmd_CodigoBaixa 
               Height          =   300
               Left            =   4680
               Picture         =   "frmParametroLancamento.frx":018E
               Style           =   1  'Graphical
               TabIndex        =   47
               TabStop         =   0   'False
               Tag             =   "590"
               ToolTipText     =   "Ativa Cadastro de Composição da Receita"
               Top             =   1200
               Width           =   360
            End
            Begin MSDataListLib.DataCombo dbc_intCodigoBaixa 
               Height          =   315
               Left            =   3240
               TabIndex        =   46
               Top             =   1200
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.TextBox txt_dtmDtVencimento3 
               Height          =   315
               Left            =   1050
               TabIndex        =   44
               Top             =   1170
               Width           =   975
            End
            Begin VB.TextBox txt_strTitulo3 
               Height          =   315
               Left            =   1050
               TabIndex        =   38
               Top             =   240
               Width           =   4350
            End
            Begin TrueOleDBGrid70.TDBGrid tdb_ParcelaCancelamento 
               Height          =   1410
               Left            =   5550
               TabIndex        =   48
               Top             =   240
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   2487
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Pkid"
               Columns(0).DataField=   "Pkid"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Parcelas"
               Columns(1).DataField=   "intParcela"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Vencimento"
               Columns(2).DataField=   "dtmDtVencimento"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).ScrollBars=   2
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=1931"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1852"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(13)=   "Column(2).Width=2355"
               Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2275"
               Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
               Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(21)  =   ":id=8,.fgcolor=&H8000000E&"
               _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
               _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
               _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
               _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
               _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
               _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(52)  =   "Named:id=33:Normal"
               _StyleDefs(53)  =   ":id=33,.parent=0"
               _StyleDefs(54)  =   "Named:id=34:Heading"
               _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(56)  =   ":id=34,.wraptext=-1"
               _StyleDefs(57)  =   "Named:id=35:Footing"
               _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(59)  =   "Named:id=36:Selected"
               _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(61)  =   "Named:id=37:Caption"
               _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(63)  =   "Named:id=38:HighlightRow"
               _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
               _StyleDefs(65)  =   "Named:id=39:EvenRow"
               _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(67)  =   "Named:id=40:OddRow"
               _StyleDefs(68)  =   ":id=40,.parent=33"
               _StyleDefs(69)  =   "Named:id=41:RecordSelector"
               _StyleDefs(70)  =   ":id=41,.parent=34"
               _StyleDefs(71)  =   "Named:id=42:FilterBar"
               _StyleDefs(72)  =   ":id=42,.parent=33"
            End
            Begin MSDataListLib.DataCombo dbc_intParcela 
               Height          =   315
               Left            =   1050
               TabIndex        =   40
               Top             =   720
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dbc_intParcelaBaixa 
               Height          =   315
               Left            =   4200
               TabIndex        =   42
               Top             =   660
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label Label1 
               Caption         =   "Tipo de baixa"
               Height          =   255
               Left            =   2160
               TabIndex        =   45
               Top             =   1260
               Width           =   1095
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Parcela a baixar"
               Height          =   195
               Left            =   3000
               TabIndex        =   41
               Top             =   750
               Width           =   1140
            End
            Begin VB.Label lbl_vencimento3 
               AutoSize        =   -1  'True
               Caption         =   "Vencimento"
               Height          =   195
               Left            =   120
               TabIndex        =   43
               Top             =   1275
               Width           =   840
            End
            Begin VB.Label lbl_parcela3 
               AutoSize        =   -1  'True
               Caption         =   "Parcela"
               Height          =   195
               Left            =   420
               TabIndex        =   39
               Top             =   795
               Width           =   540
            End
            Begin VB.Label lbl_strTitulo3 
               AutoSize        =   -1  'True
               Caption         =   "Título"
               Height          =   195
               Left            =   540
               TabIndex        =   37
               Top             =   360
               Width           =   420
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1770
            Left            =   -74940
            TabIndex        =   28
            Top             =   315
            Width           =   8415
            Begin VB.TextBox txt_strTitulo2 
               Height          =   315
               Left            =   735
               TabIndex        =   30
               Top             =   240
               Width           =   4590
            End
            Begin VB.TextBox txt_intParcelas 
               Height          =   315
               Left            =   735
               TabIndex        =   32
               Top             =   690
               Width           =   1425
            End
            Begin VB.TextBox txt_dtmDtVencimento 
               Height          =   315
               Left            =   3330
               TabIndex        =   34
               Top             =   690
               Width           =   1215
            End
            Begin TrueOleDBGrid70.TDBGrid tdb_ParcelaVencimento 
               Height          =   1410
               Left            =   5550
               TabIndex        =   35
               Top             =   240
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   2487
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Parcelas"
               Columns(0).DataField=   "intParcelas"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Vencimento"
               Columns(1).DataField=   "dtmDtVencimento"
               Columns(1).NumberFormat=   "Short Date"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "pkid"
               Columns(2).DataField=   "pkid"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).ScrollBars=   2
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2381"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2302"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
               Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
               Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(17)=   "Column(2).Visible=0"
               Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(21)  =   ":id=8,.fgcolor=&H8000000E&"
               _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
               _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
               _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
               _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
               _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
               _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
               _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(52)  =   "Named:id=33:Normal"
               _StyleDefs(53)  =   ":id=33,.parent=0"
               _StyleDefs(54)  =   "Named:id=34:Heading"
               _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(56)  =   ":id=34,.wraptext=-1"
               _StyleDefs(57)  =   "Named:id=35:Footing"
               _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(59)  =   "Named:id=36:Selected"
               _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(61)  =   "Named:id=37:Caption"
               _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(63)  =   "Named:id=38:HighlightRow"
               _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
               _StyleDefs(65)  =   "Named:id=39:EvenRow"
               _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(67)  =   "Named:id=40:OddRow"
               _StyleDefs(68)  =   ":id=40,.parent=33"
               _StyleDefs(69)  =   "Named:id=41:RecordSelector"
               _StyleDefs(70)  =   ":id=41,.parent=34"
               _StyleDefs(71)  =   "Named:id=42:FilterBar"
               _StyleDefs(72)  =   ":id=42,.parent=33"
            End
            Begin VB.Label lbl_strTitulo2 
               AutoSize        =   -1  'True
               Caption         =   "Título"
               Height          =   195
               Left            =   240
               TabIndex        =   29
               Top             =   360
               Width           =   420
            End
            Begin VB.Label lbl_parcela2 
               AutoSize        =   -1  'True
               Caption         =   "Parcela"
               Height          =   195
               Left            =   105
               TabIndex        =   31
               Top             =   795
               Width           =   540
            End
            Begin VB.Label lbl_vencimento2 
               AutoSize        =   -1  'True
               Caption         =   "Vencimento"
               Height          =   195
               Left            =   2340
               TabIndex        =   33
               Top             =   795
               Width           =   840
            End
         End
         Begin VB.Frame fra_FormaDePagamento 
            Height          =   1770
            Left            =   60
            TabIndex        =   14
            Top             =   315
            Width           =   8415
            Begin VB.TextBox txt_intDesctoTaxas 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4500
               TabIndex        =   19
               Top             =   630
               Width           =   795
            End
            Begin VB.CheckBox chk_bytParcelado 
               Caption         =   "Parcelamento Principal"
               Height          =   255
               Left            =   1575
               TabIndex        =   26
               Top             =   1410
               Width           =   1935
            End
            Begin VB.TextBox txt_strTitulo 
               Height          =   315
               Left            =   1590
               TabIndex        =   16
               Top             =   240
               Width           =   3720
            End
            Begin VB.TextBox txt_dblDesconto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1590
               TabIndex        =   18
               Top             =   630
               Width           =   705
            End
            Begin VB.TextBox txt_dblValorMoeda 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4350
               TabIndex        =   25
               Top             =   1020
               Width           =   945
            End
            Begin VB.CommandButton cmd_CodMoeda 
               Height          =   300
               Left            =   2475
               Picture         =   "frmParametroLancamento.frx":02AC
               Style           =   1  'Graphical
               TabIndex        =   23
               TabStop         =   0   'False
               Tag             =   "590"
               ToolTipText     =   "Indexador Econônico."
               Top             =   1020
               Width           =   360
            End
            Begin MSDataListLib.DataCombo dbc_strCodMoeda 
               Height          =   315
               Left            =   1575
               TabIndex        =   22
               Top             =   1020
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   556
               _Version        =   393216
               IntegralHeight  =   0   'False
               Text            =   ""
            End
            Begin TrueOleDBGrid70.TDBGrid tdb_ParcelaVencimento2 
               Height          =   1410
               Left            =   5535
               TabIndex        =   27
               Top             =   240
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   2487
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Parcelas"
               Columns(0).DataField=   "intParcelas"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Vencimento"
               Columns(1).DataField=   "dtmDtVencimento"
               Columns(1).NumberFormat=   "Short Date"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "pkid"
               Columns(2).DataField=   "pkid"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).MarqueeStyle=   3
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).ScrollBars=   2
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2381"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2302"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
               Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
               Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(17)=   "Column(2).Visible=0"
               Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(21)  =   ":id=8,.fgcolor=&H8000000E&"
               _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
               _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
               _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
               _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
               _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
               _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
               _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
               _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(52)  =   "Named:id=33:Normal"
               _StyleDefs(53)  =   ":id=33,.parent=0"
               _StyleDefs(54)  =   "Named:id=34:Heading"
               _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(56)  =   ":id=34,.wraptext=-1"
               _StyleDefs(57)  =   "Named:id=35:Footing"
               _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(59)  =   "Named:id=36:Selected"
               _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(61)  =   "Named:id=37:Caption"
               _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(63)  =   "Named:id=38:HighlightRow"
               _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
               _StyleDefs(65)  =   "Named:id=39:EvenRow"
               _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(67)  =   "Named:id=40:OddRow"
               _StyleDefs(68)  =   ":id=40,.parent=33"
               _StyleDefs(69)  =   "Named:id=41:RecordSelector"
               _StyleDefs(70)  =   ":id=41,.parent=34"
               _StyleDefs(71)  =   "Named:id=42:FilterBar"
               _StyleDefs(72)  =   ":id=42,.parent=33"
            End
            Begin VB.Label lbl_intDesctoTaxas 
               Caption         =   "Desconto Taxas"
               Height          =   225
               Left            =   3240
               TabIndex        =   54
               Top             =   720
               Width           =   1185
            End
            Begin VB.Label lbl_Porcentagem 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Left            =   2370
               TabIndex        =   20
               Top             =   690
               Width           =   120
            End
            Begin VB.Label lbl_strTitulo 
               AutoSize        =   -1  'True
               Caption         =   "Título"
               Height          =   195
               Left            =   1035
               TabIndex        =   15
               Top             =   360
               Width           =   420
            End
            Begin VB.Label lbl_dblDesconto 
               AutoSize        =   -1  'True
               Caption         =   "Desconto Impostos"
               Height          =   195
               Left            =   105
               TabIndex        =   17
               Top             =   720
               Width           =   1365
            End
            Begin VB.Label lbl_CodMoeda 
               AutoSize        =   -1  'True
               Caption         =   "Indexador"
               Height          =   195
               Left            =   810
               TabIndex        =   21
               Top             =   1125
               Width           =   705
            End
            Begin VB.Label lbl_dblValorMoeda 
               AutoSize        =   -1  'True
               Caption         =   "Valor do Indexador"
               Height          =   195
               Left            =   2940
               TabIndex        =   24
               Top             =   1095
               Width           =   1335
            End
         End
      End
      Begin VB.Label lbl_intNumAvisoFinal 
         AutoSize        =   -1  'True
         Caption         =   "Nº Aviso Final"
         Height          =   195
         Left            =   5850
         TabIndex        =   11
         Top             =   1365
         Width           =   990
      End
      Begin VB.Label lbl_intNumAvisoInicial 
         AutoSize        =   -1  'True
         Caption         =   "Nº Aviso Inicial"
         Height          =   195
         Left            =   3300
         TabIndex        =   9
         Top             =   1365
         Width           =   1065
      End
      Begin VB.Label lbl_Emissao 
         AutoSize        =   -1  'True
         Caption         =   "Emissão"
         Height          =   195
         Left            =   1860
         TabIndex        =   7
         Top             =   1365
         Width           =   585
      End
      Begin VB.Label lbl_Exercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   330
         TabIndex        =   5
         Top             =   1365
         Width           =   675
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Parametros 
      Height          =   1620
      Left            =   45
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5370
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   2858
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "PKID"
      Columns(0).DataField=   "PKID"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Composição"
      Columns(1).DataField=   "Composicao"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Exercício"
      Columns(2).DataField=   "Exercicio"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Emissão"
      Columns(3).DataField=   "Emissao"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nº Aviso Inicial"
      Columns(4).DataField=   "NumAvisoInicial"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Nº Aviso Final"
      Columns(5).DataField=   "NumAvisoFinal"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=7514"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7435"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1482"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1402"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1535"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1455"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2223"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2143"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1984"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1905"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(31)  =   ":id=18,.fgcolor=&H8000000E&"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(34)  =   ":id=19,.fgcolor=&H8000000E&"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(63)  =   "Named:id=33:Normal"
      _StyleDefs(64)  =   ":id=33,.parent=0"
      _StyleDefs(65)  =   "Named:id=34:Heading"
      _StyleDefs(66)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   ":id=34,.wraptext=-1"
      _StyleDefs(68)  =   "Named:id=35:Footing"
      _StyleDefs(69)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   "Named:id=36:Selected"
      _StyleDefs(71)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(72)  =   "Named:id=37:Caption"
      _StyleDefs(73)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(74)  =   "Named:id=38:HighlightRow"
      _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(76)  =   "Named:id=39:EvenRow"
      _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(78)  =   "Named:id=40:OddRow"
      _StyleDefs(79)  =   ":id=40,.parent=33"
      _StyleDefs(80)  =   "Named:id=41:RecordSelector"
      _StyleDefs(81)  =   ":id=41,.parent=34"
      _StyleDefs(82)  =   "Named:id=42:FilterBar"
      _StyleDefs(83)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmParametroLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnClickGridParametros      As Boolean
Dim blnClickGridPagamentos      As Boolean
Dim blnClickGridParcelas        As Boolean
Dim blnAlterandoParcela         As Boolean
Dim blnAlterandoFormaPagto      As Boolean
Dim blnNovaParcela              As Boolean
Dim intPosicaoCursorParametro   As Integer
Dim blnTabPagamentosParcelas    As Boolean
Dim blnParametros               As Boolean
Dim tdbGrdSelecionada           As TrueOleDBGrid70.TDBGrid
Dim vetTag()                    As String
Dim strTituloAtual              As String
Dim strParcelaAtual             As String
Dim strEmissaoAtual             As String
Dim blnDadosCancelamento        As Boolean

Private Sub cmd_CodigoBaixa_Click()
    ChamaFormCadastro frmCodigosDeBaixa, dbc_intCodigoBaixa
End Sub

Private Sub cmd_CodMoeda_Click()
    ChamaFormCadastro frmIndexadorEconomico, dbc_strCodMoeda
    blnTabPagamentosParcelas = True
End Sub

Private Sub cmd_Composicao_Click()
    ChamaFormCadastro frmCadComposicaoDaReceita, dbc_intComposicao
    blnTabPagamentosParcelas = False
End Sub

Private Sub dbc_intComposicao_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicao, Me, Area
End Sub

Private Sub dbc_intComposicao_GotFocus()
    MarcaCampo dbc_intComposicao
    blnTabPagamentosParcelas = False
End Sub

Private Sub dbc_intComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicao
End Sub

Private Sub dbc_intParcela_Click(Area As Integer)
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    
    blnTabPagamentosParcelas = True
    
    If Area = 2 Then 'dbc_intParcela.MatchedWithList
        strSql = "SELECT "
        strSql = strSql & " FV.dtmDtVencimento"
        strSql = strSql & " FROM "
        strSql = strSql & gstrFormaPagtoVencimentos & " FV"
        strSql = strSql & " WHERE FV.Pkid ='" & dbc_intParcela.BoundText & "'"
        strSql = strSql & " ORDER BY FV.dtmDtVencimento "
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            txt_dtmDtVencimento3.Text = gstrENulo(adoResultado!Dtmdtvencimento)
        End If
        LeDaTabelaParaObj "", dbc_intParcelaBaixa, strQueryParcelaCancelar
        LeDaTabelaParaObj "", tdb_ParcelaCancelamento, strQueryParcelaCancelarGrid
    End If
    
End Sub

Private Sub dbc_intParcela_GotFocus()
    blnTabPagamentosParcelas = True
End Sub

Private Sub dbc_intParcelaBaixa_Click(Area As Integer)
    blnTabPagamentosParcelas = True
End Sub

Private Sub dbc_strCodMoeda_Click(Area As Integer)
    DropDownDataCombo dbc_strCodMoeda, Me, Area
    blnTabPagamentosParcelas = True
End Sub

Private Sub dbc_strCodMoeda_GotFocus()
    MarcaCampo dbc_strCodMoeda
End Sub

Private Sub dbc_strCodMoeda_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strCodMoeda, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strCodMoeda_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strCodMoeda
End Sub

Private Sub dbc_strCodMoeda_LostFocus()
    blnTabPagamentosParcelas = False
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1246
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar, gstrDeletar
    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"
    dbc_strCodMoeda.Tag = strQueryCodMoeda & ";strAbreviatura"
    dbc_intCodigoBaixa.Tag = ""
    TrocaCorObjeto txt_strTitulo2, True
    TrocaCorObjeto txt_strTitulo3, True
    TrocaCorObjeto txt_dtmDtVencimento3, True
    dbc_intCodigoBaixa.Tag = "SELECT pkid, strAbreviatura FROM " & gstrCodigoDeBaixa & " WHERE bytTipo = 2;strAbreviatura "
    
    'Cláudio - Tag dos objetos grid para saber qual tabela haverá a exclusão
    tdb_Parametros.Tag = gstrParametroIPTU & ";Parametros"
    tdb_Pagamentos.Tag = gstrParametroIPTUPagto & ";Forma de Pagamento"
    tdb_ParcelaVencimento.Tag = gstrFormaPagtoVencimentos & ";Parcela/Vencimento"
    tdb_ParcelaCancelamento.Tag = gstrFormaPagtoCancelamentos & ";Parcela/Cancelamento"
    
    blnNovaParcela = False
    blnAlterandoFormaPagto = False
    blnAlterandoParcela = False
    blnClickGridParametros = False
    blnClickGridPagamentos = False
    blnClickGridParcelas = False
    
    chk_bytParcelado.Value = 1
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If txtPkidParametros.Text <> "" Then
        If tdb_Pagamentos.EOF Then
            ExibeMensagem "Dados incompletos para Parâmetros."
            Cancel = 1
        Else
            HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
        End If
    End If
    
End Sub

Private Sub fra_ComposicaoDaReceita_Click()
    blnTabPagamentosParcelas = False
End Sub

Private Sub fra_FormaDePagamento_Click()
    blnTabPagamentosParcelas = True
End Sub

Private Sub tab_3DPasta_GotFocus()
    blnTabPagamentosParcelas = False
End Sub

Private Sub tab_PagamentoParcelas_Click(PreviousTab As Integer)
    blnTabPagamentosParcelas = True
'    If tab_PagamentoParcelas.Tab = 1 Then
'        MantemForm gstrNovo
'    ElseIf tab_PagamentoParcelas.Tab = 0 Then
'        blnNovaParcela = False
'        blnAlterandoParcela = False
'    End If
End Sub

Private Sub tab_PagamentoParcelas_GotFocus()
    blnTabPagamentosParcelas = True
End Sub

Private Sub tab_PagamentoParcelas_LostFocus()
    blnTabPagamentosParcelas = False
End Sub

Private Sub tdb_Pagamentos_BeforeRowColChange(Cancel As Integer)
    If txtPkidParametros.Text <> "" And txtPkidPagamentos.Text <> "" Then
        If tdb_ParcelaVencimento.EOF Then
            ExibeMensagem "Deve existir no mínimo uma Parcela/Pagamento."
            Cancel = 1
        End If
    End If
End Sub

Private Sub tdb_Pagamentos_Click()
    
    blnClickGridParametros = True
    blnClickGridPagamentos = True
    blnClickGridParcelas = False
    blnAlterandoParcela = False
    blnAlterandoFormaPagto = True
    
    blnTabPagamentosParcelas = True
    
    blnNovaParcela = False
    
    Set tdbGrdSelecionada = tdb_Pagamentos
    
    tdb_Pagamentos_RowColChange 0, 0
    
End Sub

Private Sub tdb_Pagamentos_FilterChange()
    gblnFilraCampos tdb_Pagamentos
End Sub

Private Sub tdb_Pagamentos_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Pagamentos, ColIndex
End Sub

Private Sub tdb_Pagamentos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If Not tdb_Pagamentos.EOF And blnClickGridPagamentos Then
    txtPkidPagamentos.Text = tdb_Pagamentos.Columns("Pkid").Value
    chk_bytParcelado = IIf(gstrENulo(tdb_Pagamentos.Columns("bytParcelado").Value) <> "", tdb_Pagamentos.Columns("bytParcelado").Value, 0)
    PreencheCamposPagamentos Val(txtPkidPagamentos.Text)
    PreencheGrdParcelas Val(txtPkidPagamentos.Text)
    LeDaTabelaParaObj "", dbc_intParcela, strQueryParcela
    Set dbc_intParcelaBaixa.RowSource = Nothing
    dbc_intParcelaBaixa.Text = ""
    blnClickGridPagamentos = True
    blnTabPagamentosParcelas = True
    'Não deve preencher na tela até que o usuário clique no grid
    'tdb_ParcelaVencimento_RowColChange 0, 0
    strTituloAtual = tdb_Pagamentos.Columns("Título").Value
    strEmissaoAtual = tdb_Parametros.Columns("Emissão").Value
    If tab_PagamentoParcelas.Tab = 1 Then
        MantemForm gstrNovo
    End If
Else
    txt_strTitulo.Text = ""
    txt_strTitulo2.Text = ""
    txt_strTitulo3.Text = ""
    txt_dblDesconto.Text = ""
    dbc_strCodMoeda.Text = ""
    txt_dblValorMoeda.Text = ""
    txt_intDesctoTaxas.Text = ""
End If
    
End Sub

Private Sub tdb_Parametros_BeforeRowColChange(Cancel As Integer)
    If txtPkidParametros.Text <> "" Then
        If tdb_Pagamentos.EOF Then
            ExibeMensagem "Dados incompletos para Parâmetros."
            Cancel = 1
        End If
    End If
End Sub

Private Sub tdb_Parametros_Click()
    
    blnClickGridParametros = True
    blnClickGridPagamentos = False
    blnClickGridParcelas = False
    blnAlterandoParcela = False
    blnAlterandoFormaPagto = False
   
    Set tdbGrdSelecionada = tdb_Parametros
End Sub

Private Sub tdb_Parametros_FilterChange()
    gblnFilraCampos tdb_Parametros
End Sub

Private Sub tdb_Parametros_GotFocus()
    blnTabPagamentosParcelas = False
End Sub

Private Sub tdb_Parametros_HeadClick(ByVal ColIndex As Integer)
    blnClickGridParametros = False
    gOrdenaGrid tdb_Parametros, ColIndex
End Sub

Private Sub tdb_Parametros_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tdb_Parametros.EOF Then
        If blnClickGridParametros Then
            limpaGrds
            txtPkidParametros.Text = tdb_Parametros.Columns("Pkid").Value
            PreencheCamposParametros Val(txtPkidParametros.Text)
            PreencheGrdPagamentos Val(txtPkidParametros.Text)
            txt_dtmDtVencimento3.Text = ""
            Set tdb_ParcelaCancelamento.DataSource = Nothing
            txtPkidPagamentos.Text = ""
            txtPkidParcelas = ""
        End If
    End If

End Sub

Private Sub tdb_ParcelaVencimento_Click()
    blnNovaParcela = False
    blnAlterandoParcela = True
    blnClickGridParcelas = True
    blnTabPagamentosParcelas = True
    Set tdbGrdSelecionada = tdb_ParcelaVencimento
    tdb_ParcelaVencimento_RowColChange 0, 0
End Sub

Private Sub tdb_ParcelaVencimento_FilterChange()
    gblnFilraCampos tdb_ParcelaVencimento
End Sub

Private Sub tdb_ParcelaVencimento_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tdb_ParcelaVencimento.EOF And blnClickGridParcelas Then
        txtPkidParcelas.Text = Val(tdb_ParcelaVencimento.Columns("Pkid").Value)
        txt_intParcelas.Text = Val(tdb_ParcelaVencimento.Columns("Parcelas").Value)
        txt_dtmDtVencimento.Text = tdb_ParcelaVencimento.Columns("Vencimento").Value
        strParcelaAtual = tdb_ParcelaVencimento.Columns("Parcelas").Value
        blnClickGridParcelas = True
        blnTabPagamentosParcelas = True
    Else
        LimpaTabParcela
        blnClickGridParcelas = False
    End If
End Sub

Private Sub tdb_ParcelaCancelamento_Click()
    blnTabPagamentosParcelas = True
    Set tdbGrdSelecionada = tdb_ParcelaCancelamento
End Sub

Private Sub tdb_ParcelaVencimento2_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_ParcelaVencimento2, ColIndex
End Sub

Private Sub txt_dblDesconto_GotFocus()
    MarcaCampo txt_dblDesconto
    blnTabPagamentosParcelas = True
End Sub

Private Sub txt_dblDesconto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblDesconto
End Sub

Private Sub txt_dblDesconto_LostFocus()
    txt_dblDesconto.Text = gstrConvVrDoSql(txt_dblDesconto.Text, 2)
    blnTabPagamentosParcelas = False
End Sub

Private Sub txt_dblValorMoeda_GotFocus()
    MarcaCampo txt_dblValorMoeda
    blnTabPagamentosParcelas = True
End Sub

Private Sub txt_dblValorMoeda_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorMoeda
End Sub

Private Sub txt_dtmDtVencimento_GotFocus()
    MarcaCampo txt_dtmDtVencimento
    blnTabPagamentosParcelas = True
End Sub

Private Sub txt_dtmDtVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDtVencimento
End Sub

Private Sub txt_dblValorMoeda_LostFocus()
    blnTabPagamentosParcelas = False
    txt_dblValorMoeda.Text = gstrConvVrDoSql(txt_dblValorMoeda.Text, 4)
End Sub

Private Sub txt_dtmDtVencimento_LostFocus()
    txt_dtmDtVencimento.Text = gstrDataFormatada(txt_dtmDtVencimento.Text)
End Sub

Private Sub txt_intDesctoTaxas_GotFocus()
    MarcaCampo txt_intDesctoTaxas
End Sub

Private Sub txt_intDesctoTaxas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_intDesctoTaxas
End Sub

Private Sub txt_intExercicio_GotFocus()
    If txt_intExercicio.Text = "" Then
        txt_intExercicio = gintExercicio
    End If
    MarcaCampo txt_intExercicio
    blnTabPagamentosParcelas = False
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_intNumAvisoFinal_GotFocus()
    MarcaCampo txt_intNumAvisoFinal
    blnTabPagamentosParcelas = False
End Sub

Private Sub txt_intNumAvisoFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intNumAvisoFinal
End Sub

Private Sub txt_intNumAvisoInicial_GotFocus()
    MarcaCampo txt_intNumAvisoInicial
    blnTabPagamentosParcelas = False
End Sub

Private Sub txt_intNumAvisoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intNumAvisoInicial
End Sub

Private Sub txt_intParcelas_GotFocus()
    MarcaCampo txt_intParcelas
    blnTabPagamentosParcelas = True
    tab_PagamentoParcelas.Tab = 1
End Sub

Private Sub txt_intParcelas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelas
End Sub

Private Sub txt_strEmissao_GotFocus()
    MarcaCampo txt_strEmissao
    blnTabPagamentosParcelas = False
End Sub

Private Sub txt_strEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strEmissao
End Sub

Private Function strQueryComposicao() As String
Dim strSql As String

strSql = "SELECT Pkid,"
strSql = strSql & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
strSql = strSql & " FROM "
strSql = strSql & gstrComposicaoDaReceita
strSql = strSql & " ORDER BY intCodigo"

strQueryComposicao = strSql

End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark             As Variant
Dim blnCriticarCancelamento As Boolean

    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        Case Is = UCase(gstrImprimir)
            ImprimeRelatorio rptParametrosLancamento, strQueryRelatorio
            Exit Sub
        Case Is = UCase(gstrLocalizar)
            LeDaTabelaParaObj "", tdb_Parametros, strQueryParametros(False)
            If Not tdb_Parametros.EOF Then
               blnClickGridParametros = True
               tdb_Parametros_RowColChange 0, 0
            End If
        Case Is = UCase(gstrSalvar)
            If blnDadosOk Then
                If gblnExclusaoGravacaoOk("" & IIf(txtPkidParametros.Text <> "", "A", "I")) Then
                    Set gobjBanco = New clsBanco
                    
                    gobjBanco.ExecutaBeginTrans
                    
                    If Not GravaParametros(blnClickGridParametros) Then Exit Sub
                    
                    If txtPkidParametros.Text = "" Then
                        GravaPagamentos IIf(Val(txtPkidPagamentos.Text) <> 0, True, False), PegaUltimoPkidParametro
                    Else
                        GravaPagamentos IIf(Val(txtPkidPagamentos.Text) <> 0, True, False), Val(txtPkidParametros.Text)
                    End If
                    
                    blnCriticarCancelamento = chk_bytParcelado = 0 And blnNovaParcela
                    
                    If (blnAlterandoFormaPagto And blnAlterandoParcela) _
                        Or (Not blnAlterandoFormaPagto And Not blnAlterandoParcela) _
                        Or (blnAlterandoFormaPagto And blnNovaParcela And tab_PagamentoParcelas.Tab = 1) Then
                    
                        If txtPkidPagamentos.Text = "" Then
                            GravaParcelas blnClickGridParcelas, PegaUltimoPkidPagamentos
                            txtPkidPagamentos.Text = PegaUltimoPkidPagamentos
                            strTituloAtual = txt_strTitulo.Text
                        Else
                            GravaParcelas blnClickGridParcelas, Val(txtPkidPagamentos.Text)
                        End If
                    
                        blnNovaParcela = False
                    
                    End If
                    
                    If blnDadosCancelamento Then
                        GravaParcelasCanceladas
                    End If
                    
                    gobjBanco.ExecutaCommitTrans
                    
                    PreencheGrdParcelas Val(txtPkidPagamentos.Text)
                    
                    If blnDadosCancelamento Then
                        blnDadosCancelamento = False
                        LeDaTabelaParaObj "", dbc_intParcelaBaixa, strQueryParcelaCancelar
                        LeDaTabelaParaObj "", tdb_ParcelaCancelamento, strQueryParcelaCancelarGrid
                    End If
                        
                    If Val(txtPkidParametros.Text) = 0 Then
                        MantemForm gstrRefresh
                        blnDadosCancelamento = False
                        blnTabPagamentosParcelas = True
                        MantemForm gstrNovo
                        PreencheGrdPagamentos Val(txtPkidParametros.Text)
                    End If
                    
                End If
                
                If Not blnTabPagamentosParcelas And tab_PagamentoParcelas.Tab = 1 Then
                    Limpa_Controles Me, True, False, False, True, False
                    limpaGrds
                    Set tdb_Parametros.DataSource = Nothing
                    LeDaTabelaParaObj "", tdb_Parametros, strQueryParametros(True)
                End If
                
               If tab_PagamentoParcelas.Tab = 0 Then
                    blnClickGridParametros = True
                    blnClickGridPagamentos = True
                    blnClickGridParcelas = False
                    blnAlterandoFormaPagto = True
                    blnAlterandoParcela = False
                    
                ElseIf tab_PagamentoParcelas.Tab = 1 Then

                    txt_intParcelas.Text = ""
                    txt_dtmDtVencimento.Text = ""
                    txt_intParcelas.SetFocus
                        
                    blnAlterandoFormaPagto = True
                    blnAlterandoParcela = False
                        
                End If
                                
                'Vamos posicionar no registro em que esta posicionado
                varBookMark = tdb_Pagamentos.Bookmark

                LeDaTabelaParaObj "", tdb_Parametros, strQueryParametros(False)
                
                If txtPkidParametros <> "" Then PreencheGrdPagamentos txtPkidParametros
                
                DoEvents
                If Len(Str(varBookMark)) > 0 Then
                    tdb_Pagamentos.Bookmark = varBookMark
                End If
                tdb_Pagamentos_Click
                
                'txtPkidPagamentos.Text = ""
                
                If tab_PagamentoParcelas.Tab = 0 Then
                    txt_strTitulo.SetFocus
                ElseIf tab_PagamentoParcelas.Tab = 1 Then
                    txt_intParcelas.SetFocus
                Else
                    dbc_intParcela.SetFocus
                End If
                
                'Caso nao seja o parcelamento principal e estiver incluindo uma parcela,
                'vamos fazer com que cadastre um cancelamento
                If blnCriticarCancelamento Then
                    ExibeMensagem "É necessário existir(em) parcela(s) cancelada(s) para forma de pagamento que não seja parcelamento principal."
                    tab_PagamentoParcelas.Tab = 2
                    dbc_intParcela.SetFocus
                End If

            End If
            
        Case Is = UCase(gstrNovo)
            If Not blnTabPagamentosParcelas Then
                Limpa_Controles Me, True, False, False, True, False
                limpaGrds
                dbc_intComposicao.SetFocus
                tab_PagamentoParcelas.Tab = 0
                
                blnClickGridParametros = False
                blnClickGridPagamentos = False
                blnClickGridParcelas = False
                blnAlterandoParcela = False
                blnAlterandoFormaPagto = False
                strEmissaoAtual = ""
                
                blnNovaParcela = False
                chk_bytParcelado.Value = 1
            Else
                If tab_PagamentoParcelas.Tab = 0 Then
                    LimpaTabParcela
                    LimpaTabPagamento
                    LimpaTabCancelamento
                    
                    blnClickGridParametros = True
                    blnClickGridPagamentos = False
                    blnClickGridParcelas = False
                    blnAlterandoParcela = False
                    blnAlterandoFormaPagto = False
                    blnNovaParcela = True
                                        
                    tab_PagamentoParcelas.Tab = 0
                    txt_strTitulo.SetFocus
                    chk_bytParcelado.Value = 1
                    
                ElseIf tab_PagamentoParcelas.Tab = 1 Then
                    LimpaTabParcela
                    
                    blnClickGridParametros = True
                    blnClickGridPagamentos = True
                    blnClickGridParcelas = False
                    blnAlterandoFormaPagto = True
                    blnAlterandoParcela = False
                    
                    blnNovaParcela = True
                    tab_PagamentoParcelas.Tab = 1
                    txt_intParcelas.SetFocus
                Else
                    tab_PagamentoParcelas.Tab = 2
                    dbc_intParcela.SetFocus
                    dbc_intParcelaBaixa.Text = ""
                    dbc_intParcela.Text = ""
                    dbc_intCodigoBaixa.Text = ""
                    txt_dtmDtVencimento3.Text = ""
                    Set tdb_ParcelaVencimento2.DataSource = Nothing
                End If
            End If
        Case Is = UCase(gstrDeletar)
        
            If blnTabPagamentosParcelas And tab_PagamentoParcelas.Tab = 0 Then
            
                Set tdbGrdSelecionada = tdb_Pagamentos
                
            ElseIf blnTabPagamentosParcelas And tab_PagamentoParcelas.Tab = 1 Then
            
                Set tdbGrdSelecionada = tdb_ParcelaVencimento
                
            ElseIf blnTabPagamentosParcelas And tab_PagamentoParcelas.Tab = 2 Then
            
                Set tdbGrdSelecionada = tdb_ParcelaCancelamento
                    
            Else
            
                Set tdbGrdSelecionada = tdb_Parametros
                
            End If
            
            vetTag = Split(Trim(tdbGrdSelecionada.Tag), ";")
            
            If Not tdbGrdSelecionada.EOF Then
                If gblnExclusaoGravacaoOk("E", "Confirma Exclusão de " & vetTag(1) & "?", True) Then
                    If vetTag(0) = gstrParametroIPTUPagto Then
                        DeletaPagamento tdbGrdSelecionada.Columns("Pkid").Value 'Deleta 1º guia de Forma de Pagamento
                    ElseIf vetTag(0) = gstrFormaPagtoVencimentos Then
                        DeletaParcelasVencimentos tdbGrdSelecionada.Columns("Pkid").Value, 0, 0 'Deleta 2º Guia das parcelas
                    ElseIf vetTag(0) = gstrFormaPagtoCancelamentos Then
                        DeletaParcelaCancelamento tdbGrdSelecionada.Columns("Pkid").Value 'Deleta 3º guia de Cancelamento
                    Else
                        DeletaParametro tdbGrdSelecionada.Columns("Pkid").Value 'Deleta um Lancamento
                    End If
                    
                    If Not blnTabPagamentosParcelas Then
                        Limpa_Controles Me, True, False, False, True, False
                        limpaGrds
                        Set tdb_Parametros.DataSource = Nothing
                        LeDaTabelaParaObj "", tdb_Parametros, strQueryParametros(True)
                    End If
                    
                End If
            Else
                ExibeMensagem "Não há registros a serem excluidos."
            End If
            
        Case Is = UCase(gstrRefresh)
            LeDaTabelaParaObj "", tdb_Parametros, strQueryParametros(True)
    End Select
                 
End Sub

Private Function strQueryRelatorio() As String
'RESPONSAVEL LEANDRO 30/06/2004
Dim strSql As String
    
strSql = "SELECT "
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CE.intCodigo") & strCONCAT & "' - '" & strCONCAT & " CE.strDescricao Composicao,"
    strSql = strSql & " PI.intExercicio Exercicio,"
    strSql = strSql & " PI.strEmissao Emissao,"
    strSql = strSql & " PI.intNumAvisoInicial NumAvisoInicial,"
    strSql = strSql & " PI.intNumAvisoFinal NumAvisoFinal"
strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita & " CE, "
    strSql = strSql & gstrParametroIPTU & " PI"
    
strSql = strSql & " WHERE"
    strSql = strSql & " PI.intComposicaoDaReceita " & strOUTJSQLServer & "=" & " CE.Pkid " & strOUTJOracle
    If dbc_intComposicao.MatchedWithList Then strSql = strSql & " AND PI.intComposicaoDaReceita " & IIf(Val(dbc_intComposicao.BoundText) <> 0, "= '" & dbc_intComposicao.BoundText & "'", "LIKE '" & dbc_intComposicao.BoundText & "%'")
    If Len(Trim(txt_intExercicio.Text)) > 0 Then strSql = strSql & " AND PI.intExercicio LIKE " & "'" & txt_intExercicio.Text & "%'"
    If Len(Trim(txt_strEmissao.Text)) > 0 Then strSql = strSql & " AND UPPER(PI.strEmissao) LIKE " & "'" & String(gintLenEmissao - Len(Trim(txt_strEmissao.Text)), "0") & UCase(txt_strEmissao.Text) & "%'"
    If Len(Trim(txt_intNumAvisoInicial.Text)) > 0 Then strSql = strSql & " AND PI.intNumAvisoInicial = " & Val(txt_intNumAvisoInicial.Text)
    If Len(Trim(txt_intNumAvisoFinal.Text)) > 0 Then strSql = strSql & " AND (PI.intNumAvisoFinal = " & Val(txt_intNumAvisoFinal.Text) & " OR PI.intNumAvisoFinal is null)"
    
strSql = strSql & " ORDER BY"
    strSql = strSql & " CE.intCodigo,"
    strSql = strSql & " CE.strDescricao"
strQueryRelatorio = strSql
' '
 '
End Function

Private Function strQueryParametros(blnRefresh As Boolean)
    Dim strSql As String
    
    strSql = "SELECT " & gstrCONVERT(CDT_VARCHAR, "CE.intCodigo") & strCONCAT & "' - '" & strCONCAT & " CE.strDescricao Composicao,"
    strSql = strSql & " PI.intExercicio Exercicio,"
    strSql = strSql & " PI.strEmissao Emissao,"
    strSql = strSql & " PI.intNumAvisoInicial NumAvisoInicial,"
    strSql = strSql & " PI.intNumAvisoFinal NumAvisoFinal,"
    strSql = strSql & " PI.Pkid"
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita & " CE, "
    strSql = strSql & gstrParametroIPTU & " PI"
    strSql = strSql & " WHERE"
    strSql = strSql & " PI.intComposicaoDaReceita " & strOUTJSQLServer & "=" & " CE.Pkid " & strOUTJOracle

    If Not blnRefresh Then
        If dbc_intComposicao.MatchedWithList Then strSql = strSql & " AND PI.intComposicaoDaReceita " & IIf(Val(dbc_intComposicao.BoundText) <> 0, "= '" & dbc_intComposicao.BoundText & "'", "LIKE '" & dbc_intComposicao.BoundText & "%'")
        If Len(Trim(txt_intExercicio.Text)) > 0 Then strSql = strSql & " AND PI.intExercicio LIKE " & "'" & txt_intExercicio.Text & "%'"
        If Len(Trim(txt_strEmissao.Text)) > 0 Then strSql = strSql & " AND UPPER(PI.strEmissao) LIKE " & "'" & String(gintLenEmissao - Len(Trim(txt_strEmissao.Text)), "0") & UCase(txt_strEmissao.Text) & "%'"
        If Len(Trim(txt_intNumAvisoInicial.Text)) > 0 Then strSql = strSql & " AND PI.intNumAvisoInicial = " & Val(txt_intNumAvisoInicial.Text)
        If Len(Trim(txt_intNumAvisoFinal.Text)) > 0 Then strSql = strSql & " AND (PI.intNumAvisoFinal = " & Val(txt_intNumAvisoFinal.Text) & " OR PI.intNumAvisoFinal is null)"
    End If

    strSql = strSql & " ORDER BY intCodigo"
    
    strQueryParametros = strSql

End Function

Private Sub PreencheGrdPagamentos(lngPkidParametro As Long)
Dim strSql As String
    
    strSql = "SELECT PP.strTitulo Titulo, "
    strSql = strSql & " PP.dblDesconto Desconto, "
    strSql = strSql & " PP.bytParcelado, "
    strSql = strSql & " IE.Strabreviatura CodMoeda , "
    strSql = strSql & " PP.dblValorMoeda ValorMoeda, "
    strSql = strSql & " PP.intDesctoTaxas intDesctoTaxas, "
    strSql = strSql & " PP.Pkid"
    strSql = strSql & " FROM "
    strSql = strSql & gstrParametroIPTUPagto & " PP, "
    strSql = strSql & gstrIndexadorEconomico & " IE"
    strSql = strSql & " WHERE PP.intParametroIPTU ='" & lngPkidParametro & "'"
    strSql = strSql & " AND PP.intIndexadorEconomico " & strOUTJSQLServer & "=" & " IE.Pkid " & strOUTJOracle
    strSql = strSql & " ORDER BY PP.strTitulo"
    
    LeDaTabelaParaObj "", tdb_Pagamentos, strSql

End Sub

Private Sub PreencheGrdParcelas(lngPkidPagamento As Long)
    Dim strSql As String
    
    strSql = "SELECT FV.Pkid, FV.intParcela intParcelas,"
    strSql = strSql & " FV.dtmDtVencimento"
    strSql = strSql & " FROM "
    strSql = strSql & gstrFormaPagtoVencimentos & " FV"
    strSql = strSql & " WHERE FV.intFormaPagto ='" & lngPkidPagamento & "'"
    strSql = strSql & " ORDER BY FV.dtmDtVencimento "
    
    LeDaTabelaParaObj "", tdb_ParcelaVencimento, strSql
    LeDaTabelaParaObj "", tdb_ParcelaVencimento2, strSql

End Sub

Private Function blnDadosOk() As Boolean

    blnDadosOk = False
    blnDadosCancelamento = False
    
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "Selecione uma Composição da Receita válido."
        dbc_intComposicao.SetFocus
        Exit Function
    End If
    
    If txt_intExercicio.Text = "" Then
        ExibeMensagem "É necessário informar o Exercício."
        txt_intExercicio.SetFocus
        Exit Function
    End If
        
    If txt_strEmissao.Text = "" Then
        ExibeMensagem "É necessário informar a Emissão."
        txt_strEmissao.SetFocus
        Exit Function
    End If
        
    If txt_intNumAvisoInicial.Text = "" Then
        ExibeMensagem "É necessário informar o Número de Aviso Inicial."
        txt_intNumAvisoInicial.SetFocus
        Exit Function
    End If
        
    If txt_intNumAvisoFinal.Text <> "" Then
        If Val(txt_intNumAvisoInicial.Text) > Val(txt_intNumAvisoFinal.Text) Then
            ExibeMensagem "O número de aviso inicial não pode ser menor quer o número de aviso final."
            txt_intNumAvisoFinal.SetFocus
            Exit Function
        End If
    End If
        
    If txt_strTitulo <> "" Then
    
        If Val(txtPkidPagamentos.Text) <> 0 And (Val(strEmissaoAtual) <> Val(txt_strEmissao.Text)) Then
            If Not blnParcelaOK Then
                ExibeMensagem "Já existe uma Forma de Pagamento com a Composição da Receita selecionada."
                txt_strTitulo.SetFocus
                Exit Function
            End If
        End If

        If (blnAlterandoFormaPagto And blnAlterandoParcela) _
            Or (Not blnAlterandoFormaPagto And Not blnAlterandoParcela) _
            Or (blnAlterandoFormaPagto And blnNovaParcela And tab_PagamentoParcelas.Tab = 1) Then
                        
            If txt_intParcelas.Text = "" Then
                ExibeMensagem "É necessário informar um número de Parcela."
                txt_intParcelas.SetFocus
                Exit Function
            End If
            
            If Val(txtPkidPagamentos.Text) = 0 Or (Val(txtPkidPagamentos.Text) <> 0 And strParcelaAtual <> RTrim(LTrim(txt_intParcelas.Text))) Then
                If gblnExisteCodigo(2, gstrFormaPagtoVencimentos, "intParcela", "'" & Val(txt_intParcelas.Text) & "'", "intFormaPagto", "'" & Val(txtPkidPagamentos.Text) & "'") Then
                    ExibeMensagem "Já existe o número de Parcela informado."
                    txt_intParcelas.SetFocus
                    Exit Function
                End If
            End If
            
            If txt_dtmDtVencimento.Text = "" Then
                ExibeMensagem "É necessário informar a data de Vencimento."
                txt_dtmDtVencimento.SetFocus
                Exit Function
            Else
                If Not gblnDataValida(txt_dtmDtVencimento.Text) Then
                    ExibeMensagem "A Data de Vencimento não é válida."
                    txt_dtmDtVencimento.SetFocus
                    Exit Function
                End If
            End If
            If Not blnVencimentoParcela(blnAlterandoParcela) Then
                ExibeMensagem "Já existe uma Parcela a vencer no dia " & txt_dtmDtVencimento.Text & "."
                MarcaCampo txt_dtmDtVencimento
                Exit Function
            End If
        End If
    Else
        ExibeMensagem "É necessário informar um Título."
        txt_strTitulo.SetFocus
        Exit Function
    End If
    
    If (dbc_intParcela.MatchedWithList = True) And (dbc_intParcelaBaixa.MatchedWithList = True) And (dbc_intCodigoBaixa.MatchedWithList) Then
        blnDadosCancelamento = True
    End If
    
    If txtPkidParametros <> "" And chk_bytParcelado = 1 Then
        If blnParcelaPrincipal Then
            ExibeMensagem "Já existe uma forma de pagamento principal"
            Exit Function
        End If
    End If
    
    If tab_PagamentoParcelas.Tab = 2 Then
    
        If Not dbc_intParcela.MatchedWithList Then
            ExibeMensagem "Selecione uma parcela antes de salvar."
            dbc_intParcela.SetFocus
            Exit Function
        End If
        
        If Not dbc_intParcelaBaixa.MatchedWithList Then
            ExibeMensagem "Selecione uma parcela a ser cancelada"
            dbc_intParcelaBaixa.SetFocus
            Exit Function
        End If
        
        If Not dbc_intCodigoBaixa.MatchedWithList Then
            ExibeMensagem "Selecione um tipo de baixa"
            dbc_intCodigoBaixa.SetFocus
            Exit Function
        End If
    
    End If
    
    If chk_bytParcelado = 0 Then
        If blnAlterandoParcela Then
            If Not blnExisteCancelamento And Not blnDadosCancelamento Then
                ExibeMensagem "É necessário existir(em) parcela(s) cancelada(s) para forma de pagamento que não seja parcelamento principal."
                tab_PagamentoParcelas.Tab = 2
                dbc_intParcela.SetFocus
                Exit Function
            End If
         End If
    End If
    
    blnDadosOk = True
 
End Function

Private Function strQueryCodMoeda() As String

    Dim strSql As String
    
    strSql = "SELECT Pkid, strAbreviatura"
    strSql = strSql & " FROM "
    strSql = strSql & gstrIndexadorEconomico
    strSql = strSql & " ORDER BY strAbreviatura"
    
    strQueryCodMoeda = strSql

End Function

Private Sub txt_strTitulo_GotFocus()
    MarcaCampo txt_strTitulo
    blnTabPagamentosParcelas = True
    blnParametros = False
    tab_PagamentoParcelas.Tab = 0
End Sub

Private Sub txt_strTitulo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strTitulo
End Sub

Private Function GravaParametros(blnAlterando As Boolean) As Boolean
    Dim strSql As String
    
    GravaParametros = False
    
    If Not blnAlterando Then
    
        strSql = "INSERT INTO " & gstrParametroIPTU & "(intComposicaoDaReceita,"
        strSql = strSql & " intExercicio,"
        strSql = strSql & " strEmissao,"
        strSql = strSql & " intNumAvisoInicial,"
        strSql = strSql & " intNumAvisoFinal,"
        strSql = strSql & " dtmDtAtualizacao,"
        strSql = strSql & " lngCodUsr)"
        strSql = strSql & " VALUES ("
        strSql = strSql & "'" & dbc_intComposicao.BoundText & "',"
        strSql = strSql & " '" & Val(txt_intExercicio.Text) & "',"
        strSql = strSql & " '" & txt_strEmissao.Text & "',"
        strSql = strSql & " " & txt_intNumAvisoInicial.Text & ","
        strSql = strSql & " " & gstrENulo(txt_intNumAvisoFinal.Text, , True) & ", "
        strSql = strSql & strGETDATE & ", "
        strSql = strSql & glngCodUsr
        strSql = strSql & ")"
    Else
        strSql = "UPDATE " & gstrParametroIPTU
        strSql = strSql & " SET "
        strSql = strSql & " intComposicaoDaReceita ='" & dbc_intComposicao.BoundText & "',"
        strSql = strSql & " intExercicio ='" & Val(txt_intExercicio.Text) & "',"
        strSql = strSql & " strEmissao ='" & txt_strEmissao.Text & "',"
        strSql = strSql & " intNumAvisoInicial =" & txt_intNumAvisoInicial.Text & ","
        strSql = strSql & " intNumAvisoFinal =" & gstrENulo(txt_intNumAvisoFinal.Text, , True) & ","
        strSql = strSql & " dtmDtAtualizacao =" & strGETDATE & ","
        strSql = strSql & " lngCodUsr =" & glngCodUsr
        strSql = strSql & " WHERE"
        strSql = strSql & " Pkid = " & Val(txtPkidParametros.Text)
    End If
    
    
    If Not gobjBanco.Execute(strSql) Then
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    GravaParametros = True
    
End Function

Private Sub GravaPagamentos(blnAlterando As Boolean, lngPkidParametro As Long)
    Dim strSql As String
    
    If Not blnAlterando Then
        strSql = "INSERT INTO " & gstrParametroIPTUPagto & "(intParametroIPTU,"
        strSql = strSql & " strTitulo,"
        strSql = strSql & " dblDesconto,"
        strSql = strSql & " intIndexadorEconomico,"
        strSql = strSql & " dblValorMoeda,"
        strSql = strSql & " bytParcelado,"
        strSql = strSql & " intDesctoTaxas, "
        strSql = strSql & " dtmDtAtualizacao,"
        strSql = strSql & " lngCodUsr)"
        strSql = strSql & " VALUES("
        strSql = strSql & Val(lngPkidParametro) & ", "
        strSql = strSql & "'" & RTrim(LTrim(txt_strTitulo.Text)) & "', "
        strSql = strSql & gstrConvVrParaSql(txt_dblDesconto.Text) & ", "
        strSql = strSql & "'" & dbc_strCodMoeda.BoundText & "', "
        strSql = strSql & gstrConvVrParaSql(IIf(Trim(txt_dblValorMoeda.Text) = "", 0, txt_dblValorMoeda.Text)) & ", "
        strSql = strSql & IIf(chk_bytParcelado.Value = 1, 1, 0) & ", "
        strSql = strSql & gstrConvVrParaSql(txt_intDesctoTaxas) & ", "
        strSql = strSql & strGETDATE & ", "
        strSql = strSql & glngCodUsr
        strSql = strSql & ")"
    Else
        strSql = "UPDATE " & gstrParametroIPTUPagto
        strSql = strSql & " SET "
        strSql = strSql & " intParametroIPTU = " & lngPkidParametro & ","
        strSql = strSql & " strTitulo = '" & RTrim(LTrim(txt_strTitulo.Text)) & "',"
        strSql = strSql & " dblDesconto = " & gstrConvVrParaSql(txt_dblDesconto.Text) & ","
        strSql = strSql & " intIndexadorEconomico = '" & dbc_strCodMoeda.BoundText & "',"
        strSql = strSql & " dblValorMoeda = " & gstrConvVrParaSql(IIf(Trim(txt_dblValorMoeda.Text) = "", 0, txt_dblValorMoeda.Text)) & ", "
        strSql = strSql & " bytParcelado = " & IIf(chk_bytParcelado.Value = 1, 1, 0) & ", "
        strSql = strSql & " intDesctoTaxas = " & gstrConvVrParaSql(txt_intDesctoTaxas) & ", "
        strSql = strSql & " dtmDtAtualizacao =" & strGETDATE & ", "
        strSql = strSql & " lngCodUsr =" & glngCodUsr
        strSql = strSql & " WHERE"
        strSql = strSql & " Pkid = " & Val(txtPkidPagamentos.Text)
    End If
    
    Set gobjbjanco = New clsBanco
    If Not gobjBanco.Execute(strSql) Then
        gobjBanco.ExecutaRollbackTrans
    End If

End Sub

Private Function PegaUltimoPkidParametro() As String
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT MAX(Pkid) Pkid "
    strSql = strSql & " FROM "
    strSql = strSql & gstrParametroIPTU
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            PegaUltimoPkidParametro = adoResultado!Pkid
        Else
            PegaUltimoPkidParametro = ""
        End If
    End If
    
End Function

Private Sub GravaParcelas(blnAlterando As Boolean, lngPkidPagamento As Long)
    Dim strSql As String
    
    If Not blnAlterando Then
        strSql = "INSERT INTO " & gstrFormaPagtoVencimentos & "(intFormaPagto,"
        strSql = strSql & " intParcela,"
        strSql = strSql & " dtmDtVencimento,"
        strSql = strSql & " dtmDtAtualizacao,"
        strSql = strSql & " lngCodUsr)"
        strSql = strSql & " VALUES("
        strSql = strSql & Val(lngPkidPagamento) & ", "
        strSql = strSql & txt_intParcelas.Text & ", "
        strSql = strSql & gstrConvDtParaSql(txt_dtmDtVencimento.Text) & ", "
        strSql = strSql & strGETDATE & ", "
        strSql = strSql & glngCodUsr
        strSql = strSql & ")"
    Else
        strSql = "UPDATE " & gstrFormaPagtoVencimentos
        strSql = strSql & " SET "
        strSql = strSql & " intFormaPagto = " & lngPkidPagamento & ","
        strSql = strSql & " intparcela = " & txt_intParcelas.Text & ","
        strSql = strSql & " dtmDtVencimento = " & gstrConvDtParaSql(txt_dtmDtVencimento.Text) & ", "
        strSql = strSql & " dtmDtAtualizacao = " & strGETDATE & ", "
        strSql = strSql & " lngCodUsr = " & glngCodUsr
        strSql = strSql & " WHERE"
        strSql = strSql & " Pkid = " & tdb_ParcelaVencimento.Columns("Pkid").Value
    End If
    
    Set gobjBanco = New clsBanco
    
    If Not gobjBanco.Execute(strSql) Then
        gobjBanco.ExecutaRollbackTrans
    End If

End Sub

Private Function PegaUltimoPkidPagamentos() As String
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT MAX(Pkid) Pkid"
    strSql = strSql & " FROM "
    strSql = strSql & gstrParametroIPTUPagto
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            PegaUltimoPkidPagamentos = adoResultado!Pkid
        Else
            PegaUltimoPkidPagamentos = ""
        End If
    End If

End Function
Private Sub limpaGrds()
    Set tdb_Pagamentos.DataSource = Nothing
    Set tdb_ParcelaVencimento.DataSource = Nothing
    Set tdb_ParcelaVencimento2.DataSource = Nothing
End Sub

Private Sub txt_strTitulo_LostFocus()
    blnTabPagamentosParcelas = False
End Sub

Private Sub LimpaTabPagamento()

    txtPkidPagamentos.Text = ""
    txt_strTitulo.Text = ""
    txt_strTitulo2.Text = ""
    txt_strTitulo3.Text = ""
    txt_dblDesconto.Text = ""
    dbc_strCodMoeda.Text = ""
    dbc_strCodMoeda.ListField = ""
    txt_dblValorMoeda.Text = ""
    txt_intDesctoTaxas.Text = ""
    Set tdb_ParcelaVencimento.DataSource = Nothing
    Set tdb_ParcelaVencimento2.DataSource = Nothing
    tab_PagamentoParcelas.Tab = 0
    txt_strTitulo.SetFocus
    
End Sub

Private Sub LimpaTabParcela()
    txtPkidParcelas.Text = ""
    txt_intParcelas.Text = ""
    txt_dtmDtVencimento.Text = ""
End Sub

Private Sub PreencheCamposParametros(lngPkidParametro As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "SELECT " & gstrCONVERT(CDT_VARCHAR, "CE.intCodigo") & strCONCAT & "' - '" & strCONCAT & " CE.strDescricao Composicao,"
    strSql = strSql & " PI.intExercicio Exercicio,"
    strSql = strSql & " PI.strEmissao Emissao,"
    strSql = strSql & " PI.intNumAvisoInicial NumAvisoInicial,"
    strSql = strSql & " PI.intNumAvisoFinal NumAvisoFinal,"
    strSql = strSql & " PI.Pkid,"
    strSql = strSql & " CE.Pkid PkidCE"
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita & " CE, "
    strSql = strSql & gstrParametroIPTU & " PI"
    strSql = strSql & " WHERE"
    strSql = strSql & " PI.intComposicaoDaReceita " & strOUTJSQLServer & "=" & " CE.Pkid " & strOUTJOracle
    strSql = strSql & " AND PI.Pkid =" & lngPkidParametro
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            PreencherListaDeOpcoes dbc_intComposicao, Val(adoResultado!PkidCE)
            txt_intExercicio.Text = gstrENulo(adoResultado!EXERCICIO)
            txt_strEmissao.Text = gstrENulo(adoResultado!Emissao)
            txt_intNumAvisoInicial.Text = gstrENulo(adoResultado!NumAvisoInicial)
            txt_intNumAvisoFinal.Text = gstrENulo(adoResultado!NumAvisoFinal)
        Else
            dbc_intComposicao.Text = ""
            txt_intExercicio.Text = ""
            txt_strEmissao.Text = ""
            txt_intNumAvisoInicial.Text = ""
            txt_intNumAvisoFinal.Text = ""
        End If
    End If

End Sub

Private Sub PreencheCamposPagamentos(lngPkidPagamentos As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "SELECT PP.strTitulo Titulo,"
    strSql = strSql & " PP.dblDesconto Desconto,"
    strSql = strSql & " IE.strAbreviatura CodMoeda ,"
    strSql = strSql & " PP.dblValorMoeda ValorMoeda,"
    strSql = strSql & " PP.intDesctoTaxas intDesctoTaxas, "
    strSql = strSql & " PP.Pkid,"
    strSql = strSql & " IE.Pkid PkidMoeda"
    strSql = strSql & " FROM "
    strSql = strSql & gstrParametroIPTUPagto & " PP, "
    strSql = strSql & gstrIndexadorEconomico & " IE"
    strSql = strSql & " WHERE PP.intIndexadorEconomico " & strOUTJSQLServer & "=" & " IE.Pkid " & strOUTJOracle
    strSql = strSql & " AND PP.Pkid = " & lngPkidPagamentos
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_strTitulo.Text = gstrENulo(adoResultado!Titulo)
            txt_strTitulo2.Text = gstrENulo(adoResultado!Titulo)
            txt_strTitulo3.Text = gstrENulo(adoResultado!Titulo)
            txt_dblDesconto.Text = gstrENulo(gstrConvVrDoSql(adoResultado!Desconto, 2))
            PreencherListaDeOpcoes dbc_strCodMoeda, Val(gstrENulo(adoResultado!PkidMoeda))
            txt_dblValorMoeda.Text = gstrENulo(gstrConvVrDoSql(adoResultado!ValorMoeda, 4))
            txt_intDesctoTaxas.Text = IIf(IsNull(adoResultado!intDesctoTaxas), "", gstrConvVrDoSql(adoResultado!intDesctoTaxas, 2, 3, False))
        Else
            txt_strTitulo.Text = ""
            txt_strTitulo2.Text = ""
            txt_strTitulo3.Text = ""
            txt_dblDesconto.Text = ""
            dbc_strCodMoeda.Text = ""
            txt_dblValorMoeda.Text = ""
            txt_intDesctoTaxas.Text = ""
        End If
    End If

End Sub

Private Sub DeletaParcelasVencimentos(Optional lngPkidParcelaVencimeto As Long, Optional lngPkidFormaPagamento As Long, Optional lngPkidParametroIPTU)
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    If Val(lngPkidFormaPagamento) <> 0 Then
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
        strSql = strSql & "Delete "
        strSql = strSql & "From "
        strSql = strSql & gstrFormaPagtoCancelamentos
        strSql = strSql & " Where "
        strSql = strSql & "Intformapagtovencimentos in"
        strSql = strSql & "(select PKID from tblFormaPagtoVencimentos WHERE  intFormaPagto = " & Val(lngPkidFormaPagamento) & IIf(bytDBType = Oracle, ");", ")")
        strSql = strSql & "DELETE FROM " & gstrFormaPagtoVencimentos
        strSql = strSql & " WHERE "
        strSql = strSql & " intFormaPagto = " & Val(lngPkidFormaPagamento) & IIf(bytDBType = Oracle, ";", "")
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    ElseIf Val(lngPkidParametroIPTU) <> 0 Then
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
        strSql = strSql & "Delete from " & gstrFormaPagtoCancelamentos & " Where intformapagtovencimentos "
        strSql = strSql & "in(SELECT Pkid FROM " & gstrFormaPagtoVencimentos & " WHERE intFormaPagto IN "
        strSql = strSql & "(SELECT PKID FROM " & gstrParametroIPTUPagto & " WHERE intParametroIPTU = " & Val(lngPkidParametroIPTU) & "))" & IIf(bytDBType = Oracle, ";", "")
        strSql = strSql & "DELETE FROM " & gstrFormaPagtoVencimentos
        strSql = strSql & " WHERE intFormaPagto IN "
        strSql = strSql & "(SELECT PKID FROM " & gstrParametroIPTUPagto
        strSql = strSql & " WHERE intParametroIPTU = " & Val(lngPkidParametroIPTU) & ")" & IIf(bytDBType = Oracle, ";", "")
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    Else
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
        strSql = strSql & "DELETE FROM " & gstrFormaPagtoCancelamentos
        strSql = strSql & " WHERE "
        strSql = strSql & " intformapagtovencimentos = " & Val(lngPkidParcelaVencimeto) & IIf(bytDBType = Oracle, ";", "")
        strSql = strSql & "DELETE FROM " & gstrFormaPagtoVencimentos
        strSql = strSql & " WHERE "
        strSql = strSql & " PKID = " & Val(lngPkidParcelaVencimeto) & IIf(bytDBType = Oracle, ";", "")
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    End If
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute strSql
    
    PreencheGrdParcelas Val(txtPkidPagamentos.Text)
    LimpaTabParcela
    LimpaTabCancelamento
    blnClickGridParcelas = False

End Sub

Private Sub DeletaPagamento(Optional lngPkidPagamento As Long, Optional lngPkidParametro As Long)
    Dim strSql As String
    
    DeletaParcelasVencimentos 0, lngPkidPagamento, 0
    
    If Val(lngPkidParametro) <> 0 Then
        strSql = "DELETE FROM " & gstrParametroIPTUPagto
        strSql = strSql & " WHERE"
        strSql = strSql & " intParametroIPTU =" & Val(lngPkidParametro)
    Else
        strSql = "DELETE FROM " & gstrParametroIPTUPagto
        strSql = strSql & " WHERE"
        strSql = strSql & " Pkid =" & Val(lngPkidPagamento)
    End If
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute strSql
    
    PreencheGrdPagamentos Val(txtPkidParametros.Text)
    
    LimpaTabPagamento
    
    blnClickGridPagamentos = False

End Sub

Private Sub DeletaParametro(Optional lngPkidParametro As Long)
    Dim strSql As String
    
    DeletaParcelasVencimentos , , Val(lngPkidParametro)
    DeletaPagamento , Val(lngPkidParametro)
    
    strSql = "DELETE FROM " & gstrParametroIPTU
    strSql = strSql & " WHERE"
    strSql = strSql & " Pkid = " & Val(lngPkidParametro)
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute strSql
    
    blnClickGridParametros = False
    
    Set tdb_Parametros.DataSource = Nothing
    MantemForm gstrRefresh

End Sub

Private Function strQueryParcela() As String
    Dim strSql As String
    strSql = "SELECT FV.Pkid, FV.intParcela intParcelas "
    strSql = strSql & " FROM "
    strSql = strSql & gstrFormaPagtoVencimentos & " FV"
    strSql = strSql & " WHERE FV.intFormaPagto ='" & txtPkidPagamentos.Text & "'"
    strSql = strSql & " ORDER BY FV.dtmDtVencimento "
    strQueryParcela = strSql
End Function

Private Function strQueryParcelaCancelar() As String
    Dim strSql As String
    
    strSql = strSql & "SELECT "
    strSql = strSql & "FV.Pkid, "
    strSql = strSql & "FV.intParcela "
    strSql = strSql & "From "
    strSql = strSql & gstrParametroIPTU & " PI, "
    strSql = strSql & gstrParametroIPTUPagto & " PP, "
    strSql = strSql & gstrFormaPagtoVencimentos & " FV "
    strSql = strSql & "Where "
    strSql = strSql & "PP.Bytparcelado = 1 AND "
    strSql = strSql & "PI.Pkid = PP.Intparametroiptu AND "
    strSql = strSql & "PP.Pkid = FV.Intformapagto AND "
    strSql = strSql & "PI.Pkid =" & tdb_Parametros.Columns("Pkid").Value & " AND "
    strSql = strSql & "not FV.intParcela = " & dbc_intParcela.Text
        strSql = strSql & " AND FV.Pkid not in(SELECT "
        strSql = strSql & "CL.INTFORMAPAGTOVENCIMENTOSCANCEL "
        strSql = strSql & "From "
        strSql = strSql & gstrFormaPagtoCancelamentos & " CL "
        strSql = strSql & "Where "
        strSql = strSql & "Cl.Intformapagtovencimentos = " & dbc_intParcela.BoundText & " ) "
    strSql = strSql & " Order By "
    strSql = strSql & "FV.intParcela"

    strQueryParcelaCancelar = strSql
    
End Function

Private Sub GravaParcelasCanceladas()
    Dim strSql As String
    
    strSql = "INSERT INTO " & gstrFormaPagtoCancelamentos & "(intformapagtovencimentos,"
    strSql = strSql & " intformapagtovencimentoscancel,"
    strSql = strSql & " intCodigoBaixa,"
    strSql = strSql & " dtmDtAtualizacao,"
    strSql = strSql & " lngCodUsr)"
    strSql = strSql & " VALUES("
    strSql = strSql & dbc_intParcela.BoundText & ", "
    strSql = strSql & dbc_intParcelaBaixa.BoundText & ", "
    strSql = strSql & dbc_intCodigoBaixa.BoundText & ", "
    strSql = strSql & strGETDATE & ", "
    strSql = strSql & glngCodUsr
    strSql = strSql & ")"
    
    Set gobjBanco = New clsBanco
    
    If Not gobjBanco.Execute(strSql) Then
        gobjBanco.ExecutaRollbackTrans
    End If
End Sub

Private Function strQueryParcelaCancelarGrid() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "ParcelaCancelada.Pkid, "
    strSql = strSql & "ParcelaCancelada.intParcela , "
    strSql = strSql & "ParcelaCancelada.Dtmdtvencimento "
    strSql = strSql & "From "
    strSql = strSql & gstrFormaPagtoVencimentos & " FV, "
    
      strSql = strSql & "(SELECT "
             strSql = strSql & "CL.Pkid, "
             strSql = strSql & "CL.INTFORMAPAGTOVENCIMENTOS, "
             strSql = strSql & "FV.intParcela as intParcela, "
             strSql = strSql & "FV.Dtmdtvencimento "
      strSql = strSql & "From "
           strSql = strSql & gstrFormaPagtoVencimentos & " FV, "
           strSql = strSql & gstrFormaPagtoCancelamentos & " CL "
      strSql = strSql & "Where "
            strSql = strSql & "FV.Pkid = CL.Intformapagtovencimentoscancel "
      strSql = strSql & ") ParcelaCancelada "
    strSql = strSql & "Where "
    strSql = strSql & "FV.Pkid = ParcelaCancelada.INTFORMAPAGTOVENCIMENTOS AND "
    strSql = strSql & "FV.Pkid = " & dbc_intParcela.BoundText
    strSql = strSql & "Order By "
    strSql = strSql & "ParcelaCancelada.intParcela "

    strQueryParcelaCancelarGrid = strSql
    
End Function

Private Sub LimpaTabCancelamento()
    txt_strTitulo3.Text = ""
    Set dbc_intParcela.RowSource = Nothing
    Set dbc_intParcelaBaixa.RowSource = Nothing
    dbc_intParcela.Text = ""
    dbc_intParcelaBaixa.Text = ""
    txt_dtmDtVencimento3.Text = ""
    Set tdb_ParcelaCancelamento.DataSource = Nothing
End Sub

Private Function DeletaParcelaCancelamento(Optional lngPkidCancelamento As Long)

    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "DELETE FROM " & gstrFormaPagtoCancelamentos & " "
    strSql = strSql & "WHERE "
    strSql = strSql & "Pkid = " & Val(lngPkidCancelamento)
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute strSql
    
    LeDaTabelaParaObj "", dbc_intParcelaBaixa, strQueryParcelaCancelar
    LeDaTabelaParaObj "", tdb_ParcelaCancelamento, strQueryParcelaCancelarGrid
    
End Function

Private Function blnExisteCancelamento() As Boolean
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    
    blnExisteCancelamento = False
    
    strSql = "SELECT FV.Pkid, FV.intParcela intParcelas "
    strSql = strSql & " FROM "
    strSql = strSql & gstrFormaPagtoVencimentos & " FV, "
    strSql = strSql & gstrFormaPagtoCancelamentos & " C "
    strSql = strSql & " WHERE FV.Pkid =" & tdb_ParcelaVencimento.Columns("Pkid").Value
    strSql = strSql & " AND FV.Pkid = C.intFormaPagtoVencimentos "
        
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .RecordCount > 0 Then
                blnExisteCancelamento = True
            End If
        End With
    End If
    
    Set gobjBanco = Nothing
    
End Function

Private Function blnParcelaPrincipal() As Boolean
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    blnParcelaPrincipal = False
    
    strSql = ""
    strSql = strSql & "Select bytParcelado "
    strSql = strSql & "From "
    strSql = strSql & gstrParametroIPTU & " P, "
    strSql = strSql & gstrParametroIPTUPagto & " PP "
    strSql = strSql & "Where "
    strSql = strSql & "P.Pkid = PP.Intparametroiptu and "
    strSql = strSql & "PP.Intparametroiptu = " & txtPkidParametros & " And PP.Pkid <> " & Val(txtPkidPagamentos.Text)
        
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                Do While Not .EOF
                If IIf(IsNull(!bytParcelado), 0, !bytParcelado) = 1 Then
                    blnParcelaPrincipal = True
                    Exit Do
                    Set gobjBanco = Nothing
                End If
                .MoveNext
                Loop
            End If
        End With
    End If
    Set gobjBanco = Nothing
    
End Function

Private Function blnParcelaOK() As Boolean

    blnParcelaOK = False
       
    Dim strSql As String
    Dim adoRec As New ADODB.Recordset
    
    strSql = "SELECT   IP.PKID," & _
                        "FV.Pkid, " & _
                        "FV.intParcela intParcelas, " & _
                        "FV.Dtmdtvencimento " & _
              "FROM " & gstrParametroIPTUPagto & " IP, " & _
                        gstrParametroIPTU & " PI, " & _
                        gstrFormaPagtoVencimentos & " FV " & _
              "WHERE    IP.Intparametroiptu = PI.Pkid" & _
                      " AND PI.Intcomposicaodareceita = " & dbc_intComposicao.BoundText & _
                      " AND PI.Intexercicio = " & Trim(txt_intExercicio.Text) & _
                      " AND PI.Stremissao = " & String(gintLenEmissao - Len(Trim(txt_strEmissao.Text)), "0") & UCase(txt_strEmissao.Text) & _
                      " AND FV.intFormaPagto = IP.Pkid"
                      '" AnD IP.strTitulo = '" & Trim(txt_strTitulo.Text) & "'"
'                      IIf(Trim(txt_intParcelas.Text) <> "", " AND FV.Intparcela = " & Trim(txt_intParcelas.Text), "") & _

    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        If adoRec.RecordCount = 0 Then
            blnParcelaOK = True
        End If
    End If

End Function

Private Function blnVencimentoParcela(Optional blnAlterando As Boolean = False) As Boolean

    blnVencimentoParcela = False
       
    Dim strSql As String
    Dim adoRec As New ADODB.Recordset
    
    strSql = "SELECT   IP.PKID," & _
                        "FV.Pkid, " & _
                        "FV.intParcela intParcelas, " & _
                        "FV.Dtmdtvencimento " & _
              "FROM " & gstrParametroIPTUPagto & " IP, " & _
                        gstrParametroIPTU & " PI, " & _
                        gstrFormaPagtoVencimentos & " FV " & _
              "WHERE    IP.Intparametroiptu = PI.Pkid" & _
                      " AND PI.Intcomposicaodareceita = " & dbc_intComposicao.BoundText & _
                      " AND IP.strTitulo = '" & Trim(txt_strTitulo.Text) & "'" & _
                      " AND PI.Intexercicio = " & Trim(txt_intExercicio.Text) & _
                      " AND PI.Stremissao = " & String(gintLenEmissao - Len(Trim(txt_strEmissao.Text)), "0") & UCase(txt_strEmissao.Text) & _
                      " AND FV.intFormaPagto = IP.Pkid" & _
                      " AND FV.dtmdtVencimento = " & gstrConvDtParaSql(txt_dtmDtVencimento.Text)
                      
    If blnAlterando Then
        strSql = strSql & " AND FV.intParcela <> " & tdb_ParcelaVencimento.Columns("intParcelas")
    End If
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        If adoRec.RecordCount = 0 Then
            blnVencimentoParcela = True
        End If
    End If

End Function
