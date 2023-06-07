VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadPrecoPublico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preço Público"
   ClientHeight    =   8115
   ClientLeft      =   330
   ClientTop       =   2025
   ClientWidth     =   10455
   HelpContextID   =   5
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2730
      TabIndex        =   49
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6525
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   11509
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Preço Público"
      TabPicture(0)   =   "frmCadPrecoPublico.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_DomicilioFiscal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_PrescricaoDoDebito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Fra_Titulo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_receitas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Parcelas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "frmCadPrecoPublico.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Historico"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra_Parcelas 
         Caption         =   "Parcelas"
         Height          =   1695
         Left            =   120
         TabIndex        =   44
         Top             =   4770
         Width           =   10095
         Begin TrueOleDBGrid70.TDBGrid tdb_Parcelas 
            Height          =   1395
            Left            =   60
            TabIndex        =   45
            Top             =   210
            Width           =   9945
            _ExtentX        =   17542
            _ExtentY        =   2461
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   "PKId"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nº"
            Columns(1).DataField=   "intParcela"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Valor "
            Columns(2).DataField=   "dblValor"
            Columns(2).NumberFormat=   "Standard"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Vencimento"
            Columns(3).DataField=   "dtmDtVencimento"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "D.A."
            Columns(4).DataField=   "intLancamentoAlfaDAtiva"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Baixa"
            Columns(5).DataField=   "dtmDtPagamento"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Desc. Baixa"
            Columns(6).DataField=   "STRDESCRICAO"
            Columns(6).NumberFormat=   "Standard"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Observação"
            Columns(7).DataField=   "Strobservacao"
            Columns(7).NumberFormat=   "Standard"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1191"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1111"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2514"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2434"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1693"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1614"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=609"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=529"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2143"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2064"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=4868"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=4789"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=0"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=10769"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=10689"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=0"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTips        =   1
            CellTipsWidth   =   0
            MultiSelect     =   0
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
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
            _StyleDefs(16)  =   ":id=8,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
            _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
            _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
            _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=0"
            _StyleDefs(58)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(59)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(60)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(61)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(62)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
            _StyleDefs(63)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
            _StyleDefs(64)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
            _StyleDefs(65)  =   "Named:id=33:Normal"
            _StyleDefs(66)  =   ":id=33,.parent=0"
            _StyleDefs(67)  =   "Named:id=34:Heading"
            _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   ":id=34,.wraptext=-1"
            _StyleDefs(70)  =   "Named:id=35:Footing"
            _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   "Named:id=36:Selected"
            _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(74)  =   "Named:id=37:Caption"
            _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(76)  =   "Named:id=38:HighlightRow"
            _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(78)  =   "Named:id=39:EvenRow"
            _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(80)  =   "Named:id=40:OddRow"
            _StyleDefs(81)  =   ":id=40,.parent=33"
            _StyleDefs(82)  =   "Named:id=41:RecordSelector"
            _StyleDefs(83)  =   ":id=41,.parent=34"
            _StyleDefs(84)  =   "Named:id=42:FilterBar"
            _StyleDefs(85)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_Historico 
         Caption         =   "Histórico"
         Height          =   3795
         Left            =   -74850
         TabIndex        =   46
         Top             =   480
         Width           =   10035
         Begin VB.TextBox txtHistorico 
            Height          =   3345
            Left            =   120
            MaxLength       =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   47
            Top             =   270
            Width           =   9825
         End
      End
      Begin VB.Frame fra_receitas 
         Caption         =   "Receitas"
         Height          =   1425
         Left            =   120
         TabIndex        =   42
         Top             =   3330
         Width           =   10095
         Begin TrueOleDBGrid70.TDBGrid tdb_Receitas 
            Height          =   1125
            Left            =   60
            TabIndex        =   43
            Top             =   210
            Width           =   9945
            _ExtentX        =   17542
            _ExtentY        =   1984
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   "PKId"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Receita"
            Columns(1).DataField=   "STRRECEITA"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Indexador"
            Columns(2).DataField=   "STRINDEXADOR"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Qtde Index"
            Columns(3).DataField=   "DBLQTDEINDEXADOR"
            Columns(3).NumberFormat=   "FormatText Event"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "VL Indexador"
            Columns(4).DataField=   "DBLVALORINDEXADOR"
            Columns(4).NumberFormat=   "FormatText Event"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Valor"
            Columns(5).DataField=   "dblValorUnit"
            Columns(5).NumberFormat=   "FormatText Event"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Qtde Receita"
            Columns(6).DataField=   "INTQTDERECEITA"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Total"
            Columns(7).DataField=   "dblTotal"
            Columns(7).NumberFormat=   "FormatText Event"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=5186"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=5106"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1376"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1296"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=0"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1852"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1773"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2143"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2064"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2355"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2275"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1799"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1720"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=1"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=2"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTips        =   1
            CellTipsWidth   =   0
            MultiSelect     =   0
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
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
            _StyleDefs(16)  =   ":id=8,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=0"
            _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=70,.parent=13,.alignment=1"
            _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
            _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
            _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
            _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
            _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
            _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
            _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(58)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(59)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(60)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(61)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(62)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
            _StyleDefs(63)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
            _StyleDefs(64)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
            _StyleDefs(65)  =   "Named:id=33:Normal"
            _StyleDefs(66)  =   ":id=33,.parent=0"
            _StyleDefs(67)  =   "Named:id=34:Heading"
            _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   ":id=34,.wraptext=-1"
            _StyleDefs(70)  =   "Named:id=35:Footing"
            _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   "Named:id=36:Selected"
            _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(74)  =   "Named:id=37:Caption"
            _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(76)  =   "Named:id=38:HighlightRow"
            _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(78)  =   "Named:id=39:EvenRow"
            _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(80)  =   "Named:id=40:OddRow"
            _StyleDefs(81)  =   ":id=40,.parent=33"
            _StyleDefs(82)  =   "Named:id=41:RecordSelector"
            _StyleDefs(83)  =   ":id=41,.parent=34"
            _StyleDefs(84)  =   "Named:id=42:FilterBar"
            _StyleDefs(85)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame Fra_Titulo 
         Height          =   945
         Left            =   90
         TabIndex        =   1
         Top             =   330
         Width           =   10095
         Begin VB.TextBox txtstrCodigoP 
            Alignment       =   1  'Right Justify
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
            Left            =   6300
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   570
            Width           =   825
         End
         Begin VB.TextBox txtintExercicioP 
            Alignment       =   1  'Right Justify
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
            Left            =   7140
            MaxLength       =   4
            TabIndex        =   16
            Top             =   570
            Width           =   465
         End
         Begin VB.TextBox txtbitDigitoP 
            Alignment       =   1  'Right Justify
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
            Left            =   7620
            MaxLength       =   2
            TabIndex        =   17
            Top             =   570
            Width           =   285
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Left            =   8295
            TabIndex        =   9
            Top             =   150
            Width           =   1425
         End
         Begin VB.TextBox txtintExercicio 
            Height          =   285
            Left            =   6960
            MaxLength       =   8
            TabIndex        =   7
            Top             =   150
            Width           =   495
         End
         Begin VB.TextBox txtdtmdtVencimento 
            Height          =   285
            Left            =   4200
            TabIndex        =   13
            Top             =   570
            Width           =   1005
         End
         Begin VB.TextBox txtstrAviso 
            Height          =   285
            Left            =   4950
            MaxLength       =   10
            TabIndex        =   5
            Top             =   150
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dbc_intReceita 
            Height          =   315
            Left            =   1710
            TabIndex        =   3
            Top             =   150
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskstrInscricao 
            Height          =   285
            Left            =   1710
            TabIndex        =   11
            Top             =   570
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin VB.Label lblProcesso 
            AutoSize        =   -1  'True
            Caption         =   "Processo"
            Height          =   195
            Left            =   5550
            TabIndex        =   14
            Top             =   660
            Width           =   660
         End
         Begin VB.Label lbl_compreceita 
            AutoSize        =   -1  'True
            Caption         =   "Composição da receita"
            Height          =   195
            Left            =   60
            TabIndex        =   2
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label lbl_cadastro 
            AutoSize        =   -1  'True
            Caption         =   "Cadastro"
            Height          =   195
            Left            =   7590
            TabIndex        =   8
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lbl_strInscricao 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição"
            Height          =   195
            Left            =   1035
            TabIndex        =   10
            Top             =   660
            Width           =   645
         End
         Begin VB.Label lbl_exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   6210
            TabIndex        =   6
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lbl_inscricao 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento"
            Height          =   195
            Left            =   3180
            TabIndex        =   12
            Top             =   660
            Width           =   840
         End
         Begin VB.Label lbl_aviso 
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Left            =   4470
            TabIndex        =   4
            Top             =   240
            Width           =   390
         End
      End
      Begin VB.Frame fra_PrescricaoDoDebito 
         Caption         =   "Contribuinte"
         Height          =   915
         Left            =   120
         TabIndex        =   18
         Top             =   1260
         Width           =   10095
         Begin VB.TextBox txtstridentidade 
            Height          =   285
            Left            =   6780
            TabIndex        =   22
            Top             =   180
            Width           =   1155
         End
         Begin VB.TextBox txtstrcnpjcpf 
            Height          =   285
            Left            =   8760
            TabIndex        =   24
            Top             =   180
            Width           =   1155
         End
         Begin VB.TextBox txtstrnomeproprietario 
            Height          =   285
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   20
            Top             =   180
            Width           =   4635
         End
         Begin VB.TextBox txtstrpromissario 
            Height          =   285
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   26
            Top             =   540
            Width           =   8625
         End
         Begin VB.Label lbl_identidade 
            AutoSize        =   -1  'True
            Caption         =   "Identidade"
            Height          =   195
            Left            =   6000
            TabIndex        =   21
            Top             =   270
            Width           =   750
         End
         Begin VB.Label lbl_CNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ/CPF"
            Height          =   195
            Left            =   7950
            TabIndex        =   23
            Top             =   270
            Width           =   780
         End
         Begin VB.Label lbl_nome 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   450
            TabIndex        =   19
            Top             =   270
            Width           =   795
         End
         Begin VB.Label lbl_Prescricao 
            AutoSize        =   -1  'True
            Caption         =   "Promissário"
            Height          =   195
            Left            =   450
            TabIndex        =   25
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.Frame fra_DomicilioFiscal 
         Caption         =   "Local"
         Height          =   1155
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Width           =   10095
         Begin VB.TextBox txtstrComplemento 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   7470
            MaxLength       =   20
            TabIndex        =   33
            Top             =   150
            Width           =   1935
         End
         Begin VB.TextBox txtstrNumero 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   5850
            MaxLength       =   10
            TabIndex        =   31
            Top             =   150
            Width           =   1005
         End
         Begin VB.TextBox txt_Logradouro 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   29
            Top             =   150
            Width           =   4125
         End
         Begin VB.TextBox txt_Bairro 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   35
            Top             =   480
            Width           =   5535
         End
         Begin VB.TextBox txt_Municipio 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   39
            Top             =   810
            Width           =   5535
         End
         Begin VB.TextBox txt_Cep 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   7470
            MaxLength       =   20
            TabIndex        =   37
            Top             =   480
            Width           =   1005
         End
         Begin VB.TextBox txt_UF 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   7470
            MaxLength       =   2
            TabIndex        =   41
            Top             =   810
            Width           =   405
         End
         Begin VB.Label lbl_strComplemento 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6960
            TabIndex        =   32
            Top             =   180
            Width           =   480
         End
         Begin VB.Label lbl_numero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5520
            TabIndex        =   30
            Top             =   180
            Width           =   180
         End
         Begin VB.Label lbl_Logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   195
            Left            =   570
            TabIndex        =   28
            Top             =   180
            Width           =   690
         End
         Begin VB.Label lbl_Bairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   855
            TabIndex        =   34
            Top             =   510
            Width           =   405
         End
         Begin VB.Label lbl_Municipio 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   555
            TabIndex        =   38
            Top             =   840
            Width           =   705
         End
         Begin VB.Label lbl_Cep 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   7080
            TabIndex        =   36
            Top             =   540
            Width           =   285
         End
         Begin VB.Label lbl_UF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   7170
            TabIndex        =   40
            Top             =   840
            Width           =   210
         End
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1485
      Left            =   90
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6570
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   2619
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "IntPPublico"
      Columns(0).DataField=   "IntPPublico"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "IntLancamentoAlfa"
      Columns(1).DataField=   "IntLancamentoAlfa"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Comp. da Receita"
      Columns(2).DataField=   "Strcomposicaodareceita"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Exercício"
      Columns(3).DataField=   "Intexercicio"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Aviso"
      Columns(4).DataField=   "Strnumeroaviso"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   16
      Columns(5)._MaxComboItems=   5
      Columns(5).ValueItems(0)._DefaultItem=   0
      Columns(5).ValueItems(0).Value=   " "
      Columns(5).ValueItems(0).Value.vt=   8
      Columns(5).ValueItems(0).DisplayValue=   " "
      Columns(5).ValueItems(0).DisplayValue.vt=   8
      Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(1)._DefaultItem=   0
      Columns(5).ValueItems(1).Value=   "0"
      Columns(5).ValueItems(1).Value.vt=   8
      Columns(5).ValueItems(1).DisplayValue=   ""
      Columns(5).ValueItems(1).DisplayValue.vt=   8
      Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(2)._DefaultItem=   0
      Columns(5).ValueItems(2).Value=   "1"
      Columns(5).ValueItems(2).Value.vt=   8
      Columns(5).ValueItems(2).DisplayValue=   "Imobiliário"
      Columns(5).ValueItems(2).DisplayValue.vt=   8
      Columns(5).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(3)._DefaultItem=   0
      Columns(5).ValueItems(3).Value=   "2"
      Columns(5).ValueItems(3).Value.vt=   8
      Columns(5).ValueItems(3).DisplayValue=   "Mobiliário"
      Columns(5).ValueItems(3).DisplayValue.vt=   8
      Columns(5).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems.Count=   4
      Columns(5).Caption=   "Cadastro"
      Columns(5).DataField=   "strCadastro"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Inscrição"
      Columns(6).DataField=   "Strinscricao"
      Columns(6).NumberFormat=   "FormatText Event"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Vencimento"
      Columns(7).DataField=   "Dtmdtvencimento"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Total"
      Columns(8).DataField=   "dblTotal"
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "intUtilizacao"
      Columns(9).DataField=   "intUtilizacao"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=6244"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6165"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=1482"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1402"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=1693"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1614"
      Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=2037"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=1958"
      Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=2355"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2275"
      Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(41)=   "Column(7).Width=1667"
      Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=1588"
      Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=2461"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2381"
      Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(52)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(53)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(54)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(56)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(57)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(58)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(59)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTips        =   1
      CellTipsWidth   =   0
      MultiSelect     =   0
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
      _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(77)  =   "Named:id=33:Normal"
      _StyleDefs(78)  =   ":id=33,.parent=0"
      _StyleDefs(79)  =   "Named:id=34:Heading"
      _StyleDefs(80)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   ":id=34,.wraptext=-1"
      _StyleDefs(82)  =   "Named:id=35:Footing"
      _StyleDefs(83)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   "Named:id=36:Selected"
      _StyleDefs(85)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(86)  =   "Named:id=37:Caption"
      _StyleDefs(87)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(88)  =   "Named:id=38:HighlightRow"
      _StyleDefs(89)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(90)  =   "Named:id=39:EvenRow"
      _StyleDefs(91)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(92)  =   "Named:id=40:OddRow"
      _StyleDefs(93)  =   ":id=40,.parent=33"
      _StyleDefs(94)  =   "Named:id=41:RecordSelector"
      _StyleDefs(95)  =   ":id=41,.parent=34"
      _StyleDefs(96)  =   "Named:id=42:FilterBar"
      _StyleDefs(97)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadPrecoPublico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mobjAux                   As Object
Dim mblnClickOk               As Boolean
Dim mblnSelecionou            As Boolean
Dim mblnPrimeiraVez           As Boolean

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
'    If ColIndex = tdb_Lista.Columns("Identificação").Order Then
'        If tdb_Lista.Columns(5).Value = 1 Then
'            Value = gstrFormataInscricao(CStr(Value), TYP_IMOBILIARIA)
'        ElseIf tdb_Lista.Columns(5).Value = 2 Then
'            Value = gstrFormataInscricao(CStr(Value), TYP_ECONOMICA)
'        ElseIf tdb_Lista.Columns(5).Value = 0 Then
'            Value = gstrENulo(Value)
'        End If
'    End If
End Sub

Private Sub tdb_Receitas_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    
    Select Case ColIndex
        Case 3 'Qtde de Indexador
            Value = gstrConvVrDoSql(Value, 6, , True)
        Case 4 'Vl do indexador
            Value = gstrConvVrDoSql(Value, 6, , True)
        Case 5 'ValorUnit
            Value = gstrConvVrDoSql(Value, 6)
            Value = Mid(gstrConvVrDoSql(Value, 6), 1, InStr(Value, ",") - 1) & Mid(gstrConvVrDoSql(Value, 6), InStr(Value, ","), 3)
        Case 7 'Total
            Value = gstrConvVrDoSql(Value, 2, , True)
    End Select
End Sub

Private Sub txtbitDigitoP_GotFocus()
    MarcaCampo txtbitDigitoP
End Sub

Private Sub txtbitDigitoP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigitoP
End Sub

Private Sub txtHistorico_GotFocus()
    MarcaCampo txtHistorico
End Sub

Private Sub txtHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtHistorico
End Sub

Private Sub tdb_Parcelas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tdb_Parcelas.EOF And Not tdb_Parcelas.BOF Then
        gCorLinhaSelecionada tdb_Parcelas
    End If
End Sub

Private Sub txtintExercicioP_GotFocus()
    MarcaCampo txtintExercicioP
End Sub

Private Sub txtintExercicioP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicioP
End Sub

Private Sub txtstrcnpjcpf_GotFocus()
    MarcaCampo txtstrcnpjcpf
End Sub

Private Sub txtstrcnpjcpf_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrcnpjcpf
End Sub

Private Sub txtstrCodigoP_GotFocus()
    MarcaCampo txtstrCodigoP
End Sub

Private Sub txtstrCodigoP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigoP
End Sub

Private Sub txtstridentidade_GotFocus()
    MarcaCampo txtstridentidade
End Sub

Private Sub txtstrIdentidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstridentidade
End Sub


Private Sub dbc_intReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intReceita, Me, Area
End Sub

Private Sub dbc_intReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intReceita, Me, , KeyCode, Shift
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnPrimeiraVez = True
    mblnClickOk = True
End Sub

Private Sub txt_Bairro_GotFocus()
    MarcaCampo txt_Bairro
End Sub

Private Sub txt_Bairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Bairro
End Sub

Private Sub txt_Cep_GotFocus()
    MarcaCampo txt_Cep
End Sub

Private Sub txt_Cep_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txt_Cep
End Sub

Private Sub txt_Cep_LostFocus()
    txt_Cep = gstrCEPFormatado(txt_Cep)
    CepLogradouro txt_Cep, txt_Logradouro, txt_Bairro, txt_Municipio, txt_UF, , , True, False, False, False, False
End Sub

Private Sub txt_Logradouro_GotFocus()
    MarcaCampo txt_Logradouro
End Sub

Private Sub txt_Logradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Logradouro
End Sub

Private Sub txt_Municipio_GotFocus()
    MarcaCampo txt_Municipio
End Sub

Private Sub txt_Municipio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Municipio
End Sub

Private Sub txt_UF_GotFocus()
    MarcaCampo txt_UF
End Sub

Private Sub txt_UF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "U", txt_UF
End Sub

Private Sub txtcadastro_GotFocus()
    MarcaCampo txtcadastro
End Sub

Private Sub txtcadastro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtcadastro
End Sub

Private Sub txtdtmdtVencimento_GotFocus()
    If txtdtmdtVencimento = "" Then
        txtdtmdtVencimento = gstrDataFormatada(Date)
    End If
    MarcaCampo txtdtmdtVencimento
End Sub

Private Sub txtdtmdtVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdtVencimento
End Sub

Private Sub txtdtmdtVencimento_LostFocus()
    txtdtmdtVencimento = gstrDataFormatada(txtdtmdtVencimento)
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtstrAviso_GotFocus()
    MarcaCampo txtstrAviso
End Sub

Private Sub txtstrAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrAviso
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
End Sub

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub mskstrInscricao_GotFocus()
    MarcaCampo mskstrInscricao
End Sub

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricao
End Sub

Private Sub txtstrNomeProprietario_GotFocus()
    MarcaCampo txtstrnomeproprietario
End Sub

Private Sub txtstrNomeProprietario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrnomeproprietario
End Sub

Private Sub txtstrNumero_GotFocus()
    MarcaCampo txtstrNumero
End Sub

Private Sub txtstrNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumero
End Sub

Private Sub txtstrPromissario_GotFocus()
    MarcaCampo txtstrpromissario
End Sub

Private Sub txtstrPromissario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrpromissario
End Sub

Private Sub dbc_intReceita_GotFocus()
    MarcaCampo dbc_intReceita
End Sub

Private Sub dbc_intReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intReceita
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1241

    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    VerificaObjParaAplicar mobjAux
    dbc_intReceita.Tag = strQueryComposicaoReceita & ";strDescricao"
    TrocaCorObjeto txt_Logradouro, True
    TrocaCorObjeto txtstrNumero, True
    TrocaCorObjeto txtstrComplemento, True
    TrocaCorObjeto txt_Bairro, True
    TrocaCorObjeto txt_Cep, True
    TrocaCorObjeto txt_Municipio, True
    TrocaCorObjeto txt_UF, True
    TrocaCorObjeto txtcadastro, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Function strQueryComposicaoReceita()
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT PKId, Ltrim(Rtrim(strDescricao)) as strDescricao "
    strSql = strSql & "FROM " & gstrComposicaoDaReceita & " "
    strSql = strSql & "WHERE "
    strSql = strSql & "intUtilizacao = 5 "
    strSql = strSql & "ORDER BY strDescricao"
    
    strQueryComposicaoReceita = strSql
    
End Function

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
   gOrdenaGrid tdb_Lista, ColIndex
'   mblnPrimeiraVez = False
'   mblnClickOk = False
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If Not .EOF And Not .BOF Then
            'If mblnPrimeiraVez Then
                'If mblnClickOk Then
                    gCorLinhaSelecionada tdb_Lista
                    mblnClickOk = False
                    mblnSelecionou = True
                    gCorLinhaSelecionada tdb_Lista
                    txtPKId = .Columns(0).Value
                    PreencheCampos
                    LeDaTabelaParaObj "", tdb_Parcelas, strQueryParcela
                    LeDaTabelaParaObj "", tdb_Receitas, strQueryReceita
                    If mobjAux Is Nothing Then
                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                    End If
                'End If
            'End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(gstrImprimir) = UCase(strModoOperacao) Then

    ElseIf UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    ElseIf UCase(strModoOperacao) = gstrLocalizar Then
        LeDaTabelaParaObj "", tdb_Lista, strQuery(True)
    ElseIf UCase(strModoOperacao) = gstrRefresh Then
        LeDaTabelaParaObj "", tdb_Lista, strQuery(False)
    ElseIf UCase(gstrFechar) = UCase(strModoOperacao) Then
        Unload Me
    ElseIf UCase(gstrNovo) = UCase(strModoOperacao) Then
        Limpa_Controles Me, True, True, True, True, True
        mskstrInscricao.Mask = ""
        mskstrInscricao.Text = ""
        Set tdb_Parcelas.DataSource = Nothing
        Set tdb_Receitas.DataSource = Nothing
        tab_3DPasta.Tab = 0
        dbc_intReceita.SetFocus
    ElseIf UCase(gstrSalvar) = UCase(strModoOperacao) Then
    
    Else
        
    End If
End Sub

Private Sub PreencheCampos()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "LA.Pkid, "
    strSql = strSql & "LA.Intcomposicaodareceita, "
    strSql = strSql & "LPP.Intutilizacao, "
    strSql = strSql & gstrRIGHT("LPP.Strinscricao", gintRetornaTamanhoMascara(Val(tdb_Lista.Columns("intUtilizacao")))) & " strInscricao, "
    strSql = strSql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso,"
    strSql = strSql & "LA.intExercicio, "
    strSql = strSql & "LPP.Dtmdtvencimento, "
    strSql = strSql & "LA.strnomeproprietario, "
    strSql = strSql & "LA.strcnpjcpf, "
    strSql = strSql & "LA.stridentidade, "
    strSql = strSql & "LA.strlogradouro, "
    strSql = strSql & "LA.strnumero, "
    strSql = strSql & "LA.strcomplemento, "
    strSql = strSql & "LA.strbairro, "
    strSql = strSql & "LA.strmunicipio, "
    strSql = strSql & "LA.struf, "
    strSql = strSql & "LA.intcep, "
    strSql = strSql & "LA.strpromissario, "
    strSql = strSql & "LPP.STRCODIGO, "
    strSql = strSql & "LPP.INTEXERCICIO intExercicioP, "
    strSql = strSql & "LPP.bitDigito, "
    strSql = strSql & "LPP.strHistorico "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoPPublico & " LPP, "
    strSql = strSql & gstrComposicaoDaReceita & " CR "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LPP.Intlancamentoalfa AND "
    
    'Alterado por Salsis: Todo Lançamento Preço Público deve ter uma Composição da Receita
    'strSQL = strSQL & "CR.Pkid" & strOUTJSQLServer & "= LA.Intcomposicaodareceita" & strOUTJOracle & " AND "
    
    strSql = strSql & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strSql = strSql & "LA.pkid = " & tdb_Lista.Columns(1).Value
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                
                PreencherListaDeOpcoes dbc_intReceita, gstrENulo(!intComposicaoDaReceita)
                Select Case gstrENulo(!intUtilizacao)
                    Case 1
                        txtcadastro = "Imobiliário"
                        VerificaMascaraInscricao TYP_IMOBILIARIA
                    Case 2
                        txtcadastro = "Econômico"
                        VerificaMascaraInscricao TYP_ECONOMICA
                    Case Else
                        txtcadastro = ""
                        mskstrInscricao.Mask = ""
                End Select
                mskstrInscricao = gstrENulo(!strInscricao)
                txtintExercicio = gstrENulo(!intExercicio)
                txtstrAviso = gstrENulo(!strNumeroAviso)
                txtstrAviso.Text = Space$(0) & adoResultado!strNumeroAviso
                txtdtmdtVencimento = gstrENulo(!Dtmdtvencimento)
                txtstrnomeproprietario = gstrENulo(!strnomeproprietario)
                txtstrcnpjcpf = gstrENulo(!StrCnpjCpf)
                txtstridentidade = gstrENulo(!STRIDENTIDADE)
                txt_Logradouro = gstrENulo(!strLogradouro)
                txtstrNumero = gstrENulo(!strNumero)
                txtstrComplemento = gstrENulo(!STRCOMPLEMENTO)
                txt_Bairro = gstrENulo(!strBairro)
                txt_Municipio = gstrENulo(!STRMUNICIPIO)
                txt_UF = gstrENulo(!STRUF)
                txt_Cep = gstrCEPFormatado(gstrENulo(!INTCEP))
                txtstrpromissario = gstrENulo(!strpromissario)
                txtstrCodigoP = gstrENulo(!strCodigo)
                txtintExercicioP = gstrENulo(!intExercicioP)
                txtbitDigitoP = gstrENulo(!bitDigito)
                txtHistorico.Text = gstrENulo(!STRHISTORICO)
            End If
        End With
    End If
    
End Sub

Private Function strQuery(blnFiltro As Boolean) As String
    Dim strSql   As String
    Dim strIndex As String
    Dim strWhere As String
    
    'Indexes utilizados p/ otimização da consulta
    'LA.intComposicaoDaReceita --IDX_TBLLCTOALFA_INTCOMPRECEITA
    'LA.intExercicio --IDX_TBLLCTOALFA_INTEXERCICIO
    'LA.strNumeroAviso --IDX_TBLLCALFA_AVIS_COMPREC_ANO
    'LA.strInscricao --IDX_TBLLCTOALFA_STRINSCRICAO
    'LA.strNomeProprietario LIKE  --IDX_TBLLCALFA_PROPRIET
    'LA.strPromissario LIKE  --IDX_TBLLCALFA_PROMIS
    'LPP.Dtmdtvencimento --IDX_TBLLCTPP_DTMDTVENC
    
    If blnFiltro Then
        strIndex = strIndex & "/*+"
        
        'Index usado p/ o join
        strIndex = strIndex & " index (LPP IDX_TBLLCTPP_TBLLCTALFA)"
        
        If dbc_intReceita.MatchedWithList = True Then
           strWhere = strWhere & " AND LA.intComposicaoDaReceita = " & dbc_intReceita.BoundText
           strIndex = strIndex & " index (LA IDX_TBLLCTOALFA_INTCOMPRECEITA)"
        End If
        
        If Trim(txtintExercicio) <> "" Then
           strWhere = strWhere & " AND LA.intExercicio = " & txtintExercicio.Text
           strIndex = strIndex & " index (LA IDX_TBLLCTOALFA_INTEXERCICIO)"
        End If
        
        If Trim(txtstrAviso) <> "" Then
           strWhere = strWhere & " AND LA.strNumeroAviso ='" & UCase(String(gintLenNumAviso - Len(txtstrAviso), "0") & txtstrAviso.Text) & "'"
           strIndex = strIndex & " index (LA IDX_TBLLCALFA_AVIS_COMPREC_ANO)"
        End If
        
        If Trim(mskstrInscricao) <> "" Then
           strWhere = strWhere & " AND LA.Strinscricao ='" & String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text) & "'"
           strIndex = strIndex & " index (LA IDX_TBLLCTOALFA_STRINSCRICAO)"
        End If
        
        If Trim(txtstrnomeproprietario) <> "" Then
           strWhere = strWhere & " AND LA.strnomeproprietario Like '" & UCase(txtstrnomeproprietario.Text) & "%'"
           strIndex = strIndex & " index (LA IDX_TBLLCALFA_PROPRIET)"
        End If
        
        If Trim(txtstrpromissario) <> "" Then
           strWhere = strWhere & " AND LA.strpromissario Like '" & UCase(txtstrpromissario.Text) & "%'"
           strIndex = strIndex & " index (LA IDX_TBLLCALFA_PROMIS)"
        End If
        
        If Trim(txtdtmdtVencimento) <> "" Then
           strWhere = strWhere & " AND LPP.Dtmdtvencimento = " & gstrConvDtParaSql(txtdtmdtVencimento.Text)
           strIndex = strIndex & " index (LPP IDX_TBLLCTPP_DTMDTVENC)"
        End If
        
        strIndex = strIndex & " */ "
    End If
    
    
    strSql = "Select "
    If bytDBType = Oracle Then
       strSql = strSql & strIndex 'Parâmetro adicional inserido para otimizar a consulta à pedido do DBA
    End If
    strSql = strSql & "LPP.Pkid As IntPPublico, "
    strSql = strSql & "LA.Pkid As IntLancamentoAlfa, "
    strSql = strSql & "LA.Strcomposicaodareceita strComposicaoDaReceita, "
    strSql = strSql & "LA.Intexercicio, "
    strSql = strSql & "LPP.intUtilizacao, "
    strSql = strSql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, "
    strSql = strSql & "'Imobiliário' As strCadastro, "
    strSql = strSql & gstrRIGHT("LPP.Strinscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
    strSql = strSql & "LPP.Dtmdtvencimento, "
    strSql = strSql & "LPP.DBLVALOR As dblTotal "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoPPublico & " LPP "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LPP.Intlancamentoalfa AND "
    strSql = strSql & "lpp.intutilizacao = " & TYP_IMOBILIARIA & " AND "
    strSql = strSql & "LA.intUtilizacao = " & TYP_IMOBILIARIA
    strSql = strSql & strWhere
    
    strSql = strSql & " UNION "
    
    strSql = strSql & "Select "
    If bytDBType = Oracle Then
       strSql = strSql & strIndex 'Parâmetro adicional inserido para otimizar a consulta à pedido do DBA
    End If
    strSql = strSql & "LPP.Pkid As IntPPublico, "
    strSql = strSql & "LA.Pkid As IntLancamentoAlfa, "
    strSql = strSql & "LA.Strcomposicaodareceita strComposicaoDaReceita, "
    strSql = strSql & "LA.Intexercicio, "
    strSql = strSql & "LPP.intUtilizacao, "
    strSql = strSql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, "
    strSql = strSql & "'Econômico' As strCadastro, "
    strSql = strSql & gstrRIGHT("LPP.Strinscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao, "
    strSql = strSql & "LPP.Dtmdtvencimento, "
    strSql = strSql & "LPP.DBLVALOR As dblTotal "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoPPublico & " LPP "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LPP.Intlancamentoalfa "
    strSql = strSql & " AND lpp.intutilizacao = " & TYP_ECONOMICA & " AND "
    strSql = strSql & "LA.intUtilizacao = " & TYP_ECONOMICA
    strSql = strSql & strWhere
    
    strSql = strSql & " UNION "
    
    strSql = strSql & "Select "
    If bytDBType = Oracle Then
       strSql = strSql & strIndex 'Parâmetro adicional inserido para otimizar a consulta à pedido do DBA
    End If
    strSql = strSql & "LPP.Pkid As IntPPublico, "
    strSql = strSql & "LA.Pkid As IntLancamentoAlfa, "
    strSql = strSql & "LA.Strcomposicaodareceita strComposicaoDaReceita, "
    strSql = strSql & "LA.Intexercicio, "
    strSql = strSql & "LPP.intUtilizacao, "
    strSql = strSql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, "
    strSql = strSql & "'' As strCadastro, "
    strSql = strSql & gstrRIGHT("LPP.Strinscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao, "
    strSql = strSql & "LPP.Dtmdtvencimento, "
    strSql = strSql & "LPP.DBLVALOR As dblTotal "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoPPublico & " LPP "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LPP.Intlancamentoalfa "
    strSql = strSql & " AND (lpp.intutilizacao = 0 or lpp.intutilizacao is null)"
    'strSql = strSql & " AND (LA.intutilizacao = 0 or LA.intutilizacao is null)"
    strSql = strSql & strWhere
    
    strSql = strSql & " ORDER BY strComposicaoDaReceita ASC"
    
    strQuery = strSql
    
End Function

Private Function strQueryParcela() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "LV.Pkid, "
    strSql = strSql & "LV.intParcela, "
    strSql = strSql & "LV.dblValor, "
    strSql = strSql & "CASE WHEN LV.intLancamentoAlfaDAtiva IS NULL THEN '' ELSE 'X' END intLancamentoAlfaDAtiva ,"
    strSql = strSql & "LV.dtmDtVencimento, "
    strSql = strSql & "LP.dtmDtPagamento, "
    strSql = strSql & "CB.STRDESCRICAO, "
    strSql = strSql & "LP.Strobservacao  "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrCodigoDeBaixa & " CB, "
    strSql = strSql & gstrMoedas & " M, "
    strSql = strSql & gstrLancamentoPagamento & " LP "
    strSql = strSql & "WHERE LV.Pkid =" & " LP.intLancamentoValor " & strOUTJOracle & " "
    strSql = strSql & "AND CB.Pkid " & strOUTJOracle & "=" & " LP.Intcodigobaixa And "
    strSql = strSql & "M.Pkid = LV.Intmoeda AND "
    strSql = strSql & "LV.intLancamentoAlfa = " & tdb_Lista.Columns(1).Value & " "
    
    
    ' Esta condição foi especificada para complementar a query em sqlserver que não suporta o
    ' join feita na consulta acima
    If bytDBType = SQLServer Then
        strSql = strSql & " UNION ALL "
        
        strSql = strSql & "SELECT "
        strSql = strSql & "LV.Pkid, "
        strSql = strSql & "LV.intParcela, "
        strSql = strSql & "LV.dblValor, "
        strSql = strSql & "CASE WHEN LV.intLancamentoAlfaDAtiva IS NULL THEN '' ELSE 'X' END intLancamentoAlfaDAtiva ,"
        strSql = strSql & "LV.dtmDtVencimento, "
        strSql = strSql & "NULL, "
        strSql = strSql & "'', "
        strSql = strSql & "''  "
        strSql = strSql & "FROM "
        strSql = strSql & gstrLancamentoValor & " LV, "
        strSql = strSql & gstrMoedas & " M "
        strSql = strSql & "WHERE "
        strSql = strSql & "M.Pkid = LV.Intmoeda AND "
        strSql = strSql & "LV.pkId NOT IN (SELECT intLancamentoValor FROM tblLancamentoPagamento) "
        strSql = strSql & "AND LV.intLancamentoAlfa = " & tdb_Lista.Columns(1).Value & " "
    End If
    
    strSql = strSql & "ORDER BY LV.intParcela"
    
    strQueryParcela = strSql
    
End Function

Sub VerificaMascaraInscricao(intTipoComposicao As Integer)
Dim strSql As String
Dim adoResultado As ADODB.Recordset
Dim strMascara   As String
    
    strMascara = ""
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & intTipoComposicao
    strSql = strSql & " Order By intSequencia"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    mskstrInscricao.Mask = strMascara
    
End Sub

Private Function strQueryReceita() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "LPPR.PKID, "
    strSql = strSql & "LPPR.STRRECEITA, "
    strSql = strSql & "LPPR.STRINDEXADOR, "
    strSql = strSql & "LPPR.DBLQTDEINDEXADOR, "
    strSql = strSql & "LPPR.DBLVALORINDEXADOR, "
    strSql = strSql & gstrCASEWHEN("LPPR.DBLVALORINDEXADOR", "0, LPPR.DBLVALOR / LPPR.INTQTDERECEITA", gstrISNULL("LPPR.DBLQTDEINDEXADOR", "0") & "*" & gstrISNULL("LPPR.DBLVALORINDEXADOR", "0")) & " As dblValorUnit, "
    strSql = strSql & "LPPR.INTQTDERECEITA, "
    strSql = strSql & "LPPR.DBLVALOR As dblTotal "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoPPublicoReceita & " LPPR "
    strSql = strSql & "WHERE "
    strSql = strSql & "LPPR.Intlancamentoppublico = " & txtPKId & " "
    strSql = strSql & "ORDER BY LPPR.STRRECEITA"
    
    strQueryReceita = strSql
    
End Function
