VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadImobiliarioRural 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imobiliário Rural"
   ClientHeight    =   6480
   ClientLeft      =   90
   ClientTop       =   3225
   ClientWidth     =   8535
   HelpContextID   =   117
   Icon            =   "frmCadImobiliarioRural.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8535
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6255
      Left            =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   120
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   529
      TabCaption(0)   =   "Imobiliário Rural"
      TabPicture(0)   =   "frmCadImobiliarioRural.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Contribuinte"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdb_ImobiliarioRural"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Arrendatário"
      TabPicture(1)   =   "frmCadImobiliarioRural.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_Inscricao"
      Tab(1).Control(1)=   "txt_Proprietario"
      Tab(1).Control(2)=   "fra_Promissario"
      Tab(1).Control(3)=   "lbl_Label"
      Tab(1).Control(4)=   "lbl_Label1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Geral"
      TabPicture(2)   =   "frmCadImobiliarioRural.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt_Proprietario2"
      Tab(2).Control(1)=   "txt_Inscricao2"
      Tab(2).Control(2)=   "lvw_Caracteristica(0)"
      Tab(2).Control(3)=   "lvw_Detalhe(0)"
      Tab(2).Control(4)=   "lbl_Proprietario2"
      Tab(2).Control(5)=   "lbl_InscricaoCadastral3"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Terreno"
      TabPicture(3)   =   "frmCadImobiliarioRural.frx":1096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txt_Proprietario3"
      Tab(3).Control(1)=   "txt_Inscricao3"
      Tab(3).Control(2)=   "lvw_Caracteristica(1)"
      Tab(3).Control(3)=   "lvw_Detalhe(1)"
      Tab(3).Control(4)=   "lbl_Proprietario3"
      Tab(3).Control(5)=   "lbl_InscricaoCadastral4"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Construção"
      TabPicture(4)   =   "frmCadImobiliarioRural.frx":10B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txt_Proprietario4"
      Tab(4).Control(1)=   "txt_Inscricao4"
      Tab(4).Control(2)=   "lvw_Caracteristica(2)"
      Tab(4).Control(3)=   "lvw_Detalhe(2)"
      Tab(4).Control(4)=   "lbl_Proprietario4"
      Tab(4).Control(5)=   "lbl_InscricaoCadastral5"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Produção"
      TabPicture(5)   =   "frmCadImobiliarioRural.frx":10CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txt_Inscricao6"
      Tab(5).Control(1)=   "txt_Proprietario6"
      Tab(5).Control(2)=   "fra_Producao"
      Tab(5).Control(3)=   "lbl_Ble"
      Tab(5).Control(4)=   "lbl_Bla"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Históricos"
      TabPicture(6)   =   "frmCadImobiliarioRural.frx":10EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txt_Proprietario5"
      Tab(6).Control(1)=   "txt_Inscricao5"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "fra_Valores"
      Tab(6).Control(3)=   "fra_Historico"
      Tab(6).Control(4)=   "lbl_KJKHU"
      Tab(6).Control(5)=   "lbl_kjhkjh"
      Tab(6).ControlCount=   6
      Begin TrueOleDBGrid70.TDBGrid tdb_ImobiliarioRural 
         Height          =   2970
         Left            =   120
         TabIndex        =   109
         Top             =   3165
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   5239
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   "bytNaturezaJuridica"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Código Reduzido"
         Columns(1).DataField=   "UM"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Inscrição Anterior"
         Columns(2).DataField=   "strInscricaoAnterior"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Contribuinte"
         Columns(3).DataField=   "strNome"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "CNPJ / CPF"
         Columns(4).DataField=   "strCNPJCPF"
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2355"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2275"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2752"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2672"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=6059"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=5980"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=2540"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2461"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txt_Inscricao6 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   85
         Top             =   540
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario6 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   84
         Top             =   900
         Width           =   4650
      End
      Begin VB.Frame fra_Producao 
         Caption         =   "Produção Rural"
         Height          =   4695
         Left            =   -74850
         TabIndex        =   82
         Top             =   1380
         Width           =   8025
         Begin VB.TextBox txt_ValorTotalEstimado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   6585
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   4290
            Width           =   1290
         End
         Begin TrueOleDBGrid70.TDBDropDown tdd_UnidadeMedida 
            Height          =   1695
            Left            =   2070
            TabIndex        =   89
            Top             =   2250
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   2990
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   "strDescricao"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Unidades de Medida"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
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
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   4
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   ""
            DataField       =   ""
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
            ValueTranslate  =   0   'False
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
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBDropDown tdd_Detalhe 
            Height          =   1695
            Left            =   1200
            TabIndex        =   88
            Top             =   1500
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   2990
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
            Columns(1).Caption=   "Detalhe da característica"
            Columns(1).DataField=   "strNomeDoDetalhe"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
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
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   4
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   ""
            DataField       =   ""
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
            ValueTranslate  =   0   'False
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
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBDropDown tdd_Caracteristica 
            Height          =   1695
            Left            =   480
            TabIndex        =   83
            Top             =   810
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   2990
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   "strNomeDaCaracteristica"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Características"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
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
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   4
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   ""
            DataField       =   ""
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
            ValueTranslate  =   0   'False
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
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Producao 
            Height          =   3885
            Left            =   150
            TabIndex        =   24
            Top             =   300
            Width           =   7725
            _ExtentX        =   13626
            _ExtentY        =   6853
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Característica"
            Columns(0).DataField=   ""
            Columns(0).DropDown=   "tdd_Caracteristica"
            Columns(0).DropDown.vt=   8
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Detalhe"
            Columns(1).DataField=   ""
            Columns(1).NumberFormat=   "Standard"
            Columns(1).EditMaskUpdate=   -1  'True
            Columns(1).EditMaskRight=   -1  'True
            Columns(1).DropDown=   "tdd_Detalhe"
            Columns(1).DropDown.vt=   8
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Qtd."
            Columns(2).DataField=   ""
            Columns(2).DataWidth=   10
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Und. de Medida"
            Columns(3).DataField=   ""
            Columns(3).DropDown=   "tdd_UnidadeMedida"
            Columns(3).DropDown.vt=   8
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Valor Estimado"
            Columns(4).DataField=   ""
            Columns(4).DataWidth=   15
            Columns(4).NumberFormat=   "Standard"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3122"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3043"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=3201"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3122"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=260"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1429"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1349"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=260"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2143"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2064"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=260"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=260"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowDelete     =   -1  'True
            AllowAddNew     =   -1  'True
            DataMode        =   4
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTips        =   1
            CellTipsWidth   =   0
            DeadAreaBackColor=   -2147483644
            RowDividerColor =   11059392
            RowSubDividerColor=   11059392
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
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=67,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=69,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=72,.parent=6,.namedParent=38"
            _StyleDefs(26)  =   "Splits(0).EditorStyle:id=71,.parent=7"
            _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
            _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
            _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
            _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
            _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=82,.parent=67"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68,.alignment=0"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=86,.parent=67"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=68,.alignment=0"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=69"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=71"
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=102,.parent=67"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=68,.alignment=0"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=69"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=71"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=106,.parent=67"
            _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=103,.parent=68,.alignment=0"
            _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=104,.parent=69"
            _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=105,.parent=71"
            _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=110,.parent=67"
            _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=107,.parent=68,.alignment=0"
            _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=108,.parent=69"
            _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=109,.parent=71"
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
         Begin VB.Label lbl_ValorTotal 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total Estimado"
            Height          =   195
            Left            =   5040
            TabIndex        =   91
            Top             =   4380
            Width           =   1455
         End
      End
      Begin VB.TextBox txt_Proprietario4 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   78
         Top             =   900
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao4 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   77
         Top             =   540
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario3 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   74
         Top             =   900
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao3 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   73
         Top             =   540
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario2 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   70
         Top             =   900
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao2 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   69
         Top             =   540
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario5 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   66
         Top             =   900
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao5 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   540
         Width           =   2130
      End
      Begin VB.Frame fra_Valores 
         Caption         =   "Valores"
         Height          =   915
         Left            =   -74850
         TabIndex        =   59
         Top             =   5190
         Width           =   8025
         Begin VB.TextBox txtdblValorEdificacao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            MaxLength       =   25
            TabIndex        =   27
            Top             =   180
            Width           =   1770
         End
         Begin VB.TextBox txtdblValorTerreno 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6120
            MaxLength       =   25
            TabIndex        =   28
            Top             =   180
            Width           =   1770
         End
         Begin VB.TextBox txtdblValorImovel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   525
            Width           =   1770
         End
         Begin VB.TextBox txtdblValorITBI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6120
            MaxLength       =   25
            TabIndex        =   29
            Top             =   525
            Width           =   1770
         End
         Begin VB.Label lbldblValorEdificacao 
            AutoSize        =   -1  'True
            Caption         =   "Valor venal da edificação"
            Height          =   195
            Left            =   135
            TabIndex        =   64
            Top             =   270
            Width           =   1800
         End
         Begin VB.Label lbldblValorTerreno 
            AutoSize        =   -1  'True
            Caption         =   "Valor venal do Terreno"
            Height          =   195
            Left            =   4365
            TabIndex        =   63
            Top             =   270
            Width           =   1620
         End
         Begin VB.Label lbldblValorImovel 
            AutoSize        =   -1  'True
            Caption         =   "Valor venal do Imóvel"
            Height          =   195
            Left            =   405
            TabIndex        =   62
            Top             =   615
            Width           =   1530
         End
         Begin VB.Label lbldblValorITBI 
            AutoSize        =   -1  'True
            Caption         =   "Valor para efeito de ITBI"
            Height          =   195
            Left            =   4260
            TabIndex        =   61
            Top             =   615
            Width           =   1725
         End
      End
      Begin VB.Frame fra_Historico 
         Caption         =   "Históricos"
         Height          =   3930
         Left            =   -74835
         TabIndex        =   56
         Top             =   1230
         Width           =   8025
         Begin VB.TextBox txt_Historico 
            Height          =   1440
            Left            =   150
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   225
            Width           =   7740
         End
         Begin Threed.SSPanel ssp_TipoComunicacao 
            Height          =   390
            Left            =   150
            TabIndex        =   57
            Top             =   1800
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   688
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin MSComctlLib.Toolbar tlb_Historico 
               Height          =   330
               Left            =   30
               TabIndex        =   58
               Top             =   30
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               ImageList       =   "img_Aux"
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   3
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Novo"
                     Object.ToolTipText     =   "Novo"
                     ImageIndex      =   1
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Salvar"
                     Object.ToolTipText     =   "Adicionar / Alterar"
                     ImageIndex      =   2
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Deletar"
                     Object.ToolTipText     =   "Remover"
                     ImageIndex      =   3
                  EndProperty
               EndProperty
            End
         End
         Begin MSComctlLib.ListView lvw_Historico 
            Height          =   1545
            Left            =   90
            TabIndex        =   31
            Top             =   2280
            Width           =   7830
            _ExtentX        =   13811
            _ExtentY        =   2725
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ImageList img_Aux 
            Left            =   1440
            Top             =   1680
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCadImobiliarioRural.frx":1106
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCadImobiliarioRural.frx":1266
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCadImobiliarioRural.frx":13C2
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txt_Inscricao 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         Top             =   540
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Top             =   900
         Width           =   4650
      End
      Begin VB.Frame fra_Promissario 
         Height          =   2640
         Left            =   -74880
         TabIndex        =   42
         Top             =   1290
         Width           =   8070
         Begin VB.Frame fra_Endereco 
            Caption         =   "Endereço de Correspondência"
            Height          =   1395
            Left            =   90
            TabIndex        =   92
            Top             =   1170
            Width           =   7875
            Begin VB.TextBox txt_Distrito 
               Height          =   285
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   100
               Top             =   960
               Width           =   3525
            End
            Begin VB.TextBox txt_Bairro 
               Height          =   285
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   99
               Top             =   600
               Width           =   3105
            End
            Begin VB.TextBox txt_Logradouro 
               Height          =   285
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   98
               Top             =   240
               Width           =   4005
            End
            Begin VB.TextBox txt_Numero 
               Height          =   285
               Left            =   5460
               MaxLength       =   8
               TabIndex        =   97
               Top             =   240
               Width           =   795
            End
            Begin VB.TextBox txt_Complemento 
               Height          =   285
               Left            =   6870
               MaxLength       =   20
               TabIndex        =   96
               Top             =   240
               Width           =   870
            End
            Begin VB.TextBox txt_Cep 
               Height          =   285
               Left            =   6675
               MaxLength       =   9
               TabIndex        =   95
               Top             =   960
               Width           =   1080
            End
            Begin VB.TextBox txt_Municipio 
               Height          =   285
               Left            =   5130
               MaxLength       =   50
               TabIndex        =   94
               Top             =   600
               Width           =   2625
            End
            Begin VB.TextBox txt_UF 
               Height          =   285
               Left            =   5130
               MaxLength       =   2
               TabIndex        =   93
               Top             =   960
               Width           =   510
            End
            Begin VB.Label lblstrDistritoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               Height          =   195
               Left            =   450
               TabIndex        =   108
               Top             =   1050
               Width           =   480
            End
            Begin VB.Label lblintMunicipioC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   4290
               TabIndex        =   107
               Top             =   690
               Width           =   705
            End
            Begin VB.Label lblintBairroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   525
               TabIndex        =   106
               Top             =   690
               Width           =   405
            End
            Begin VB.Label lblintLogradouroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   120
               TabIndex        =   105
               Top             =   330
               Width           =   810
            End
            Begin VB.Label lblintNumeroC 
               AutoSize        =   -1  'True
               Caption         =   "Nº"
               Height          =   195
               Left            =   5190
               TabIndex        =   104
               Top             =   330
               Width           =   180
            End
            Begin VB.Label lblstrComplementoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6360
               TabIndex        =   103
               Top             =   330
               Width           =   480
            End
            Begin VB.Label lblintUFC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   4770
               TabIndex        =   102
               Top             =   1065
               Width           =   210
            End
            Begin VB.Label lblintCepC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   6240
               TabIndex        =   101
               Top             =   1050
               Width           =   285
            End
         End
         Begin VB.CommandButton cmd_Contribuinte 
            Height          =   315
            Left            =   5910
            Picture         =   "frmCadImobiliarioRural.frx":151E
            Style           =   1  'Graphical
            TabIndex        =   46
            TabStop         =   0   'False
            Tag             =   "15"
            ToolTipText     =   "Ativa cadastro de Contribuintes"
            Top             =   240
            Width           =   360
         End
         Begin VB.OptionButton optbytNaturezaJuridica 
            BackColor       =   &H80000004&
            Caption         =   "Pessoa Jurídica"
            Height          =   240
            Index           =   1
            Left            =   4410
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   690
            Width           =   1500
         End
         Begin VB.OptionButton optbytNaturezaJuridica 
            BackColor       =   &H80000004&
            Caption         =   "Pessoa Física"
            Height          =   315
            Index           =   0
            Left            =   3045
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   660
            Width           =   1365
         End
         Begin VB.TextBox txtstrCNPJCPFP 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   660
            Width           =   1845
         End
         Begin MSDataListLib.DataCombo dbcintArrendatario 
            Height          =   315
            Left            =   1020
            TabIndex        =   17
            Top             =   255
            Width           =   4860
            _ExtentX        =   8573
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin VB.Label lblstrCNPJCPFP 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   150
            TabIndex        =   48
            Top             =   735
            Width           =   780
         End
         Begin VB.Label lblintArrendatario 
            AutoSize        =   -1  'True
            Caption         =   "Arrendatário"
            Height          =   195
            Left            =   75
            TabIndex        =   47
            Top             =   375
            Width           =   855
         End
      End
      Begin VB.Frame fra_Contribuinte 
         Height          =   2745
         Left            =   120
         TabIndex        =   26
         Top             =   330
         Width           =   8070
         Begin VB.TextBox txtstrAreaPropriedade 
            Height          =   285
            Left            =   4215
            MaxLength       =   15
            TabIndex        =   15
            Top             =   2280
            Width           =   1080
         End
         Begin VB.TextBox txtstrAreaConstruida 
            Height          =   285
            Left            =   1545
            MaxLength       =   15
            TabIndex        =   14
            Top             =   2280
            Width           =   1080
         End
         Begin VB.TextBox txtstrCNPJCPF 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   6600
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   705
            Width           =   1365
         End
         Begin VB.TextBox txt_PKIdContribuinte 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   1545
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   675
            Width           =   885
         End
         Begin VB.TextBox txtintNumero 
            Height          =   285
            Left            =   1545
            MaxLength       =   6
            TabIndex        =   8
            Top             =   1485
            Width           =   1035
         End
         Begin VB.TextBox txtintCep 
            Height          =   285
            Left            =   5910
            MaxLength       =   9
            TabIndex        =   12
            Top             =   1890
            Width           =   1005
         End
         Begin VB.TextBox txtstrComplemento 
            Height          =   285
            Left            =   3240
            MaxLength       =   10
            TabIndex        =   9
            Top             =   1485
            Width           =   1005
         End
         Begin VB.TextBox txtPKId 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   315
            Width           =   2130
         End
         Begin MSMask.MaskEdBox mskstrInscricaoAnterior 
            Height          =   285
            Left            =   5835
            TabIndex        =   1
            Top             =   315
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo cbointContribuinte 
            Height          =   315
            Left            =   2430
            TabIndex        =   3
            Top             =   675
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cbointMunicipio 
            Height          =   315
            Left            =   1545
            TabIndex        =   11
            Top             =   1860
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintLogradouro 
            Height          =   315
            Left            =   3690
            TabIndex        =   7
            Top             =   1080
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTipoC 
            Height          =   315
            Left            =   1545
            TabIndex        =   5
            Top             =   1080
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTituloC 
            Height          =   315
            Left            =   2265
            TabIndex        =   6
            Top             =   1080
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintBairro 
            Height          =   315
            Left            =   4845
            TabIndex        =   10
            Top             =   1455
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintUf 
            Height          =   315
            Left            =   7320
            TabIndex        =   13
            Top             =   1860
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintOcorrrencia 
            Height          =   315
            Left            =   6240
            TabIndex        =   16
            Top             =   2265
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin VB.Label lblintUf 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   7050
            TabIndex        =   81
            Top             =   1980
            Width           =   210
         End
         Begin VB.Label lblintMunicipio 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   765
            TabIndex        =   55
            Top             =   1995
            Width           =   705
         End
         Begin VB.Label lblstrAreaPropriedade 
            AutoSize        =   -1  'True
            Caption         =   "Área da propriedade"
            Height          =   195
            Left            =   2700
            TabIndex        =   54
            Top             =   2370
            Width           =   1440
         End
         Begin VB.Label lblstrAreaConstruida 
            AutoSize        =   -1  'True
            Caption         =   "Área construída"
            Height          =   195
            Left            =   330
            TabIndex        =   53
            Top             =   2370
            Width           =   1140
         End
         Begin VB.Label lblintOcorrencia 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrência"
            Height          =   195
            Left            =   5400
            TabIndex        =   41
            Top             =   2370
            Width           =   780
         End
         Begin VB.Label lblintCep 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Left            =   5550
            TabIndex        =   40
            Top             =   1980
            Width           =   315
         End
         Begin VB.Label lblstrComplemento 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   2670
            TabIndex        =   39
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label lblintNumero 
            AutoSize        =   -1  'True
            Caption         =   "N°"
            Height          =   195
            Left            =   1290
            TabIndex        =   38
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label lblintBairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   4395
            TabIndex        =   37
            Top             =   1560
            Width           =   405
         End
         Begin VB.Label lblintLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   660
            TabIndex        =   36
            Top             =   1200
            Width           =   810
         End
         Begin VB.Label lblstrCNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   5745
            TabIndex        =   35
            Top             =   780
            Width           =   780
         End
         Begin VB.Label lblintContribunte 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   675
            TabIndex        =   34
            Top             =   780
            Width           =   795
         End
         Begin VB.Label lblPKId 
            AutoSize        =   -1  'True
            Caption         =   "Código Reduzido"
            Height          =   195
            Left            =   255
            TabIndex        =   33
            Top             =   405
            Width           =   1215
         End
         Begin VB.Label lblstrInscricaoAnterior 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   4425
            TabIndex        =   32
            Top             =   405
            Width           =   1350
         End
      End
      Begin MSComctlLib.ListView lvw_Caracteristica 
         Height          =   4665
         Index           =   0
         Left            =   -74880
         TabIndex        =   18
         Top             =   1470
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Característica"
            Object.Width           =   6368
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Detalhe 
         Height          =   4665
         Index           =   0
         Left            =   -70830
         TabIndex        =   19
         Top             =   1470
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Detalhes"
            Object.Width           =   6368
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Caracteristica 
         Height          =   4665
         Index           =   1
         Left            =   -74880
         TabIndex        =   20
         Top             =   1470
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Característica"
            Object.Width           =   6368
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Detalhe 
         Height          =   4665
         Index           =   1
         Left            =   -70830
         TabIndex        =   21
         Top             =   1470
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Detalhes"
            Object.Width           =   6368
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Caracteristica 
         Height          =   4665
         Index           =   2
         Left            =   -74880
         TabIndex        =   22
         Top             =   1470
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Característica"
            Object.Width           =   6368
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Detalhe 
         Height          =   4665
         Index           =   2
         Left            =   -70860
         TabIndex        =   23
         Top             =   1470
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Detalhes"
            Object.Width           =   6368
         EndProperty
      End
      Begin VB.Label lbl_Ble 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74700
         TabIndex        =   87
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label lbl_Bla 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -74145
         TabIndex        =   86
         Top             =   1005
         Width           =   795
      End
      Begin VB.Label lbl_Proprietario4 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -74235
         TabIndex        =   80
         Top             =   975
         Width           =   795
      End
      Begin VB.Label lbl_InscricaoCadastral5 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74790
         TabIndex        =   79
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label lbl_Proprietario3 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -74235
         TabIndex        =   76
         Top             =   975
         Width           =   795
      End
      Begin VB.Label lbl_InscricaoCadastral4 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74790
         TabIndex        =   75
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label lbl_Proprietario2 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -74235
         TabIndex        =   72
         Top             =   975
         Width           =   795
      End
      Begin VB.Label lbl_InscricaoCadastral3 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74790
         TabIndex        =   71
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label lbl_KJKHU 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -74235
         TabIndex        =   68
         Top             =   975
         Width           =   795
      End
      Begin VB.Label lbl_kjhkjh 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74790
         TabIndex        =   67
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label lbl_Label 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74790
         TabIndex        =   52
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label lbl_Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -74235
         TabIndex        =   51
         Top             =   975
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCadImobiliarioRural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim PKId_Temporario    As Double
    Dim mblnPrimeiraVez    As Boolean
    Dim mblnAlterando      As Boolean
    Dim mcboAux            As ComboBox
    Dim adoResultado       As ADODB.Recordset
    Dim strSql             As String
    Dim intMaxPKId         As Integer
    Dim oList              As Object
    Dim mblnSelecionou     As Boolean
    Dim objList            As Object
    Dim objList1           As Object
    Dim adoRec             As ADODB.Recordset
    Dim adoTdb             As ADODB.Recordset
    Dim X                  As XArrayDB
    Dim Y                  As XArrayDB
    Dim Z                  As XArrayDB
    Dim A                  As XArrayDB
    Dim intCodImobiliario  As Integer
    'Auxiliar para montagem do grid de detalhes da característica
    Dim xDet               As XArrayDB
    Dim blnButaoNovo       As Boolean
    'Guarda o tipo do imóvel de acordo com o tab selecionado
    '1 = Imobiliário Geral
    '2 = Imobiliário Terreno
    '3 = Imobiliário Construção
    Dim intCaractImobil    As Integer
    Dim intIndiceLVW       As Integer
    Dim intCodContribuinte As Integer
    'Variáveis para somar os campos respectivos
    Dim dblEdificacao      As Double
    Dim dblTerreno         As Double
    Dim txtDataReforma     As TextBox
    Dim mobjAux            As Object

Private Sub dbcintArrendatario_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintArrendatario, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBairro_Click(Area As Integer)
    DropDownDataCombo dbcintBairro, Me, Area
End Sub

Private Sub dbcintBairro_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintBairro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintBairro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintBairro
End Sub

Private Sub cbointContribuinte_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub cbointContribuinte_LostFocus()
    strPreencheContribuinteTabs
End Sub

Private Sub dbcintLogradouro_Click(Area As Integer)
    DropDownDataCombo dbcintLogradouro, Me, Area
End Sub

Private Sub dbcintLogradouro_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintLogradouro
End Sub

Private Sub cbointMunicipio_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub cbointMunicipio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cbointMunicipio
End Sub

Private Sub dbcintOcorrrencia_Click(Area As Integer)
    DropDownDataCombo dbcintOcorrrencia, Me, Area
End Sub

Private Sub dbcintOcorrrencia_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintOcorrrencia_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintOcorrrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrrencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintOcorrrencia
End Sub

Private Sub dbcintTipoC_Click(Area As Integer)
    DropDownDataCombo dbcintTipoC, Me, Area
End Sub

Private Sub dbcintTipoC_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintTipoC_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTipoC, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTipoC
End Sub

Private Sub dbcintTituloC_Click(Area As Integer)
    DropDownDataCombo dbcintTituloC, Me, Area
End Sub

Private Sub dbcintTituloC_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintTituloC_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTituloC, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTituloC
End Sub

Private Sub dbcintUf_Click(Area As Integer)
    DropDownDataCombo dbcintUF, Me, Area
End Sub

Private Sub dbcintUF_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintUf_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUF, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintUF
End Sub

Private Sub cmd_Contribuinte_Click()
    ChamaFormCadastro frmCadContribuinte, dbcintArrendatario
End Sub

Private Sub dbcintArrendatario_GotFocus()
tab_3dPasta.Tab = 1
End Sub

Private Sub Form_Activate()
    If MDIMenu.Tag = "Ouvidoria" Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir
    End If
    gintCodSeguranca = 736
      VirificaGradeListView Me
'=============
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
'=============

End Sub

Private Sub dbcintArrendatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not dbcintArrendatario.MatchedWithList Then
            dbcintArrendatario.BoundText = Trim(dbcintArrendatario.Text)
            If dbcintArrendatario.MatchedWithList Then
                dbcintArrendatario_Click 2
            Else
                dbcintArrendatario.BoundText = ""
            End If
        Else
            dbcintArrendatario_Click 2
        End If
    End If
    CaracterValido KeyAscii, "A", dbcintArrendatario
End Sub

Private Sub cbointContribuinte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not cbointContribuinte.MatchedWithList Then
            cbointContribuinte.BoundText = Trim(cbointContribuinte.Text)
            If cbointContribuinte.MatchedWithList Then
                cbointContribuinte_Click 2
            Else
                cbointContribuinte.BoundText = ""
            End If
        Else
            cbointContribuinte_Click 2
        End If
    End If
    CaracterValido KeyAscii, "A", cbointContribuinte
End Sub

Function strQuerryEmOrdemAlfabetica() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao "
    strSql = strSql & " FROM " & gstrLogradouro
    strSql = strSql & " WHERE Dtmdtexclusao is null "
    strSql = strSql & " ORDER BY strDescricao"
strQuerryEmOrdemAlfabetica = strSql
End Function

Function strQuerryMunicipioEmOrdemAlfabetica() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao "
    strSql = strSql & " FROM " & gstrCidade
    strSql = strSql & " ORDER BY strDescricao"
    strQuerryMunicipioEmOrdemAlfabetica = strSql
End Function

Function strQuerryBairroEmOrdemAlfabetica() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao "
    strSql = strSql & " FROM " & gstrBairro
    strSql = strSql & " ORDER BY strDescricao"
    strQuerryBairroEmOrdemAlfabetica = strSql
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
    If KeyCode = vbKeyF1 Then
       Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
    If MDIMenu.Tag = "Ouvidoria" Then
        cmd_Contribuinte.Enabled = False
    End If

    mblnAlterando = False
    MontaColumnHeaders
   'TAB GERAL
  
    LeDaTabelaParaObj gstrOcorrencia, dbcintOcorrrencia, strQueryTabelaOcorrencia
    LeDaTabelaParaObj gstrContribuinte, cbointContribuinte
    LeDaTabelaParaObj gstrTipoLogradouro, dbcintTipoC, "PKID, strSigla"
    LeDaTabelaParaObj gstrTituloLogradouro, dbcintTituloC
    LeDaTabelaParaObj gstrLogradouro, dbcintLogradouro, strQuerryEmOrdemAlfabetica
    LeDaTabelaParaObj gstrBairro, dbcintBairro, strQuerryBairroEmOrdemAlfabetica
    LeDaTabelaParaObj gstrCidade, cbointMunicipio, strQuerryMunicipioEmOrdemAlfabetica
    LeDaTabelaParaObj gstrUF, dbcintUF
   LeDaTabelaParaObj gstrImobiliarioRural, tdb_ImobiliarioRural, strQueryImobiliarioRural
    VerificaMascaraInscricao
    
'   'TAB Arrendatario
    LeDaTabelaParaObj gstrContribuinte, dbcintArrendatario, "PKId, strNome"

   ''CarregaComboSecoes
    blnButaoNovo = False
    PreencheGRD2
    
    txt_Bairro.Enabled = False
    TrocaCorObjeto txt_Bairro, True
    txt_Cep.Enabled = False
    TrocaCorObjeto txt_Cep, True
    txt_Complemento.Enabled = False
    TrocaCorObjeto txt_Complemento, True
    txt_Distrito.Enabled = False
    TrocaCorObjeto txt_Distrito, True
    txt_Logradouro.Enabled = False
    TrocaCorObjeto txt_Logradouro, True
    txt_Municipio.Enabled = False
    TrocaCorObjeto txt_Municipio, True
    txt_Numero.Enabled = False
    TrocaCorObjeto txt_Numero, True
    txt_UF.Enabled = False
    TrocaCorObjeto txt_UF, True
End Sub

Function strQueryImobiliarioRural() As String

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'        pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT CO.bytNaturezaJuridica, IM.PKId as UM,IM.PKId , IM.strInscricaoAnterior, "
'    strSql = strSql & " CO.strNome, (CONVERT(NVARCHAR,CO.bytNaturezaJuridica) + IM.strCNPJCPF) AS strCNPJCPF"
    strSql = strSql & " CO.strNome, (" & gstrCONVERT(CDT_NVARCHAR, "CO.bytNaturezaJuridica") & strCONCAT & " IM.strCNPJCPF) AS strCNPJCPF"
    strSql = strSql & " FROM "
    strSql = strSql & gstrImobiliarioRural & " IM, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE IM.intContribuinte = CO.PKId "
strQueryImobiliarioRural = strSql
End Function

Private Function strQueryRelatorio() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT IM.PKId as CR, IM.strInscricaoAnterior IA, "
    strSql = strSql & " CO.strNome CO, IM.strCNPJCPF AS CC"
    strSql = strSql & " FROM "
    strSql = strSql & gstrImobiliarioRural & " IM, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE IM.intContribuinte = CO.PKId "
strQueryRelatorio = strSql
End Function


Private Function ContadorDetalhe() As Boolean
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT COUNT(*) as Contador FROM "
    strSql = strSql & gstrCaracteristicaDoImovelRural
    strSql = strSql & " WHERE intCodigoImobiliario = " & Val(PKId_Temporario)
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                If adoResultado!Contador <> 0 Then
                    ContadorDetalhe = True
                    Exit Function
                End If
            End If
        End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MDIMenu.Tag = "Ouvidoria" Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
  End If
 If PKId_Temporario = 0 Then
    Exit Sub
 End If
    If ContadorDetalhe = True Then
        DeletaTemporario
        DoEvents
    Else
        gobjBanco.ExecutaBeginTrans
        gobjBanco.ExecutaCommitTrans
    End If
End Sub

Private Sub DeletaTemporario()
Dim strSql As String
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM " & gstrCaracteristicaDoImovelRural
    strSql = strSql & " WHERE "
    strSql = strSql & " intCodigoImobiliario = " & Val(PKId_Temporario)
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSql
        PKId_Temporario = 0
End Sub

Private Sub lvw_Caracteristica_GotFocus(Index As Integer)
    If Index = 0 Then
        tab_3dPasta.Tab = 2
    ElseIf Index = 1 Then
        tab_3dPasta.Tab = 3
    ElseIf Index = 2 Then
        tab_3dPasta.Tab = 4
    End If
End Sub

'
'Private Sub lbl_ContribuinteGeral_Click()
'    If fra_Contribuinte.Enabled = False Then
'        ExibeMensagem "Selecione um imóvel ou entre com um novo registro."
'    End If
'End Sub

Private Sub lvw_Caracteristica_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    CarregaDetalhes
    If mblnAlterando = True Then
        SelecionaDetalhe
    Else
        SelecionaDetalheApoio
    End If
End Sub

Private Sub lvw_Caracteristica_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_Caracteristica(Index)
End Sub

Private Sub lvw_Detalhe_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_Detalhe(Index)
End Sub

Private Sub lvw_Detalhe_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    For i = 1 To lvw_Detalhe(Index).ListItems.Count
        lvw_Detalhe(Index).ListItems(i).Checked = False
    Next
End Sub

Private Sub lvw_Detalhe_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnAlterando = True Then
        If lvw_Detalhe(Index).ListItems.Count <> 0 Then
            GravaDetalhe
        End If
    Else
        If lvw_Detalhe(Index).ListItems.Count <> 0 Then
            GravaDetalheApoio
        End If
    End If
End Sub

Private Function UpdateGravaDetalhe(Indice As Integer, strOperacao As String)
Dim i      As Integer
Dim strSql As String
    If strOperacao = "SALVAR" Then
    gobjBanco.ExecutaBeginTrans
        strSql = ""
        strSql = strSql & " UPDATE " & gstrCaracteristicaDoImovelRural & " SET intCodigoImobiliario = " & Indice
        strSql = strSql & " WHERE intCodigoImobiliario = " & Val(PKId_Temporario)
        Set gobjBanco = New clsBanco
            gobjBanco.Execute strSql
            gobjBanco.ExecutaCommitTrans
    End If
End Function

Function GravaDetalheApoio()

'******************************************************************************************
' Data: 09/04/2003
' Alteração: - Alteração do nome do atributo intCodigoDetalheDaCaracteristica por
'            intCodigoDetalheDaCaracteristi a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/04/2003
' Alteração: - Alteração do nome do atributo intCodigoUtilizacaoDaTabelaDeValor por
'            intCodigoUtilizacaoDaTabelaDeV a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim i      As Integer
Dim strSql As String
gobjBanco.ExecutaBeginTrans

    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM " & gstrCaracteristicaDoImovelRural
    strSql = strSql & " WHERE "
    'Código do imobiliário
    strSql = strSql & " intCodigoImobiliario = " & Val(PKId_Temporario)
    'Código da utilização
'    strSql = strSql & " AND intCodigoUtilizacaoDaTabelaDeValor = " & (intCaractImobil + 1)
    strSql = strSql & " AND intCodigoUtilizacaoDaTabelaDeV = " & (intCaractImobil + 1)
    'Código da Característica geral
    strSql = strSql & " AND intCodigoCaracteristicaGeral = " & lvw_Caracteristica(intIndiceLVW).SelectedItem.Tag
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql

    For i = 1 To lvw_Detalhe(intIndiceLVW).ListItems.Count
        strSql = ""
        If lvw_Detalhe(intIndiceLVW).ListItems(i).Checked = True Then
        strSql = ""
        strSql = strSql & " INSERT INTO " & gstrCaracteristicaDoImovelRural & "(intCodigoImobiliario,intCodigoCaracteristicaGeral,"
'        strSql = strSql & " intCodigoDetalheDaCaracteristica,intCodigoUtilizacaoDaTabelaDeValor) VALUES ("
        strSql = strSql & " intCodigoDetalheDaCaracteristi,intCodigoUtilizacaoDaTabelaDeV) VALUES ("
        'Código do imobiliário
        strSql = strSql & Val(PKId_Temporario)
        'Código da Característica geral
        strSql = strSql & "," & lvw_Caracteristica(intIndiceLVW).SelectedItem.Tag
        'Código do detalhe da característica
        strSql = strSql & "," & lvw_Detalhe(intIndiceLVW).ListItems(i).Tag
        'Código da utilização
        strSql = strSql & "," & (intCaractImobil + 1)
        strSql = strSql & " )"

            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSql
            gobjBanco.ExecutaCommitTrans
        End If
    Next
End Function

Function SelecionaDetalheApoio()

'******************************************************************************************
' Data: 09/04/2003
' Alteração: - Alteração do nome do atributo intCodigoDetalheDaCaracteristica por
'            intCodigoDetalheDaCaracteristi a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/04/2003
' Alteração: - Alteração do nome do atributo intCodigoUtilizacaoDaTabelaDeValor por
'            intCodigoUtilizacaoDaTabelaDeV a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    strSql = ""
'    strSql = strSql & " SELECT D.intCodigoDetalheDaCaracteristica Detalhe "
    strSql = strSql & " SELECT D.intCodigoDetalheDaCaracteristi Detalhe "
    strSql = strSql & " FROM " & gstrImobiliarioRural & " A," & gstrCaracteristicaGeral & " B,"
    strSql = strSql & gstrUtilizacaoDaTabelaDeValor & " C, " & gstrCaracteristicaDoImovelRural & " D"
    strSql = strSql & " WHERE D.intCodigoImobiliario = " & Val(PKId_Temporario)
    strSql = strSql & " AND D.intCodigoCaracteristicaGeral = B.PKId"
'    strSql = strSql & " AND D.intCodigoUtilizacaoDaTabelaDeValor = C.PKId"
    strSql = strSql & " AND D.intCodigoUtilizacaoDaTabelaDeV = C.PKId"
    strSql = strSql & " AND D.intCodigoCaracteristicaGeral = " & lvw_Caracteristica(intIndiceLVW).SelectedItem.Tag
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                Call MarcaDetalhe(!Detalhe)
                .MoveNext
            Loop
        End With
    End If
End Function

Private Sub lvw_Historico_GotFocus()
    tab_3dPasta.Tab = 6
End Sub

Private Sub lvw_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_Historico
End Sub

Function MarcaDetalhe(intTag As Integer)
Dim i As Integer
    For i = 1 To lvw_Detalhe(intIndiceLVW).ListItems.Count
        If lvw_Detalhe(intIndiceLVW).ListItems(i).Tag = intTag Then
            lvw_Detalhe(intIndiceLVW).ListItems(i).Checked = True
            lvw_Detalhe(intIndiceLVW).ListItems(i).Selected = True
        End If
    Next
End Function

Function SelecionaDetalhe()

'******************************************************************************************
' Data: 09/04/2003
' Alteração: - Alteração do nome do atributo intCodigoDetalheDaCaracteristica por
'            intCodigoDetalheDaCaracteristi a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/04/2003
' Alteração: - Alteração do nome do atributo intCodigoUtilizacaoDaTabelaDeValor por
'            intCodigoUtilizacaoDaTabelaDeV a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    strSql = ""
'    strSql = strSql & " SELECT D.intCodigoDetalheDaCaracteristica Detalhe "
    strSql = strSql & " SELECT D.intCodigoDetalheDaCaracteristi Detalhe "
    strSql = strSql & " FROM " & gstrImobiliarioRural & " A," & gstrCaracteristicaGeral & " B,"
    strSql = strSql & gstrUtilizacaoDaTabelaDeValor & " C, " & gstrCaracteristicaDoImovelRural & " D"
    strSql = strSql & " WHERE D.intCodigoImobiliario = " & Val(txtPKID)
    strSql = strSql & " AND D.intCodigoCaracteristicaGeral = B.PKId"
'    strSql = strSql & " AND D.intCodigoUtilizacaoDaTabelaDeValor = C.PKId"
    strSql = strSql & " AND D.intCodigoUtilizacaoDaTabelaDeV = C.PKId"
    strSql = strSql & " AND D.intCodigoCaracteristicaGeral = " & lvw_Caracteristica(intIndiceLVW).SelectedItem.Tag
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                Call MarcaDetalhe(!Detalhe)
                .MoveNext
            Loop
        End With
    End If
End Function

'''''''''''''$$$$$$$$$$$$$'Query para montar detalhes$$$$$$$$$$$$$$$
Private Function CarregaDetalhes()
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strNomeDoDetalhe Detalhe"
    strSql = strSql & " FROM " & gstrDetalheDaCaracteristica
    'Apenas para característica selecionada no grid de características
    strSql = strSql & " WHERE intCaracteristica = " & lvw_Caracteristica(intIndiceLVW).SelectedItem.Tag
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            lvw_Detalhe(intIndiceLVW).ListItems.Clear
            Do While .EOF = False
                Set objList1 = lvw_Detalhe(intIndiceLVW).ListItems.Add(, , !Detalhe)
                objList1.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Function

Private Sub mskstrInscricaoAnterior_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricaoAnterior
End Sub

Private Sub mskstrInscricaoAnterior_LostFocus()
    If blnDuplicataInscricao(mskstrInscricaoAnterior.Text) = False Then
        Exit Sub
    End If
    strPreencheContribuinteTabs
End Sub

Function blnDuplicataInscricao(strInscricao As String) As Boolean
Dim strSql      As String
Dim strSqlAux   As String
Dim INT_PKIDI   As Integer
    If strInscricao = "" Then
        blnDuplicataInscricao = False
        Exit Function
    End If
    strSql = ""
    strSql = strSql & "SELECT count(*) as Contador FROM " & gstrImobiliarioRural & " WHERE strInscricaoAnterior = '" & strInscricao & "'"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If adoResultado!Contador <= 0 Then
                blnDuplicataInscricao = False
                Exit Function
                Else
                strSqlAux = ""
                strSqlAux = strSqlAux & "SELECT PKId as PP FROM " & gstrImobiliarioRural & " WHERE strInscricaoAnterior = '" & strInscricao & "'"
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSqlAux, 5, adoResultado) Then
                    If Not adoResultado.EOF Then
                        INT_PKIDI = adoResultado!PP
                        blnDuplicataInscricao = True
                    End If
                End If
            End If
        End If
    End If
End Function

'Private Sub mskstrInscricaoAnterior_Validate(Cancel As Boolean)
'If blnButaoNovo Then Exit Sub
'    If mblnAlterando = True Then Exit Sub
'    If Trim(mskstrInscricaoAnterior.ClipText) = "" Then
'        ExibeMensagem "O número da inscrição deve ser digitado."
'        mskstrInscricaoAnterior.SetFocus
'        Cancel = True
'    End If
'End Sub

Function PegaMaxPKId()
Dim strSql As String
        strSql = ""
        strSql = strSql & "SELECT MAX(PKId) as PKId "
        strSql = strSql & " FROM " & gstrImobiliarioRural
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
             intMaxPKId = adoResultado!Pkid
        End If
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSql As String
Dim intAux As Integer
Dim i As Integer

    If UCase(strModoOperacao) = gstrLocalizar Then
        LocalizarImobiliarioRural
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = "NOVO" Then
        If PKId_Temporario <> 0 Then
            strSql = ""
            strSql = strSql & " DELETE "
            strSql = strSql & " FROM " & gstrCaracteristicaDoImovelRural
            strSql = strSql & " WHERE "
            strSql = strSql & " intCodigoImobiliario = " & Val(PKId_Temporario)
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSql
            PKId_Temporario = 0
            strSql = ""
        End If
        strLimpaContribuintesTabs
        txt_PKIdContribuinte = ""
        For i = 1 To lvw_Detalhe(0).ListItems.Count
            lvw_Detalhe(0).ListItems(i).Checked = False
            lvw_Detalhe(0).ListItems(1).Selected = True
        Next
        i = 0
        For i = 1 To lvw_Detalhe(1).ListItems.Count
            lvw_Detalhe(1).ListItems(i).Checked = False
            lvw_Detalhe(1).ListItems(1).Selected = True
        Next
        i = 0
        For i = 1 To lvw_Detalhe(2).ListItems.Count
            lvw_Detalhe(2).ListItems(i).Checked = False
            lvw_Detalhe(2).ListItems(1).Selected = True
        Next
    End If
    If UCase(strModoOperacao) = "SALVAR" And mskstrInscricaoAnterior.ClipText = "" Then
        MsgBox "O número da Inscrição Cadastral tem que ser digitado."
        mskstrInscricaoAnterior.SetFocus
        Exit Sub
    End If
    If UCase(strModoOperacao) = "SALVAR" And dbcintLogradouro.BoundText = "" Then
        MsgBox "O logradouro do Imóvel tem que ser selecionado."
        dbcintLogradouro.SetFocus
        Exit Sub
    End If
    If UCase(strModoOperacao) = "IMPRIMIR" Then
        ImprimeRelatorio rptImobiliarioRural, strQueryRelatorio
        Exit Sub
    End If
intAux = 0
    If mblnAlterando Then
        intAux = Val(tdb_ImobiliarioRural.Columns("UM").Value)
    End If
        If ToolBarGeral(strModoOperacao, gstrImobiliarioRural, mblnAlterando, tdb_ImobiliarioRural, Me, mobjAux, strQueryImobiliarioRural) Then
            strLimpaContribuintesTabs
            lvw_Caracteristica(0).ListItems.Clear
            lvw_Caracteristica(1).ListItems.Clear
            lvw_Caracteristica(2).ListItems.Clear
            If intAux > 0 Then
                If GravaHistoricos((intAux), strModoOperacao) Or DeletaHistoricos((intAux), strModoOperacao) Then
                End If
                If GravaValores2((intAux), strModoOperacao) Or DeletaValores2((intAux), strModoOperacao) Then
                End If
            Else
                PegaMaxPKId
                DoEvents
                UpdateGravaDetalhe intMaxPKId, strModoOperacao
                If GravaHistoricos((intMaxPKId), strModoOperacao) Or DeletaHistoricos((intMaxPKId), strModoOperacao) Then
                End If
                If GravaValores2((intMaxPKId), strModoOperacao) Or DeletaValores2((intMaxPKId), strModoOperacao) Then
                End If
            End If
            intMaxPKId = 0
            PKId_Temporario = 0
            tab_3dPasta.Tab = 0
            txt_PKIdContribuinte = ""
            mblnAlterando = False
            mblnPrimeiraVez = False
        End If
        If UCase(strModoOperacao) = "NOVO" And mblnAlterando = False Then
            cbointContribuinte.Locked = False
        End If
        If UCase(strModoOperacao) = "EXCLUIR" Then
            cbointContribuinte.Locked = False
            mblnAlterando = False
        End If
        If UCase(strModoOperacao) = "NOVO" Then
            lvw_Historico.ListItems.Clear
            txt_Historico = ""
            txt_ValorTotalEstimado = ""
            PreencheGRD2
            mblnAlterando = False
            tab_3dPasta.Tab = 0
        End If
End Sub

Sub MontaColumnHeaders()
     With lvw_Historico
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Historico", 7000
    End With
End Sub

 'TAB GERAL

Private Sub cbointContribuinte_Click(Area As Integer)

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela variável
'            gstrISNULL.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    If Area = 2 Then
        If cbointContribuinte.Locked Then Exit Sub
        If cbointContribuinte.BoundText <> "" Then
            txt_PKIdContribuinte = cbointContribuinte.BoundText

            strSql = ""
'            strSql = strSql & "Select PKId, ISNULL(strCNPJCPF, 0) AS strCNPJCPF From " & gstrContribuinte & " "
            strSql = strSql & "Select PKId, " & gstrISNULL("strCNPJCPF", "0") & " AS strCNPJCPF From " & gstrContribuinte & " "
            strSql = strSql & "Where PKId = " & cbointContribuinte.BoundText

            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                txtstrcnpjcpf = adoResultado!StrCnpjCpf
            End If
        End If
        
    End If
End Sub

Private Function strQueryTabelaOcorrencia() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "Select PKId, strDescricao "
    strSql = strSql & "From " & gstrOcorrencia & " "
    strSql = strSql & "Where intUtilizacaoDaOcorrencia = 6 "
    strQueryTabelaOcorrencia = strSql
End Function

Sub VerificaMascaraInscricao()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
    
    strMascara = ""
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & TYP_IMOBILIARIA
    strSql = strSql & "Order By intSequencia"

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    mskstrInscricaoAnterior.Mask = strMascara
End Sub

Private Sub tdb_ImobiliarioRural_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_ImobiliarioRural_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_ImobiliarioRural
End Sub

Private Sub tdb_ImobiliarioRural_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 4 Then
        If Mid(Value, 1, 1) Then
            Value = Format(Mid(Value, 2), "@@.@@@.@@@/@@@@-@@")
        Else
            Value = Format(Mid(Value, 2), "@@@.@@@.@@@-@@")
        End If
    End If
End Sub

Private Sub tdb_ImobiliarioRural_HeadClick(ByVal ColIndex As Integer)
gOrdenaGrid tdb_ImobiliarioRural, ColIndex
End Sub

Private Sub tdb_ImobiliarioRural_KeyPress(KeyAscii As Integer)
    If tdb_ImobiliarioRural.Col = 1 Or tdb_ImobiliarioRural.Col = 2 Then
        CaracterValido KeyAscii, "N", tdb_ImobiliarioRural
    Else
        CaracterValido KeyAscii, "A", tdb_ImobiliarioRural
    End If
End Sub

Private Sub tdb_ImobiliarioRural_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_ImobiliarioRural
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                cbointContribuinte.Locked = True
                txtPKID.Text = tdb_ImobiliarioRural.Columns("UM").Value
                LeDaTabelaParaObj gstrImobiliarioRural, Me
                'LeDaTabelaParaObj gstrHistoricoImobiliarioRural, lvw_Historico, CarregaHistoricos(lvw_Lista.SelectedItem.Tag)
                CarregaHistoricos (txtPKID.Text)
                PreencheGRD2
                TotalEstimado
                strPreencheContribuinteTabs
                txt_PKIdContribuinte = cbointContribuinte.BoundText
                lvw_Detalhe(0).ListItems.Clear
                lvw_Detalhe(1).ListItems.Clear
                lvw_Detalhe(2).ListItems.Clear
            End If
        End If
    End With
End Sub

Private Sub tdb_Producao_GotFocus()
    tab_3dPasta.Tab = 5
End Sub

Private Sub tlb_Historico_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim alterando As Boolean

    'If txtPKId.Text = "" Then Exit Sub

    If lvw_Historico.ListItems.Count <> 0 Then
        If lvw_Historico.SelectedItem.Selected Then
            alterando = True
            Else
            alterando = False
        End If
    End If
    Select Case UCase(Button.Key)
        Case gstrSalvar
            If txt_Historico = "" Then Exit Sub
            If alterando Then
                lvw_Historico.SelectedItem.Text = txt_Historico
            Else
                lvw_Historico.ListItems.Add , , txt_Historico
            End If
            txt_Historico = ""
        Case gstrNovo
            txt_Historico = ""
            txt_Historico.SetFocus
        Case gstrDeletar
            If txt_Historico = "" Then Exit Sub

            If lvw_Historico.SelectedItem.Selected Then
                lvw_Historico.ListItems.Remove (lvw_Historico.SelectedItem.Index)
                txt_Historico = ""
            End If

    End Select

    If lvw_Historico.ListItems.Count <> 0 Then
        lvw_Historico.SelectedItem.Selected = False
    End If
End Sub

Private Sub txt_Historico_GotFocus()
    tab_3dPasta.Tab = 6
End Sub

Private Sub txt_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Historico, True
    If KeyAscii = vbKeyReturn Then
        mskstrInscricaoAnterior.SetFocus
    End If
End Sub

Private Sub txt_ValorTotalEstimado_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_ValorTotalEstimado
End Sub

Private Sub txtdblValorITBI_LostFocus()
    txtdblValorITBI = gvntConvVrDoSql(txtdblValorITBI)
End Sub

Private Sub txtstrAreaPropriedade_GotFocus()
    MarcaCampo txtstrAreaPropriedade
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrAreaPropriedade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrAreaPropriedade
End Sub

Private Sub txtstrAreaPropriedade_LostFocus()
    txtstrAreaPropriedade = gvntConvVrDoSql(txtstrAreaPropriedade)
End Sub

Private Sub txtstrAreaConstruida_GotFocus()
    MarcaCampo txtstrAreaConstruida
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrAreaConstruida_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtstrAreaConstruida
End Sub

Private Sub txtstrAreaConstruida_LostFocus()
    txtstrAreaConstruida = gvntConvVrDoSql(txtstrAreaConstruida)
End Sub

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub txtintNumero_GotFocus()
    MarcaCampo txtintNumero
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCEP
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCEP
End Sub

Private Sub txtintCEP_LostFocus()
    If txtintCEP.Text = "" Then
        Exit Sub
    End If
        txtintCEP = gstrCEPFormatado(txtintCEP)
    If gblnCepValido(txtintCEP, dbcintLogradouro) = False Then
        MsgBox "CEP inválido para o logradouro cadastrado "
    End If
End Sub

Private Sub mskstrInscricaoAnterior_GotFocus()
    blnButaoNovo = False
    MarcaCampo mskstrInscricaoAnterior
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumero
End Sub

'TAB PROMISSARIO

Function strEncheCamposArrendatario()

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela variável
'            gstrISNULL.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    If dbcintArrendatario.MatchedWithList Then
        strSql = ""
'        strSql = strSql & " SELECT bytNaturezaJuridica, ISNULL(strCNPJCPF, 0) as strCNPJCPFP "
        strSql = strSql & " SELECT bytNaturezaJuridica, " & gstrISNULL("strCNPJCPF", "0") & " as strCNPJCPFP "
        strSql = strSql & " FROM " & gstrContribuinte & " "
        strSql = strSql & " WHERE PKId = " & dbcintArrendatario.BoundText
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    optbytNaturezaJuridica(!bytNaturezaJuridica) = True
                    txtstrCNPJCPFP = !strCNPJCPFP
                    .MoveNext
                Loop
            End With
        End If
    End If
End Function

Private Sub dbcintArrendatario_Click(Area As Integer)
    DropDownDataCombo dbcintArrendatario, Me, Area
    If Area = 2 Then
       strEncheCamposArrendatario
       MostraDadosContribuinte (dbcintArrendatario.BoundText)
    End If
End Sub

Private Function MostraDadosContribuinte(intBound As Integer) As Boolean
Dim strSql As String
On Error Resume Next
    strSql = ""
    strSql = strSql & "SELECT CO.strBairroC, CO.strLogradouroC, CO.intNumeroC,"
    strSql = strSql & " CO.strComplementoC , CO.intCEPC, CO.strDistritoC, "
    strSql = strSql & " CD.strDescricao, UF.strSigla "
    strSql = strSql & " FROM "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrCidade & " CD, "
    strSql = strSql & gstrUF & " UF "
    strSql = strSql & "WHERE intMunicipioC = CD.PKId "
    strSql = strSql & " AND intUFC = UF.PKId "
    strSql = strSql & " AND CO.PKId = " & intBound
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                txt_Bairro = gstrVerificaCampoNulo(!strBairroC)
                txt_Cep = gstrVerificaCampoNulo(!intCepC)
                txt_Complemento = gstrVerificaCampoNulo(!strComplementoC)
                txt_Distrito = gstrVerificaCampoNulo(!strDistritoC)
                txt_Logradouro = gstrVerificaCampoNulo(!strLogradouroC)
                txt_Municipio = gstrVerificaCampoNulo(!strDescricao)
                txt_Numero = gstrVerificaCampoNulo(!intNumeroC)
                txt_UF = gstrVerificaCampoNulo(!strSigla)
                MostraDadosContribuinte = True
                .MoveNext
            Loop
        End With
    End If
End Function

Private Sub dbcintArrendatario_Validate(Cancel As Boolean)
    strEncheCamposArrendatario
End Sub

'TAB ÁREAS E HISTÓRICOS
Private Sub tab_3dPasta_Click(PreviousTab As Integer)
Dim intCodImobiliario As Integer

'If txtPKId <> "" Then

    Select Case tab_3dPasta.Tab
        Case 0
        Case 1
        If dbcintArrendatario.BoundText <> "" Then
            MostraDadosContribuinte (dbcintArrendatario.BoundText)
        End If
        Case 2
            '7 = Rural Geral
            If mblnAlterando = False And PKId_Temporario = 0 Then
               PKId_Temporario = Timer * 100
            End If
            intCaractImobil = 6
            intIndiceLVW = 0
            CarregaCaracteristica
            'CarregaDetalheDaCaracteristica
        Case 3
            '8 = Rural Terreno
            If mblnAlterando = False And PKId_Temporario = 0 Then
               PKId_Temporario = Timer * 100
            End If
            intCaractImobil = 7
            intIndiceLVW = 1
            CarregaCaracteristica
        Case 4
            '9 = Rural Construção
            If mblnAlterando = False And PKId_Temporario = 0 Then
               PKId_Temporario = Timer * 100
            End If
            intCaractImobil = 8
            intIndiceLVW = 2
            CarregaCaracteristica
        Case 5
            TotalEstimado
        Case 6

        Case 7
'                txtdblValorTerreno = gstrConvVrParaSql(txtdblValorTerreno)
'                txtdblValorEdificacao = gstrConvVrParaSql(txtdblValorEdificacao)
'                txtdblValorImovel = Val(txtdblValorEdificacao) + Val(txtdblValorTerreno)
'                txtdblValorImovel = gvntConvVrDoSql(txtdblValorImovel)
    End Select
'End If
End Sub

Function GravaDetalhe()

'******************************************************************************************
' Data: 09/04/2003
' Alteração: - Alteração do nome do atributo intCodigoDetalheDaCaracteristica por
'            intCodigoDetalheDaCaracteristi a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/04/2003
' Alteração: - Alteração do nome do atributo intCodigoUtilizacaoDaTabelaDeValor por
'            intCodigoUtilizacaoDaTabelaDeV a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim i      As Integer
    Dim strSql As String
    
    gobjBanco.ExecutaBeginTrans
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM " & gstrCaracteristicaDoImovelRural
    strSql = strSql & " WHERE "
    'Código do imobiliário
    strSql = strSql & " intCodigoImobiliario = " & Trim(txtPKID.Text)
    'Código da utilização
'    strSql = strSql & " AND intCodigoUtilizacaoDaTabelaDeValor = " & (intCaractImobil + 1)
    strSql = strSql & " AND intCodigoUtilizacaoDaTabelaDeV = " & (intCaractImobil + 1)
    'Código da Característica geral
    strSql = strSql & " AND intCodigoCaracteristicaGeral = " & lvw_Caracteristica(intIndiceLVW).SelectedItem.Tag
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql

    For i = 1 To lvw_Detalhe(intIndiceLVW).ListItems.Count
        strSql = ""
        If lvw_Detalhe(intIndiceLVW).ListItems(i).Checked = True Then
        strSql = ""
        strSql = strSql & " INSERT INTO " & gstrCaracteristicaDoImovelRural & "(intCodigoImobiliario,intCodigoCaracteristicaGeral,"
'        strSql = strSql & " intCodigoDetalheDaCaracteristica,intCodigoUtilizacaoDaTabelaDeValor) VALUES ("
        strSql = strSql & " intCodigoDetalheDaCaracteristi,intCodigoUtilizacaoDaTabelaDeV) VALUES ("
        'Código do imobiliário
        strSql = strSql & Trim(txtPKID)
        'Código da Característica geral
        strSql = strSql & "," & lvw_Caracteristica(intIndiceLVW).SelectedItem.Tag
        'Código do detalhe da característica
        strSql = strSql & "," & lvw_Detalhe(intIndiceLVW).ListItems(i).Tag
        'Código da utilização
        strSql = strSql & "," & (intCaractImobil + 1)
        strSql = strSql & ")"

            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSql
            gobjBanco.ExecutaCommitTrans
        End If
    Next
End Function

Private Sub CarregaCaracteristica()
    Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strNomeDaCaracteristica Caracteristica"
    strSql = strSql & " FROM " & gstrCaracteristicaGeral
    '1 = Rural Geral
    '2 = Rural Terreno
    '3 = Rural Construção
    strSql = strSql & " WHERE intUtilizacaoDaCaracteristica = " & (intCaractImobil + 1)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            lvw_Caracteristica(intIndiceLVW).ListItems.Clear
            Do While .EOF = False
                Set objList1 = lvw_Caracteristica(intIndiceLVW).ListItems.Add(, , !Caracteristica)
                objList1.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
    lvw_Caracteristica(intIndiceLVW).Refresh
    lvw_Caracteristica(intIndiceLVW).SetFocus
    If lvw_Caracteristica(intIndiceLVW).ListItems.Count > 0 Then
        lvw_Caracteristica(intIndiceLVW).ListItems(1).Selected = True
    End If
    SendKeys " "
    DoEvents
End Sub

Function strQuerryGrid() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT VL.intX, VL.strY FROM "
    strSql = strSql & gstrValorCompoRec & " VL, "
    strSql = strSql & gstrComposicaoDaReceita & " CP "
    strSql = strSql & " WHERE CP.intCodigo = VL.intComposicao "
    strSql = strSql & " AND CP.PKId = " & tdb_ImobiliarioRural.Columns("UM").Value
strQuerryGrid = strSql
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ShiftDown, AltDown, CtrlDown
    Select Case KeyCode
        Case vbKeyEscape
'            If Not IsNull(tdd_Area.SelectedItem) Then
'                grd_Area.SelStart = Len(grd_Area.Text)
'            End If
            SendKeys "{RIGHT}"
            Exit Sub
    End Select
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
End Sub

Private Function blnDadosOk() As Boolean
    blnDadosOk = True
End Function

Private Sub lvw_Historico_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_Historico
        txt_Historico = .SelectedItem.Text
    End With
End Sub

Private Sub tbl_Historico_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim alterando As Boolean
'    If lvw_Historico.SelectedItem = True Then
'        alterando = True
'    End If
'
'    Select Case UCase(Button.Key)
'        Case gstrSalvar
'            If txt_Historico = "" Then Exit Sub
'            If alterando Then
'                lvw_Historico.SelectedItem.Text = txt_Historico
'            Else
'                lvw_Historico.ListItems.Add , , txt_Historico
'            End If
'
'        Case gstrNovo
'
'        Case gstrDeletar
'            If txt_Historico = "" Then Exit Sub
'            If alterando Then
'                lvw_Historico.SelectedItem.Text = txt_Historico
'            Else
'                lvw_Historico.ListItems.Remove (Tag)
'            End If
'
'    End Select
'    txt_Historico = ""
End Sub

Function GravaHistoricos(intCodImobiliario As Integer, strOperacao As String) As Boolean
    Dim strSql As String
    Dim intI   As Integer

    On Error GoTo err_GravaHistoricos
'If strOperacao = "Novo" Then
'    lvw_Historico.ListItems.Clear
'    txt_Historico = ""
'End If


If UCase(strOperacao) = "SALVAR" Then

    strSql = ""
    strSql = strSql & "Delete From " & gstrHistoricoImobiliarioRural & " "
    strSql = strSql & "Where intImobiliario = " & intCodImobiliario

    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql

    GravaHistoricos = True

    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans


    With lvw_Historico
        For intI = 1 To .ListItems.Count
            strSql = ""
            strSql = strSql & "Insert Into " & gstrHistoricoImobiliarioRural & " "
            strSql = strSql & "(intImobiliario, strDescricao "
            strSql = strSql & ") Values ("
            strSql = strSql & intCodImobiliario & ",'"
            strSql = strSql & .ListItems(intI).Text & "' "
            strSql = strSql & ")"
            Set gobjBanco = New clsBanco
            If Not gobjBanco.Execute(strSql) Then
                gobjBanco.ExecutaRollbackTrans
            End If
        Next
    End With
    lvw_Historico.ListItems.Clear
    txt_Historico = ""


    gobjBanco.ExecutaCommitTrans

Exit Function
err_GravaHistoricos:
    gobjBanco.ExecutaRollbackTrans
    GravaHistoricos = False
End If
End Function

Private Function DeletaHistoricos(intCodImobiliario As Integer, Optional strOperacao As String) As Boolean
If UCase(strOperacao) = "DELETAR" Then
    Dim strSql As String
    strSql = ""
    strSql = strSql & "Delete From " & gstrHistoricoImobiliarioRural & " "
    strSql = strSql & "Where intImobiliario = " & intCodImobiliario

    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
End If
End Function

Private Function CarregaHistoricos(intCodImobiliario As Integer)
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset

    lvw_Historico.ListItems.Clear
    txt_Historico = ""

    strSql = ""
    strSql = strSql & "Select HI.strDescricao Historico "
    strSql = strSql & "From " & gstrHistoricoImobiliarioRural & " HI "
    strSql = strSql & "Where HI.intImobiliario = " & tdb_ImobiliarioRural.Columns("UM").Value
    'CarregaHistoricos =
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set oList = lvw_Historico.ListItems.Add(, , Trim(!Historico))
                .MoveNext
            Loop
        End With
    End If
    If lvw_Historico.ListItems.Count <> 0 Then
        lvw_Historico.SelectedItem.Selected = False
    End If
End Function

Private Sub txtdblValorEdificacao_LostFocus()
dblEdificacao = 0
dblTerreno = 0
    If txtdblValorEdificacao = "Null" Then
       txtdblValorEdificacao = 0
    End If
    If txtdblValorTerreno = "Null" Then
       txtdblValorTerreno = 0
    End If
    If txtdblValorEdificacao = "" Then
        dblEdificacao = 0
        Else
        dblEdificacao = CDbl(txtdblValorEdificacao)
    End If
    If txtdblValorTerreno = "" Then
        dblTerreno = 0
        Else
        dblTerreno = CDbl(txtdblValorTerreno)
    End If
    txtdblValorImovel = (dblEdificacao) + (dblTerreno)
    txtdblValorEdificacao = gvntConvVrDoSql(txtdblValorEdificacao)
End Sub

Private Sub txtdblValorImovel_Change()
    If txtdblValorImovel = "" Then
        Exit Sub
    End If
    txtdblValorImovel = gvntConvVrDoSql(txtdblValorImovel)
End Sub


Private Sub txtdblValorTerreno_LostFocus()
dblEdificacao = 0
dblTerreno = 0
    If txtdblValorEdificacao = "Null" Then
       txtdblValorEdificacao = 0
    End If
    If txtdblValorTerreno = "Null" Then
       txtdblValorTerreno = 0
    End If
    If txtdblValorEdificacao = "" Then
        dblEdificacao = 0
        Else
        dblEdificacao = CDbl(txtdblValorEdificacao)
    End If
    If txtdblValorTerreno = "" Then
        dblTerreno = 0
        Else
        dblTerreno = CDbl(txtdblValorTerreno)
    End If
    txtdblValorImovel = (dblEdificacao) + (dblTerreno)
    txtdblValorTerreno = gvntConvVrDoSql(txtdblValorTerreno)
End Sub

Private Sub txtdblValorTerreno_GotFocus()
    MarcaCampo txtdblValorTerreno
    tab_3dPasta.Tab = 6
End Sub

Private Sub txtdblValorTerreno_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValorTerreno
End Sub

Private Sub txtdblValorEdificacao_GotFocus()
    MarcaCampo txtdblValorEdificacao
    tab_3dPasta.Tab = 6
End Sub

Private Sub txtdblValorEdificacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValorEdificacao
End Sub

Private Sub txtdblValorITBI_GotFocus()
    MarcaCampo txtdblValorITBI
    tab_3dPasta.Tab = 6
End Sub

Private Sub txtdblValorITBI_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValorITBI
End Sub

Private Sub txtdblValorImovel_LostFocus()
    txtdblValorImovel = gvntConvVrDoSql(txtdblValorImovel)
End Sub

'    strSql = ""
'    strSql = strSql & "SELECT SL.PKId as PKID2, ISNULL(SL.strInscricaoCadastral, '') + ' - ' + RTRIM(LTRIM(ISNULL(TL.strSigla, '') + ' ' + ISNULL(U.strDescricao,'') + ' ' + L.strDescricao)) AS Logradouro "
'    strSql = strSql & "FROM " & gstrSecaoLogradouro & " SL, " & "(" & gstrLogradouro & " L "
'    strSql = strSql & "LEFT JOIN  " & gstrTituloLogradouro & " U "
'    strSql = strSql & "ON L.intTituloLogradouro = U.PKId) "
'    strSql = strSql & "LEFT JOIN " & gstrTipoLogradouro & " TL "
'    strSql = strSql & "ON L.intTipoLogradouro = TL.PKId "
'    strSql = strSql & "WHERE SL.intLogradouro = L.PKId "
'    strSql = strSql & "ORDER BY Logradouro"

Sub strLimpaContribuintesTabs()
txt_Inscricao = ""
txt_Proprietario = ""
'txt_Inscricao1 = ""
'txt_Proprietario1 = ""
txt_Inscricao2 = ""
txt_Proprietario2 = ""
txt_Inscricao3 = ""
txt_Proprietario3 = ""
txt_Inscricao4 = ""
txt_Proprietario4 = ""
txt_Inscricao5 = ""
txt_Proprietario5 = ""
txt_Inscricao6 = ""
txt_Proprietario6 = ""

txt_Bairro = ""
txt_Cep = ""
txt_Complemento = ""
txt_Distrito = ""
txt_Logradouro = ""
txt_Municipio = ""
txt_Numero = ""
txt_UF = ""
End Sub

Sub strPreencheContribuinteTabs()
    txt_Inscricao = gstrVerificaCampoNulo(mskstrInscricaoAnterior.ClipText)
    txt_Proprietario = gstrVerificaCampoNulo(cbointContribuinte.Text)
    'txt_Inscricao1 = gstrVerificaCampoNulo(mskstrInscricaoAnterior.ClipText)
    'txt_Proprietario1 = gstrVerificaCampoNulo(cbointContribuinte.Text)
    txt_Inscricao2 = gstrVerificaCampoNulo(mskstrInscricaoAnterior.ClipText)
    txt_Proprietario2 = gstrVerificaCampoNulo(cbointContribuinte.Text)
    txt_Inscricao3 = gstrVerificaCampoNulo(mskstrInscricaoAnterior.ClipText)
    txt_Proprietario3 = gstrVerificaCampoNulo(cbointContribuinte.Text)
    txt_Inscricao4 = gstrVerificaCampoNulo(mskstrInscricaoAnterior.ClipText)
    txt_Proprietario4 = gstrVerificaCampoNulo(cbointContribuinte.Text)
    txt_Inscricao5 = gstrVerificaCampoNulo(mskstrInscricaoAnterior.ClipText)
    txt_Proprietario5 = gstrVerificaCampoNulo(cbointContribuinte.Text)
    txt_Inscricao6 = gstrVerificaCampoNulo(mskstrInscricaoAnterior.ClipText)
    txt_Proprietario6 = gstrVerificaCampoNulo(cbointContribuinte.Text)
End Sub

'umw2345

Function PreencheGRD2()
Dim strSql As String
    
LimpaGrid2
'GRID PRODUCOES
        strSql = ""
        strSql = strSql & "SELECT strCaracteristica, strDetalhe, intQtd, "
        strSql = strSql & " strUnidadeMedida, dblValor, dblValorTotal "
        strSql = strSql & " FROM " & gstrProducaoImobiliarioRural
    If mblnAlterando = True Then
        strSql = strSql & " WHERE intImobiliario = " & tdb_ImobiliarioRural.Columns("UM").Value
    End If
        Set gobjBanco = New clsBanco
        gobjBanco.CriaADO strSql, 5, adoRec
        MontaArray2
        
'DROP CARACTERISTICA

        strSql = ""
        strSql = strSql & "SELECT PKId, strNomeDaCaracteristica "
        strSql = strSql & "FROM " & gstrCaracteristicaGeral
        strSql = strSql & " WHERE intUtilizacaoDaCaracteristica = 10 "

        Set gobjBanco = New clsBanco
        gobjBanco.CriaADO strSql, 5, adoTdb
    If Not adoTdb.EOF Then
        Y.ReDim 0, adoTdb.RecordCount - 1, 0, 1
        Dim varAux1  As Variant
        Dim varAux11 As Variant
        Do While Not adoTdb.EOF

            varAux1 = adoTdb!Pkid
            varAux11 = adoTdb!strNomeDaCaracteristica
            Y(adoTdb.AbsolutePosition - 1, 0) = varAux1
            Y(adoTdb.AbsolutePosition - 1, 1) = varAux11

            adoTdb.MoveNext
        Loop
    End If
        Set tdd_Caracteristica.Array = Y
        tdd_Caracteristica.Rebind
        tdd_Caracteristica.Refresh

' DROP UNIDADE
        
        strSql = ""
        strSql = strSql & "SELECT PKId, strDescricao "
        strSql = strSql & " FROM " & gstrUnidadeMedida
        strSql = strSql & " ORDER BY strDescricao"

        Set gobjBanco = New clsBanco
        gobjBanco.CriaADO strSql, 5, adoTdb
    If Not adoTdb.EOF Then
        A.ReDim 0, adoTdb.RecordCount - 1, 0, 1
        Dim varAux3  As Variant
        Dim varAux31 As Variant
        Do While Not adoTdb.EOF

            varAux3 = adoTdb!Pkid
            varAux31 = adoTdb!strDescricao
            A(adoTdb.AbsolutePosition - 1, 0) = varAux3
            A(adoTdb.AbsolutePosition - 1, 1) = varAux31
            
            adoTdb.MoveNext
        Loop
    End If
        Set tdd_UnidadeMedida.Array = A
        tdd_UnidadeMedida.Rebind
        tdd_UnidadeMedida.Refresh





End Function

Private Sub tdb_Producao_KeyPress(KeyAscii As Integer)
    
    Select Case tdb_Producao.Col
    Case 0
    Case 1
    Case 2
        CaracterValido KeyAscii, "N", tdb_Producao.Columns(2)
    Case 3
    Case 4
        CaracterValido KeyAscii, "V", tdb_Producao.Columns(4)
    End Select
End Sub


Private Sub MontaArray2()
    Dim varAux As Variant

    Set X = New XArrayDB
    X.Clear
    With adoRec
        If Not .EOF And mblnAlterando Then
            X.ReDim 0, .RecordCount - 1, 0, 4
            txt_ValorTotalEstimado = gvntConvVrDoSql(!dblValorTotal)
            Do While Not .EOF
                varAux = .Fields(0)
                X(.AbsolutePosition - 1, 0) = varAux
                varAux = .Fields(1)
                X(.AbsolutePosition - 1, 1) = varAux
                varAux = .Fields(2)
                X(.AbsolutePosition - 1, 2) = varAux
                varAux = .Fields(3)
                X(.AbsolutePosition - 1, 3) = varAux
                varAux = .Fields(4)
                X(.AbsolutePosition - 1, 4) = varAux
                .MoveNext
            Loop
        Else
            X.ReDim 0, 0, 0, 4
            X(0, 0) = ""
            X(0, 1) = ""
            X(0, 2) = ""
            X(0, 3) = ""
            X(0, 4) = ""
        End If
    End With

    Set tdb_Producao.Array = X
    tdb_Producao.Rebind
    tdb_Producao.Refresh
End Sub




Private Sub LimpaGrid2()
    Set X = New XArrayDB
    Set Y = New XArrayDB
    Set Z = New XArrayDB
    Set A = New XArrayDB

    X.Clear
    Y.Clear
    Z.Clear
    A.Clear

    Set tdb_Producao.Array = X
    tdb_Producao.Rebind
    tdb_Producao.Refresh

    Set tdd_Caracteristica.Array = Y
    tdd_Caracteristica.Rebind
    tdd_Caracteristica.Refresh

    Set tdd_Detalhe.Array = Z
    tdd_Detalhe.Rebind
    tdd_Detalhe.Refresh

    Set tdd_UnidadeMedida.Array = A
    tdd_UnidadeMedida.Rebind
    tdd_UnidadeMedida.Refresh


End Sub

Private Sub tdd_Caracteristica_DropDownClose()
    Dim intRow As Integer
    Dim PPKKid As Integer
    PPKKid = 0
    On Error GoTo Err_Handle
    If Not IsNull(tdd_Caracteristica.SelectedItem) Or Not IsEmpty(tdd_Caracteristica.SelectedItem) Then
        tdb_Producao.Columns(0) = tdd_Caracteristica.Columns(1)
        PPKKid = Val(tdd_Caracteristica.Columns(0))
    Else
        tdb_Producao.Columns(0) = ""
    End If
        
        If PPKKid <> 0 Then
            strSql = ""
            strSql = strSql & "SELECT PKId, strNomeDoDetalhe "
            strSql = strSql & "FROM " & gstrDetalheDaCaracteristica
            strSql = strSql & " WHERE intCaracteristica = " & PPKKid
            strSql = strSql & " ORDER BY strNomeDoDetalhe "
    
            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strSql, 5, adoTdb
            If Not adoTdb.EOF Then
                Z.ReDim 0, adoTdb.RecordCount - 1, 0, 1
                Dim varAux2  As Variant
                Dim varAux21 As Variant
                Do While Not adoTdb.EOF
        
                    varAux2 = adoTdb!Pkid
                    varAux21 = adoTdb!strNomeDoDetalhe
                    Z(adoTdb.AbsolutePosition - 1, 0) = varAux2
                    Z(adoTdb.AbsolutePosition - 1, 1) = varAux21
        
                    adoTdb.MoveNext
                Loop
            End If
            Set tdd_Detalhe.Array = Z
            tdd_Detalhe.Rebind
            tdd_Detalhe.Refresh
        End If
    Exit Sub
Err_Handle:
End Sub

Private Sub tdd_Detalhe_DropDownClose()
    Dim intRow As Integer
    On Error GoTo Err_Handle
    If Not IsNull(tdd_Detalhe.SelectedItem) Or Not IsEmpty(tdd_Detalhe.SelectedItem) Then
        tdb_Producao.Columns(1) = tdd_Detalhe.Columns(1)
    Else
        tdb_Producao.Columns(1) = ""
    End If

    Exit Sub
Err_Handle:
End Sub

Private Sub tdd_UnidadeMedida_DropDownClose()
    Dim intRow As Integer
    On Error GoTo Err_Handle
    If Not IsNull(tdd_UnidadeMedida.SelectedItem) Or Not IsEmpty(tdd_UnidadeMedida.SelectedItem) Then
        tdb_Producao.Columns(3) = tdd_UnidadeMedida.Columns(1)
    Else
        tdb_Producao.Columns(3) = ""
    End If

    Exit Sub
Err_Handle:
End Sub


Private Function DeletaValores2(intCodImobiliario As Integer, strOperacao As String) As Boolean
    Dim strSql As String
If strOperacao = "DELETAR" Then
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrProducaoImobiliarioRural & " "
        strSql = strSql & "WHERE  intImobiliario = " & intCodImobiliario
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSql

        LimpaGrid2
        txt_ValorTotalEstimado = ""

End If
End Function

Private Function GravaValores2(intCodImobiliario As Integer, strOperacao As String) As Boolean
    Dim strSql As String
    Dim strMsg As String
    Dim i      As Integer

On Error GoTo err_GravaValores2
If strOperacao = "SALVAR" Then

    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans

    strSql = ""
    strSql = strSql & "DELETE FROM " & gstrProducaoImobiliarioRural & " "
    strSql = strSql & "WHERE  intImobiliario = " & intCodImobiliario
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql

    tdb_Producao.MoveFirst

    For i = 0 To X.Count(1) - 1
        strSql = ""
        strSql = strSql & "INSERT INTO " & gstrProducaoImobiliarioRural & " "
        strSql = strSql & "(intImobiliario, strCaracteristica, strDetalhe, "
        strSql = strSql & " intQtd, strUnidadeMedida, dblValor, dblValorTotal "
        strSql = strSql & ") Values ("
        strSql = strSql & intCodImobiliario & ", "
        strSql = strSql & "'" & X(i, 0) & "', "
        strSql = strSql & "'" & X(i, 1) & "', "
        strSql = strSql & Val(X(i, 2)) & ", "
        strSql = strSql & "'" & X(i, 3) & "', "
        strSql = strSql & gstrConvVrParaSql(X(i, 4)) & ", "
        strSql = strSql & gstrConvVrParaSql(txt_ValorTotalEstimado) & " "
        strSql = strSql & ")"

        If Not gobjBanco.Execute(strSql, False) Then
            gobjBanco.ExecutaRollbackTrans
        End If
    Next i
End If
    gobjBanco.ExecutaCommitTrans
    LimpaGrid2
    txt_ValorTotalEstimado = ""

Exit Function
err_GravaValores2:
    gobjBanco.ExecutaRollbackTrans
End Function

Private Sub tdb_Producao_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 4 Then
        TotalEstimado
    End If
End Sub

Private Sub tdb_Producao_AfterDelete()
    TotalEstimado
End Sub

Sub TotalEstimado()
    Dim i As Integer
    Dim vntTotal
    On Error GoTo Err_Handle
    vntTotal = 0
    tdb_Producao.Update
    For i = 0 To X.Count(1) - 1
        vntTotal = vntTotal + CDbl(X(i, 4))
    Next
    txt_ValorTotalEstimado = gvntConvVrDoSql(vntTotal)
Err_Handle:
End Sub


Private Sub LocalizarImobiliarioRural()
Dim strSql As String
Dim strCondicao As String
Dim strValor As String
Dim strCampo As String

Dim i As Integer

strCondicao = ""

With Me
    For i = 0 To .Controls.Count - 1
        
        If Not TypeOf .Controls(i) Is Label Then 'Elimina os Label's da pesquisa
            'Elimina objetos indesejáveis
            If UCase(.Controls(i).Name) <> "TXTPKId" _
            And UCase(Left(.Controls(i).Name, 3)) <> "IMG" _
            And UCase(Left(.Controls(i).Name, 3)) <> "LVW" _
            And UCase(Left(.Controls(i).Name, 3)) <> "TLB" _
            And UCase(Left(.Controls(i).Name, 3)) <> "TDD" _
            And UCase(Left(.Controls(i).Name, 3)) <> "GRD" _
            And UCase(.Controls(i).Name) <> UCase("txt_PKIdContribuinte") _
            And UCase(.Controls(i).Name) <> UCase("txtstrCNPJCPF") _
            And InStr(1, UCase(.Controls(i).Name), UCase("_Inscricao")) = 0 _
            And InStr(1, UCase(.Controls(i).Name), UCase("_Proprietario")) = 0 _
            And UCase(.Controls(i).Name) <> UCase("txtstrCNPJCPFP") _
            And InStr(1, UCase(.Controls(i).Name), UCase("txt_")) = 0 _
            And UCase(.Controls(i).Name) <> UCase("optbytNaturezaJuridica") _
            And UCase(.Controls(i).Name) <> UCase("txtdblValorImovel") Then
            
                If Not (TypeOf .Controls(i) Is OptionButton) Or .Controls(i) = True Then 'Elimina OptionButton desmarcado
                    If TypeOf .Controls(i) Is TextBox Then
                        If Trim(.Controls(i).Text) <> "" Then
                            If InStr(1, .Controls(i).Name, "Cep") > 0 Then
                                strValor = gstrValorSemMascara(Trim(.Controls(i).Text))
                            Else
                                strValor = Trim(.Controls(i).Text)
                            End If
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, "_", strCampo) > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            strCampo = "IM." & strCampo
                            
                            If InStr(1, "%", strValor) > 0 Then
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & strValor & "'"
                                Else
                                    strCondicao = strCampo & " LIKE '" & strValor & "'"
                                End If
                            ElseIf InStr(1, UCase(.Controls(i).Name), "DTM") > 0 Then
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & gstrConvDtParaSql(strValor)
                                Else
                                    strCondicao = strCampo & " = " & gstrConvDtParaSql(strValor)
                                End If
                            ElseIf InStr(1, UCase(.Controls(i).Name), "DBL") > 0 Then
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & gstrConvVrParaSql(strValor)
                                Else
                                    strCondicao = strCampo & " = " & gstrConvVrParaSql(strValor)
                                End If
                            Else
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & strValor & "%'"
                                Else
                                    strCondicao = strCampo & " LIKE '" & strValor & "%'"
                                End If
                            End If
                        End If
                    'If TypeOf .Controls(i) Is TextBox Then
                    ElseIf TypeOf .Controls(i) Is OptionButton Then
                        strValor = .Controls(i).Index
                        strCampo = Trim(.Controls(i).Name)
                        If InStr(1, "_", strCampo) > 0 Then
                            strCampo = Mid(strCampo, 5, Len(strCampo))
                        Else
                            strCampo = Mid(strCampo, 4, Len(strCampo))
                        End If
                        strCampo = "IM." & strCampo
                        
                        If strCondicao <> "" Then
                            strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                        Else
                            strCondicao = strCampo & " = " & strValor
                        End If
                    'ElseIf TypeOf .Controls(i) Is OptionButton Then
                    ElseIf TypeOf .Controls(i) Is CheckBox Then
                        If .Controls(i).Value = 1 Then
                            strValor = .Controls(i).Value
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, "_", strCampo) > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            strCampo = "IM." & strCampo
                            
                            If strCondicao <> "" Then
                                strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                            Else
                                strCondicao = strCampo & " = " & strValor
                            End If
                        End If
                    'ElseIf TypeOf .Controls(i) Is CheckBox Then
                    ElseIf TypeOf .Controls(i) Is DataCombo Then
                        If .Controls(i).MatchedWithList Then
                            strValor = .Controls(i).BoundText
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, "_", strCampo) > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            strCampo = "IM." & strCampo
                            
                            If strCondicao <> "" Then
                                strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                            Else
                                strCondicao = strCampo & " = " & strValor
                            End If
                        End If
                    ElseIf TypeOf .Controls(i) Is MaskEdBox Then
                        If Trim(.Controls(i).ClipText) <> "" Then
                            strValor = .Controls(i).ClipText
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, "_", strCampo) > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            strCampo = "IM." & strCampo
                            
                            If strCondicao <> "" Then
                                strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & strValor & "%'"
                            Else
                                strCondicao = strCampo & " LIKE '" & strValor & "%'"
                            End If
                        End If
                    End If 'If TypeOf .Controls(i) Is TextBox Then
                End If
            End If 'If Not (TypeOf .Controls(I) Is OptionButton) Or .Controls(I) = True Then
        End If 'If Not TypeOf .Controls(I) Is Label Then
    Next i
End With

strSql = ""
If strCondicao <> "" Then
    strSql = strSql & strQueryImobiliarioRural & " AND " & strCondicao
Else
    strSql = strQueryImobiliarioRural
End If

LeDaTabelaParaObj gstrImobiliarioRural, tdb_ImobiliarioRural, strSql

End Sub




