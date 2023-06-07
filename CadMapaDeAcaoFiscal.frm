VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadMapaDeAcaoFiscal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mapa de Ação Fiscal"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "CadMapaDeAcaoFiscal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_PKIdOS 
      Height          =   315
      Left            =   4740
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPKId 
      Height          =   315
      Left            =   5730
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4635
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8176
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Mapa de Ação Fiscal"
      TabPicture(0)   =   "CadMapaDeAcaoFiscal.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Tdb_Documento"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdd_Documento"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Cadastro"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_bytOrigem"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Relato de Ação Fiscal"
      TabPicture(1)   =   "CadMapaDeAcaoFiscal.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Fra_Somatorio"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   2925
         Left            =   -74850
         TabIndex        =   19
         Top             =   1620
         Width           =   6375
         Begin VB.TextBox txtstrAcaoFiscal 
            Height          =   2295
            Left            =   120
            TabIndex        =   5
            Top             =   510
            Width           =   2955
         End
         Begin VB.TextBox txtstrSancaoLegal 
            Height          =   2295
            Left            =   3240
            TabIndex        =   6
            Top             =   510
            Width           =   2955
         End
         Begin VB.Label lblstrAcaoFiscal 
            AutoSize        =   -1  'True
            Caption         =   "Relato de Ação Fiscal"
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label lblstrSancaoLegal 
            AutoSize        =   -1  'True
            Caption         =   "Sanção Legal"
            Height          =   195
            Left            =   3270
            TabIndex        =   20
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.Frame fra_bytOrigem 
         Caption         =   " Origem "
         Height          =   705
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   6285
         Begin VB.OptionButton optbytOrigem 
            Caption         =   "Econômico"
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   18
            Top             =   300
            Width           =   1155
         End
         Begin VB.OptionButton optbytOrigem 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   1
            Left            =   2130
            TabIndex        =   17
            Top             =   300
            Width           =   1695
         End
         Begin VB.OptionButton optbytOrigem 
            Caption         =   "Imobiliário Rural"
            Height          =   195
            Index           =   2
            Left            =   4290
            TabIndex        =   16
            Top             =   300
            Width           =   1425
         End
      End
      Begin VB.Frame Fra_Somatorio 
         Caption         =   "Somatório"
         Height          =   945
         Left            =   -74850
         TabIndex        =   11
         Top             =   630
         Width           =   6375
         Begin VB.TextBox txt_dblValorDocumento 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2010
            TabIndex        =   15
            Top             =   420
            Width           =   1335
         End
         Begin VB.TextBox txt_dblValorApurado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4890
            TabIndex        =   14
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label lbl_dblValorDocumento 
            AutoSize        =   -1  'True
            Caption         =   "Valores de Documentos"
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   465
            Width           =   1695
         End
         Begin VB.Label lbl_dblValorApurado 
            AutoSize        =   -1  'True
            Caption         =   "Valores Apurados"
            Height          =   285
            Left            =   3510
            TabIndex        =   12
            Top             =   450
            Width           =   1245
         End
      End
      Begin VB.Frame fra_Cadastro 
         Height          =   1335
         Left            =   120
         TabIndex        =   7
         Top             =   1290
         Width           =   6285
         Begin VB.TextBox txtstrNumeroProtocolo 
            Height          =   285
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   4
            Top             =   900
            Width           =   1725
         End
         Begin MSDataListLib.DataCombo dbc_intInscricaoCadastral 
            Height          =   315
            Left            =   2070
            TabIndex        =   2
            Top             =   240
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intFiscal 
            Height          =   315
            Left            =   2070
            TabIndex        =   3
            Top             =   570
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_intFiscal 
            AutoSize        =   -1  'True
            Caption         =   "Fiscal"
            Height          =   195
            Left            =   1545
            TabIndex        =   10
            Top             =   600
            Width           =   405
         End
         Begin VB.Label lblstrNumeroProtocolo 
            AutoSize        =   -1  'True
            Caption         =   "Número de Protocolo"
            Height          =   195
            Left            =   450
            TabIndex        =   9
            Top             =   930
            Width           =   1500
         End
         Begin VB.Label lbl_intInscricaoCadastral 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   600
            TabIndex        =   8
            Top             =   300
            Width           =   1350
         End
      End
      Begin TrueOleDBGrid70.TDBDropDown tdd_Documento 
         Height          =   1425
         Left            =   420
         TabIndex        =   23
         Top             =   3090
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2514
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PkId"
         Columns(0).DataField=   "PkId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Abreviatura"
         Columns(1).DataField=   "strAbreviatura"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1614"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1535"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits.Count    =   1
         AllowRowSizing  =   0   'False
         Appearance      =   1
         BorderStyle     =   1
         ColumnHeaders   =   -1  'True
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowDividerStyle =   2
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   0   'False
         ListField       =   "strDescricao"
         DataField       =   "PkId"
         IntegralHeight  =   0   'False
         FetchRowStyle   =   0   'False
         AlternatingRowStyle=   0   'False
         DataMember      =   ""
         ColumnFooters   =   0   'False
         FootLines       =   1
         DeadAreaBackColor=   12632256
         ValueTranslate  =   -1  'True
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Named:id=33:Normal"
         _StyleDefs(43)  =   ":id=33,.parent=0"
         _StyleDefs(44)  =   "Named:id=34:Heading"
         _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(46)  =   ":id=34,.wraptext=-1"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid Tdb_Documento 
         Height          =   1845
         Left            =   120
         TabIndex        =   24
         Top             =   2700
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   3254
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   1
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tipo de Documentos"
         Columns(0).DataField=   "intDocumento"
         Columns(0).DropDown=   "tdd_Documento"
         Columns(0).DropDown.vt=   8
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Valor do Documento"
         Columns(1).DataField=   "dblValorDocumento"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor Apurado"
         Columns(2).DataField=   "dblValorApurado"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=5133"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5054"
         Splits(0)._ColumnProps(4)=   "Column(0).Button=1"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(0).AutoDropDown=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AutoCompletion=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(14)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33"
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Geral 
      Height          =   1695
      Left            =   90
      TabIndex        =   22
      Top             =   4770
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   2990
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "PKId"
      Columns(0).DataField=   "PKID"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nº Protocolo"
      Columns(1).DataField=   "strProtocolo"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Inscrição Cadastral"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Fiscal"
      Columns(3).DataField=   "strNome"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2170"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2090"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=6350"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=6271"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadMapaDeAcaoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando       As Boolean
Dim mobjAux             As Object
Dim mblnSelecionou      As Boolean
Dim mblnPrimeiraVez     As Boolean
Dim opt                 As Integer
Dim Matriz              As XArrayDB
Dim bytCont             As Byte

Private Sub dbc_intFiscal_Click(Area As Integer)
    DropDownDataCombo dbc_intFiscal, Me, Area
    If Area = 2 And dbc_intFiscal.MatchedWithList Then
        HabilitaDesabilitaCamposInterno True
        PreenchePKIdOS
    End If
End Sub

Private Sub dbc_intFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intFiscal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intInscricaoCadastral_Click(Area As Integer)
    DropDownDataCombo dbc_intInscricaoCadastral, Me, Area
    If Area = 2 And dbc_intInscricaoCadastral.MatchedWithList Then
        Dim strSql As String
        strSql = ""
        strSql = strSql & " SELECT FC.PKId, CC.strNome "
        strSql = strSql & " FROM " & gstrFiscais & " FC,"
        strSql = strSql & gstrContribuinte & " CC,"
        strSql = strSql & gstrOrdemServico & " OS,"
        strSql = strSql & gstrOrdemServicoFiscal & " OSF"
        strSql = strSql & " WHERE CC.PKId = FC.intContribuinte "
        strSql = strSql & " AND FC.PKId = OSF.intFiscal"
        strSql = strSql & " AND OS.PKId = OSF.intOrdemServico"
        strSql = strSql & " AND OS.intInscricaoCadastral = " & dbc_intInscricaoCadastral.BoundText
        LeDaTabelaParaObj "", dbc_intFiscal, strSql
        dbc_intFiscal.Enabled = True
        TrocaCorObjeto dbc_intFiscal, False
    End If
End Sub

Private Sub dbc_intInscricaoCadastral_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intInscricaoCadastral, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 650
    VirificaGradeListView Me
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
    opt = 3
    mblnAlterando = False
    VerificaObjParaAplicar mobjAux
    LeDaTabelaParaObj "", tdd_Documento, Query
    HabilitaDesabilitaCamposInterno False
    txt_dblValorApurado.Enabled = False
    TrocaCorObjeto txt_dblValorApurado, True
    txt_dblValorDocumento.Enabled = False
    TrocaCorObjeto txt_dblValorDocumento, True
    Set Matriz = New XArrayDB
    Matriz.ReDim 0, 0, 0, 3
    Set Tdb_Documento.Array = Matriz
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
Dim ValorDocumento As Double
Dim ValorApurado   As Double
    ValorDocumento = 0
    ValorApurado = 0
    Tdb_Documento.Update
    
    Select Case tab_3dPasta.Tab
        Case 1
            Tdb_Documento.MoveFirst
            Do While Not Tdb_Documento.EOF
                If Tdb_Documento.Columns(1) <> "" And Tdb_Documento.Columns(2) <> "" Then
                    ValorDocumento = ValorDocumento + CDbl(Tdb_Documento.Columns(1).Text)
                    ValorApurado = ValorApurado + CDbl(Tdb_Documento.Columns(2).Text)
                End If
                Tdb_Documento.MoveNext
            Loop
            txt_dblValorDocumento.Text = gstrConvVrDoSql(ValorDocumento)
            txt_dblValorApurado.Text = gstrConvVrDoSql(ValorApurado)
    End Select
End Sub

Private Sub Tdb_Documento_KeyPress(KeyAscii As Integer)
    Select Case Tdb_Documento.col
    Case 1
        CaracterValido KeyAscii, "V", Tdb_Documento.Columns(1)
    Case 2
        CaracterValido KeyAscii, "V", Tdb_Documento.Columns(2)
    End Select
End Sub

Private Sub tdb_Geral_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Geral
        If Not .EOF And Not .BOF Then
            txtPKId.Text = .Columns("PKID").Value

            If mblnPrimeiraVez Then

                LeDadosDoTDB
                gCorLinhaSelecionada tdb_Geral

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar

                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
            End If

        End If
    End With
End Sub

Private Sub tdb_Geral_Click()
    mblnPrimeiraVez = True
    With tdb_Geral
        If Not .EOF And Not .BOF Then
            If .Bookmark = 1 Then
                tdb_Geral_RowColChange 0, 0
            End If
        End If
    End With
End Sub

Sub tdb_Geral_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Geral_FilterChange()
    gblnFilraCampos tdb_Geral
End Sub

Private Function Query() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strAbreviatura,  strDescricao "
    strSql = strSql & " FROM " & gstrTabelaDocumento
Query = strSql
End Function
Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSql As String

If UCase(strModoOperacao) = gstrPreencherLista Then
    PreencherListaDeOpcoes Me.ActiveControl
    Exit Sub
End If

If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
    mblnPrimeiraVez = False
End If

If UCase(strModoOperacao) = "SALVAR" Then
    If DadosOk Then
        If mblnAlterando Then
            If MsgBox("Confirma Alteração ?", vbYesNo + vbQuestion) = vbYes Then
                Query_UPDATE
            End If
        Else
            If MsgBox("Confirma Inclusão ?", vbYesNo + vbQuestion) = vbYes Then
                Query_INSERT
            End If
        End If
        Query_NOVO
        If opt <> 3 Then
            optbytOrigem_Click opt
        Else
            optbytOrigem_Click 0
        End If
        tab_3dPasta.Tab = 0
    End If
End If
If UCase(strModoOperacao) = "NOVO" Then
    Query_NOVO
End If
If UCase(strModoOperacao) = "DELETAR" Then
    If txtPKId <> "" Then
        If MsgBox("Confirma Exclusão ?", vbYesNo + vbQuestion) = vbYes Then
            Query_DELETE
            Query_NOVO
            If opt <> 3 Then
                optbytOrigem_Click opt
            Else
                optbytOrigem_Click 0
            End If
            tab_3dPasta.Tab = 0
        End If
    Else
        ExibeMensagem "Deve ser Selecionado algum Registro"
    End If
End If
HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
End Sub

Private Function HabilitaDesabilitaCamposInterno(Habilitar As Boolean)
    Dim Enabled As Boolean
    Dim TrocaCor As Boolean
    If Habilitar Then
        Enabled = True
        TrocaCor = False
    Else
        Enabled = False
        TrocaCor = True
    End If
    dbc_intInscricaoCadastral.Enabled = Enabled
    TrocaCorObjeto dbc_intInscricaoCadastral, TrocaCor
    dbc_intFiscal.Enabled = Enabled
    TrocaCorObjeto dbc_intFiscal, TrocaCor
    txtstrNumeroProtocolo.Enabled = Enabled
    TrocaCorObjeto txtstrNumeroProtocolo, TrocaCor
    txtstrAcaoFiscal.Enabled = Enabled
    TrocaCorObjeto txtstrAcaoFiscal, TrocaCor
    txtstrSancaoLegal.Enabled = Enabled
    TrocaCorObjeto txtstrSancaoLegal, TrocaCor
End Function

Private Function DadosOk() As Boolean
Dim strSql As String
    If txtstrAcaoFiscal = "" Then
        ExibeMensagem "Deve ser digitado alguma Ação Fiscal"
        DadosOk = False
        Exit Function
    End If
    If txtstrNumeroProtocolo.Text = "" Then
        ExibeMensagem "Deve ser digitado algum Número de Protocolo"
        DadosOk = False
        Exit Function
    End If
    If txtstrSancaoLegal.Text = "" Then
        ExibeMensagem "Deve ser digitado alguma Sanção Legal"
        DadosOk = False
        Exit Function
    End If
    If dbc_intFiscal.BoundText = "" Then
        ExibeMensagem "Deve ser Selecionado algum Fiscal"
        DadosOk = False
        Exit Function
    End If
    If bytCont = 0 Then
        ExibeMensagem "Deve ser Selecionado Alguma Documento"
        DadosOk = False
        Exit Function
    End If
DadosOk = True
End Function

Private Sub Tdb_Documento_AfterColUpdate(ByVal ColIndex As Integer)
'If Tdb_Documento.Columns(1).Text = "" Or Tdb_Documento.Columns(2).Text = "" Then Exit Sub
    Tdb_Documento.Columns(1).Text = gstrConvVrDoSql(Tdb_Documento.Columns(1).Text)
    Tdb_Documento.Columns(2).Text = gstrConvVrDoSql(Tdb_Documento.Columns(2).Text)

Tdb_Documento.Update
End Sub

Private Sub Tdb_Documento_AfterUpdate()
Tdb_Documento.Update
End Sub

Private Sub PreenchePKIdOS()
Dim adoResultado As ADODB.Recordset
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT OSF.PKId FROM " & gstrOrdemServicoFiscal & " OSF,"
    strSql = strSql & gstrOrdemServico & " OS"
    strSql = strSql & " WHERE OS.bytOrigem = " & opt
    strSql = strSql & " AND OS.intInscricaoCadastral = " & dbc_intInscricaoCadastral.BoundText
    strSql = strSql & " AND OSF.intOrdemServico = OS.PKId "
    strSql = strSql & " AND OSF.intFiscal = " & dbc_intFiscal.BoundText
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_PKIdOS = adoResultado!PKId
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub optbytOrigem_Click(Index As Integer)
Dim strSql As String
strSql = ""
HabilitaDesabilitaCamposInterno False
dbc_intInscricaoCadastral.Enabled = True
TrocaCorObjeto dbc_intInscricaoCadastral, False
dbc_intInscricaoCadastral.BoundText = ""
Set dbc_intInscricaoCadastral.RowSource = Nothing
opt = Index
If Index = 0 Then 'Economico
    tdb_Geral.Columns(2).DataField = "strInscricaoCadastral"
    Insc_Economico
    Query_Economico
Else
    tdb_Geral.Columns(2).DataField = "strInscricaoAnterior"
    If Index = 1 Then 'Imobiliario Urbano
        Insc_Urbano
        Query_Urbano
    Else 'Imobiliario Rural
        Insc_Rural
        Query_Rural
    End If
End If
Query_NOVO
mblnPrimeiraVez = False
End Sub

Private Sub Insc_Economico()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT EC.PKId, EC.strInscricaoCadastral "
    strSql = strSql & " FROM " & gstrEconomico & " EC,"
    strSql = strSql & gstrOrdemServico & " OS"
    strSql = strSql & " WHERE EC.PKId = OS.intInscricaoCadastral "
    strSql = strSql & " AND OS.BytOrigem = " & opt
    dbc_intInscricaoCadastral.Tag = strSql & ";EC.strInscricaoCadastral "
End Sub

Private Sub Insc_Urbano()
Dim strSql As String
    strSql = strSql & " SELECT IM.PKId, IM.strInscricaoAnterior "
    strSql = strSql & " FROM " & gstrImobiliario & " IM,"
    strSql = strSql & gstrOrdemServico & " OS"
    strSql = strSql & " WHERE IM.PKId = OS.intInscricaoCadastral "
    strSql = strSql & " AND OS.BytOrigem = " & opt
    dbc_intInscricaoCadastral.Tag = strSql & ";IM.strInscricaoAnterior"
End Sub

Private Sub Insc_Rural()
Dim strSql As String
    strSql = strSql & " SELECT IM.PKId, IM.strInscricaoAnterior "
    strSql = strSql & " FROM " & gstrImobiliarioRural & " IM,"
    strSql = strSql & gstrOrdemServico & " OS"
    strSql = strSql & " WHERE IM.PKId = OS.intInscricaoCadastral "
    strSql = strSql & " AND OS.BytOrigem = " & opt
    dbc_intInscricaoCadastral.Tag = strSql & ";IM.strInscricaoAnterior"
End Sub

Private Sub Query_Economico()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT MA.PKId, MA.strProtocolo, EC.strInscricaoCadastral, CC.strNome "
    strSql = strSql & " FROM " & gstrMapaAcaoFiscal & " MA,"
    strSql = strSql & gstrEconomico & " EC," & gstrContribuinte & " CC,"
    strSql = strSql & gstrFiscais & " FC," & gstrOrdemServicoFiscal & " OSF,"
    strSql = strSql & gstrOrdemServico & " OS"
    strSql = strSql & " WHERE CC.PKId = FC.intContribuinte "
    strSql = strSql & " AND FC.PKId = OSF.intFiscal "
    strSql = strSql & " AND EC.PKId = OS.intInscricaoCadastral "
    strSql = strSql & " AND OS.PKId = OSF.intOrdemServico "
    strSql = strSql & " AND OSF.PKId = MA.intOSFiscalizacao " 'intOSFiscalizacao
    strSql = strSql & " AND OS.bytOrigem = " & opt
    LeDaTabelaParaObj "", tdb_Geral, strSql
End Sub

Private Sub Query_Urbano()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT MA.PKId, MA.strProtocolo, IM.strInscricaoAnterior, CC.strNome "
    strSql = strSql & " FROM " & gstrMapaAcaoFiscal & " MA,"
    strSql = strSql & gstrImobiliario & " IM," & gstrContribuinte & " CC,"
    strSql = strSql & gstrFiscais & " FC," & gstrOrdemServicoFiscal & " OSF,"
    strSql = strSql & gstrOrdemServico & " OS"
    strSql = strSql & " WHERE CC.PKId = FC.intContribuinte "
    strSql = strSql & " AND FC.PKId = OSF.intFiscal "
    strSql = strSql & " AND IM.PKId = OS.intInscricaoCadastral "
    strSql = strSql & " AND OS.PKId = OSF.intOrdemServico "
    strSql = strSql & " AND OSF.PKId = MA.intOSFiscalizacao "
    strSql = strSql & " AND OS.bytOrigem = " & opt
    LeDaTabelaParaObj "", tdb_Geral, strSql
End Sub

Private Sub Query_Rural()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT MA.PKId, MA.strProtocolo, IM.strInscricaoAnterior, CC.strNome "
    strSql = strSql & " FROM " & gstrMapaAcaoFiscal & " MA,"
    strSql = strSql & gstrImobiliarioRural & " IM," & gstrContribuinte & " CC,"
    strSql = strSql & gstrFiscais & " FC," & gstrOrdemServicoFiscal & " OSF,"
    strSql = strSql & gstrOrdemServico & " OS"
    strSql = strSql & " WHERE CC.PKId = FC.intContribuinte "
    strSql = strSql & " AND FC.PKId = OSF.intFiscal "
    strSql = strSql & " AND IM.PKId = OS.intInscricaoCadastral "
    strSql = strSql & " AND OS.PKId = OSF.intOrdemServico "
    strSql = strSql & " AND OSF.PKId = MA.intOSFiscalizacao "
    strSql = strSql & " AND OS.bytOrigem = " & opt
    LeDaTabelaParaObj "", tdb_Geral, strSql
End Sub

Private Sub Query_INSERT()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 12/05/2003
' Alteração: - Substituição do comando INSERT INTO SELECT pelo comando INSERT INTO VALUES.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 12/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
Dim i As Integer
Dim MaxPKId As Integer
Dim adoResultado    As ADODB.Recordset
    'set Tdb_Documento.DataSource =
    
    'Incluir Na tablea MapaAcaoFiscal
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
  
'    strSql = strSql & " INSERT INTO " & gstrMapaAcaoFiscal & " SELECT "
    strSql = strSql & " INSERT INTO " & gstrMapaAcaoFiscal
    
    strSql = strSql & "(intOSFiscalizacao, strAcaoFiscal, strSancaoLegal, "
    strSql = strSql & "strProtocolo, dtmDtAtualizacao, lngCodUsr) VALUES ("
    
    strSql = strSql & txt_PKIdOS & ",'" & txtstrAcaoFiscal & "','"
    strSql = strSql & txtstrSancaoLegal & "','" & txtstrNumeroProtocolo.Text
'    strSql = strSql & "',GETDATE() , " & glngCodUsr
    strSql = strSql & "'," & strGETDATE & " , " & glngCodUsr
    
    strSql = strSql & ")"
  
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
  
    MaxPKId = MaximoPKId

    strSql = ""
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    'Incluir na Tabela  MapaAcaoFiscalDocumento (em Loop)
    'Set Matriz = Tdb_Documento.Array
    For i = 0 To Matriz.Count(1) - 1
'        strSql = strSql & " INSERT INTO " & gstrMapaAcaoFiscalDocumento & " SELECT "
        strSql = strSql & " INSERT INTO " & gstrMapaAcaoFiscalDocumento
        
        strSql = strSql & "(intMapaAcaoFiscal, intDocumento, dblValorDocumento, "
        strSql = strSql & "dblValorApurado, dtmDtAtualizacao, lngCodUsr) VALUES ("
        
        strSql = strSql & MaxPKId & ","
        strSql = strSql & Matriz.Value(i, 0)       'intDocumento - Da Matriz
        strSql = strSql & "," & gstrConvVrParaSql(Matriz.Value(i, 1)) 'dblValorDocumento - Da Matriz
        strSql = strSql & "," & gstrConvVrParaSql(Matriz.Value(i, 2))  'dblValorApurado - Da Matriz
'        strSql = strSql & ", GETDATE()," & glngCodUsr
        strSql = strSql & ", " & strGETDATE & "," & glngCodUsr
        
        strSql = strSql & ")"
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
    Next i
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
    If Not gobjBanco.Execute(strSql, False) Then
        gobjBanco.ExecutaRollbackTrans
    Else
        gobjBanco.ExecutaCommitTrans
    End If
End Sub

Private Sub Query_DELETE()

'******************************************************************************************
' Data: 12/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSql As String
    strSql = ""
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    'Deleta da Tabela de MapaAcaoFiscalDocumento
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    strSql = strSql & " DELETE FROM " & gstrMapaAcaoFiscalDocumento
    strSql = strSql & " WHERE intMapaAcaoFiscal = " & txtPKId
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
    
    'Deleta da Tabela de MapaAcaoFiscal
    strSql = strSql & " DELETE FROM " & gstrMapaAcaoFiscal
    strSql = strSql & " WHERE PKId = " & txtPKId
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
    Set gobjBanco = New clsBanco
    If Not gobjBanco.Execute(strSql, False) Then
        gobjBanco.ExecutaRollbackTrans
    Else
        gobjBanco.ExecutaCommitTrans
    End If
End Sub

Private Sub Query_UPDATE()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 12/05/2003
' Alteração: - Substituição do comando INSERT INTO SELECT pelo comando INSERT INTO VALUES.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 12/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
Dim i As Integer
    strSql = ""
    
    'Deleta da Tabela de MapaAcaoFiscalDocumento
    strSql = strSql & " DELETE FROM " & gstrMapaAcaoFiscalDocumento
    strSql = strSql & " WHERE intMapaAcaoFiscal = " & txtPKId
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    strSql = ""
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    'Atualiza a Tablea MapaAcaoFiscal
    strSql = strSql & " UPDATE " & gstrMapaAcaoFiscal & " SET "
    strSql = strSql & " intOSFiscalizacao = " & txt_PKIdOS
    strSql = strSql & ", strAcaoFiscal = '" & txtstrAcaoFiscal
    strSql = strSql & "', strSancaoLegal =  '" & txtstrSancaoLegal
    strSql = strSql & "', strProtocolo = '" & txtstrNumeroProtocolo
'    strSql = strSql & "', dtmDtAtualizacao = GETDATE(), LngCodUsr = "
    strSql = strSql & "', dtmDtAtualizacao = " & strGETDATE & ", LngCodUsr = "
    strSql = strSql & glngCodUsr
    strSql = strSql & " WHERE PKId = " & txtPKId
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
    
    'Incluir na Tabela  MapaAcaoFiscalDocumento (em Loop)
    For i = 0 To Matriz.Count(1) - 1
'        strSql = strSql & " INSERT INTO " & gstrMapaAcaoFiscalDocumento & " SELECT "
        strSql = strSql & " INSERT INTO " & gstrMapaAcaoFiscalDocumento
        
        strSql = strSql & "(intMapaAcaoFiscal, intDocumento, dblValorDocumento, "
        strSql = strSql & "dblValorApurado, dtmDtAtualizacao, lngCodUsr) VALUES ("
        
        strSql = strSql & txtPKId & ","
        strSql = strSql & Matriz.Value(i, 0)       'intDocumento - Da Matriz
        strSql = strSql & "," & gstrConvVrParaSql(Matriz.Value(i, 1)) 'dblValorDocumento - Da Matriz
        strSql = strSql & "," & gstrConvVrParaSql(Matriz.Value(i, 2))  'dblValorApurado - Da Matriz
'        strSql = strSql & ", GETDATE()," & glngCodUsr
        strSql = strSql & ", " & strGETDATE & "," & glngCodUsr
    
        strSql = strSql & ")"
    
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
    
    Next i
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
    If Not gobjBanco.Execute(strSql, False) Then
        gobjBanco.ExecutaRollbackTrans
    Else
        gobjBanco.ExecutaCommitTrans
    End If
End Sub

Private Sub Query_NOVO()
dbc_intInscricaoCadastral.BoundText = ""
dbc_intFiscal.BoundText = ""
txtstrNumeroProtocolo.Text = ""
txt_dblValorApurado.Text = ""
txt_dblValorDocumento.Text = ""
txtstrAcaoFiscal.Text = ""
txtstrSancaoLegal.Text = ""
txt_PKIdOS.Text = ""
txtPKId.Text = ""
Matriz.Clear
Set Tdb_Documento.Array = Matriz
Tdb_Documento.ReBind
Tdb_Documento.Refresh
'If opt <> 3 Then
'    optbytOrigem_Click opt
'Else
'    optbytOrigem_Click 0
'End If
mblnAlterando = False
End Sub

Private Function MaximoPKId() As Integer
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    strSql = ""
    strSql = strSql & " SELECT MAX(PKId) AS PKId FROM " & gstrMapaAcaoFiscal
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not (.BOF And .EOF) Then
                 MaximoPKId = (!PKId)
            End If
        End With
    End If
End Function

Private Function LeDadosDoTDB()
Dim strSql As String
Dim adoResultado As ADODB.Recordset

strSql = strSql & " SELECT OSF.PKId, OS.intInscricaoCadastral, OSF.intFiscal, "
strSql = strSql & " MA.strProtocolo, MA.strAcaoFiscal, MA.strSancaoLegal "
strSql = strSql & " FROM " & gstrMapaAcaoFiscal & " MA,"
strSql = strSql & gstrOrdemServicoFiscal & " OSF,"
strSql = strSql & gstrOrdemServico & " OS"
strSql = strSql & " WHERE OS.PKId = OSF.intOrdemServico "
strSql = strSql & " AND OSF.PKId = MA.intOSFiscalizacao "
strSql = strSql & " AND MA.PKId = " & txtPKId
strSql = strSql & " AND OS.bytOrigem = " & opt

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
    With adoResultado
        If Not .EOF Then
            If IsNull(!PKId) Then
                txt_PKIdOS = ""
            Else
                txt_PKIdOS = (!PKId)
            End If
            If IsNull(!intInscricaoCadastral) Then
                dbc_intInscricaoCadastral.BoundText = ""
            Else
                dbc_intInscricaoCadastral.BoundText = (!intInscricaoCadastral)
            End If
            dbc_intInscricaoCadastral_Click 2
            If IsNull(!intFiscal) Then
                dbc_intFiscal.BoundText = ""
            Else
                dbc_intFiscal.BoundText = (!intFiscal)
            End If
            If IsNull(!strProtocolo) Then
                txtstrNumeroProtocolo = ""
            Else
                txtstrNumeroProtocolo = (!strProtocolo)
            End If
            If IsNull(!strAcaoFiscal) Then
                txtstrAcaoFiscal = ""
            Else
                txtstrAcaoFiscal = (!strAcaoFiscal)
            End If
            If IsNull(!strSancaoLegal) Then
                txtstrSancaoLegal = ""
            Else
                txtstrSancaoLegal = (!strSancaoLegal)
            End If
            HabilitaDesabilitaCamposInterno True
            PreenchePKIdOS
            PreencheMatriz
       End If
    End With
End If
Set gobjBanco = Nothing
Set adoResultado = Nothing
End Function

Private Sub PreencheMatriz()
Dim strSql As String
Dim adoRec As ADODB.Recordset
Dim varAux As String

bytCont = 0
On Error GoTo Err_Handle

Set Matriz = New XArrayDB
Matriz.Clear

Matriz.ReDim 0, 0, 0, 3


strSql = ""
strSql = strSql & " SELECT MAD.intDocumento, MAD.dblValorDocumento, MAD.dblValorApurado "
strSql = strSql & " FROM " & gstrMapaAcaoFiscal & " MA,"
strSql = strSql & gstrMapaAcaoFiscalDocumento & " MAD "
strSql = strSql & " WHERE MA.PKId = MAD.intMapaAcaoFiscal "
strSql = strSql & " AND MA.PKId = " & txtPKId

txt_dblValorDocumento = ""
txt_dblValorApurado = ""

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSql, 5, adoRec) Then
    With adoRec
        If Not .EOF Then
            Matriz.ReDim 0, .RecordCount - 1, 0, 3
            Do While Not .EOF
                bytCont = 1
                varAux = !intDocumento
                Matriz(.AbsolutePosition - 1, 0) = varAux
                
                varAux = !dblValorDocumento
                Matriz(.AbsolutePosition - 1, 1) = gstrConvVrDoSql(varAux)
                txt_dblValorDocumento = gstrConvVrDoSql(Val(gstrConvVrParaSql(txt_dblValorDocumento)) + Val(gstrConvVrParaSql(varAux)))
                
                varAux = !dblValorApurado
                Matriz(.AbsolutePosition - 1, 2) = gstrConvVrDoSql(varAux)
                txt_dblValorApurado = gstrConvVrDoSql(Val(gstrConvVrParaSql(txt_dblValorApurado)) + Val(gstrConvVrParaSql(varAux)))
                .MoveNext
            Loop
        End If
    End With
End If

Set Tdb_Documento.Array = Matriz
Tdb_Documento.ReBind
Tdb_Documento.Refresh

Exit Sub
Err_Handle:
End Sub

Private Sub txt_dblValorApurado_KeyPress(KeyAscii As Integer)
        CaracterValido KeyAscii, "V", txt_dblValorApurado
End Sub

Private Sub txt_dblValorDocumento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorDocumento
End Sub

Private Sub txtstrNumeroProtocolo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNumeroProtocolo
End Sub
