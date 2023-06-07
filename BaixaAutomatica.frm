VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmBaixaAutomatica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Baixa Automática"
   ClientHeight    =   6810
   ClientLeft      =   840
   ClientTop       =   1605
   ClientWidth     =   8535
   Icon            =   "BaixaAutomatica.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_Parametros 
      Height          =   6675
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   11774
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros"
      TabPicture(0)   =   "BaixaAutomatica.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_BaixaAutomatica"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Dados"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Guias"
      TabPicture(1)   =   "BaixaAutomatica.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_DadosdoDebito"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "tdb_Baixa"
      Tab(1).ControlCount=   3
      Begin TrueOleDBGrid70.TDBGrid tdb_Baixa 
         Height          =   2235
         Left            =   -74820
         TabIndex        =   16
         Top             =   4260
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   3942
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Inscrição Cadastral"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Exercício"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Sequência"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Número Parcela"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Composição da Receita"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Contribuinte"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Lançamento"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Vencimento"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Valor Parcela"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Juros"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Multa"
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Correção"
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Desconto"
         Columns(12).DataField=   ""
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Total Pago"
         Columns(13).DataField=   ""
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Ocorrência"
         Columns(14).DataField=   ""
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Pagamento"
         Columns(15).DataField=   ""
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Código Lançamento Cálculo"
         Columns(16).DataField=   ""
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Nome"
         Columns(17).DataField=   ""
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   18
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=18"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2805"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1402"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1323"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1614"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1535"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2275"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2196"
         Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(25)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(30)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(33)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(38)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(41)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(43)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(44)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(46)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=770"
         Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(49)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(52)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=770"
         Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(55)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(56)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(58)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(10)._ColStyle=770"
         Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(61)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(62)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(64)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(11)._ColStyle=770"
         Splits(0)._ColumnProps(66)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(67)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(68)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(70)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(12)._ColStyle=770"
         Splits(0)._ColumnProps(72)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(73)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(74)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(76)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(77)=   "Column(13)._ColStyle=770"
         Splits(0)._ColumnProps(78)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(79)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(80)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(82)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(83)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(84)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(85)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(86)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(88)=   "Column(15)._EditAlways=0"
         Splits(0)._ColumnProps(89)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(90)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(91)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(92)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(93)=   "Column(16)._EditAlways=0"
         Splits(0)._ColumnProps(94)=   "Column(16).Visible=0"
         Splits(0)._ColumnProps(95)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(96)=   "Column(17).Width=9234"
         Splits(0)._ColumnProps(97)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(98)=   "Column(17)._WidthInPix=9155"
         Splits(0)._ColumnProps(99)=   "Column(17)._EditAlways=0"
         Splits(0)._ColumnProps(100)=   "Column(17).Order=18"
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
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14,.alignment=1"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14,.alignment=1"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=1"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14,.alignment=1"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.alignment=1"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14,.alignment=1"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=86,.parent=13,.alignment=1"
         _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14,.alignment=1"
         _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=90,.parent=13,.alignment=1"
         _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14,.alignment=1"
         _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
         _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
         _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
         _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
         _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
         _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
         _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
         _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
         _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=102,.parent=13"
         _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
         _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
         _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
         _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=106,.parent=13"
         _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
         _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
         _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
         _StyleDefs(109) =   "Named:id=33:Normal"
         _StyleDefs(110) =   ":id=33,.parent=0"
         _StyleDefs(111) =   "Named:id=34:Heading"
         _StyleDefs(112) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(113) =   ":id=34,.wraptext=-1"
         _StyleDefs(114) =   "Named:id=35:Footing"
         _StyleDefs(115) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(116) =   "Named:id=36:Selected"
         _StyleDefs(117) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(118) =   "Named:id=37:Caption"
         _StyleDefs(119) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(120) =   "Named:id=38:HighlightRow"
         _StyleDefs(121) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(122) =   "Named:id=39:EvenRow"
         _StyleDefs(123) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(124) =   "Named:id=40:OddRow"
         _StyleDefs(125) =   ":id=40,.parent=33"
         _StyleDefs(126) =   "Named:id=41:RecordSelector"
         _StyleDefs(127) =   ":id=41,.parent=34"
         _StyleDefs(128) =   "Named:id=42:FilterBar"
         _StyleDefs(129) =   ":id=42,.parent=33"
      End
      Begin VB.Frame Frame2 
         Caption         =   " Dados da Guia/Contribuinte "
         Height          =   1995
         Left            =   -74820
         TabIndex        =   39
         Top             =   450
         Width           =   8025
         Begin VB.TextBox txt_Contribuinte 
            Height          =   285
            Left            =   1755
            TabIndex        =   6
            Top             =   1260
            Width           =   4845
         End
         Begin VB.TextBox txt_dtmDataLancamento 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1755
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1605
            Width           =   975
         End
         Begin VB.TextBox txt_dblValorParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5205
            TabIndex        =   9
            Top             =   1605
            Width           =   1395
         End
         Begin VB.TextBox txt_dtmDataVencimento 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3705
            MaxLength       =   15
            TabIndex        =   8
            Top             =   1605
            Width           =   975
         End
         Begin VB.TextBox txtintNumeroParcela 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   5355
            MaxLength       =   5
            TabIndex        =   4
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtintExercicio 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1755
            MaxLength       =   4
            TabIndex        =   2
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txt_strSequencia 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   3
            Top             =   570
            Width           =   555
         End
         Begin MSMask.MaskEdBox mskInscricaoCadastral 
            Height          =   285
            Left            =   1755
            TabIndex        =   1
            Top             =   240
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo dbc_intComposicaoReceita 
            Height          =   315
            Left            =   1755
            TabIndex        =   5
            Top             =   900
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lbldtmDataLancamento 
            AutoSize        =   -1  'True
            Caption         =   "Data do Lançamento"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   1695
            Width           =   1500
         End
         Begin VB.Label lblintComposicaoReceita 
            AutoSize        =   -1  'True
            Caption         =   "Origem da Receita"
            Height          =   195
            Left            =   300
            TabIndex        =   47
            Top             =   990
            Width           =   1320
         End
         Begin VB.Label lbldblValorParcela 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   4740
            TabIndex        =   46
            Top             =   1695
            Width           =   360
         End
         Begin VB.Label lbldtmDataVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento"
            Height          =   195
            Left            =   2760
            TabIndex        =   45
            Top             =   1695
            Width           =   840
         End
         Begin VB.Label lblintContribuinte 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte"
            Height          =   195
            Left            =   780
            TabIndex        =   44
            Top             =   1350
            Width           =   840
         End
         Begin VB.Label lbldblNumeroParcela 
            AutoSize        =   -1  'True
            Caption         =   "Número da Parcela"
            Height          =   195
            Left            =   3885
            TabIndex        =   43
            Top             =   660
            Width           =   1365
         End
         Begin VB.Label lbldtmExercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   975
            TabIndex        =   42
            Top             =   660
            Width           =   675
         End
         Begin VB.Label lbl_InscricaoCadastral 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   300
            TabIndex        =   41
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label lblstrSequencia 
            AutoSize        =   -1  'True
            Caption         =   "Sequência"
            Height          =   195
            Left            =   2400
            TabIndex        =   40
            Top             =   660
            Width           =   765
         End
      End
      Begin VB.Frame fra_DadosdoDebito 
         Caption         =   " Dados do Débito "
         Height          =   1695
         Left            =   -74820
         TabIndex        =   32
         Top             =   2490
         Width           =   8025
         Begin VB.TextBox txtdblDesconto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4185
            TabIndex        =   13
            Top             =   570
            Width           =   1605
         End
         Begin VB.TextBox txtdblTotalPago 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1725
            TabIndex        =   14
            Top             =   900
            Width           =   1605
         End
         Begin VB.TextBox txtdblCorrecao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1725
            TabIndex        =   12
            Top             =   570
            Width           =   1605
         End
         Begin VB.TextBox txtdblMulta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4185
            TabIndex        =   11
            Top             =   210
            Width           =   1605
         End
         Begin VB.TextBox txtdblJuros 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1725
            TabIndex        =   10
            Top             =   210
            Width           =   1605
         End
         Begin MSDataListLib.DataCombo dbcintOcorrencia 
            Height          =   315
            Left            =   1725
            TabIndex        =   15
            Top             =   1230
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
            Height          =   195
            Left            =   3435
            TabIndex        =   38
            Top             =   660
            Width           =   690
         End
         Begin VB.Label lbldblTotalPago 
            AutoSize        =   -1  'True
            Caption         =   "Total Pago"
            Height          =   195
            Left            =   885
            TabIndex        =   37
            Top             =   990
            Width           =   780
         End
         Begin VB.Label lbldblCorrecao 
            AutoSize        =   -1  'True
            Caption         =   "Correção"
            Height          =   195
            Left            =   1020
            TabIndex        =   36
            Top             =   660
            Width           =   645
         End
         Begin VB.Label lbldblMulta 
            AutoSize        =   -1  'True
            Caption         =   "Multa"
            Height          =   195
            Left            =   3735
            TabIndex        =   35
            Top             =   300
            Width           =   390
         End
         Begin VB.Label lbldblJuros 
            AutoSize        =   -1  'True
            Caption         =   "Juros"
            Height          =   195
            Left            =   1290
            TabIndex        =   34
            Top             =   300
            Width           =   375
         End
         Begin VB.Label lblintOcorrencia 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrência"
            Height          =   195
            Left            =   885
            TabIndex        =   33
            Top             =   1350
            Width           =   780
         End
      End
      Begin VB.Frame fra_Dados 
         Caption         =   " Dados da Guia / Banco "
         Height          =   1875
         Left            =   780
         TabIndex        =   27
         Top             =   2790
         Width           =   6765
         Begin VB.TextBox txtCapaDeLote 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1305
            MaxLength       =   21
            TabIndex        =   19
            Top             =   330
            Width           =   1785
         End
         Begin MSDataListLib.DataCombo dbcintBanco 
            Height          =   315
            Left            =   1305
            TabIndex        =   20
            Top             =   660
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintAgencia 
            Height          =   315
            Left            =   1305
            TabIndex        =   21
            Top             =   1020
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintConta 
            Height          =   315
            Left            =   1305
            TabIndex        =   22
            Top             =   1380
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblCapaDeLote 
            AutoSize        =   -1  'True
            Caption         =   "Capa de Lote"
            Height          =   195
            Left            =   270
            TabIndex        =   31
            Top             =   420
            Width           =   960
         End
         Begin VB.Label lblintConta 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   810
            TabIndex        =   30
            Top             =   1500
            Width           =   420
         End
         Begin VB.Label lblintAgencia 
            AutoSize        =   -1  'True
            Caption         =   "Agência"
            Height          =   195
            Left            =   645
            TabIndex        =   29
            Top             =   1140
            Width           =   585
         End
         Begin VB.Label lblintBanco 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   765
            TabIndex        =   28
            Top             =   780
            Width           =   465
         End
      End
      Begin VB.Frame fra_BaixaAutomatica 
         Caption         =   " Arquivo "
         Height          =   1275
         Left            =   780
         TabIndex        =   23
         Top             =   1080
         Width           =   6765
         Begin VB.CommandButton cmd_Arquivo 
            Height          =   285
            Left            =   5760
            Picture         =   "BaixaAutomatica.frx":107A
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Localiza Arquivo de Baixa Automática"
            Top             =   300
            Width           =   360
         End
         Begin VB.TextBox txt_Arquivo 
            Height          =   285
            Left            =   1140
            TabIndex        =   17
            Top             =   300
            Width           =   4635
         End
         Begin MSDataListLib.DataCombo dbc_intLayOut 
            Height          =   315
            Left            =   1140
            TabIndex        =   18
            Top             =   690
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_Arquivo 
            AutoSize        =   -1  'True
            Caption         =   "Localização"
            Height          =   195
            Left            =   195
            TabIndex        =   26
            Top             =   390
            Width           =   855
         End
         Begin VB.Label lbl_LayOut 
            AutoSize        =   -1  'True
            Caption         =   "Layout"
            Height          =   195
            Left            =   570
            TabIndex        =   25
            Top             =   810
            Width           =   480
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgArquivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBaixaAutomatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim mblnPrimeiraVez  As Boolean
Dim vetDescricao(8)  As String
Dim X                As XArrayDB

Dim dblValorTotal    As Double

Private Sub cmd_Arquivo_Click()
    dlgArquivo.CancelError = True
    dlgArquivo.DialogTitle = "Selecione o arquivo"
    dlgArquivo.InitDir = "C:\"
    dlgArquivo.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    dlgArquivo.flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    
    On Error GoTo err_cmd_Arquivo_Click
    
    dlgArquivo.ShowOpen
    txt_Arquivo = dlgArquivo.Filename
    Exit Sub

err_cmd_Arquivo_Click:
    If Err.Number = 32755 Then
        txt_Arquivo = ""
    End If
End Sub

Private Sub dbc_intLayOut_Click(Area As Integer)
    DropDownDataCombo dbc_intLayOut, Me, Area
End Sub

Private Sub dbc_intLayOut_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intLayOut, Me, , KeyCode, Shift
End Sub

Private Sub dbcintAgencia_Click(Area As Integer)
    DropDownDataCombo dbcintAgencia, Me, Area
    If Area = 2 And dbcintAgencia.MatchedWithList Then
        LeDaTabelaParaObj gstrContaBancaria, dbcintConta, strQueryConta
    End If
End Sub

Private Function strQueryConta() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    strSQL = ""
'    strSql = strSql & "SELECT C.PKId, RTRIM(C.strConta) + ' - ' + RTRIM(C.strDigitoVerificador) AS Conta "
    strSQL = strSQL & "SELECT C.PKId, RTRIM(C.strConta) " & strCONCAT & " ' - ' " & strCONCAT & " RTRIM(C.strDigitoVerificador) AS Conta "
    strSQL = strSQL & "FROM " & gstrContaBancaria & " C "
    strSQL = strSQL & "WHERE C.blnContaPublica = 1 "
    If dbcintBanco.MatchedWithList Then
        strSQL = strSQL & "AND C.intBanco = " & dbcintBanco.BoundText & " "
    End If
    If dbcintAgencia.MatchedWithList Then
        strSQL = strSQL & "AND C.intAgencia = " & dbcintAgencia.BoundText & " "
    End If
    strSQL = strSQL & "Order By C.strConta"
    strQueryConta = strSQL
End Function

Private Sub dbcintAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintAgencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintAgencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintAgencia
End Sub

Private Function strQueryAgencia() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT AG.PKId, AG.strDescricao FROM "
    strSQL = strSQL & gstrAgencia & " AG "
    If dbcintBanco.MatchedWithList Then
        strSQL = strSQL & "WHERE AG.intBanco = " & dbcintBanco.BoundText
    End If
    strSQL = strSQL & " ORDER BY AG.strDescricao"
    strQueryAgencia = strSQL
End Function

Private Sub dbcintBanco_Click(Area As Integer)
    DropDownDataCombo dbcintBanco, Me, Area
    If Area = 2 And dbcintBanco.MatchedWithList Then
        LeDaTabelaParaObj gstrAgencia, dbcintAgencia, strQueryAgencia
        dbcintConta.BoundText = ""
        Set dbcintConta.DataSource = Nothing
    End If
End Sub

Private Sub dbcintBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintBanco, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBanco_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintBanco
End Sub

Private Sub dbcintConta_Click(Area As Integer)
    DropDownDataCombo dbcintConta, Me, Area
End Sub

Private Sub dbcintConta_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intLayOut, Me, , KeyCode, Shift
End Sub

Private Sub dbcintConta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintConta
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 678
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrLerArquivo
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrAplicar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrLerArquivo
End Sub

Private Sub Form_Load()
    dbcintBanco.Tag = gstrQueryDataComboBanco & ";strDescricao"
    dbc_intLayOut.Tag = strQueryDataComboLayout & ";strDescricao"
    dbcintOcorrencia.Tag = strQueryOcorrencia & ";O.strDescricao"
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoReceita, strQueryDataComboComposicaoReceita
    
    CarregaDescricaoColuna
    
    TrocaCorObjeto txt_Contribuinte, True
    TrocaCorObjeto dbcintOcorrencia, True
    TrocaCorObjeto dbc_intComposicaoReceita, True
    TrocaCorObjeto mskInscricaoCadastral, True
    TrocaCorObjeto txtIntexercicio, True
    TrocaCorObjeto txt_strSequencia, True
    TrocaCorObjeto txtintNumeroParcela, True
    TrocaCorObjeto txt_dtmDataLancamento, True
    TrocaCorObjeto txt_dtmDataVencimento, True
    TrocaCorObjeto txt_dblValorParcela, True
    TrocaCorObjeto txtDbljuros, True
    TrocaCorObjeto txtDblmulta, True
    TrocaCorObjeto txtDblcorrecao, True
    TrocaCorObjeto txtdblDesconto, True
    TrocaCorObjeto txtdblTotalPago, True
End Sub

Public Function strQueryDataComboComposicaoReceita()
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrComposicaoDaReceita & " "
    strSQL = strSQL & "ORDER BY strDescricao"
    strQueryDataComboComposicaoReceita = strSQL
End Function

Public Function strQueryDataComboLayout()
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrDescricaoLayout & " "
    strSQL = strSQL & "ORDER BY strDescricao"
    strQueryDataComboLayout = strSQL
End Function

Private Function strQueryOcorrencia() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo CAST do SQL Server pela função gstrCONVERT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    strSQL = ""
'    strSql = strSql & "SELECT O.PKID, RTRIM(CAST(O.intCodigo AS CHAR)) + ' - ' + O.strDescricao AS Ocorrencia "
    strSQL = strSQL & "SELECT O.PKID, RTRIM(" & gstrCONVERT(CDT_VARCHAR, "O.intCodigo") & ") " & strCONCAT & " ' - ' " & strCONCAT & " O.strDescricao AS Ocorrencia "
    strSQL = strSQL & "FROM " & gstrOcorrencia & " O "
    strSQL = strSQL & "ORDER BY O.intCodigo"
    strQueryOcorrencia = strSQL
End Function

Private Sub ArquivoBaixa()
    Dim lngLancamentoCalculo   As Long
    Dim strCodParcela          As String
    Dim adoBaixa               As ADODB.Recordset
    Dim strDigitoSeparador     As String
    Dim blnPularHeader         As Boolean
    Dim blnPularTrailer        As Boolean
    Dim adoResultado           As ADODB.Recordset
    Dim vetCampos(8)           As Campo
    Dim strSQL                 As String
    Dim FileNumber
    Dim strLinha               As String
    Dim lngLinha               As Long
    
    Dim i                      As Integer
    Dim varAux                 As Variant

    Dim strInscricaoCadastral  As String
    Dim intExercicio           As Integer
    Dim strSequencia           As String
    Dim intNumeroDaParcela     As Integer
    Dim intComposicaoDaReceita As Integer
    Dim intContribuinte        As Integer
    Dim dtmDataLancamento      As String
    Dim dtmDataVencimento      As String
    Dim dblValorParcela        As Double
    Dim dblJuros               As Double
    Dim dblMulta               As Double
    Dim dblCorrecao            As Double
    Dim dblDesconto            As Double
    Dim dblTotalPago           As Double
    Dim intOcorrencia          As Integer
    Dim dtmDataPagamento       As String
    Dim strNome                As String
    Dim strDataLimiteDesconto  As String
    
    On Error GoTo err_BaixaAutomatica
    
    If Not mblnDadosOK Then
        Exit Sub
    End If
    
    dblValorTotal = 0
    
    strSQL = ""
    strSQL = strSQL & "SELECT C.*, D.strSeparadorColuna, D.blnPularHeader, D.blnPularTrailer "
    strSQL = strSQL & "FROM " & gstrDescricaoLayout & " D, " & gstrLayoutColuna & " C "
    strSQL = strSQL & "WHERE D.pkid = C.intDescricaoLayout "
    strSQL = strSQL & "AND D.PKId = " & dbc_intLayOut.BoundText & " "
    strSQL = strSQL & "ORDER BY intPosicaoCampo"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If Not (.BOF Or .EOF) Then
                strDigitoSeparador = gstrVerificaCampoNulo(!strSeparadorColuna)
                blnPularHeader = Abs(!blnPularHeader)
                blnPularTrailer = Abs(!blnPularTrailer)
                Do While Not .EOF
                    If !intPosicaoCampo = 0 Then
                        ExibeMensagem "Não foi possível completar a leitura do arquivo. O campo '" & vetDescricao(!intDescricaoColuna) & "' possui a posição inicial 0."
                        Exit Sub
                    ElseIf !intTamanhoCampo = 0 Then
                        ExibeMensagem "Não foi possível completar a leitura do arquivo. O campo '" & vetDescricao(!intDescricaoColuna) & "' possui a posição tamanho 0."
                        Exit Sub
                    End If
                    vetCampos(!intDescricaoColuna).blnVirgula = !blnContemVirgula
                    vetCampos(!intDescricaoColuna).strCasasDecimais = IIf(IsNull(!bytPosicaoDaVirgula), 0, !bytPosicaoDaVirgula)
                    vetCampos(!intDescricaoColuna).intPosicao = !intPosicaoCampo
                    vetCampos(!intDescricaoColuna).intTamanho = !intTamanhoCampo
                    vetCampos(!intDescricaoColuna).intTipo = !bytTipoDado
                    vetCampos(!intDescricaoColuna).strDescricao = !intDescricaoColuna
                    .MoveNext
                Loop
                For i = 1 To 8
                    If Trim(vetCampos(i).intPosicao) = 0 Then
                        ExibeMensagem "Não foi possível completar a leitura do arquivo. O campo '" & vetDescricao(i) & "' não foi cadastrado corretamente."
                        Exit Sub
                    End If
                Next i
            Else
                ExibeMensagem "Não há colunas cadastradas para o lay-out selecionado."
                Exit Sub
            End If
        End With
    End If
    
    Set X = New XArrayDB
    
    FileNumber = FreeFile
    lngLinha = 0
    
    Open txt_Arquivo For Input As #FileNumber
    Do While Not EOF(FileNumber)
        Line Input #FileNumber, strLinha
        lngLinha = lngLinha + 1
        
        If lngLinha = 1 Then
            If blnPularHeader = True Then
                GoTo Proximo
            End If
        ElseIf EOF(FileNumber) Then
            If blnPularTrailer Then
                GoTo Proximo
            End If
        End If
        
        strCodParcela = Mid(strLinha, vetCampos(1).intPosicao, vetCampos(1).intTamanho)
        
        'Seleciona os dados da parcela
        strSQL = ""
        strSQL = strSQL & "SELECT C.strNome, LC.PKId, LC.strInscricaoCadastral, LC.intExercicio, LC.intOcorrencia, "
        strSQL = strSQL & "LC.dtmLancamento, LC.strSequencia, LC.intContribuinte, "
        strSQL = strSQL & "PR.intNumeroParcela, PR.intComposicaoDaReceita, PR.dtmDataVencimento, "
        strSQL = strSQL & "PR.dblValorParcela "
        strSQL = strSQL & "FROM " & gstrLancamentoCalculo & " LC, "
        strSQL = strSQL & gstrParcelaReceita & " PR, " & gstrContribuinte & " C "
        strSQL = strSQL & "WHERE LC.PKID = PR.intLancamentoCalculo "
        strSQL = strSQL & "AND LC.intContribuinte = C.PKId "
        strSQL = strSQL & "AND PR.PKId = " & strCodParcela
        Set adoBaixa = New ADODB.Recordset
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoBaixa) Then
            With adoBaixa
                If .BOF Or .EOF Then
                    '????? Incluir parcela na inconsistência
                    GoTo Proximo
                End If
            End With
        Else
            '????? Incluir parcela na inconsistência
            GoTo Proximo
        End If
        Set gobjBanco = Nothing
        
        strDataLimiteDesconto = strTrataData(Mid(strLinha, vetCampos(2).intPosicao, vetCampos(2).intTamanho))
        
        lngLancamentoCalculo = adoBaixa!Pkid
        strInscricaoCadastral = adoBaixa!strInscricaoCadastral
        intExercicio = adoBaixa!intExercicio
        strSequencia = adoBaixa!strSequencia
        intNumeroDaParcela = adoBaixa!intNumeroParcela
        intComposicaoDaReceita = adoBaixa!intComposicaoDaReceita
        intContribuinte = adoBaixa!intContribuinte
        dtmDataLancamento = gstrDataFormatada(adoBaixa!dtmLancamento)
        dtmDataVencimento = gstrDataFormatada(adoBaixa!dtmDataVencimento)
        dblValorParcela = gstrConvVrDoSql(adoBaixa!dblValorParcela)
        strNome = Trim(adoBaixa!strNome)
        
        If vetCampos(3).blnVirgula = True Then
            dblJuros = gstrConvVrDoSql(Mid(strLinha, vetCampos(3).intPosicao, vetCampos(3).intTamanho))
        Else
            dblJuros = CDbl(gstrConvVrDoSql(Mid(strLinha, vetCampos(3).intPosicao, vetCampos(3).intTamanho))) / 10 ^ vetCampos(3).strCasasDecimais
        End If
        
        If vetCampos(4).blnVirgula = True Then
            dblMulta = gstrConvVrDoSql(Mid(strLinha, vetCampos(4).intPosicao, vetCampos(4).intTamanho))
        Else
            dblMulta = CDbl(gstrConvVrDoSql(Mid(strLinha, vetCampos(4).intPosicao, vetCampos(4).intTamanho))) / 10 ^ vetCampos(4).strCasasDecimais
        End If
        
        If vetCampos(5).blnVirgula = True Then
            dblCorrecao = gstrConvVrDoSql(Mid(strLinha, vetCampos(5).intPosicao, vetCampos(5).intTamanho))
        Else
            dblCorrecao = CDbl(gstrConvVrDoSql(Mid(strLinha, vetCampos(5).intPosicao, vetCampos(5).intTamanho))) / 10 ^ vetCampos(5).strCasasDecimais
        End If
        
        If vetCampos(6).blnVirgula = True Then
            dblDesconto = gstrConvVrDoSql(Mid(strLinha, vetCampos(6).intPosicao, vetCampos(6).intTamanho))
        Else
            dblDesconto = CDbl(gstrConvVrDoSql(Mid(strLinha, vetCampos(6).intPosicao, vetCampos(6).intTamanho))) / 10 ^ vetCampos(6).strCasasDecimais
        End If
        
        If vetCampos(7).blnVirgula = True Then
            dblTotalPago = Mid(strLinha, vetCampos(7).intPosicao, vetCampos(7).intTamanho)
        Else
            dblTotalPago = CDbl(gstrConvVrDoSql(Mid(strLinha, vetCampos(7).intPosicao, vetCampos(7).intTamanho))) / 10 ^ vetCampos(7).strCasasDecimais
        End If
        
        dblValorTotal = dblValorTotal + dblTotalPago
        
        intOcorrencia = adoBaixa!intOcorrencia
        
        dtmDataPagamento = strTrataData(Mid(strLinha, vetCampos(8).intPosicao, vetCampos(8).intTamanho))

        i = X.Count(1)
        X.ReDim 0, i, 0, 17
        
        varAux = strInscricaoCadastral
        X(i, 0) = varAux
        
        varAux = intExercicio
        X(i, 1) = varAux
        
        varAux = strSequencia
        X(i, 2) = varAux
        
        varAux = intNumeroDaParcela
        X(i, 3) = varAux
        
        varAux = intComposicaoDaReceita
        X(i, 4) = varAux
        
        varAux = intContribuinte
        X(i, 5) = varAux
        
        varAux = dtmDataLancamento
        X(i, 6) = varAux
        
        varAux = dtmDataVencimento
        X(i, 7) = varAux
        
        varAux = gvntConvVrDoSql(dblValorParcela)
        X(i, 8) = varAux
        
        varAux = gvntConvVrDoSql(dblJuros)
        X(i, 9) = varAux

        varAux = gvntConvVrDoSql(dblMulta)
        X(i, 10) = varAux

        varAux = gvntConvVrDoSql(dblCorrecao)
        X(i, 11) = varAux

        varAux = gvntConvVrDoSql(dblDesconto)
        X(i, 12) = varAux

        varAux = gvntConvVrDoSql(dblTotalPago)
        X(i, 13) = varAux

        varAux = intOcorrencia
        X(i, 14) = varAux

        varAux = dtmDataPagamento
        X(i, 15) = varAux

        varAux = lngLancamentoCalculo
        X(i, 16) = varAux
    
        varAux = strNome
        X(i, 17) = varAux
        
Proximo:
        Set adoBaixa = Nothing
    Loop
    
    Set tdb_Baixa.Array = X
    tdb_Baixa.ReBind
    tdb_Baixa.Refresh

    Close #FreeFile
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Erro desconhecido na leitura do arquivo."
End Sub

Private Function mblnDadosOK() As Boolean
    On Error GoTo err_mblnDadosOK
    If Trim(txt_Arquivo) = "" Then
        ExibeMensagem "Indique a localização do arquivo de retorno."
        Exit Function
    ElseIf Dir(txt_Arquivo) = "" Then
        ExibeMensagem "Arquivo não encontrado no local especificado."
        Exit Function
    ElseIf Trim(dbc_intLayOut.BoundText) = "" Then
        ExibeMensagem "Selecione o lay-out do arquivo de retorno."
        Exit Function
    ElseIf Trim(txtCapaDeLote) = "" Then
        ExibeMensagem "Informe a capa de lote."
        Exit Function
    ElseIf dbcintBanco.MatchedWithList = False Then
        ExibeMensagem "Selecione o banco."
        Exit Function
    ElseIf dbcintAgencia.MatchedWithList = False Then
        ExibeMensagem "Selecione a agência."
        Exit Function
    ElseIf dbcintConta.MatchedWithList = False Then
        ExibeMensagem "Selecione a conta."
        Exit Function
    End If
    
    mblnDadosOK = True
    Exit Function
    
err_mblnDadosOK:
    mblnDadosOK = False
End Function

Private Function strTrataData(strData As String) As String
    If Len(strData) = 6 Then
        strTrataData = gstrDataFormatada(Mid(strData, 1, 2) & "/" & Mid(strData, 3, 2) & "/" & Mid(strData, 5, 2))
    Else
        strTrataData = gstrDataFormatada(Mid(strData, 1, 2) & "/" & Mid(strData, 3, 2) & "/" & Mid(strData, 5, 4))
    End If
End Function

Private Sub CarregaDescricaoColuna()
    vetDescricao(1) = "Código da Parcela"
    vetDescricao(2) = "Data Limite para Desconto"
    vetDescricao(3) = "Juros"
    vetDescricao(4) = "Multa"
    vetDescricao(5) = "Correção"
    vetDescricao(6) = "Desconto"
    vetDescricao(7) = "Total Pago"
    vetDescricao(8) = "Data de Pagamento"
End Sub

Private Function GravaPagamentos() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL    As String
    Dim i         As Integer
    Dim intLinhas As Integer
    Dim strQueryParcelaZero As String
    On Error GoTo err_GravaPagamentos
    
    If MsgBox("Confirma gravação do lançamentos?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    Set gobjBanco = New clsBanco
    
    intLinhas = X.Count(1) - 1
    
    For i = 0 To intLinhas
        gobjBanco.ExecutaBeginTrans
        
        strQueryParcelaZero = ""
        Select Case Val(gstrConvVrParaSql(X(i, 3)))  'Número da parcela
            Case 0 'Deleta todas as parcelas
            Case Else 'Deleta somente a parcela selecionada e a parcela 0, caso exista
                strQueryParcelaZero = strQueryParcelaZero & " AND (intNumeroParcela = " & gstrConvVrParaSql(X(i, 3))
                strQueryParcelaZero = strQueryParcelaZero & " OR intNumeroParcela = 0)"
        End Select
        
        'Deleta as registros da tabela tblParcelaReceita
        strSQL = ""
        strSQL = strSQL & " DELETE FROM " & gstrParcelaReceita
        strSQL = strSQL & " WHERE intLancamentoCalculo = " & X(i, 16)
        strSQL = strSQL & strQueryParcelaZero
        If Not gobjBanco.Execute(strSQL, False) Then
            gobjBanco.ExecutaRollbackTrans
            GoTo Proximo
        End If
        
        'Deleta as registros da tabela tblParcelaTaxa
        strSQL = ""
        strSQL = strSQL & " DELETE FROM " & gstrParcelaTaxa
        strSQL = strSQL & " WHERE intLancamentoCalculo = " & X(i, 16)
        strSQL = strSQL & strQueryParcelaZero
        If Not gobjBanco.Execute(strSQL, False) Then
            gobjBanco.ExecutaRollbackTrans
            GoTo Proximo
        End If
                
        'Grava o pagamento da parcela
        strSQL = ""
        strSQL = strSQL & " INSERT INTO " & gstrPagamentoParcela
        strSQL = strSQL & " (strInscricaoCadastral, intExercicio, strSequencia, intNumeroDaParcela, intComposicaoDaReceita,"
        strSQL = strSQL & " intContribuinte, dtmDataLancamento, dtmDataVencimento, dblValorParcela, dblJuros, dblMulta, "
        strSQL = strSQL & " dblCorrecao, dblDesconto, dblTotalPago, intOcorrencia, dtmDataPagamento, "
        strSQL = strSQL & " intBanco, intAgencia, intContaBancaria, dtmDtAtualizacao, lngCodUsr) VALUES ( "
        
        'Inscrição
        strSQL = strSQL & gstrConvVrParaSql(X(i, 0))
        'Exercício
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 1))
        'Sequência
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 2))
        'Número da parcela
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 3))
        'Composição da receita
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 4))
        'Contribuinte
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 5))
        'Data de Lançamento
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 6))
        'Data de vencimento
        strSQL = strSQL & "," & gstrConvDtParaSql(X(i, 7))
        'Valor da Parcela
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 8))
        'Juros
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 9))
        'Multa
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 10))
        'Correção
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 11))
        'Desconto
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 12))
        'Total pago
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 13))
        'Ocorrência
        strSQL = strSQL & "," & gstrConvVrParaSql(X(i, 14))
        'Data do Pagamento
        strSQL = strSQL & "," & gstrConvDtParaSql(X(i, 15))
        
        'Banco
        strSQL = strSQL & "," & gstrConvVrParaSql(dbcintBanco.BoundText)
        'Agencia
        strSQL = strSQL & "," & gstrConvVrParaSql(dbcintAgencia.BoundText)
        'Conta Bancária
        strSQL = strSQL & "," & gstrConvVrParaSql(dbcintConta.BoundText)
        
        'Atualização da tabela
'        strSql = strSql & ", GETDATE()"
        strSQL = strSQL & ", " & strGETDATE
        'Usuário
        strSQL = strSQL & "," & glngCodUsr
        strSQL = strSQL & ")"
        
        If gobjBanco.Execute(strSQL, False) Then
            gobjBanco.ExecutaCommitTrans
        Else
            gobjBanco.ExecutaRollbackTrans
        End If
Proximo:
    Next i
    
    Screen.MousePointer = 0
    ExibeMensagem "Operação realizada com sucesso."
    Exit Function
    
err_GravaPagamentos:
    Screen.MousePointer = 0
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrLerArquivo
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_Baixa_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_Baixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Baixa
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                CarregaDadosParcela
                mblnPrimeiraVez = False
            End If
        End If
    End With
End Sub

Private Sub CarregaDadosParcela()
    On Error GoTo err_CarregaDadosParcela
    txt_Contribuinte = tdb_Baixa.Columns("Nome")
    mskInscricaoCadastral = tdb_Baixa.Columns("Inscrição Cadastral")
    txtIntexercicio = tdb_Baixa.Columns("Exercício")
    txt_strSequencia = tdb_Baixa.Columns("Sequência")
    txtintNumeroParcela = tdb_Baixa.Columns("Número Parcela")
    dbc_intComposicaoReceita.BoundText = tdb_Baixa.Columns("Composição da Receita")
    txt_dtmDataLancamento = tdb_Baixa.Columns("Lançamento")
    txt_dtmDataVencimento = tdb_Baixa.Columns("Vencimento")
    txt_dblValorParcela = tdb_Baixa.Columns("Valor Parcela")
    txtDbljuros = tdb_Baixa.Columns("Juros")
    txtDblmulta = tdb_Baixa.Columns("Multa")
    txtDblcorrecao = tdb_Baixa.Columns("Correção")
    txtdblDesconto = tdb_Baixa.Columns("Desconto")
    txtdblTotalPago = tdb_Baixa.Columns("Total Pago")
    dbcintOcorrencia.BoundText = tdb_Baixa.Columns("Ocorrência")
err_CarregaDadosParcela:
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case strModoOperacao
        Case gstrLerArquivo
            ArquivoBaixa
            
        Case gstrSalvar
            If VerificaLancamento Then
                GravaPagamentos
            End If
            
        Case gstrPreencherLista
            dbcintAgencia.Tag = strQueryAgencia & ";strDescricao"
            dbcintConta.Tag = strQueryConta & ";C.strConta"
            
            PreencherListaDeOpcoes Me.ActiveControl
            
        Case gstrFechar
            Unload Me
    End Select
End Sub

Private Sub txtCapaDeLote_GotFocus()
    MarcaCampo txtCapaDeLote
End Sub

Private Sub txtCapaDeLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtCapaDeLote
End Sub

Private Function VerificaLancamento() As Boolean
    On Error Resume Next
    If CDbl(txtCapaDeLote.Text) <> dblValorTotal Then
        ExibeMensagem "Os valores lançados não fecham com o valor da capa de lote!"
        VerificaLancamento = False
    Else
        VerificaLancamento = True
    End If
End Function

Private Sub txtCapaDeLote_LostFocus()
    txtCapaDeLote = gvntConvVrDoSql(txtCapaDeLote)
End Sub
