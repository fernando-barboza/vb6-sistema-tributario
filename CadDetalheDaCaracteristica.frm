VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadDetalheDaCaracteristica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Características de Boletins"
   ClientHeight    =   6825
   ClientLeft      =   2925
   ClientTop       =   3705
   ClientWidth     =   7845
   HelpContextID   =   42
   Icon            =   "CadDetalheDaCaracteristica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7845
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6675
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   11774
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Dados"
      TabPicture(0)   =   "CadDetalheDaCaracteristica.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintCaracteristica"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Detalhe"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintUtilizacao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintReferenciaTributo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintCategoriaConstrucao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintCategoriaConstrucao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbcintReferenciaTributo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "grd_Detalhes"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dbcintUtilizacaoDaCaracteristica"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbointCaracteristica"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_Caracteristica"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "tdd_Detalhes"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin TrueOleDBGrid70.TDBDropDown tdd_Detalhes 
         Height          =   1770
         Left            =   930
         TabIndex        =   12
         Top             =   3570
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   3122
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
         Columns(1).Caption=   "Valor"
         Columns(1).DataField=   "dblvalor"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nome"
         Columns(2).DataField=   "strnomedovalor"
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
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1667"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1588"
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
         DataMode        =   4
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowDividerStyle =   2
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   0   'False
         ListField       =   "dblValor"
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H8000000E&"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Named:id=33:Normal"
         _StyleDefs(50)  =   ":id=33,.parent=0"
         _StyleDefs(51)  =   "Named:id=34:Heading"
         _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   ":id=34,.wraptext=-1"
         _StyleDefs(54)  =   "Named:id=35:Footing"
         _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   "Named:id=36:Selected"
         _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=37:Caption"
         _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(60)  =   "Named:id=38:HighlightRow"
         _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton cmd_Caracteristica 
         Height          =   315
         Left            =   6075
         Picture         =   "CadDetalheDaCaracteristica.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ativa cadastro de Características Gerais"
         Top             =   1245
         Width           =   360
      End
      Begin MSDataListLib.DataCombo cbointCaracteristica 
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Top             =   1260
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintUtilizacaoDaCaracteristica 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   540
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid grd_Detalhes 
         Height          =   4050
         Left            =   180
         TabIndex        =   11
         Top             =   2445
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   7144
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nome"
         Columns(1).DataField=   ""
         Columns(1).DataWidth=   50
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   ""
         Columns(2).DropDown=   "tdd_Detalhes"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "PKId"
         Columns(3).DataField=   ""
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1270"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1191"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=6429"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6350"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4366"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4286"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(2).AutoCompletion=1"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
         TabAction       =   1
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=116,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(54)  =   "Named:id=33:Normal"
         _StyleDefs(55)  =   ":id=33,.parent=0"
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
         _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=39:EvenRow"
         _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(69)  =   "Named:id=40:OddRow"
         _StyleDefs(70)  =   ":id=40,.parent=33"
         _StyleDefs(71)  =   "Named:id=41:RecordSelector"
         _StyleDefs(72)  =   ":id=41,.parent=34"
         _StyleDefs(73)  =   "Named:id=42:FilterBar"
         _StyleDefs(74)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintReferenciaTributo 
         Height          =   315
         Left            =   1260
         TabIndex        =   9
         Top             =   1635
         Visible         =   0   'False
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintCategoriaConstrucao 
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Top             =   900
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Label lblintCategoriaConstrucao 
         AutoSize        =   -1  'True
         Caption         =   "Categoria"
         Height          =   195
         Left            =   465
         TabIndex        =   3
         Top             =   1005
         Width           =   675
      End
      Begin VB.Label lblintReferenciaTributo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   810
         TabIndex        =   8
         Top             =   1755
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblintUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   450
         TabIndex        =   1
         Top             =   630
         Width           =   690
      End
      Begin VB.Label lbl_Detalhe 
         AutoSize        =   -1  'True
         Caption         =   "Detalhes:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label lblintCaracteristica 
         AutoSize        =   -1  'True
         Caption         =   "Característica"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   1365
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmCadDetalheDaCaracteristica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando             As Boolean
Dim mobjAux                   As Object
Dim adoResultado              As ADODB.Recordset
Dim strSQL                    As String
Dim intUtilizacao             As Integer
Dim strExclusao               As String
    
Dim mblnSelecionou            As Boolean
Dim mblnPrimeiraVez           As Boolean
    
Dim adoRec                    As ADODB.Recordset
Dim adoTdb                    As ADODB.Recordset
Dim x                         As XArrayDB
Dim y                         As New XArrayDB

Private Sub cbointCaracteristica_Change()

    If cbointCaracteristica.MatchedWithList Then
        cbointCaracteristica_Click 2
    End If

End Sub

Private Sub cbointCaracteristica_Click(Area As Integer)
   
   intUtilizacao = 0
   
   If Area = 2 Then
       
       If cbointCaracteristica.MatchedWithList Then
           
           LimpaGrid
           
           strSQL = ""
           strSQL = strSQL & "SELECT DC.intCodigoDoDetalhe, DC.strNomeDoDetalhe, "
           strSQL = strSQL & "DC.intTabelaDeValores, DC.PKId, DC.intReferenciaTributo "
           strSQL = strSQL & "FROM " & gstrDetalheDaCaracteristica & " DC "
           strSQL = strSQL & "WHERE DC.intCaracteristica = " & cbointCaracteristica.BoundText & " "
           strSQL = strSQL & "ORDER BY DC.intCodigoDoDetalhe"
            
           Set gobjBanco = New clsBanco
           gobjBanco.CriaADO strSQL, 5, adoRec
           strSQL = ""
           strSQL = strSQL & "SELECT intUtilizacaoDaCaracteristica "
           strSQL = strSQL & "FROM " & gstrCaracteristicaGeral & " "
           strSQL = strSQL & "WHERE PKId = " & cbointCaracteristica.BoundText
           Set gobjBanco = New clsBanco
           If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
               If Not adoResultado.EOF Then
                   intUtilizacao = adoResultado!intUtilizacaoDaCaracteristica
               End If
           End If
                      
           If Not adoRec.EOF Then
               dbcintReferenciaTributo.BoundText = IIf(adoRec("intReferenciaTributo").Value > 0, adoRec("intReferenciaTributo").Value, "")
           End If
           
           MontaArray

           strSQL = ""
           strSQL = strSQL & "SELECT pkiD, strnomedovalor, dblvalor "
           strSQL = strSQL & "FROM " & gstrTabelaDeValor & " "
           strSQL = strSQL & "WHERE intCodigoDaUtilizacao = " & intUtilizacao & " "
           strSQL = strSQL & "ORDER BY strnomedovalor"
           gobjBanco.CriaADO strSQL, 5, adoTdb

           y.ReDim 0, adoTdb.RecordCount - 1, 0, 2
           Dim varAux As Variant
           Do While Not adoTdb.EOF
               varAux = adoTdb!Pkid
               y(adoTdb.AbsolutePosition - 1, 0) = varAux

               varAux = adoTdb!DBLVALOR
               y(adoTdb.AbsolutePosition - 1, 1) = varAux

               varAux = adoTdb!strnomedovalor
               y(adoTdb.AbsolutePosition - 1, 2) = varAux

               adoTdb.MoveNext
           Loop
           
           Set tdd_Detalhes.Array = y
           tdd_Detalhes.Rebind
           tdd_Detalhes.Refresh
           
           grd_Detalhes.SetFocus
       Else
           LimpaGrid
       End If
   
   End If
   
End Sub

Private Sub cbointCaracteristica_KeyPress(KeyAscii As Integer)
Dim intPkid As Integer
    
    If KeyAscii = vbKeyReturn Then
        If Not cbointCaracteristica.MatchedWithList Then
            strSQL = ""
            strSQL = strSQL & "SELECT PKId "
            strSQL = strSQL & "FROM " & gstrCaracteristicaGeral & " "
            strSQL = strSQL & "WHERE intCodigoDaCaracteristica = " & Val(cbointCaracteristica.Text)
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If Not adoResultado.EOF Then
                    intPkid = adoResultado!Pkid
                End If
            End If

            cbointCaracteristica.BoundText = CStr(intPkid)
            If cbointCaracteristica.MatchedWithList Then
                cbointCaracteristica_Click 2
                grd_Detalhes.SetFocus
            Else
                LimpaGrid
                cbointCaracteristica.BoundText = ""
            End If
        Else
            cbointCaracteristica_Click 2
        End If
    End If
    
End Sub


Private Sub dbcintUtilizacaoDaCaracteristica_Change()

    'So vamos exibir a combo de tipo caso seja Imobiliario Terreno
    If dbcintUtilizacaoDaCaracteristica.MatchedWithList Then
       dbcintReferenciaTributo.Visible = dbcintUtilizacaoDaCaracteristica.BoundText = 2
       lblintReferenciaTributo.Visible = dbcintUtilizacaoDaCaracteristica.BoundText = 2
    Else
       LimpaGrid
       dbcintReferenciaTributo.Visible = False
       lblintReferenciaTributo.Visible = False
       dbcintCategoriaConstrucao.BoundText = ""
       Set dbcintCategoriaConstrucao.RowSource = Nothing
       cbointCaracteristica.BoundText = ""
       Set cbointCaracteristica.RowSource = Nothing
    End If
    
End Sub

Private Sub dbcintUtilizacaoDaCaracteristica_Click(Area As Integer)
   
   DropDownDataCombo dbcintUtilizacaoDaCaracteristica, Me, Area
   
   If Area = 2 And dbcintUtilizacaoDaCaracteristica.MatchedWithList Then
       
       dbcintCategoriaConstrucao.Tag = strQueryCategoriaConstrucao & ";strDescricao"
       LeDaTabelaParaObj gstrCategoriaConstrucao, dbcintCategoriaConstrucao, strQueryCategoriaConstrucao
       mblnAlterando = True
       
   End If
   
End Sub

Private Sub dbcintUtilizacaoDaCaracteristica_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUtilizacaoDaCaracteristica, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUtilizacaoDaCaracteristica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintUtilizacaoDaCaracteristica
End Sub
    
Private Sub dbcintCategoriaConstrucao_Change()

    If Not dbcintCategoriaConstrucao.MatchedWithList Then
       LimpaGrid
       cbointCaracteristica.BoundText = ""
       Set cbointCaracteristica.RowSource = Nothing
    End If

End Sub
    
Private Sub dbcintCategoriaConstrucao_Click(Area As Integer)
   
   DropDownDataCombo dbcintCategoriaConstrucao, Me, Area
   
   If Area = 2 And dbcintCategoriaConstrucao.MatchedWithList Then
       
       cbointCaracteristica.Tag = strQueryCaracteristicas & ";strNomeDaCaracteristica"
       LeDaTabelaParaObj gstrCaracteristicaGeral, cbointCaracteristica, strQueryCaracteristicas
       mblnAlterando = True
       
   End If
   
End Sub

Private Sub dbcintCategoriaConstrucao_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintCategoriaConstrucao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCategoriaConstrucao_KeyPress(KeyAscii As Integer)
 CaracterValido KeyAscii, "A", dbcintCategoriaConstrucao
End Sub
    
Private Sub cmd_Caracteristica_Click()
    
    If Not dbcintUtilizacaoDaCaracteristica.MatchedWithList Then
        ExibeMensagem "Selecione a utilização."
        Exit Sub
    Else
        If Not dbcintCategoriaConstrucao.MatchedWithList Then
            ExibeMensagem "Selecione a Categoria"
            Exit Sub
        End If
    End If
    
    'PreencherListaDeOpcoes frmCadCaracteristicasGerais.dbcintUtilizacaoDaCaracteristica, dbcintUtilizacaoDaCaracteristica.BoundText
    'frmCadCaracteristicasGerais.dbcintCategoriaConstrucao.Text = dbcintCategoriaConstrucao.Text
    'TrocaCorObjeto frmCadCaracteristicasGerais.dbcintUtilizacaoDaCaracteristica, True
    'frmCadCaracteristicasGerais.MantemForm gstrLocalizar
    'frmCadCaracteristicasGerais.MantemForm (gstrPreencherLista)
    
    'CarregaForm frmCadCaracteristicasGerais
    
    
    ChamaFormCadastro frmCadCaracteristicasGerais, cbointCaracteristica
    VerificaFormAtivo = True
    'CarregaForm frmCadCaracteristicasGerais
    
    
    
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 598
    VirificaGradeListView Me
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
End Sub

Private Sub Form_Deactivate()
    'HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftDown, AltDown, CtrlDown
    
    Select Case KeyCode
        Case vbKeyEscape
            If Not IsNull(tdd_Detalhes.SelectedItem) Then
                grd_Detalhes.SelStart = Len(grd_Detalhes.Text)
            End If
            SendKeys "{RIGHT}"
            Exit Sub
    End Select
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    LeDaTabelaParaObj "", dbcintUtilizacaoDaCaracteristica, strQueryDataComboUtilizacaoDaTabelaDeValor
    LeDaTabelaParaObj "", dbcintReferenciaTributo, strQueryDataComboReferenciaTributo
    dbcintUtilizacaoDaCaracteristica.Tag = strQueryDataComboUtilizacaoDaTabelaDeValor & ";strNomeDaUtilizacao"
    VerificaObjParaAplicar mobjAux
    VerificaFormAtivo = False
End Sub

Private Function strQueryDataComboUtilizacaoDaTabelaDeValor() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strNomeDaUtilizacao "
    strSQL = strSQL & "FROM " & gstrUtilizacaoDaTabelaDeValor & " "
    strSQL = strSQL & "ORDER BY strNomeDaUtilizacao"
    strQueryDataComboUtilizacaoDaTabelaDeValor = strSQL
End Function

Private Function strQueryDataComboReferenciaTributo() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrReferenciasDeTributos & " "
    strSQL = strSQL & "WHERE bytExibir = 1 AND bytGrupo = " & GRUPO_IMOB_TERRENO & " "
    strSQL = strSQL & "ORDER BY strDescricao"
    strQueryDataComboReferenciaTributo = strSQL
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
VerificaFormAtivo = False
End Sub

Private Sub grd_Detalhes_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If ColIndex = 2 Then
        If Trim(grd_Detalhes.Columns(ColIndex).Text) = 0 Or Trim(grd_Detalhes.Columns(ColIndex).Text) = "" Or blnValorCadastrado(grd_Detalhes.Columns(ColIndex).Text) = True Then
            Exit Sub
        Else
            ExibeMensagem "Valor não cadastrado na tabela de valores."
            grd_Detalhes.Columns(ColIndex).Text = ""
            Cancel = True
        End If
    End If
End Sub

Private Sub grd_Detalhes_BeforeDelete(Cancel As Integer)
    If Val(grd_Detalhes.Columns(3).Value) <> 0 Then
        strExclusao = strExclusao & grd_Detalhes.Columns(3).Value & ", "
    End If
End Sub

Private Sub grd_Detalhes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    Else
        Select Case grd_Detalhes.Col
            Case 0
                CaracterValido KeyAscii, "N", grd_Detalhes
            Case 1
                CaracterValido KeyAscii, "A", grd_Detalhes
            Case 2
                If KeyAscii = vbKeyReturn Then
                    KeyAscii = 0
                    SendKeys "{TAB}"
                End If
                CaracterValido KeyAscii, "V", grd_Detalhes
        End Select
    End If
End Sub


Private Sub tdd_Detalhes_DropDownClose()
Dim intRow As Integer
    On Error GoTo Err_Handle
    With grd_Detalhes
        intRow = .Row + 1
        If .Row = (x.Count(1)) Then
            .MoveFirst
        End If
        .Row = intRow
        .Col = 0
        .SetFocus
    End With
Err_Handle:
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case UCase(strModoOperacao)
        Case "NOVO"
            LimpaGrid
            LimpaTela
            Exit Sub
        Case "SALVAR"
            If blnDadosOk Then
                If GravaValores = True Then
                   cbointCaracteristica_Click 2
                   LimpaTela
                End If
            End If
        Case "DELETAR"
                If DeletaValores = True Then
                   cbointCaracteristica_Click 2
                   LimpaTela
                End If
        Case gstrLocalizar
            Exit Sub
        Case gstrPreencherLista
            PreencherListaDeOpcoes Me.ActiveControl
            Exit Sub
        Case Else
            ToolBarGeral strModoOperacao, gstrDetalheDaCaracteristica, mblnAlterando, grd_Detalhes, Me, mobjAux, strSQL, , rptcadDetalhedaCaracteristica, strQueryRelatorio
            
    End Select
    
End Sub

Private Sub LimpaTela()
    dbcintCategoriaConstrucao.BoundText = ""
    cbointCaracteristica.BoundText = ""
    Set cbointCaracteristica.RowSource = Nothing
    dbcintReferenciaTributo.BoundText = ""
    mblnAlterando = False
    strExclusao = ""
    dbcintUtilizacaoDaCaracteristica.SetFocus
End Sub

Private Function strQueryRelatorio() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " SELECT CG.PKId, CG.intCodigoDaCaracteristica, CG.strNomeDaCaracteristica, DC.intCodigoDoDetalhe, "
    strSQL = strSQL & " DC.strNomeDoDetalhe, TV.bytTipoDoValor, TV.dblValor, UV.PKId AS intUtil, UV.strNomeDaUtilizacao"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrCaracteristicaGeral & " CG, "
    strSQL = strSQL & gstrUtilizacaoDaTabelaDeValor & " UV, "
    strSQL = strSQL & gstrDetalheDaCaracteristica & " DC, "
    strSQL = strSQL & gstrTabelaDeValor & " TV "
    strSQL = strSQL & " WHERE UV.PKId = CG.intUtilizacaoDaCaracteristica"
    strSQL = strSQL & " AND CG.PKId = DC.intCaracteristica"
    strSQL = strSQL & " AND TV.PKId = DC.intTabelaDeValores"
    strSQL = strSQL & " AND UV.PKId = TV.intCodigoDaUtilizacao"
    strSQL = strSQL & " AND bytCaracteristica = 1"
    If dbcintUtilizacaoDaCaracteristica.MatchedWithList Then
        strSQL = strSQL & " AND intUtilizacaoDaCaracteristica = " & dbcintUtilizacaoDaCaracteristica.BoundText
        If cbointCaracteristica.MatchedWithList Then
            strSQL = strSQL & " AND CG.PKId = " & cbointCaracteristica.BoundText
        End If
    End If
    strSQL = strSQL & " ORDER BY CG.intCodigoDaCaracteristica, CG.strNomeDaCaracteristica, DC.intCodigoDoDetalhe, DC.strNomeDoDetalhe"
   
    strQueryRelatorio = strSQL
    
End Function

Private Sub MontaArray()
Dim varAux As Variant

    Set x = New XArrayDB
    x.Clear

    With adoRec
        If Not .EOF Then
            x.ReDim 0, .RecordCount - 1, 0, 3
            Do While Not .EOF
                varAux = .Fields(0)
                x(.AbsolutePosition - 1, 0) = varAux
                varAux = .Fields(1)
                x(.AbsolutePosition - 1, 1) = varAux
                varAux = .Fields(2)
                x(.AbsolutePosition - 1, 2) = varAux
                varAux = .Fields(3)
                x(.AbsolutePosition - 1, 3) = varAux
                .MoveNext
            Loop
        Else
            x.ReDim 0, 0, 0, 3
            x(0, 0) = ""
            x(0, 1) = ""
            x(0, 2) = ""
            x(0, 3) = ""
        End If
    End With

    Set grd_Detalhes.Array = x
    grd_Detalhes.Rebind
    grd_Detalhes.Refresh
    
    strExclusao = ""
    
End Sub

Private Function DeletaValores() As Boolean
Dim strSQL As String

    If cbointCaracteristica.BoundText = "" Then
        ExibeMensagem "A característica tem que ser selecionada."
        Exit Function
    End If

    If MsgBox("Confirma exclusão do Item Selecionado?", vbQuestion + vbYesNo) = vbYes Then
        strSQL = ""
        strSQL = strSQL & "DELETE FROM " & gstrDetalheDaCaracteristica & " "
        strSQL = strSQL & "WHERE Pkid = " & grd_Detalhes.Columns(3) & " "

        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        
        If Not gobjBanco.Execute(strSQL) Then
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
        
        gobjBanco.ExecutaCommitTrans
        DeletaValores = True
        
        'X.Clear
        'X.ReDim 0, 0, 0, 3
                
        'Set grd_Detalhes.Array = X
        'grd_Detalhes.Rebind
        'grd_Detalhes.Refresh
        
'        LimpaGrid
'        dbcintUtilizacaoDaCaracteristica.BoundText = ""
'        cbointCaracteristica.BoundText = ""
'        Set cbointCaracteristica.RowSource = Nothing
    End If
End Function

Private Function GravaValores() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSQL As String
Dim strMsg As String
Dim i      As Integer

    strMsg = "Confirma gravação destes detalhes?"

    If gblnExclusaoGravacaoOk("", strMsg, True) Then
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans

        If Len(strExclusao) <> 0 Then
            strSQL = ""
            strSQL = strSQL & "DELETE FROM " & gstrDetalheDaCaracteristica & " "
            strSQL = strSQL & "WHERE PKID IN (" & Mid(strExclusao, 1, Len(strExclusao) - 2) & ")"
    
            If Not gobjBanco.Execute(strSQL) Then
                gobjBanco.ExecutaRollbackTrans
                Exit Function
            End If
        End If
        
        If dbcintReferenciaTributo.MatchedWithList Then
            If blnTipoJaCadastrado(dbcintReferenciaTributo.BoundText) Then
                ExibeMensagem "Este Tipo já está relacionado."
                gobjBanco.ExecutaRollbackTrans
                Exit Function
            End If
        End If
        
        For i = 0 To x.Count(1) - 1
            If Val(x(i, 3)) = 0 Then    'Inclusão
                strSQL = ""
                strSQL = strSQL & "INSERT INTO " & gstrDetalheDaCaracteristica & " "
                strSQL = strSQL & "(intCaracteristica, intCodigoDoDetalhe, strNomeDoDetalhe, "
                strSQL = strSQL & "intTabelaDeValores, intReferenciaTributo, dtmDtAtualizacao, lngCodUsr"
                strSQL = strSQL & ") Values ("
                strSQL = strSQL & cbointCaracteristica.BoundText & ", "
                strSQL = strSQL & gstrConvVrParaSql(x(i, 0)) & ", '"
                strSQL = strSQL & gstrConvVrParaSql(x(i, 1)) & "', "
    
                If x(i, 2) = "" Or IsNull(x(i, 2) = "") Or x(i, 2) = Empty Or x(i, 2) = "0" Then
                    strSQL = strSQL & "NULL, "
                Else
                    strSQL = strSQL & x(i, 2) & ", "
                End If
                strSQL = strSQL & gstrENulo(dbcintReferenciaTributo.BoundText, , True) & ", "
                strSQL = strSQL & strGETDATE & ", "
                strSQL = strSQL & glngCodUsr
                strSQL = strSQL & ")"
    
                If Not gobjBanco.Execute(strSQL, False) Then
                    gobjBanco.ExecutaRollbackTrans
                    Exit Function
                End If
                
            Else    'Alteraçao
                strSQL = ""
                strSQL = strSQL & "UPDATE " & gstrDetalheDaCaracteristica & " SET "
                strSQL = strSQL & "intCaracteristica = " & cbointCaracteristica.BoundText & ", "
                strSQL = strSQL & "intCodigoDoDetalhe = " & gstrConvVrParaSql(x(i, 0)) & ", "
                strSQL = strSQL & "strNomeDoDetalhe = '" & gstrConvVrParaSql(x(i, 1)) & "', "
                strSQL = strSQL & "intReferenciaTributo = " & gstrENulo(dbcintReferenciaTributo.BoundText, , True) & ", "
                
                If x(i, 2) = "" Or IsNull(x(i, 2) = "") Or x(i, 2) = Empty Or x(i, 2) = "0" Then
                    strSQL = strSQL & "intTabelaDeValores = NULL, "
                Else
                    strSQL = strSQL & "intTabelaDeValores = " & x(i, 2) & ", "
                End If
                
'                strSql = strSql & "dtmDtAtualizacao = GETDATE(), "
                strSQL = strSQL & "dtmDtAtualizacao = " & strGETDATE & ", "
                strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
                strSQL = strSQL & "WHERE PKId = " & x(i, 3)
    
                If Not gobjBanco.Execute(strSQL, False) Then
                    gobjBanco.ExecutaRollbackTrans
                    Exit Function
                End If
            End If
        Next i
        gobjBanco.ExecutaCommitTrans
        GravaValores = True
        
'        dbcintUtilizacaoDaCaracteristica.BoundText = ""
'        cbointCaracteristica.BoundText = ""
'        Set cbointCaracteristica.RowSource = Nothing
'        LimpaGrid
        
    End If
    
End Function

Private Function blnDadosOk() As Boolean
Dim i, j As Integer
    
    If Not dbcintUtilizacaoDaCaracteristica.MatchedWithList Then
        ExibeMensagem "Selecione alguma utilização válida."
        dbcintUtilizacaoDaCaracteristica.SetFocus
        Exit Function
    End If
    
    If Not dbcintCategoriaConstrucao.MatchedWithList Then
        ExibeMensagem "Selecione alguma categoria válida."
        dbcintCategoriaConstrucao.SetFocus
        Exit Function
    End If
    
    If Not cbointCaracteristica.MatchedWithList Then
        ExibeMensagem "Selecione alguma característica válida."
        cbointCaracteristica.SetFocus
        Exit Function
    End If
    
    If dbcintUtilizacaoDaCaracteristica.BoundText = 2 And Not dbcintReferenciaTributo.MatchedWithList Then
        ExibeMensagem "Para esta utilização deve ser informado o Tipo."
        dbcintReferenciaTributo.SetFocus
        Exit Function
    End If
    
    grd_Detalhes.MoveFirst
    i = 0
    Do While i <= x.Count(1) - 1
      If x(i, 0) = "" And x(i, 1) = "" Then
           x.DeleteRows (i)
      Else
        grd_Detalhes.MoveNext
        If x(i, 0) = "" Then
            grd_Detalhes.Col = 0
            grd_Detalhes.Row = j
            grd_Detalhes.SetFocus
            ExibeMensagem "O código tem que ser digitado."
            Exit Function
        ElseIf x(i, 1) = "" Then
            grd_Detalhes.Col = 1
            grd_Detalhes.Row = j
            grd_Detalhes.SetFocus
            ExibeMensagem "O nome tem que ser digitado."
            Exit Function
        End If
        i = i + 1
      End If
      j = j + 1
      
    Loop
    grd_Detalhes.Rebind
    grd_Detalhes.Refresh
    DoEvents
    
    blnDadosOk = True
End Function

Private Sub LimpaGrid()
    
    Set x = New XArrayDB
    Set y = New XArrayDB

'    Y.ReDim 0, 0, 0, 2

    x.Clear
    y.Clear

    x.ReDim 0, 0, 0, 3

    Set grd_Detalhes.Array = x
    grd_Detalhes.Rebind
    grd_Detalhes.Refresh

    Set tdd_Detalhes.Array = y
    tdd_Detalhes.Rebind
    tdd_Detalhes.Refresh
    
    strExclusao = ""
    
End Sub

Private Function blnValorCadastrado(DBLVALOR As Variant) As Boolean
Dim i As Integer

    For i = 0 To y.Count(1) - 1
        If gvntConvVrDoSql(y(i, 1)) = gvntConvVrDoSql(DBLVALOR) Then
            blnValorCadastrado = True
            Exit Function
        End If
    Next
    blnValorCadastrado = False
End Function

Private Function strQueryCaracteristicas() As String
Dim strSQL As String
    
    LimpaGrid
    
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strNomeDaCaracteristica "
    strSQL = strSQL & " FROM " & gstrCaracteristicaGeral & " "
    strSQL = strSQL & " WHERE bytCaracteristica = 1"
    If dbcintUtilizacaoDaCaracteristica.MatchedWithList Then
        strSQL = strSQL & " AND intUtilizacaoDaCaracteristica = " & gstrItemData(dbcintUtilizacaoDaCaracteristica) & " "
    End If
    If dbcintCategoriaConstrucao.MatchedWithList Then
        strSQL = strSQL & " AND intCategoriaConstrucao = " & gstrItemData(dbcintCategoriaConstrucao) & " "
    End If
    
    strSQL = strSQL & "ORDER BY strNomeDaCaracteristica "
    
    strQueryCaracteristicas = strSQL

End Function

Function strQuerryRelatorio() As String
Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT UT.strNomeDaUtilizacao Utilizacao, CG.strNomeDaCaracteristica Caracteristica, DC.strNomeDoDetalhe Detalhe, DC.* "
    strSQL = strSQL & "FROM " & gstrUtilizacaoDaTabelaDeValor & " UT,"
    strSQL = strSQL & gstrCaracteristicaGeral & " CG,"
    strSQL = strSQL & gstrDetalheDaCaracteristica & " DC "
    If mblnAlterando = True Then
        strSQL = strSQL & "WHERE CG.intUtilizacaoDaCaracteristica = " & gstrItemData(dbcintUtilizacaoDaCaracteristica) & " and CG.intUtilizacaoDaCaracteristica = UT.PKId and DC.intCaracteristica = CG.PKId "
        Else
            strSQL = strSQL & "WHERE CG.intUtilizacaoDaCaracteristica = UT.PKId and DC.intCaracteristica = CG.PKId "
    End If
    strSQL = strSQL & " ORDER BY Utilizacao"
    
    strQuerryRelatorio = strSQL
    
End Function

Private Function blnTipoJaCadastrado(lngTipo As Long) As Boolean
Dim adoConsulta  As New ADODB.Recordset
Dim strSQL       As String

    strSQL = "SELECT Pkid FROM " & gstrDetalheDaCaracteristica & " WHERE intReferenciaTributo = " & lngTipo & " AND intCaracteristica <> " & cbointCaracteristica.BoundText
    
    If gobjBanco.CriaADO(strSQL, 5, adoConsulta) Then
        blnTipoJaCadastrado = Not adoConsulta.EOF
    End If
    
End Function

Private Function strQueryCategoriaConstrucao() As String
Dim strSQL As String

    strSQL = "SELECT Pkid,"
    strSQL = strSQL & " strDescricao"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrCategoriaConstrucao
    strSQL = strSQL & " WHERE intUtilizacaoTabelaValor = '" & dbcintUtilizacaoDaCaracteristica.BoundText & "'"

    strQueryCategoriaConstrucao = strSQL

End Function

Private Function strQueryCondicao() As String
'Função usada para pegar a condição
End Function
