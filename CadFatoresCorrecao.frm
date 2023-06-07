VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadFatoresCorrecao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fatores de Correção"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   HelpContextID   =   35
   Icon            =   "CadFatoresCorrecao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6750
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   11906
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Fatores de Correção"
      TabPicture(0)   =   "CadFatoresCorrecao.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintUtilizacao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dbcintUtilizacao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_CaracteristicaVertical"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_DetalheCaracteristicaVertical"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_CaracteristicaHorizontal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fra_CaracteristicaHorizontal 
         Caption         =   " Caracteristica Horizontal "
         Height          =   1095
         Left            =   150
         TabIndex        =   7
         Top             =   900
         Width           =   6375
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   5850
            Picture         =   "CadFatoresCorrecao.frx":105E
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Ativa cadastro de Características Gerais"
            Top             =   660
            Width           =   360
         End
         Begin VB.CommandButton cmd_Caracteristica 
            Height          =   315
            Left            =   5850
            Picture         =   "CadFatoresCorrecao.frx":117C
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Características Gerais"
            Top             =   270
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintCaracteristicaHorizontal 
            Height          =   315
            Left            =   1020
            TabIndex        =   8
            Top             =   270
            Width           =   4830
            _ExtentX        =   8520
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintDetalheCaracteristicaHorizontal 
            Height          =   315
            Left            =   1020
            TabIndex        =   9
            Top             =   660
            Width           =   4830
            _ExtentX        =   8520
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblintCaracteristica 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   210
            TabIndex        =   11
            Top             =   390
            Width           =   720
         End
         Begin VB.Label lblintDetalheCaracteristicaHorizontal 
            AutoSize        =   -1  'True
            Caption         =   "Detalhe"
            Height          =   195
            Left            =   375
            TabIndex        =   10
            Top             =   750
            Width           =   555
         End
      End
      Begin VB.Frame fra_DetalheCaracteristicaVertical 
         Caption         =   " Detalhes "
         Height          =   4485
         Left            =   4230
         TabIndex        =   5
         Top             =   2100
         Width           =   4905
         Begin TrueOleDBGrid70.TDBGrid tdb_DetalheVertical 
            Height          =   4035
            Left            =   180
            TabIndex        =   6
            Top             =   300
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   7117
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKID"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descrição"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Valor"
            Columns(3).DataField=   ""
            Columns(3).DataWidth=   7
            Columns(3).NumberFormat=   "Standard"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1164"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1085"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=4551"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=4471"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8196"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1799"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1720"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowDelete     =   -1  'True
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
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
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
      Begin VB.Frame fra_CaracteristicaVertical 
         Caption         =   " Característica Vertical "
         Height          =   4485
         Left            =   120
         TabIndex        =   3
         Top             =   2100
         Width           =   4035
         Begin TrueOleDBGrid70.TDBGrid tdb_CaracteristicaVertical 
            Height          =   4035
            Left            =   180
            TabIndex        =   4
            Top             =   300
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   7117
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKID"
            Columns(0).DataField=   "PKID"
            Columns(0).DropDown=   "tdd_Valores"
            Columns(0).DropDown.vt=   8
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "intCodigoDaCaracteristica"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descrição"
            Columns(2).DataField=   "strNomeDaCaracteristica"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=4022"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3942"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1191"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1111"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=4789"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4710"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowDelete     =   -1  'True
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
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
            _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=39:EvenRow"
            _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(57)  =   "Named:id=40:OddRow"
            _StyleDefs(58)  =   ":id=40,.parent=33"
            _StyleDefs(59)  =   "Named:id=41:RecordSelector"
            _StyleDefs(60)  =   ":id=41,.parent=34"
            _StyleDefs(61)  =   "Named:id=42:FilterBar"
            _StyleDefs(62)  =   ":id=42,.parent=33"
         End
      End
      Begin MSDataListLib.DataCombo dbcintUtilizacao 
         Height          =   315
         Left            =   1170
         TabIndex        =   1
         Top             =   450
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblintUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   570
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmCadFatoresCorrecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando             As Boolean
Dim mobjAux                   As Object
    
Dim X                         As New XArrayDB 'Grid Detalhes
    
Dim mblnSelecionou            As Boolean
Dim mblnPrimeiraVez           As Boolean

Private Function strQueryCaracteristicaVertical() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT PKId, intCodigoDaCaracteristica, strNomeDaCaracteristica FROM "
    strSql = strSql & gstrCaracteristicaGeral & " "
    strSql = strSql & "WHERE intUtilizacaoDaCaracteristica = 3 "
    strSql = strSql & "AND PKId <> " & dbcintCaracteristicaHorizontal.BoundText & " "
    strSql = strSql & " AND bytFator = 1"
    strSql = strSql & " ORDER BY strNomeDaCaracteristica"
    strQueryCaracteristicaVertical = strSql
End Function

Private Function strQueryDetalheCaracteristicaHorizontal() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql  As String
    strSql = ""
'    strSql = strSql & "SELECT PKId, RTRIM(Convert(char(10),intCodigoDoDetalhe)) + ' - ' + strNomeDoDetalhe AS Detalhe "
    strSql = strSql & "SELECT PKId, RTRIM(" & gstrCONVERT(CDT_VARCHAR, "intCodigoDoDetalhe") & ") " & strCONCAT & " ' - ' " & strCONCAT & " strNomeDoDetalhe AS Detalhe "
    strSql = strSql & "FROM " & gstrDetalheDaCaracteristica & " "
    strSql = strSql & "WHERE intCaracteristica = " & dbcintCaracteristicaHorizontal.BoundText & " "
    strSql = strSql & "ORDER BY intCodigoDoDetalhe"
    strQueryDetalheCaracteristicaHorizontal = strSql
End Function

Private Sub cmd_Caracteristica_Click()
    If dbcintUtilizacao.BoundText = "" Then
        ExibeMensagem "Selecione a utilização."
        Exit Sub
    End If
    ChamaFormCadastro frmCadCaracteristicasGerais, dbcintCaracteristicaHorizontal
End Sub

Private Sub dbcintCaracteristicaHorizontal_Click(Area As Integer)
   DropDownDataCombo dbcintCaracteristicaHorizontal, Me, Area
   If Area = 2 And dbcintCaracteristicaHorizontal.MatchedWithList Then
       LeDaTabelaParaObj gstrCaracteristicaGeral, tdb_CaracteristicaVertical, strQueryCaracteristicaVertical
       LeDaTabelaParaObj gstrDetalheDaCaracteristica, dbcintDetalheCaracteristicaHorizontal, strQueryDetalheCaracteristicaHorizontal
       LimpaGrid
       tdb_CaracteristicaVertical.HighlightRowStyle.BackColor = vbWhite
       tdb_CaracteristicaVertical.HighlightRowStyle.ForeColor = vbBlack
   End If
End Sub

Private Sub dbcintCaracteristicaHorizontal_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintCaracteristicaHorizontal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCaracteristicaHorizontal_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintCaracteristicaHorizontal
End Sub

Private Sub dbcintDetalheCaracteristicaHorizontal_Click(Area As Integer)
   DropDownDataCombo dbcintDetalheCaracteristicaHorizontal, Me, Area
End Sub

Private Sub dbcintDetalheCaracteristicaHorizontal_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintDetalheCaracteristicaHorizontal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintDetalheCaracteristicaHorizontal_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintDetalheCaracteristicaHorizontal
End Sub

Private Sub dbcintUtilizacao_Click(Area As Integer)
   DropDownDataCombo dbcintUtilizacao, Me, Area
End Sub

Private Sub dbcintUtilizacao_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUtilizacao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUtilizacao_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", dbcintUtilizacao
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    gintCodSeguranca = 613
    VirificaGradeListView Me
    dbcintCaracteristicaHorizontal.SetFocus
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    mblnPrimeiraVez = False
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrAplicar
    
    LeDaTabelaParaObj gstrUtilizacaoDaTabelaDeValor, dbcintUtilizacao, "PKId, strNomeDaUtilizacao"
    dbcintUtilizacao.BoundText = 3
    TrocaCorObjeto dbcintUtilizacao, True
    
    LeDaTabelaParaObj gstrCaracteristicaGeral, dbcintCaracteristicaHorizontal, strQueryCaracteristica
    
    Set tdb_CaracteristicaVertical.DataSource = Nothing
    VerificaObjParaAplicar mobjAux
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Function strQuery() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao, strAbreviatura FROM "
    strSql = strSql & gstrUnidadeMedida & " ORDER BY strDescricao"
    strQuery = strSql
End Function

Private Function strQueryCaracteristica() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNomeDaCaracteristica FROM "
    strSql = strSql & gstrCaracteristicaGeral & " "
    strSql = strSql & "WHERE intUtilizacaoDaCaracteristica = 3 "
    strSql = strSql & "ORDER BY strNomeDaCaracteristica"
    strQueryCaracteristica = strSql
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
        Case gstrSalvar
            If blnDadosOk Then
                GravaValores
            End If
        Case gstrDeletar
        Case gstrNovo
            Novo
    End Select
End Sub

Private Function blnDadosOk() As Boolean
    If dbcintCaracteristicaHorizontal.BoundText = "" Then
        ExibeMensagem "A característica horizontal tem que ser selecionada."
        dbcintCaracteristicaHorizontal.SetFocus
        Exit Function
    ElseIf dbcintDetalheCaracteristicaHorizontal.BoundText = "" Then
        ExibeMensagem "O detalhe da característica horizontal tem que ser selecionado."
        dbcintDetalheCaracteristicaHorizontal.SetFocus
        Exit Function
    End If
    
    With tdb_CaracteristicaVertical
        If .BOF Or .EOF Then
            Exit Function
        ElseIf .Columns("PKID").Value = "" Then
            Exit Function
        End If
    End With
    
    With tdb_DetalheVertical
        If .BOF Or .EOF Then
            Exit Function
        End If
    End With
    
    blnDadosOk = True
End Function

Private Sub GravaValores()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 06/05/2003
' Alteração: - Alteração do nome do atributo intDetalheCaracteristicaHorizontal e
'            intDetalheCaracteristicaVertical da tabela tblFatorDeCorrecao para
'            intDetalheCaracteristicaHorizo e intDetalheCaracteristicaVertic
'            respectivamente.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim i      As Integer
    
    On Error GoTo Err_Handle
    
    If MsgBox("Confirma gravação?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    strSql = ""
    strSql = strSql & "DELETE FROM " & gstrFatorDeCorrecao & " "
    strSql = strSql & "WHERE intCaracteristicaHorizontal = " & dbcintCaracteristicaHorizontal.BoundText & " "
'    strSql = strSql & "AND intDetalheCaracteristicaHorizontal = " & dbcintDetalheCaracteristicaHorizontal.BoundText & " "
    strSql = strSql & "AND intDetalheCaracteristicaHorizo = " & dbcintDetalheCaracteristicaHorizontal.BoundText & " "
    strSql = strSql & "AND intCaracteristicaVertical = " & tdb_CaracteristicaVertical.Columns("PKID").Value
    
    If Not gobjBanco.Execute(strSql) Then
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If
    
    tdb_DetalheVertical.MoveFirst
    tdb_DetalheVertical.Update
    
    For i = 0 To X.Count(1) - 1
        If X(i, 3) = "" Or IsNull(X(i, 3)) Or X(i, 3) = Empty Or X(i, 3) = "0" Then
            GoTo ProximaLinha
        End If
    
        strSql = ""
        strSql = strSql & "INSERT INTO " & gstrFatorDeCorrecao & " ("
'        strSql = strSql & "intCaracteristicaHorizontal, intDetalheCaracteristicaHorizontal, "
        strSql = strSql & "intCaracteristicaHorizontal, intDetalheCaracteristicaHorizo, "
'        strSql = strSql & "intCaracteristicaVertical, intDetalheCaracteristicaVertical, "
        strSql = strSql & "intCaracteristicaVertical, intDetalheCaracteristicaVertic, "
        strSql = strSql & "dblValor, "
        strSql = strSql & "dtmDtAtualizacao, lngCodUsr"
        strSql = strSql & ") Values ("
        strSql = strSql & dbcintCaracteristicaHorizontal.BoundText & ", "
        strSql = strSql & dbcintDetalheCaracteristicaHorizontal.BoundText & ", "
        strSql = strSql & tdb_CaracteristicaVertical.Columns("PKID").Value & ", "
        
        strSql = strSql & X(i, 0) & ", " 'PKID do Detalhe Vertical

        strSql = strSql & gstrConvVrParaSql(X(i, 3)) & ", " 'Valor

'        strSql = strSql & "getdate()" & ", "
        strSql = strSql & strGETDATE & ", "
        strSql = strSql & glngCodUsr
        strSql = strSql & ")"

        If Not gobjBanco.Execute(strSql, False) Then
            gobjBanco.ExecutaRollbackTrans
            Exit Sub
        End If
ProximaLinha:
    Next i
    
    gobjBanco.ExecutaCommitTrans
    
Exit Sub
Err_Handle:
    gobjBanco.ExecutaRollbackTrans
End Sub

Private Sub tdb_CaracteristicaVertical_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_CaracteristicaVertical_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", tdb_CaracteristicaVertical
End Sub

Private Sub tdb_CaracteristicaVertical_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_CaracteristicaVertical
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnPrimeiraVez = False
                If dbcintDetalheCaracteristicaHorizontal.BoundText = "" Then
                    ExibeMensagem "O detalhe da característica horizontal tem que ser selecionado."
                    Exit Sub
                End If
                MontaArray 0
                gCorLinhaSelecionada tdb_CaracteristicaVertical
            End If
        End If
    End With
End Sub

Private Sub MontaArray(intFlag As Integer)

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 06/05/2003
' Alteração: - Alteração do nome do atributo intDetalheCaracteristicaHorizontal e
'            intDetalheCaracteristicaVertical da tabela tblFatorDeCorrecao para
'            intDetalheCaracteristicaHorizo e intDetalheCaracteristicaVertic
'            respectivamente.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 06/05/2003
' Alteração: - Inclusão de comandos de outer join na instrução SELECT para que esta
'            funcionasse igual ao SQL Server.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim varAux       As Variant
    Dim strSql       As String
    Dim AdoResultado As ADODB.Recordset
    
    On Error GoTo Err_Handle
    
    Select Case intFlag
        Case 0  'Grid Detalhes
            Set X = New XArrayDB
            X.Clear
            
            strSql = ""
            strSql = strSql & "SELECT D.PKId, D.intCodigoDoDetalhe, D.strNomeDoDetalhe, F.dblValor "
            strSql = strSql & "FROM " & gstrDetalheDaCaracteristica & " D, "
            strSql = strSql & gstrFatorDeCorrecao & " F "
'            strSql = strSql & "WHERE D.PKID *= F.intDetalheCaracteristicaVertical "
            strSql = strSql & "WHERE D.PKID " & strOUTJSQLServer & "= F.intDetalheCaracteristicaVertic " & strOUTJOracle
            strSql = strSql & "AND D.intCaracteristica = " & tdb_CaracteristicaVertical.Columns("PKID").Value & " "
'            strSql = strSql & "AND F.intCaracteristicaHorizontal = " & dbcintCaracteristicaHorizontal.BoundText & " "
            strSql = strSql & "AND F.intCaracteristicaHorizontal " & strOUTJOracle & "= " & dbcintCaracteristicaHorizontal.BoundText & " "
'            strSql = strSql & "AND F.intDetalheCaracteristicaHorizontal = " & dbcintDetalheCaracteristicaHorizontal.BoundText & " "
            strSql = strSql & "AND F.intDetalheCaracteristicaHorizo " & strOUTJOracle & "= " & dbcintDetalheCaracteristicaHorizontal.BoundText & " "
            strSql = strSql & "ORDER BY D.intCodigoDoDetalhe"
            
            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strSql, 5, AdoResultado
            With AdoResultado
                If Not .EOF Then
                    X.ReDim 0, .RecordCount - 1, 0, 3
                    Do While Not .EOF
                        varAux = !Pkid
                        X(.AbsolutePosition - 1, 0) = varAux

                        varAux = !intCodigoDoDetalhe
                        X(.AbsolutePosition - 1, 1) = varAux

                        varAux = !strNomeDoDetalhe
                        X(.AbsolutePosition - 1, 2) = varAux

                        varAux = gvntConvVrDoSql(!dblValor)
                        X(.AbsolutePosition - 1, 3) = varAux

                        .MoveNext
                    Loop
                Else
                    X.ReDim 0, 0, 0, 3
                    X(0, 0) = ""
                    X(0, 1) = ""
                    X(0, 2) = ""
                    X(0, 3) = ""
                End If
            End With

            Set tdb_DetalheVertical.Array = X
            tdb_DetalheVertical.ReBind
            tdb_DetalheVertical.Refresh
    End Select
    
Exit Sub
Err_Handle:
    ExibeDetalheErro ""
End Sub

Private Sub LimpaGrid()
    Set X = New XArrayDB 'Grid Detalhe
    
    X.Clear
    
'    X.ReDim 0, 0, 0, 2
    
    Set tdb_DetalheVertical.Array = X
    tdb_DetalheVertical.ReBind
    tdb_DetalheVertical.Refresh
End Sub

Private Sub tdb_DetalheVertical_KeyPress(KeyAscii As Integer)
    Select Case tdb_DetalheVertical.Col
        Case 3
            CaracterValido KeyAscii, "V", tdb_DetalheVertical
    End Select
End Sub

Sub Novo()
    dbcintCaracteristicaHorizontal.BoundText = ""
    dbcintDetalheCaracteristicaHorizontal.BoundText = ""
    Set dbcintDetalheCaracteristicaHorizontal.DataSource = Nothing
    LimpaGrid
    Set tdb_CaracteristicaVertical.DataSource = Nothing
    mblnPrimeiraVez = False
    tdb_CaracteristicaVertical.HighlightRowStyle.BackColor = vbWhite
    tdb_CaracteristicaVertical.HighlightRowStyle.ForeColor = vbBlack
    dbcintCaracteristicaHorizontal.SetFocus
End Sub
