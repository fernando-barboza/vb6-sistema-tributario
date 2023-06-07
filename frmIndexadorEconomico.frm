VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmIndexadorEconomico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indexador Econômico"
   ClientHeight    =   5535
   ClientLeft      =   3660
   ClientTop       =   3555
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3450
      Left            =   45
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   15
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   6085
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Indexador Econômico"
      TabPicture(0)   =   "frmIndexadorEconomico.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAbreviatura"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNome"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdb_FormasAtualizacaoValor"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrAbreviatura"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtstrNome"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPKId"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_Tipo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame fra_Tipo 
         Caption         =   "Tipos"
         Height          =   645
         Left            =   165
         TabIndex        =   12
         Top             =   1215
         Width           =   5805
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Correção"
            Height          =   225
            Index           =   3
            Left            =   4395
            TabIndex        =   5
            Top             =   270
            Width           =   1050
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Índice Corrigido"
            Height          =   225
            Index           =   2
            Left            =   2655
            TabIndex        =   4
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Percentual"
            Height          =   225
            Index           =   1
            Left            =   1230
            TabIndex        =   3
            Top             =   270
            Width           =   1185
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Valor"
            Height          =   225
            Index           =   0
            Left            =   345
            TabIndex        =   2
            Top             =   270
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   15
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtstrNome 
         Height          =   315
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   0
         Top             =   450
         Width           =   4950
      End
      Begin VB.TextBox txtstrAbreviatura 
         Height          =   315
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   1
         Top             =   825
         Width           =   1380
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_FormasAtualizacaoValor 
         Height          =   1365
         Left            =   165
         Negotiate       =   -1  'True
         TabIndex        =   6
         Top             =   1950
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   2408
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   ""
         Columns(0).ConvertEmptyCell=   1
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Data"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(8)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(9)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=1799"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=1720"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=3149"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=3069"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
         ExposeCellMode  =   1
         TabAction       =   1
         WrapCellPointer =   -1  'True
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
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=33"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   555
         TabIndex        =   11
         Top             =   555
         Width           =   420
      End
      Begin VB.Label lblAbreviatura 
         AutoSize        =   -1  'True
         Caption         =   "Abreviatura"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   900
         Width           =   810
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_FormasDeAtualizacao 
      Height          =   2055
      Left            =   45
      TabIndex        =   7
      Top             =   3480
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   3625
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
      Columns(1).Caption=   "Nome"
      Columns(1).DataField=   "Nome"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Abreviatura"
      Columns(2).DataField=   "Abreviatura"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tipo"
      Columns(3).DataField=   "Tipo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=7197"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=7117"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1984"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1905"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2619"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2540"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
      TabAction       =   1
      WrapCellPointer =   -1  'True
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
End
Attribute VB_Name = "frmIndexadorEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mobjAux                 As Object
    Dim blnAlterando            As Boolean
    Dim bytOrdenacao            As Byte
    Dim blnOrdenacaoAsc         As Boolean
    Dim blnPrimeiraVez          As Boolean
    Dim strNomeAtual            As String
    Dim strAbreviaturaAtual     As String
    Dim blnOrdenacaoAscValor    As Boolean
    Dim bytOrdenacaoValor       As Byte

Private Sub Form_Activate()
    gintCodSeguranca = 1117
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    LimpaValores
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub tdb_FormasAtualizacaoValor_AfterColEdit(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case Is = 1
            If tdb_FormasAtualizacaoValor.Columns(ColIndex).Value <> "" Then
                tdb_FormasAtualizacaoValor.Columns(ColIndex).Value = gstrDataFormatada(tdb_FormasAtualizacaoValor.Columns(ColIndex).Value)
            End If
        Case Is = 2
            If tdb_FormasAtualizacaoValor.Columns(ColIndex).Value <> "" Then
                tdb_FormasAtualizacaoValor.Columns(ColIndex).Value = gstrConvVrDoSql(tdb_FormasAtualizacaoValor.Columns(ColIndex).Value, 6)
            End If
    End Select
End Sub

Private Sub tdb_FormasAtualizacaoValor_AfterColUpdate(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case Is = 1
            If tdb_FormasAtualizacaoValor.Columns(ColIndex).Value <> "" Then
                tdb_FormasAtualizacaoValor.Columns(ColIndex).Value = gstrDataFormatada(tdb_FormasAtualizacaoValor.Columns(ColIndex).Value)
            End If
        Case Is = 2
            If tdb_FormasAtualizacaoValor.Columns(ColIndex).Value <> "" Then
                tdb_FormasAtualizacaoValor.Columns(ColIndex).Value = gstrConvVrDoSql(tdb_FormasAtualizacaoValor.Columns(ColIndex).Value, 6)
            End If
    End Select
End Sub


Private Sub tdb_FormasAtualizacaoValor_GotFocus()
    tdb_FormasAtualizacaoValor.Col = 1
    tdb_FormasAtualizacaoValor.SetFocus
    tdb_FormasAtualizacaoValor.EditActive = True
End Sub

Private Sub tdb_FormasAtualizacaoValor_KeyPress(KeyAscii As Integer)
    Select Case tdb_FormasAtualizacaoValor.Col
        Case 1
            CaracterValido KeyAscii, "D", tdb_FormasAtualizacaoValor
        Case 2
            If KeyAscii <> 8 Then
                If Len(Mid(tdb_FormasAtualizacaoValor.Columns(2).Value, InStr(1, tdb_FormasAtualizacaoValor.Columns(2).Value, ",") + 1, Len(tdb_FormasAtualizacaoValor.Columns(2).Value))) = 6 Then
                    KeyAscii = 0
                End If
            End If
            CaracterValido KeyAscii, "V", tdb_FormasAtualizacaoValor
    End Select

End Sub


Private Sub tdb_FormasDeAtualizacao_Click()
    blnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_FormasDeAtualizacao) = 1 Then
        tdb_FormasDeAtualizacao_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_FormasDeAtualizacao_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_FormasDeAtualizacao_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_FormasDeAtualizacao
End Sub

Private Sub tdb_FormasDeAtualizacao_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_FormasDeAtualizacao, ColIndex
End Sub

Private Sub tdb_FormasDeAtualizacao_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
            blnPrimeiraVez = True
    End Select
    
End Sub

Private Sub tdb_FormasDeAtualizacao_KeyPress(KeyAscii As Integer)
    Select Case tdb_FormasDeAtualizacao.Col
        Case 1
            CaracterValido KeyAscii, "A", tdb_FormasDeAtualizacao
        Case 2
            CaracterValido KeyAscii, "A", tdb_FormasDeAtualizacao
        Case 3
            CaracterValido KeyAscii, "A", tdb_FormasDeAtualizacao
    End Select
End Sub

Private Sub tdb_FormasDeAtualizacao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_FormasDeAtualizacao
        If Not .EOF And blnPrimeiraVez Then
            txtPKId.Text = .Columns("PKID").Value
            blnAlterando = True
            LeDaTabelaParaObj gstrIndexadorEconomico, Me
            MontaArrayValores Val(txtPKId.Text)
            strNomeAtual = tdb_FormasDeAtualizacao.Columns("Nome").Value
            strAbreviaturaAtual = tdb_FormasDeAtualizacao.Columns("Abreviatura").Value
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSql As String
    Dim blnAlterandoAux As Boolean
    strSql = strQueryRelatorio
    
    If strModoOperacao = UCase("IMPRIMIR") Then
        ToolBarGeral strModoOperacao, gstrIndexadorEconomico, blnAlterando, tdb_FormasDeAtualizacao, Me, mobjAux, strSql, , rptIndexadorEconomico, strQueryRelatorio
        Exit Sub
    End If
    
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrSalvar)
            If Not blnDadosOk Then Exit Sub
            blnAlterandoAux = blnAlterando
            If ToolBarGeral(strModoOperacao, gstrIndexadorEconomico, blnAlterando, tdb_FormasDeAtualizacao, Me, mobjAux, strQuery(gstrSalvar), strQueryAplicar, , , Not blnAlterando) Then
                GravaValores (blnAlterandoAux)
                LimpaValores
            End If
            blnPrimeiraVez = False
            tdb_FormasDeAtualizacao_RowColChange 0, 0
        Case Is = UCase(gstrNovo)
            LimpaObjeto Me
            blnPrimeiraVez = False
            blnAlterando = False
            LimpaValores
            txtstrNome.SetFocus
        Case Is = UCase(gstrDeletar)
            
            Set gobjBanco = New clsBanco
            
            gobjBanco.ExecutaBeginTrans
            
            strSql = "DELETE FROM " & gstrFormaAtualizacaoValor
            strSql = strSql & " WHERE"
            strSql = strSql & " intIndexadorEconomico = " & Val(txtPKId.Text)
            
            gobjBanco.Execute strSql
            
            If ToolBarGeral(strModoOperacao, gstrIndexadorEconomico, blnAlterando, tdb_FormasDeAtualizacao, Me, mobjAux, strQuery, strQueryAplicar) Then
                gobjBanco.ExecutaCommitTrans
                MantemForm gstrNovo
            Else
                gobjBanco.ExecutaRollbackTrans
            End If
            
        Case Else
            ToolBarGeral strModoOperacao, gstrIndexadorEconomico, blnAlterando, tdb_FormasDeAtualizacao, Me, mobjAux, strQuery, strQueryAplicar
    End Select
                 
End Sub

Private Function blnDadosOk()
    blnDadosOk = False
    
    If txtstrNome.Text = "" Then
        ExibeMensagem "É necessário preencher o campo Nome."
        If txtstrNome.Enabled Then txtstrNome.SetFocus
        Exit Function
    End If
    
    If txtstrAbreviatura.Text = "" Then
        ExibeMensagem "É necessário preencher a Abreviatura."
        If txtstrAbreviatura.Enabled Then txtstrAbreviatura.SetFocus
        Exit Function
    End If
    
    If Not blnAlterando Or (blnAlterando And RTrim(LTrim(strNomeAtual)) <> LTrim(RTrim(txtstrNome.Text))) Then
        If gblnExisteCodigo(1, gstrIndexadorEconomico, "strNome", "'" & txtstrNome.Text & "'") Then
            ExibeMensagem "Já existe um registro com o mesmo nome informado."
            If txtstrNome.Enabled Then txtstrNome.SetFocus
            Exit Function
        End If
    End If
    If Not blnAlterando Or (blnAlterando And LTrim(RTrim(strAbreviaturaAtual)) <> LTrim(RTrim(txtstrAbreviatura.Text))) Then
        If gblnExisteCodigo(1, gstrIndexadorEconomico, "strAbreviatura", "'" & txtstrAbreviatura.Text & "'") Then
            ExibeMensagem "Já existe um registro com a mesma abreviatura informado."
            If txtstrAbreviatura.Enabled Then txtstrAbreviatura.SetFocus
            Exit Function
        End If
    End If

    If Not VerificaValores Then Exit Function
        
    blnDadosOk = True
    
End Function

Private Function strQuery(Optional strModoOperacao As String) As String
Dim strSql As String

    strSql = "SELECT FA.Pkid,"
    strSql = strSql & " FA.strNome Nome,"
    strSql = strSql & " FA.strAbreviatura Abreviatura, "
    strSql = strSql & gstrCASEWHEN("BYTTipo", " 0, 'Valor', 1, 'Percentual', 2, 'Índice Corrigido', 3, 'Correção'") & " Tipo"
    strSql = strSql & " FROM "
    strSql = strSql & gstrIndexadorEconomico & " FA"

    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If Not blnAlterando Then
            If glngPegaUltimaChave(gstrIndexadorEconomico, "Pkid") + 1 = 1 Then
                MantemForm gstrLocalizar
            Else
                strSql = strSql & " WHERE FA.Pkid = " & glngPegaUltimaChave(gstrIndexadorEconomico, "Pkid") + 1
            End If
        Else
            strSql = strSql & " WHERE FA.Pkid = " & txtPKId.Text
        End If
    End If
    
    Select Case bytOrdenacao
        Case Is = 1
            strSql = strSql & " ORDER BY FA.strNome " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSql = strSql & " ORDER BY FA.strAbreviatura " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strSql = strSql & " ORDER BY Tipo " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSql

End Function

Private Function strQueryRelatorio()
Dim strSql As String
    strSql = Empty
    strSql = "select strnome,strabreviatura from " & gstrIndexadorEconomico
    strQueryRelatorio = strSql
End Function

Private Sub txtstrAbreviatura_GotFocus()
    MarcaCampo txtstrAbreviatura
End Sub

Private Sub txtstrAbreviatura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrAbreviatura
End Sub

Private Sub txtstrNome_GotFocus()
    MarcaCampo txtstrNome
End Sub

Private Sub txtstrNome_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNome
End Sub

Private Sub MontaArrayValores(lngPkidFormaAtualizacao As Long)

Dim strSql          As String
Dim x               As XArrayDB
Dim adoResultado    As ADODB.Recordset

    Set x = New XArrayDB
    
    strSql = "SELECT FV.Pkid, "
    strSql = strSql & " FV.dtmData, "
    strSql = strSql & " FV.dblValor"
    strSql = strSql & " FROM "
    strSql = strSql & gstrIndexadorEconomico & " FA, "
    strSql = strSql & gstrFormaAtualizacaoValor & " FV"
    strSql = strSql & " WHERE FV.intIndexadorEconomico = FA.Pkid AND"
    strSql = strSql & " FA.Pkid = " & Val(lngPkidFormaAtualizacao)
    strSql = strSql & " ORDER BY FV.dtmDAta DESC"

    Set gobjBanco = New clsBanco
    
    If Not gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        Exit Sub
    End If
            
    If Not adoResultado.EOF Then
        x.ReDim 0, adoResultado.RecordCount - 1, 0, 2
        Dim varAux As Variant
        Do While Not adoResultado.EOF
            
            varAux = adoResultado!Pkid
            x(adoResultado.AbsolutePosition - 1, 0) = varAux
            
            varAux = adoResultado!DTMDATA
            x(adoResultado.AbsolutePosition - 1, 1) = varAux
                    
            varAux = gstrConvVrDoSql(adoResultado!dblValor, 6)
            x(adoResultado.AbsolutePosition - 1, 2) = varAux
            
            adoResultado.MoveNext
        Loop
    
    Else
    
        x.ReDim 0, 0, 0, 2
        x(0, 0) = 0
        x(0, 1) = ""
        x(0, 2) = "0,000000"
        
    End If
    
    
            
    Set tdb_FormasAtualizacaoValor.Array = x
    tdb_FormasAtualizacaoValor.ReBind
    tdb_FormasAtualizacaoValor.Refresh
    
End Sub
Private Function GravaValores(blnAltera As Boolean)
    Dim A       As XArrayDB
    Dim intCont As Integer
    Dim strSql  As String
    Dim strAux  As String

    tdb_FormasAtualizacaoValor.Refresh
    tdb_FormasAtualizacaoValor.Update
    tdb_FormasAtualizacaoValor.MoveFirst
        
    Set A = New XArrayDB
    Set A = tdb_FormasAtualizacaoValor.Array
    
    strSql = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", " ")
    
    If blnAltera Then
        For intCont = 0 To A.Count(1) - 1
            If Val(A.Value(intCont, 0)) <> 0 Then
                strAux = strAux & Val(A.Value(intCont, 0)) & ", "
            End If
        Next
         
        strSql = strSql & "Delete from " & gstrFormaAtualizacaoValor
        strSql = strSql & " Where intIndexadorEconomico = " & Val(txtPKId.Text)
        If Trim(strAux) <> "" Then
            strAux = Mid(strAux, 1, Len(strAux) - 2)
            strSql = strSql & " AND not Pkid in(" & strAux & ")"
        End If
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
    End If
    
    For intCont = 0 To A.Count(1) - 1
        
        If Val(A.Value(intCont, 0)) = 0 Then
            strSql = strSql & "INSERT INTO " & gstrFormaAtualizacaoValor
            strSql = strSql & " (intIndexadorEconomico, dtmData, dblValor, dtmDtAtualizacao, lngCodUsr)"
            strSql = strSql & " VALUES("
            If Val(txtPKId.Text) = 0 Then
                strSql = strSql & glngPegaUltimaChave(gstrIndexadorEconomico, "Pkid") & ", "
            Else
                strSql = strSql & Val(txtPKId.Text) & ", "
            End If
            strSql = strSql & gstrConvDtParaSql(A.Value(intCont, 1)) & ", "
            strSql = strSql & gstrConvVrParaSql(Val(gstrConvVrParaSql(A.Value(intCont, 2)))) & ", "
            strSql = strSql & strGETDATE & ", "
            strSql = strSql & glngCodUsr & ")"
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), ";", " ")
        Else
            strSql = strSql & "UPDATE " & gstrFormaAtualizacaoValor
            strSql = strSql & " SET intIndexadorEconomico = " & Val(txtPKId) & ", "
            strSql = strSql & " dtmData = " & gstrConvDtParaSql(A.Value(intCont, 1)) & ", "
            strSql = strSql & " dblValor = " & gstrConvVrParaSql(Val(gstrConvVrParaSql(A.Value(intCont, 2)))) & ", "
            strSql = strSql & " dtmDtAtualizacao = " & strGETDATE & ", "
            strSql = strSql & " lngCodUsr = " & glngCodUsr
            strSql = strSql & " WHERE Pkid = " & A.Value(intCont, 0)
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), ";", " ")
        End If
    Next
    
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), " END; ", " ")
    
    Set gobjBanco = New clsBanco
    
    If Not gobjBanco.Execute(strSql) Then
        ExibeMensagem "Não foi possivel gravar valores."
    End If
    

End Function

Private Sub LimpaValores()
Dim x As New XArrayDB

    Set x = New XArrayDB

    x.ReDim 0, 0, 0, 2
    x(0, 0) = 0
    x(0, 1) = ""
    x(0, 2) = "0,00"
            
    Set tdb_FormasAtualizacaoValor.Array = x
    tdb_FormasAtualizacaoValor.ReBind
    tdb_FormasAtualizacaoValor.Refresh

End Sub

Private Function VerificaValores() As Boolean

Dim A       As XArrayDB
Dim intCont As Integer

    VerificaValores = False
    
    tdb_FormasAtualizacaoValor.Update
    tdb_FormasAtualizacaoValor.MoveFirst
        
    Set A = tdb_FormasAtualizacaoValor.Array
    If (A.Count(1)) = 0 Then
        ExibeMensagem "É obrigatório preencher grid de valores."
        tdb_FormasAtualizacaoValor.Refresh
        Exit Function
    End If
    For intCont = 0 To A.Count(1) - 1
        If Not gblnDataValida(A.Value(intCont, 1)) Then
            ExibeMensagem "Dados inválidos para valores."
            tdb_FormasAtualizacaoValor.Refresh
            Exit Function
        End If
    Next

    VerificaValores = True

End Function
Private Sub tdb_FormasAtualizacaoValor_HeadClick(ByVal ColIndex As Integer)

Dim x As New XArrayDB

   
   blnOrdenacaoAscValor = IIf(bytOrdenacaoValor = ColIndex, Not blnOrdenacaoAscValor, True)
   bytOrdenacaoValor = ColIndex
    
    Set x = tdb_FormasAtualizacaoValor.Array
 
    Select Case ColIndex
        Case Is = 1
            x.QuickSort x.LowerBound(1), x.UpperBound(1), ColIndex, Abs(blnOrdenacaoAscValor), XTYPE_DATE

        Case Is = 2
            x.QuickSort x.LowerBound(1), x.UpperBound(1), ColIndex, Abs(blnOrdenacaoAscValor), XTYPE_CURRENCY
    End Select
    
    tdb_FormasAtualizacaoValor.Refresh

End Sub

Private Function strQueryAplicar() As String
Dim strSql  As String
    
    strSql = "SELECT PKId, strAbreviatura FROM "
    strSql = strSql & gstrIndexadorEconomico & " ORDER BY strAbreviatura"
    
    strQueryAplicar = strSql
End Function

