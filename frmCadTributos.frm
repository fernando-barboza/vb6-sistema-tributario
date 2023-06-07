VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadTributos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tributos"
   ClientHeight    =   4500
   ClientLeft      =   3180
   ClientTop       =   3255
   ClientWidth     =   6420
   Icon            =   "frmCadTributos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4335
      Left            =   90
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tributos"
      TabPicture(0)   =   "frmCadTributos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTipo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcinttributotipo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_Lista"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPKId"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrDescricao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd_Tipo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtintCodigoTributo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Valores"
      TabPicture(1)   =   "frmCadTributos.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_pkidTributoExercicio"
      Tab(1).Control(1)=   "cmd_indexador"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt_dblvalorexcedente"
      Tab(1).Control(3)=   "txt_intExercicio"
      Tab(1).Control(4)=   "txt_dblvalor"
      Tab(1).Control(5)=   "lvw_Itens"
      Tab(1).Control(6)=   "dbc_intindexadoreconomico"
      Tab(1).Control(7)=   "Label2"
      Tab(1).Control(8)=   "lbldblvalorexcedente"
      Tab(1).Control(9)=   "lbldblValor"
      Tab(1).Control(10)=   "Label1"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Faixa de Valores"
      TabPicture(2)   =   "frmCadTributos.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt_dblValorExcedenteF"
      Tab(2).Control(1)=   "txt_dblValorF"
      Tab(2).Control(2)=   "txt_dblQuantidadeFinalF"
      Tab(2).Control(3)=   "txt_dblQuantidadeInicialF"
      Tab(2).Control(4)=   "txt_DescricaoF"
      Tab(2).Control(5)=   "txt_intTributoExercicioF"
      Tab(2).Control(6)=   "cmd_IndexadorF"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "dbc_intIndexadorF"
      Tab(2).Control(8)=   "tdb_FaixaValores"
      Tab(2).Control(9)=   "Label10"
      Tab(2).Control(10)=   "Label9"
      Tab(2).Control(11)=   "lblQtdeFinal"
      Tab(2).Control(12)=   "lblQtdeInicial"
      Tab(2).Control(13)=   "Label6"
      Tab(2).Control(14)=   "lblExercicioF"
      Tab(2).Control(15)=   "lblDescricaoF"
      Tab(2).ControlCount=   16
      Begin VB.TextBox txt_pkidTributoExercicio 
         Height          =   285
         Left            =   -69960
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_dblValorExcedenteF 
         Height          =   285
         Left            =   -70200
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt_dblValorF 
         Height          =   285
         Left            =   -71760
         TabIndex        =   15
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt_dblQuantidadeFinalF 
         Height          =   285
         Left            =   -73320
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt_dblQuantidadeInicialF 
         Height          =   285
         Left            =   -74880
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt_DescricaoF 
         Height          =   285
         Left            =   -74100
         TabIndex        =   11
         Top             =   630
         Width           =   3720
      End
      Begin VB.TextBox txt_intTributoExercicioF 
         Height          =   285
         Left            =   -69480
         TabIndex        =   12
         Top             =   630
         Width           =   615
      End
      Begin VB.CommandButton cmd_IndexadorF 
         Height          =   315
         Left            =   -71445
         Picture         =   "frmCadTributos.frx":1096
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro Único"
         Top             =   1560
         Width           =   360
      End
      Begin VB.TextBox txtintCodigoTributo 
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
         Left            =   900
         MaxLength       =   6
         TabIndex        =   0
         Top             =   600
         Width           =   900
      End
      Begin VB.CommandButton cmd_indexador 
         Height          =   315
         Left            =   -71160
         Picture         =   "frmCadTributos.frx":1420
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro Único"
         Top             =   1500
         Width           =   360
      End
      Begin VB.TextBox txt_dblvalorexcedente 
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
         Left            =   -70590
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1110
         Width           =   1710
      End
      Begin VB.TextBox txt_intExercicio 
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
         Left            =   -73805
         MaxLength       =   4
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txt_dblvalor 
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
         Left            =   -73800
         MaxLength       =   18
         TabIndex        =   6
         Top             =   1110
         Width           =   1660
      End
      Begin VB.CommandButton cmd_Tipo 
         Height          =   315
         Left            =   5595
         Picture         =   "frmCadTributos.frx":17AA
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "585"
         ToolTipText     =   "Ativa Cadastro Tipo de Tributo"
         Top             =   1395
         Width           =   360
      End
      Begin VB.TextBox txtstrDescricao 
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
         Left            =   900
         MaxLength       =   55
         TabIndex        =   1
         Top             =   990
         Width           =   5040
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2175
         Left            =   195
         TabIndex        =   4
         Top             =   1935
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   3836
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "intCodigoTributo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
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
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1693"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1614"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=5530"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5450"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=3942"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=3863"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
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
      Begin MSDataListLib.DataCombo dbcinttributotipo 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Top             =   1380
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSComctlLib.ListView lvw_Itens 
         Height          =   2235
         Left            =   -74910
         TabIndex        =   10
         Top             =   1980
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Exercício"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor Excedente"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Indexador"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "intIndexador"
            Object.Width           =   0
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbc_intindexadoreconomico 
         Height          =   315
         Left            =   -73800
         TabIndex        =   8
         Top             =   1500
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intIndexadorF 
         Height          =   315
         Left            =   -74085
         TabIndex        =   17
         Top             =   1560
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_FaixaValores 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   19
         Top             =   2040
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   3836
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
         Columns(1).Caption=   "Exercicio"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Quantidade Inicial"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Quantidade Final"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Valor Excedente"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Indexador"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "intIndexador"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Posicao"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Alterado"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
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
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=2487"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2408"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=2328"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2249"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2302"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2223"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=2249"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2170"
         Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(35)=   "Column(6).Width=1508"
         Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=1429"
         Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(40)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(43)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(44)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(45)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(47)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(51)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(52)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(54)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(55)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(57)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(58)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(59)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
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
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(25)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(26)  =   ":id=13,.fontname=MS Sans Serif"
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
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(78)  =   "Named:id=33:Normal"
         _StyleDefs(79)  =   ":id=33,.parent=0"
         _StyleDefs(80)  =   "Named:id=34:Heading"
         _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(82)  =   ":id=34,.wraptext=-1"
         _StyleDefs(83)  =   "Named:id=35:Footing"
         _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(85)  =   "Named:id=36:Selected"
         _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(87)  =   "Named:id=37:Caption"
         _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(89)  =   "Named:id=38:HighlightRow"
         _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(91)  =   "Named:id=39:EvenRow"
         _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(93)  =   "Named:id=40:OddRow"
         _StyleDefs(94)  =   ":id=40,.parent=33"
         _StyleDefs(95)  =   "Named:id=41:RecordSelector"
         _StyleDefs(96)  =   ":id=41,.parent=34"
         _StyleDefs(97)  =   "Named:id=42:FilterBar"
         _StyleDefs(98)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valor Excedente"
         Height          =   195
         Left            =   -70200
         TabIndex        =   35
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -71760
         TabIndex        =   34
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblQtdeFinal 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Final"
         Height          =   195
         Left            =   -73320
         TabIndex        =   33
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblQtdeInicial 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Inicial"
         Height          =   195
         Left            =   -74880
         TabIndex        =   32
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   315
         Left            =   -74880
         TabIndex        =   31
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label lblExercicioF 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   -70200
         TabIndex        =   30
         Top             =   630
         Width           =   675
      End
      Begin VB.Label lblDescricaoF 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   -74880
         TabIndex        =   29
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   195
         Left            =   -74595
         TabIndex        =   27
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label lbldblvalorexcedente 
         AutoSize        =   -1  'True
         Caption         =   "Valor Excedente"
         Height          =   195
         Left            =   -71850
         TabIndex        =   26
         Top             =   1170
         Width           =   1170
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -74250
         TabIndex        =   25
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   -74550
         TabIndex        =   24
         Top             =   780
         Width           =   675
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   525
         TabIndex        =   23
         Top             =   1425
         Width           =   315
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1035
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadTributos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando       As Boolean
    Dim mblnAlterandoF      As Boolean
    Dim mobjAux             As Object
    Dim mblnSelecionou      As Boolean
    Dim mblnClickOk         As Boolean
    Dim bytOrdenacao        As Byte
    Dim blnOrdenacaoAsc     As Boolean
    Dim mobjLista           As Object
    Dim mblnAlterandoLista  As Boolean
    Dim intPkid             As Long
    Dim intPKIDF            As Long
    Dim mblnAlterandoAux    As Boolean
    Dim strDescricao        As String
    Dim strCodigo           As String
    Dim xadbFaixaValores    As New XArrayDB
    Dim PKIDNaoDeletar      As String
    Dim intPosicao          As Integer
    Dim blnOrdenacaoAscF    As Boolean
    

Private Function strQuery() As String

Dim strSQL  As String
   
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "TB.pkID, "
    strSQL = strSQL & gstrREPLICATE("TB.intCodigoTributo", "0", 6) & " intCodigoTributo, "
    strSQL = strSQL & "TB.strDescricao, "
    strSQL = strSQL & "TB.intTributoTipo, "
    strSQL = strSQL & "TBT.strDescricao as Tipo "
    strSQL = strSQL & "From " & gstrTributo & " TB, "
    strSQL = strSQL & gstrTributoTipo & " TBT "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "TBT.Pkid = TB.inttributotipo"
   
    If Len(Trim$(txtPKId.Text)) > 0 Then
        strSQL = strSQL & " And TB.PkID = " & txtPKId.Text
    End If
    
    Select Case bytOrdenacao
        Case Is = 1
            strSQL = strSQL & " Order by TB.intCodigoTributo" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSQL = strSQL & " Order by TB.strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strSQL = strSQL & " Order by TBT.strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSQL
    
End Function

Private Function strQueryAplicar() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao FROM "
    strSQL = strSQL & gstrTributo & " ORDER BY strDescricao"
    strQueryAplicar = strSQL
End Function

Private Sub cmd_indexador_Click()
    CarregaForm frmIndexadorEconomico, dbc_intindexadoreconomico
End Sub

Private Sub cmd_IndexadorF_Click()
    CarregaForm frmIndexadorEconomico, dbc_intindexadoreconomico
End Sub

Private Sub cmd_Tipo_Click()
CarregaForm frmCadTipoTributo
End Sub

Private Sub dbc_intindexadoreconomico_GotFocus()
    MarcaCampo dbc_intindexadoreconomico
End Sub

Private Sub dbc_intindexadoreconomico_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intindexadoreconomico, Me, , , Shift
End Sub

Private Sub dbc_intindexadoreconomico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intindexadoreconomico
End Sub

Private Sub dbc_intIndexadorF_GotFocus()
   MarcaCampo dbc_intIndexadorF
End Sub

Private Sub dbc_intIndexadorF_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intIndexadorF, Me, , , Shift
End Sub

Private Sub dbc_intIndexadorF_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", dbc_intIndexadorF
End Sub

Private Sub dbcinttributotipo_GotFocus()
    MarcaCampo dbcinttributotipo
End Sub

Private Sub dbcinttributotipo_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcinttributotipo, Me, , KeyCode, Shift
End Sub

Private Sub dbcinttributotipo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcinttributotipo
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1157
    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    
    If tab_3dPasta.Tab = 1 Or tab_3dPasta.Tab = 2 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub

Private Sub Form_Load()
    xadbFaixaValores.ReDim 0, 0, 0, 9
    bytOrdenacao = 1: blnOrdenacaoAsc = True
    blnOrdenacaoAscF = True
    dbcinttributotipo.Tag = strQueryTipo & ";strdescricao"
    dbc_intindexadoreconomico.Tag = strQueryIndexEconomico & ";strabreviatura"
    dbc_intIndexadorF.Tag = strQueryIndexEconomico & ";PKID"
    VerificaObjParaAplicar mobjAux
    TrocaCorObjeto txt_intTributoExercicioF, True
    TrocaCorObjeto txt_DescricaoF, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub lvw_Itens_Click()
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset

    If lvw_Itens.ListItems.Count > 0 Then
       txt_intExercicio.Text = lvw_Itens.SelectedItem.Text
       txt_dblvalor = lvw_Itens.SelectedItem.SubItems(1)
       txt_dblvalorexcedente = lvw_Itens.SelectedItem.SubItems(2)
       PreencherListaDeOpcoes dbc_intindexadoreconomico, lvw_Itens.SelectedItem.SubItems(4)
       mblnAlterandoLista = True
       If txtPKId.Text <> "" Then
         strSQL = "SELECT PKID FROM " & gstrTributoExercicio & " WHERE intTributo=" & txtPKId.Text
         strSQL = strSQL & " AND intExercicio=" & txt_intExercicio.Text & " AND dblValor=" & gstrConvVrParaSql(txt_dblvalor.Text)
         strSQL = strSQL & " AND dblValorExcedente=" & gstrConvVrParaSql(txt_dblvalorexcedente.Text)
         strSQL = strSQL & " AND intIndexadorEconomico=" & lvw_Itens.SelectedItem.SubItems(4)
         Set gobjBanco = New clsBanco
         If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
            If Not adoResultado.EOF Then
               txt_pkidTributoExercicio = adoResultado!Pkid
            End If
         End If
      End If
    End If
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    mblnAlterandoF = False
    If tab_3dPasta.Tab <> 2 Then
        MantemForm2 (gstrNovo)
    End If
    If tab_3dPasta.Tab = 1 Or tab_3dPasta.Tab = 2 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        If tab_3dPasta.Tab = 2 Then
           txt_DescricaoF = txtstrDescricao
           txt_intTributoExercicioF = txt_intExercicio
           txt_intTributoExercicioF.Tag = txt_pkidTributoExercicio
           If txt_pkidTributoExercicio = "" Then
              MantemForm2 (gstrNovo)
              xadbFaixaValores.Clear
              xadbFaixaValores.ReDim 0, 0, 0, 9
              tdb_FaixaValores.Array = xadbFaixaValores
              tdb_FaixaValores.ReBind
              tdb_FaixaValores.Refresh
           Else
              CarregaRegistrosFaixaValores
           End If
        End If
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
End Sub

Private Sub tdb_FaixaValores_Click()
   mblnAlterandoF = True
End Sub

Private Sub tdb_FaixaValores_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If Not mblnAlterandoF Then Exit Sub
   If tdb_FaixaValores.EOF Or tdb_FaixaValores.BOF Then Exit Sub
   gCorLinhaSelecionada tdb_FaixaValores
   mblnAlterandoF = True
   txt_dblQuantidadeInicialF = tdb_FaixaValores.Columns("Quantidade Inicial")
   txt_dblQuantidadeFinalF = tdb_FaixaValores.Columns("Quantidade Final")
   txt_dblValorF = tdb_FaixaValores.Columns("Valor")
   txt_dblValorExcedenteF = tdb_FaixaValores.Columns("Valor Excedente")
   dbc_intIndexadorF.Text = tdb_FaixaValores.Columns("Indexador")
End Sub

Private Sub tdb_Lista_Click()
    mblnClickOk = True
End Sub

Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrTributo, Me
            gCorLinhaSelecionada tdb_Lista
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            strDescricao = Trim(tdb_Lista.Columns("strDescricao").Value)
            strCodigo = Val(Trim(txtintCodigoTributo.Text))
            txtintCodigoTributo.Text = tdb_Lista.Columns("Código").Value
            LimpaItens
            PreencheListItens
            mblnSelecionou = True
            mblnAlterando = True
            mblnAlterandoLista = False
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSQL As String
    
    If tab_3dPasta.Tab = 2 Then
       MantemForm2 (strModoOperacao)
       Exit Sub
    End If
    
    Select Case UCase(strModoOperacao)
        Case UCase(gstrNovo)
            If tab_3dPasta.Tab = 0 Then
                mblnClickOk = False
                mblnSelecionou = False
                mblnAlterando = False
                mblnAlterandoLista = False
                LimpaObjeto Me
                tab_3dPasta.Tab = 0
                strDescricao = ""
                txt_intExercicio = ""
                txt_dblvalor = ""
                txt_dblvalorexcedente = ""
                dbc_intindexadoreconomico.Text = ""
                Set dbcinttributotipo.RowSource = Nothing
                lvw_Itens.ListItems.Clear
            Else
                LimpaItens
                txt_intExercicio.SetFocus
            End If
        Case UCase(gstrSalvar)
            If Not blnDadosOk Then Exit Sub
            If mblnAlterando Then
                mblnAlterandoAux = mblnAlterando
                intPkid = txtPKId
            Else
                mblnAlterandoAux = False
            End If
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            If ToolBarGeral(strModoOperacao, gstrTributo, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar) Then
'                strSQL = StrSalvaItem
'                If strSQL <> "" Then
'                    If gobjBanco.Execute(strSQL) Then
               If StrSalvaItem Then
                  gobjBanco.ExecutaCommitTrans
                  tab_3dPasta.Tab = 0
                  LeDaTabelaParaObj "", tdb_Lista, strQuery
                  MantemForm gstrNovo
               Else
                  Set gobjBanco = New clsBanco
                  gobjBanco.ExecutaRollbackTrans
                End If
'                Else
'                    gobjBanco.ExecutaCommitTrans
'                    tab_3dPasta.Tab = 0
'                    MantemForm gstrNovo
'                    LeDaTabelaParaObj "", tdb_Lista, strQuery
'                End If
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
            End If
            Set gobjBanco = Nothing
        Case UCase(gstrIncluirItem)
            IncluirItemNoGrid
        Case UCase(gstrExcluirItem)
            ExcluirItemNoGrid
            mblnAlterandoLista = False
        Case UCase(gstrDeletar)
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            strSQL = " Delete from " & gstrTributoExercicio & " Where inttributo = " & txtPKId
            If gobjBanco.Execute(strSQL) Then
                If ToolBarGeral(strModoOperacao, gstrTributo, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar) Then
                    gobjBanco.ExecutaCommitTrans
                    tab_3dPasta.Tab = 0
                    LeDaTabelaParaObj "", tdb_Lista, strQuery
                    MantemForm gstrNovo
                Else
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaRollbackTrans
                End If
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
            End If
        Case Else
            ToolBarGeral strModoOperacao, gstrTributo, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar
    End Select
    strSQL = strQueryRelatorio
    If strModoOperacao = UCase("IMPRIMIR") Then
        ToolBarGeral strModoOperacao, gstrTributo, mblnAlterando, tdb_Lista, Me, mobjAux, strSQL, , rpttributos, strQueryRelatorio
        Exit Sub
    End If
End Sub


Function strQueryRelatorio() As String
    
Dim strSQL As String
   
   strSQL = ""
   strSQL = "select strdescricao,inttributotipo from " & gstrTributo
   
     
   Select Case bytOrdenacao
      
      Case Is = 1

      Case Is = 2

      
      Case Is = 3

         
   End Select
   
   strQueryRelatorio = strSQL
   
End Function

Private Function blnDadosOk() As Boolean
Dim adoResultado As ADODB.Recordset
Dim strSQL As String

    blnDadosOk = False
    
    If Trim(txtintCodigoTributo.Text) = "" Then
       ExibeMensagem "O Código deve ser informado."
       txtintCodigoTributo.SetFocus
       Exit Function
    End If
    
    If Trim(txtstrDescricao.Text) = "" Then
       ExibeMensagem "A Descrição deve ser informada."
       txtstrDescricao.SetFocus
       Exit Function
    End If
    
    If dbcinttributotipo.MatchedWithList = False Then
       ExibeMensagem "O Tipo do Tributo deve ser informado."
       dbcinttributotipo.SetFocus
       Exit Function
    End If
    
    strSQL = ""
    strSQL = strSQL & "SELECT pkID "
    strSQL = strSQL & "FROM " & gstrTributo & " TR, "
    strSQL = strSQL & gstrTributoTipo & " TT "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & ""
    
    
    
    If Not mblnAlterando Or (mblnAlterando And Val(strCodigo) <> Val(Trim(txtintCodigoTributo.Text))) Then
        If gblnExisteCodigo(1, gstrTributo, "intCodigoTributo", Val(Trim(txtintCodigoTributo.Text)), "intTributoTipo", dbcinttributotipo.BoundText) Then
            ExibeMensagem "O Código informado já se encontra cadastrado nesse Tipo de Tributo."
            txtintCodigoTributo.SetFocus
            Exit Function
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strDescricao) <> UCase$(Trim(txtstrDescricao.Text))) Then
        If gblnExisteCodigo(1, gstrTributo, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
    
End Function

Private Function strQueryTipo() As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select Pkid, Strdescricao From " & gstrTributoTipo
    strSQL = strSQL & " Order By strDescricao"
    
    strQueryTipo = strSQL
End Function

Private Sub txt_dblValor_GotFocus()
    MarcaCampo txt_dblvalor
End Sub

Private Sub txt_DBLVALOR_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblvalor
End Sub

Private Sub txt_DBLVALOR_LostFocus()
    txt_dblvalor = gstrConvVrDoSql(txt_dblvalor, 6)
End Sub

Private Sub txt_dblvalorexcedente_GotFocus()
    MarcaCampo txt_dblvalorexcedente
End Sub

Private Sub txt_dblvalorexcedente_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblvalorexcedente
End Sub


Private Sub txt_dblvalorexcedente_LostFocus()
    txt_dblvalorexcedente = gstrConvVrDoSql(txt_dblvalorexcedente, 6)
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub
Private Function strQueryIndexEconomico() As String
    Dim strSQL As String
    
    strSQL = "SELECT Pkid,"
    strSQL = strSQL & " strabreviatura "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrIndexadorEconomico
    strSQL = strSQL & " ORDER BY strAbreviatura"
    
    strQueryIndexEconomico = strSQL
End Function

Private Function IncluirItemNoGrid()
    Dim intInd          As Integer
    If blnDadosItens = False Then Exit Function
    With lvw_Itens
        If mblnAlterandoLista Then
            For intInd = 1 To .ListItems.Count
                If .SelectedItem.Index <> intInd Then
                    If Trim(txt_intExercicio) = .ListItems(intInd).Text Then
                        ExibeMensagem "Não é possível incluir itens com exercícios iguais."
                        Exit Function
                    End If
                End If
            Next
            .SelectedItem.Text = txt_intExercicio
            .SelectedItem.SubItems(1) = gstrConvVrDoSql(txt_dblvalor, 6)
            .SelectedItem.SubItems(2) = gstrConvVrDoSql(txt_dblvalorexcedente, 6)
            .SelectedItem.SubItems(3) = dbc_intindexadoreconomico.Text
            .SelectedItem.SubItems(4) = dbc_intindexadoreconomico.BoundText
            mblnAlterandoLista = False
        Else
            For intInd = 1 To .ListItems.Count
                If Trim(txt_intExercicio) = .ListItems(intInd).Text Then
                    ExibeMensagem "Não é possível incluir itens com exercícios iguais."
                    Exit Function
                End If
            Next

            Set mobjLista = .ListItems.Add(, , txt_intExercicio)
            mobjLista.SubItems(1) = gstrConvVrDoSql(txt_dblvalor, 6)
            mobjLista.SubItems(2) = gstrConvVrDoSql(txt_dblvalorexcedente, 6)
            mobjLista.SubItems(3) = dbc_intindexadoreconomico.Text
            mobjLista.SubItems(4) = dbc_intindexadoreconomico.BoundText
        End If
    End With
    txt_intExercicio.Text = ""
    txt_dblvalor.Text = ""
    txt_dblvalorexcedente.Text = ""
    dbc_intindexadoreconomico.Text = ""
    
End Function

Private Function ExcluirItemNoGrid()
    With lvw_Itens
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
End Function

Private Function blnDadosItens() As Boolean
    blnDadosItens = False
    If Trim(Len(txt_intExercicio.Text)) <> 4 Then
        ExibeMensagem "O exercício deve ser preenchido corretamente."
        txt_intExercicio.SetFocus
        Exit Function
    ElseIf Trim(txt_dblvalor) = "" Then
        ExibeMensagem "O valor deve ser preenchido corretamente."
        txt_dblvalor.SetFocus
        Exit Function
    ElseIf Trim(txt_dblvalorexcedente) = "" Then
        ExibeMensagem "O valor excedente deve ser preenchido corretamente."
        txt_dblvalorexcedente.SetFocus
        Exit Function
    ElseIf dbc_intindexadoreconomico.MatchedWithList = False Then
        ExibeMensagem "O indexador deve ser preenchido corretamente."
        dbc_intindexadoreconomico.SetFocus
        Exit Function
    End If
    blnDadosItens = True
End Function
Private Function StrSalvaItem() As Boolean
    Dim strSQL  As String
    Dim strSql2 As String
    Dim adoResultado As New ADODB.Recordset
    Dim intInd  As Integer
    Dim strPKId As String
    Dim strPKIDAtualizar As String
    Dim vetPKIDDeletar() As String
    
    strSQL = ""
'    If lvw_Itens.ListItems.Count > 0 Then
'        strSQL = IIf(bytDBType = Oracle, "Begin", "")
'    End If
    
'    If mblnAlterandoAux Then
'        strSQL = strSQL & " Delete from " & gstrTributoExercicio & " Where inttributo = " & intPkid
'        If lvw_Itens.ListItems.Count > 0 Then
'           strSQL = strSQL & IIf(bytDBType = Oracle, ";", "")
'        End If
'    End If
    If lvw_Itens.ListItems.Count > 0 Then
        With lvw_Itens
            For intInd = 1 To .ListItems.Count
                
                strSql2 = "SELECT PKID FROM " & gstrTributoExercicio & " WHERE intExercicio=" & .ListItems(intInd).Text
                strSql2 = strSql2 & " AND intTributo = "
                If mblnAlterandoAux Then
                    strSql2 = strSql2 & intPkid
                Else
                    strSql2 = strSql2 & glngPegaUltimaChave(gstrTributo, "pkid")
                End If
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSql2, 10, adoResultado) Then
                    If Not adoResultado.EOF Then
                        If strPKId = "" Then
                            strPKId = adoResultado!Pkid
                        Else
                            strPKId = strPKId & "," & adoResultado!Pkid
                        End If
                        strPKIDAtualizar = adoResultado!Pkid
                    Else
                        strPKIDAtualizar = ""
                    End If
                    adoResultado.Close
                End If
                
                If strPKIDAtualizar <> "" Then
                    strSQL = " UPDATE "
                    strSQL = strSQL & gstrTributoExercicio & " SET "
                    strSQL = strSQL & "inttributo = "
                    If mblnAlterandoAux Then
                        strSQL = strSQL & intPkid & ", "
                    Else
                        strSQL = strSQL & glngPegaUltimaChave(gstrTributo, "pkid") & ", "
                    End If
                    strSQL = strSQL & "intexercicio = " & .ListItems(intInd).Text & ", "
                    strSQL = strSQL & "dblvalor = " & gstrConvVrParaSql(.ListItems(intInd).SubItems(1)) & ", "
                    strSQL = strSQL & "dblvalorexcedente = " & gstrConvVrParaSql(.ListItems(intInd).SubItems(2)) & ", "
                    strSQL = strSQL & "intindexadoreconomico = " & .ListItems(intInd).SubItems(4) & ", "
                    strSQL = strSQL & "dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
                    strSQL = strSQL & "WHERE PKID=" & strPKIDAtualizar
                 Else
                    strSQL = " INSERT INTO "
                    strSQL = strSQL & gstrTributoExercicio & " ("
                    strSQL = strSQL & "inttributo, "
                    strSQL = strSQL & "intexercicio, "
                    strSQL = strSQL & "dblvalor, "
                    strSQL = strSQL & "dblvalorexcedente, "
                    strSQL = strSQL & "intindexadoreconomico, "
                    strSQL = strSQL & "dtmDtAtualizacao, "
                    strSQL = strSQL & "lngCodUsr) "
                    strSQL = strSQL & "Values("
                    If mblnAlterandoAux Then
                        strSQL = strSQL & intPkid & ", "
                    Else
                        strSQL = strSQL & glngPegaUltimaChave(gstrTributo, "pkid") & ", "
                    End If
                    strSQL = strSQL & .ListItems(intInd).Text & ", "
                    strSQL = strSQL & gstrConvVrParaSql(.ListItems(intInd).SubItems(1)) & ", "
                    strSQL = strSQL & gstrConvVrParaSql(.ListItems(intInd).SubItems(2)) & ", "
                    strSQL = strSQL & .ListItems(intInd).SubItems(4) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSQL = strSQL & glngCodUsr & ") "
'                    strSQL = strSQL & ")" & IIf(bytDBType = Oracle, ";", "")
                End If
                
                Set gobjBanco = New clsBanco
                If gobjBanco.Execute(strSQL) Then
                    If strPKIDAtualizar = "" Then
                        strSql2 = "SELECT MAX(PKID) as PKID FROM " & gstrTributoExercicio & " WHERE lngCodUsr = " & glngCodUsr
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strSql2, 10, adoResultado) Then
                            If strPKId = "" Then
                                strPKId = adoResultado!Pkid
                            Else
                                strPKId = strPKId & "," & adoResultado!Pkid
                            End If
                            adoResultado.Close
                        End If
                    End If
                Else
                    StrSalvaItem = False
                    Exit Function
                End If
            Next
        End With
    End If
    
    strSql2 = "SELECT PKID FROM " & gstrTributoExercicio & " WHERE intTributo = "
    If mblnAlterandoAux Then
        strSql2 = strSql2 & intPkid
    Else
        strSql2 = strSql2 & glngPegaUltimaChave(gstrTributo, "pkid")
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql2, 10, adoResultado) Then
        Do While Not adoResultado.EOF
            strSQL = "DELETE FROM " & gstrTributosFaixa & " WHERE intTributoExercicio = " & adoResultado!Pkid
            If strPKId <> "" Then
                If InStr(1, strPKId, ",") = 0 Then
                    If strPKId = adoResultado!Pkid Then
                        strSQL = ""
                    End If
                Else
                    vetPKIDDeletar = Split(strPKId, ",")
                    For intInd = 0 To UBound(vetPKIDDeletar)
                        If adoResultado!Pkid = vetPKIDDeletar(intInd) Then
                            strSQL = ""
                            Exit For
                        End If
                    Next
                End If
            End If
            If strSQL <> "" Then
                Set gobjBanco = New clsBanco
                If Not gobjBanco.Execute(strSQL) Then
                    StrSalvaItem = False
                    Exit Function
                End If
            End If
            adoResultado.MoveNext
        Loop
        adoResultado.Close
    End If
    
    strSQL = "DELETE FROM " & gstrTributoExercicio & " WHERE intTributo = "
    If mblnAlterandoAux Then
        strSQL = strSQL & intPkid
    Else
        strSQL = strSQL & glngPegaUltimaChave(gstrTributo, "pkid")
    End If
    If strPKId <> "" Then
        strSQL = strSQL & " AND PKID NOT IN(" & strPKId & ")"
    End If
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.Execute(strSQL) Then
        StrSalvaItem = True
    Else
        StrSalvaItem = False
        Exit Function
    End If
    
'    If lvw_Itens.ListItems.Count > 0 Then
'        strSQL = strSQL & IIf(bytDBType = Oracle, "End;", "")
'    End If
'    StrSalvaItem = strSQL
End Function

Private Sub PreencheListItens()
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = strSQL & "Select TE.*,IE.Strabreviatura,IE.Pkid As intIndex "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrTributo & " T, "
    strSQL = strSQL & gstrTributoExercicio & " TE, "
    strSQL = strSQL & gstrIndexadorEconomico & " IE "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "T.Pkid = TE.intTributo AND "
    strSQL = strSQL & "IE.Pkid = TE.Intindexadoreconomico AND "
    strSQL = strSQL & "TE.intTributo = " & txtPKId
    strSQL = strSQL & " Order By TE.intExercicio"
    lvw_Itens.ListItems.Clear
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                Do While Not .EOF
                    Set mobjLista = lvw_Itens.ListItems.Add(, , gstrENulo(!intExercicio))
                    mobjLista.SubItems(1) = gstrConvVrDoSql(gstrENulo(!dblValor), 6)
                    mobjLista.SubItems(2) = gstrConvVrDoSql(gstrENulo(!dblValorExcedente), 6)
                    mobjLista.SubItems(3) = gstrENulo(!Strabreviatura)
                    mobjLista.SubItems(4) = gstrENulo(!intIndex)
                    .MoveNext
                Loop
            End If
        End With
    End If
End Sub

Private Sub LimpaItens()
  txt_pkidTributoExercicio = ""
  txt_intExercicio = ""
  txt_dblvalor = ""
  txt_dblvalorexcedente = ""
  dbc_intindexadoreconomico.Text = ""
  Set dbc_intindexadoreconomico.RowSource = Nothing
  mblnAlterandoLista = False
End Sub

Private Sub txtintCodigoTributo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigoTributo
End Sub

Private Sub txtintCodigoTributo_LostFocus()
  txtintCodigoTributo.Text = Format(txtintCodigoTributo.Text, "000000")
End Sub

Private Sub txt_dblValorF_GotFocus()
    MarcaCampo txt_dblValorF
End Sub

Private Sub txt_dblValorF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorF
End Sub

Private Sub txt_dblValorF_LostFocus()
    txt_dblValorF = gstrConvVrDoSql(txt_dblValorF, 4)
End Sub

Private Sub txt_dblvalorexcedenteF_GotFocus()
    MarcaCampo txt_dblValorExcedenteF
End Sub

Private Sub txt_dblvalorexcedenteF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorExcedenteF
End Sub

Private Sub txt_dblvalorexcedenteF_LostFocus()
    txt_dblValorExcedenteF = gstrConvVrDoSql(txt_dblValorExcedenteF, 4)
End Sub

Private Sub txt_intTributoExercicioF_GotFocus()
    MarcaCampo txt_intTributoExercicioF
End Sub

Private Sub txt_intTributoExercicioF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intTributoExercicioF
End Sub

Private Sub txt_dblQuantidadeInicialF_GotFocus()
    MarcaCampo txt_dblQuantidadeInicialF
End Sub

Private Sub txt_dblQuantidadeInicialF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_dblQuantidadeInicialF
End Sub

Private Sub txt_dblQuantidadeInicialF_LostFocus()
    txt_dblQuantidadeInicialF = gstrConvVrDoSql(txt_dblQuantidadeInicialF, 2)
End Sub

Private Sub txt_dblQuantidadeFinalF_GotFocus()
    MarcaCampo txt_dblQuantidadeFinalF
End Sub

Private Sub txt_dblQuantidadeFinalF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_dblQuantidadeFinalF
End Sub

Private Sub txt_dblQuantidadeFinalF_LostFocus()
    txt_dblQuantidadeFinalF = gstrConvVrDoSql(txt_dblQuantidadeFinalF, 2)
End Sub

Private Sub tdb_FaixaValores_HeadClick(ByVal ColIndex As Integer)
   
   Select Case ColIndex
   Case 6
      xadbFaixaValores.QuickSort xadbFaixaValores.LowerBound(1), xadbFaixaValores.UpperBound(1), ColIndex, IIf(blnOrdenacaoAscF, XORDER_ASCEND, XORDER_DESCEND), XTYPE_STRING
   Case Else
      xadbFaixaValores.QuickSort xadbFaixaValores.LowerBound(1), xadbFaixaValores.UpperBound(1), ColIndex, IIf(blnOrdenacaoAscF, XORDER_ASCEND, XORDER_DESCEND), XTYPE_DOUBLE
   End Select
   
   blnOrdenacaoAscF = Not blnOrdenacaoAscF
   
   tdb_FaixaValores.Array = xadbFaixaValores
   tdb_FaixaValores.ReBind
   tdb_FaixaValores.Refresh
End Sub

Private Sub MantemForm2(strModoOperacao As String)
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset
Dim intFor As Integer
Dim strAux As String

Select Case UCase(strModoOperacao)
Case UCase(gstrNovo)
    txt_dblQuantidadeInicialF = ""
    txt_dblQuantidadeFinalF = ""
    txt_dblValorF = ""
    txt_dblValorExcedenteF = ""
    dbc_intIndexadorF.Text = ""
    mblnAlterandoF = False

Case UCase(gstrSalvar)
   mblnAlterandoF = False
   For intFor = 0 To xadbFaixaValores.UpperBound(1)
      If xadbFaixaValores(intFor, 9) = 1 Then
         If xadbFaixaValores(intFor, 0) <> "0" Then
            strSQL = "UPDATE " & gstrTributosFaixa & " SET "
            If InStr(1, xadbFaixaValores(intFor, 1), "*") <> 0 Then
               strSQL = strSQL & "intTributoExercicio=" & Left(xadbFaixaValores(intFor, 1), InStr(1, xadbFaixaValores(intFor, 1), "*") - 1) & ", "
            Else
               strSQL = strSQL & "intTributoExercicio=" & xadbFaixaValores(intFor, 1)
            End If
            strSQL = strSQL & "dblQuantidadeInicial=" & gstrConvVrParaSql(xadbFaixaValores(intFor, 2)) & ", "
            strSQL = strSQL & "dblQuantidadeFinal=" & gstrConvVrParaSql(xadbFaixaValores(intFor, 3)) & ", "
            strSQL = strSQL & "dblValor=" & gstrConvVrParaSql(xadbFaixaValores(intFor, 4)) & ", "
            strSQL = strSQL & "dblValorExcedente=" & gstrConvVrParaSql(xadbFaixaValores(intFor, 5)) & ", "
            strSQL = strSQL & "intIndexador=" & gstrConvVrParaSql(xadbFaixaValores(intFor, 7)) & ", "
         
            If bytDBType = EDatabases.Oracle Then
               strSQL = strSQL & "dtmDtAtualizacao=SYSDATE, lngCodUsr=" & glngCodUsr
            Else
               strSQL = strSQL & "dtmDtAtualizacao=GETDATE(), lngCodUsr=" & glngCodUsr
            End If
            strSQL = strSQL & " WHERE PKID=" & xadbFaixaValores(intFor, 0)
            
            If PKIDNaoDeletar = "" Then
               PKIDNaoDeletar = xadbFaixaValores(intFor, 0)
            Else
               PKIDNaoDeletar = PKIDNaoDeletar & "," & xadbFaixaValores(intFor, 0)
            End If
         Else
            strSQL = "INSERT INTO " & gstrTributosFaixa & "(intTributoExercicio,dblQuantidadeInicial,dblQuantidadeFinal"
            strSQL = strSQL & ",dblValor,dblValorExcedente,intIndexador,dtmDtAtualizacao,lngCodUsr) VALUES("
            strSQL = strSQL & xadbFaixaValores(intFor, 1) & "," & gstrConvVrParaSql(xadbFaixaValores(intFor, 2)) & ","
            strSQL = strSQL & gstrConvVrParaSql(xadbFaixaValores(intFor, 3)) & "," & gstrConvVrParaSql(xadbFaixaValores(intFor, 4)) & ","
            strSQL = strSQL & gstrConvVrParaSql(xadbFaixaValores(intFor, 5)) & "," & gstrConvVrParaSql(xadbFaixaValores(intFor, 7)) & ","
            
            If bytDBType = EDatabases.Oracle Then
               strSQL = strSQL & "SYSDATE," & glngCodUsr & ")"
            Else
               strSQL = strSQL & "GETDATE()," & glngCodUsr & ")"
            End If
         
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSQL) Then xadbFaixaValores(intFor, 9) = 1
            
            strSQL = "SELECT MAX(PKID) AS PKID FROM " & gstrTributosFaixa & " WHERE lngCodUsr = " & glngCodUsr
            If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
               If PKIDNaoDeletar = "" Then
                  PKIDNaoDeletar = adoResultado!Pkid
               Else
                  PKIDNaoDeletar = PKIDNaoDeletar & "," & adoResultado!Pkid
               End If
               strAux = adoResultado!Pkid
               xadbFaixaValores(intFor, 0) = strAux
            End If
            adoResultado.Close
         End If
         
      End If
   Next

  If txt_intTributoExercicioF.Tag <> "" Then
      strSQL = "DELETE FROM " & gstrTributosFaixa & " WHERE lngCodUsr = " & glngCodUsr & " AND intTributoExercicio = "
      strSQL = strSQL & txt_intTributoExercicioF.Tag
      If PKIDNaoDeletar <> "" Then
         strSQL = strSQL & " AND PKID NOT IN(" & PKIDNaoDeletar & ")"
      End If
      If gobjBanco.Execute(strSQL) Then PKIDNaoDeletar = ""
   End If
   MantemForm2 (gstrNovo)
   ExibeMensagem "As alterações foram salvas."
            
Case UCase(gstrIncluirItem)
   If txt_DescricaoF.Text = "" Or txt_intTributoExercicioF.Text = "" Then Exit Sub
   mblnAlterandoF = False
   If xadbFaixaValores.UpperBound(1) > -1 Then
      If xadbFaixaValores(xadbFaixaValores.UpperBound(1), 8) = "" Then
         intPosicao = 0
         xadbFaixaValores.ReDim 0, 0, 0, 9
      Else
         intPosicao = xadbFaixaValores(xadbFaixaValores.UpperBound(1), 8) + 1
         xadbFaixaValores.ReDim 0, xadbFaixaValores.Count(1), 0, 9
      End If
   Else
      xadbFaixaValores.ReDim 0, 0, 0, 9
      intPosicao = 0
   End If
    
   strSQL = "SELECT * FROM " & gstrTributosFaixa & " WHERE dblValor=" & gstrConvVrParaSql(txt_dblQuantidadeInicialF)
   strSQL = strSQL & " AND intTributoExercicio=" & txt_intTributoExercicioF.Tag
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      If adoResultado.EOF Then
         xadbFaixaValores(xadbFaixaValores.UpperBound(1), 0) = 0
      Else
         strAux = adoResultado!Pkid
         xadbFaixaValores(xadbFaixaValores.UpperBound(1), 0) = strAux
      End If
   End If
   strAux = txt_intTributoExercicioF.Tag
   xadbFaixaValores(xadbFaixaValores.UpperBound(1), 1) = strAux
   strAux = gstrConvVrDoSql(txt_dblQuantidadeInicialF, 2)
   xadbFaixaValores(xadbFaixaValores.UpperBound(1), 2) = strAux
   strAux = gstrConvVrDoSql(txt_dblQuantidadeFinalF, 2)
   xadbFaixaValores(xadbFaixaValores.UpperBound(1), 3) = strAux
   strAux = gstrConvVrDoSql(txt_dblValorF, 4)
   xadbFaixaValores(xadbFaixaValores.UpperBound(1), 4) = strAux
   strAux = gstrConvVrDoSql(txt_dblValorExcedenteF, 4)
   xadbFaixaValores(xadbFaixaValores.UpperBound(1), 5) = strAux
   strAux = dbc_intIndexadorF.Text
   xadbFaixaValores(xadbFaixaValores.UpperBound(1), 6) = strAux
   strAux = dbc_intIndexadorF.BoundText
   xadbFaixaValores(xadbFaixaValores.UpperBound(1), 7) = strAux
   xadbFaixaValores(xadbFaixaValores.UpperBound(1), 8) = intPosicao
   xadbFaixaValores(xadbFaixaValores.UpperBound(1), 9) = "1"
   
   tdb_FaixaValores.Array = xadbFaixaValores
   tdb_FaixaValores.ReBind
   tdb_FaixaValores.Refresh
   MantemForm2 (gstrNovo)
         
Case UCase(gstrExcluirItem)
   mblnAlterandoF = False
   If tdb_FaixaValores.BOF Or tdb_FaixaValores.EOF Then ExibeMensagem "É preciso selecionar um registro.": Exit Sub
   For intFor = 0 To xadbFaixaValores.UpperBound(1)
      If xadbFaixaValores(intFor, 8) = tdb_FaixaValores.Columns("Posicao") Then
         xadbFaixaValores.DeleteRows intFor
         Exit For
      End If
   Next
   tdb_FaixaValores.Array = xadbFaixaValores
   tdb_FaixaValores.ReBind
   tdb_FaixaValores.Refresh
   MantemForm2 (gstrNovo)

Case UCase(gstrPreencherLista)
   PreencherListaDeOpcoes Me.ActiveControl

End Select
End Sub

Private Sub CarregaRegistrosFaixaValores()
Dim adoResultado As New ADODB.Recordset
Dim strSQL As String
Dim strAux As String

strSQL = "SELECT TF.*, I.strAbreviatura, TE.intExercicio FROM " & gstrTributosFaixa & " TF, " & gstrTributoExercicio & " TE, " & gstrIndexadorEconomico & " I WHERE "
If bytDBType = EDatabases.Oracle Then
   strSQL = strSQL & "(+) I.PKID = TF.intIndexador "
Else
   strSQL = strSQL & "I.PKID =* TF.intIndexador "
End If
strSQL = strSQL & "AND TF.intTributoExercicio = TE.PKID AND TF.intTributoExercicio=" & txt_intTributoExercicioF.Tag

Set gobjBanco = New clsBanco
If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
   With adoResultado
      xadbFaixaValores.Clear
      xadbFaixaValores.ReDim 0, 0, 0, 9
      If Not .EOF Then
         intPosicao = 0
         Do While Not .EOF
            strAux = !Pkid
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 0) = strAux
            
            strAux = !intTributoExercicio & "*" & !intExercicio
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 1) = gstrConvVrDoSql(strAux, 2)
            
            strAux = IIf(IsNull(!dblQuantidadeInicial), "", !dblQuantidadeInicial)
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 2) = gstrConvVrDoSql(strAux, 2)
            
            strAux = IIf(IsNull(!dblQuantidadeFinal), "", !dblQuantidadeFinal)
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 3) = gstrConvVrDoSql(strAux, 2)
            
            strAux = IIf(IsNull(!dblValor), "", !dblValor)
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 4) = gstrConvVrDoSql(strAux, 4)
            
            strAux = IIf(IsNull(!dblValorExcedente), "", !dblValorExcedente)
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 5) = gstrConvVrDoSql(strAux, 4)
            
            strAux = IIf(IsNull(!Strabreviatura), "", !Strabreviatura)
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 6) = strAux
            
            strAux = IIf(IsNull(!intIndexador), "", !intIndexador)
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 7) = strAux
            
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 8) = intPosicao
            xadbFaixaValores(xadbFaixaValores.UpperBound(1), 9) = "1"
            
            .MoveNext
            If Not .EOF Then
               xadbFaixaValores.ReDim 0, xadbFaixaValores.Count(1), 0, 9
            End If
         Loop
         
      End If
      tdb_FaixaValores.Array = xadbFaixaValores
      tdb_FaixaValores.ReBind
      tdb_FaixaValores.Refresh
      
   End With
End If
            
End Sub
