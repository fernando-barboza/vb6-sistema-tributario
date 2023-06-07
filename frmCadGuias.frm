VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadGuias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guias"
   ClientHeight    =   6030
   ClientLeft      =   2400
   ClientTop       =   1665
   ClientWidth     =   8415
   HelpContextID   =   18
   Icon            =   "frmCadGuias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5865
      Left            =   90
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   10345
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Guias"
      TabPicture(0)   =   "frmCadGuias.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStrConta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNumero"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbdtEmissao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbldtVencimento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblSTRCODBARRA"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDBLVALOR"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tdb_SubLista"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbc_strConta"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dbcintContaBancaria"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "tdb_Lista"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_ContaBacaria"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPKId"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dtmDtEmissao"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtstrCodigo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Dtmdtvencimento"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_STRCODBARRA"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_DBLVALOR"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.TextBox txt_DBLVALOR 
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
         Left            =   6690
         MaxLength       =   50
         TabIndex        =   5
         Top             =   900
         Width           =   1425
      End
      Begin VB.TextBox txt_STRCODBARRA 
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
         Left            =   1200
         MaxLength       =   44
         TabIndex        =   6
         Top             =   1290
         Width           =   4095
      End
      Begin VB.TextBox Dtmdtvencimento 
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
         Left            =   7125
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtstrCodigo 
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
         Left            =   1215
         MaxLength       =   10
         OLEDragMode     =   1  'Automatic
         TabIndex        =   0
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox dtmDtEmissao 
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
         Left            =   4275
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   1005
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_ContaBacaria 
         Height          =   315
         Left            =   5760
         Picture         =   "frmCadGuias.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "585"
         ToolTipText     =   "Ativa Cadastro de Contas Bancarias"
         Top             =   900
         Width           =   360
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2025
         Left            =   150
         TabIndex        =   8
         Top             =   3645
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3572
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
         Columns(1).Caption=   "Número"
         Columns(1).DataField=   "intNumero"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Emissão"
         Columns(2).DataField=   "dtmDtEmissao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Vencimento"
         Columns(3).DataField=   "dtmDtVencimento"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Conta"
         Columns(4).DataField=   "strConta"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Descrição"
         Columns(5).DataField=   "strDescricao"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Valor"
         Columns(6).DataField=   "dblValor"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "CodigoBarra"
         Columns(7).DataField=   "STRCODBARRA"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "intUtilizacao"
         Columns(8).DataField=   "intUtilizacao"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1270"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1191"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=1826"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1746"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=1958"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1879"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=2117"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2037"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=4445"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=4366"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=2249"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2170"
         Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(42)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(45)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(48)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(49)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(51)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(52)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Named:id=33:Normal"
         _StyleDefs(73)  =   ":id=33,.parent=0"
         _StyleDefs(74)  =   "Named:id=34:Heading"
         _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   ":id=34,.wraptext=-1"
         _StyleDefs(77)  =   "Named:id=35:Footing"
         _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   "Named:id=36:Selected"
         _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=37:Caption"
         _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(83)  =   "Named:id=38:HighlightRow"
         _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(85)  =   "Named:id=39:EvenRow"
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintContaBancaria 
         Height          =   315
         Left            =   2790
         TabIndex        =   4
         Top             =   900
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strConta 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   900
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_SubLista 
         Height          =   1695
         Left            =   150
         TabIndex        =   7
         Top             =   1740
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2990
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Composição da Receita"
         Columns(0).DataField=   "strcomposicaodareceita"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Exercício"
         Columns(1).DataField=   "intExercicio"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Inscrição Cadastral"
         Columns(2).DataField=   "strInscricao"
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Aviso"
         Columns(3).DataField=   "strNumeroAviso"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Parcela"
         Columns(4).DataField=   "intParcela"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Valor"
         Columns(5).DataField=   "dblValor"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "intUtilizacao"
         Columns(6).DataField=   "intUtilizacao"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4180"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4101"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1349"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1270"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=3254"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1640"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1561"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=1111"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1032"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=2"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=2328"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2249"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(42)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(43)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblDBLVALOR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   6270
         TabIndex        =   17
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblSTRCODBARRA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código Barras"
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lbldtVencimento 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vencimento"
         Height          =   195
         Left            =   5550
         TabIndex        =   15
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label lbdtEmissao 
         AutoSize        =   -1  'True
         Caption         =   "Data da Emissão"
         Height          =   195
         Left            =   2790
         TabIndex        =   14
         Top             =   570
         Width           =   1200
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   585
         TabIndex        =   13
         Top             =   570
         Width           =   555
      End
      Begin VB.Label lblStrConta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Conta Bancaria"
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCadGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando     As Boolean
    Dim mblnSelecionou    As Boolean
    Dim mblnClickOk       As Boolean
    Dim bytOrdenacao      As Byte
    Dim blnOrdenacaoAsc   As Boolean

Private Function strQuery() As String

'******************************************************************************************
' Data: 27/03/2003
' Alteração: - Substituição do comando CONVERT do SQL Server pela função gstrCONVERT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSQL  As String
   
   strSQL = ""
   
   Select Case bytOrdenacao
   
      Case Is = 1
         
      Case Is = 2
      
      Case Is = 3
      
   End Select
   
   strQuery = strSQL
    
End Function

Private Function strQueryAplicar() As String
End Function

Private Sub cmd_ContaBacaria_Click()
    CarregaForm frmCadContasBancarias, dbcintContaBancaria
End Sub

Private Sub dbc_strConta_Click(Area As Integer)
    If Area = 2 And dbc_strConta.MatchedWithList Then
        PreencherListaDeOpcoes dbcintContaBancaria, dbc_strConta.BoundText
    End If
    DropDownDataCombo dbc_strConta, Me, Area
End Sub

Private Sub dbcintContaBancaria_Click(Area As Integer)
    If Area = 2 And dbcintContaBancaria.MatchedWithList Then
        PreencherListaDeOpcoes dbc_strConta, dbcintContaBancaria.BoundText
    End If
    DropDownDataCombo dbcintContaBancaria, Me, Area
End Sub


Private Sub dtmDtEmissao_GotFocus()
    MarcaCampo dtmDtEmissao
End Sub

Private Sub dtmDtEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", dtmDtEmissao
End Sub

Private Sub dtmDtEmissao_LostFocus()
    dtmDtEmissao = gstrDataFormatada(dtmDtEmissao)
End Sub

Private Sub Dtmdtvencimento_GotFocus()
    MarcaCampo Dtmdtvencimento
End Sub

Private Sub Dtmdtvencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", Dtmdtvencimento
End Sub

Private Sub Dtmdtvencimento_LostFocus()
    Dtmdtvencimento = gstrDataFormatada(Dtmdtvencimento)
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1115
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
   bytOrdenacao = 1: blnOrdenacaoAsc = True
   dbc_strConta.Tag = strQueryConta("strConta") & " ;strConta"
   dbcintContaBancaria.Tag = strQueryConta("strDescricao") & " ;strDescricao"
   
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub tdb_Lista_Click()
    mblnClickOk = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 6 Then
        Value = gstrConvVrDoSql(Value)
    End If
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
    mblnClickOk = False
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)

    Select Case tdb_Lista.Col
    
    '    Case Is = 1
    '        CaracterValido KeyAscii, "N", tdb_Lista
        Case Is = 2
            CaracterValido KeyAscii, "D", tdb_Lista
        Case Is = 3
            CaracterValido KeyAscii, "D", tdb_Lista
    '    Case Is = 4
    '        CaracterValido KeyAscii, "A", tdb_Lista
    '    Case Is = 5
    '        CaracterValido KeyAscii, "A", tdb_Lista
    '    Case Is = 6
    '        CaracterValido KeyAscii, "N", tdb_Lista
    End Select
    
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            PreencheCampos
            LeDaTabelaParaObj "", tdb_SubLista, strLocalizarSub
            gCorLinhaSelecionada tdb_Lista
            mblnSelecionou = True
            mblnAlterando = True
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
        
    If UCase(strModoOperacao) = gstrLocalizar Then
        mblnClickOk = True
        LeDaTabelaParaObj "", tdb_Lista, strLocalizar(True)
    ElseIf UCase(strModoOperacao) = gstrPreencherLista Then
        If Me.ActiveControl.Name = "dbc_strConta" Then
            LeDaTabelaParaObj gstrContaBancaria, dbc_strConta, strQueryConta("strConta", , Trim(Me.ActiveControl.Text))
        ElseIf Me.ActiveControl.Name = "dbcintContaBancaria" Then
            LeDaTabelaParaObj gstrContaBancaria, dbcintContaBancaria, strQueryConta("strDescricao", , Trim(Me.ActiveControl.Text))
        End If
    ElseIf UCase(strModoOperacao) = gstrNovo Then
        Limpa_Controles Me, True, False, False, True, False
        Set tdb_SubLista.DataSource = Nothing
        dbc_strConta.ListField = ""
        dbcintContaBancaria.ListField = ""
        TrocaCorObjeto txtstrCodigo, False
        TrocaCorObjeto dtmDtEmissao, False
        TrocaCorObjeto Dtmdtvencimento, False
        TrocaCorObjeto dbc_strConta, False
        TrocaCorObjeto dbcintContaBancaria, False
        TrocaCorObjeto txt_DBLVALOR, False
        TrocaCorObjeto txt_STRCODBARRA, False
    ElseIf UCase(strModoOperacao) = gstrRefresh Then
        LeDaTabelaParaObj "", tdb_Lista, strLocalizar(False)
    ElseIf UCase(strModoOperacao) = gstrImprimir Then
        ImprimeRelatorio rptGuias, strQueryRelatorio
    End If
    
    'ToolBarGeral strModoOperacao, gstrBairro, mblnAlterando, tdb_Lista, _
    '             Me, mobjAux, strQuery, strQueryAplicar, rptBairro, strQueryRelatorio
                 
End Sub


Function strQueryConta(strCampo As String, Optional strPKId As String, Optional strfiltro As String) As String
   
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select Pkid, " & strCampo & " From "
    strSQL = strSQL & gstrContaBancaria
    
    If Trim(strPKId) <> "" Then strSQL = strSQL & " Where Pkid = " & strPKId
    
    If Trim(strfiltro) <> "" And Trim(strPKId) <> "" Then
        strSQL = strSQL & " Where Pkid = " & strPKId
        strSQL = strSQL & " UPPER(" & strCampo & " Like '" & UCase(strfiltro) & "%'"
    ElseIf Trim(strfiltro) <> "" And Trim(strPKId) = "" Then
        strSQL = strSQL & " Where "
        strSQL = strSQL & " UPPER(" & strCampo & ") Like '" & UCase(strfiltro) & "%'"
    End If
    strSQL = strSQL & " Order by strDescricao"
    
    strQueryConta = strSQL
   
End Function

Private Sub tdb_SubLista_Click()
    gCorLinhaSelecionada tdb_SubLista
End Sub

Private Sub tdb_SubLista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 2 Then
        Value = gstrFormataInscricao(CStr(Value), tdb_SubLista.Columns("intUtilizacao"))
    End If
End Sub

Private Sub txt_DBLVALOR_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_DBLVALOR
End Sub

Private Sub txt_DBLVALOR_LostFocus()
    txt_DBLVALOR = gstrConvVrDoSql(txt_DBLVALOR)
End Sub

Private Function strQueryRelatorio() As String
'RESPONSÁVEL    LEANDRO 30/06/2004

Dim strSQL As String

strSQL = ""
strSQL = strSQL & "Select "
    strSQL = strSQL & " G.INTNUMERO Numero,"
    strSQL = strSQL & " G.DTMDTEMISSAO Emissao,"
    strSQL = strSQL & " G.DTMDTVENCIMENTO Vencimento,"
    strSQL = strSQL & " CB.Strconta Conta,"
    strSQL = strSQL & " CB.Strdescricao Descricao,"
    strSQL = strSQL & " G.DBLVALOR Valor"
    
strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrGuias & " G, "
    strSQL = strSQL & gstrContaBancaria & " CB"
    
strSQL = strSQL & " WHERE"
    strSQL = strSQL & " G.Intcontabancaria " & strOUTJSQLServer & "= " & strOUTJOracle & " CB.pkid "

strQueryRelatorio = strSQL

End Function

Private Function strLocalizar(Optional blnFiltro As Boolean) As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "G.PKID, "
    strSQL = strSQL & "G.INTNUMERO, "
    strSQL = strSQL & "G.DTMDTEMISSAO, "
    strSQL = strSQL & "G.DTMDTVENCIMENTO, "
    strSQL = strSQL & "CB.Strconta, "
    strSQL = strSQL & "CB.Strdescricao, "
    strSQL = strSQL & "G.STRCODBARRA, "
    strSQL = strSQL & "G.DBLVALOR, "
    strSQL = strSQL & "LA.intUtilizacao "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrGuias & " G, "
    strSQL = strSQL & gstrContaBancaria & " CB, "
    
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLancamentoValor & " LV, "
    strSQL = strSQL & gstrLancamentoGuias & " LG "
    
    strSQL = strSQL & "Where "
    strSQL = strSQL & "G.PKID = LG.INTGUIAS AND "
    strSQL = strSQL & "LV.INTLANCAMENTOALFA = LA.PKID AND "
    strSQL = strSQL & "LG.INTLANCAMENTOVALOR = LV.PKID AND "
    strSQL = strSQL & "G.Intcontabancaria" & strOUTJSQLServer & "= CB.pkid" & strOUTJOracle & " "
    
    If blnFiltro = True Then
        If Trim(txtstrCodigo) <> "" Then
            strSQL = strSQL & " And G.INTNUMERO = " & txtstrCodigo
        End If
        
        If Len(dtmDtEmissao) = 10 Then
            strSQL = strSQL & " And G.DTMDTEMISSAO = " & gstrConvDtParaSql(dtmDtEmissao)
        End If
        
        If Len(Dtmdtvencimento) = 10 Then
            strSQL = strSQL & " And G.DTMDTVENCIMENTO = " & gstrConvDtParaSql(Dtmdtvencimento)
        End If
        
        If Trim(dbc_strConta) <> "" And dbc_strConta.MatchedWithList Then
            strSQL = strSQL & " And CB.pkid = " & dbc_strConta.BoundText
        ElseIf Trim(dbcintContaBancaria) <> "" And dbcintContaBancaria.MatchedWithList Then
            strSQL = strSQL & " And CB.pkid = " & dbcintContaBancaria.BoundText
        End If
        
        If Trim(txt_DBLVALOR) <> "" Then
            strSQL = strSQL & " And G.DBLVALOR = " & gstrConvVrParaSql(txt_DBLVALOR)
        End If
        
        If Trim(txt_STRCODBARRA) <> "" Then
            strSQL = strSQL & " And UPPER(G.STRCODBARRA) Like'" & UCase(txt_STRCODBARRA) & "'"
        End If
    End If
    
    strSQL = strSQL & "GROUP BY "
    strSQL = strSQL & "G.PKID, "
    strSQL = strSQL & "G.INTNUMERO, "
    strSQL = strSQL & "G.DTMDTEMISSAO, "
    strSQL = strSQL & "G.DTMDTVENCIMENTO, "
    strSQL = strSQL & "CB.Strconta, "
    strSQL = strSQL & "CB.Strdescricao, "
    strSQL = strSQL & "G.STRCODBARRA, "
    strSQL = strSQL & "G.DBLVALOR, "
    strSQL = strSQL & "LA.intUtilizacao "
    
    Select Case bytOrdenacao
    
        Case Is = 1
            strSQL = strSQL & " ORDER BY G.INTNUMERO " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSQL = strSQL & " ORDER BY G.dtmDtEmissao " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strSQL = strSQL & " ORDER BY G.DTMDTVENCIMENTO " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 4
            strSQL = strSQL & " ORDER BY CB.Strconta " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 5
            strSQL = strSQL & " ORDER BY CB.Strdescricao " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 6
            strSQL = strSQL & " ORDER BY G.DBLVALOR " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    
    strLocalizar = strSQL



End Function
Private Function strLocalizarSub() As String
Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "LA.strcomposicaodareceita, "
    strSQL = strSQL & "LA.intexercicio, "
    strSQL = strSQL & "LA.intUtilizacao, "
    strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(gstrConvVrDoSql(tdb_Lista.Columns("intUtilizacao").Value, , , True))) & "strInscricao, "
    strSQL = strSQL & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso,"
    strSQL = strSQL & "LV.intparcela, "
    strSQL = strSQL & "(lg.dblvalorprincipal + lg.dblvalormulta + lg.dblvalorjuros + lg.dblvalorcorrecao) - lg.dblvalordesconto dblValor "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrGuias & " G, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLancamentoValor & " LV, "
    strSQL = strSQL & gstrLancamentoGuias & " LG "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "G.Pkid   = LG.Intguias AND "
    strSQL = strSQL & "LV.Pkid  = LG.Intlancamentovalor AND "
    strSQL = strSQL & "LA.Pkid = LV.Intlancamentoalfa AND "
    strSQL = strSQL & "G.Pkid = " & txtPKId
    
    strLocalizarSub = strSQL

End Function
    
Private Sub PreencheCampos()
With tdb_Lista
    
    txtstrCodigo = .Columns("INTNUMERO").Value
    dtmDtEmissao = .Columns("dtmDtEmissao").Value
    Dtmdtvencimento = .Columns("DTMDTVENCIMENTO").Value
    dbc_strConta.Text = .Columns("Strconta").Value
    dbcintContaBancaria.Text = .Columns("Strdescricao").Value
    txt_DBLVALOR = gstrConvVrDoSql(.Columns("DBLVALOR").Value)
    txt_STRCODBARRA = .Columns("STRCODBARRA").Value
    
    TrocaCorObjeto txtstrCodigo, True
    TrocaCorObjeto dtmDtEmissao, True
    TrocaCorObjeto Dtmdtvencimento, True
    TrocaCorObjeto dbc_strConta, True
    TrocaCorObjeto dbcintContaBancaria, True
    TrocaCorObjeto txt_DBLVALOR, True
    TrocaCorObjeto txt_STRCODBARRA, True
    
End With

End Sub

