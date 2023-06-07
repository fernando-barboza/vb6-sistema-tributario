VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadAlteracaoEndContribuinte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Endereço do Contribuinte"
   ClientHeight    =   5160
   ClientLeft      =   930
   ClientTop       =   4320
   ClientWidth     =   9645
   FillColor       =   &H80000012&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9645
   Begin VB.TextBox txtstrNome 
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
      Left            =   5670
      MaxLength       =   50
      TabIndex        =   3
      Top             =   90
      Width           =   3720
   End
   Begin TabDlg.SSTab tab_3DCorrespondencia 
      Height          =   2445
      Left            =   195
      TabIndex        =   4
      Top             =   480
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   4313
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Endereço Notificação"
      TabPicture(0)   =   "frmCadAlteracaoEndContribuinte.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDistritoC"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintMunicipioC"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintBairroC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintLogradouroC"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintNumeroC"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrComplementoC"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblintUFC"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblintCepC"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dbcstrLogradouroC"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dbcintTituloLogradouro"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbcintTipoLogradouro"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dbcintUFC"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dbcintMunicipioC"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtintCodigoLogradouro"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtstrDistritoC"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtstrBairroC"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtintNumeroC"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtstrComplementoC"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtintCepC"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmd_TipoLogradouro"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmd_TituloLogradouro"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmd_MunicipioC"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      Begin VB.CommandButton cmd_MunicipioC 
         Height          =   315
         Left            =   4650
         Picture         =   "frmCadAlteracaoEndContribuinte.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Ativa Cadastro de Municipios"
         Top             =   1380
         Width           =   360
      End
      Begin VB.CommandButton cmd_TituloLogradouro 
         Height          =   315
         Left            =   3660
         Picture         =   "frmCadAlteracaoEndContribuinte.frx":013A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Ativa Cadastro de Titulo de Logradouro"
         Top             =   600
         Width           =   360
      End
      Begin VB.CommandButton cmd_TipoLogradouro 
         Height          =   315
         Left            =   1830
         Picture         =   "frmCadAlteracaoEndContribuinte.frx":0258
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Ativa Cadastro de Tipo de Logradouro"
         Top             =   600
         Width           =   360
      End
      Begin VB.TextBox txtintCepC 
         Height          =   285
         Left            =   7890
         MaxLength       =   9
         TabIndex        =   24
         Top             =   1425
         Width           =   1080
      End
      Begin VB.TextBox txtstrComplementoC 
         Height          =   285
         Left            =   2790
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1020
         Width           =   1260
      End
      Begin VB.TextBox txtintNumeroC 
         Height          =   285
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   13
         Top             =   1020
         Width           =   855
      End
      Begin VB.TextBox txtstrBairroC 
         Height          =   285
         Left            =   5595
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1020
         Width           =   3375
      End
      Begin VB.TextBox txtstrDistritoC 
         Height          =   285
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1800
         Width           =   3525
      End
      Begin VB.TextBox txtintCodigoLogradouro 
         Height          =   285
         Left            =   4020
         MaxLength       =   8
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dbcintMunicipioC 
         Height          =   315
         Left            =   1110
         TabIndex        =   19
         Top             =   1380
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintUFC 
         Height          =   315
         Left            =   5595
         TabIndex        =   22
         Top             =   1395
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintTipoLogradouro 
         Height          =   315
         Left            =   1110
         TabIndex        =   6
         Top             =   585
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintTituloLogradouro 
         Height          =   315
         Left            =   2220
         TabIndex        =   8
         Top             =   585
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcstrLogradouroC 
         Height          =   315
         Left            =   4830
         TabIndex        =   11
         Top             =   570
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblintCepC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cep"
         Height          =   195
         Left            =   7545
         TabIndex        =   23
         Top             =   1515
         Width           =   285
      End
      Begin VB.Label lblintUFC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "UF"
         Height          =   195
         Left            =   5250
         TabIndex        =   21
         Top             =   1500
         Width           =   210
      End
      Begin VB.Label lblstrComplementoC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Compl."
         Height          =   195
         Left            =   2190
         TabIndex        =   14
         Top             =   1110
         Width           =   480
      End
      Begin VB.Label lblintNumeroC 
         AutoSize        =   -1  'True
         Caption         =   "Nº"
         Height          =   195
         Left            =   780
         TabIndex        =   12
         Top             =   1110
         Width           =   180
      End
      Begin VB.Label lblintLogradouroC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Logradouro"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   810
      End
      Begin VB.Label lblintBairroC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   5040
         TabIndex        =   16
         Top             =   1110
         Width           =   405
      End
      Begin VB.Label lblintMunicipioC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Município"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   1515
         Width           =   705
      End
      Begin VB.Label lblstrDistritoC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         Height          =   195
         Left            =   450
         TabIndex        =   25
         Top             =   1890
         Width           =   480
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1890
      Left            =   210
      TabIndex        =   27
      Top             =   3060
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   3334
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
      Columns(1).Caption=   "Inscrição Cadastral"
      Columns(1).DataField=   "strInscricao"
      Columns(1).NumberFormat=   "FormatText Event"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Contribuinte"
      Columns(2).DataField=   "strNome"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "CNPJ / CPF"
      Columns(3).DataField=   "strCNPJCPF"
      Columns(3).NumberFormat=   "FormatText Event"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "intContribuinte"
      Columns(4).DataField=   "intContribuinte"
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
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2672"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2593"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=8943"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=8864"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=3651"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3572"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2302"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2223"
      Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(29)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
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
   Begin MSMask.MaskEdBox mskstrInscricao 
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   90
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   24
      PromptChar      =   " "
   End
   Begin VB.Label lbl_strNome 
      AutoSize        =   -1  'True
      Caption         =   "Contribuinte"
      Height          =   195
      Left            =   4770
      TabIndex        =   2
      Top             =   135
      Width           =   840
   End
   Begin VB.Label lblstrInscricao 
      AutoSize        =   -1  'True
      Caption         =   "Inscrição Cadastral"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   135
      Width           =   1350
   End
End
Attribute VB_Name = "frmCadAlteracaoEndContribuinte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnClick         As Boolean
Dim mblnSelecionou    As Boolean
Dim mobjAux           As Object
Dim mblnPrimeiraVez    As Boolean

Private Sub cmd_MunicipioC_Click()
    ChamaFormCadastro frmCadCidade, dbcintMunicipioC
End Sub

Private Sub cmd_TipoLogradouro_Click()
    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouro
End Sub

Private Sub cmd_TituloLogradouro_Click()
    CarregaForm frmCadTituloLogradouro, dbcintTituloLogradouro
End Sub

Private Sub dbcintTituloLogradouro_GotFocus()
    MarcaCampo dbcintTipoLogradouro
End Sub

Private Sub dbcstrLogradouroC_Click(Area As Integer)
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    
    On Error GoTo TrataErro
    
    If (Area = 2 Or Area = 1) And dbcstrLogradouroC.BoundText <> "" And IsNumeric(dbcstrLogradouroC.BoundText) Then
       strSql = strSql & "SELECT "
       strSql = strSql & "TL.strDescricao strTituloLogradouro, "
       strSql = strSql & "TP.strSigla strTipoLogradouro, "
       strSql = strSql & "BA.strDescricao strBairro, "
       strSql = strSql & "MU.strDescricao strMunicipio, "
       strSql = strSql & "UF.strSigla strUF, "
       strSql = strSql & "LO.intCEP intCEP "
       
       strSql = strSql & "FROM "
       strSql = strSql & gstrLogradouro & " LO, "
       strSql = strSql & gstrTituloLogradouro & " TL, "
       strSql = strSql & gstrTipoLogradouro & " TP, "
       strSql = strSql & gstrBairro & " BA, "
       strSql = strSql & gstrCidade & " MU, "
       strSql = strSql & gstrUF & " UF "
       
       strSql = strSql & "WHERE "
       strSql = strSql & "TL.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LO.intTituloLogradouro AND "
       strSql = strSql & "TP.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LO.intTipoLogradouro AND "
       strSql = strSql & "BA.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LO.intBairro AND "
       strSql = strSql & "MU.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " BA.intMunicipio AND "
       strSql = strSql & "UF.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " MU.intUF AND "
       strSql = strSql & "LO.pkID = " & dbcstrLogradouroC.BoundText
       
       Set adoResultado = New ADODB.Recordset
       Set gobjBanco = New clsBanco
       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
          If Not adoResultado.EOF Then
             If Not IsNull(adoResultado("strTituloLogradouro")) Then
                dbcintTituloLogradouro.Text = adoResultado("strTituloLogradouro")
                PreencherListaDeOpcoes dbcintTituloLogradouro
                dbcintTituloLogradouro.Text = adoResultado("strTituloLogradouro")
             End If
             
             If Not IsNull(adoResultado("strTipoLogradouro")) Then
                dbcintTipoLogradouro.Text = adoResultado("strTipoLogradouro")
                PreencherListaDeOpcoes dbcintTipoLogradouro
                dbcintTipoLogradouro.Text = adoResultado("strTipoLogradouro")
             End If
             
             If Not IsNull(adoResultado("strMunicipio")) Then
                dbcintMunicipioC.Text = adoResultado("strMunicipio")
                PreencherListaDeOpcoes dbcintMunicipioC
                dbcintMunicipioC.Text = adoResultado("strMunicipio")
             End If
             
             If Not IsNull(adoResultado("strUF")) Then
                dbcintUFC.Text = adoResultado("strUF")
                PreencherListaDeOpcoes dbcintUFC
                dbcintUFC.Text = adoResultado("strUF")
             End If
             
             txtstrBairroC.Text = gstrENulo(adoResultado("strBairro"))
             
             txtintCepC.Text = gstrENulo(adoResultado("intCEP"))
             txtintCepC.Text = gstrCEPFormatado(txtintCepC.Text)
          End If
       End If
    End If
    
    txtintCodigoLogradouro.Text = ""
    txtstrDistritoC.Text = ""
    txtintNumeroC.Text = ""
    txtstrComplementoC.Text = ""
    
Exit Sub
TrataErro:
    
    If Err.Number = 7 Then 'Out of Memory (dbcstrLogradouroC.BoundText)
       Exit Sub
    Else
       ExibeMensagem Err.Number & " - " & Err.Description
    End If
    
End Sub

Private Sub dbcstrLogradouroC_GotFocus()
    MarcaCampo dbcstrLogradouroC
End Sub

Private Sub dbcstrLogradouroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcstrLogradouroC
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = 1262
    
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
    dbcintTipoLogradouro.Tag = gstrQueryDataComboTipoLogradouro & ";strSigla"
    dbcintMunicipioC.Tag = gstrQueryDataComboMunicipio & ";strDescricao"
    dbcintUFC.Tag = gstrQueryDataComboUF & ";strSigla"
    dbcintTituloLogradouro.Tag = gstrQueryDataComboTituloLogradouro & ";strDescricao"
   
    VerificaMascaraInscricao
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Function blnDadosOk()
    
    blnDadosOk = False
    If Len(Trim(dbcstrLogradouroC.Text)) = 0 Then
        ExibeMensagem "O campo logradouro deve ser preenchido corretamente."
        dbcstrLogradouroC.SetFocus
        Exit Function
    ElseIf Trim(txtintNumeroC) = "" Then
        ExibeMensagem "O campo número deve ser preenchido corretamente."
        txtintNumeroC.SetFocus
        Exit Function
    ElseIf Trim(txtstrBairroC) = "" Then
        ExibeMensagem "O campo bairro deve ser preenchido corretamente."
        txtstrBairroC.SetFocus
        Exit Function
    ElseIf Not dbcintMunicipioC.MatchedWithList Then
        ExibeMensagem "O campo município deve ser preenchido corretamente."
        dbcintMunicipioC.SetFocus
        Exit Function
    ElseIf Not dbcintUFC.MatchedWithList Then
        ExibeMensagem "O campo UF deve ser preenchido corretamente."
        dbcintUFC.SetFocus
        Exit Function
    ElseIf Trim(txtintCepC) = "" Then
        ExibeMensagem "O campo CEP deve ser preenchido corretamente."
        txtintCepC.SetFocus
        Exit Function
    ElseIf Trim(dbcintTipoLogradouro.Text) <> "" And Not dbcintTipoLogradouro.MatchedWithList Then
        ExibeMensagem "O campo tipo do logradouro deve ser preenchido corretamente."
        dbcintTipoLogradouro.SetFocus
        Exit Function
    ElseIf Trim(dbcintTituloLogradouro.Text) <> "" And Not dbcintTituloLogradouro.MatchedWithList Then
        ExibeMensagem "O campo título do logradouro deve ser preenchido corretamente."
        dbcintTituloLogradouro.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
    
End Function

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
Dim strInscricao As String
    On Error Resume Next
    Select Case ColIndex
        Case 1
            strInscricao = Value
            Value = gstrFormataInscricao(strInscricao, TYP_IMOBILIARIA)
        Case 4
            Value = gstrCGCCPFFormatado(CStr(Value))
    End Select
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    mblnClick = False
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Lista
        If Not .EOF And Not .BOF Then
            If mblnClick Then
                If mblnPrimeiraVez Then
                    
                    Screen.MousePointer = vbHourglass

                    mblnClick = False
                    Limpa_Controles Me, False, False, False, False, True
                    dbcstrLogradouroC.BoundText = ""
                    
                    mblnSelecionou = True
                    'dbcintLogradouro_Click 2
                    
                    Set gobjBanco = New clsBanco
                    
                    LeDaTabelaParaObj gstrImobiliario, Me, "SELECT CO.*, " & gstrRIGHT("IM.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao FROM " & gstrImobiliario & " IM, " & gstrContribuinte & " CO WHERE IM.Pkid = " & .Columns("PKID").Value & " AND CO.Pkid = IM.intContribuinte "
                    
                    gCorLinhaSelecionada tdb_Lista

                    'dbcintLogradouro_Click 2
                    
                    Screen.MousePointer = vbDefault
                    
                End If
            End If
        End If
    End With
    
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    mblnClick = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrImobiliario
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    mblnClick = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Lista
End Sub

Private Sub txtintCepC_LostFocus()
    txtintCepC = gstrCEPFormatado(txtintCepC)
    dbcintTipoLogradouro.Text = ""
    dbcintTituloLogradouro.Text = ""
    txtintCodigoLogradouro.Text = ""
    txtstrBairroC.Text = ""
    txtstrDistritoC.Text = ""
    txtintNumeroC.Text = ""
    txtstrComplementoC.Text = ""
    dbcstrLogradouroC.Tag = gstrQueryLogradouro & ";L.strDescricao"
    CepLogradouro txtintCepC, dbcstrLogradouroC, txtstrBairroC, dbcintMunicipioC, dbcintUFC, dbcintTipoLogradouro, dbcintTituloLogradouro, , True, False, True, True, True, True
    dbcstrLogradouroC.Tag = ""
    
End Sub

Private Sub txtintCepC_GotFocus()
    MarcaCampo txtintCepC
End Sub

Private Sub txtintCepC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepC
End Sub

Private Sub dbcintTipoLogradouro_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintTipoLogradouro, Me, Area
End Sub

Private Sub dbcintTipoLogradouro_GotFocus()
    MarcaCampo dbcintTipoLogradouro
End Sub

Private Sub dbcintTipoLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTipoLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouro_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintTituloLogradouro, Me, Area
End Sub

Private Sub dbcintTituloLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTituloLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTituloLogradouro
End Sub

Private Sub txtintCodigoLogradouro_GotFocus()
    MarcaCampo txtintCodigoLogradouro
End Sub

Private Sub txtintCodigoLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintNumeroC_GotFocus()
    MarcaCampo txtintNumeroC
End Sub

Private Sub txtintNumeroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumeroC
End Sub

Private Sub txtstrBairroC_GotFocus()
    MarcaCampo txtstrBairroC
End Sub

Private Sub txtstrComplementoC_GotFocus()
    MarcaCampo txtstrComplementoC
End Sub

Private Sub txtstrComplementoC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplementoC
End Sub

Private Sub txtstrBairroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrBairroC
End Sub

Private Sub dbcintMunicipioC_Click(Area As Integer)
    DropDownDataCombo dbcintMunicipioC, Me, Area
End Sub

Private Sub dbcintMunicipioC_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintMunicipioC, Me, , KeyCode, Shift
End Sub

Private Sub dbcintMunicipioC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintUFC_Click(Area As Integer)
    DropDownDataCombo dbcintUFC, Me, Area
End Sub

Private Sub dbcintUFC_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUFC, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUFC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)


    If UCase(strModoOperacao) = gstrPreencherLista Then
        dbcstrLogradouroC.Tag = gstrQueryLogradouro & ";L.strDescricao"
        PreencherListaDeOpcoes Me.ActiveControl
        dbcstrLogradouroC.Tag = ""
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = gstrLocalizar Then
        LeDaTabelaParaObj gstrImobiliario, tdb_Lista, strQuery
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = gstrSalvar Then
        If Val(tdb_Lista.Columns("intContribuinte").Value) > 0 Then
            If Not blnDadosOk Then Exit Sub
            If gblnExclusaoGravacaoOk("A") Then
                AlteraNotificacao (tdb_Lista.Columns("intContribuinte").Value)
                Exit Sub
            End If
        End If
    End If
    
    ToolBarGeral strModoOperacao, gstrImobiliario, False, tdb_Lista, Me, mobjAux, strQuery
                 
End Sub

Private Function strQuery() As String
Dim strSql As String
    
    strSql = "SELECT IM.Pkid,  " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, IM.strCnpjCpf, IM.intContribuinte, CO.strNome "
    strSql = strSql & "FROM " & gstrImobiliario & " IM, " & gstrContribuinte & " CO "
    strSql = strSql & "WHERE IM.intContribuinte = CO.pkid "
    
    If Len(mskstrInscricao) > 0 Then
        strSql = strSql & " AND IM.strInscricao LIKE '" & String(gintLenInscricao - gintRetornaTamanhoMascara(TYP_IMOBILIARIA), "0") & UCase(mskstrInscricao.Text) & "%' "
    End If
    If Len(txtstrNome) > 0 Then
        strSql = strSql & " AND UPPER(CO.strNome) LIKE '" & UCase(txtstrNome) & "%' "
    End If
    
    strSql = strSql & " ORDER BY strInscricao "
    
    strQuery = strSql
    
End Function

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao
End Sub

Private Sub txtstrNome_GotFocus()
    MarcaCampo txtstrNome
End Sub

Private Sub txtstrNome_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNome
End Sub

Private Sub VerificaMascaraInscricao()
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
    
    mskstrInscricao.Mask = strMascara
End Sub

Private Function AlteraNotificacao(lngPkid As Long) As Boolean
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "UPDATE " & gstrContribuinte & " Set "
    strSql = strSql & "intTipoLogradouro = " & gstrENulo(dbcintTipoLogradouro.BoundText, , True) & ", "
    strSql = strSql & "intTituloLogradouro = " & gstrENulo(dbcintTituloLogradouro.BoundText, , True) & ", "
    strSql = strSql & "intCodigoLogradouro = " & gstrENulo(txtintCodigoLogradouro, , True) & ", "
    strSql = strSql & "strlogradouroc = '" & gstrENulo(dbcstrLogradouroC) & "', "
    strSql = strSql & "intNumeroC = " & gstrENulo(txtintNumeroC, , True) & ", "
    strSql = strSql & "strComplementoC = '" & gstrENulo(txtstrComplementoC) & "', "
    strSql = strSql & "strBairroC = '" & gstrENulo(txtstrBairroC) & "', "
    strSql = strSql & "intMunicipioC = " & gstrENulo(dbcintMunicipioC.BoundText, , True) & ", "
    strSql = strSql & "intUFC = " & gstrENulo(dbcintUFC.BoundText, , True) & ", "
    strSql = strSql & "intcepc = " & gstrENulo(Replace(txtintCepC, "-", ""), , True) & ", "
    strSql = strSql & "strDistritoC = '" & gstrENulo(txtstrDistritoC) & "' "
    strSql = strSql & "Where "
    strSql = strSql & "Pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    If gobjBanco.Execute(strSql) Then
        gobjBanco.ExecutaCommitTrans
        MantemForm gstrNovo
    Else
        ExibeMensagem "Não foi possível concluir alteração do endereço de notificação."
        gobjBanco.ExecutaRollbackTrans
        mskstrInscricao.SetFocus
    End If
    
End Function

