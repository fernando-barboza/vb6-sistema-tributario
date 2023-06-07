VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadContasBancarias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contas Bancárias"
   ClientHeight    =   5970
   ClientLeft      =   2865
   ClientTop       =   2715
   ClientWidth     =   7755
   HelpContextID   =   107
   Icon            =   "CadContasBancarias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7755
   Begin VB.CheckBox chkblnDebitoAutomatico 
      Caption         =   "Debito Automatico"
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5910
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   2895
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   5106
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Contas Bancárias"
      TabPicture(0)   =   "CadContasBancarias.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintAgencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintBanco"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrConta"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblstrDigitoVerificador"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintNumConta"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbcintTipoContaBancaria"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Agencia"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmd_Banco"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtstrConta"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtstrDigitoVerificador"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dbcintBanco"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dbcintAgencia"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtstrdescricao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtintnumeroconta"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkblnContaPublica"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmd_TipoConta"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Impressão/Recebimento"
      TabPicture(1)   =   "CadContasBancarias.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Impressao"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra_Impressao 
         Height          =   1485
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   7245
         Begin VB.TextBox txtstrContaRetorno 
            Height          =   285
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   19
            Top             =   390
            Width           =   2835
         End
         Begin MSDataListLib.DataCombo dbcintTipoCodigoBarra 
            Height          =   315
            Left            =   1920
            TabIndex        =   20
            Top             =   870
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo do código de Barra"
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   960
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Identificação do retorno"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   480
            Width           =   1680
         End
      End
      Begin VB.CommandButton cmd_TipoConta 
         Height          =   315
         Left            =   5595
         Picture         =   "CadContasBancarias.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Conta"
         Top             =   945
         Width           =   360
      End
      Begin VB.CheckBox chkblnContaPublica 
         Caption         =   "Conta Publica"
         Height          =   255
         Left            =   4980
         TabIndex        =   18
         Top             =   2460
         Width           =   1335
      End
      Begin VB.TextBox txtintnumeroconta 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox txtstrdescricao 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1332
         Width           =   5895
      End
      Begin MSDataListLib.DataCombo dbcintAgencia 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   2064
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintBanco 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   1683
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtstrDigitoVerificador 
         Height          =   285
         Left            =   3780
         MaxLength       =   6
         TabIndex        =   17
         Top             =   2445
         Width           =   525
      End
      Begin VB.TextBox txtstrConta 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2445
         Width           =   1815
      End
      Begin VB.CommandButton cmd_Banco 
         Height          =   315
         Left            =   6990
         Picture         =   "CadContasBancarias.frx":1198
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Banco"
         Top             =   1683
         Width           =   360
      End
      Begin VB.CommandButton cmd_Agencia 
         Height          =   315
         Left            =   6990
         Picture         =   "CadContasBancarias.frx":12B6
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "591"
         ToolTipText     =   "Ativa Cadastro de Agências"
         Top             =   2064
         Width           =   360
      End
      Begin MSDataListLib.DataCombo dbcintTipoContaBancaria 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   945
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo da Conta"
         Height          =   195
         Left            =   255
         TabIndex        =   3
         Top             =   1005
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descricao"
         Height          =   195
         Left            =   540
         TabIndex        =   6
         Top             =   1377
         Width           =   720
      End
      Begin VB.Label lblintNumConta 
         Alignment       =   1  'Right Justify
         Caption         =   "Conta"
         Height          =   255
         Left            =   585
         TabIndex        =   1
         Top             =   630
         Width           =   675
      End
      Begin VB.Label lblstrDigitoVerificador 
         AutoSize        =   -1  'True
         Caption         =   "DV"
         Height          =   195
         Left            =   3450
         TabIndex        =   16
         Top             =   2490
         Width           =   225
      End
      Begin VB.Label lblstrConta 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   2490
         Width           =   1065
      End
      Begin VB.Label lblintBanco 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   795
         TabIndex        =   8
         Top             =   1743
         Width           =   465
      End
      Begin VB.Label lblintAgencia 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   675
         TabIndex        =   11
         Top             =   2124
         Width           =   585
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   2925
      Left            =   90
      TabIndex        =   26
      Top             =   2970
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5159
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
      Columns(1).Caption=   "Conta"
      Columns(1).DataField=   "intnumeroconta"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Banco"
      Columns(2).DataField=   "strSigla"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Descrição"
      Columns(3).DataField=   "strdescricao"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Agência"
      Columns(4).DataField=   "strAgencia"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Conta Corrente"
      Columns(5).DataField=   "strConta"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "DV"
      Columns(6).DataField=   "strDigitoVerificador"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Públ."
      Columns(7).DataField=   "contapublica"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
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
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1482"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1402"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=4048"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3969"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=3545"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=3466"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2646"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2566"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=3228"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=3149"
      Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(33)=   "Column(6).Width=714"
      Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=635"
      Splits(0)._ColumnProps(36)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=1058"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=979"
      Splits(0)._ColumnProps(41)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(69)  =   "Named:id=33:Normal"
      _StyleDefs(70)  =   ":id=33,.parent=0"
      _StyleDefs(71)  =   "Named:id=34:Heading"
      _StyleDefs(72)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   ":id=34,.wraptext=-1"
      _StyleDefs(74)  =   "Named:id=35:Footing"
      _StyleDefs(75)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(76)  =   "Named:id=36:Selected"
      _StyleDefs(77)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(78)  =   "Named:id=37:Caption"
      _StyleDefs(79)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(80)  =   "Named:id=38:HighlightRow"
      _StyleDefs(81)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(82)  =   "Named:id=39:EvenRow"
      _StyleDefs(83)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(84)  =   "Named:id=40:OddRow"
      _StyleDefs(85)  =   ":id=40,.parent=33"
      _StyleDefs(86)  =   "Named:id=41:RecordSelector"
      _StyleDefs(87)  =   ":id=41,.parent=34"
      _StyleDefs(88)  =   "Named:id=42:FilterBar"
      _StyleDefs(89)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadContasBancarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando    As Boolean
Dim mobjAux          As Object
Dim mblnGuardaUltimo As Boolean
Dim mlngUltimo       As Long
Dim mblnSelecionou   As Boolean
Dim mblnPrimeiraVez  As Boolean
Dim FlagOperacao     As String
Dim bytOrdenacao     As Byte
Dim blnOrdenacaoAsc  As Boolean
Dim blnOrcamentario  As Boolean
Dim blnClick         As Boolean

Private Sub chkblnContaPublica_GotFocus()
    tab_3DPasta.Tab = 0
End Sub

Private Sub chkblnContaPublica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkblnContaPublica
End Sub

Private Sub chkblnDebitoAutomatico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkblnDebitoAutomatico
End Sub

Private Sub cmd_TipoConta_Click()
    CarregaForm frmCadTipoContaBancaria, dbcintTipoContaBancaria
    Set gobjGeral = dbcintTipoContaBancaria
End Sub

Private Sub dbcintAgencia_Click(Area As Integer)
    DropDownDataCombo dbcintAgencia, Me, Area
    mblnPrimeiraVez = False
End Sub

Private Sub dbcintAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintAgencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintAgencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintAgencia
End Sub

Private Sub dbcintBanco_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintBanco, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBanco_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintBanco
End Sub

Private Sub dbcintTipoCodigoBarra_GotFocus()
    MarcaCampo dbcintTipoCodigoBarra
End Sub

Private Sub dbcintTipoCodigoBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTipoCodigoBarra, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoCodigoBarra_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTipoCodigoBarra
End Sub

Private Sub dbcintTipoContaBancaria_Click(Area As Integer)
    DropDownDataCombo dbcintTipoContaBancaria, Me, Area
End Sub

Private Sub dbcintTipoContaBancaria_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTipoContaBancaria, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoContaBancaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTipoContaBancaria
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 628
    VirificaGradeListView Me
    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()

    bytOrdenacao = 1: blnOrdenacaoAsc = True

    mblnAlterando = False
    dbcintBanco.Tag = "Select Pkid,strDescricao From " & gstrBanco & " Order By strdescricao;strdescricao"
    dbcintAgencia.Tag = "Select Pkid,strDescricao From " & gstrAgencia & " Order By strdescricao;strdescricao"
    dbcintTipoContaBancaria.Tag = "Select Pkid,strDescricao From " & gstrTipoContaBancaria & " Order By strdescricao;strdescricao"
    dbcintTipoCodigoBarra.Tag = "Select Pkid, strDescricao From tblTipoCodigoBarra Order By strDescricao" & ";strDescricao"
    
    VerificaListaAutomatica gstrBanco, dbcintBanco
    VerificaObjParaAplicar mobjAux
    txtintnumeroconta = strProximaConta
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Err_Form_QueryUnload
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
    
Err_Form_QueryUnload:
End Sub

Private Function strQueryCampos() As String

'******************************************************************************************
' Data: 07/03/2003
' Alteração: - Retirada a palavra chave "AS" das cláusulas FROM, pois o Oracle não permite
'            a utilização desta palavra chave nesta cláusula.
' Responsável: Everton Bianchini
'******************************************************************************************
Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select CB.PKId, BO.strdescricao as strSigla, AG.strdescricao as strAgencia, CB.strConta, "
    strSQL = strSQL & "CB.strDigitoVerificador, CB.strdescricao, CB.intnumeroconta ,"
    strSQL = strSQL & gstrCASEWHEN("CB.blncontapublica", "1,'Sim',0,'Não'") & " contapublica "
    strSQL = strSQL & "From " & gstrContaBancaria & " CB, "
    strSQL = strSQL & gstrBanco & " BO, "
    strSQL = strSQL & gstrAgencia & " AG "
    strSQL = strSQL & "Where "
    
    strSQL = strSQL & "CB.intAgencia = AG.PKId "
    strSQL = strSQL & "AND CB.intBanco = BO.PKId "
    
    If dbcintBanco.MatchedWithList Then
        strSQL = strSQL & "AND CB.intBanco =" & dbcintBanco.BoundText & " "
    End If

    If dbcintAgencia.MatchedWithList Then
        strSQL = strSQL & "AND CB.intAgencia =" & dbcintAgencia.BoundText & " "
    End If
    
    If Trim(dbcintTipoContaBancaria.Text) <> "" Then
        strSQL = strSQL & " AND CB.intTipoContaBancaria = " & gstrItemData(dbcintTipoContaBancaria)
    End If
    
    If chkblnContaPublica.Value = 1 Then
        strSQL = strSQL & " AND CB.blncontapublica = 1 "
    End If
    
    
    strSQL = strSQL & " Order by CB.intnumeroconta"
    
    strQueryCampos = strSQL

End Function

Private Sub dbcintBanco_Click(Area As Integer)
    DropDownDataCombo dbcintBanco, Me, Area
End Sub

Private Function strQuery() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "Select CB.PKId, BO.strdescricao as strSigla, AG.strdescricao as strAgencia, CB.strConta, "
    strSQL = strSQL & "CB.strDigitoVerificador, CB.strdescricao, CB.intnumeroconta "
    strSQL = strSQL & "From " & gstrContaBancaria & " CB, "
    strSQL = strSQL & gstrBanco & " BO, "
    strSQL = strSQL & gstrAgencia & " AG "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "CB.intAgencia = AG.PKId "
    strSQL = strSQL & "AND CB.intBanco = BO.PKId "
    
    Select Case bytOrdenacao
        Case Is = 1
            strSQL = strSQL & " ORDER BY CB.intnumeroconta" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSQL = strSQL & " ORDER BY BO.strSigla" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strSQL = strSQL & " ORDER BY strAgencia" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 4
            strSQL = strSQL & " ORDER BY CB.strConta" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 5
            strSQL = strSQL & " ORDER BY CB.strDigitoVerificador" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSQL
    
End Function

Private Function strQueryAgenciaContribuinte() As String
Dim strSQL       As String
Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT intBanco "
    strSQL = strSQL & "FROM " & gstrContaBancaria & " "
    strSQL = strSQL & "WHERE PKId = " & txtPKId
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If Not .BOF And Not .EOF Then
                strSQL = ""
                strSQL = strSQL & "SELECT AG.PKId, AG.strDescricao FROM "
                strSQL = strSQL & gstrBanco & " BC, "
                strSQL = strSQL & gstrAgencia & " AG "
                strSQL = strSQL & "WHERE BC.PKId = AG.intBanco "
                strSQL = strSQL & "AND BC.PKId = " & !intBanco & " "
                strSQL = strSQL & "ORDER BY AG.strDescricao"
                strQueryAgenciaContribuinte = strSQL
                Exit Function
            End If
        End With
    End If
    
    strQueryAgenciaContribuinte = ""
    
End Function

Private Sub cmd_Agencia_Click()
    CarregaForm frmCadAgenciaBanco, dbcintAgencia
End Sub

Private Sub cmd_Banco_Click()
    CarregaForm frmCadBanco, dbcintBanco
End Sub

Sub LimpaCampos(blnFlag As Boolean)
    If blnFlag Then
        dbcintBanco.BoundText = ""
        dbcintAgencia.BoundText = ""
    End If
    txtstrConta = ""
    txtstrDigitoVerificador = ""
    chkblnContaPublica.Value = 0
    chkblnDebitoAutomatico = 0
    mblnAlterando = False
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tab_3DPasta
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    blnClick = True
    'tdb_Lista_RowColChange 0, 0
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    gblnFilraCampos tdb_Lista
    blnClick = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnPrimeiraVez = True
    blnClick = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Lista
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If Not .EOF And Not .BOF And blnClick Then
            If mblnPrimeiraVez Then
                mblnPrimeiraVez = False
                mblnAlterando = True
                txtPKId = .Columns("PKID").Value
                LeDaTabelaParaObj gstrContaBancaria, Me
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar, gstrDeletar
                mblnSelecionou = True
            End If
        End If
    End With
End Sub

Private Sub txtintnumeroconta_GotFocus()
    txtintnumeroconta = strProximaConta
    MarcaCampo txtintnumeroconta
End Sub

Private Sub txtintnumeroconta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintnumeroconta
End Sub

Private Sub txtstrConta_GotFocus()
    MarcaCampo txtstrConta
End Sub

Private Sub txtstrConta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrConta
End Sub

Private Sub txtstrContaRetorno_GotFocus()
    tab_3DPasta.Tab = 1
    MarcaCampo txtstrContaRetorno
End Sub

Private Sub txtstrContaRetorno_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrContaRetorno
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrdescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrdescricao
End Sub

Private Sub txtstrDigitoVerificador_GotFocus()
    MarcaCampo txtstrDigitoVerificador
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSQL As String
    
    FlagOperacao = ""
    mblnGuardaUltimo = True
    Select Case UCase(strModoOperacao)
        Case UCase(gstrSalvar)
            If blnDadosOk Then
                If mblnAlterando <> True Then
                    If gblnExisteCodigo(2, gstrContaBancaria, "intnumeroconta", Trim(txtintnumeroconta), "intnumeroconta", txtintnumeroconta) Then
                        ExibeMensagem "Este número da Conta já existe "
                        txtintnumeroconta.SetFocus
                        Exit Sub
                    End If
                End If
            Else
                Exit Sub
            End If
            
        Case UCase(gstrNovo)
            ToolBarGeral strModoOperacao, gstrContaBancaria, mblnAlterando, tdb_Lista, Me, mobjAux
            txtintnumeroconta = strProximaConta
            tab_3DPasta.Tab = 0
            txtintnumeroconta.SetFocus
           
            
        Case gstrPreencherLista
            PreencherListaDeOpcoes Me.ActiveControl
            Exit Sub
            
        Case gstrLocalizar
            ToolBarGeral strModoOperacao, gstrContaBancaria, mblnAlterando, tdb_Lista, Me, mobjAux, strQueryCampos
            Exit Sub
    End Select
    
    If UCase(strModoOperacao) = UCase(gstrAplicar) Then
        ToolBarGeral strModoOperacao, gstrContaBancaria, mblnAlterando, tdb_Lista, Me, mobjAux, strQueryCampos, strQueryAplicar
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
        mblnPrimeiraVez = False
        If UCase(strModoOperacao) = UCase(gstrSalvar) Then
            If Not blnDadosOk Then Exit Sub
        End If
        If ToolBarGeral(strModoOperacao, gstrContaBancaria, mblnAlterando, tdb_Lista, Me, mobjAux) Then
            tab_3DPasta.Tab = 0
        End If
        If gblnCancelarInclusao = False Then
           LeDaTabelaParaObj gstrContaBancaria, tdb_Lista, strQueryCampos
           txtintnumeroconta = strProximaConta
        End If
    ElseIf UCase(strModoOperacao) <> UCase(gstrNovo) Then
        ToolBarGeral strModoOperacao, gstrContaBancaria, mblnAlterando, tdb_Lista, Me, mobjAux, strQueryCampos, , rptCadContaBancaria, strQueryRelatorio
    End If
    
    FlagOperacao = UCase(strModoOperacao)
    mblnGuardaUltimo = False
    
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
End Sub

Private Sub txtstrDigitoVerificador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDigitoVerificador
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If Trim(txtintnumeroconta) = "" Then
        ExibeMensagem "O número da conta tem que ser digitado."
        txtintnumeroconta.SetFocus
        Exit Function
    End If
    
    If dbcintTipoContaBancaria.MatchedWithList = False Or Trim(dbcintTipoContaBancaria.Text) = "" Then
        ExibeMensagem "É necessário escolher algum tipo de conta bancária da lista de opções."
        If dbcintTipoContaBancaria.Enabled Then dbcintTipoContaBancaria.SetFocus
        Exit Function
    End If
    
    If Trim(txtstrdescricao) = "" Then
        ExibeMensagem "A descrição é obrigatória."
        If txtstrdescricao.Enabled Then txtstrdescricao.SetFocus
        Exit Function
    End If
    
    If dbcintBanco.MatchedWithList = False Or Trim(dbcintBanco.Text) = "" Then
        ExibeMensagem "É necessário escolher algum banco da lista de opções."
        If dbcintBanco.Enabled Then dbcintBanco.SetFocus
        Exit Function
    End If
    
    If dbcintAgencia.MatchedWithList = False Or Trim(dbcintAgencia.Text) = "" Then
        ExibeMensagem "É necessário escolher alguma agência da lista de opções."
        If dbcintAgencia.Enabled Then dbcintAgencia.SetFocus
        Exit Function
    End If
    
    If Trim(txtstrConta) = "" Then
        ExibeMensagem "A conta tem que ser digitada."
        txtstrConta.SetFocus
        Exit Function
    End If
    
    If Trim(txtstrDigitoVerificador.Text) = "" Then
        ExibeMensagem "O dígito verificador deve ser digitado."
        txtstrDigitoVerificador.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True

End Function

Private Function strProximaConta() As String
Dim strSQL       As String
Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "Max(" & gstrISNULL("intnumeroconta", "0") & ")+1 as intnumeroconta FROM "
    strSQL = strSQL & gstrContaBancaria
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount >= 1 Then
            If Not IsNull(adoResultado!intNumeroConta) Then
                strProximaConta = adoResultado!intNumeroConta
            Else
                strProximaConta = 1
            End If
        End If
    End If
End Function

Private Function strQueryRelatorio() As String
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "CB.PKId, "
    strSQL = strSQL & "CB.intnumeroconta, "
    strSQL = strSQL & "BO.strdescricao as strSigla,"
    strSQL = strSQL & "AG.strdescricao as strAgencia, "
    If bytDBType = Oracle Then
        strSQL = strSQL & "Trim(CB.strConta) as strConta, "
        strSQL = strSQL & "Trim(CB.strDigitoVerificador) as strDigitoVerificador, "
    Else
        strSQL = strSQL & "Ltrim(Rtrim(CB.strConta)) as strConta, "
        strSQL = strSQL & "Ltrim(Rtrim(CB.strDigitoVerificador)) as strDigitoVerificador, "
    End If
    strSQL = strSQL & "CB.strdescricao "
    strSQL = strSQL & "From " & gstrContaBancaria & " CB, "
    strSQL = strSQL & gstrBanco & " BO, "
    strSQL = strSQL & gstrAgencia & " AG "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "CB.intAgencia = AG.PKId "
    strSQL = strSQL & "AND CB.intBanco = BO.PKId "
    strSQL = strSQL & " Order by CB.intnumeroconta"
    
    strQueryRelatorio = strSQL
    
End Function

Private Function strQueryAplicar() As String
Dim strSQL As String

    strSQL = "SELECT Pkid, "
    strSQL = strSQL & "LTRIM(RTRIM(" & gstrCONVERT(CDT_VARCHAR, "strConta") & "))" & strCONCAT & "'-'" & strCONCAT & " strDigitoVerificador ContaCorrente"
    strSQL = strSQL & " FROM " & gstrContaBancaria
    strSQL = strSQL & " Where Pkid = " & Trim(txtPKId) & " "
    strSQL = strSQL & " ORDER BY intNumeroConta, strDigitoVerificador"
    
    strQueryAplicar = strSQL

End Function
