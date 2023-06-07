VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadPlanoConta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plano de Contas"
   ClientHeight    =   8520
   ClientLeft      =   2100
   ClientTop       =   2340
   ClientWidth     =   9600
   HelpContextID   =   168
   Icon            =   "CadPlanoConta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1650
      Left            =   180
      TabIndex        =   109
      Top             =   6750
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   2910
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
      Columns(1).Caption=   "Conta"
      Columns(1).DataField=   "strContaContabil"
      Columns(1).NumberFormat=   "FormatText Event"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Descrição"
      Columns(2).DataField=   "strDescricao"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Natureza"
      Columns(3).DataField=   "strNaturezaDaConta"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2566"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=11615"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=11536"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1614"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1535"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DefColWidth     =   0
      EditDropDown    =   0   'False
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTips        =   1
      CellTipsWidth   =   0
      MultiSelect     =   0
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=97,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(53)  =   "Named:id=33:Normal"
      _StyleDefs(54)  =   ":id=33,.parent=0"
      _StyleDefs(55)  =   "Named:id=34:Heading"
      _StyleDefs(56)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   ":id=34,.wraptext=-1"
      _StyleDefs(58)  =   "Named:id=35:Footing"
      _StyleDefs(59)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   "Named:id=36:Selected"
      _StyleDefs(61)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(62)  =   "Named:id=37:Caption"
      _StyleDefs(63)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(64)  =   "Named:id=38:HighlightRow"
      _StyleDefs(65)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   "Named:id=39:EvenRow"
      _StyleDefs(67)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(68)  =   "Named:id=40:OddRow"
      _StyleDefs(69)  =   ":id=40,.parent=33"
      _StyleDefs(70)  =   "Named:id=41:RecordSelector"
      _StyleDefs(71)  =   ":id=41,.parent=34"
      _StyleDefs(72)  =   "Named:id=42:FilterBar"
      _StyleDefs(73)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame frm_DadosConta 
      Height          =   555
      Left            =   180
      TabIndex        =   72
      Top             =   630
      Width           =   9285
      Begin VB.TextBox txt_strConta 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   630
         TabIndex        =   75
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txt_strDescricao 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   2970
         TabIndex        =   74
         Top             =   180
         Width           =   3645
      End
      Begin VB.TextBox txt_dblSaldoInicial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   7740
         TabIndex        =   73
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label lbl_strConta 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         Height          =   195
         Left            =   120
         TabIndex        =   79
         Top             =   210
         Width           =   420
      End
      Begin VB.Label lbl_strDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descricao"
         Height          =   195
         Left            =   2130
         TabIndex        =   78
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lbl_dblValorInicial 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Inicial"
         Height          =   195
         Left            =   6780
         TabIndex        =   77
         Top             =   210
         Width           =   855
      End
      Begin VB.Label lbl_Natureza 
         AutoSize        =   -1  'True
         Caption         =   "CR"
         Height          =   195
         Left            =   8985
         TabIndex        =   76
         Top             =   210
         Width           =   225
      End
   End
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5520
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   8385
      Left            =   75
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   90
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   14790
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Plano "
      TabPicture(0)   =   "CadPlanoConta.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrContaContabil"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintContaReduzida"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbldblSaldoDaConta"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_SaldoAtual"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblstrDescricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtstrContaContabil"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_blnNaturezaDaConta"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtdblSaldoDaConta"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtintContaReduzida"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_Tipo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_SaldoAtual"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtstrDescricao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fra_Itegrante"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "frm_Deducao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fra_Aplicavel"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkblnContaAtiva"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "fra_contaBancaria"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "fra_MovimentaSistemas"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Resumo Movimentado"
      TabPicture(1)   =   "CadPlanoConta.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvw_Resumo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cruzamento de Sistemas"
      TabPicture(2)   =   "CadPlanoConta.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tab_3DPastaMovimento"
      Tab(2).ControlCount=   1
      Begin VB.Frame fra_MovimentaSistemas 
         Caption         =   "Movimentar Sistemas"
         Height          =   675
         Left            =   6615
         TabIndex        =   107
         Top             =   4260
         Width           =   2700
         Begin VB.CheckBox chkbytMovimentaSistema 
            Caption         =   "Não movimentar no Sistema"
            Enabled         =   0   'False
            Height          =   195
            Left            =   225
            TabIndex        =   45
            Top             =   300
            Width           =   2280
         End
      End
      Begin VB.Frame fra_contaBancaria 
         Caption         =   "Conta Bancária"
         Height          =   1785
         Left            =   90
         TabIndex        =   71
         Top             =   4260
         Width           =   6435
         Begin VB.TextBox txt_intConvenio 
            Height          =   315
            Left            =   510
            MaxLength       =   2
            TabIndex        =   51
            Top             =   1380
            Width           =   315
         End
         Begin VB.TextBox txt_intModalidade 
            Height          =   315
            Left            =   90
            MaxLength       =   3
            TabIndex        =   50
            Top             =   1380
            Width           =   405
         End
         Begin VB.TextBox txt_intFonteRecurso 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            TabIndex        =   47
            Top             =   810
            Width           =   1155
         End
         Begin MSDataListLib.DataCombo dbcintcontabancaria 
            Height          =   315
            Left            =   1290
            TabIndex        =   44
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intNumeroConta 
            Height          =   315
            Left            =   90
            TabIndex        =   43
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intFonteRecurso 
            Height          =   315
            Left            =   1290
            TabIndex        =   48
            Top             =   810
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintModalidade 
            Height          =   315
            Left            =   870
            TabIndex        =   52
            Top             =   1380
            Width           =   5475
            _ExtentX        =   9657
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblintModalidade 
            AutoSize        =   -1  'True
            Caption         =   "Modalidade de Aplicação"
            Height          =   195
            Left            =   150
            TabIndex        =   49
            Top             =   1140
            Width           =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fontes de Recurso "
            Height          =   195
            Left            =   150
            TabIndex        =   46
            Top             =   570
            Width           =   1395
         End
      End
      Begin VB.CheckBox chkblnContaAtiva 
         Caption         =   "Conta Ativa"
         Height          =   195
         Left            =   7995
         TabIndex        =   5
         Top             =   930
         Width           =   1215
      End
      Begin VB.Frame fra_Aplicavel 
         Caption         =   " Aplicar "
         Height          =   2115
         Left            =   90
         TabIndex        =   69
         Top             =   2130
         Width           =   4635
         Begin VB.CheckBox chkbytProjecaoAtuarial 
            Caption         =   "Projeção Atuarial"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1830
            Width           =   1545
         End
         Begin VB.CheckBox chkbytRegimePrevidenciario 
            Caption         =   "Regime próprio de prev.social"
            Height          =   195
            Left            =   1980
            TabIndex        =   26
            Top             =   1602
            Width           =   2445
         End
         Begin VB.CheckBox chkbytDisponibilidadeDeCaixa 
            Caption         =   "Disponibilidade de caixa"
            Height          =   195
            Left            =   1980
            TabIndex        =   22
            Top             =   1148
            Width           =   2445
         End
         Begin VB.CheckBox chkbytDemaisAtivoFinanceiro 
            Caption         =   "Demais ativos financeiros"
            Height          =   195
            Left            =   1980
            TabIndex        =   24
            Top             =   1375
            Width           =   2445
         End
         Begin VB.CheckBox chkbytDividaConsolidada 
            Caption         =   "Dívida consolidada"
            Height          =   195
            Left            =   1980
            TabIndex        =   18
            Top             =   694
            Width           =   2445
         End
         Begin VB.CheckBox chkbytAplicacaoFinanceira 
            Caption         =   "Aplicações Financeiras"
            Height          =   195
            Left            =   1980
            TabIndex        =   20
            Top             =   921
            Width           =   2445
         End
         Begin VB.CheckBox chkbytPensionista 
            Caption         =   "Pensionista"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   1375
            Width           =   1455
         End
         Begin VB.CheckBox chkbytPessoalInativo 
            Caption         =   "Pessoal inativo"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1148
            Width           =   1455
         End
         Begin VB.CheckBox chkbytAntecipacaoDeReceitaOrcamen 
            Caption         =   "Antecipação de receita"
            Height          =   195
            Left            =   1980
            TabIndex        =   16
            Top             =   467
            Width           =   2445
         End
         Begin VB.CheckBox chkbytPrevidenciaria 
            Caption         =   "Previdenciária"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1602
            Width           =   1455
         End
         Begin VB.CheckBox chkbytEducacao 
            Caption         =   "Educação"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkbytSaude 
            Caption         =   "Saúde"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   467
            Width           =   1455
         End
         Begin VB.CheckBox chkbytFundef 
            Caption         =   "Fundef"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   694
            Width           =   1455
         End
         Begin VB.CheckBox chkbytPessoal 
            Caption         =   "Pessoal ativo"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   921
            Width           =   1455
         End
         Begin VB.CheckBox chkbytServicoDeTerceiro 
            Caption         =   "Serviços de terceiros"
            Height          =   195
            Left            =   1980
            TabIndex        =   14
            Top             =   240
            Width           =   2445
         End
      End
      Begin VB.Frame frm_Deducao 
         Caption         =   " Deduzir "
         Height          =   2115
         Left            =   4800
         TabIndex        =   68
         Top             =   2130
         Width           =   4515
         Begin VB.CheckBox chkbytDeduzProjecaoAtuarial 
            Caption         =   "Projeção Atuarial"
            Height          =   195
            Left            =   90
            TabIndex        =   42
            Top             =   1830
            Width           =   1635
         End
         Begin VB.CheckBox chkbytDeduzRegimePrevidenciario 
            Caption         =   "Regime próprio de prev.social"
            Height          =   195
            Left            =   1920
            TabIndex        =   41
            Top             =   1632
            Width           =   2475
         End
         Begin VB.CheckBox chkbytDeduzDisponibilidadeDeCaixa 
            Caption         =   "Disponibilidade de caixa"
            Height          =   195
            Left            =   1920
            TabIndex        =   37
            Top             =   1168
            Width           =   2475
         End
         Begin VB.CheckBox chkbytDeduzDemaisAtivoFinanceiro 
            Caption         =   "Demais ativos financeiros"
            Height          =   195
            Left            =   1920
            TabIndex        =   39
            Top             =   1400
            Width           =   2475
         End
         Begin VB.CheckBox chkbytDeduzDividaConsolidada 
            Caption         =   "Dívida consolidada"
            Height          =   195
            Left            =   1920
            TabIndex        =   33
            Top             =   704
            Width           =   2475
         End
         Begin VB.CheckBox chkbytDeduzAplicacaoFinanceira 
            Caption         =   "Aplicações Financeiras"
            Height          =   195
            Left            =   1920
            TabIndex        =   35
            Top             =   936
            Width           =   2475
         End
         Begin VB.CheckBox chkbytDeduzPensionista 
            Caption         =   "Pensionista"
            Height          =   225
            Left            =   90
            TabIndex        =   38
            Top             =   1375
            Width           =   1365
         End
         Begin VB.CheckBox chkbytDeduzPessoalInativo 
            Caption         =   "Pessoal inativo"
            Height          =   225
            Left            =   90
            TabIndex        =   36
            Top             =   1148
            Width           =   1365
         End
         Begin VB.CheckBox chkbytDeduzAntecipacaoDeReceitaOr 
            Caption         =   "Antecipação de receita"
            Height          =   195
            Left            =   1920
            TabIndex        =   31
            Top             =   472
            Width           =   2475
         End
         Begin VB.CheckBox chkbytDeduzPrevidenciaria 
            Caption         =   "Previdenciária"
            Height          =   225
            Left            =   90
            TabIndex        =   40
            Top             =   1602
            Width           =   1365
         End
         Begin VB.CheckBox chkbytDeduzPessoal 
            Caption         =   "Pessoal"
            Height          =   225
            Left            =   90
            TabIndex        =   34
            Top             =   921
            Width           =   1365
         End
         Begin VB.CheckBox chkbytDeduzFundef 
            Caption         =   "Fundef"
            Height          =   225
            Left            =   90
            TabIndex        =   32
            Top             =   694
            Width           =   1365
         End
         Begin VB.CheckBox chkbytDeduzSaude 
            Caption         =   "Saúde"
            Height          =   225
            Left            =   90
            TabIndex        =   30
            Top             =   467
            Width           =   1365
         End
         Begin VB.CheckBox chkbytDeduzEducacao 
            Caption         =   "Educação"
            Height          =   225
            Left            =   90
            TabIndex        =   28
            Top             =   240
            Width           =   1365
         End
         Begin VB.CheckBox chkbytDeduzServicoDeTerceiro 
            Caption         =   "Serviços de terceiros"
            Height          =   195
            Left            =   1920
            TabIndex        =   29
            Top             =   240
            Width           =   2475
         End
      End
      Begin VB.Frame fra_Itegrante 
         Caption         =   " Integrante "
         Height          =   525
         Left            =   90
         TabIndex        =   67
         Top             =   6060
         Width           =   9225
         Begin VB.CheckBox chkblnPatrimonial 
            Caption         =   "Patrimonial"
            Height          =   195
            Left            =   8010
            TabIndex        =   57
            Top             =   210
            Width           =   1110
         End
         Begin VB.CheckBox chkblnOrcamentario 
            Caption         =   "Orçamentário"
            Height          =   195
            Left            =   6339
            TabIndex        =   56
            Top             =   210
            Width           =   1320
         End
         Begin VB.CheckBox chkblnFinanceira 
            Caption         =   "Financeira"
            Height          =   195
            Left            =   4941
            TabIndex        =   55
            Top             =   210
            Width           =   1050
         End
         Begin VB.CheckBox chkbytVariacaoPatrimonial 
            Caption         =   "Variação Patrimonial"
            Height          =   195
            Left            =   990
            TabIndex        =   53
            Top             =   210
            Width           =   1770
         End
         Begin VB.CheckBox chkblnIntegraDividaFundada 
            Caption         =   "Dívida Fundada"
            Height          =   195
            Left            =   3108
            TabIndex        =   54
            Top             =   210
            Width           =   1485
         End
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
         Height          =   510
         Left            =   90
         MaxLength       =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   930
         Width           =   7770
      End
      Begin VB.TextBox txt_SaldoAtual 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   7860
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   390
         Width           =   1530
      End
      Begin VB.Frame fra_Tipo 
         Caption         =   " Tipo "
         Height          =   670
         Left            =   75
         TabIndex        =   64
         Top             =   1440
         Width           =   7785
         Begin VB.TextBox txtIntExtraMaua 
            Height          =   285
            Left            =   4830
            MaxLength       =   10
            TabIndex        =   9
            Top             =   330
            Width           =   1095
         End
         Begin VB.CheckBox chkblnAnalitica 
            Caption         =   "Analítica"
            Height          =   195
            Left            =   1590
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   300
            Width           =   930
         End
         Begin VB.CheckBox chkblnInversaoDeSaldo 
            Caption         =   "Inverte Saldo"
            Height          =   195
            Left            =   6360
            TabIndex        =   10
            Top             =   300
            Width           =   1245
         End
         Begin VB.CheckBox chkblnRetificadora 
            Caption         =   "Retificadora"
            Height          =   195
            Left            =   60
            TabIndex        =   6
            Top             =   300
            Width           =   1185
         End
         Begin VB.CheckBox chkblnExtraOrcamentaria 
            Caption         =   "Extra-orçamentária"
            Height          =   195
            Left            =   3075
            TabIndex        =   8
            Top             =   300
            Width           =   1650
         End
         Begin VB.Label lblContaExtra 
            Caption         =   "Conta Extra"
            Height          =   165
            Left            =   4830
            TabIndex        =   108
            Top             =   120
            Width           =   825
         End
      End
      Begin VB.TextBox txtintContaReduzida 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3285
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox txtdblSaldoDaConta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   6
         EndProperty
         Height          =   285
         HideSelection   =   0   'False
         Left            =   5325
         MaxLength       =   18
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   390
         WhatsThisHelpID =   1
         Width           =   1530
      End
      Begin VB.Frame fra_blnNaturezaDaConta 
         Caption         =   " Natureza "
         Height          =   670
         Left            =   7995
         TabIndex        =   58
         Top             =   1440
         Width           =   1395
         Begin VB.OptionButton optblnNaturezaDaConta 
            Caption         =   "Devedora"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   12
            Top             =   430
            Width           =   1080
         End
         Begin VB.OptionButton optblnNaturezaDaConta 
            Caption         =   "Credora"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   11
            Top             =   200
            Width           =   960
         End
      End
      Begin VB.TextBox txtstrContaContabil 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   6
         EndProperty
         Height          =   285
         HideSelection   =   0   'False
         Left            =   705
         MaxLength       =   25
         OLEDropMode     =   1  'Manual
         TabIndex        =   0
         Tag             =   "1"
         Top             =   390
         WhatsThisHelpID =   1
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvw_Resumo 
         Height          =   2885
         Left            =   -74910
         TabIndex        =   70
         Top             =   1440
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   5080
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
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
            Text            =   "Mês"
            Object.Width           =   3440
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Crédito"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Débito"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Tipo Saldo"
            Object.Width           =   3476
         EndProperty
      End
      Begin TabDlg.SSTab tab_3DPastaMovimento 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   80
         Top             =   1230
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   7435
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Movimento de Crédito"
         TabPicture(0)   =   "CadPlanoConta.frx":1096
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblintContaCredito"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lvw_Credito"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "frm_SistemaCredito"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chk_ContaGrupoCredito"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cbo_DescricaoContaCredito"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cbo_intContaCredito"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "fra_bytTipoCredito"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Movimento de Débito"
         TabPicture(1)   =   "CadPlanoConta.frx":10B2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fra_bytTipoDebito"
         Tab(1).Control(1)=   "cbo_intContaDebito"
         Tab(1).Control(2)=   "frm_SistemaDebito"
         Tab(1).Control(3)=   "chk_ContaGrupoDebito"
         Tab(1).Control(4)=   "cbo_DescricaoContaDebito"
         Tab(1).Control(5)=   "lvw_Debito"
         Tab(1).Control(6)=   "lblintContaDebito"
         Tab(1).ControlCount=   7
         Begin VB.Frame fra_bytTipoDebito 
            Caption         =   "Tipo Movimentação"
            Height          =   675
            Left            =   -71340
            TabIndex        =   102
            Top             =   330
            Width           =   3675
            Begin VB.OptionButton opt_bytTipoDebito 
               Caption         =   "Mutação"
               CausesValidation=   0   'False
               Height          =   195
               Index           =   2
               Left            =   510
               TabIndex        =   106
               Top             =   420
               Width           =   1005
            End
            Begin VB.OptionButton opt_bytTipoDebito 
               Caption         =   "Independente"
               CausesValidation=   0   'False
               Height          =   195
               Index           =   1
               Left            =   2010
               TabIndex        =   105
               Top             =   180
               Width           =   1335
            End
            Begin VB.OptionButton opt_bytTipoDebito 
               Caption         =   "Cancelamento"
               CausesValidation=   0   'False
               Height          =   195
               Index           =   3
               Left            =   2010
               TabIndex        =   104
               Top             =   420
               Width           =   1395
            End
            Begin VB.OptionButton opt_bytTipoDebito 
               Caption         =   "Diversos"
               Height          =   195
               Index           =   0
               Left            =   510
               TabIndex        =   103
               Top             =   180
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.Frame fra_bytTipoCredito 
            Caption         =   "Tipo Movimentação"
            Height          =   675
            Left            =   3660
            TabIndex        =   97
            Top             =   330
            Width           =   3675
            Begin VB.OptionButton opt_bytTipoCredito 
               Caption         =   "Mutação"
               CausesValidation=   0   'False
               Height          =   195
               Index           =   2
               Left            =   510
               TabIndex        =   101
               Top             =   420
               Width           =   1005
            End
            Begin VB.OptionButton opt_bytTipoCredito 
               Caption         =   "Independente"
               CausesValidation=   0   'False
               Height          =   195
               Index           =   1
               Left            =   2010
               TabIndex        =   100
               Top             =   180
               Width           =   1335
            End
            Begin VB.OptionButton opt_bytTipoCredito 
               Caption         =   "Cancelamento"
               CausesValidation=   0   'False
               Height          =   195
               Index           =   3
               Left            =   2010
               TabIndex        =   99
               Top             =   420
               Width           =   1395
            End
            Begin VB.OptionButton opt_bytTipoCredito 
               Caption         =   "Diversos"
               Height          =   195
               Index           =   0
               Left            =   510
               TabIndex        =   98
               Top             =   180
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.ComboBox cbo_intContaDebito 
            Height          =   315
            Left            =   -74325
            Sorted          =   -1  'True
            TabIndex        =   89
            Top             =   1080
            Width           =   1575
         End
         Begin VB.ComboBox cbo_intContaCredito 
            Height          =   315
            Left            =   675
            Sorted          =   -1  'True
            TabIndex        =   87
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Frame frm_SistemaDebito 
            Caption         =   "Sistema"
            Height          =   675
            Left            =   -74820
            TabIndex        =   92
            Top             =   330
            Width           =   3285
            Begin VB.OptionButton opt_FinanceiroDebito 
               Caption         =   "Financeiro"
               Height          =   225
               Left            =   270
               TabIndex        =   94
               Top             =   270
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton opt_EconomicoDebito 
               Caption         =   "Econômico"
               Height          =   225
               Left            =   1620
               TabIndex        =   93
               Top             =   270
               Width           =   1095
            End
         End
         Begin VB.CheckBox chk_ContaGrupoDebito 
            Caption         =   "Conta de Grupo"
            Height          =   255
            Left            =   -67440
            TabIndex        =   91
            Top             =   600
            Width           =   1485
         End
         Begin VB.ComboBox cbo_DescricaoContaDebito 
            Height          =   315
            Left            =   -72795
            Sorted          =   -1  'True
            TabIndex        =   90
            Top             =   1080
            Width           =   6855
         End
         Begin VB.ComboBox cbo_DescricaoContaCredito 
            Height          =   315
            Left            =   2205
            Sorted          =   -1  'True
            TabIndex        =   86
            Top             =   1080
            Width           =   6855
         End
         Begin VB.CheckBox chk_ContaGrupoCredito 
            Caption         =   "Conta de Grupo"
            Height          =   255
            Left            =   7560
            TabIndex        =   83
            Top             =   600
            Width           =   1485
         End
         Begin VB.Frame frm_SistemaCredito 
            Caption         =   "Sistema"
            Height          =   675
            Left            =   180
            TabIndex        =   82
            Top             =   330
            Width           =   3285
            Begin VB.OptionButton opt_EconomicoCredito 
               Caption         =   "Econômico"
               Height          =   225
               Left            =   1620
               TabIndex        =   85
               Top             =   270
               Width           =   1095
            End
            Begin VB.OptionButton opt_FinanceiroCredito 
               Caption         =   "Financeiro"
               Height          =   225
               Left            =   270
               TabIndex        =   84
               Top             =   270
               Value           =   -1  'True
               Width           =   1035
            End
         End
         Begin MSComctlLib.ListView lvw_Credito 
            Height          =   2535
            Left            =   90
            TabIndex        =   81
            Top             =   1530
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   4471
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "pkID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Conta"
               Object.Width           =   2866
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Descrição"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Sistema"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Movimentação"
               Object.Width           =   2295
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Grupo"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "intTipoMovimentacao"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Debito 
            Height          =   2535
            Left            =   -74910
            TabIndex        =   95
            Top             =   1530
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   4471
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "pkID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Conta"
               Object.Width           =   2866
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Descrição"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Sistema"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Movimentação"
               Object.Width           =   2295
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Grupo"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "intTipoMovimentacao"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblintContaDebito 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   -74820
            TabIndex        =   96
            Top             =   1140
            Width           =   420
         End
         Begin VB.Label lblintContaCredito 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   180
            TabIndex        =   88
            Top             =   1140
            Width           =   420
         End
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   90
         TabIndex        =   66
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl_SaldoAtual 
         AutoSize        =   -1  'True
         Caption         =   "Saldo atual"
         Height          =   195
         Left            =   6990
         TabIndex        =   65
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lbldblSaldoDaConta 
         AutoSize        =   -1  'True
         Caption         =   "Saldo inicial"
         Height          =   195
         Left            =   4380
         TabIndex        =   63
         Top             =   420
         Width           =   840
      End
      Begin VB.Label lblintContaReduzida 
         AutoSize        =   -1  'True
         Caption         =   "Reduzido"
         Height          =   195
         Left            =   2535
         TabIndex        =   62
         Top             =   420
         Width           =   675
      End
      Begin VB.Label lblstrContaContabil 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   90
         TabIndex        =   61
         Top             =   420
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmCadPlanoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando           As Boolean
    Dim mblnClickOk             As Boolean
    Dim mobjAux                 As Object
    Dim mstrQueryAplicar        As String
    Dim mblnPrimeiraVez         As Boolean
    
    Dim strContaAtual           As String
    Dim strDescricaoAtual       As String
    Dim strContaExtraAtual      As String
    Dim mblnExistemMovimentos   As Boolean
    
Private Function strQueryAplicar() As String

    Dim strSQL  As String
    If Trim(mstrQueryAplicar) = "" Then
        strSQL = ""
        strSQL = strSQL & "SELECT PC.PKId, PC.strDescricao "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrPlanoConta & " PC "
        strSQL = strSQL & "WHERE blnFinanceira = 1 "
        strSQL = strSQL & "ORDER BY PC.strDescricao"
    Else
        strQueryAplicar = Trim(mstrQueryAplicar)
    End If
End Function

Private Sub VerificaSeEAnalitica()
    Dim intInd  As Integer
    Dim vntAux  As Variant
    vntAux = Split(txtstrContaContabil, ".")
    For intInd = UBound(vntAux) To 0 Step -1
        If vntAux(intInd) <> 0 Then
            chkblnAnalitica = 1
            Exit For
        Else
            chkblnAnalitica = 0
            Exit For
        End If
    Next
End Sub

Private Function strQuery() As String
    Dim strSQL              As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao, "
    strSQL = strSQL & gstrCASEWHEN("PC.blnnaturezadaconta", "1,'Devedora',0,'Credora'")
    strSQL = strSQL & " AS strNaturezaDaConta "
    strSQL = strSQL & "FROM " & gstrPlanoConta & " PC "
    
    If dbc_intFonteRecurso.MatchedWithList Then
        strSQL = strSQL & "," & gstrPlanoContaSaldo & " PS, "
        strSQL = strSQL & gstrFonteRecurso & " FR "
        strSQL = strSQL & "Where "
        strSQL = strSQL & "PC.pkid = PS.intPlanoConta and "
        strSQL = strSQL & "PS.INTFONTERECURSO = FR.Pkid and "
        strSQL = strSQL & "FR.Pkid = " & dbc_intFonteRecurso.BoundText
    End If
    
    strSQL = strSQL & "ORDER BY strContaContabil"
    strQuery = strSQL
End Function

Private Function strQueryModalidade() As String

    Dim strSQL  As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT Pkid, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrModalidade

    strQueryModalidade = strSQL
    
End Function

Private Function blnPreencheControles(ByVal strTabela As String, ByVal strCampo1 As String, ByVal strCampo2 As String, ByVal strCondicao As String, ByRef objControleDestino1 As Object, _
                                      Optional ByVal strCampo3 As String, Optional ByRef objControleDestino2 As Object) As Boolean

    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    blnPreencheControles = False
    
    strSQL = "Select " & strCampo1 & ", " & strCampo2
    
    If Len(Trim$(strCampo3)) > 0 Then
        strSQL = strSQL & ", " & strCampo3
    End If
    
    strSQL = strSQL & " FROM " & strTabela
     
    If Len(Trim$(strCondicao)) > 0 Then
        strSQL = strSQL & " Where " & strCondicao
    End If
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            If TypeOf objControleDestino1 Is TextBox Then
                objControleDestino1.Tag = adoResultado.Fields(0).Value
                objControleDestino1.Text = adoResultado.Fields(1).Value
            ElseIf TypeOf objControleDestino1 Is DataCombo Then
                PreencherListaDeOpcoes objControleDestino1, adoResultado.Fields(0).Value
            End If
            
            If Not objControleDestino2 Is Nothing Then
                If TypeOf objControleDestino2 Is TextBox Then
                    objControleDestino2.Tag = adoResultado.Fields(0).Value
                    objControleDestino2.Text = adoResultado.Fields(2).Value
                ElseIf TypeOf objControleDestino2 Is DataCombo Then
                    PreencherListaDeOpcoes objControleDestino2, adoResultado.Fields(2).Value
                End If
            End If
            
            adoResultado.Close: blnPreencheControles = True
        Else
            If TypeOf objControleDestino1 Is TextBox Then
                objControleDestino1.Tag = Space$(0)
                objControleDestino1.Text = Space$(0)
            ElseIf TypeOf objControleDestino1 Is DataCombo Then
                objControleDestino1.Text = Space$(0)
                Set objControleDestino1.RowSource = Nothing
            End If
            
            If Not objControleDestino2 Is Nothing Then
                If TypeOf objControleDestino2 Is TextBox Then
                    objControleDestino2.Tag = Space$(0)
                    objControleDestino2.Text = Space$(0)
                ElseIf TypeOf objControleDestino2 Is DataCombo Then
                    objControleDestino2.Text = Space$(0)
                    Set objControleDestino2.RowSource = Nothing
                End If
            End If
            
        End If
    End If
    
    Set adoResultado = Nothing
    
End Function

Private Sub chk_ContaGrupoCredito_Click()
    PreencheComboCredito
End Sub

Private Sub chkblnExtraOrcamentaria_Click()
    If chkblnExtraOrcamentaria.Value = 1 Then
        TrocaCorObjeto txtIntExtraMaua, False
    Else
        txtIntExtraMaua.Text = ""
        TrocaCorObjeto txtIntExtraMaua, True
    End If
    
End Sub

Private Sub chkblnExtraOrcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkblnExtraOrcamentaria
End Sub

Private Sub chkblnFinanceira_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkblnIntegraDividaFundada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkblnInversaoDeSaldo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkblnPatrimonial_Click()
    If chkblnPatrimonial.Value = 0 Then
        chkbytMovimentaSistema.Enabled = False
        chkbytMovimentaSistema.Value = 0
        If tab_3dPasta.Tab = 2 Then tab_3dPasta.Tab = 0
        tab_3dPasta.TabEnabled(2) = False
    Else
        chkbytMovimentaSistema.Enabled = True
        tab_3dPasta.TabEnabled(2) = True
    End If
End Sub

Private Sub chkblnRetificadora_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytAntecipacaoDeReceitaOrcamen_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytAplicacaoFinanceira_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzAntecipacaoDeReceitaOr_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzAplicacaoFinanceira_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzDemaisAtivoFinanceiro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzDisponibilidadeDeCaixa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzDividaConsolidada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzEducacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzFundef_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzPensionista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzPessoal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzPessoalInativo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzPrevidenciaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzProjecaoAtuarial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzRegimePrevidenciario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzSaude_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzServicoDeTerceiro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDemaisAtivoFinanceiro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDisponibilidadeDeCaixa_Click()
    If chkbytDisponibilidadeDeCaixa.Value = vbChecked Then
        TrocaCorObjeto dbcintcontabancaria, False
        TrocaCorObjeto dbc_intNumeroConta, False
        TrocaCorObjeto dbc_intFonteRecurso, False
        TrocaCorObjeto txt_intFonteRecurso, False
    Else
        TrocaCorObjeto dbcintcontabancaria, True
        TrocaCorObjeto dbc_intNumeroConta, True
        TrocaCorObjeto dbc_intFonteRecurso, True
        TrocaCorObjeto txt_intFonteRecurso, True
        dbc_intFonteRecurso = ""
        txt_intFonteRecurso = ""
        dbcintcontabancaria.Text = ""
        dbc_intNumeroConta.Text = ""
    End If
End Sub

Private Sub chkbytDisponibilidadeDeCaixa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDividaConsolidada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytEducacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytFundef_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytPensionista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytPessoal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkbytPessoal
End Sub

Private Sub chkbytPessoalInativo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkbytPessoalInativo
End Sub

Private Sub chkbytPrevidenciaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkbytPrevidenciaria
End Sub

Private Sub chkbytProjecaoAtuarial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkbytProjecaoAtuarial
End Sub

Private Sub chkbytRegimePrevidenciario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkbytRegimePrevidenciario
End Sub

Private Sub chkbytSaude_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkbytSaude
End Sub

Private Sub chkbytServicoDeTerceiro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkbytServicoDeTerceiro
End Sub

Private Sub chkbytVariacaoPatrimonial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbc_intFonteRecurso_Change()
    If dbc_intFonteRecurso.MatchedWithList Then
        If Val(dbc_intFonteRecurso.BoundText) > 0 Then
            txt_intFonteRecurso = LeFonteRecurso(dbc_intFonteRecurso.BoundText)
        Else
            txt_intFonteRecurso.Text = ""
        End If
    Else
        txt_intFonteRecurso.Text = ""
    End If
End Sub

Private Sub dbc_intFonteRecurso_GotFocus()
    MarcaCampo dbc_intFonteRecurso
End Sub

Private Sub dbc_intFonteRecurso_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intFonteRecurso, Me, , , Shift
End Sub

Private Sub dbc_intFonteRecurso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intFonteRecurso
End Sub

Private Sub dbc_intNumeroConta_Change()
    If dbc_intNumeroConta.MatchedWithList Then
        If dbc_intNumeroConta.BoundText <> dbcintcontabancaria.BoundText Then
            PreencherListaDeOpcoes dbcintcontabancaria, dbc_intNumeroConta.BoundText
        End If
    End If

End Sub

Private Sub dbc_intNumeroConta_Click(Area As Integer)
    PreencherListaDeOpcoes dbcintcontabancaria, dbc_intNumeroConta.BoundText
End Sub

Private Sub dbc_intNumeroConta_GotFocus()
    If Len(dbc_intNumeroConta.Text) = 0 Then MantemForm gstrPreencherLista
End Sub

Private Sub dbc_intNumeroConta_Validate(Cancel As Boolean)
    If dbc_intNumeroConta.MatchedWithList Then
        PreencherListaDeOpcoes dbcintcontabancaria, dbc_intNumeroConta.BoundText
    Else
        dbcintcontabancaria.BoundText = Space$(0)
        dbc_intNumeroConta.BoundText = Space$(0)
    End If
End Sub

Private Sub dbcintcontabancaria_Change()
    'dbc_intNumeroConta.BoundText = dbcintcontabancaria.BoundText
End Sub

Private Sub dbcintcontabancaria_Click(Area As Integer)
    PreencherListaDeOpcoes dbc_intNumeroConta, dbcintcontabancaria.BoundText
End Sub

Private Sub dbcintModalidade_Change()
    If dbcintModalidade.MatchedWithList Then
       'blnPreencheControles gstrModalidade & " tm, " & gstrConvenio & " tc", "tm.Pkid", gstrRIGHT("'000'" & strCONCAT & "tm.strCodigo", 3), "tc.Pkid " & IIf(Val(txt_intConvenio.Text) > 0, "=", strOUTJOracle & "=" & strOUTJSQLServer) & " tm.intConvenio And tm.Pkid = " & dbcintModalidade.BoundText, txt_intModalidade, gstrISNULL(gstrRIGHT("'00'" & strCONCAT & "tc.strCodigo", 2), "'00'"), txt_intConvenio
        blnPreencheControles gstrModalidade & " tm, " & gstrConvenio & " tc", "tm.Pkid", gstrRIGHT("'000'" & strCONCAT & "tm.strCodigo", 3), "tc.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " tm.intConvenio And tm.Pkid = " & dbcintModalidade.BoundText, txt_intModalidade, gstrISNULL(gstrRIGHT("'00'" & strCONCAT & "tc.strCodigo", 2), "'00'"), txt_intConvenio
    Else
        txt_intModalidade.Text = Space$(0)
        txt_intConvenio.Text = Space$(0)
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 322
    VirificaGradeListView Me
    If MDIMenu.Tag <> "PATRIMONIO" Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir
        HabilitaDesabilitaBotao1 mblnAlterando, gstrBtnArquivo, gstrDeletar
        If mobjAux Is Nothing Then
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrAplicar
        Else
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrAplicar
        End If
        
        If mblnExistemMovimentos = True Then
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
        End If
        
                
        If tab_3dPasta.Tab = 2 And chkblnPatrimonial.Value = 1 Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        End If
        
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrNovo, gstrSalvar, _
                                 gstrDeletar, gstrImprimir, gstrIncluirItem, gstrExcluirItem
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrAplicar, gstrDeletar, gstrImprimir
    If MDIMenu.Tag = "PATRIMONIO" Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrNovo, gstrSalvar
    End If
End Sub

Private Sub Form_Load()
    If MDIMenu.Tag = "PATRIMONIO" Then
        Me.Caption = "Classificação de Bens"
        tab_3dPasta.TabCaption(0) = "Classificação de Bens"
    End If
    mblnAlterando = False
    ' A linha abaixo estava identada MRA
    VerificaListaAutomatica gstrPlanoConta, tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux, mstrQueryAplicar
    tab_3dPasta.TabEnabled(1) = False
    tab_3dPasta.TabEnabled(2) = False
    dbcintcontabancaria.Tag = "Select pkid,strdescricao from " & gstrContaBancaria & " ORDER BY strdescricao;strdescricao"
    dbc_intNumeroConta.Tag = "SELECT pkid, intNumeroConta FROM " & gstrContaBancaria & " ORDER BY intNumeroConta;intNumeroConta"
    dbc_intFonteRecurso.Tag = "select Pkid, strDescricao from " & gstrFonteRecurso & " Where intExercicio = " & gintExercicio & " Order By strDescricao;strDescricao"
    dbcintModalidade.Tag = strQueryModalidade & ";strDescricao"
    
    TrocaCorObjeto dbc_intNumeroConta, True
    TrocaCorObjeto dbcintcontabancaria, True
    TrocaCorObjeto dbc_intFonteRecurso, True
    TrocaCorObjeto txt_intFonteRecurso, True
    tab_3dPasta.Tab = 0
    frm_DadosConta.Visible = False
    TrocaCorObjeto txtIntExtraMaua, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MDIMenu.Tag = "PATRIMONIO" Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar, gstrNovo, gstrImprimir
    End If
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrAplicar, gstrDeletar
    mblnPrimeiraVez = False
End Sub

Private Sub lvw_Credito_ItemClick(ByVal Item As MSComctlLib.ListItem)

    With Item
        If .SubItems(3) = "Financeiro" Then
            opt_FinanceiroCredito.Value = True
        ElseIf .SubItems(3) = "Econômico" Then
            opt_EconomicoCredito.Value = True
        End If
        
        opt_bytTipoCredito(CInt(.SubItems(6))).Value = True
        
        chk_ContaGrupoCredito.Value = IIf(.SubItems(5) = "Sim", 1, 0)
        
        If cbo_intContaCredito.ListCount = 0 Then
            PreencheComboCredito
        End If
        
        cbo_intContaCredito.ListIndex = gintIndiceCBO(cbo_intContaCredito, .Tag)
    End With
    
End Sub


Private Sub lvw_Debito_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
        If .SubItems(3) = "Financeiro" Then
            opt_FinanceiroDebito.Value = True
        ElseIf .SubItems(3) = "Econômico" Then
            opt_EconomicoDebito.Value = True
        End If
        opt_bytTipoDebito(CInt(.SubItems(6))).Value = True
        chk_ContaGrupoDebito.Value = IIf(.SubItems(5) = "Sim", 1, 0)
        
        If cbo_intContaDebito.ListCount = 0 Then
            PreencheComboDebito
        End If
        
        cbo_intContaDebito.ListIndex = gintIndiceCBO(cbo_intContaDebito, .Tag)
    End With
End Sub

Private Sub opt_EconomicoCredito_Click()
    PreencheComboCredito
    If opt_EconomicoCredito.Value = True Then
        chk_ContaGrupoCredito.Enabled = True
    Else
        chk_ContaGrupoCredito.Value = 0
        chk_ContaGrupoCredito.Enabled = False
    End If

End Sub

Private Sub opt_FinanceiroCredito_Click()
    PreencheComboCredito
    If opt_FinanceiroCredito.Value = True Then
        chk_ContaGrupoCredito.Enabled = True
    Else
        chk_ContaGrupoCredito.Value = 0
        chk_ContaGrupoCredito.Enabled = False
    End If
End Sub

Private Sub optblnNaturezaDaConta_Click(Index As Integer)
    If mblnAlterando = True Then AjustaResumoMovimentado
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3dPasta.Tab = 0 Then
        frm_DadosConta.Visible = False
    Else
        frm_DadosConta.Visible = True
    End If
    
    If tab_3dPasta.Tab = 0 Then
        frm_DadosConta.Visible = False
    Else
        frm_DadosConta.Visible = True
    End If
    
    If tab_3dPasta.Tab = 2 And chkblnPatrimonial.Value = 1 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
    
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
    If tab_3dPasta.Tab = 0 Then
        frm_DadosConta.Visible = False
    Else
        frm_DadosConta.Visible = True
    End If
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

'******************************************************************************************
' Data: 09/04/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
'    EnviaTeclaTab vbKeyReturn Identada por MRA
'    ToolBarGeral strModoOperacao, gstrPlanoConta, mblnAlterando, _
'                 tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar, _
'                 rptPlanoContas, "sp_PlanoContas"
    
    Dim strSQL As String
    
    If strModoOperacao = gstrNovo Then
       LimpaObjetos
    End If
    
    If strModoOperacao = gstrSalvar Then
        If Not blnDadosOk Then Exit Sub
    End If
    
    Select Case UCase(strModoOperacao)
        Case UCase(gstrPreencherLista)
            If Me.ActiveControl.Name = cbo_intContaDebito.Name Or Me.ActiveControl.Name = cbo_DescricaoContaDebito.Name Then
                PreencheComboDebito
            End If
            
            If Me.ActiveControl.Name = cbo_intContaCredito.Name Or Me.ActiveControl.Name = cbo_DescricaoContaCredito.Name Then
                PreencheComboCredito
            End If
            
            'Exit Sub
        Case UCase(gstrIncluirItem)
            If tab_3dPasta.Tab = 2 Then
                If tab_3DPastaMovimento.Tab = 0 Then
                    IncluiAlteraListaCredito
                ElseIf tab_3DPastaMovimento.Tab = 1 Then
                    IncluiAlteraListaDebito
                End If
            End If
        Case UCase(gstrExcluirItem)
            If tab_3dPasta.Tab = 2 Then
                If tab_3DPastaMovimento.Tab = 0 Then
                    ExcluirListaCredito
                ElseIf tab_3DPastaMovimento.Tab = 1 Then
                    ExcluirListaDebito
                End If
            End If
        
    End Select
    
    
    
    
    
    If strModoOperacao = gstrSalvar Then
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        If SalvarGeral(gstrPlanoConta, IIf(mblnAlterando, "A", "I"), Me, tdb_Lista, strQuery, False, mblnAlterando) = True Then
            If Not blnAlteraPlanoContaSaldo(IIf(Trim(txtPKId) = "", glngPegaUltimaChave(gstrPlanoConta, "PKID"), txtPKId)) Then
                gobjBanco.ExecutaRollbackTrans
                Exit Sub
            End If
            
            If chkblnPatrimonial.Value = 1 Then
                If gravaCruzamentos(IIf(Trim(txtPKId) = "", glngPegaUltimaChave(gstrPlanoConta, "PKID"), txtPKId)) = True Then
                        gobjBanco.ExecutaCommitTrans
                        MantemForm gstrNovo
                        MantemForm gstrLocalizar
                        Exit Sub
                Else
                    gobjBanco.ExecutaRollbackTrans
                    Exit Sub
                End If
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaCommitTrans
                MantemForm gstrNovo
                MantemForm gstrLocalizar
                Exit Sub
            End If
        Else
            gobjBanco.ExecutaRollbackTrans
            Exit Sub
        End If
    End If
                 
    If strModoOperacao = gstrDeletar And Len(Trim$(txtPKId.Text)) > 0 Then
    
        If gblnExclusaoGravacaoOk("E", txtstrDescricao.Text) Then
        
            Set gobjBanco = New clsBanco
            
            strSQL = IIf(bytDBType = Oracle, "Begin ", Space$(0))
            
            strSQL = strSQL & "Delete From " & gstrPlanoContaSaldo & " Where intPlanoConta = " & txtPKId.Text
            
            strSQL = strSQL & IIf(bytDBType = Oracle, ";", Space$(1))
            
            strSQL = strSQL & "Delete From " & gstrPlanoConta & " Where Pkid = " & txtPKId.Text
            
            strSQL = strSQL & IIf(bytDBType = Oracle, ";", Space$(1))
            
            strSQL = strSQL & IIf(bytDBType = Oracle, "End; ", Space$(0))
            
            gobjBanco.ExecutaBeginTrans
            
            If gobjBanco.Execute(strSQL) Then
                gobjBanco.ExecutaCommitTrans
                MantemForm gstrLocalizar
                LimpaObjeto Me: LimpaObjetos
            Else
                gobjBanco.ExecutaRollbackTrans
            End If
            
        End If
        
        Exit Sub
        
    End If
    
    ToolBarGeral strModoOperacao, gstrPlanoConta, mblnAlterando, _
                 tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar, _
                 rptPlanoContas, gstrStoredProcedure("sp_PlanoContas", , True)

End Sub

Private Function gravaCruzamentos(ByVal ContaPKID As String) As Boolean
    Dim i As Integer
    Dim strSQL As String
    
    gravaCruzamentos = False
    
    strSQL = ""
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    strSQL = strSQL & " DELETE " & gstrCruzamentos & " WHERE IntPlanoContaOrigem  = " & ContaPKID & " ;"
    
    For i = 1 To lvw_Credito.ListItems.Count
        With lvw_Credito.ListItems.Item(i)
            strSQL = strSQL & " INSERT INTO " & gstrCruzamentos
            strSQL = strSQL & " (intplanocontaorigem, intplanocontadestino, bytTipoMovimento, "
            strSQL = strSQL & " byttipocontadestino, byttiposistema,"
            strSQL = strSQL & " byttipolancamento, dtmdtatualizacao , lngcodusr) "
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & ContaPKID & " ," 'intplanocontaorigem
            
            If InStr(1, .Tag, "-") = 0 Then
               strSQL = strSQL & .Tag & "," 'intplanocontadestino
            Else
               strSQL = strSQL & Left(.Tag, InStr(1, .Tag, "-") - 1) & ", " 'intplanocontadestino
            End If
            
            'strSql = strSql & IIf(InStr(1, .Tag, "-") = 0, .Tag, Left(.Tag, InStr(1, .Tag, "-") - 1)) & " ,"  'intplanocontadestino
            
            strSQL = strSQL & .SubItems(6) & " ,"   'bytTipoMovimentacao
            strSQL = strSQL & IIf(.SubItems(5) = "Sim", "1", "0") & " ," 'byttipocontadestino
            strSQL = strSQL & IIf(.SubItems(3) = "Financeiro", "1", "2") & " ," 'byttiposistema
            strSQL = strSQL & " 0 ," 'byttipolancamento
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & " ," 'dtmdtatualizacao
            strSQL = strSQL & glngCodUsr 'lngcodusr
            strSQL = strSQL & " );"
        End With
    Next
    
    For i = 1 To lvw_Debito.ListItems.Count
        With lvw_Debito.ListItems.Item(i)
            strSQL = strSQL & " INSERT INTO " & gstrCruzamentos
            strSQL = strSQL & " (intplanocontaorigem, intplanocontadestino, bytTipoMovimento, "
            strSQL = strSQL & " byttipocontadestino, byttiposistema,"
            strSQL = strSQL & " byttipolancamento, dtmdtatualizacao , lngcodusr) "
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & ContaPKID & " ," 'intplanocontaorigem
            
            
            If InStr(1, .Tag, "-") = 0 Then
               strSQL = strSQL & .Tag & "," 'intplanocontadestino
            Else
               strSQL = strSQL & Left(.Tag, InStr(1, .Tag, "-") - 1) & ", " 'intplanocontadestino
            End If
            
            'strSql = strSql & IIf(InStr(1, .Tag, "-") = 0, .Tag, Left(.Tag, InStr(1, .Tag, "-") - 1)) & " ,"  'intplanocontadestino
            
            strSQL = strSQL & .SubItems(6) & " ,"   'bytTipoMovimentacao
            strSQL = strSQL & IIf(.SubItems(5) = "Sim", "1", "0") & " ," 'byttipocontadestino
            strSQL = strSQL & IIf(.SubItems(3) = "Financeiro", "1", "2") & " ," 'byttiposistema
            strSQL = strSQL & " 1 ," 'byttipolancamento
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & " ," 'dtmdtatualizacao
            strSQL = strSQL & glngCodUsr 'lngcodusr
            strSQL = strSQL & " );"
        End With
    Next
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
    
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSQL) Then
        gravaCruzamentos = True
    End If
End Function


Private Function VerificaMovimentos() As Boolean
    Dim strSQL          As String
    Dim adoResultado    As New ADODB.Recordset
    
    VerificaMovimentos = False

    strSQL = "SELECT "

    If (bytDBType = EDatabases.SQLServer) Then
        strSQL = strSQL & " TOP 1 "
    End If
    
    strSQL = strSQL & " PKID FROM " & gstrLancamentoContabil & " WHERE intconta =" & txtPKId.Text
    
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " AND Rownum = 1"
    End If
            
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                VerificaMovimentos = True
            End If
        End With
    End If
    
End Function


Private Sub PreencheComboCredito()
    Dim strSQL As String
    Dim adoResultado As New ADODB.Recordset
    
    strSQL = "SELECT PKID,  strDescricao , strContaContabil FROM "
    strSQL = strSQL & gstrPlanoConta & " PC "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " ABS(PC.blnPatrimonial) = 0 "
    
    If opt_FinanceiroCredito.Value = True Then
        strSQL = strSQL & " AND ABS(PC.blnFinanceira) = 1 "
    End If
    
    If opt_EconomicoCredito.Value = True Then
        strSQL = strSQL & " AND ABS(PC.bytvariacaopatrimonial) = 1 "
    End If
    
    If chk_ContaGrupoCredito.Value = 0 Then
        strSQL = strSQL & " AND ABS(PC.blnAnalitica) = 1 "
    Else
        strSQL = strSQL & " AND ABS(PC.blnAnalitica) = 0 "
    End If
    
    cbo_intContaCredito.Clear
    cbo_DescricaoContaCredito.Clear
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                cbo_intContaCredito.AddItem gvntFormatacaoEspecifica(!strContaContabil)
                cbo_intContaCredito.ItemData(cbo_intContaCredito.NewIndex) = !Pkid

                cbo_DescricaoContaCredito.AddItem !strDescricao
                cbo_DescricaoContaCredito.ItemData(cbo_DescricaoContaCredito.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If
    
End Sub

Private Sub PreencheComboDebito()
    Dim strSQL As String
    Dim adoResultado As New ADODB.Recordset
    
    strSQL = "SELECT PKID, strDescricao , strContaContabil FROM "
    strSQL = strSQL & gstrPlanoConta & " PC "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " ABS(PC.blnPatrimonial) = 0 "
    
    If opt_FinanceiroDebito.Value = True Then
        strSQL = strSQL & " AND ABS(PC.blnFinanceira) = 1 "
    End If
    
    If opt_EconomicoDebito.Value = True Then
        strSQL = strSQL & " AND ABS(PC.bytvariacaopatrimonial) = 1 "
    End If
    
    If chk_ContaGrupoDebito.Value = 0 Then
        strSQL = strSQL & " AND ABS(PC.blnAnalitica) = 1 "
    Else
        strSQL = strSQL & " AND ABS(PC.blnAnalitica) = 0 "
    End If
    
    
    cbo_intContaDebito.Clear
    cbo_DescricaoContaDebito.Clear
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                cbo_intContaDebito.AddItem gvntFormatacaoEspecifica(!strContaContabil)
                cbo_intContaDebito.ItemData(cbo_intContaDebito.NewIndex) = !Pkid

                cbo_DescricaoContaDebito.AddItem !strDescricao
                cbo_DescricaoContaDebito.ItemData(cbo_DescricaoContaDebito.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If

End Sub


Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 1 Then
        Value = gvntFormatacaoEspecifica(Value, 1)
    End If
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    If tdb_Lista.Col = 1 Then
        CaracterValido KeyAscii, "N", tdb_Lista
    Else
        CaracterValido KeyAscii
    End If
End Sub

Private Sub optblnNaturezaDaConta_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii
    If mblnAlterando = True Then AjustaResumoMovimentado
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            If mblnPrimeiraVez Then
                limpaCamposCruzamentos
                mblnClickOk = False
                txtPKId = .Columns(0).Value
                LeDaTabelaParaObj gstrPlanoConta, Me
                
                PreencherListaDeOpcoes dbc_intFonteRecurso, RetornaFonteRecurso(txtPKId)
                
                'LeSaldoDaConta txtPKId
                If blnVerificaMovi Then
                    chkblnFinanceira.Enabled = False
                    chkbytDisponibilidadeDeCaixa.Enabled = False
                    chkblnAnalitica.Enabled = False
                End If
                VerificaAnalitica

                mblnExistemMovimentos = VerificaMovimentos
                
                If MDIMenu.Tag = "PATRIMONIO" Then
                    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
                Else
                    If mblnExistemMovimentos = True Then
                        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
                    End If
                    
                    If tab_3dPasta.Tab = 2 And chkblnPatrimonial.Value = 1 Then
                        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
                    Else
                        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
                    End If
                    
                End If
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrAplicar
                End If
                mblnAlterando = True
                
                PreencheLIstaCredito
                PreencheLIstaDebito
                
                strContaAtual = txtstrContaContabil.Text
                strDescricaoAtual = txtstrDescricao.Text
                strContaExtraAtual = txtIntExtraMaua.Text
                tab_3dPasta.TabEnabled(1) = True
                CalculaSaldoInicial
                leResumoMovimentado
                PreencherListaDeOpcoes dbc_intNumeroConta, dbcintcontabancaria.BoundText
                If chkblnPatrimonial.Value = 0 Then
                    If tab_3dPasta.Tab = 2 Then tab_3dPasta.Tab = 0
                    tab_3dPasta.TabEnabled(2) = False
                End If
            End If
        End If
    End With
End Sub

Private Sub PreencheLIstaCredito()
    Dim strSQL As String
    
    strSQL = " SELECT"
    strSQL = strSQL & " PC.PKID " & strCONCAT & "'-'" & strCONCAT & " CR.BytTipoMovimento PKID,"
    strSQL = strSQL & " PC.PKID " & strCONCAT & "'-'" & strCONCAT & " CR.BytTipoMovimento PKID,"
    strSQL = strSQL & " PC.STRCONTACONTABIL,"
    strSQL = strSQL & " PC.Strdescricao,"
    strSQL = strSQL & gstrCASEWHEN("CR.BYTTIPOSISTEMA", "1,'Financeiro',2,'Econômico'") & " Sistema,"
    strSQL = strSQL & gstrCASEWHEN("CR.BytTipoMovimento", _
            "0, 'Diversos', 1, 'Independente', 2, 'Mutação', 3, 'Cancelamento'") & " AS strTipoMovimento, "
    strSQL = strSQL & gstrCASEWHEN("Cr.BytTipoContaDestino", "0,'Não',1,'Sim'") & " Grupo ,"
    strSQL = strSQL & "CR.BytTipoMovimento "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrCruzamentos & " CR "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " PC.PKID = CR.Intplanocontadestino"
    strSQL = strSQL & " AND CR.BYTTIPOLANCAMENTO = 0"
    strSQL = strSQL & " AND CR.Intplanocontaorigem = " & txtPKId.Text
    
    
    LeDaTabelaParaObj "", lvw_Credito, strSQL
    
End Sub

Private Sub PreencheLIstaDebito()
    Dim strSQL As String
    
    strSQL = " SELECT"
    strSQL = strSQL & " PC.PKID " & strCONCAT & "'-'" & strCONCAT & " CR.BytTipoMovimento PKID,"
    strSQL = strSQL & " PC.PKID " & strCONCAT & "'-'" & strCONCAT & " CR.BytTipoMovimento PKID,"
    strSQL = strSQL & " PC.STRCONTACONTABIL,"
    strSQL = strSQL & " PC.Strdescricao,"
    strSQL = strSQL & gstrCASEWHEN("CR.BYTTIPOSISTEMA", "1,'Financeiro',2,'Econômico'") & " Sistema,"
    strSQL = strSQL & gstrCASEWHEN("CR.BytTipoMovimento", _
            "0, 'Diversos', 1, 'Independente', 2, 'Mutação', 3, 'Cancelamento'") & " AS strTipoMovimento, "
    strSQL = strSQL & gstrCASEWHEN("CR.BytTipoContaDestino", "0,'Não',1,'Sim'") & " Grupo , "
    strSQL = strSQL & "CR.BytTipoMovimento "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrCruzamentos & " CR "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " PC.PKID = CR.Intplanocontadestino"
    strSQL = strSQL & " AND CR.BYTTIPOLANCAMENTO = 1"
    strSQL = strSQL & " AND CR.Intplanocontaorigem = " & txtPKId.Text
    
    
    LeDaTabelaParaObj "", lvw_Debito, strSQL
    
End Sub


Private Sub ExcluirListaCredito()
    If Not lvw_Credito.SelectedItem Is Nothing Then
        lvw_Credito.ListItems.Remove lvw_Credito.SelectedItem.Index
        cbo_intContaCredito.ListIndex = -1
        cbo_DescricaoContaCredito.ListIndex = -1
    End If
End Sub

Private Sub ExcluirListaDebito()
    If Not lvw_Debito.SelectedItem Is Nothing Then
        lvw_Debito.ListItems.Remove lvw_Debito.SelectedItem.Index
        cbo_intContaDebito.ListIndex = -1
        cbo_DescricaoContaDebito.ListIndex = -1
    End If
End Sub

Private Sub IncluiAlteraListaCredito()
    Dim mobjLista As Object
    Dim strTipoMovimentacao As String
    Dim intTipoMovimentacao As Integer

    If blnDadosOKCredito = True Then
    
        '"0, 'Diversos', 1, 'Independente', 2, 'Mutação', 3, 'Cancelamento'")
        If opt_bytTipoCredito(0).Value = True Then strTipoMovimentacao = "Diversos": intTipoMovimentacao = 0
        If opt_bytTipoCredito(1).Value = True Then strTipoMovimentacao = "Independente": intTipoMovimentacao = 1
        If opt_bytTipoCredito(2).Value = True Then strTipoMovimentacao = "Mutação": intTipoMovimentacao = 2
        If opt_bytTipoCredito(3).Value = True Then strTipoMovimentacao = "Cancelamento": intTipoMovimentacao = 3
    
        Set mobjLista = lvw_Credito.ListItems.Add(, , gstrItemData(cbo_intContaCredito) & "-" & intTipoMovimentacao)
        mobjLista.SubItems(1) = cbo_intContaCredito.Text 'Conta
        mobjLista.SubItems(2) = cbo_DescricaoContaCredito.Text 'Descrição
        mobjLista.SubItems(3) = IIf(opt_FinanceiroCredito.Value, "Financeiro", "Econômico") 'Sistema
        
    
        mobjLista.SubItems(4) = strTipoMovimentacao 'Tipo Movimentacao
        mobjLista.SubItems(5) = IIf(chk_ContaGrupoCredito.Value = 1, "Sim", "Não") ' Grupo
        mobjLista.SubItems(6) = intTipoMovimentacao ' N° Tipo Movimentacao (invisivel)
        mobjLista.Tag = gstrItemData(cbo_intContaCredito)
        
        cbo_intContaCredito.ListIndex = -1
        cbo_DescricaoContaCredito.ListIndex = -1
    End If
End Sub

Private Sub IncluiAlteraListaDebito()
    Dim mobjLista As Object
    Dim strTipoMovimentacao As String
    Dim intTipoMovimentacao As Integer
    
    If blnDadosOKDebito = True Then
    
        '0, 'Diversos', 1, 'Independente', 2, 'Mutação', 3, 'Cancelamento'
        If opt_bytTipoDebito(0).Value = True Then strTipoMovimentacao = "Diversos": intTipoMovimentacao = 0
        If opt_bytTipoDebito(1).Value = True Then strTipoMovimentacao = "Independente": intTipoMovimentacao = 1
        If opt_bytTipoDebito(2).Value = True Then strTipoMovimentacao = "Mutação": intTipoMovimentacao = 2
        If opt_bytTipoDebito(3).Value = True Then strTipoMovimentacao = "Cancelamento": intTipoMovimentacao = 3
    
        Set mobjLista = lvw_Debito.ListItems.Add(, , gstrItemData(cbo_intContaDebito) & "-" & intTipoMovimentacao)
        mobjLista.SubItems(1) = cbo_intContaDebito.Text 'Conta
        mobjLista.SubItems(2) = cbo_DescricaoContaDebito.Text 'Descrição
        mobjLista.SubItems(3) = IIf(opt_FinanceiroDebito.Value, "Financeiro", "Econômico") 'Sistema
        
        
        mobjLista.SubItems(4) = strTipoMovimentacao 'Tipo Movimentacao
        mobjLista.SubItems(5) = IIf(chk_ContaGrupoDebito.Value = 1, "Sim", "Não") ' Grupo
        mobjLista.SubItems(6) = intTipoMovimentacao ' N° Tipo Movimentacao (invisivel)
        mobjLista.Tag = gstrItemData(cbo_intContaDebito)
        cbo_intContaDebito.ListIndex = -1
        cbo_DescricaoContaDebito.ListIndex = -1
    End If
End Sub


Private Function blnDadosOKCredito() As Boolean
    
    Dim intTipoMovimentacao As Integer
    
    blnDadosOKCredito = False
    
    '"0, 'Diversos', 1, 'Independente', 2, 'Mutação', 3, 'Cancelamento'")
    If opt_bytTipoCredito(0).Value = True Then intTipoMovimentacao = 0
    If opt_bytTipoCredito(1).Value = True Then intTipoMovimentacao = 1
    If opt_bytTipoCredito(2).Value = True Then intTipoMovimentacao = 2
    If opt_bytTipoCredito(3).Value = True Then intTipoMovimentacao = 3
    
    
    If gstrItemData(cbo_intContaCredito) = 0 Then
        ExibeMensagem "A conta deve ser informada corretamente."
        If cbo_intContaCredito.Enabled Then cbo_intContaCredito.SetFocus
        Exit Function
    End If
    
    If gblnEncontroItemNoListView(lvw_Credito, gstrItemData(cbo_intContaCredito) & "-" & intTipoMovimentacao, 0) Then
        ExibeMensagem "Esta conta já se encontra na lista."
        If cbo_intContaCredito.Enabled Then cbo_intContaCredito.SetFocus
        Exit Function
    End If
    
    blnDadosOKCredito = True
        
End Function


Private Function blnDadosOKDebito() As Boolean
    Dim intTipoMovimentacao As Integer
    blnDadosOKDebito = False
    
    '0, 'Diversos', 1, 'Independente', 2, 'Mutação', 3, 'Cancelamento'
    If opt_bytTipoDebito(0).Value = True Then intTipoMovimentacao = 0
    If opt_bytTipoDebito(1).Value = True Then intTipoMovimentacao = 1
    If opt_bytTipoDebito(2).Value = True Then intTipoMovimentacao = 2
    If opt_bytTipoDebito(3).Value = True Then intTipoMovimentacao = 3
    
    
    If gstrItemData(cbo_intContaDebito) = 0 Then
        ExibeMensagem "A conta deve ser informada corretamente."
        If cbo_intContaDebito.Enabled Then cbo_intContaDebito.SetFocus
        Exit Function
    End If
        
        
    If gblnEncontroItemNoListView(lvw_Debito, gstrItemData(cbo_intContaDebito) & "-" & intTipoMovimentacao, 0) Then
        ExibeMensagem "Esta conta já se encontra na lista."
        If cbo_intContaDebito.Enabled Then cbo_intContaDebito.SetFocus
        Exit Function
    End If
    
    blnDadosOKDebito = True
End Function

Private Sub txt_intConvenio_Validate(Cancel As Boolean)
    If Len(Trim$(txt_intModalidade.Text)) > 0 And Len(Trim$(txt_intConvenio.Text)) > 0 Then
        If Not blnPreencheControles(gstrModalidade & " tm, " & gstrConvenio & " tc", "tm.Pkid", "tm.strDescricao", "tc.Pkid " & IIf(Val(txt_intConvenio.Text) > 0, "=", strOUTJOracle & "=" & strOUTJSQLServer) & " tm.intConvenio And " & gstrCONVERT(CDT_INT, "Ltrim(Rtrim(tm.strCodigo))") & " = '" & Trim$(txt_intModalidade.Text) & "'" & IIf(Val(txt_intConvenio.Text) > 0, " And " & gstrCONVERT(CDT_INT, "Ltrim(Rtrim(tc.strCodigo))") & " = '" & Val(txt_intConvenio.Text) & "'", Space$(0)), dbcintModalidade) Then
            txt_intConvenio.Text = Space$(0): Cancel = True
        End If
    ElseIf Len(Trim$(txt_intModalidade.Text)) > 0 Then
        If Not blnPreencheControles(gstrModalidade, "Pkid", "strDescricao", "Ltrim(Rtrim(strCodigo)) = '" & Trim$(txt_intModalidade.Text) & "'", dbcintModalidade) Then
            txt_intConvenio.Text = "00"
        End If
    End If
End Sub

Private Sub txt_intFonteRecurso_GotFocus()
    MarcaCampo txt_intFonteRecurso
End Sub

Private Sub txt_intFonteRecurso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intFonteRecurso
End Sub

Private Sub txt_intFonteRecurso_LostFocus()
    If Len(txt_intFonteRecurso) > 0 Then
        PreencherListaDeOpcoes dbc_intFonteRecurso, LeFonteRecurso(, txt_intFonteRecurso)
    Else
        dbc_intFonteRecurso.Text = ""
    End If
End Sub

Private Sub txt_intModalidade_Validate(Cancel As Boolean)
    If Len(Trim$(txt_intModalidade.Text)) > 0 And Len(Trim$(txt_intConvenio.Text)) > 0 Then
        If Not blnPreencheControles(gstrModalidade & " tm, " & gstrConvenio & " tc", "tm.Pkid", "tm.strDescricao", "tc.Pkid " & IIf(Val(txt_intConvenio.Text) > 0, "=", strOUTJOracle & "=" & strOUTJSQLServer) & " tm.intConvenio And " & gstrCONVERT(CDT_INT, "Ltrim(Rtrim(tm.strCodigo))") & " = '" & Val(txt_intModalidade.Text) & "'" & IIf(Val(txt_intConvenio.Text) > 0, " And " & gstrCONVERT(CDT_INT, "Ltrim(Rtrim(tc.strCodigo))") & " = '" & Val(txt_intConvenio.Text) & "'", Space$(0)), dbcintModalidade) Then
            txt_intModalidade.Text = Space$(0): Cancel = True
        End If
'    ElseIf Len(Trim$(txt_intModalidade.Text)) > 0 Then
'        If Not blnPreencheControles(gstrModalidade, "Pkid", "strDescricao", "Ltrim(Rtrim(strCodigo)) = '" & Trim$(txt_intModalidade.Text) & "'", dbcintModalidade) Then
'            txt_intConvenio.Text = "00"
'        End If
    Else
        dbcintModalidade.Text = Space$(0)
    End If
End Sub

Private Sub txt_SaldoAtual_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_SaldoAtual
End Sub

Private Sub txtintContaReduzida_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtIntExtraMaua_GotFocus()
    MarcaCampo txtIntExtraMaua
End Sub

Private Sub txtIntExtraMaua_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtIntExtraMaua
End Sub

Private Sub txtstrContaContabil_GotFocus()
    txtstrContaContabil = gstrValorSemMascara(txtstrContaContabil)
    MarcaCampo txtstrContaContabil
End Sub

Private Sub txtstrContaContabil_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrContaContabil
End Sub

Private Sub txtstrContaContabil_LostFocus()
    If Val(InStr(1, txtstrContaContabil, ".")) < 1 Then
        txtstrContaContabil = gvntFormatacaoEspecifica(txtstrContaContabil)
        
    End If
    'VerificaSeEAnalitica
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtdblSaldoDaConta_GotFocus()
    MarcaCampo txtdblSaldoDaConta
End Sub

Private Sub txtdblSaldoDaConta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii > 47 And KeyAscii < 58 And chkblnAnalitica = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    CaracterValido KeyAscii, "V", txtdblSaldoDaConta
End Sub

Private Sub txtdblSaldoDaConta_LostFocus()
    txtdblSaldoDaConta = gstrConvVrDoSql(txtdblSaldoDaConta)
End Sub

Private Sub LeSaldoDaConta(intPkid As Integer)

'******************************************************************************************
' Data: 09/04/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/04/2003
' Alteração: - Substituição da chamada à função CriaADO por uma chamada à função
'            ExecuteStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL          As String
'    Dim adoResultado    As ADODB.Recordset
    Dim adoResultado    As New ADODB.Recordset
'    strSql = "sp_SaldoAtualDaConta " & intPKID

    strSQL = gstrStoredProcedure("sp_SaldoAtualDaConta", CStr(intPkid) & ",NULL,NULL,NULL,NULL,NULL,NULL," & gintExercicio, True)
            
    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txt_SaldoAtual = gstrConvVrDoSql(!dblSaldo)
            End If
        End With
    End If
End Sub

Private Function blnDadosOk() As Boolean
    Dim strDescricao As String
    
    strDescricao = Replace(Replace(Replace(txtstrDescricao.Text, Chr(13), ""), Chr(9), ""), Chr(10), "")
    
    If Trim(strDescricao) = "" Then
        ExibeMensagem "O campo Descrição deve ser preenchido!"
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    
    If Not CampoObrigatorio(txtstrContaContabil, "Conta contábil") Then Exit Function
    
    If optblnNaturezaDaConta(0).Value = False And optblnNaturezaDaConta(1).Value = False Then
        ExibeMensagem "O campo Natureza da Conta deve ser preenchido!"
        'optblnNaturezaDaConta.SetFocus
        Exit Function
    End If
        
    If Not mblnAlterando Or (mblnAlterando And UCase(txtstrContaContabil.Text) <> UCase(strContaAtual)) Then

        If gblnExisteCodigo(1, gstrPlanoConta, "strContaContabil", "'" & gvntConvFormatoEspecificoParaSQL(txtstrContaContabil.Text, 1) & "'") Then
            ExibeMensagem "O valor digitado para o campo Conta Contábil já se encontra cadastrado!"
            txtstrContaContabil.SetFocus
            Exit Function
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase(txtstrDescricao.Text) <> UCase(strDescricaoAtual)) Then
        If gblnExisteCodigo(1, gstrPlanoConta, "strDescricao", "'" & txtstrContaContabil.Text & "'") Then
            ExibeMensagem "O valor digitado para o campo Descrição já se encontra cadastrado!"
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    
    If chkblnExtraOrcamentaria.Value = 1 And Val(txtIntExtraMaua) = 0 Then
        ExibeMensagem "É necessário informar uma Conta Extra."
        txtIntExtraMaua.SetFocus
        Exit Function
    End If
    
    
    If Not mblnAlterando Or (mblnAlterando And UCase(txtIntExtraMaua.Text) <> UCase(strContaExtraAtual)) Then
        If gblnExisteCodigo(1, gstrPlanoConta, "IntExtraMaua", txtIntExtraMaua.Text) Then
            ExibeMensagem "A conta extra informada já se encontra cadastrada!"
            txtIntExtraMaua.SetFocus
            Exit Function
        End If
    End If
    
    If chkbytDisponibilidadeDeCaixa = vbChecked Then
        If Not dbc_intFonteRecurso.MatchedWithList Then
            ExibeMensagem "A Fonte de Recurso deve ser preenchida corretamente."
            If dbc_intFonteRecurso.Enabled Then dbc_intFonteRecurso.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True
    
End Function
Private Sub VerificaAnalitica()

   Dim strSQL       As String
   Dim adoResultado As New ADODB.Recordset
   
   strSQL = "SELECT * FROM " & gstrEventoContaContabilCredito
   strSQL = strSQL & " WHERE intContaContabil = " & txtPKId
           
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         chkblnAnalitica.Enabled = False
         adoResultado.Close
         Exit Sub
      End If
      adoResultado.Close
   End If
   
   strSQL = "SELECT * FROM " & gstrEventoContaContabilDebito
   strSQL = strSQL & " WHERE intContaContabil = " & txtPKId
           
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         chkblnAnalitica.Enabled = False
         adoResultado.Close
         Exit Sub
      End If
      adoResultado.Close
   End If
   
   chkblnAnalitica.Enabled = True

End Sub
Private Sub LimpaObjetos()
    txt_SaldoAtual = ""
    chkblnAnalitica.Enabled = True
    chkblnFinanceira.Enabled = True
    chkbytDisponibilidadeDeCaixa.Enabled = True
    tab_3dPasta.Tab = 0
    tab_3dPasta.TabEnabled(1) = False
    
    txt_intFonteRecurso.Text = ""
    dbc_intFonteRecurso.Text = ""
    Set dbc_intFonteRecurso.RowSource = Nothing
    
    txt_intModalidade.Text = ""
    txt_intConvenio.Text = ""
    dbcintModalidade.Text = ""
    Set dbcintModalidade.RowSource = Nothing
    
   limpaCamposCruzamentos
End Sub

Private Sub limpaCamposCruzamentos()
    opt_FinanceiroCredito.Value = True
    opt_bytTipoCredito(0).Value = True
    chk_ContaGrupoCredito.Value = 0
    cbo_intContaCredito.Text = ""
    cbo_DescricaoContaCredito.Text = ""
    lvw_Credito.ListItems.Clear
    
    opt_FinanceiroDebito.Value = True
    opt_bytTipoDebito(0).Value = True
    chk_ContaGrupoDebito.Value = 0
    cbo_intContaDebito.Text = ""
    cbo_DescricaoContaDebito.Text = ""
    lvw_Debito.ListItems.Clear

End Sub

Private Sub leResumoMovimentado()
    txt_strConta.Text = txtstrContaContabil
    txt_strDescricao.Text = gstrENulo(tdb_Lista.Columns(2).Value)
    'txt_dblSaldoInicial.Text = IIf(Trim(txtdblSaldoDaConta.Text) = "", "0,00", txtdblSaldoDaConta.Text)
    LeDaTabelaParaObj gstrPlanoConta, lvw_Resumo, strQueryResumo
    AjustaResumoMovimentado
End Sub

Private Function strQueryResumo() As String

    Dim strSQL              As String
    Dim i                   As Integer
    Dim intPosicao          As Integer
    Dim mstrCodigoSemPontos As String
    Dim mstrMaiorHierarquia As String
    
    
    
    mstrCodigoSemPontos = gstrValorSemMascara(tdb_Lista.Columns(1).Value)
    mstrMaiorHierarquia = mstrCodigoSemPontos
    
    intPosicao = 0
    
    For i = 1 To Len(gstrMascaraContaContabil)
        If Mid(gstrMascaraContaContabil, i, 1) <> "0" Then
           If Val(Mid(mstrMaiorHierarquia, intPosicao + 1, Len(mstrMaiorHierarquia))) = 0 Then
              mstrMaiorHierarquia = Mid(mstrMaiorHierarquia, 1, intPosicao)
              Exit For
           End If
        Else
           intPosicao = intPosicao + 1
        End If
    Next
    
    strSQL = ""

    strSQL = strSQL & " SELECT 0 PKID, "
    strSQL = strSQL & gstrCASEWHEN("RE.Mes", "1,'Janeiro',2,'Fevereiro',3,'Março'," & _
    "4,'Abril',5,'Maio',6,'Junho',7,'Julho',8,'Agosto',9,'Setembro',10,'Outubro',11,'Novembro',12,'Dezembro'") & "Meses"
    strSQL = strSQL & ", SUM( RE.Credito) Credito,  SUM(RE.Debito) Debito, (SUM( RE.Credito) - SUM(RE.Debito)) Saldo, "
    'strSQL = strSQL & gstrCASEWHEN("(SUM( RE.Credito) - SUM(RE.Debito))", ">=0 ,'Credito', < 0,'Debito'") & "TipoSaldo "
    strSQL = strSQL & " '' TipoSaldo "
        
    strSQL = strSQL & "FROM ( "
    strSQL = strSQL & " SELECT 0 PKID, "
    
    
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " TO_NUMBER(TO_CHAR(PP.dtmData , 'MM')) "
    Else
        strSQL = strSQL & " Month (PP.dtmData)"
    End If
    strSQL = strSQL & " Mes, LC.dblValor Credito, 0 Debito,0 Saldo, '' tipoSaldo FROM "
    strSQL = strSQL & " " & gstrProcessoPagamento & " PP,"
    strSQL = strSQL & " " & gstrPlanoConta & " PC,"
    strSQL = strSQL & " " & gstrLancamentoContabil & " LC "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " LC.intProcesso = PP.PKID AND "
    strSQL = strSQL & " LC.intConta = PC.PKID AND "
    strSQL = strSQL & " PP.bytNormal = 1 AND "
    strSQL = strSQL & " LC.bytNatureza = 0 AND "
    
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " TO_NUMBER(TO_CHAR(PP.dtmData , 'YYYY')) = " & CStr(gintExercicio) & " AND "
    Else
        strSQL = strSQL & " YEAR(PP.dtmData) = " & CStr(gintExercicio) & " AND "
    End If
    
    'strSQL = strSQL & " PC.Pkid = " & txtPKId
    strSQL = strSQL & " " & strSUBSTRING & " (PC.strContaContabil,1," + CStr(Len(mstrMaiorHierarquia)) & ") = '" & mstrMaiorHierarquia & "'"
    
    
    strSQL = strSQL & " UNION All"
    strSQL = strSQL & " SELECT 0 ,"
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " TO_NUMBER(TO_CHAR(PP.dtmData , 'MM')) "
    Else
        strSQL = strSQL & " MONTH (PP.dtmData)"
    End If
    strSQL = strSQL & " Mes, 0 Credito, LC.dblValor Debito, 0 Saldo,  '' tipoSaldo FROM"
    strSQL = strSQL & " " & gstrProcessoPagamento & " PP,"
    strSQL = strSQL & " " & gstrPlanoConta & " PC,"
    strSQL = strSQL & " " & gstrLancamentoContabil & " LC"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " LC.intProcesso = PP.PKID AND"
    strSQL = strSQL & " LC.intConta = PC.PKID AND"
    strSQL = strSQL & " PP.bytNormal = 1 AND"
    strSQL = strSQL & " LC.bytNatureza = 1 AND"
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " TO_NUMBER(TO_CHAR(PP.dtmData , 'YYYY'))= " & CStr(gintExercicio) & " AND "
    Else
        strSQL = strSQL & " YEAR(PP.dtmData)  =" & CStr(gintExercicio) & " AND "
    End If
    'strSQL = strSQL & " PC.PKID =" & txtPKId
    strSQL = strSQL & " " & strSUBSTRING & " (PC.strContaContabil,1," + CStr(Len(mstrMaiorHierarquia)) & ") = '" & mstrMaiorHierarquia & "'"
       
       
' UNION ACRESCENTADA PARA CRUZAMENTOS ############################################
       
    'Credito ************************
    strSQL = strSQL & " UNION All"
    strSQL = strSQL & " SELECT 0 ,"
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " TO_NUMBER(TO_CHAR(MS.dtmData , 'MM')) "
    Else
        strSQL = strSQL & " MONTH (MS.dtmData)"
    End If
    strSQL = strSQL & " Mes, CM.dblValor Credito, 0 Debito, 0 Saldo,  '' tipoSaldo FROM"
    strSQL = strSQL & " " & gstrMovimentoSistemas & " MS,"
    strSQL = strSQL & " " & gstrPlanoConta & " PC,"
    strSQL = strSQL & " " & gstrContaMovimentoSistemas & " CM"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " CM.intmovimentoSistema = MS.PKID AND"
    strSQL = strSQL & " CM.intPlanoConta = PC.PKID AND"
    strSQL = strSQL & " CM.bytTipo = 0 AND"
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " TO_NUMBER(TO_CHAR(MS.dtmData , 'YYYY'))= " & CStr(gintExercicio) & " AND "
    Else
        strSQL = strSQL & " YEAR(MS.dtmData)  =" & CStr(gintExercicio) & " AND "
    End If
    'strSQL = strSQL & " PC.PKID =" & txtPKId
    strSQL = strSQL & " " & strSUBSTRING & " (PC.strContaContabil,1," + CStr(Len(mstrMaiorHierarquia)) & ") = '" & mstrMaiorHierarquia & "'"
       
       
    'DEBITO ************************
    strSQL = strSQL & " UNION All"
    strSQL = strSQL & " SELECT 0 ,"
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " TO_NUMBER(TO_CHAR(MS.dtmData , 'MM')) "
    Else
        strSQL = strSQL & " MONTH (MS.dtmData)"
    End If
    strSQL = strSQL & " Mes, 0 Credito, CM.dblValor Debito, 0 Saldo,  '' tipoSaldo FROM"
    strSQL = strSQL & " " & gstrMovimentoSistemas & " MS,"
    strSQL = strSQL & " " & gstrPlanoConta & " PC,"
    strSQL = strSQL & " " & gstrContaMovimentoSistemas & " CM"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " CM.intmovimentoSistema = MS.PKID AND"
    strSQL = strSQL & " CM.intPlanoConta = PC.PKID AND"
    strSQL = strSQL & " CM.bytTipo = 1 AND"
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " TO_NUMBER(TO_CHAR(MS.dtmData , 'YYYY'))= " & CStr(gintExercicio) & " AND "
    Else
        strSQL = strSQL & " YEAR(MS.dtmData)  =" & CStr(gintExercicio) & " AND "
    End If
    'strSQL = strSQL & " PC.PKID =" & txtPKId
    strSQL = strSQL & " " & strSUBSTRING & " (PC.strContaContabil,1," + CStr(Len(mstrMaiorHierarquia)) & ") = '" & mstrMaiorHierarquia & "'"
       
'#################################################################################
       
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " UNION ALL SELECT 0,1,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,2,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,3,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,4,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,5,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,6,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,7,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,8,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,9,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,10,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,11,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
        strSQL = strSQL & " UNION ALL SELECT 0,12,0,0,0,'' FROM " & gstrProcessoPagamento & " WHERE ROWNUM =1"
    Else
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,1,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,2,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,3,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,4,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,5,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,6,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,7,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,8,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,9,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,10,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,11,0,0,0,'' FROM " & gstrProcessoPagamento
        strSQL = strSQL & " UNION ALL SELECT TOP 1 0,12,0,0,0,'' FROM " & gstrProcessoPagamento
    End If
    
    
    strSQL = strSQL & ") RE GROUP BY RE.Mes"
    strSQL = strSQL & " ORDER BY RE.Mes"
    
    strQueryResumo = strSQL
End Function


Private Sub AjustaResumoMovimentado()
    Dim i  As Integer
    Dim mdblSaldoInicial As Double
    Dim dblValorTotal As Double
    Dim dblvalorAcumulado As Double
    
    If Val(txt_dblSaldoInicial) = 0 Then
        mdblSaldoInicial = 0
    Else
        mdblSaldoInicial = CDbl(txt_dblSaldoInicial)
    End If
    
    
    For i = 1 To lvw_Resumo.ListItems.Count
        With lvw_Resumo
            
            If i = 1 Then
            
                'If optblnNaturezaDaConta(0).Value = False Then
                If mdblSaldoInicial < 0 Then
                    dblValorTotal = Abs(mdblSaldoInicial) * -1 + CDbl(.ListItems(i).ListSubItems(1)) - CDbl(.ListItems(i).ListSubItems(2))
                    lbl_Natureza.Caption = "DB"
                Else
                    dblValorTotal = Abs(mdblSaldoInicial) + CDbl(.ListItems(i).ListSubItems(1)) - CDbl(.ListItems(i).ListSubItems(2))
                    lbl_Natureza.Caption = "CR"
                End If
                
                If dblValorTotal < 0 Then
                    .ListItems(i).ListSubItems(4) = "Débito"
                    .ListItems(i).ListSubItems(3) = gstrConvVrDoSql(dblValorTotal * -1)
                Else
                    .ListItems(i).ListSubItems(4) = "Crédito"
                    .ListItems(i).ListSubItems(3) = gstrConvVrDoSql(dblValorTotal)
                End If
                dblvalorAcumulado = dblValorTotal
            Else
                dblValorTotal = dblValorTotal + CDbl(.ListItems(i).ListSubItems(1)) - CDbl(.ListItems(i).ListSubItems(2))
                'dblvalorAcumulado = dblvalorAcumulado + dblvalorTotal
                If dblValorTotal < 0 Then
                    .ListItems(i).ListSubItems(4) = "Débito"
                    .ListItems(i).ListSubItems(3) = gstrConvVrDoSql(dblValorTotal * -1)
                Else
                    .ListItems(i).ListSubItems(4) = "Crédito"
                    .ListItems(i).ListSubItems(3) = gstrConvVrDoSql(dblValorTotal)
                End If
            End If
            
            .ListItems(i).ListSubItems(1) = gstrConvVrDoSql(.ListItems(i).ListSubItems(1))
            .ListItems(i).ListSubItems(2) = gstrConvVrDoSql(.ListItems(i).ListSubItems(2))
            
            
            
        End With
    Next
    If lvw_Resumo.ListItems.Count >= 12 Then
        txt_SaldoAtual = lvw_Resumo.ListItems(12).ListSubItems(3)
    End If
    
    If mdblSaldoInicial < 0 Then
       txt_dblSaldoInicial = gstrConvVrDoSql(Val(gstrConvVrParaSql(txt_dblSaldoInicial)) * (-1))
    End If
    
End Sub
Private Function blnVerificaMovi() As Boolean
    Dim strSQL As String
    Dim adoTemp As ADODB.Recordset
    
    'chkblnFinanceira
    'chkbytDisponibilidadeDeCaixa
    'chkblnAnalitica
    blnVerificaMovi = False
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "Count(*) IntNumero "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLancamentoContabil & " LC, "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrProcessoPagamento & " PG "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "LC.INTPROCESSO = PG.PKID AND "
    strSQL = strSQL & "LC.Intconta = PC.pkid AND "
    strSQL = strSQL & "Not PG.INTLANCAMENTOCONTABIL is null AND "
    strSQL = strSQL & "PC.Pkid = " & txtPKId
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoTemp) Then
        If adoTemp!INTNUMERO >= 1 Then
            blnVerificaMovi = True
        End If
    End If
        
    


End Function

Private Function CalculaSaldoInicial()
   Dim strSQL              As String
   Dim i                   As Integer
   Dim intPosicao          As Integer
   Dim adoResultado        As New ADODB.Recordset
   Dim mstrCodigoSemPontos As String
   Dim mstrMaiorHierarquia As String
    
   mstrCodigoSemPontos = gstrValorSemMascara(tdb_Lista.Columns(1).Value)
   mstrMaiorHierarquia = mstrCodigoSemPontos
    
    intPosicao = 0
    
    For i = 1 To Len(gstrMascaraContaContabil)
        If Mid(gstrMascaraContaContabil, i, 1) <> "0" Then
           If Val(Mid(mstrMaiorHierarquia, intPosicao + 1, Len(mstrMaiorHierarquia))) = 0 Then
              mstrMaiorHierarquia = Mid(mstrMaiorHierarquia, 1, intPosicao)
              Exit For
           End If
        Else
           intPosicao = intPosicao + 1
        End If
    Next
   
   strSQL = "SELECT SUM(TMP.DBLVALOR) DBLVALOR FROM (SELECT " & gstrISNULL("SUM(PCS.DBLVALOR)", "0") & " DBLVALOR "
   strSQL = strSQL & " FROM " & gstrPlanoConta & " PC ," & gstrPlanoContaSaldo & " PCS "
   strSQL = strSQL & " WHERE PCS.blnNaturezaDaConta = 0 AND "
   strSQL = strSQL & " PCS.intPlanoConta = PC.PKID AND "
   strSQL = strSQL & " PCS.intExercicio = " & gintExercicio & " AND "
   
'   If (bytDBType = EDatabases.Oracle) Then
'        strSql = strSql & " TO_NUMBER(TO_CHAR(PP.dtmData , 'YYYY')) = " & CStr(gintExercicio) & " AND "
'   Else
'        strSql = strSql & " YEAR(PP.dtmData) = " & CStr(gintExercicio) & " AND "
'   End If
    
   'strSQL = strSQL & " PC.Pkid = " & txtPKId
   strSQL = strSQL & " " & strSUBSTRING & " (PC.strContaContabil,1," + CStr(Len(mstrMaiorHierarquia)) & ") = '" & mstrMaiorHierarquia & "'"
   
   
   strSQL = strSQL & "UNION ALL SELECT " & gstrISNULL("SUM(PCS.DBLVALOR)", "0") & " * (-1) DBLVALOR "
   strSQL = strSQL & " FROM " & gstrPlanoConta & " PC ," & gstrPlanoContaSaldo & " PCS "
   strSQL = strSQL & " WHERE PCS.blnNaturezaDaConta = 1 AND "
   strSQL = strSQL & " PCS.intPlanoConta = PC.PKID AND "
   strSQL = strSQL & " PCS.intExercicio = " & gintExercicio & " AND "
   
'   If (bytDBType = EDatabases.Oracle) Then
'        strSql = strSql & " TO_NUMBER(TO_CHAR(PP.dtmData , 'YYYY')) = " & CStr(gintExercicio) & " AND "
'   Else
'        strSql = strSql & " YEAR(PP.dtmData) = " & CStr(gintExercicio) & " AND "
'   End If
'
   strSQL = strSQL & " " & strSUBSTRING & " (PC.strContaContabil,1," + CStr(Len(mstrMaiorHierarquia)) & ") = '" & mstrMaiorHierarquia & "'"
   strSQL = strSQL & ")TMP"
   
   Set gobjBanco = New clsBanco

   If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txt_dblSaldoInicial = gstrConvVrDoSql(!dblValor)
                txtdblSaldoDaConta = gstrConvVrDoSql(Abs(!dblValor))
            End If
        End With
    End If
End Function

Private Sub cbo_DescricaoContaCredito_Click()
    cbo_intContaCredito.ListIndex = gintIndiceCBO(cbo_intContaCredito, _
                            gstrItemData(cbo_DescricaoContaCredito))
End Sub

Private Sub cbo_intContaCredito_Click()
    cbo_DescricaoContaCredito.ListIndex = gintIndiceCBO(cbo_DescricaoContaCredito, _
                                   gstrItemData(cbo_intContaCredito))

End Sub

Private Sub cbo_intContaCredito_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaMovimento, 0
End Sub

Private Sub cbo_intContaCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_DescricaoContaCredito_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaMovimento, 0
End Sub

Private Sub cbo_DescricaoContaCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_DescricaoContaDebito_Click()
    cbo_intContaDebito.ListIndex = gintIndiceCBO(cbo_intContaDebito, _
                            gstrItemData(cbo_DescricaoContaDebito))
End Sub

Private Sub cbo_intContaDebito_Click()
    cbo_DescricaoContaDebito.ListIndex = gintIndiceCBO(cbo_DescricaoContaDebito, _
                                   gstrItemData(cbo_intContaDebito))

End Sub

Private Sub cbo_intContaDebito_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaMovimento, 1
End Sub

Private Sub cbo_intContaDebito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_DescricaoContaDebito_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaMovimento, 1
End Sub

Private Sub cbo_DescricaoContaDebito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub opt_EconomicoDebito_Click()
    PreencheComboDebito
    If opt_EconomicoDebito.Value = True Then
        chk_ContaGrupoDebito.Enabled = True
    Else
        chk_ContaGrupoDebito.Value = 0
        chk_ContaGrupoDebito.Enabled = False
    End If

End Sub

Private Sub opt_FinanceiroDebito_Click()
    PreencheComboDebito
    If opt_FinanceiroDebito.Value = True Then
        chk_ContaGrupoDebito.Enabled = True
    Else
        chk_ContaGrupoDebito.Value = 0
        chk_ContaGrupoDebito.Enabled = False
    End If
End Sub

Private Sub chk_ContaGrupoDebito_Click()
    PreencheComboDebito
End Sub

Public Function LeFonteRecurso(Optional intPkid As Long, Optional strCodigo As String) As String
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = ""
    
    If intPkid > 0 Then
        strSQL = "select strCodigo intNumero From " & gstrFonteRecurso & " Where intExercicio = " & gintExercicio & " And pkid = " & intPkid
    ElseIf Len(strCodigo) > 0 Then
        strSQL = "select Pkid intNumero From " & gstrFonteRecurso & " Where intExercicio = " & gintExercicio & " And strCodigo = '" & strCodigo & "'"
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            LeFonteRecurso = gstrENulo(adoResultado!INTNUMERO)
        Else
            LeFonteRecurso = ""
        End If
    End If
End Function

Private Function blnAlteraPlanoContaSaldo(intPkid As Long) As Boolean
    Dim strSQL As String
    
    blnAlteraPlanoContaSaldo = False
    
    strSQL = ""
    strSQL = "Update " & gstrPlanoContaSaldo & " Set intfonteRecurso = "
    
    If chkbytDisponibilidadeDeCaixa.Value = vbChecked Then
        strSQL = strSQL & dbc_intFonteRecurso.BoundText
    Else
        strSQL = strSQL & "Null"
    End If
    
    strSQL = strSQL & " Where intPlanoConta = " & intPkid
    
    
    
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSQL) Then
        blnAlteraPlanoContaSaldo = True
    End If
    
End Function

Public Function RetornaFonteRecurso(intPkid As Long) As String
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = ""
    
    strSQL = "select intFonteRecurso From " & gstrPlanoContaSaldo & " Where intPlanoConta = " & intPkid
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            RetornaFonteRecurso = gstrENulo(adoResultado!INTFONTERECURSO)
        Else
            RetornaFonteRecurso = "0"
        End If
    End If
End Function

Private Sub txt_intModalidade_GotFocus()
    MarcaCampo txt_intModalidade
End Sub

Private Sub txt_intModalidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intModalidade
End Sub

Private Sub dbcintModalidade_Click(Area As Integer)
    DropDownDataCombo dbcintModalidade, Me, Area
End Sub

Private Sub dbcintModalidade_GotFocus()
    MarcaCampo dbcintModalidade
End Sub

Private Sub dbcintModalidade_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintModalidade, Me, , KeyCode, Shift
End Sub

Private Sub dbcintModalidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintModalidade
End Sub

Private Sub txt_intConvenio_GotFocus()
    MarcaCampo txt_intConvenio
End Sub

Private Sub txt_intConvenio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intConvenio
End Sub
