VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadDividaAtivaManual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inscrição de Dívida Ativa Manual"
   ClientHeight    =   9135
   ClientLeft      =   1230
   ClientTop       =   1845
   ClientWidth     =   11205
   HelpContextID   =   5
   Icon            =   "CadDividaAtivaManual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2580
      TabIndex        =   88
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   7245
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   12779
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dívida Ativa"
      TabPicture(0)   =   "CadDividaAtivaManual.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_Parcelas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_DomicilioFiscal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Notificação"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_PrescricaoDoDebito"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Fra_Titulo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Observação"
      TabPicture(1)   =   "CadDividaAtivaManual.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_strindexador"
      Tab(1).Control(1)=   "lbl_dblvlindexador"
      Tab(1).Control(2)=   "fra_Historico"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(4)=   "txtstrindexador"
      Tab(1).Control(5)=   "txtdblvlindexador"
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame2 
         Caption         =   "Valores"
         Height          =   1065
         Left            =   180
         TabIndex        =   91
         Top             =   4860
         Width           =   10755
         Begin VB.TextBox txtdblValorTaxas 
            Alignment       =   1  'Right Justify
            DataField       =   "dblValorTaxas"
            Height          =   285
            Left            =   3540
            TabIndex        =   94
            Top             =   435
            Width           =   1545
         End
         Begin VB.TextBox txtdblValorImposto 
            Alignment       =   1  'Right Justify
            DataField       =   "dblValorImposto"
            Height          =   285
            Left            =   1320
            TabIndex        =   92
            Top             =   435
            Width           =   1545
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Receitas 
            Height          =   855
            Left            =   5250
            TabIndex        =   96
            Top             =   180
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   1508
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
            Columns(1).Caption=   "Receita"
            Columns(1).DataField=   "strdescricao"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Sigla"
            Columns(2).DataField=   "Strsigla"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=6297"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6218"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
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
            AllowUpdate     =   0   'False
            DefColWidth     =   0
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
            _StyleDefs(16)  =   ":id=8,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(45)  =   "Named:id=33:Normal"
            _StyleDefs(46)  =   ":id=33,.parent=0"
            _StyleDefs(47)  =   "Named:id=34:Heading"
            _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(49)  =   ":id=34,.wraptext=-1"
            _StyleDefs(50)  =   "Named:id=35:Footing"
            _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   "Named:id=36:Selected"
            _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(54)  =   "Named:id=37:Caption"
            _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(56)  =   "Named:id=38:HighlightRow"
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=39:EvenRow"
            _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(60)  =   "Named:id=40:OddRow"
            _StyleDefs(61)  =   ":id=40,.parent=33"
            _StyleDefs(62)  =   "Named:id=41:RecordSelector"
            _StyleDefs(63)  =   ":id=41,.parent=34"
            _StyleDefs(64)  =   "Named:id=42:FilterBar"
            _StyleDefs(65)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Taxas"
            Height          =   195
            Left            =   2940
            TabIndex        =   95
            Top             =   480
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Impostos"
            Height          =   195
            Left            =   600
            TabIndex        =   93
            Top             =   480
            Width           =   630
         End
      End
      Begin VB.TextBox txtdblvlindexador 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -71520
         MaxLength       =   25
         TabIndex        =   84
         Top             =   1500
         Width           =   1605
      End
      Begin VB.TextBox txtstrindexador 
         Height          =   285
         Left            =   -73950
         MaxLength       =   20
         TabIndex        =   82
         Top             =   1500
         Width           =   945
      End
      Begin VB.Frame Frame1 
         Height          =   1035
         Left            =   -74820
         TabIndex        =   62
         Top             =   330
         Width           =   10755
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   4680
            Picture         =   "CadDividaAtivaManual.frx":107A
            Style           =   1  'Graphical
            TabIndex        =   90
            TabStop         =   0   'False
            Tag             =   "617"
            ToolTipText     =   "Ativa Cadastro de Editais"
            Top             =   150
            Width           =   360
         End
         Begin VB.TextBox txtstrAviso2 
            Height          =   285
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   72
            Top             =   570
            Width           =   1545
         End
         Begin VB.TextBox txtdtmdtinscricao2 
            Height          =   285
            Left            =   4680
            TabIndex        =   74
            Top             =   570
            Width           =   1005
         End
         Begin VB.TextBox txtintExercicio2 
            Height          =   285
            Left            =   10140
            MaxLength       =   8
            TabIndex        =   70
            Top             =   150
            Width           =   495
         End
         Begin VB.TextBox txtintlivro2 
            Height          =   285
            Left            =   9690
            MaxLength       =   8
            TabIndex        =   80
            Top             =   570
            Width           =   945
         End
         Begin VB.TextBox txtintfolha2 
            Height          =   285
            Left            =   8460
            MaxLength       =   4
            TabIndex        =   78
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtstrIncricao2 
            Height          =   285
            Left            =   8010
            MaxLength       =   20
            TabIndex        =   68
            Top             =   150
            Width           =   1305
         End
         Begin VB.TextBox txtcadastro2 
            Height          =   285
            Left            =   5835
            TabIndex        =   66
            Top             =   150
            Width           =   1425
         End
         Begin VB.TextBox txtintcertidao2 
            Height          =   285
            Left            =   6420
            TabIndex        =   76
            Top             =   570
            Width           =   1365
         End
         Begin MSDataListLib.DataCombo dbc_intReceita2 
            Height          =   315
            Left            =   1710
            TabIndex        =   64
            Top             =   150
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_aviso2 
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Left            =   1230
            TabIndex        =   71
            Top             =   660
            Width           =   390
         End
         Begin VB.Label lbl_inscricao2 
            AutoSize        =   -1  'True
            Caption         =   "Data de Inscrição"
            Height          =   195
            Left            =   3360
            TabIndex        =   73
            Top             =   660
            Width           =   1260
         End
         Begin VB.Label lbl_exercicio2 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   9390
            TabIndex        =   69
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lbl_livro2 
            AutoSize        =   -1  'True
            Caption         =   "Livro"
            Height          =   195
            Left            =   9270
            TabIndex        =   79
            Top             =   660
            Width           =   345
         End
         Begin VB.Label lbl_folha2 
            AutoSize        =   -1  'True
            Caption         =   "Folha"
            Height          =   195
            Left            =   7980
            TabIndex        =   77
            Top             =   660
            Width           =   390
         End
         Begin VB.Label lbl_strInscricao2 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição"
            Height          =   195
            Left            =   7335
            TabIndex        =   67
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lbl_cadastro2 
            AutoSize        =   -1  'True
            Caption         =   "Cadastro"
            Height          =   195
            Left            =   5130
            TabIndex        =   65
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lbl_certidao2 
            AutoSize        =   -1  'True
            Caption         =   "Certidão"
            Height          =   195
            Left            =   5760
            TabIndex        =   75
            Top             =   660
            Width           =   585
         End
         Begin VB.Label lbl_compreceita2 
            AutoSize        =   -1  'True
            Caption         =   "Composição da receita"
            Height          =   195
            Left            =   60
            TabIndex        =   63
            Top             =   240
            Width           =   1620
         End
      End
      Begin VB.Frame Fra_Titulo 
         Height          =   1245
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Width           =   10755
         Begin VB.CheckBox chk_InscricaoAcordo 
            Caption         =   "Exibir inscrições em acordo"
            Height          =   195
            Left            =   4680
            TabIndex        =   21
            Top             =   930
            Width           =   2865
         End
         Begin VB.TextBox txtintQtdCertidaoUltFolha 
            Height          =   285
            Left            =   10380
            MaxLength       =   8
            TabIndex        =   97
            Top             =   900
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmd_TabelaComposicaoDaReceita 
            Height          =   315
            Left            =   4680
            Picture         =   "CadDividaAtivaManual.frx":1404
            Style           =   1  'Graphical
            TabIndex        =   89
            TabStop         =   0   'False
            Tag             =   "617"
            ToolTipText     =   "Ativa Cadastro de Composição Da Receita"
            Top             =   150
            Width           =   360
         End
         Begin VB.CheckBox chk_Atualizacao 
            Caption         =   "Exibir valores atualizados"
            Height          =   195
            Left            =   1710
            TabIndex        =   20
            Top             =   930
            Width           =   2085
         End
         Begin VB.TextBox txtintCertidao 
            Height          =   285
            Left            =   6420
            TabIndex        =   15
            Top             =   570
            Width           =   1365
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Left            =   5835
            TabIndex        =   5
            Top             =   150
            Width           =   1425
         End
         Begin VB.TextBox txtintFolha 
            Height          =   285
            Left            =   8460
            MaxLength       =   4
            TabIndex        =   17
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtintLivro 
            Height          =   285
            Left            =   9690
            MaxLength       =   8
            TabIndex        =   19
            Top             =   570
            Width           =   945
         End
         Begin VB.TextBox txtintExercicio 
            Height          =   285
            Left            =   10140
            MaxLength       =   8
            TabIndex        =   9
            Top             =   150
            Width           =   495
         End
         Begin VB.TextBox txtdtmdtinscricao 
            Height          =   285
            Left            =   4680
            TabIndex        =   13
            Top             =   570
            Width           =   1005
         End
         Begin VB.TextBox txtstrAviso 
            Height          =   285
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   11
            Top             =   570
            Width           =   1545
         End
         Begin MSDataListLib.DataCombo dbc_intReceita 
            Height          =   315
            Left            =   1710
            TabIndex        =   3
            Top             =   150
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskstrInscricao 
            Height          =   285
            Left            =   8010
            TabIndex        =   7
            Top             =   150
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin VB.Label lbl_compreceita 
            AutoSize        =   -1  'True
            Caption         =   "Composição da receita"
            Height          =   195
            Left            =   60
            TabIndex        =   2
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label lbl_certidao 
            AutoSize        =   -1  'True
            Caption         =   "Certidão"
            Height          =   195
            Left            =   5760
            TabIndex        =   14
            Top             =   660
            Width           =   585
         End
         Begin VB.Label lbl_cadastro 
            AutoSize        =   -1  'True
            Caption         =   "Cadastro"
            Height          =   195
            Left            =   5130
            TabIndex        =   4
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lbl_strInscricao 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição"
            Height          =   195
            Left            =   7335
            TabIndex        =   6
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lbl_folha 
            AutoSize        =   -1  'True
            Caption         =   "Folha"
            Height          =   195
            Left            =   7980
            TabIndex        =   16
            Top             =   660
            Width           =   390
         End
         Begin VB.Label lbl_livro 
            AutoSize        =   -1  'True
            Caption         =   "Livro"
            Height          =   195
            Left            =   9270
            TabIndex        =   18
            Top             =   660
            Width           =   345
         End
         Begin VB.Label lbl_exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   9390
            TabIndex        =   8
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lbl_inscricao 
            AutoSize        =   -1  'True
            Caption         =   "Data de Inscrição"
            Height          =   195
            Left            =   3360
            TabIndex        =   12
            Top             =   660
            Width           =   1260
         End
         Begin VB.Label lbl_aviso 
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Left            =   1230
            TabIndex        =   10
            Top             =   660
            Width           =   390
         End
      End
      Begin VB.Frame fra_Historico 
         Caption         =   "Histórico"
         Height          =   3795
         Left            =   -74820
         TabIndex        =   85
         Top             =   1980
         Width           =   10755
         Begin VB.TextBox txtHistorico 
            Height          =   3345
            Left            =   120
            MaxLength       =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   86
            Top             =   270
            Width           =   10485
         End
      End
      Begin VB.Frame fra_PrescricaoDoDebito 
         Caption         =   "Contribuinte"
         Height          =   885
         Left            =   180
         TabIndex        =   22
         Top             =   1560
         Width           =   10755
         Begin VB.TextBox txtstridentidade 
            Height          =   285
            Left            =   7380
            TabIndex        =   26
            Top             =   180
            Width           =   1155
         End
         Begin VB.TextBox txtstrcnpjcpf 
            Height          =   285
            Left            =   9480
            TabIndex        =   28
            Top             =   180
            Width           =   1155
         End
         Begin VB.TextBox txtstrnomeproprietario 
            Height          =   285
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   24
            Top             =   180
            Width           =   5115
         End
         Begin VB.TextBox txtstrpromissario 
            Height          =   285
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   30
            Top             =   540
            Width           =   8625
         End
         Begin VB.Label lbl_identidade 
            AutoSize        =   -1  'True
            Caption         =   "Identidade"
            Height          =   195
            Left            =   6600
            TabIndex        =   25
            Top             =   270
            Width           =   750
         End
         Begin VB.Label lbl_CNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ/CPF"
            Height          =   195
            Left            =   8670
            TabIndex        =   27
            Top             =   270
            Width           =   780
         End
         Begin VB.Label lbl_nome 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   450
            TabIndex        =   23
            Top             =   270
            Width           =   795
         End
         Begin VB.Label lbl_Prescricao 
            AutoSize        =   -1  'True
            Caption         =   "Promissário"
            Height          =   195
            Left            =   450
            TabIndex        =   29
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.Frame fra_Notificação 
         Caption         =   "Endereço de Notificação"
         Height          =   1215
         Left            =   180
         TabIndex        =   46
         Top             =   3630
         Width           =   10755
         Begin VB.TextBox txt_MunicipioN 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   58
            Top             =   840
            Width           =   6375
         End
         Begin VB.TextBox txt_UFN 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   8430
            MaxLength       =   2
            TabIndex        =   60
            Top             =   840
            Width           =   405
         End
         Begin VB.TextBox txtstrComplementoN 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   8430
            MaxLength       =   20
            TabIndex        =   52
            Top             =   210
            Width           =   1935
         End
         Begin VB.TextBox txtstrNumeroN 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   6570
            MaxLength       =   10
            TabIndex        =   50
            Top             =   210
            Width           =   1125
         End
         Begin VB.TextBox txt_CepN 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   8430
            MaxLength       =   20
            TabIndex        =   56
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txt_BairroN 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   54
            Top             =   540
            Width           =   6375
         End
         Begin VB.TextBox txt_LogradouroN 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   48
            Top             =   210
            Width           =   4725
         End
         Begin VB.Label lbl_strComplementoC 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   7890
            TabIndex        =   51
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbl_numeroC 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   6240
            TabIndex        =   49
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lbl_UFN 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   8130
            TabIndex        =   59
            Top             =   900
            Width           =   210
         End
         Begin VB.Label lbl_CepN 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   8040
            TabIndex        =   55
            Top             =   600
            Width           =   285
         End
         Begin VB.Label lbl_MunicipioN 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   555
            TabIndex        =   57
            Top             =   900
            Width           =   705
         End
         Begin VB.Label lbl_BairroN 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   855
            TabIndex        =   53
            Top             =   570
            Width           =   405
         End
         Begin VB.Label lbl_LogradouroN 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   195
            Left            =   570
            TabIndex        =   47
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.Frame fra_DomicilioFiscal 
         Caption         =   "Local"
         Height          =   1185
         Left            =   180
         TabIndex        =   31
         Top             =   2430
         Width           =   10755
         Begin VB.TextBox txtstrComplemento 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   8430
            MaxLength       =   20
            TabIndex        =   37
            Top             =   150
            Width           =   1935
         End
         Begin VB.TextBox txtstrNumero 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   6570
            MaxLength       =   10
            TabIndex        =   35
            Top             =   150
            Width           =   1125
         End
         Begin VB.TextBox txt_Logradouro 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   33
            Top             =   150
            Width           =   4725
         End
         Begin VB.TextBox txt_Bairro 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   39
            Top             =   480
            Width           =   6375
         End
         Begin VB.TextBox txt_Municipio 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   43
            Top             =   810
            Width           =   6375
         End
         Begin VB.TextBox txt_Cep 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   8430
            MaxLength       =   20
            TabIndex        =   41
            Top             =   480
            Width           =   1005
         End
         Begin VB.TextBox txt_UF 
            BackColor       =   &H8000000E&
            Height          =   285
            Left            =   8430
            MaxLength       =   2
            TabIndex        =   45
            Top             =   810
            Width           =   405
         End
         Begin VB.Label lbl_strComplemento 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   7920
            TabIndex        =   36
            Top             =   180
            Width           =   480
         End
         Begin VB.Label lbl_numero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   6240
            TabIndex        =   34
            Top             =   180
            Width           =   180
         End
         Begin VB.Label lbl_Logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   195
            Left            =   570
            TabIndex        =   32
            Top             =   180
            Width           =   690
         End
         Begin VB.Label lbl_Bairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   855
            TabIndex        =   38
            Top             =   510
            Width           =   405
         End
         Begin VB.Label lbl_Municipio 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   555
            TabIndex        =   42
            Top             =   840
            Width           =   705
         End
         Begin VB.Label lbl_Cep 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   8040
            TabIndex        =   40
            Top             =   540
            Width           =   285
         End
         Begin VB.Label lbl_UF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   8130
            TabIndex        =   44
            Top             =   840
            Width           =   210
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Parcelas 
         Height          =   1185
         Left            =   180
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   5970
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   2090
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   64
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   4
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nº"
         Columns(2).DataField=   "intNumeroParcela"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Vencimento"
         Columns(3).DataField=   "dtmVencimento"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Moeda"
         Columns(4).DataField=   "strMoeda"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Valor "
         Columns(5).DataField=   "dblValor"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Juros"
         Columns(6).DataField=   "dblJuros"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Multa"
         Columns(7).DataField=   "dblMulta"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Correc. Monetária"
         Columns(8).DataField=   "dblCorrecao"
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Total"
         Columns(9).DataField=   "dblTotal"
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "IntMoeda"
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   11
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=11"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=450"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=370"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1482"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1402"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1693"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1614"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2223"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2143"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=2275"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2196"
         Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=2117"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2037"
         Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(42)=   "Column(7).Width=2196"
         Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2117"
         Splits(0)._ColumnProps(45)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(48)=   "Column(8).Width=2540"
         Splits(0)._ColumnProps(49)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(8)._WidthInPix=2461"
         Splits(0)._ColumnProps(51)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(52)=   "Column(8)._ColStyle=2"
         Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(54)=   "Column(9).Width=3413"
         Splits(0)._ColumnProps(55)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(9)._WidthInPix=3334"
         Splits(0)._ColumnProps(57)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(58)=   "Column(9)._ColStyle=2"
         Splits(0)._ColumnProps(59)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(60)=   "Column(10).Width=132"
         Splits(0)._ColumnProps(61)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(62)=   "Column(10)._WidthInPix=53"
         Splits(0)._ColumnProps(63)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(64)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(65)=   "Column(10).Order=11"
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
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=74,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(81)  =   "Named:id=33:Normal"
         _StyleDefs(82)  =   ":id=33,.parent=0"
         _StyleDefs(83)  =   "Named:id=34:Heading"
         _StyleDefs(84)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(85)  =   ":id=34,.wraptext=-1"
         _StyleDefs(86)  =   "Named:id=35:Footing"
         _StyleDefs(87)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(88)  =   "Named:id=36:Selected"
         _StyleDefs(89)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(90)  =   "Named:id=37:Caption"
         _StyleDefs(91)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(92)  =   "Named:id=38:HighlightRow"
         _StyleDefs(93)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(94)  =   "Named:id=39:EvenRow"
         _StyleDefs(95)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(96)  =   "Named:id=40:OddRow"
         _StyleDefs(97)  =   ":id=40,.parent=33"
         _StyleDefs(98)  =   "Named:id=41:RecordSelector"
         _StyleDefs(99)  =   ":id=41,.parent=34"
         _StyleDefs(100) =   "Named:id=42:FilterBar"
         _StyleDefs(101) =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_dblvlindexador 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Indexador"
         Height          =   195
         Left            =   -72930
         TabIndex        =   83
         Top             =   1590
         Width           =   1335
      End
      Begin VB.Label lbl_strindexador 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   195
         Left            =   -74700
         TabIndex        =   81
         Top             =   1590
         Width           =   705
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1815
      Left            =   60
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   7290
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   3201
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "PkitdAlfa"
      Columns(0).DataField=   "intAlfa"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Comp. Receita"
      Columns(1).DataField=   "strComposicao"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Inscrição"
      Columns(2).DataField=   "strInscricao"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Exercício"
      Columns(3).DataField=   "intExercicio"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Aviso"
      Columns(4).DataField=   "strNumeroAviso"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Proprietário"
      Columns(5).DataField=   "strNomeProprietario"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=6350"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6271"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=3201"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3122"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1429"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1349"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2381"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2302"
      Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=8176"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=8096"
      Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
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
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(61)  =   "Named:id=33:Normal"
      _StyleDefs(62)  =   ":id=33,.parent=0"
      _StyleDefs(63)  =   "Named:id=34:Heading"
      _StyleDefs(64)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   ":id=34,.wraptext=-1"
      _StyleDefs(66)  =   "Named:id=35:Footing"
      _StyleDefs(67)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   "Named:id=36:Selected"
      _StyleDefs(69)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(70)  =   "Named:id=37:Caption"
      _StyleDefs(71)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(72)  =   "Named:id=38:HighlightRow"
      _StyleDefs(73)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(74)  =   "Named:id=39:EvenRow"
      _StyleDefs(75)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(76)  =   "Named:id=40:OddRow"
      _StyleDefs(77)  =   ":id=40,.parent=33"
      _StyleDefs(78)  =   "Named:id=41:RecordSelector"
      _StyleDefs(79)  =   ":id=41,.parent=34"
      _StyleDefs(80)  =   "Named:id=42:FilterBar"
      _StyleDefs(81)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadDividaAtivaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mobjAux                     As Object
Dim mblnClickOk                 As Boolean
Dim mblnSelecionou              As Boolean
Dim mblnPrimeiraVez             As Boolean
Dim vetImpostoTaxa()            As Double
Dim xadbParcelas                As XArrayDB
Dim blnClickGridParcelas        As Boolean

Private Sub chk_Atualizacao_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not tdb_Parcelas.EOF And Not tdb_Parcelas.BOF And Len(Trim(tdb_Parcelas.Columns(0).Value)) > 0 Then
        If chk_Atualizacao.Value Then
            If blnDadosAtualizacao Then
                chk_Atualizacao.Enabled = False
                AtualizaParcelas
            Else
                chk_Atualizacao.Value = 0
            End If
        Else
            chk_Atualizacao.Enabled = False
            PreencheGridParcela
            If xadbParcelas(0, 0) > 0 Then
                AdicionaRemoveTaixasImpostos
            Else
                txtdblValorImposto = ""
                txtdblValorTaxas.Text = ""
            End If
        End If
    Else
        If Val(Trim(txtPKId)) > 0 Then
            chk_Atualizacao.Enabled = False
            PreencheGridParcela
        End If
    End If
    chk_Atualizacao.Enabled = True
End Sub

Private Sub cmd_TabelaDeEdital_Click()

End Sub

Private Sub chk_InscricaoAcordo_Click()
    If Not tdb_Parcelas.EOF And Not tdb_Parcelas.BOF And Len(Trim(tdb_Parcelas.Columns(0).Value)) > 0 Then
        If chk_Atualizacao.Value Then
            If blnDadosAtualizacao Then
                AtualizaParcelas
            End If
        Else
            PreencheGridParcela
            If xadbParcelas(0, 0) > 0 Then
                AdicionaRemoveTaixasImpostos
            Else
                txtdblValorImposto = ""
                txtdblValorTaxas.Text = ""
            End If
        End If
    Else
        If Val(Trim(txtPKId)) > 0 Then
            PreencheGridParcela
        End If
    End If
End Sub

Private Sub cmd_TabelaComposicaoDaReceita_Click()
    ChamaFormCadastro frmCadComposicaoDaReceita, dbc_intReceita
End Sub

Private Sub dbc_intReceita_Change()
    If dbc_intReceita.MatchedWithList = True And dbc_intReceita.BoundText <> "" Then
        PreencherListaDeOpcoes dbc_intReceita2, dbc_intReceita.BoundText
        PreencheCadastro CLng(dbc_intReceita.BoundText)
    End If
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        txtstrindexador.SetFocus
    ElseIf PreviousTab = 1 Then
        dbc_intReceita.SetFocus
    End If
End Sub

Private Sub tdb_Parcelas_ColEdit(ByVal ColIndex As Integer)
    Dim intFor As Integer
    
    With tdb_Parcelas
        .Enabled = False
        .RowDividerColor = dbgLightGrayLine
        If Not .EOF And Not .BOF And Len(Trim(.Columns(0).Value)) > 0 Then
            If ColIndex = 1 Then
                For intFor = 0 To xadbParcelas.UpperBound(1)
                    If xadbParcelas(intFor, 0) = .Columns(0).Value Then
                        If xadbParcelas(intFor, 1) = -1 Then
                            xadbParcelas(intFor, 1) = 0
                            'Vamos atualizar Imposto / Taxa subtraindo
                            If chk_Atualizacao.Value Then
                                If CDbl(vetImpostoTaxa(0, 0)) > 0 Then
                                    txtdblValorImposto = CDbl(txtdblValorImposto) - (CDbl(xadbParcelas(intFor, 9)) * vetImpostoTaxa(0, 2))
                                Else
                                    txtdblValorImposto = 0
                                End If
                                txtdblValorImposto = gstrConvVrDoSql(txtdblValorImposto, , , True)
                                If CDbl(vetImpostoTaxa(0, 1)) > 0 Then
                                    txtdblValorTaxas = CDbl(txtdblValorTaxas) - (CDbl(xadbParcelas(intFor, 9)) * vetImpostoTaxa(0, 3))
                                Else
                                    txtdblValorTaxas = 0
                                End If
                                txtdblValorTaxas = gstrConvVrDoSql(txtdblValorTaxas, , , True)
                            Else
                                RemoveTaxasImpostos intFor
                                txtdblValorImposto = gstrConvVrDoSql(vetImpostoTaxa(0, 0), 2)
                                txtdblValorTaxas = gstrConvVrDoSql(vetImpostoTaxa(0, 1), 2)
                            End If
                            Exit For
                        Else
                            xadbParcelas(intFor, 1) = -1
                            'Vamos atualizar Imposto / Taxa somando
                            If chk_Atualizacao.Value Then
                                If CDbl(vetImpostoTaxa(0, 0)) > 0 Then
                                    txtdblValorImposto = CDbl(txtdblValorImposto) - (CDbl(xadbParcelas(intFor, 9)) * vetImpostoTaxa(0, 2))
                                Else
                                    txtdblValorImposto = 0
                                End If
                                txtdblValorImposto = gstrConvVrDoSql(txtdblValorImposto, , , True)
                                If CDbl(vetImpostoTaxa(0, 1)) > 0 Then
                                    txtdblValorTaxas = CDbl(txtdblValorTaxas) + (CDbl(xadbParcelas(intFor, 9)) * vetImpostoTaxa(0, 3))
                                Else
                                    txtdblValorTaxas = 0
                                End If
                                txtdblValorTaxas = gstrConvVrDoSql(txtdblValorTaxas, , , True)
                            Else
                                AdicionaTaxasImpostos intFor
                                txtdblValorImposto = gstrConvVrDoSql(vetImpostoTaxa(0, 0), 2)
                                txtdblValorTaxas = gstrConvVrDoSql(vetImpostoTaxa(0, 1), 2)
                            End If
                            Exit For
                        End If
                    End If
                Next
                DoEvents
            End If
        End If
        .Enabled = True
    End With
End Sub

Private Sub tdb_Parcelas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    gCorLinhaSelecionada tdb_Parcelas
End Sub


Private Sub txtHistorico_GotFocus()
    MarcaCampo txtHistorico
    tab_3DPasta.Tab = 1
End Sub

Private Sub txtHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtHistorico
End Sub

Private Sub txtintcertidao_Change()
    txtintcertidao2 = txtintCertidao
End Sub

Private Sub txtintExercicio_Change()
    txtintExercicio2 = txtintExercicio
End Sub

Private Sub txtintfolha_Change()
    txtintfolha2 = txtintFolha
End Sub

Private Sub txtintlivro_Change()
    txtintlivro2 = txtintLivro
End Sub

Private Sub txtstrAviso_Change()
    txtstrAviso2 = txtstrAviso
End Sub

Private Sub mskstrInscricao_Change()
    txtstrIncricao2 = mskstrInscricao
End Sub

Private Sub txtstrcnpjcpf_GotFocus()
    MarcaCampo txtstrcnpjcpf
End Sub

Private Sub txtstrcnpjcpf_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrcnpjcpf
End Sub

Private Sub txtstridentidade_GotFocus()
    MarcaCampo txtstridentidade
End Sub

Private Sub txtstrIdentidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstridentidade
End Sub

Private Sub txtstrindexador_GotFocus()
    MarcaCampo txtstrindexador
End Sub

Private Sub txtstrNumeroN_GotFocus()
    MarcaCampo txtstrNumeroN
End Sub

Private Sub txtstrNumeroN_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumeroN
End Sub
Private Sub txt_BairroN_GotFocus()
    MarcaCampo txt_BairroN
End Sub

Private Sub txt_BairroN_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_BairroN
End Sub

Private Sub txt_CepN_GotFocus()
    MarcaCampo txt_CepN
End Sub

Private Sub txt_CepN_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txt_CepN
End Sub

Private Sub txt_LogradouroN_GotFocus()
    MarcaCampo txt_LogradouroN
End Sub

Private Sub txt_LogradouroN_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_LogradouroN
End Sub

Private Sub txt_UFN_GotFocus()
    MarcaCampo txt_UFN
    tab_3DPasta.Tab = 0
End Sub

Private Sub txt_UFN_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "U", txt_UFN
End Sub

Private Sub txt_MunicipioN_GotFocus()
    MarcaCampo txt_MunicipioN
End Sub

Private Sub txt_MunicipioN_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_MunicipioN
End Sub

Private Sub txtstrComplementoN_GotFocus()
    MarcaCampo txtstrComplementoN
End Sub

Private Sub txtstrComplementoN_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplementoN
End Sub

Private Sub dbc_intReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intReceita, Me, Area
End Sub

Private Sub dbc_intReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intReceita, Me, , KeyCode, Shift
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnPrimeiraVez = True
    mblnClickOk = True
End Sub

Private Sub txt_Bairro_GotFocus()
    MarcaCampo txt_Bairro
End Sub

Private Sub txt_Bairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Bairro
End Sub

Private Sub txt_Cep_GotFocus()
    MarcaCampo txt_Cep
End Sub

Private Sub txt_Cep_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txt_Cep
End Sub

Private Sub txt_Cep_LostFocus()
    txt_Cep = gstrCEPFormatado(txt_Cep)
    CepLogradouro txt_Cep, txt_Logradouro, txt_Bairro, txt_Municipio, txt_UF, , , True, False, False, False, False
End Sub

Private Sub txt_CepN_LostFocus()
    txt_CepN = gstrCEPFormatado(txt_CepN)
    CepLogradouro txt_CepN, txt_LogradouroN, txt_BairroN, txt_MunicipioN, txt_UFN, , , True, False, False, False, False
End Sub

Private Sub txt_Logradouro_GotFocus()
    MarcaCampo txt_Logradouro
End Sub

Private Sub txt_Logradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Logradouro
End Sub

Private Sub txt_Municipio_GotFocus()
    MarcaCampo txt_Municipio
End Sub

Private Sub txt_Municipio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Municipio
End Sub

Private Sub txt_UF_GotFocus()
    MarcaCampo txt_UF
End Sub

Private Sub txt_UF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "U", txt_UF
End Sub

Private Sub txtcadastro_GotFocus()
    MarcaCampo txtcadastro
End Sub

Private Sub txtcadastro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtcadastro
End Sub

Private Sub txtdblvlindexador_GotFocus()
    MarcaCampo txtdblvlindexador
End Sub

Private Sub txtDblVlIndexador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblvlindexador
End Sub

Private Sub txtdblvlindexador_LostFocus()
    txtdblvlindexador = gstrConvVrDoSql(txtdblvlindexador, 6)
End Sub

Private Sub txtdtmdtinscricao_GotFocus()
    If txtdtmdtinscricao = "" Then
        txtdtmdtinscricao = gstrDataFormatada(Date)
    End If
    MarcaCampo txtdtmdtinscricao
End Sub

Private Sub txtdtmdtinscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdtinscricao
End Sub

Private Sub txtdtmdtinscricao_LostFocus()
    txtdtmdtinscricao = gstrDataFormatada(txtdtmdtinscricao)
    txtdtmdtinscricao2 = txtdtmdtinscricao
    chk_Atualizacao_MouseUp 1, 0, 825, 120
End Sub

Private Sub txtintCertidao_GotFocus()
    MarcaCampo txtintCertidao
End Sub

Private Sub txtintcertidao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtintCertidao
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintFolha_GotFocus()
    MarcaCampo txtintFolha
End Sub

Private Sub txtintfolha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintFolha
End Sub

Private Sub txtintLivro_GotFocus()
    MarcaCampo txtintLivro
End Sub

Private Sub txtintlivro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintLivro
End Sub

Private Sub txtstrAviso_GotFocus()
    MarcaCampo txtstrAviso
End Sub

Private Sub txtstrAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrAviso
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
End Sub

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub mskstrInscricao_GotFocus()
    MarcaCampo mskstrInscricao
End Sub

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricao
End Sub

Private Sub txtstrindexador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrindexador
End Sub

Private Sub txtstrNomeProprietario_GotFocus()
    MarcaCampo txtstrnomeproprietario
End Sub

Private Sub txtstrNomeProprietario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrnomeproprietario
End Sub

Private Sub txtstrNumero_GotFocus()
    MarcaCampo txtstrNumero
End Sub

Private Sub txtstrNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumero
End Sub

Private Sub txtstrPromissario_GotFocus()
    MarcaCampo txtstrpromissario
End Sub

Private Sub txtstrPromissario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrpromissario
End Sub

Private Sub dbc_intReceita_GotFocus()
    tab_3DPasta.Tab = 0
    MarcaCampo dbc_intReceita
End Sub

Private Sub dbc_intReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intReceita
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1295

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
    
    VerificaObjParaAplicar mobjAux
    TrocaCorObjeto dbc_intReceita2, True
    TrocaCorObjeto txtcadastro2, True
    TrocaCorObjeto txtstrIncricao2, True
    TrocaCorObjeto txtintExercicio2, True
    TrocaCorObjeto txtstrAviso2, True
    TrocaCorObjeto txtintfolha2, True
    TrocaCorObjeto txtintlivro2, True
    TrocaCorObjeto txtdtmdtinscricao2, True
    TrocaCorObjeto txtintcertidao2, True
    TrocaCorObjeto txtcadastro, True
    TrocaCorObjeto txtintCertidao, True
    TrocaCorObjeto txtintFolha, True
    TrocaCorObjeto txtintLivro, True
    
    TrocaCorObjeto txtdblValorImposto, True
    TrocaCorObjeto txtdblValorTaxas, True
   
    dbc_intReceita.Tag = strQueryComposicaoReceita & ";strDescricao"
    dbc_intReceita2.Tag = strQueryComposicaoReceita & ";strDescricao"
       
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Function strQueryComposicaoReceita()
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT PKId," & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " Ltrim(Rtrim(strDescricao)) as strDescricao "
    strSql = strSql & "FROM " & gstrComposicaoDaReceita & " "
    strSql = strSql & "WHERE bytDividaAtiva = 1 " ' And ""
    'strSQL = strSQL & "intUtilizacao <> 3 "
    strSql = strSql & "ORDER BY strDescricao"
    
    strQueryComposicaoReceita = strSql
    
End Function

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
   gOrdenaGrid tdb_Lista, ColIndex
   mblnPrimeiraVez = False
   mblnClickOk = False
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                If mblnClickOk Then
                    Limpa_Controles Me, True, True, True, True, True
                    gCorLinhaSelecionada tdb_Lista
                    mblnClickOk = False
                    mblnSelecionou = True
                    gCorLinhaSelecionada tdb_Lista
                    LimpaGrids
                    Set tdb_Receitas.DataSource = Nothing
                    tdb_Receitas.Refresh

                    txtPKId = .Columns(0).Value
                    PreencheCampos
                    PreencheGridParcela
                    PreencheGridReceita
                    If xadbParcelas(0, 0) > 0 Then
                        AdicionaRemoveTaixasImpostos
                    Else
                        txtdblValorImposto = ""
                        txtdblValorTaxas.Text = ""
                    End If
                    If mobjAux Is Nothing Then
                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(gstrImprimir) = UCase(strModoOperacao) Then

    ElseIf UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    ElseIf UCase(strModoOperacao) = gstrLocalizar Then
        LeDaTabelaParaObj "", tdb_Lista, strQuery(True)
    ElseIf UCase(strModoOperacao) = gstrRefresh Then
        LeDaTabelaParaObj "", tdb_Lista, strQuery(False)
    ElseIf UCase(gstrFechar) = UCase(strModoOperacao) Then
        Unload Me
    ElseIf UCase(gstrNovo) = UCase(strModoOperacao) Then
        Limpa_Controles Me, True, True, True, True, True
        LimpaGrids
        Set tdb_Receitas.DataSource = Nothing
        tdb_Receitas.Refresh
        ReDim vetImpostoTaxa(0, 3)
        mskstrInscricao.Mask = ""
        mskstrInscricao.Text = ""
        dbc_intReceita.SetFocus
    ElseIf UCase(gstrSalvar) = UCase(strModoOperacao) Then
        If blnDadosOk Then
            If gblnExclusaoGravacaoOk("SALVAR", "Deseja realmente Inscrever a inscrição " & mskstrInscricao.FormattedText & " em Dívida Ativa") Then
                
                If blnSalvarLancamentoDA Then
                    Limpa_Controles Me, True, True, True, True, True
                    LimpaGrids
                    Set tdb_Receitas.DataSource = Nothing
                    tdb_Receitas.Refresh

                    mskstrInscricao.Mask = ""
                    mskstrInscricao.Text = ""
                    dbc_intReceita.SetFocus
                Else
                    ExibeMensagem "Não foi possível gravar Inscrição de Dívida Ativa Manual."
                End If
            End If
        End If
    Else
        
    End If
End Sub

Private Sub PreencheCampos()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select "
    
    If bytDBType = Oracle Then
        strSql = strSql & "/*+ index(A) */ " 'Parâmetro adicional inserido para otimizar a consulta à pedido do DBA
    End If
    
    strSql = strSql & "LA.Pkid, "
    strSql = strSql & "LA.Intcomposicaodareceita, "
    strSql = strSql & "da.dblvalorimposto, "
    strSql = strSql & "da.dblvalortaxas, "
    strSql = strSql & "CR.INTUTILIZACAO, "
    strSql = strSql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
    strSql = strSql & "LA.strNumeroAviso, "
    strSql = strSql & "LA.intExercicio, "
    strSql = strSql & "LA.strnomeproprietario, "
    strSql = strSql & "PDA.Intlivro, "
    strSql = strSql & "pda.intfolha, "
    strSql = strSql & "pda.intcertidao, "
    strSql = strSql & "LA.strcnpjcpf, "
    strSql = strSql & "LA.stridentidade, "
    strSql = strSql & "LA.strlogradouro, "
    strSql = strSql & "LA.strnumero, "
    strSql = strSql & "LA.strcomplemento, "
    strSql = strSql & "LA.strbairro, "
    strSql = strSql & "LA.strmunicipio, "
    strSql = strSql & "LA.struf, "
    strSql = strSql & "LA.intcep, "
    strSql = strSql & "LA.strlogradouroc, "
    strSql = strSql & "LA.strnumeroc, "
    strSql = strSql & "LA.strcomplementoc, "
    strSql = strSql & "LA.strbairroc, "
    strSql = strSql & "LA.strmunicipioc, "
    strSql = strSql & "LA.strufc, "
    strSql = strSql & "LA.intcepc, "
    strSql = strSql & "LA.strpromissario, "
    strSql = strSql & "LA.strindexador, "
    strSql = strSql & "LA.dblvlindexador "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
    strSql = strSql & gstrParametroDividaAtiva & " PDA, "
    strSql = strSql & gstrDativa & " DA "
    strSql = strSql & "Where "
    strSql = strSql & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strSql = strSql & "CR.Pkid " & strOUTJSQLServer & "= PDA.Intcomposicaodareceita " & strOUTJOracle & " AND "
    strSql = strSql & "LA.pkid " & strOUTJSQLServer & "= DA.intLancamentoAlfa " & strOUTJOracle & " AND "
    strSql = strSql & "LA.pkid = " & Trim(txtPKId)
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                
                PreencherListaDeOpcoes dbc_intReceita, gstrENulo(!intComposicaoDaReceita)
                Select Case gstrENulo(!intUtilizacao)
                    Case 1
                        txtcadastro = "Imobiliário"
                        txtcadastro2 = "Imobiliário"
                    Case 2
                        txtcadastro = "Econômico"
                        txtcadastro2 = "Econômico"
                    Case 3
                        txtcadastro = "Dívida Ativa"
                        txtcadastro2 = "Dívida Ativa"
                    Case 4
                        txtcadastro = "Acordo"
                        txtcadastro2 = "Acordo"
                    Case 5
                        txtcadastro = "Preco Público"
                        txtcadastro2 = "Preco Público"
                    Case 6
                        txtcadastro = "ISS Construção"
                        txtcadastro2 = "ISS Construção"
                    Case Else
                        txtcadastro = ""
                        txtcadastro2 = ""
                End Select
                
                mskstrInscricao = gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao)
                txtintExercicio = gstrENulo(!intExercicio)
                txtstrAviso = gstrENulo(!strNumeroAviso)
                txtdtmdtinscricao = IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, gstrDataFormatada(Date))
                
                PreencheCertidaoLivroFolha !intComposicaoDaReceita
                
                PreencherListaDeOpcoes dbc_intReceita2, gstrENulo(!intComposicaoDaReceita)
                txtstrIncricao2 = gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao)
                txtintExercicio2 = gstrENulo(!intExercicio)
                txtstrAviso2 = gstrENulo(!strNumeroAviso)
                txtdtmdtinscricao2 = gstrDataFormatada(Date)
                txtstrnomeproprietario = gstrENulo(!strnomeproprietario)
                txtstrcnpjcpf = gstrENulo(!StrCnpjCpf)
                txtstridentidade = gstrENulo(!STRIDENTIDADE)
                txt_Logradouro = gstrENulo(!strLogradouro)
                txtstrNumero = gstrENulo(!strNumero)
                txtstrComplemento = gstrENulo(!STRCOMPLEMENTO)
                txt_Bairro = gstrENulo(!strBairro)
                txt_Municipio = gstrENulo(!STRMUNICIPIO)
                txt_UF = gstrENulo(!STRUF)
                txt_Cep = gstrCEPFormatado(gstrENulo(!INTCEP))
                
                'txtdblValorImposto = gstrENulo(!dblValorImposto)
                'txtdblValorTaxas = gstrENulo(!dblvalortaxas)
                
                txt_LogradouroN = gstrENulo(!strlogradouroc)
                txtstrNumeroN = gstrENulo(!strNumeroC)
                txtstrComplementoN = gstrENulo(!strComplementoC)
                txt_BairroN = gstrENulo(!strBairroC)
                txt_MunicipioN = gstrENulo(!strMunicipioC)
                txt_UFN = gstrENulo(!strUFC)
                txt_CepN = gstrCEPFormatado(gstrENulo(!intcepc))
                txtstrpromissario = gstrENulo(!strpromissario)
                txtstrindexador = gstrENulo(!Strindexador)
                txtdblvlindexador = gstrConvVrDoSql(gstrENulo(!dblvlIndexador), 6)
            End If
        End With
    End If
    
End Sub

Private Function strQuery(blnFiltro As Boolean) As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT DISTINCT "
    If bytDBType = Oracle Then
        strSql = strSql & "/*+ index(A) */ " 'Parâmetro adicional inserido para otimizar a consulta à pedido do DBA
    End If
    strSql = strSql & "Count (LP.Pkid), "
    strSql = strSql & "LA.Pkid AS intAlfa, "
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CR.intCodigo ") & strCONCAT & "' - '" & strCONCAT & " CR.Strdescricao AS strComposicao, "
    strSql = strSql & "LA.Intexercicio, "
    strSql = strSql & "LA.strNumeroAviso, "
    strSql = strSql & "LA.strInscricao, "
    strSql = strSql & "LA.Strnomeproprietario "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoPagamento & " LP, "
    strSql = strSql & gstrDativa & " A "
    strSql = strSql & "WHERE "
    strSql = strSql & "CR.PKID = LA.Intcomposicaodareceita AND "
    strSql = strSql & "La.pkid = LV.Intlancamentoalfa" & strOUTJOracle
    strSql = strSql & " AND LV.PKID " & strOUTJSQLServer & "= LP.INTLANCAMENTOVALOR " & strOUTJOracle
    strSql = strSql & " AND "
    strSql = strSql & " LV.BITPARCELAVALIDA = 1 AND "
    strSql = strSql & " LA.pkid " & strOUTJSQLServer & "= A.intLancamentoAlfa " & strOUTJOracle & " AND "
    strSql = strSql & "LV.IntlancamentoalfaDativa Is Null AND "
    strSql = strSql & "LV.dtmdtVencimento < " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date)) & " AND "
    strSql = strSql & "CR.bytDividaAtiva = 1 "
        
    If blnFiltro Then
        If dbc_intReceita.MatchedWithList = True Then strSql = strSql & " AND LA.intComposicaoDaReceita = " & dbc_intReceita.BoundText
        If Trim(txtcadastro) <> "" Then strSql = strSql & ""
        If Trim(mskstrInscricao) <> "" Then strSql = strSql & " AND LA.strInscricao ='" & (String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text)) & "'"
        If Trim(txtintExercicio) <> "" Then strSql = strSql & " AND LA.intExercicio = " & txtintExercicio.Text
        'If Trim(txtstrAviso) <> "" Then strSql = strSql & " AND LA.strNumeroAviso = " & txtstrAviso.Text
        If Trim(txtstrAviso) <> "" Then strSql = strSql & " AND LA.strNumeroAviso = '" & (String(gintLenNumAviso = 10 - Len(Trim(txtstrAviso)), "0") & Trim(txtstrAviso)) & "'"
        If Trim(txtdtmdtinscricao) <> "" Then strSql = strSql & " AND A.dtmdtinscricao = " & gstrConvDtParaSql(txtdtmdtinscricao.Text)
        If Trim(txtintCertidao) <> "" Then strSql = strSql & " AND A.intcertidao = " & txtintCertidao.Text
        If Trim(txtintFolha) <> "" Then strSql = strSql & " AND A.intfolha = " & txtintFolha.Text
        If Trim(txtintLivro) <> "" Then strSql = strSql & " AND A.intlivro = " & txtintLivro.Text
        If Trim(txtstrnomeproprietario) <> "" Then strSql = strSql & " AND UPPER(LA.strnomeproprietario) Like '" & UCase(txtstrnomeproprietario.Text) & "%'"
        If Trim(txtstrpromissario) <> "" Then strSql = strSql & " AND UPPER(LA.strpromissario) Like '" & UCase(txtstrpromissario.Text) & "%'"
    End If
    
    strSql = strSql & " GROUP BY "
    strSql = strSql & "LV.Pkid, "
    strSql = strSql & "LA.Pkid, "
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CR.intCodigo ") & strCONCAT & "' - '" & strCONCAT & " CR.Strdescricao, "
    strSql = strSql & "LA.Intexercicio, "
    strSql = strSql & "LA.strNumeroAviso, "
    strSql = strSql & "LA.strInscricao, "
    strSql = strSql & "LA.Strnomeproprietario "
    
    strSql = strSql & " Having Count(LP.pkid) = 0 "
    
    strSql = strSql & " ORDER BY strComposicao, strInscricao, LA.Intexercicio ASC"
    
    strQuery = strSql
    
End Function

Private Function strQueryParcela() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "LV.Pkid, "
    strSql = strSql & "LV.Intparcela as intNumeroParcela, "
    strSql = strSql & "M.Strabreviatura as strMoeda, "
    strSql = strSql & "LV.Dtmdtvencimento as dtmVencimento, "
    strSql = strSql & gstrISNULL("LV.Dblvalor", 0) & " as dblValor, "
    strSql = strSql & gstrISNULL("LV.DBLVALOR", 0) & " as dblTotal, "
    strSql = strSql & "M.Pkid as intMoeda, "
    strSql = strSql & "COUNT(LP.Pkid) "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoPagamento & " LP, "
    strSql = strSql & gstrMoedas & " M "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.INTLANCAMENTOALFA AND "
    strSql = strSql & "LV.INTLANCAMENTOALFADATIVA IS NULL AND "
        
    If chk_InscricaoAcordo.Value = vbUnchecked Then
        strSql = strSql & "LV.INTLANCAMENTOALFAACORDO IS NULL AND "
    End If
    
    strSql = strSql & "LV.bitParcelaValida = 1 AND "
    strSql = strSql & "LV.Pkid" & strOUTJSQLServer & "= LP.Intlancamentovalor" & strOUTJOracle & " AND "
    strSql = strSql & "M.Pkid" & strOUTJOracle & "=" & strOUTJSQLServer & "LV.Intmoeda AND "
    strSql = strSql & "LV.Dtmdtvencimento < " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date)) & " AND "
    strSql = strSql & "LA.Pkid = " & Trim(txtPKId)
    strSql = strSql & " GROUP BY LV.pkId, LV.intParcela, M.strAbreviatura, M.pkid, LV.dtmDtVencimento, "
    strSql = strSql & gstrISNULL("LV.Dblvalor", 0) & ", "
    strSql = strSql & gstrISNULL("LV.DBLVALOR", 0)
    strSql = strSql & " HAVING COUNT(LP.Pkid) = 0 "
    strSql = strSql & " Order By LV.Intparcela"
    
    strQueryParcela = strSql
    
End Function

Private Sub PreencheCadastro(lngPkid As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select * From " & gstrComposicaoDaReceita & " Where pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Select Case gstrENulo(adoResultado!intUtilizacao)
                Case 1
                    txtcadastro = "Imobiliário"
                    txtcadastro2 = "Imobiliário"
                Case 2
                    txtcadastro = "Econômico"
                    txtcadastro2 = "Econômico"
                Case 3
                    txtcadastro = "Dívida Ativa"
                    txtcadastro2 = "Dívida Ativa"
                Case 4
                    txtcadastro = "Acordo"
                    txtcadastro2 = "Acordo"
                Case 5
                    txtcadastro = "Preco Público"
                    txtcadastro2 = "Preco Público"
                Case 6
                    txtcadastro = "ISS Construção"
                    txtcadastro2 = "ISS Construção"
                Case Else
                    txtcadastro = ""
                    txtcadastro2 = ""
            End Select
            VerificaMascaraInscricao CInt(gstrENulo(adoResultado!intUtilizacao))
        End If
    End If
End Sub

Private Sub LimpaGrids()

    Set xadbParcelas = New XArrayDB
    xadbParcelas.Clear
    xadbParcelas.ReDim 0, 0, 0, 10
    
    Set tdb_Parcelas.Array = xadbParcelas
    tdb_Parcelas.ReBind
    tdb_Parcelas.Refresh
    
End Sub

Private Sub PreencheGridParcela()
    Dim adoResultado    As ADODB.Recordset
    Dim varAux          As Variant
    Dim intPosition     As Integer
    
    LimpaGrids
    intPosition = 0
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strQueryParcela, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                xadbParcelas.ReDim 0, .RecordCount - 1, 0, 10
                Do While Not .EOF
                        varAux = Space$(0) & !Pkid                                  'Pkid da TbllancamentoValor
                        xadbParcelas(intPosition, 0) = varAux
                        
                        xadbParcelas(intPosition, 1) = -1                           'Traz todas as parcelas checadas
                        
                        varAux = Space$(0) & !intNumeroParcela                      'Número das parcelas
                        xadbParcelas(intPosition, 2) = varAux
                        
                        varAux = Space$(0) & !strMoeda                              'Abreviatura da Moeda
                        xadbParcelas(intPosition, 3) = varAux
                        
                        varAux = Space$(0) & gstrDataFormatada(!dtmVencimento)      'Vencimento da parcela
                        xadbParcelas(intPosition, 4) = varAux
                        
                        varAux = Space$(0) & gstrConvVrDoSql(!dblValor, 2, , True)  'Valor da parcela
                        xadbParcelas(intPosition, 5) = varAux
                        
                        varAux = "0,00"                                             'Não tem o valor de Juros
                        xadbParcelas(intPosition, 6) = varAux
                        
                        varAux = "0,00"                                             'Não tem o valor de Multa
                        xadbParcelas(intPosition, 7) = varAux
                        
                        varAux = "0,00"                                             'Não tem o valor de Correção
                        xadbParcelas(intPosition, 8) = varAux
                        
                        varAux = Space$(0) & gstrConvVrDoSql(!dblValor, 2, , True)  'Valor Total
                        xadbParcelas(intPosition, 9) = varAux
                        
                        varAux = Space$(0) & gstrENulo(!intMoeda)                   'Pkid da tblMoeda
                        xadbParcelas(intPosition, 10) = varAux

                        intPosition = intPosition + 1
                        
                    .MoveNext
                Loop
                
                Set tdb_Parcelas.Array = xadbParcelas
                tdb_Parcelas.ReBind
                tdb_Parcelas.Refresh
                
            End If
        End With
    End If
    
End Sub

Private Sub AtualizaParcelas()
    Dim intFor          As Integer
    Dim adoParcelas     As ADODB.Recordset
    Dim adoResultado    As ADODB.Recordset
    Dim xadbAtualizadas As XArrayDB
    Dim varAux          As Variant
    Dim strSql          As String
    Dim blnAtualiza     As Boolean
    Dim dblTotal        As Double
    
    Set xadbAtualizadas = New XArrayDB
    xadbAtualizadas.Clear
    xadbAtualizadas.ReDim 0, 0, 0, 10
    blnAtualiza = False
    intFor = 0
    If gobjBanco.CriaADO(strQueryParcela, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            Do While Not adoResultado.EOF
            
                If Val(xadbParcelas(adoResultado.AbsolutePosition - 1, 0)) > 0 And Val(xadbParcelas(adoResultado.AbsolutePosition - 1, 1)) = -1 Then
                    
                    strSql = gstrStoredProcedure("sp_AtualizaParcela", dbc_intReceita.BoundText & ", " & txtintExercicio & ", " & gstrENulo(adoResultado!intNumeroParcela) & ", " & gstrConvDtParaSql(adoResultado!dtmVencimento) & ", " & gstrConvDtParaSql(txtdtmdtinscricao) & ", " & gstrConvVrParaSql(gstrConvVrDoSql(gstrENulo(adoResultado!dblValor), 2, , True)) & ", " & Val(gstrENulo(adoResultado!intMoeda)), True)
                                                                      'COMPOSICAO DA RECEITA,           EXERCICIO,              INTPARCELA,                                         DTMDTVENCIMENTO,
                    If gobjBanco.CriaADO(strSql, 80, adoParcelas) Then
                        With adoParcelas
                            If Not .EOF Then
                            
                                blnAtualiza = True
                                
                                xadbAtualizadas.ReDim 0, intFor, 0, 10
                                
                                varAux = Space$(0) & adoResultado!Pkid                                              'Pkid da TbllancamentoValor
                                xadbAtualizadas(intFor, 0) = varAux
                                
                                xadbAtualizadas(intFor, 1) = -1                                                     'Traz todas as parcelas checadas
                                
                                varAux = Space$(0) & adoResultado!intNumeroParcela                                  'Número das parcelas
                                xadbAtualizadas(intFor, 2) = varAux
                                
                                varAux = Space$(0) & adoResultado!strMoeda                                          'Abreviatura da Moeda
                                xadbAtualizadas(intFor, 3) = varAux
                                
                                varAux = Space$(0) & gstrDataFormatada(adoResultado!dtmVencimento)                  'Vencimento da parcela
                                xadbAtualizadas(intFor, 4) = varAux
                                
                                varAux = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))  'Valor da parcela
                                xadbAtualizadas(intFor, 5) = varAux
                                
                                'varAux = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value))      'Valor de Juros Atualizado
                                'xadbAtualizadas(intFor, 6) = varAux
                                
                                varAux = Space$(0) & CCur(gstrConvVrDoSql("0"))                                     'Valor de Juros Atualizado
                                xadbAtualizadas(intFor, 6) = varAux

                                
                                'varAux = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value))      'Valor de Multa Atualizado
                                'xadbAtualizadas(intFor, 7) = varAux
                                
                                varAux = Space$(0) & CCur(gstrConvVrDoSql("0"))      'Valor de Multa Atualizado
                                xadbAtualizadas(intFor, 7) = varAux
                                
                                varAux = CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))               'Valor de Correção Atualizado
                                xadbAtualizadas(intFor, 8) = varAux
                                
                                'varAux = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)) 'Valor de Total Atualizado
                                varAux = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql("0")) + CCur(gstrConvVrDoSql("0")) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)) 'Valor de Total Atualizado
                                dblTotal = dblTotal + varAux
                                xadbAtualizadas(intFor, 9) = varAux
                                
                                varAux = Space$(0) & gstrENulo(adoResultado!intMoeda)                                           'Pkid da tblMoeda
                                xadbAtualizadas(intFor, 10) = varAux
                            End If
                        End With
                    Else
                        LimpaGrids
                        Exit Sub
                    End If
                    intFor = intFor + 1
                End If
                adoResultado.MoveNext
            Loop
            txtdblValorImposto = gstrConvVrDoSql(vetImpostoTaxa(0, 2) * dblTotal, , , True)
            txtdblValorTaxas = gstrConvVrDoSql(vetImpostoTaxa(0, 3) * dblTotal, , , True)
        Else
            LimpaGrids
            Exit Sub
        End If
    Else
        LimpaGrids
        Exit Sub
    End If
    
    If blnAtualiza Then
        LimpaGrids
        
        Set tdb_Parcelas.Array = xadbAtualizadas
        tdb_Parcelas.ReBind
        tdb_Parcelas.Refresh
        
        Set xadbParcelas = tdb_Parcelas.Array
    End If
    
End Sub

Private Function blnDadosAtualizacao() As Boolean
    blnDadosAtualizacao = False
    
    If dbc_intReceita.MatchedWithList = False Then
        ExibeMensagem "O campo de Composição da Receita é obrigatório."
        dbc_intReceita.SetFocus
        Exit Function
    ElseIf Trim(Len(txtintExercicio)) <> 4 Then
        ExibeMensagem "O campo de exercício deve ser preenchido corretamente."
        txtintExercicio.SetFocus
        Exit Function
    ElseIf Trim(txtdtmdtinscricao) = "" Then
        ExibeMensagem "O campo de data da inscrição deve ser preenchido corretamente."
        txtdtmdtinscricao.SetFocus
        Exit Function
    ElseIf Trim(txtdtmdtinscricao) <> "" Then
        If gblnDataValida(txtdtmdtinscricao, True) = False Then
            txtdtmdtinscricao.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosAtualizacao = True
End Function

Sub VerificaMascaraInscricao(intTipoComposicao As Integer)
Dim strSql As String
Dim adoResultado As ADODB.Recordset
Dim strMascara   As String
    
    strMascara = ""
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & intTipoComposicao
    strSql = strSql & " Order By intSequencia"
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

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    
    If Trim(txtPKId) = "" Then
        ExibeMensagem "Não foi selecionado nenhum lançamento."
        Exit Function
    ElseIf Not blnReceitasOK Then
        ExibeMensagem "Não encontrado receitas para impostos e taxas."
        Exit Function
    ElseIf Not blnVerificaParcelas Then
        ExibeMensagem "Não existem parcelas para gerar o Inscrição de Dívida Ativa."
        Exit Function
    ElseIf Not dbc_intReceita.MatchedWithList Then
        ExibeMensagem "O campo de Composição da Receita é obrigatório."
        dbc_intReceita.SetFocus
        Exit Function
    ElseIf Not blnExisteParametroAtualizacao Then
        ExibeMensagem "Não foi encontrado PARAMETROS DE DÍVIDA ATÍVA para inscrever em dívida atíva esta inscrição cadastral!"
        Exit Function
    ElseIf Trim(mskstrInscricao.Text) = "" Then
        ExibeMensagem "O campo de Inscrição é obrigatório."
        mskstrInscricao.SetFocus
        Exit Function
    ElseIf Trim(txtintExercicio.Text) = "" Then
        ExibeMensagem "O campo de Exercício é obrigatório."
        txtintExercicio.SetFocus
        Exit Function
    ElseIf Trim(txtstrAviso.Text) = "" Then
        ExibeMensagem "O campo Aviso é obrigatório."
        txtstrAviso.SetFocus
        Exit Function
    ElseIf Trim(txtdtmdtinscricao) = "" Then
        ExibeMensagem "O campo de Data da Inscrição deve ser preenchido corretamente."
        txtdtmdtinscricao.SetFocus
        Exit Function
    ElseIf Trim(txtdtmdtinscricao) <> "" Then
        If gblnDataValida(txtdtmdtinscricao, True) = False Then
            txtdtmdtinscricao.SetFocus
            Exit Function
        End If
    End If
    If Trim(txtstrnomeproprietario) = "" Then
        ExibeMensagem "O campo Proprietário é obrigatório."
        txtstrnomeproprietario.SetFocus
        Exit Function
    ElseIf Trim(txt_LogradouroN) = "" Then
        ExibeMensagem "O campo Logradoruro é obrigatório."
        txt_LogradouroN.SetFocus
        Exit Function
    ElseIf Trim(txtstrNumeroN) = "" Then
        ExibeMensagem "O campo Número é obrigatório."
        txtstrNumeroN.SetFocus
        Exit Function
    ElseIf Trim(txt_BairroN) = "" Then
        ExibeMensagem "O campo Bairro é obrigatório."
        txt_BairroN.SetFocus
        Exit Function
    ElseIf Trim(txt_CepN) = "" Then
        ExibeMensagem "O campo de CEP é obrigatório."
        txt_CepN.SetFocus
        Exit Function
    ElseIf Trim(txt_MunicipioN) = "" Then
        ExibeMensagem "O campo de Município é obrigatório."
        txt_MunicipioN.SetFocus
        Exit Function
    ElseIf Trim(txt_UFN) = "" Then
        ExibeMensagem "O campo de unidade federativa  é obrigatório."
        txt_UFN.SetFocus
        Exit Function
   ElseIf blnInscritoEmDividaAtiva Then
        ExibeMensagem "Este tributo já se encontra cadastrado"
        Exit Function
   End If
    
    blnDadosOk = True
End Function

Private Function blnVerificaParcelas() As Boolean
    Dim intFor As Integer
    blnVerificaParcelas = False
    For intFor = 0 To xadbParcelas.Count(1) - 1
        If Val(xadbParcelas(intFor, 0)) > 0 And Val(xadbParcelas(intFor, 1)) = -1 Then
            blnVerificaParcelas = True
            Exit For
        End If
    Next
End Function

Private Function blnSalvarLancamentoDA() As Boolean
    Dim strSql           As String
    Dim intFor           As Integer
    Dim adoResultado     As ADODB.Recordset
    Dim strIDLanctoValor As String
    Dim strUpdate        As String
    
    blnSalvarLancamentoDA = False

    If bnlVerificaParametros = False Then Exit Function
    
    strUpdate = strParametrosDividaAtiva(dbc_intReceita.BoundText)
    If strUpdate = "" Then
       Exit Function
    End If
    
    Set gobjBanco = New clsBanco
    Set adoResultado = New ADODB.Recordset
    
    'If gobjBanco.CriaADO("SELECT seqTBLDATIVA.NEXTVAL as Pkid FROM DUAL", 5, adoResultado) Then
     '   If Not adoResultado.EOF Then
            strSql = ""
            'strSql = strSql & IIf(bytDBType = Oracle, "Begin ", "")
            strSql = strSql & "Insert Into " & gstrDativa
            strSql = strSql & "(intlancamentoalfa, intfolha, intlivro, dtmdtinscricao, strobservacao, strnomeproprietario, strcnpjcpf, stridentidade, strlogradouro, strnumero, strcomplemento, strbairro, strmunicipio, struf, intcep, strlogradouroc, strnumeroc, strcomplementoc, strbairroc, strmunicipioc, strufc, intcepc, strpromissario, strindexador, dblvlindexador, dtmdtatualizacao, lngcodusr, intcertidao, dblValorImposto, dblValorTaxas)"
            strSql = strSql & " Values "
            'strSql = strSql & "(" & gstrENulo(adoResultado!Pkid) & ", "
            strSql = strSql & "(" & txtPKId & ", "
            strSql = strSql & gstrENulo(txtintFolha, , True) & ", "
            strSql = strSql & gstrENulo(txtintLivro, , True) & ", "
            strSql = strSql & gstrConvDtParaSql(gstrENulo(txtdtmdtinscricao, , True)) & ", "
            strSql = strSql & "'" & gstrENulo(txtHistorico) & "', "
            strSql = strSql & "'" & gstrENulo(txtstrnomeproprietario) & "', "
            strSql = strSql & "'" & gstrENulo(txtstrcnpjcpf) & "', "
            strSql = strSql & "'" & gstrENulo(txtstridentidade) & "', "
            strSql = strSql & "'" & gstrENulo(txt_Logradouro) & "', "
            strSql = strSql & "'" & gstrENulo(txtstrNumero) & "', "
            strSql = strSql & "'" & gstrENulo(txtstrComplemento) & "', "
            strSql = strSql & "'" & gstrENulo(txt_Bairro) & "', "
            strSql = strSql & "'" & gstrENulo(txt_Municipio) & "', "
            strSql = strSql & "'" & gstrENulo(txt_UF) & "', "
            strSql = strSql & gstrENulo(Replace(txt_Cep, "-", ""), , True) & ", "
            strSql = strSql & "'" & gstrENulo(txt_LogradouroN) & "', "
            strSql = strSql & "'" & gstrENulo(txtstrNumeroN) & "', "
            strSql = strSql & "'" & gstrENulo(txtstrComplementoN) & "', "
            strSql = strSql & "'" & gstrENulo(txt_BairroN) & "', "
            strSql = strSql & "'" & gstrENulo(txt_MunicipioN) & "', "
            strSql = strSql & "'" & gstrENulo(txt_UFN) & "', "
            strSql = strSql & gstrENulo(Replace(txt_CepN, "-", ""), , True) & ", "
            strSql = strSql & "'" & gstrENulo(txtstrpromissario) & "', "
            strSql = strSql & "'" & gstrENulo(txtstrindexador) & "', "
            strSql = strSql & gstrENulo(gstrConvVrParaSql(gstrConvVrDoSql(txtdblvlindexador)), , True) & ", "
            strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSql = strSql & glngCodUsr & ", "
            strSql = strSql & gstrENulo(txtintCertidao, , True) & ", "
            strSql = strSql & gstrConvVrParaSql(txtdblValorImposto) & ", "
            strSql = strSql & gstrConvVrParaSql(txtdblValorTaxas) & ")"
            'strSql = strSql & IIf(bytDBType = Oracle, ";", "")
            If Not gobjBanco.Execute(strSql) Then
                Exit Function
            End If

            strSql = IIf(bytDBType = Oracle, "Begin ", "")
            
            For intFor = 0 To xadbParcelas.UpperBound(1)
                If xadbParcelas(intFor, 1) = -1 Then
                    strSql = strSql & "Insert Into " & gstrDaParcel
                    strSql = strSql & "(intdativa, intparcela, dblvalor, dtmdtvencimento, dblmulta, dblcorrecaomonet, dbljuros, intmoeda, dtmdtatualizacao, lngcodusr) "
                    strSql = strSql & "Values("
                    strSql = strSql & glngRetornaPkidTabelaPai("seqTBLDATIVA", gstrDativa) & ", "
                    strSql = strSql & xadbParcelas(intFor, 2) & ", "
                    strSql = strSql & gstrConvVrParaSql(xadbParcelas(intFor, 5)) & ", "
                    strSql = strSql & gstrConvDtParaSql(xadbParcelas(intFor, 4)) & ", "
                    strSql = strSql & gstrConvVrParaSql(xadbParcelas(intFor, 7)) & ", "
                    strSql = strSql & gstrConvVrParaSql(xadbParcelas(intFor, 8)) & ", "
                    strSql = strSql & gstrConvVrParaSql(xadbParcelas(intFor, 6)) & ", "
                    strSql = strSql & xadbParcelas(intFor, 10) & ", "
                    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSql = strSql & glngCodUsr
                    strSql = strSql & ")" & IIf(bytDBType = Oracle, ";", "")
                End If
            Next
            
            
            
            strSql = strSql & strUpdate & " " & IIf(bytDBType = Oracle, ";", "")
            
            For intFor = 0 To xadbParcelas.UpperBound(1)
                If xadbParcelas(intFor, 1) = -1 Then
                    strIDLanctoValor = strIDLanctoValor & xadbParcelas(intFor, 0) & ","
                End If
            Next
        
            strIDLanctoValor = Mid(strIDLanctoValor, 1, Len(strIDLanctoValor) - 1)
            
            strSql = strSql & "UPDATE " & gstrLancamentoValor & " SET INTLANCAMENTOALFADATIVA = " & txtPKId & " WHERE pkid in(" & strIDLanctoValor & ");"
                    
            strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    
        'Else
        '    Exit Function
        'End If
    'Else
        'Exit Function
    'End If
    
    If Not gobjBanco.Execute(strSql) Then
        Exit Function
    End If
    
    blnSalvarLancamentoDA = True
End Function

Private Sub PreencheCertidaoLivroFolha(intComposicaoDaReceita As Double)
    Dim adoResultado           As New ADODB.Recordset
    Dim strSql                 As String
    Dim intCertidao            As Double
    Dim intFolha               As Double
    Dim intLivro               As Double
    Dim intQtdCertidaoUltFolha As Double
    Dim blnProximoLivro        As Boolean
    
    strSql = "SELECT "
    strSql = strSql & "PDA.intCertidao, "
    strSql = strSql & "PDA.intFolha, "
    strSql = strSql & "PDA.intLivro, "
    strSql = strSql & "PDA.intFolhaPorLivro, "
    strSql = strSql & "PDA.intCertidaoPorFolha, "
    strSql = strSql & "PDA.intQtdCertidaoUltFolha "
    strSql = strSql & "FROM "
    strSql = strSql & gstrParametroDividaAtiva & " PDA "
    strSql = strSql & "WHERE "
    
    'Parâmetros p/ composições específicas
    If gobjBanco.CriaADO(strSql & "PDA.intComposicaoDaReceita = " & intComposicaoDaReceita, 10, adoResultado) Then
       If adoResultado.EOF Then
        
          'Parâmetros p/ composições diversas
          Set adoResultado = New ADODB.Recordset
          If gobjBanco.CriaADO(strSql & "PDA.intComposicaoDaReceita IS NULL ", 10, adoResultado) Then
             If adoResultado.EOF Then
                txtintCertidao.Text = ""
                txtintFolha.Text = ""
                txtintLivro.Text = ""
                ExibeMensagem "Não há parâmetros para inscrição em Dívida Ativa."
                Exit Sub
             End If
          End If
       End If
    End If
           
    With adoResultado
         intCertidao = !intCertidao
         intFolha = !intFolha
         intLivro = !intLivro
         intQtdCertidaoUltFolha = !intQtdCertidaoUltFolha
               
         If (intQtdCertidaoUltFolha Mod !intCertidaoPorFolha = 0) And intQtdCertidaoUltFolha <> 0 Then
            intQtdCertidaoUltFolha = 0
            If (intFolha Mod !intFolhaPorLivro = 0) And intFolha <> 0 Then
               intFolha = 1
               intLivro = intLivro + 1
               
                'Vamos verificar se o livro esta disponivel
                blnProximoLivro = False
                Set adoResultado = New ADODB.Recordset
                Do Until blnProximoLivro
                  
                   If gblnExisteCodigo(1, gstrDativa, "intLivro", Val(intLivro)) Then
                      intLivro = intLivro + 1
                   Else
                      blnProximoLivro = True
                   End If
                Loop

            Else
               intFolha = intFolha + 1
            End If
         End If
                      
    End With
    
    intCertidao = intCertidao + 1
    intQtdCertidaoUltFolha = intQtdCertidaoUltFolha + 1
    
    txtintCertidao.Text = intCertidao
    txtintFolha.Text = intFolha
    txtintLivro.Text = intLivro
    txtintQtdCertidaoUltFolha.Text = intQtdCertidaoUltFolha
                
End Sub

Private Function strParametrosDividaAtiva(intComposicaoDaReceita As Double) As String
    Dim adoResultado As ADODB.Recordset
    Dim strSql As String
    Dim blnExisteRegistro As Boolean
    Dim intCertidao            As Double
    Dim intFolha               As Double
    Dim intLivro               As Double
    Dim intQtdCertidaoUltFolha As Double
    Dim intFolhaPorLivro       As Double
    Dim intCertidaoPorFolha    As Double
    Dim blnProximoLivro        As Boolean
    Dim blnComposicao          As Boolean
    
    intCertidao = Val(txtintCertidao.Text)
    intFolha = Val(txtintFolha.Text)
    intLivro = Val(txtintLivro.Text)
    intQtdCertidaoUltFolha = Val(txtintQtdCertidaoUltFolha.Text)
    
    strSql = "SELECT "
    strSql = strSql & "PDA.intFolhaPorLivro, "
    strSql = strSql & "PDA.intCertidaoPorFolha "
    strSql = strSql & "FROM "
    strSql = strSql & gstrParametroDividaAtiva & " PDA "
    strSql = strSql & "WHERE "
    
    'Parâmetros p/ composições específicas
    blnComposicao = False
    Set adoResultado = New ADODB.Recordset
    If gobjBanco.CriaADO(strSql & "PDA.intComposicaoDaReceita = " & intComposicaoDaReceita, 10, adoResultado) Then
       If adoResultado.EOF Then
          
          'Parâmetros p/ composições diversas
          Set adoResultado = New ADODB.Recordset
          If gobjBanco.CriaADO(strSql & "PDA.intComposicaoDaReceita IS NULL ", 10, adoResultado) Then
             If adoResultado.EOF Then
                Exit Function
             End If
          End If
       Else
          blnComposicao = True
       End If
    End If
    
    'intFolhaPorLivro = adoResultado!intFolhaPorLivro
    'intCertidaoPorFolha = adoResultado!intCertidaoPorFolha
    
    'blnExisteRegistro = True
    'Do While blnExisteRegistro
       If gblnExisteCodigo(2, gstrDativa, "intCertidao", Val(intCertidao), "intFolha", Val(intFolha), "intLivro", Val(intLivro)) Then
          
          'If (intQtdCertidaoUltFolha Mod intCertidaoPorFolha = 0) And intQtdCertidaoUltFolha <> 0 Then
          '   intQtdCertidaoUltFolha = 0
          '   If (intFolha Mod intFolhaPorLivro = 0) And intFolha <> 0 Then
          '      intFolha = 1
          '      intLivro = intLivro + 1
          '
                'Procura o próximo livro disponível
          '      strSQL = "SELECT "
          '      strSQL = strSQL & "Pkid "
          '      strSQL = strSQL & "FROM "
          '      strSQL = strSQL & gstrDividaAtiva & " "
          '      strSQL = strSQL & "WHERE "
            
          '      blnProximoLivro = False
          '      Set adoResultado = New ADODB.Recordset
          '      Do Until blnProximoLivro
                   
          '         If gobjBanco.CriaADO(strSQL & "intLivro = " & intLivro, 10, adoResultado) Then
          '            If adoResultado.EOF Then
          '               intLivro = intLivro + 1
          '            Else
          '               blnProximoLivro = True
          '            End If
          '         End If
          '      Loop
          '
          '   Else
          '      intFolha = intFolha + 1
          '   End If
          'End If
          'intCertidao = intCertidao + 1
          'intQtdCertidaoUltFolha = intQtdCertidaoUltFolha + 1
             
          ExibeMensagem "A Certidão: " & intCertidao & ", Folha: " & intFolha & ", Livro: " & intLivro & _
                        " já estão cadastrados. Não é possível inscrever em Dívida Ativa. Reinicie a inscrição. "
          Exit Function
       'Else
       '   blnExisteRegistro = False
       End If
    'Loop
    
    strSql = "UPDATE " & gstrParametroDividaAtiva
    strSql = strSql & " SET intCertidao = " & intCertidao & ", intFolha = " & intFolha & ", intLivro = " & intLivro & ", intQtdCertidaoUltFolha = " & intQtdCertidaoUltFolha
    
    If blnComposicao Then
        strSql = strSql & " WHERE intComposicaoDaReceita = " & intComposicaoDaReceita
    Else
        strSql = strSql & " WHERE intComposicaoDaReceita IS NULL"
    End If
    
    strParametrosDividaAtiva = strSql
    
End Function

Private Function blnInscritoEmDividaAtiva() As Boolean

    Dim strSql              As String
    Dim strIDLanctoValor    As String
    Dim intFor              As Integer
    Dim adoRec              As New ADODB.Recordset

    blnInscritoEmDividaAtiva = False
    
    For intFor = 0 To xadbParcelas.UpperBound(1)
        If xadbParcelas(intFor, 1) = -1 Then
            strIDLanctoValor = strIDLanctoValor & xadbParcelas(intFor, 0) & ","
        End If
    Next

    strIDLanctoValor = Mid(strIDLanctoValor, 1, Len(strIDLanctoValor) - 1)

    strSql = "SELECT DISTINCT LA.PKID "
    strSql = strSql & "FROM " & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrDativa & " DA "
    strSql = strSql & "Where DA.intLancamentoAlfa = LA.Pkid "
    strSql = strSql & "AND LA.PKID = LV.INTLANCAMENTOALFA "
    strSql = strSql & "AND LV.pkid IN(" & strIDLanctoValor & ") "
    strSql = strSql & " AND LV.INTLANCAMENTOALFADATIVA IS NOT NULL "
    strSql = strSql & " AND LA.INTEXERCICIO = " & txtintExercicio
    strSql = strSql & " AND LA.Strinscricao = '" & (String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text)) & "'"
    strSql = strSql & " AND LA.STRNUMEROAVISO = '" & String(gintLenNumAviso = 10 - Len(txtstrAviso), "0") & txtstrAviso.Text & "'"
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        If Not adoRec.EOF Then
            blnInscritoEmDividaAtiva = True
        End If
    End If
    
End Function

Private Function blnExisteParametroAtualizacao() As Boolean

    Dim strSql As String
    Dim adoRec As New ADODB.Recordset
    
    blnExisteParametroAtualizacao = False
    
    strSql = "SELECT intComposicaoDaReceita "
    strSql = strSql & "FROM " & gstrParametroDividaAtiva
    strSql = strSql & " WHERE intComposicaoDaReceita = " & dbc_intReceita.BoundText
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        If adoRec.EOF Then
            strSql = "SELECT intComposicaoDaReceita "
            strSql = strSql & "FROM " & gstrParametroDividaAtiva
            strSql = strSql & " WHERE intComposicaoDaReceita is Null"
            If gobjBanco.CriaADO(strSql, 10, adoRec) Then
                If adoRec.EOF Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    blnExisteParametroAtualizacao = True
    
End Function

Private Sub AdicionaTaxasImpostos(intPosicao As Integer)

    Dim strSql As String
    Dim adoResultado As New ADODB.Recordset
    
    'Query para buscar o valor dos impostos
    strSql = "Select "
    strSql = strSql & gstrISNULL("TT.dblImposto", "0") & " As  dblImposto, "
    strSql = strSql & gstrISNULL("TT.dblTaxa", "0") & "+" & gstrISNULL("TT.dblTaxa1", "0") & " As dblTaxa "
    strSql = strSql & "From "
    strSql = strSql & "(SELECT "
    strSql = strSql & "LV.INTPARCELA, "
    strSql = strSql & "SUM(LR.DBLVALOR) dblImposto, "
    strSql = strSql & "0 dblTaxa, "
    strSql = strSql & "0 dblTaxa1"
    strSql = strSql & " From "
    strSql = strSql & gstrLancamentoAlfa & " LA,"
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR, "
    strSql = strSql & gstrReceita & " RC "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.intLancamentoAlfa"
    strSql = strSql & " AND LV.PKID = LR.INTLANCAMENTOVALOR"
    strSql = strSql & " AND RC.PKID = LR.INTRECEITA"
    strSql = strSql & " AND LA.PKID = " & Trim(txtPKId)
    strSql = strSql & " AND LV.INTPARCELA = " & xadbParcelas(intPosicao, 2)
    strSql = strSql & " AND RC.BYTTIPO in(2,1,5,6)"
    strSql = strSql & " AND LV.Dtmdtvencimento < " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    strSql = strSql & " Group By LV.intParcela , LV.DBLVALOR"
  
    strSql = strSql & " UNION "
    
    'Query para buscar o valor das taxas
    strSql = strSql & "SELECT "
    strSql = strSql & "LV.INTPARCELA, "
    strSql = strSql & "0 dblImposto, "
    strSql = strSql & "SUM(LR.DBLVALOR) dblTaxa,"
    strSql = strSql & "0 dblTaxa1"
    strSql = strSql & " From "
    strSql = strSql & gstrLancamentoAlfa & " LA,"
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR, "
    strSql = strSql & gstrReceita & " RC "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.intLancamentoAlfa"
    strSql = strSql & " AND LV.PKID = LR.INTLANCAMENTOVALOR"
    strSql = strSql & " AND RC.PKID = LR.INTRECEITA"
    strSql = strSql & " AND LA.PKID = " & Trim(txtPKId)
    strSql = strSql & " AND LV.INTPARCELA = " & xadbParcelas(intPosicao, 2)
    strSql = strSql & " AND RC.BYTTIPO in(3,4)"
    strSql = strSql & " AND LV.Dtmdtvencimento <  " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    strSql = strSql & " Group By LV.intParcela , LV.DBLVALOR"
    
    strSql = strSql & " UNION "
    
    'Query para buscar o valor das taxas que os tipos de receitas não sejam (1,2,3,4,5,6)
    strSql = strSql & "SELECT "
    strSql = strSql & "LV.INTPARCELA, "
    strSql = strSql & "0 dblImposto, "
    strSql = strSql & "0 dblTaxa,"
    strSql = strSql & "SUM(LR.DBLVALOR) dblTaxa"
    strSql = strSql & " From "
    strSql = strSql & gstrLancamentoAlfa & " LA,"
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR, "
    strSql = strSql & gstrReceita & " RC "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.intLancamentoAlfa"
    strSql = strSql & " AND LV.PKID = LR.INTLANCAMENTOVALOR"
    strSql = strSql & " AND RC.PKID = LR.INTRECEITA"
    strSql = strSql & " AND LA.PKID = " & Trim(txtPKId)
    strSql = strSql & " AND LV.INTPARCELA = " & xadbParcelas(intPosicao, 2)
    strSql = strSql & " AND not RC.BYTTIPO in(1,2,3,4,5,6)"
    strSql = strSql & " AND LV.Dtmdtvencimento <  " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    strSql = strSql & " Group By LV.intParcela , LV.DBLVALOR ) TT"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 15, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                vetImpostoTaxa(0, 0) = vetImpostoTaxa(0, 0) + gstrConvVrDoSql(gstrENulo(!dblImposto), 2, , True)
                vetImpostoTaxa(0, 1) = vetImpostoTaxa(0, 1) + gstrConvVrDoSql(gstrENulo(!dblTaxa), 2, , True)
                
                'Vamos obter as porcentagens de Taxas e de Impostos
                If CDbl(vetImpostoTaxa(0, 0)) > 0 Then
                    vetImpostoTaxa(0, 2) = (CDbl(vetImpostoTaxa(0, 0)) * 100) / (CDbl(vetImpostoTaxa(0, 1)) + CDbl(vetImpostoTaxa(0, 0))) / 100
                Else
                    vetImpostoTaxa(0, 2) = 0
                End If
                
                If CDbl(vetImpostoTaxa(0, 1)) > 0 Then
                    vetImpostoTaxa(0, 3) = (CDbl(vetImpostoTaxa(0, 1)) * 100) / (CDbl(vetImpostoTaxa(0, 1)) + CDbl(vetImpostoTaxa(0, 0))) / 100
                Else
                    vetImpostoTaxa(0, 3) = 0
                End If
            End If
        End With
    End If

End Sub

Private Sub RemoveTaxasImpostos(intPosicao As Integer)

    Dim strSql As String
    Dim adoResultado As New ADODB.Recordset
    
    'Query para buscar o valor dos impostos
    strSql = "Select "
    strSql = strSql & gstrISNULL("TT.dblImposto", "0") & " As  dblImposto, "
    strSql = strSql & gstrISNULL("TT.dblTaxa", "0") & "+" & gstrISNULL("TT.dblTaxa1", "0") & " As dblTaxa "
    strSql = strSql & "From "
    strSql = strSql & "(SELECT "
    strSql = strSql & "LV.INTPARCELA, "
    strSql = strSql & "SUM(LR.DBLVALOR) dblImposto, "
    strSql = strSql & "0 dblTaxa, "
    strSql = strSql & "0 dblTaxa1"
    strSql = strSql & " From "
    strSql = strSql & gstrLancamentoAlfa & " LA,"
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR, "
    strSql = strSql & gstrReceita & " RC "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.intLancamentoAlfa"
    strSql = strSql & " AND LV.PKID = LR.INTLANCAMENTOVALOR"
    strSql = strSql & " AND RC.PKID = LR.INTRECEITA"
    strSql = strSql & " AND LA.PKID = " & Trim(txtPKId)
    strSql = strSql & " AND LV.INTPARCELA = " & xadbParcelas(intPosicao, 2)
    strSql = strSql & " AND RC.BYTTIPO in(2,1,5,6)"
    strSql = strSql & " AND LV.Dtmdtvencimento < " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    strSql = strSql & " Group By LV.intParcela , LV.DBLVALOR"
  
    strSql = strSql & " UNION "
    
    'Query para buscar o valor das taxas
    strSql = strSql & "SELECT "
    strSql = strSql & "LV.INTPARCELA, "
    strSql = strSql & "0 dblImposto, "
    strSql = strSql & "SUM(LR.DBLVALOR) dblTaxa,"
    strSql = strSql & "0 dblTaxa1"
    strSql = strSql & " From "
    strSql = strSql & gstrLancamentoAlfa & " LA,"
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR, "
    strSql = strSql & gstrReceita & " RC "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.intLancamentoAlfa"
    strSql = strSql & " AND LV.PKID = LR.INTLANCAMENTOVALOR"
    strSql = strSql & " AND RC.PKID = LR.INTRECEITA"
    strSql = strSql & " AND LA.PKID = " & Trim(txtPKId)
    strSql = strSql & " AND LV.INTPARCELA = " & xadbParcelas(intPosicao, 2)
    strSql = strSql & " AND RC.BYTTIPO in(3,4)"
    strSql = strSql & " AND LV.Dtmdtvencimento <  " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    strSql = strSql & " Group By LV.intParcela , LV.DBLVALOR"
    
    strSql = strSql & " UNION "
    
    'Query para buscar o valor das taxas que os tipos de receitas não sejam (1,2,3,4,5,6)
    strSql = strSql & "SELECT "
    strSql = strSql & "LV.INTPARCELA, "
    strSql = strSql & "0 dblImposto, "
    strSql = strSql & "0 dblTaxa,"
    strSql = strSql & "SUM(LR.DBLVALOR) dblTaxa"
    strSql = strSql & " From "
    strSql = strSql & gstrLancamentoAlfa & " LA,"
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR, "
    strSql = strSql & gstrReceita & " RC "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.intLancamentoAlfa"
    strSql = strSql & " AND LV.PKID = LR.INTLANCAMENTOVALOR"
    strSql = strSql & " AND RC.PKID = LR.INTRECEITA"
    strSql = strSql & " AND LA.PKID = " & Trim(txtPKId)
    strSql = strSql & " AND LV.INTPARCELA = " & xadbParcelas(intPosicao, 2)
    strSql = strSql & " AND not RC.BYTTIPO in(1,2,3,4,5,6)"
    strSql = strSql & " AND LV.Dtmdtvencimento <  " & gstrConvDtParaSql(IIf(txtdtmdtinscricao <> "", txtdtmdtinscricao, Date))
    strSql = strSql & " Group By LV.intParcela , LV.DBLVALOR ) TT"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 15, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                vetImpostoTaxa(0, 0) = vetImpostoTaxa(0, 0) - gstrConvVrDoSql(gstrENulo(!dblImposto))
                vetImpostoTaxa(0, 1) = vetImpostoTaxa(0, 1) - gstrConvVrDoSql(gstrENulo(!dblTaxa))
                
                'Vamos obter as porcentagens de Taxas e de Impostos
                If CDbl(vetImpostoTaxa(0, 0)) > 0 Then
                    vetImpostoTaxa(0, 2) = (CDbl(vetImpostoTaxa(0, 0)) * 100) / (CDbl(vetImpostoTaxa(0, 1)) + CDbl(vetImpostoTaxa(0, 0))) / 100
                Else
                    vetImpostoTaxa(0, 2) = 0
                End If
                
                If CDbl(vetImpostoTaxa(0, 1)) > 0 Then
                    vetImpostoTaxa(0, 3) = (CDbl(vetImpostoTaxa(0, 1)) * 100) / (CDbl(vetImpostoTaxa(0, 1)) + CDbl(vetImpostoTaxa(0, 0))) / 100
                Else
                    vetImpostoTaxa(0, 3) = 0
                End If
                
            End If
        End With
    End If

End Sub

Private Sub AdicionaRemoveTaixasImpostos()
    Dim intFor As Integer

        With tdb_Parcelas
            If Not .EOF And Not .BOF Then
                ReDim vetImpostoTaxa(0, 3)
                    For intFor = 0 To xadbParcelas.UpperBound(1)
                        If xadbParcelas(intFor, 1) = -1 Then
                            AdicionaTaxasImpostos intFor
                            txtdblValorImposto = gstrConvVrDoSql(vetImpostoTaxa(0, 0), 2)
                            txtdblValorTaxas = gstrConvVrDoSql(vetImpostoTaxa(0, 1), 2)
                        Else
                            RemoveTaxasImpostos intFor
                            txtdblValorImposto = vetImpostoTaxa(0, 0)
                            txtdblValorTaxas = vetImpostoTaxa(0, 1)
                        End If
                    Next
            End If
        End With

End Sub

Private Function bnlVerificaParametros() As Boolean
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Dim adoResultado1    As ADODB.Recordset
    
    bnlVerificaParametros = False
    
    strSql = "SELECT "
    strSql = strSql & "PDA.INTCERTIDAO, "
    strSql = strSql & "PDA.INTFOLHA, "
    strSql = strSql & "PDA.INTLIVRO, "
    strSql = strSql & "pda.intfolhaporlivro, "
    strSql = strSql & "pda.intcertidaoporfolha, "
    strSql = strSql & "intQtdCertidaoUltFolha "
    strSql = strSql & "FROM "
    strSql = strSql & gstrParametroDividaAtiva & " PDA "
    strSql = strSql & "Where "
    strSql = strSql & "PDA.intComposicaoDaReceita = " & dbc_intReceita.BoundText & " AND "
    strSql = strSql & "PDA.INTCERTIDAO =" & Trim(txtintCertidao) & " AND "
    strSql = strSql & "PDA.INTFOLHA =" & Trim(txtintFolha) & " AND "
    strSql = strSql & "PDA.INTLIVRO =" & Trim(txtintLivro)
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                strSql = "SELECT "
                strSql = strSql & "PDA.INTCERTIDAO, "
                strSql = strSql & "PDA.INTFOLHA, "
                strSql = strSql & "PDA.INTLIVRO, "
                strSql = strSql & "pda.intfolhaporlivro, "
                strSql = strSql & "pda.intcertidaoporfolha, "
                strSql = strSql & "intQtdCertidaoUltFolha "
                strSql = strSql & "FROM "
                strSql = strSql & gstrParametroDividaAtiva & " PDA "
                strSql = strSql & "Where "
                strSql = strSql & "PDA.intComposicaoDaReceita = " & dbc_intReceita.BoundText
                If gobjBanco.CriaADO(strSql, 10, adoResultado1) Then
                    With adoResultado1
                        If Not .EOF Then
                            If gblnExclusaoGravacaoOk("A", "Número de certidão " & txtintCertidao & " já se encontra cadastrado." & Chr(13) & _
                                                    "Deseja salvar D.A. com o número de certidão " & (!intCertidao + 1), True) Then
                                bnlVerificaParametros = True
                                txtintCertidao = !intCertidao + 1
                                txtintFolha = !intFolha
                                txtintLivro = !intLivro
                            Else
                                bnlVerificaParametros = False
                                Exit Function
                            End If
                        End If
                    End With
                End If
            Else
                strSql = "SELECT "
                strSql = strSql & "PDA.INTCERTIDAO, "
                strSql = strSql & "PDA.INTFOLHA, "
                strSql = strSql & "PDA.INTLIVRO, "
                strSql = strSql & "pda.intfolhaporlivro, "
                strSql = strSql & "pda.intcertidaoporfolha, "
                strSql = strSql & "intQtdCertidaoUltFolha "
                strSql = strSql & "FROM "
                strSql = strSql & gstrParametroDividaAtiva & " PDA "
                strSql = strSql & "Where "
                strSql = strSql & "PDA.intComposicaoDaReceita Is null AND "
                strSql = strSql & "PDA.INTCERTIDAO =" & Trim(txtintCertidao) & " AND "
                strSql = strSql & "PDA.INTFOLHA =" & Trim(txtintFolha) & " AND "
                strSql = strSql & "PDA.INTLIVRO =" & Trim(txtintLivro)
                
                Set gobjBanco = New clsBanco
                Set adoResultado = New ADODB.Recordset
                
                If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
                    With adoResultado
                        If Not .EOF Then
                            strSql = "SELECT "
                            strSql = strSql & "PDA.INTCERTIDAO, "
                            strSql = strSql & "PDA.INTFOLHA, "
                            strSql = strSql & "PDA.INTLIVRO, "
                            strSql = strSql & "pda.intfolhaporlivro, "
                            strSql = strSql & "pda.intcertidaoporfolha, "
                            strSql = strSql & "intQtdCertidaoUltFolha "
                            strSql = strSql & "FROM "
                            strSql = strSql & gstrParametroDividaAtiva & " PDA "
                            strSql = strSql & "Where "
                            strSql = strSql & "PDA.intComposicaoDaReceita Is null "
                            If gobjBanco.CriaADO(strSql, 10, adoResultado1) Then
                                With adoResultado1
                                    If Not .EOF Then
                                        If gblnExclusaoGravacaoOk("A", "Número de certidão " & txtintCertidao & " já se encontra cadastrado." & Chr(13) & _
                                                                "Deseja salvar D.A. com o número de certidão " & (!intCertidao + 1), True) Then
                                            bnlVerificaParametros = True
                                            txtintCertidao = !intCertidao + 1
                                            txtintFolha = !intFolha
                                            txtintLivro = !intLivro
                                        Else
                                            bnlVerificaParametros = False
                                            Exit Function
                                        End If
                                    End If
                                End With
                            End If
                        End If
                    End With
                End If
            End If
        End With
    End If
    bnlVerificaParametros = True
    
End Function

Private Function blnReceitasOK() As Boolean
    Dim strSql As String
    Dim adoResultado As New ADODB.Recordset
    
    blnReceitasOK = False
    
    strSql = "SELECT LR.Pkid "
    strSql = strSql & " From "
    strSql = strSql & gstrLancamentoAlfa & " LA,"
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.intLancamentoAlfa AND "
    strSql = strSql & "LV.PKID = LR.INTLANCAMENTOVALOR AND "
    strSql = strSql & "LA.PKID = " & Trim(txtPKId)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 15, adoResultado) Then
        With adoResultado
            If Not .EOF Then
               blnReceitasOK = True
            End If
        End With
    End If
End Function

Private Sub PreencheGridReceita()
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "SELECT "
    strSql = strSql & "RC.Pkid, RC.Strdescricao , Rc.Strsigla "
    strSql = strSql & "From "
    strSql = strSql & "tblLancamentoValor LV, "
    strSql = strSql & "tblLancamentoReceita LR, "
    strSql = strSql & "tblReceita RC "
    strSql = strSql & "Where "
    strSql = strSql & "LV.PKID = LR.INTLANCAMENTOVALOR AND "
    strSql = strSql & "RC.PKID = LR.INTRECEITA AND "
    strSql = strSql & "LV.Intlancamentoalfa = " & Val(txtPKId) & " "
    strSql = strSql & "Group By "
    strSql = strSql & "RC.Pkid, RC.Strdescricao , Rc.Strsigla "
    strSql = strSql & "Order by "
    strSql = strSql & "RC.Strdescricao , Rc.Strsigla "
    
    LeDaTabelaParaObj "", tdb_Receitas, strSql
End Sub
