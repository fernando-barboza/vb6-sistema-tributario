VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadLancamentoEconomico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamento Econômico"
   ClientHeight    =   8910
   ClientLeft      =   1890
   ClientTop       =   1995
   ClientWidth     =   9630
   Icon            =   "frmCadLancamentoEconomico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   7035
      Left            =   90
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   60
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   12409
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lançamento Econômico"
      TabPicture(0)   =   "frmCadLancamentoEconomico.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtPKId"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Cabecalho(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_socios"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_Lancamentos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_enquadramentos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SSTab1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Tributos / Parcelas"
      TabPicture(1)   =   "frmCadLancamentoEconomico.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_ValoresLancados"
      Tab(1).Control(1)=   "fra_Parcelas"
      Tab(1).Control(2)=   "fra_Cabecalho(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Publicidade / ISS / Feiras"
      TabPicture(2)   =   "frmCadLancamentoEconomico.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_Cabecalho(2)"
      Tab(2).Control(1)=   "fra_feiras"
      Tab(2).Control(2)=   "fra_publicidades"
      Tab(2).Control(3)=   "Frame1"
      Tab(2).ControlCount=   4
      Begin TabDlg.SSTab SSTab1 
         Height          =   1545
         Left            =   240
         TabIndex        =   86
         Top             =   1380
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2725
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Estabelecimento"
         TabPicture(0)   =   "frmCadLancamentoEconomico.frx":1096
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtStrnomeproprietario"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtSTRNOMEFANTASIA"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtStratividadebasica"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Local do Estabelecimento"
         TabPicture(1)   =   "frmCadLancamentoEconomico.frx":10B2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtstrBairro"
         Tab(1).Control(1)=   "txtstrLogradouro"
         Tab(1).Control(2)=   "txtstrComplemento"
         Tab(1).Control(3)=   "txtintCep"
         Tab(1).Control(4)=   "txtstrNumero"
         Tab(1).Control(5)=   "txtstrMunicipio"
         Tab(1).Control(6)=   "txtstrUf"
         Tab(1).Control(7)=   "lblintLogradouro"
         Tab(1).Control(8)=   "lblintBairro"
         Tab(1).Control(9)=   "lblintNumero"
         Tab(1).Control(10)=   "lblstrComplemento"
         Tab(1).Control(11)=   "lblintCep"
         Tab(1).Control(12)=   "lblstrMunicipio"
         Tab(1).Control(13)=   "lblstrUf"
         Tab(1).ControlCount=   14
         Begin VB.TextBox txtstrBairro 
            Height          =   300
            Left            =   -73995
            MaxLength       =   50
            TabIndex        =   12
            Top             =   780
            Width           =   2670
         End
         Begin VB.TextBox txtstrLogradouro 
            Height          =   300
            Left            =   -73995
            MaxLength       =   100
            TabIndex        =   9
            Top             =   420
            Width           =   4065
         End
         Begin VB.TextBox txtstrComplemento 
            Height          =   300
            Left            =   -68115
            MaxLength       =   10
            TabIndex        =   11
            Top             =   420
            Width           =   1710
         End
         Begin VB.TextBox txtintCep 
            Height          =   300
            Left            =   -67275
            MaxLength       =   9
            TabIndex        =   15
            Top             =   765
            Width           =   885
         End
         Begin VB.TextBox txtstrNumero 
            Height          =   300
            Left            =   -69600
            MaxLength       =   10
            TabIndex        =   10
            Top             =   420
            Width           =   825
         End
         Begin VB.TextBox txtstrMunicipio 
            Height          =   300
            Left            =   -70530
            MaxLength       =   50
            TabIndex        =   13
            Top             =   780
            Width           =   2115
         End
         Begin VB.TextBox txtstrUf 
            Height          =   300
            Left            =   -68070
            MaxLength       =   2
            TabIndex        =   14
            Top             =   780
            Width           =   375
         End
         Begin VB.TextBox txtStratividadebasica 
            Height          =   315
            Left            =   1380
            MaxLength       =   50
            TabIndex        =   8
            Top             =   1050
            Width           =   4650
         End
         Begin VB.TextBox txtSTRNOMEFANTASIA 
            Height          =   315
            Left            =   1380
            MaxLength       =   100
            TabIndex        =   7
            Top             =   720
            Width           =   7110
         End
         Begin VB.TextBox txtStrnomeproprietario 
            Height          =   315
            Left            =   1380
            MaxLength       =   100
            TabIndex        =   6
            Top             =   390
            Width           =   7110
         End
         Begin VB.Label lblintLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   -74880
            TabIndex        =   96
            Top             =   510
            Width           =   810
         End
         Begin VB.Label lblintBairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   -74475
            TabIndex        =   95
            Top             =   870
            Width           =   405
         End
         Begin VB.Label lblintNumero 
            AutoSize        =   -1  'True
            Caption         =   "N°"
            Height          =   195
            Left            =   -69825
            TabIndex        =   94
            Top             =   510
            Width           =   180
         End
         Begin VB.Label lblstrComplemento 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   -68640
            TabIndex        =   93
            Top             =   510
            Width           =   480
         End
         Begin VB.Label lblintCep 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Left            =   -67635
            TabIndex        =   92
            Top             =   855
            Width           =   315
         End
         Begin VB.Label lblstrMunicipio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   -71280
            TabIndex        =   91
            Top             =   855
            Width           =   705
         End
         Begin VB.Label lblstrUf 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   -68340
            TabIndex        =   90
            Top             =   870
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social"
            Height          =   195
            Left            =   360
            TabIndex        =   89
            Top             =   450
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Atividade Básica"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   1140
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nome Fantasia"
            Height          =   195
            Left            =   240
            TabIndex        =   87
            Top             =   810
            Width           =   1065
         End
      End
      Begin VB.Frame Frame4 
         Height          =   675
         Left            =   240
         TabIndex        =   81
         Top             =   2880
         Width           =   9090
         Begin VB.TextBox txtStrnaturezajuridica 
            Height          =   315
            Left            =   1440
            MaxLength       =   25
            TabIndex        =   16
            Top             =   240
            Width           =   1290
         End
         Begin VB.TextBox txtDtmdataabertura 
            Height          =   315
            Left            =   4110
            TabIndex        =   17
            Top             =   240
            Width           =   1080
         End
         Begin VB.TextBox txtDblareaocupada 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6390
            TabIndex        =   18
            Top             =   240
            Width           =   1170
         End
         Begin VB.TextBox txtDblnumeroempregados 
            Height          =   315
            Left            =   8580
            TabIndex        =   19
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Jurídica"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   330
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Data de Abertura"
            Height          =   195
            Left            =   2790
            TabIndex        =   84
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Área Ocupada"
            Height          =   195
            Left            =   5280
            TabIndex        =   83
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Empregados"
            Height          =   195
            Left            =   7650
            TabIndex        =   82
            Top             =   330
            Width           =   885
         End
      End
      Begin VB.Frame Fra_ValoresLancados 
         Caption         =   "Valores Lancados"
         Height          =   3225
         Left            =   -74760
         TabIndex        =   77
         Top             =   1320
         Width           =   9060
         Begin VB.TextBox Txt_TotalReceita 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   7230
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   2820
            Width           =   1395
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_TributoReceita 
            Height          =   2160
            Left            =   120
            TabIndex        =   78
            Top             =   300
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   3810
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Pkid Receita"
            Columns(0).DataField=   "pkid"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tributo Receita"
            Columns(1).DataField=   "strReceita"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Valor "
            Columns(2).DataField=   "dblValorReceita"
            Columns(2).NumberFormat=   "Standard"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerStyle=   7
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=1058"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=11959"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerStyle=0"
            Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=11906"
            Splits(0)._ColumnProps(13)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(16)=   "Column(2).Width=3069"
            Splits(0)._ColumnProps(17)=   "Column(2).DividerStyle=0"
            Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=3016"
            Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            BorderStyle     =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   -2147483633
            RowDividerColor =   12648447
            RowSubDividerColor=   13160660
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000004&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000004&"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1,.transparentBmp=-1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=59,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=47,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=48"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=49"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=59"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=47,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=48"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=49"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=59"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=16,.parent=47,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=48"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=49"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=59"
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
            _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=39:EvenRow"
            _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(63)  =   "Named:id=40:OddRow"
            _StyleDefs(64)  =   ":id=40,.parent=33"
            _StyleDefs(65)  =   "Named:id=41:RecordSelector"
            _StyleDefs(66)  =   ":id=41,.parent=34"
            _StyleDefs(67)  =   "Named:id=42:FilterBar"
            _StyleDefs(68)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lbl_Totalreceita 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Left            =   6660
            TabIndex        =   80
            Top             =   2610
            Width           =   360
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ISS"
         Height          =   2355
         Left            =   -74775
         TabIndex        =   62
         Top             =   3150
         Width           =   9060
         Begin VB.TextBox txt_ListaServico 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   585
            Left            =   1350
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   66
            Top             =   630
            Width           =   7545
         End
         Begin VB.Label lbl_STRMOEDAESTIMADAISS 
            Height          =   195
            Left            =   7650
            TabIndex        =   76
            Top             =   1680
            Width           =   1260
         End
         Begin VB.Label lbl_EstimativaMoeda 
            AutoSize        =   -1  'True
            Caption         =   "Estimativa Moeda :"
            Height          =   195
            Left            =   6300
            TabIndex        =   75
            Top             =   1680
            Width           =   1350
         End
         Begin VB.Label lbl_DTMDATAESTIMATIVAISS 
            Height          =   195
            Left            =   4905
            TabIndex        =   74
            Top             =   1680
            Width           =   1260
         End
         Begin VB.Label lbl_DataEstimativa 
            AutoSize        =   -1  'True
            Caption         =   "Data Estimativa :"
            Height          =   195
            Left            =   3645
            TabIndex        =   73
            Top             =   1680
            Width           =   1200
         End
         Begin VB.Label lbl_DBLVALORESTIMADOISS 
            Height          =   195
            Left            =   1350
            TabIndex        =   72
            Top             =   1680
            Width           =   2160
         End
         Begin VB.Label lbl_EstimativaValor 
            AutoSize        =   -1  'True
            Caption         =   "Estimativa Valor :"
            Height          =   195
            Left            =   90
            TabIndex        =   71
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lbl_DBLPORCENTAGEMISSVAR 
            Height          =   195
            Left            =   4770
            TabIndex        =   70
            Top             =   1380
            Width           =   1260
         End
         Begin VB.Label lbl_Mensal 
            AutoSize        =   -1  'True
            Caption         =   "ISS Mensal % :"
            Height          =   195
            Left            =   3645
            TabIndex        =   69
            Top             =   1380
            Width           =   1065
         End
         Begin VB.Label lbl_DBLVALORISSFIXO 
            Height          =   195
            Left            =   1350
            TabIndex        =   68
            Top             =   1380
            Width           =   2160
         End
         Begin VB.Label lbl_VFixo 
            AutoSize        =   -1  'True
            Caption         =   "ISS Fixo :"
            Height          =   195
            Left            =   630
            TabIndex        =   67
            Top             =   1380
            Width           =   675
         End
         Begin VB.Label lbl_lista 
            AutoSize        =   -1  'True
            Caption         =   "Lista Serviço :"
            Height          =   195
            Left            =   300
            TabIndex        =   65
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label lbl_STRTIPOISS 
            Height          =   195
            Left            =   1350
            TabIndex        =   64
            Top             =   300
            Width           =   7500
         End
         Begin VB.Label lbl_ISS 
            AutoSize        =   -1  'True
            Caption         =   "ISS - Tipo :"
            Height          =   195
            Left            =   510
            TabIndex        =   63
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Frame fra_publicidades 
         Caption         =   "Publicidades"
         Height          =   1815
         Left            =   -74775
         TabIndex        =   61
         Top             =   1320
         Width           =   9060
         Begin TrueOleDBGrid70.TDBGrid tdb_Publicidades 
            Height          =   1500
            Left            =   90
            TabIndex        =   35
            Top             =   210
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   2646
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "intPublicidade"
            Columns(0).DataField=   "Pkid"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tipo de Publicidade"
            Columns(1).DataField=   "STRTIPOPUBLICIDADE"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Quantidade"
            Columns(2).DataField=   "INTQUANTIDADE"
            Columns(2).NumberFormat=   "Standard"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Área"
            Columns(3).DataField=   "dblArea"
            Columns(3).NumberFormat=   "Standard"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Valor"
            Columns(4).DataField=   "dblValor"
            Columns(4).NumberFormat=   "Standard"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Total"
            Columns(5).DataField=   "dblTotal"
            Columns(5).NumberFormat=   "Standard"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerStyle=   7
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=1058"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=7752"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerStyle=0"
            Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=7699"
            Splits(0)._ColumnProps(13)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(16)=   "Column(2).Width=1693"
            Splits(0)._ColumnProps(17)=   "Column(2).DividerStyle=0"
            Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1640"
            Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(23)=   "Column(3).Width=1455"
            Splits(0)._ColumnProps(24)=   "Column(3).DividerStyle=0"
            Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=1402"
            Splits(0)._ColumnProps(27)=   "Column(3).AllowSizing=0"
            Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(30)=   "Column(4).Width=1931"
            Splits(0)._ColumnProps(31)=   "Column(4).DividerStyle=0"
            Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=1879"
            Splits(0)._ColumnProps(34)=   "Column(4).AllowSizing=0"
            Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(37)=   "Column(5).Width=2302"
            Splits(0)._ColumnProps(38)=   "Column(5).DividerStyle=0"
            Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=2249"
            Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            BorderStyle     =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   -2147483633
            RowDividerColor =   12648447
            RowSubDividerColor=   13160660
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000004&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000004&"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1,.transparentBmp=-1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=59,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=47,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=48"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=49"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=59"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=47,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=48"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=49"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=59"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=16,.parent=47,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=48"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=49"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=59"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=74,.parent=47,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=48"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=49"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=59"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=78,.parent=47,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=48"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=49"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=59"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=24,.parent=47,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=48"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=49"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=59"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_feiras 
         Caption         =   "Feiras"
         Height          =   1335
         Left            =   -74775
         TabIndex        =   60
         Top             =   5520
         Width           =   9060
         Begin TrueOleDBGrid70.TDBGrid tdb_Feiras 
            Height          =   1050
            Left            =   90
            TabIndex        =   36
            Top             =   210
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   1852
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "intFeira"
            Columns(0).DataField=   "Pkid"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Feira"
            Columns(1).DataField=   "strFeira"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tipo Feira"
            Columns(2).DataField=   "strTipoFeira"
            Columns(2).NumberFormat=   "Standard"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Área"
            Columns(3).DataField=   "dblArea"
            Columns(3).NumberFormat=   "Standard"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Box"
            Columns(4).DataField=   "strnrbox"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Valor"
            Columns(5).DataField=   "dblValor"
            Columns(5).NumberFormat=   "Standard"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerStyle=   7
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=1058"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=3757"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerStyle=0"
            Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=3704"
            Splits(0)._ColumnProps(13)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(16)=   "Column(2).Width=4789"
            Splits(0)._ColumnProps(17)=   "Column(2).DividerStyle=0"
            Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=4736"
            Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=0"
            Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(23)=   "Column(3).Width=1746"
            Splits(0)._ColumnProps(24)=   "Column(3).DividerStyle=0"
            Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=1693"
            Splits(0)._ColumnProps(27)=   "Column(3).AllowSizing=0"
            Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(30)=   "Column(4).Width=1535"
            Splits(0)._ColumnProps(31)=   "Column(4).DividerStyle=0"
            Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=1482"
            Splits(0)._ColumnProps(34)=   "Column(4).AllowSizing=0"
            Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(37)=   "Column(5).Width=2990"
            Splits(0)._ColumnProps(38)=   "Column(5).DividerStyle=0"
            Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=2937"
            Splits(0)._ColumnProps(41)=   "Column(5).AllowSizing=0"
            Splits(0)._ColumnProps(42)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            BorderStyle     =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   -2147483633
            RowDividerColor =   12648447
            RowSubDividerColor=   13160660
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000004&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000004&"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1,.transparentBmp=-1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=59,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=47,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=48"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=49"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=59"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=47,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=48"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=49"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=59"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=16,.parent=47,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=48"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=49"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=59"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=74,.parent=47,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=48"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=49"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=59"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=78,.parent=47,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=48"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=49"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=59"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=47,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=48"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=49"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=59"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_Parcelas 
         Caption         =   "Parcelas"
         Height          =   2385
         Left            =   -74775
         TabIndex        =   59
         Top             =   4530
         Width           =   9060
         Begin TrueOleDBGrid70.TDBGrid tdb_Parcelas 
            Height          =   2100
            Left            =   120
            TabIndex        =   28
            Top             =   210
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   3704
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Parcela"
            Columns(0).DataField=   "intParcela"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   "strMoeda"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Acordo"
            Columns(2).DataField=   "strAcordo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Valor"
            Columns(3).DataField=   "dblValor"
            Columns(3).NumberFormat=   "Standard"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Vencimento"
            Columns(4).DataField=   "dtmDtVencimento"
            Columns(4).NumberFormat=   "FormatText Event"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "D.A."
            Columns(5).DataField=   "intLancamentoAlfaDAtiva"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Baixa"
            Columns(6).DataField=   "dtmDtPagamento"
            Columns(6).NumberFormat=   "FormatText Event"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Descrição da Baixa"
            Columns(7).DataField=   "STRDESCRICAO"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Observação"
            Columns(8).DataField=   "Strobservacao"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerStyle=   7
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=1058"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=767"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerStyle=0"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=714"
            Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=2"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=2461"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerStyle=0"
            Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=2408"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(21)=   "Column(3).Width=1879"
            Splits(0)._ColumnProps(22)=   "Column(3).DividerStyle=0"
            Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1826"
            Splits(0)._ColumnProps(25)=   "Column(3).AllowSizing=0"
            Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(28)=   "Column(4).Width=1746"
            Splits(0)._ColumnProps(29)=   "Column(4).DividerStyle=0"
            Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1693"
            Splits(0)._ColumnProps(32)=   "Column(4).AllowSizing=0"
            Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(35)=   "Column(5).Width=635"
            Splits(0)._ColumnProps(36)=   "Column(5).DividerStyle=0"
            Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=582"
            Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(40)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(41)=   "Column(6).Width=1826"
            Splits(0)._ColumnProps(42)=   "Column(6).DividerStyle=0"
            Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=1773"
            Splits(0)._ColumnProps(45)=   "Column(6).AllowSizing=0"
            Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=1"
            Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(48)=   "Column(7).Width=2963"
            Splits(0)._ColumnProps(49)=   "Column(7).DividerStyle=0"
            Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=2910"
            Splits(0)._ColumnProps(52)=   "Column(7).AllowSizing=0"
            Splits(0)._ColumnProps(53)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(54)=   "Column(8).Width=3360"
            Splits(0)._ColumnProps(55)=   "Column(8).DividerStyle=0"
            Splits(0)._ColumnProps(56)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(8)._WidthInPix=3307"
            Splits(0)._ColumnProps(58)=   "Column(8).AllowSizing=0"
            Splits(0)._ColumnProps(59)=   "Column(8).Order=9"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            BorderStyle     =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   -2147483633
            RowDividerColor =   12648447
            RowSubDividerColor=   13160660
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000004&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000004&"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1,.transparentBmp=-1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=59,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=47,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=48"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=49"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=59"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=47,.alignment=1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=48"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=49"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=59"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=47,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=48"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=49"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=59"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=16,.parent=47,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=13,.parent=48"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=14,.parent=49"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=15,.parent=59"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=47,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=48"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=49"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=59"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=24,.parent=47,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=48"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=49"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=59"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=47,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=48"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=49"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=59"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=82,.parent=47"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=48"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=49"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=59"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=86,.parent=47"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=83,.parent=48"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=84,.parent=49"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=85,.parent=59"
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
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_enquadramentos 
         Caption         =   "Atividades e Enquadramentos - Tributo"
         Height          =   1095
         Left            =   188
         TabIndex        =   56
         Top             =   5850
         Width           =   9090
         Begin TrueOleDBGrid70.TDBGrid tdb_Tributos 
            Height          =   810
            Left            =   90
            TabIndex        =   22
            Top             =   210
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   1429
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "pkid"
            Columns(0).DataField=   "pkid"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Atividade"
            Columns(1).DataField=   "Strdescricaoatividade"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tipo de Tributo"
            Columns(2).DataField=   "STRTIPOTRIBUTO"
            Columns(2).NumberFormat=   "Short Date"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tributo"
            Columns(3).DataField=   "STRTRIBUTO"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=4842"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4763"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=5345"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=5265"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=4842"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=4763"
            Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
            _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
      End
      Begin VB.Frame fra_Lancamentos 
         Caption         =   "Atividades e Enquadramentos Lançados"
         Height          =   1095
         Left            =   225
         TabIndex        =   55
         Top             =   4740
         Width           =   9090
         Begin TrueOleDBGrid70.TDBGrid tdb_Atividade 
            Height          =   810
            Left            =   90
            TabIndex        =   21
            Top             =   210
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   1429
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "pkid"
            Columns(0).DataField=   "pkid"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "INTCODIGOATIVIDADE"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Atividade"
            Columns(2).DataField=   "STRDESCRICAOATIVIDADE"
            Columns(2).NumberFormat=   "Short Date"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   512
            Columns(3)._MaxComboItems=   5
            Columns(3).ValueItems(0)._DefaultItem=   0
            Columns(3).ValueItems(0).Value=   "0"
            Columns(3).ValueItems(0).Value.vt=   8
            Columns(3).ValueItems(0).DisplayValue=   "Secundária"
            Columns(3).ValueItems(0).DisplayValue.vt=   8
            Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems(1)._DefaultItem=   0
            Columns(3).ValueItems(1).Value=   "1"
            Columns(3).ValueItems(1).Value.vt=   8
            Columns(3).ValueItems(1).DisplayValue=   "Primária"
            Columns(3).ValueItems(1).DisplayValue.vt=   8
            Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems.Count=   2
            Columns(3).Caption=   "Principal / Secundária"
            Columns(3).DataField=   "BLNPRINCIPAL"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1931"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1852"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=10001"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=9922"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=3016"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2937"
            Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
            _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
      End
      Begin VB.Frame fra_socios 
         Caption         =   "Sócios"
         Height          =   1155
         Left            =   225
         TabIndex        =   54
         Top             =   3570
         Width           =   9090
         Begin TrueOleDBGrid70.TDBGrid tdb_Socios 
            Height          =   810
            Left            =   90
            TabIndex        =   20
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   1429
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Pkid"
            Columns(0).DataField=   "Pkid"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Sócio"
            Columns(1).DataField=   "strNome"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Cotas"
            Columns(2).DataField=   "intnumerocotas"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=11774"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=11695"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=3254"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3175"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
            _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
      Begin VB.Frame fra_Cabecalho 
         Enabled         =   0   'False
         Height          =   990
         Index           =   2
         Left            =   -74775
         TabIndex        =   49
         Top             =   330
         Width           =   9090
         Begin VB.TextBox txtdtmdtcancelamento2 
            Height          =   315
            Left            =   6015
            TabIndex        =   100
            Top             =   555
            Width           =   1080
         End
         Begin VB.TextBox txtstrComposicaoDaReceita3 
            Height          =   315
            Left            =   1575
            TabIndex        =   34
            Top             =   555
            Width           =   3150
         End
         Begin VB.TextBox txtstrEmissao3 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6015
            MaxLength       =   4
            TabIndex        =   32
            Top             =   165
            Width           =   570
         End
         Begin VB.TextBox txtstrNumDoAviso3 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7140
            MaxLength       =   6
            TabIndex        =   33
            Top             =   165
            Width           =   975
         End
         Begin MSMask.MaskEdBox mskstrInscricao3 
            Height          =   300
            Left            =   2220
            TabIndex        =   30
            Top             =   165
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo dbcintExercicio3 
            Height          =   315
            Left            =   4260
            TabIndex        =   31
            Top             =   165
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbldtCancelamento2 
            AutoSize        =   -1  'True
            Caption         =   "Cancelamento"
            Height          =   195
            Left            =   4950
            TabIndex        =   101
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label lblintComposicao3 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   585
            TabIndex        =   58
            Top             =   630
            Width           =   870
         End
         Begin VB.Label lbl_Exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Index           =   2
            Left            =   3540
            TabIndex        =   53
            Top             =   255
            Width           =   675
         End
         Begin VB.Label lbl_strInscricaoAnterior 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Index           =   2
            Left            =   750
            TabIndex        =   52
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label lbl_Emissao 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Index           =   2
            Left            =   5385
            TabIndex        =   51
            Top             =   255
            Width           =   585
         End
         Begin VB.Label lbl_Aviso 
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Index           =   2
            Left            =   6690
            TabIndex        =   50
            Top             =   255
            Width           =   390
         End
      End
      Begin VB.Frame fra_Cabecalho 
         Enabled         =   0   'False
         Height          =   990
         Index           =   1
         Left            =   -74775
         TabIndex        =   44
         Top             =   330
         Width           =   9090
         Begin VB.TextBox txtdtmdtcancelamento1 
            Height          =   315
            Left            =   6015
            TabIndex        =   98
            Top             =   555
            Width           =   1080
         End
         Begin VB.TextBox txtstrComposicaoDaReceita2 
            Height          =   315
            Left            =   1575
            TabIndex        =   27
            Top             =   555
            Width           =   3150
         End
         Begin VB.TextBox txtstrNumDoAviso2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7140
            MaxLength       =   6
            TabIndex        =   26
            Top             =   165
            Width           =   975
         End
         Begin VB.TextBox txtstrEmissao2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6015
            MaxLength       =   4
            TabIndex        =   25
            Top             =   165
            Width           =   570
         End
         Begin MSMask.MaskEdBox mskstrInscricao2 
            Height          =   300
            Left            =   2220
            TabIndex        =   23
            Top             =   165
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo dbcintExercicio2 
            Height          =   315
            Left            =   4260
            TabIndex        =   24
            Top             =   165
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbldtCancelamento1 
            AutoSize        =   -1  'True
            Caption         =   "Cancelamento"
            Height          =   195
            Left            =   4950
            TabIndex        =   99
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label lblintComposicao2 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   585
            TabIndex        =   57
            Top             =   630
            Width           =   870
         End
         Begin VB.Label lbl_Aviso 
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Index           =   1
            Left            =   6690
            TabIndex        =   48
            Top             =   255
            Width           =   390
         End
         Begin VB.Label lbl_Emissao 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Index           =   1
            Left            =   5385
            TabIndex        =   47
            Top             =   255
            Width           =   585
         End
         Begin VB.Label lbl_strInscricaoAnterior 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Index           =   1
            Left            =   750
            TabIndex        =   46
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label lbl_Exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Index           =   1
            Left            =   3540
            TabIndex        =   45
            Top             =   255
            Width           =   675
         End
      End
      Begin VB.Frame fra_Cabecalho 
         Height          =   990
         Index           =   0
         Left            =   225
         TabIndex        =   38
         Top             =   330
         Width           =   9090
         Begin VB.TextBox txtdtmdtCancelamento 
            Height          =   315
            Left            =   6015
            TabIndex        =   5
            Top             =   555
            Width           =   1080
         End
         Begin VB.TextBox txtstrComposicaoDaReceita 
            Height          =   315
            Left            =   1575
            TabIndex        =   4
            Top             =   555
            Width           =   3150
         End
         Begin VB.TextBox txtstrNumeroAviso 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7140
            MaxLength       =   10
            TabIndex        =   3
            Top             =   165
            Width           =   1125
         End
         Begin VB.TextBox txtstrEmissao 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6015
            MaxLength       =   4
            TabIndex        =   2
            Top             =   165
            Width           =   570
         End
         Begin MSMask.MaskEdBox mskstrInscricao 
            Height          =   300
            Left            =   2220
            TabIndex        =   0
            Top             =   165
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo dbcintExercicio 
            Height          =   315
            Left            =   4260
            TabIndex        =   1
            Top             =   165
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbldtCancelamento 
            AutoSize        =   -1  'True
            Caption         =   "Cancelamento"
            Height          =   195
            Left            =   4950
            TabIndex        =   97
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label lblintComposicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   585
            TabIndex        =   43
            Top             =   660
            Width           =   870
         End
         Begin VB.Label lbl_Aviso 
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Index           =   0
            Left            =   6690
            TabIndex        =   42
            Top             =   255
            Width           =   390
         End
         Begin VB.Label lbl_Emissao 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Index           =   0
            Left            =   5385
            TabIndex        =   41
            Top             =   255
            Width           =   585
         End
         Begin VB.Label lbl_strInscricaoAnterior 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Index           =   0
            Left            =   750
            TabIndex        =   40
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label lbl_Exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Index           =   0
            Left            =   3540
            TabIndex        =   39
            Top             =   255
            Width           =   675
         End
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6210
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1740
      Left            =   90
      TabIndex        =   102
      Top             =   7110
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   3069
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
      Columns(1).Caption=   "Inscricao"
      Columns(1).DataField=   "inscricao"
      Columns(1).NumberFormat=   "FormatText Event"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Exercício"
      Columns(2).DataField=   "Exercicio"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Aviso"
      Columns(3).DataField=   "NumeroAviso"
      Columns(3).NumberFormat=   "FormatText Event"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Razão Social"
      Columns(4).DataField=   "Proprietario"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "intLancamentoAlfa"
      Columns(5).DataField=   "intLancamentoAlfa"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Composição da Receita"
      Columns(6).DataField=   "strComposicaoDaReceita"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2196"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2117"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1376"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1296"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1296"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1217"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=5186"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=5106"
      Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(35)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=6112"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=6033"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bgcolor=&H80000009&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.namedParent=38"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
      _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadLancamentoEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando     As Boolean
    Dim mobjAux           As Object
    Dim mblnSelecionou    As Boolean
    Dim mblnClickOk       As Boolean
    Dim blnOrdenacaoAsc   As Boolean

Private Function strQueryAplicar() As String
    Dim strSQL  As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao FROM "
    strSQL = strSQL & gstrBairro & " ORDER BY strDescricao"
    
    strQueryAplicar = strSQL
    
End Function

Private Sub dbcintExercicio_GotFocus()
    MarcaCampo dbcintExercicio
End Sub

Private Sub dbcintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintExercicio
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1190
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
    TrocaCorObjeto txtStrnomeproprietario, True
    TrocaCorObjeto txtSTRNOMEFANTASIA, True
    TrocaCorObjeto txtStratividadebasica, True
    TrocaCorObjeto txtStrnaturezajuridica, True
    TrocaCorObjeto txtDtmdataabertura, True
    TrocaCorObjeto txtDblareaocupada, True
    TrocaCorObjeto txtDblnumeroempregados, True
    TrocaCorObjeto txtstrLogradouro, True
    TrocaCorObjeto txtstrNumero, True
    TrocaCorObjeto txtstrComplemento, True
    TrocaCorObjeto txtstrBairro, True
    TrocaCorObjeto txtstrMunicipio, True
    TrocaCorObjeto txtstrUf, True
    TrocaCorObjeto txtintCep, True
    VerificaMascaraInscricao
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub mskstrInscricao_GotFocus()
    MarcaCampo mskstrInscricao
End Sub

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao
End Sub

Private Sub tdb_Lista_Click()
  mblnClickOk = True
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 1 Then
        Value = gstrFormataInscricao(CStr(Value), TYP_ECONOMICA)
    End If
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            LimpaISS
            LimpaEconomico
            Screen.MousePointer = vbHourglass
            mblnClickOk = False
            mblnSelecionou = True
            mblnAlterando = True
            txtPKId.Text = .Columns("PKID").Value
            gCorLinhaSelecionada tdb_Lista
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            PreencheCabecalho
            LeDaTabelaParaObj "", tdb_Socios, strQuerySocios
            LeDaTabelaParaObj "", tdb_Atividade, strQueryAtividade
            LeDaTabelaParaObj "", tdb_Tributos, strQueryAtividadeTributoTipo
            LeDaTabelaParaObj "", tdb_TributoReceita, strQueryReceita
            PreencheTotalReceita
            LeDaTabelaParaObj "", tdb_Parcelas, strQueryParcelas
            LeDaTabelaParaObj "", tdb_Feiras, strQueryFeiras
            LeDaTabelaParaObj "", tdb_Publicidades, strQueryPublicidade
            PreencheISS
            Screen.MousePointer = vbDefault
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrLocalizar)
            mblnClickOk = True
            LeDaTabelaParaObj "", tdb_Lista, strQueryLocalizar(False)
        Case Is = UCase(gstrRefresh)
            LeDaTabelaParaObj "", tdb_Lista, strQueryLocalizar(True)
        Case Is = UCase(gstrNovo)
            tab_3dPasta.Tab = 0
            LimpaObjeto Me
            Txt_TotalReceita.Text = ""
            LimpaGrids
            LimpaISS
    End Select
End Sub

Function strQueryRelatorio() As String
    Dim strSQL As String
    
    strSQL = ""
    
    strSQL = strSQL & "SELECT * FROM "
    
    strQueryRelatorio = strSQL
   
End Function

Private Sub PreencheCabecalho()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = "SELECT " & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " Inscricao, "
    strSQL = strSQL & "LA.intExercicio Exercicio, "
    strSQL = strSQL & "LA.strEmissao Emissao, "
    strSQL = strSQL & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " NumeroAviso,"
    strSQL = strSQL & "LA.strComposicaoDaReceita ComposicaoDaReceita, "
    strSQL = strSQL & "LA.strOcorrencia Ocorrencia, "
    strSQL = strSQL & "LA.Strlogradouro, la.strnumero, la.strcomplemento, la.strbairro, la.strmunicipio, "
    strSQL = strSQL & "lA.struf, la.intcep,"
    strSQL = strSQL & "LA.Strnomeproprietario, "
    strSQL = strSQL & "LA.Dtmdtcancelamento, "
    strSQL = strSQL & "LE.STRNOMEFANTASIA, "
    strSQL = strSQL & "LE.Stratividadebasica, "
    strSQL = strSQL & "LE.Strnaturezajuridica, "
    strSQL = strSQL & "LE.Dtmdataabertura, "
    strSQL = strSQL & "LE.Dblareaocupada, "
    strSQL = strSQL & "LE.Dblnumeroempregados "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLancamentoEconomico & " LE "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "LA.Pkid = LE.intlancamentoalfa AND "
    strSQL = strSQL & "LA.intUtilizacao = " & TYP_ECONOMICA & " AND "
    strSQL = strSQL & "LE.Pkid = " & txtPKId.Text
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                mskstrInscricao.Text = gstrENulo(!Inscricao)
                mskstrInscricao2.Text = gstrENulo(!Inscricao)
                mskstrInscricao3.Text = gstrENulo(!Inscricao)
                dbcintExercicio.Text = gstrENulo(!EXERCICIO)
                dbcintExercicio2.Text = gstrENulo(!EXERCICIO)
                dbcintExercicio3.Text = gstrENulo(!EXERCICIO)
                txtstrEmissao.Text = gstrENulo(!Emissao)
                txtstrEmissao2.Text = gstrENulo(!Emissao)
                txtstrEmissao3.Text = gstrENulo(!Emissao)
                txtstrNumeroAviso.Text = gstrENulo(!NumeroAviso)
                txtstrNumDoAviso2.Text = gstrENulo(!NumeroAviso)
                txtstrNumDoAviso3.Text = gstrENulo(!NumeroAviso)
                txtstrComposicaoDaReceita.Text = gstrENulo(!ComposicaoDaReceita)
                txtstrComposicaoDaReceita2.Text = gstrENulo(!ComposicaoDaReceita)
                txtstrComposicaoDaReceita3.Text = gstrENulo(!ComposicaoDaReceita)
                txtdtmdtCancelamento.Text = gstrDataFormatada(gstrENulo(!Dtmdtcancelamento))
                txtdtmdtcancelamento1.Text = gstrDataFormatada(gstrENulo(!Dtmdtcancelamento))
                txtdtmdtcancelamento2.Text = gstrDataFormatada(gstrENulo(!Dtmdtcancelamento))
                txtStrnomeproprietario.Text = gstrENulo(!strnomeproprietario)
                txtSTRNOMEFANTASIA.Text = gstrENulo(!strNomeFantasia)
                txtStratividadebasica.Text = gstrENulo(!strAtividadeBasica)
                txtStrnaturezajuridica.Text = gstrENulo(!Strnaturezajuridica)
                txtDtmdataabertura.Text = gstrENulo(!dtmDataAbertura)
                txtDblareaocupada.Text = gstrConvVrDoSql(gstrENulo(!dblAreaOcupada))
                txtDblnumeroempregados.Text = gstrENulo(!Dblnumeroempregados)
                txtstrLogradouro.Text = gstrENulo(!strLogradouro)
                txtstrNumero.Text = gstrENulo(!strNumero)
                txtstrComplemento = gstrENulo(!STRCOMPLEMENTO)
                txtstrBairro = gstrENulo(!strBairro)
                txtstrMunicipio = gstrENulo(!STRMUNICIPIO)
                txtstrUf = gstrENulo(!STRUF)
                txtintCep = gstrENulo(!INTCEP)
            End If
        End With
    End If

End Sub

Private Function strQueryLocalizar(blnRefresh As Boolean) As String
    Dim strSQL As String
    
    strSQL = "SELECT LEA.Pkid Pkid,"
    strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " Inscricao, "
    strSQL = strSQL & "LA.intExercicio AS Exercicio, "
    strSQL = strSQL & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " NumeroAviso,"
    strSQL = strSQL & "LA.strNomeProprietario AS Proprietario, "
    strSQL = strSQL & "LA.strComposicaoDaReceita, "
    strSQL = strSQL & "LA.Pkid AS intLancamentoAlfa"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLancamentoEconomico & " LEA"
    strSQL = strSQL & " WHERE LA.Pkid = LEA.intLancamentoAlfa"
    
    If Not blnRefresh Then
        If Trim(mskstrInscricao.Text) <> "" Then
            strSQL = strSQL & " AND strInscricao LIKE " & "'" & UCase(String(gintLenInscricao - Len(mskstrInscricao), "0") & mskstrInscricao) & "'"
        End If
        If Trim(dbcintExercicio.Text) <> "" Then
            strSQL = strSQL & " AND intExercicio = " & UCase(dbcintExercicio.Text)
        End If
        If Trim(txtstrEmissao.Text) <> "" Then
            strSQL = strSQL & " AND strEmissao LIKE " & "'" & UCase(String(gintLenEmissao - Len(txtstrEmissao), "0") & txtstrEmissao) & "'"
        End If
        If Trim(txtstrNumeroAviso.Text) <> "" Then
            strSQL = strSQL & " AND strNumeroAviso LIKE " & "'" & UCase(String(gintLenNumAviso - Len(txtstrNumeroAviso), "0") & txtstrNumeroAviso.Text) & "'"
        End If
        If Trim(txtstrComposicaoDaReceita.Text) <> "" Then
            strSQL = strSQL & " AND UPPER(strComposicaoDaReceita) LIKE " & "'" & UCase(txtstrComposicaoDaReceita.Text) & "%'"
        End If
        If Trim(txtdtmdtCancelamento.Text) <> "" Then
            strSQL = strSQL & " AND LA.Dtmdtcancelamento = " & gstrConvDtParaSql(txtdtmdtCancelamento) & " "
        End If
    End If
    
    strSQL = strSQL & " ORDER BY LA.strInscricao Asc, LA.strComposicaoDaReceita Asc, LA.intExercicio Asc"

    strQueryLocalizar = strSQL

End Function


Private Function strQuerySocios() As String
    Dim strSQL          As String
    
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "Pkid, "
    strSQL = strSQL & "strnome, "
    strSQL = strSQL & "intnumerocotas "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLctEconomicoSocio
    strSQL = strSQL & " Where "
    strSQL = strSQL & "intlancamentoeconomico = " & txtPKId.Text
    strQuerySocios = strSQL
    
End Function

Private Function strQueryAtividade() As String
    Dim strSQL As String
    
    strSQL = "Select "
    strSQL = strSQL & "LEA.INTCODIGOATIVIDADE, "
    strSQL = strSQL & "LEA.STRDESCRICAOATIVIDADE, "
    strSQL = strSQL & "LEA.blnPrincipal "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLctEconomicoAtividade & " LEA "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "LEA.Intlancamentoeconomico = " & txtPKId.Text
    strSQL = strSQL & " Order By LEA.blnPrincipal Desc"
    strQueryAtividade = strSQL

End Function

Private Function strQueryAtividadeTributoTipo() As String
    Dim strSQL As String
        
    strSQL = "Select "
    strSQL = strSQL & "LET.Pkid, "
    strSQL = strSQL & "LEA.Strdescricaoatividade, "
    strSQL = strSQL & "LET.STRTIPOTRIBUTO, "
    strSQL = strSQL & "LET.STRTRIBUTO "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLancamentoEconomico & " LE, "
    strSQL = strSQL & gstrLctEconomicoAtividade & " LEA, "
    strSQL = strSQL & gstrLctEconomicoTributo & " LET "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "LE.Pkid = LEA.INTLANCAMENTOECONOMICO AND "
    strSQL = strSQL & "LEA.Pkid = LET.INTLCTECONOMICOATIVIDADE AND "
    strSQL = strSQL & "LE.Pkid = " & txtPKId.Text & " "
    strSQL = strSQL & "Order by LEA.Strdescricaoatividade, LET.STRTIPOTRIBUTO, LET.STRTRIBUTO "
    
    strQueryAtividadeTributoTipo = strSQL

End Function

Private Function strQueryParcelas() As String
    Dim strSQL As String
    
    strSQL = "SELECT LV.intParcela,"
    strSQL = strSQL & " LV.dblValor,"
    strSQL = strSQL & "CASE WHEN LV.intLancamentoAlfaDAtiva IS NULL THEN '' ELSE 'X' END intLancamentoAlfaDAtiva ,"
    strSQL = strSQL & " M.Strabreviatura as strMoeda, "
    strSQL = strSQL & " LV.dtmDtVencimento,"
    strSQL = strSQL & " LP.dtmDtPagamento,"
    strSQL = strSQL & "CB.STRDESCRICAO, "
    strSQL = strSQL & "LP.Strobservacao, "
    strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ACORDO)) & " strAcordo "
    
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrCodigoDeBaixa & " CB, "
        strSQL = strSQL & gstrMoedas & " M, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP, "
        strSQL = strSQL & gstrLancamentoAlfa & " LA "
        strSQL = strSQL & " WHERE LV.Pkid " & strOUTJSQLServer & "=" & " LP.intLancamentoValor " & strOUTJOracle
        strSQL = strSQL & " AND CB.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " LP.Intcodigobaixa  AND"
        strSQL = strSQL & " M.Pkid = LV.Intmoeda AND"
        strSQL = strSQL & " LV.intLancamentoAlfa = " & tdb_Lista.Columns("intLancamentoAlfa").Value & " AND "
        strSQL = strSQL & " lv.intlancamentoalfaacordo " & strOUTJSQLServer & "= LA.pkid " & strOUTJOracle
    ElseIf (bytDBType = EDatabases.SQLServer) Then
        strSQL = strSQL & " FROM " & gstrMoedas & " M INNER JOIN "
        strSQL = strSQL & gstrLancamentoValor & " LV ON M.PKID = LV.INTMOEDA LEFT OUTER JOIN "
        strSQL = strSQL & gstrLancamentoAlfa & " LA ON LV.intLancamentoAlfaAcordo = LA.PKId LEFT OUTER JOIN "
        strSQL = strSQL & gstrCodigoDeBaixa & " CB RIGHT OUTER JOIN "
        strSQL = strSQL & gstrLancamentoPagamento & " LP ON CB.PKID = LP.INTCODIGOBAIXA ON "
        strSQL = strSQL & " LV.Pkid = LP.intLancamentoValor "
        strSQL = strSQL & " WHERE LV.intLancamentoAlfa = " & tdb_Lista.Columns("intLancamentoAlfa").Value
    End If
    
    strSQL = strSQL & " ORDER BY LV.intParcela"
        
    strQueryParcelas = strSQL
    
End Function

Private Function strQueryFeiras() As String
    Dim strSQL As String
    
    strSQL = "Select "
    strSQL = strSQL & "LEF.PKID, "
    strSQL = strSQL & "LEF.STRFEIRA, "
    strSQL = strSQL & "LEF.STRTIPOFEIRA, "
    strSQL = strSQL & "LEF.DBLAREA, "
    strSQL = strSQL & "LEF.Strnrbox, "
    strSQL = strSQL & "LEF.dblValor "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoEconomico & " LE, "
    strSQL = strSQL & gstrLctEconomicoFeira & " LEF "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "LE.PKID = LEF.INTLANCAMENTOECONOMICO AND "
    strSQL = strSQL & "LE.Pkid = " & txtPKId & " "
    strSQL = strSQL & "ORDER BY LEF.STRFEIRA"
        
    strQueryFeiras = strSQL

End Function

Private Function strQueryPublicidade() As String
    Dim strSQL As String
    
    strSQL = "Select "
    strSQL = strSQL & "LEP.PKID, "
    strSQL = strSQL & "LEP.STRTIPOPUBLICIDADE, "
    strSQL = strSQL & "LEP.INTQUANTIDADE, "
    strSQL = strSQL & "LEP.DBLAREA, "
    strSQL = strSQL & "LEP.STROBSERVACAO, "
    strSQL = strSQL & "LEP.dblValor, "
    strSQL = strSQL & "(LEP.INTQUANTIDADE * LEP.DBLAREA * LEP.dblValor) dblTotal "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoEconomico & " LE, "
    strSQL = strSQL & gstrLctEconPublicidade & " LEP "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "LE.PKID = LEP.INTLANCAMENTOECONOMICO AND "
    strSQL = strSQL & "LE.Pkid = " & txtPKId & " "
    strSQL = strSQL & "ORDER BY LEP.STRTIPOPUBLICIDADE"
        
    strQueryPublicidade = strSQL

End Function

Private Function PreencheISS() As String
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = "Select "
    strSQL = strSQL & "LEI.DBLVALORESTIMADOISS, "
    strSQL = strSQL & "LEI.DTMDATAESTIMATIVAISS, "
    strSQL = strSQL & "LEI.STRMOEDAESTIMADAISS, "
    strSQL = strSQL & "LEI.STRTIPOISS, "
    strSQL = strSQL & "LEI.INTCODIGOLISTASERVICO, "
    strSQL = strSQL & "LEI.STRDESCRICAOLISTASERVICO, "
    strSQL = strSQL & "LEI.DBLVALORISSFIXO, "
    strSQL = strSQL & "LEI.DBLPORCENTAGEMISSVAR "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoEconomico & " LE, "
    strSQL = strSQL & gstrLancamentoEconIss & " LEI "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "LE.PKID = LEI.INTLANCAMENTOECONOMICO AND "
    strSQL = strSQL & "LE.Pkid = " & txtPKId & " "
            
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
            lbl_STRTIPOISS = gstrENulo(!STRTIPOISS)
            txt_ListaServico = IIf(Trim(gstrENulo(!INTCODIGOLISTASERVICO)) <> "", gstrENulo(!INTCODIGOLISTASERVICO) & "/", "") & gstrENulo(!STRDESCRICAOLISTASERVICO)
            lbl_DBLVALORISSFIXO = gstrConvVrDoSql(gstrENulo(!DBLVALORISSFIXO), , , True)
            lbl_DBLPORCENTAGEMISSVAR = gstrConvVrDoSql(gstrENulo(!DBLPORCENTAGEMISSVAR), , , True)
            lbl_DBLVALORESTIMADOISS = gstrConvVrDoSql(gstrENulo(!DBLVALORESTIMADOISS), , , True)
            lbl_DTMDATAESTIMATIVAISS = gstrDataFormatada(gstrENulo(!DTMDATAESTIMATIVAISS))
            lbl_STRMOEDAESTIMADAISS = gstrENulo(!STRMOEDAESTIMADAISS)
            End If
        End With
    End If
End Function

Private Sub LimpaGrids()
    Set tdb_Socios.DataSource = Nothing
    Set tdb_Atividade.DataSource = Nothing
    Set tdb_Tributos.DataSource = Nothing
    Set tdb_TributoReceita.DataSource = Nothing
    Set tdb_Parcelas.DataSource = Nothing
    Set tdb_Feiras.DataSource = Nothing
    Set tdb_Publicidades.DataSource = Nothing
    
End Sub
Private Sub LimpaISS()
    lbl_STRTIPOISS = ""
    txt_ListaServico = ""
    lbl_DBLVALORISSFIXO = ""
    lbl_DBLPORCENTAGEMISSVAR = ""
    lbl_DBLVALORESTIMADOISS = ""
    lbl_DTMDATAESTIMATIVAISS = ""
    lbl_STRMOEDAESTIMADAISS = ""
End Sub

Private Sub tdb_Parcelas_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 4 Or ColIndex = 6 Then
        Value = gstrDataFormatada(Value)
    End If
End Sub

Private Sub txtdtmDtCancelamento_GotFocus()
    MarcaCampo txtdtmdtCancelamento
End Sub

Private Sub txtdtmDtCancelamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdtCancelamento
End Sub

Private Sub txtdtmDtCancelamento_LostFocus()
    txtdtmdtCancelamento = gstrDataFormatada(txtdtmdtCancelamento)
End Sub

Private Sub txtstrComposicaoDaReceita_GotFocus()
    MarcaCampo txtstrComposicaoDaReceita
End Sub

Private Sub txtstrComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComposicaoDaReceita
End Sub

Private Sub txtstrEmissao_GotFocus()
    MarcaCampo txtstrEmissao
End Sub

Private Sub txtstrEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrEmissao
End Sub

Private Sub txtstrNumeroAviso_GotFocus()
    MarcaCampo txtstrNumeroAviso
End Sub

Private Sub txtstrNumeroAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNumeroAviso
End Sub


Private Sub VerificaMascaraInscricao()
Dim strSQL              As String
Dim adoResultado        As ADODB.Recordset
Dim strMascara          As String
Dim bytTamanhoMascara   As Integer
    
    strMascara = ""
    strSQL = ""
    strSQL = strSQL & "Select * From " & gstrCampoDeInscricao & " "
    strSQL = strSQL & "Where intTipoDeInscricao = " & TYP_ECONOMICA
    strSQL = strSQL & "Order By intSequencia"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                bytTamanhoMascara = bytTamanhoMascara + !intTamanho
                .MoveNext
            Loop
        End With
    End If
    mskstrInscricao.Mask = strMascara
    mskstrInscricao2.Mask = strMascara
    mskstrInscricao3.Mask = strMascara
    
End Sub

Private Sub LimpaEconomico()
    mskstrInscricao.Text = ""
    mskstrInscricao2.Text = ""
    mskstrInscricao3.Text = ""
    dbcintExercicio.Text = ""
    dbcintExercicio2.Text = ""
    dbcintExercicio3.Text = ""
    txtstrEmissao.Text = ""
    txtstrEmissao2.Text = ""
    txtstrEmissao3.Text = ""
    txtstrNumeroAviso.Text = ""
    txtstrNumDoAviso2.Text = ""
    txtstrNumDoAviso3.Text = ""
    txtstrComposicaoDaReceita.Text = ""
    txtstrComposicaoDaReceita2.Text = ""
    txtstrComposicaoDaReceita3.Text = ""
    txtStrnomeproprietario.Text = ""
    txtSTRNOMEFANTASIA.Text = ""
    txtStratividadebasica.Text = ""
    txtStrnaturezajuridica.Text = ""
    txtDtmdataabertura.Text = ""
    txtDblareaocupada.Text = ""
    txtDblnumeroempregados.Text = ""
    Txt_TotalReceita.Text = ""
    txtdtmdtCancelamento.Text = ""
    txtdtmdtcancelamento1.Text = ""
    txtdtmdtcancelamento2.Text = ""
End Sub

Private Function strQueryReceita() As String
    Dim strSQL As String
    
    strSQL = strSQL & "Select "
    strSQL = strSQL & "R.pkid, "
    strSQL = strSQL & "R.Strdescricao strReceita, "
    strSQL = strSQL & "Sum(" & gstrISNULL("LR.DblValor", "0") & ") as dblValorReceita "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLancamentoEconomico & " LE, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLancamentoValor & " LV, "
    strSQL = strSQL & gstrLancamentoReceita & " LR, "
    strSQL = strSQL & gstrReceita & " R "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "LA.Pkid = LE.Intlancamentoalfa And "
    strSQL = strSQL & "LA.Pkid = LV.Intlancamentoalfa And "
    strSQL = strSQL & "LV.Pkid = LR.Intlancamentovalor And "
    strSQL = strSQL & "R.Pkid = LR.Intreceita And "
    strSQL = strSQL & "LV.bitParcelaValida = 1 And "
    strSQL = strSQL & "LE.Pkid = " & txtPKId & " "
    strSQL = strSQL & "Group By "
    strSQL = strSQL & "R.pkid, "
    strSQL = strSQL & "R.Strdescricao, "
    strSQL = strSQL & "LR.dblvalor "
    strSQL = strSQL & "Order by "
    strSQL = strSQL & "strReceita "
    
    strQueryReceita = strSQL
    
End Function

Private Sub PreencheTotalReceita()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = strSQL & "Select "
    strSQL = strSQL & "Sum(" & gstrISNULL("LR.DblValor", "0") & ") as dblValorReceita "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLancamentoEconomico & " LE, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLancamentoValor & " LV, "
    strSQL = strSQL & gstrLancamentoReceita & " LR, "
    strSQL = strSQL & gstrReceita & " R "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "LA.Pkid = LE.Intlancamentoalfa And "
    strSQL = strSQL & "LA.Pkid = LV.Intlancamentoalfa And "
    strSQL = strSQL & "LV.Pkid = LR.Intlancamentovalor And "
    strSQL = strSQL & "R.Pkid = LR.Intreceita And "
    strSQL = strSQL & "LV.bitParcelaValida = 1 And "
    strSQL = strSQL & "LE.Pkid = " & txtPKId & " "
        
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Txt_TotalReceita.Text = gstrConvVrDoSql(gstrENulo(adoResultado!dblValorReceita), 2, , True)
        End If
    End If
        
End Sub
