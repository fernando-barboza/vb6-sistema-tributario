VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadExecutivosFiscais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Executivos Fiscais"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   2205
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   9615
   Begin VB.TextBox txtPkId 
      Height          =   315
      Left            =   6210
      TabIndex        =   97
      Top             =   30
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame fra_Executivo 
      Height          =   705
      Left            =   240
      TabIndex        =   59
      Top             =   480
      Width           =   9105
      Begin VB.TextBox txt_Serie 
         Height          =   285
         Left            =   8070
         TabIndex        =   117
         Top             =   270
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txt_NumDist 
         Height          =   285
         Left            =   6720
         TabIndex        =   115
         Top             =   270
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtintLoteExecutivo 
         Height          =   285
         Left            =   4650
         TabIndex        =   3
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox txtintNumeroProtocolo 
         Height          =   285
         Left            =   3240
         TabIndex        =   2
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox txtintNumSeq 
         Height          =   285
         Left            =   1230
         TabIndex        =   1
         Top             =   270
         Width           =   915
      End
      Begin VB.CheckBox chkbitDistribuicaoEletronica 
         Caption         =   "Distribuição Eletrônica"
         Height          =   195
         Left            =   5790
         TabIndex        =   4
         Top             =   330
         Width           =   1995
      End
      Begin VB.Label lbl_Serie 
         AutoSize        =   -1  'True
         Caption         =   "Série"
         Height          =   195
         Left            =   7650
         TabIndex        =   118
         Top             =   360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lbl_NumDist 
         AutoSize        =   -1  'True
         Caption         =   "Nº Distribuidor"
         Height          =   195
         Left            =   5640
         TabIndex        =   116
         Top             =   360
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lbl_intLoteExecutivo 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   4260
         TabIndex        =   62
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lbl_intNumeroProtocolo 
         AutoSize        =   -1  'True
         Caption         =   "Nº Protocolo"
         Height          =   195
         Left            =   2250
         TabIndex        =   61
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lbl_intNumSeq 
         AutoSize        =   -1  'True
         Caption         =   "Nº Seqüencial"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   1020
      End
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6675
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   11774
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Executivo"
      TabPicture(0)   =   "frmCadExecutivosFiscais.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tab_3dExecutado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_ValoresTotais"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_CartorioDistribuidor"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Débitos"
      TabPicture(1)   =   "frmCadExecutivosFiscais.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Parcelas"
      Tab(1).Control(1)=   "tdb_Debitos"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Envolvidos"
      TabPicture(2)   =   "frmCadExecutivosFiscais.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_Envolvidos"
      Tab(2).Control(1)=   "fra_EnderecoEnvolvidos(1)"
      Tab(2).Control(2)=   "tdb_Envolvidos"
      Tab(2).Control(3)=   "txt_tabtdb_Envolvidos"
      Tab(2).ControlCount=   4
      Begin VB.Frame fra_Envolvidos 
         Height          =   1185
         Left            =   -74820
         TabIndex        =   93
         Top             =   1200
         Width           =   9105
         Begin VB.TextBox txt_strEnvolvidoIdentidade 
            Height          =   315
            Left            =   4890
            MaxLength       =   50
            TabIndex        =   69
            Top             =   690
            Width           =   2340
         End
         Begin VB.TextBox txt_strEnvolvidoCPFCNPJ 
            Height          =   315
            Left            =   1170
            MaxLength       =   100
            TabIndex        =   68
            Top             =   690
            Width           =   2340
         End
         Begin VB.TextBox txt_strEnvolvidoNome 
            Height          =   315
            Left            =   1170
            MaxLength       =   100
            TabIndex        =   67
            Top             =   330
            Width           =   7410
         End
         Begin VB.Label Lab_el1 
            AutoSize        =   -1  'True
            Caption         =   "Identidade"
            Height          =   195
            Left            =   4080
            TabIndex        =   96
            Top             =   780
            Width           =   750
         End
         Begin VB.Label Lab_el43 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   675
            TabIndex        =   95
            Top             =   390
            Width           =   420
         End
         Begin VB.Label Lab_el44 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   330
            TabIndex        =   94
            Top             =   780
            Width           =   780
         End
      End
      Begin VB.Frame fra_EnderecoEnvolvidos 
         Caption         =   "Endereço"
         Height          =   1425
         Index           =   1
         Left            =   -74820
         TabIndex        =   66
         Top             =   2460
         Width           =   9105
         Begin VB.TextBox txt_strEnvolvidoTituloLogNotif 
            Height          =   300
            Left            =   2715
            TabIndex        =   71
            Top             =   270
            Width           =   795
         End
         Begin VB.TextBox txt_strEnvolvidoTipoLogNotif 
            Height          =   300
            Left            =   1185
            TabIndex        =   70
            Top             =   270
            Width           =   795
         End
         Begin VB.TextBox txt_strEnvolvidoBairroLogNotif 
            Height          =   300
            Left            =   5085
            MaxLength       =   50
            TabIndex        =   75
            Top             =   630
            Width           =   2460
         End
         Begin VB.TextBox txt_strEnvolvidoNomeLogNotif 
            Height          =   300
            Left            =   4590
            MaxLength       =   100
            TabIndex        =   72
            Top             =   270
            Width           =   3915
         End
         Begin VB.TextBox txt_strEnvolvidoComplLogNotif 
            Height          =   300
            Left            =   2730
            MaxLength       =   10
            TabIndex        =   74
            Top             =   630
            Width           =   1710
         End
         Begin VB.TextBox txt_strEnvolvidoNumLogNotif 
            Height          =   300
            Left            =   1170
            MaxLength       =   10
            TabIndex        =   73
            Top             =   630
            Width           =   825
         End
         Begin VB.TextBox txt_strEnvolvidoCidadeLogNotif 
            Height          =   300
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   76
            Top             =   990
            Width           =   2115
         End
         Begin VB.TextBox txt_strEnvolvidoUFLogNotif 
            Height          =   300
            Left            =   3630
            MaxLength       =   2
            TabIndex        =   77
            Top             =   990
            Width           =   375
         End
         Begin VB.TextBox txt_strEnvolvidoCEPLogNotif 
            Height          =   300
            Left            =   4425
            MaxLength       =   9
            TabIndex        =   78
            Top             =   975
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Título"
            Height          =   195
            Left            =   2250
            TabIndex        =   114
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   795
            TabIndex        =   113
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Lab_el51 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   3720
            TabIndex        =   87
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Lab_el50 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   4620
            TabIndex        =   86
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Lab_el49 
            AutoSize        =   -1  'True
            Caption         =   "N°"
            Height          =   195
            Left            =   930
            TabIndex        =   85
            Top             =   720
            Width           =   180
         End
         Begin VB.Label Lab_el48 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   2190
            TabIndex        =   84
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Lab_el47 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Left            =   4080
            TabIndex        =   83
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label Lab_el46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   405
            TabIndex        =   82
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Lab_el45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   3360
            TabIndex        =   81
            Top             =   1080
            Width           =   210
         End
      End
      Begin VB.Frame fra_Parcelas 
         Caption         =   "Parcelas"
         Height          =   2475
         Left            =   -74820
         TabIndex        =   63
         Top             =   3630
         Width           =   9120
         Begin TrueOleDBGrid70.TDBGrid tdb_Parcelas 
            Height          =   2220
            Left            =   30
            TabIndex        =   65
            Top             =   210
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   3916
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº"
            Columns(0).DataField=   "intParcela"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Vencimento"
            Columns(1).DataField=   "dtmDtVencimento"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Moeda"
            Columns(2).DataField=   "strAbreviatura"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Original"
            Columns(3).DataField=   "dblVlOriginal"
            Columns(3).NumberFormat=   "FormatText Event"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Principal"
            Columns(4).DataField=   "dblVlPrincipal"
            Columns(4).NumberFormat=   "FormatText Event"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Correção"
            Columns(5).DataField=   "dblVlCorrecao"
            Columns(5).NumberFormat=   "FormatText Event"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Multa"
            Columns(6).DataField=   "dblVlMulta"
            Columns(6).NumberFormat=   "FormatText Event"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Juros"
            Columns(7).DataField=   "dblVlJuros"
            Columns(7).NumberFormat=   "FormatText Event"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Total"
            Columns(8).DataField=   "dblVlTotal"
            Columns(8).NumberFormat=   "FormatText Event"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=1773"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1693"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=1244"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1164"
            Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=1"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=1931"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1852"
            Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=1931"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1852"
            Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(24)=   "Column(5).Width=1931"
            Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=1852"
            Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(29)=   "Column(6).Width=1931"
            Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=1852"
            Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(34)=   "Column(7).Width=1931"
            Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=1852"
            Splits(0)._ColumnProps(37)=   "Column(7)._ColStyle=2"
            Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(39)=   "Column(8).Width=1931"
            Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=1852"
            Splits(0)._ColumnProps(42)=   "Column(8)._ColStyle=2"
            Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=47"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=48"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=49"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=59"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=47"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=48"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=49"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=59"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=47,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=48"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=49"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=59"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=47,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=48"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=49"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=59"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=47,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=48"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=49"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=59"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=47,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=48"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=49"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=59"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=24,.parent=47,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=48"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=49"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=59"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=47,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=48"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=49"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=59"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=47,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=48"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=49"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=59"
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
      Begin VB.Frame fra_CartorioDistribuidor 
         Caption         =   "Cartório Distribuidor"
         Height          =   1785
         Left            =   180
         TabIndex        =   36
         Top             =   1170
         Width           =   9135
         Begin VB.CheckBox chkbitDistribuido 
            Caption         =   "Distribuido"
            Height          =   195
            Left            =   270
            TabIndex        =   5
            Top             =   330
            Width           =   1995
         End
         Begin VB.TextBox txtintFolhasDistribuidor 
            Height          =   285
            Left            =   7470
            TabIndex        =   10
            Top             =   600
            Width           =   915
         End
         Begin VB.TextBox txtintLivroDistribuidor 
            Height          =   285
            Left            =   5940
            TabIndex        =   9
            Top             =   600
            Width           =   915
         End
         Begin VB.TextBox txtdtmDtDistribuidor 
            Height          =   285
            Left            =   4260
            TabIndex        =   8
            Top             =   600
            Width           =   1155
         End
         Begin VB.TextBox txtstrSerieDistribuidor 
            Height          =   285
            Left            =   2730
            TabIndex        =   7
            Top             =   600
            Width           =   915
         End
         Begin VB.TextBox txtstrNumDistribuidor 
            Height          =   285
            Left            =   1230
            TabIndex        =   6
            Top             =   600
            Width           =   945
         End
         Begin VB.Frame fra_Oficio 
            Caption         =   "Ofício"
            Height          =   675
            Left            =   270
            TabIndex        =   37
            Top             =   990
            Width           =   8595
            Begin VB.TextBox txtintNumOficio 
               Height          =   285
               Left            =   990
               TabIndex        =   11
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtintLivroOficio 
               Height          =   285
               Left            =   2490
               TabIndex        =   12
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtintFolhasOficio 
               Height          =   285
               Left            =   4020
               TabIndex        =   13
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox txtintVara 
               Height          =   285
               Left            =   5700
               TabIndex        =   14
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lbl_intNumOficio 
               AutoSize        =   -1  'True
               Caption         =   "Número"
               Height          =   195
               Left            =   360
               TabIndex        =   41
               Top             =   330
               Width           =   555
            End
            Begin VB.Label lbl_intLivroOficio 
               AutoSize        =   -1  'True
               Caption         =   "Livro"
               Height          =   195
               Left            =   2070
               TabIndex        =   40
               Top             =   330
               Width           =   555
            End
            Begin VB.Label lbl_intFolhasOficio 
               AutoSize        =   -1  'True
               Caption         =   "Folha"
               Height          =   195
               Left            =   3570
               TabIndex        =   39
               Top             =   330
               Width           =   390
            End
            Begin VB.Label lbl_intVara 
               AutoSize        =   -1  'True
               Caption         =   "Vara"
               Height          =   195
               Left            =   5310
               TabIndex        =   38
               Top             =   330
               Width           =   330
            End
         End
         Begin VB.Label lbl_intFolhasDistribuidor 
            AutoSize        =   -1  'True
            Caption         =   "Folha"
            Height          =   195
            Left            =   6990
            TabIndex        =   92
            Top             =   690
            Width           =   390
         End
         Begin VB.Label lbl_intLivroDistribuidor 
            AutoSize        =   -1  'True
            Caption         =   "Livro"
            Height          =   195
            Left            =   5520
            TabIndex        =   91
            Top             =   690
            Width           =   345
         End
         Begin VB.Label lbl_dtmDtDistribuidor 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Index           =   0
            Left            =   3840
            TabIndex        =   90
            Top             =   690
            Width           =   345
         End
         Begin VB.Label lbl_strSerieDistribuidor 
            AutoSize        =   -1  'True
            Caption         =   "Série"
            Height          =   195
            Left            =   2280
            TabIndex        =   89
            Top             =   690
            Width           =   360
         End
         Begin VB.Label lbl_strNumDistribuidor 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   570
            TabIndex        =   88
            Top             =   690
            Width           =   555
         End
      End
      Begin VB.Frame fra_ValoresTotais 
         Caption         =   "Valores Totais"
         Height          =   1605
         Left            =   180
         TabIndex        =   34
         Top             =   4860
         Width           =   9135
         Begin VB.TextBox txtdblQuantIndexador 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   8010
            Locked          =   -1  'True
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   1320
            Width           =   720
         End
         Begin VB.TextBox txtdblVlIndexador 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   5490
            Locked          =   -1  'True
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1275
         End
         Begin VB.TextBox txtstrIndexadorDescr 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   1320
            Width           =   825
         End
         Begin VB.TextBox txtdtmDtCalculoPeticao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   1320
            Width           =   870
         End
         Begin VB.TextBox txtdblVlTotTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   5490
            Locked          =   -1  'True
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   960
            Width           =   1275
         End
         Begin VB.TextBox txtdblVlTotJuros 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   5490
            Locked          =   -1  'True
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   750
            Width           =   1275
         End
         Begin VB.TextBox txtdblVlTotMulta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   5490
            Locked          =   -1  'True
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   540
            Width           =   1275
         End
         Begin VB.TextBox txtdblVlTotCorrecao 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   5490
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   330
            Width           =   1275
         End
         Begin VB.TextBox txtdblVlTotPrincipal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   960
            Width           =   1275
         End
         Begin VB.TextBox txtdblVlTotOriginal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   750
            Width           =   1275
         End
         Begin VB.TextBox txtdblVlTotTaxas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   540
            Width           =   1275
         End
         Begin VB.TextBox txtdblVlTotImpostos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   330
            Width           =   1275
         End
         Begin VB.Label lbl_dblQuantIndexador 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   7080
            TabIndex        =   57
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label lbl_dblVlIndexador 
            AutoSize        =   -1  'True
            Caption         =   "Valor Unitário:"
            Height          =   195
            Left            =   4440
            TabIndex        =   56
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label lbl_strIndexadorDescr 
            AutoSize        =   -1  'True
            Caption         =   "Indexador:"
            Height          =   195
            Left            =   2670
            TabIndex        =   55
            Top             =   1320
            Width           =   750
         End
         Begin VB.Label lbl_dtmDtCalculoPeticao 
            AutoSize        =   -1  'True
            Caption         =   "Data do Cálculo:"
            Height          =   195
            Left            =   240
            TabIndex        =   54
            Top             =   1320
            Width           =   1185
         End
         Begin VB.Label lbl_dblTotCorrecao 
            AutoSize        =   -1  'True
            Caption         =   "Corr. Monetária:"
            Height          =   195
            Left            =   4290
            TabIndex        =   53
            Top             =   330
            Width           =   1125
         End
         Begin VB.Label lbl_dblTotMulta 
            AutoSize        =   -1  'True
            Caption         =   "Multa:"
            Height          =   195
            Left            =   4980
            TabIndex        =   52
            Top             =   540
            Width           =   435
         End
         Begin VB.Label lbl_dblTotJuros 
            AutoSize        =   -1  'True
            Caption         =   "Juros:"
            Height          =   195
            Left            =   4995
            TabIndex        =   51
            Top             =   750
            Width           =   420
         End
         Begin VB.Label lbl_dblTotTotal 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   5010
            TabIndex        =   50
            Top             =   960
            Width           =   405
         End
         Begin VB.Label lbl_dblVlTotPrincipal 
            AutoSize        =   -1  'True
            Caption         =   "Principal:"
            Height          =   195
            Left            =   780
            TabIndex        =   44
            Top             =   960
            Width           =   645
         End
         Begin VB.Label lbl_dblTotOriginal 
            AutoSize        =   -1  'True
            Caption         =   "Original:"
            Height          =   195
            Left            =   855
            TabIndex        =   43
            Top             =   750
            Width           =   570
         End
         Begin VB.Label lbl_dblTotTaxas 
            AutoSize        =   -1  'True
            Caption         =   "Taxas:"
            Height          =   195
            Left            =   945
            TabIndex        =   42
            Top             =   540
            Width           =   480
         End
         Begin VB.Label lbl_dblTotImpostos 
            AutoSize        =   -1  'True
            Caption         =   "Impostos:"
            Height          =   195
            Left            =   750
            TabIndex        =   35
            Top             =   330
            Width           =   675
         End
      End
      Begin TabDlg.SSTab tab_3dExecutado 
         Height          =   1755
         Left            =   180
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3030
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   3096
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Executado"
         TabPicture(0)   =   "frmCadExecutivosFiscais.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl_strExecutadoCNPJCPF"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl_strExecutadoIdentidade(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl_strExecutadoNome"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txt_tabtxtstrExecutadoIdentidade"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtstrExecutadoIdentidade"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtstrExecutadoNome"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtstrExecutadoCNPJCPF"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Endereço de Notificação"
         TabPicture(1)   =   "frmCadExecutivosFiscais.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtstrExecutadoNomeLogNotif"
         Tab(1).Control(1)=   "txtstrExecutadoTitLogNotif"
         Tab(1).Control(2)=   "txtstrExecutadoUfNotif"
         Tab(1).Control(3)=   "txtstrExecutadoCidNotif"
         Tab(1).Control(4)=   "txtstrExecutadoNumLogNotif"
         Tab(1).Control(5)=   "txtintExecutadoCepNotif"
         Tab(1).Control(6)=   "txtstrExecutadoComplNotif"
         Tab(1).Control(7)=   "txtstrExecutadoBairroNotif"
         Tab(1).Control(8)=   "txtstrExecutadoTpLogNotif"
         Tab(1).Control(9)=   "txt_tabtxtstrExecutadoNomeLogNotif"
         Tab(1).Control(10)=   "lbl_strExecutadoTitLogNotif"
         Tab(1).Control(11)=   "lbl_strExecutadoTpLogNotif"
         Tab(1).Control(12)=   "lbl_strExecutadoNomeLogNotif"
         Tab(1).Control(13)=   "lbl_strExecutadoUFNotif"
         Tab(1).Control(14)=   "lbl_strExecutadoCidNotif"
         Tab(1).Control(15)=   "lbl_strExecutadoCepNotif"
         Tab(1).Control(16)=   "lbl_strExecutadoComplLogNotif"
         Tab(1).Control(17)=   "lbl_strExecutadoNumLogNotif"
         Tab(1).Control(18)=   "lbl_strExecutadoBairroLogNotif"
         Tab(1).ControlCount=   19
         Begin VB.TextBox txtstrExecutadoNomeLogNotif 
            Height          =   300
            Left            =   -70290
            MaxLength       =   100
            TabIndex        =   23
            Top             =   540
            Width           =   3975
         End
         Begin VB.TextBox txtstrExecutadoTitLogNotif 
            Height          =   300
            Left            =   -72225
            TabIndex        =   22
            Top             =   540
            Width           =   795
         End
         Begin VB.TextBox txtstrExecutadoUfNotif 
            Height          =   300
            Left            =   -71310
            MaxLength       =   2
            TabIndex        =   32
            Top             =   1260
            Width           =   375
         End
         Begin VB.TextBox txtstrExecutadoCidNotif 
            Height          =   300
            Left            =   -73770
            MaxLength       =   50
            TabIndex        =   28
            Top             =   1260
            Width           =   2115
         End
         Begin VB.TextBox txtstrExecutadoNumLogNotif 
            Height          =   300
            Left            =   -73770
            MaxLength       =   10
            TabIndex        =   25
            Top             =   900
            Width           =   825
         End
         Begin VB.TextBox txtintExecutadoCepNotif 
            Height          =   300
            Left            =   -70515
            MaxLength       =   9
            TabIndex        =   33
            Top             =   1245
            Width           =   885
         End
         Begin VB.TextBox txtstrExecutadoComplNotif 
            Height          =   300
            Left            =   -72210
            MaxLength       =   10
            TabIndex        =   26
            Top             =   900
            Width           =   1710
         End
         Begin VB.TextBox txtstrExecutadoBairroNotif 
            Height          =   300
            Left            =   -69855
            MaxLength       =   50
            TabIndex        =   27
            Top             =   900
            Width           =   2520
         End
         Begin VB.TextBox txtstrExecutadoCNPJCPF 
            Height          =   315
            Left            =   1260
            MaxLength       =   100
            TabIndex        =   17
            Top             =   900
            Width           =   2340
         End
         Begin VB.TextBox txtstrExecutadoNome 
            Height          =   315
            Left            =   1260
            MaxLength       =   100
            TabIndex        =   16
            Top             =   540
            Width           =   7410
         End
         Begin VB.TextBox txtstrExecutadoIdentidade 
            Height          =   315
            Left            =   4980
            MaxLength       =   50
            TabIndex        =   18
            Top             =   900
            Width           =   2340
         End
         Begin VB.TextBox txtstrExecutadoTpLogNotif 
            Height          =   300
            Left            =   -73755
            TabIndex        =   21
            Top             =   540
            Width           =   795
         End
         Begin VB.TextBox txt_tabtxtstrExecutadoNomeLogNotif 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   -73710
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   600
            Width           =   555
         End
         Begin VB.TextBox txt_tabtxtstrExecutadoIdentidade 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   6450
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   960
            Width           =   555
         End
         Begin VB.Label lbl_strExecutadoTitLogNotif 
            AutoSize        =   -1  'True
            Caption         =   "Título"
            Height          =   195
            Left            =   -72690
            TabIndex        =   112
            Top             =   615
            Width           =   420
         End
         Begin VB.Label lbl_strExecutadoTpLogNotif 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   -74130
            TabIndex        =   111
            Top             =   615
            Width           =   315
         End
         Begin VB.Label lbl_strExecutadoNomeLogNotif 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   -71220
            TabIndex        =   110
            Top             =   660
            Width           =   810
         End
         Begin VB.Label lbl_strExecutadoUFNotif 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   -71580
            TabIndex        =   49
            Top             =   1350
            Width           =   210
         End
         Begin VB.Label lbl_strExecutadoCidNotif 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   -74550
            TabIndex        =   31
            Top             =   1335
            Width           =   705
         End
         Begin VB.Label lbl_strExecutadoCepNotif 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Left            =   -70875
            TabIndex        =   48
            Top             =   1335
            Width           =   315
         End
         Begin VB.Label lbl_strExecutadoComplLogNotif 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   -72720
            TabIndex        =   29
            Top             =   990
            Width           =   480
         End
         Begin VB.Label lbl_strExecutadoNumLogNotif 
            AutoSize        =   -1  'True
            Caption         =   "N°"
            Height          =   195
            Left            =   -73995
            TabIndex        =   24
            Top             =   990
            Width           =   180
         End
         Begin VB.Label lbl_strExecutadoBairroLogNotif 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   -70335
            TabIndex        =   30
            Top             =   990
            Width           =   405
         End
         Begin VB.Label lbl_strExecutadoNome 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   765
            TabIndex        =   47
            Top             =   600
            Width           =   420
         End
         Begin VB.Label lbl_strExecutadoIdentidade 
            AutoSize        =   -1  'True
            Caption         =   "Identidade"
            Height          =   195
            Index           =   1
            Left            =   4140
            TabIndex        =   46
            Top             =   990
            Width           =   780
         End
         Begin VB.Label lbl_strExecutadoCNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   420
            TabIndex        =   45
            Top             =   990
            Width           =   780
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Debitos 
         Height          =   2145
         Left            =   -74820
         TabIndex        =   64
         Top             =   1320
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   3784
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "pkId"
         Columns(0).DataField=   "pkId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Inscrição"
         Columns(1).DataField=   "strInscricao"
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Composição da Receita"
         Columns(2).DataField=   "strComposicaoDaReceita"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Exercício"
         Columns(3).DataField=   "intExercicio"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   16
         Columns(4)._MaxComboItems=   5
         Columns(4).ValueItems(0)._DefaultItem=   0
         Columns(4).ValueItems(0).Value=   "1"
         Columns(4).ValueItems(0).Value.vt=   8
         Columns(4).ValueItems(0).DisplayValue=   "Imobiliário"
         Columns(4).ValueItems(0).DisplayValue.vt=   8
         Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems(1)._DefaultItem=   0
         Columns(4).ValueItems(1).Value=   "2"
         Columns(4).ValueItems(1).Value.vt=   8
         Columns(4).ValueItems(1).DisplayValue=   "Econômico"
         Columns(4).ValueItems(1).DisplayValue.vt=   8
         Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems(2)._DefaultItem=   0
         Columns(4).ValueItems(2).Value=   "3"
         Columns(4).ValueItems(2).Value.vt=   8
         Columns(4).ValueItems(2).DisplayValue=   "Dívida Ativa"
         Columns(4).ValueItems(2).DisplayValue.vt=   8
         Columns(4).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems(3)._DefaultItem=   0
         Columns(4).ValueItems(3).Value=   "4"
         Columns(4).ValueItems(3).Value.vt=   8
         Columns(4).ValueItems(3).DisplayValue=   "Acordo"
         Columns(4).ValueItems(3).DisplayValue.vt=   8
         Columns(4).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems(4)._DefaultItem=   0
         Columns(4).ValueItems(4).Value=   "5"
         Columns(4).ValueItems(4).Value.vt=   8
         Columns(4).ValueItems(4).DisplayValue=   "Preço Público"
         Columns(4).ValueItems(4).DisplayValue.vt=   8
         Columns(4).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems.Count=   5
         Columns(4).Caption=   "Utilização"
         Columns(4).DataField=   "intUtilizacao"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2831"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2752"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=7223"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=7144"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
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
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Envolvidos 
         Height          =   2415
         Left            =   -74820
         TabIndex        =   79
         Top             =   4050
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   4260
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "pkId"
         Columns(0).DataField=   "pkId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nome"
         Columns(1).DataField=   "strEnvolvidoNome"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "CPF/CNPJ"
         Columns(2).DataField=   "strEnvolvidoCNPJCPF"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=7223"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=7144"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
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
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
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
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txt_tabtdb_Envolvidos 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   -66570
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   5820
         Width           =   555
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1485
      Left            =   30
      TabIndex        =   58
      Top             =   6750
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   2619
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "pkId"
      Columns(0).DataField=   "pkId"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nº Seqüencial"
      Columns(1).DataField=   "intNumSeq"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nº Distribuidor"
      Columns(2).DataField=   "strNumDistribuidor"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nome"
      Columns(3).DataField=   "strExecutadoNome"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=7938"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=7858"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
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
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
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
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
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
Attribute VB_Name = "frmCadExecutivosFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando As Boolean
Dim mblnPrimeiraVez As Boolean

Private Sub chkbitDistribuicaoEletronica_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
End Sub

Private Sub chkbitDistribuido_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1387
    
    If mblnAlterando Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    
    VirificaGradeListView Me
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim blnToolBarGeral As Boolean

    blnToolBarGeral = False
        
    mblnPrimeiraVez = True
    
    Select Case UCase(strModoOperacao)
        Case UCase(gstrLocalizar)
            txtdblQuantIndexador = ""
            txtdblVlIndexador = ""
            txtdblVlTotCorrecao = ""
            txtdblVlTotImpostos = ""
            txtdblVlTotJuros = ""
            txtdblVlTotMulta = ""
            txtdblVlTotOriginal = ""
            txtdblVlTotPrincipal = ""
            txtdblVlTotTaxas = ""
            txtdblVlTotTotal = ""
            txtdtmDtCalculoPeticao = ""
            
            mblnAlterando = False
            
            LimpaGrids
                        
            blnToolBarGeral = True
            
        Case UCase(gstrFechar)
            blnToolBarGeral = True
        Case UCase(gstrNovo)
            LimpaGrids
            blnToolBarGeral = True
    End Select
    
    If blnToolBarGeral Then
        ToolBarGeral strModoOperacao, gstrExecutivo, mblnAlterando, tdb_Lista, Me
    End If

    
    
End Sub


Private Sub Form_Load()
    TrocaCorObjeto txt_NumDist, True
    TrocaCorObjeto txt_Serie, True
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3dPasta.Tab = 0 Then
        TrocaCorObjeto txtintNumSeq, False
        TrocaCorObjeto txtintNumeroProtocolo, False
        TrocaCorObjeto txtintLoteExecutivo, False
        chkbitDistribuicaoEletronica.Visible = True
        txt_NumDist.Visible = False
        lbl_NumDist.Visible = False
        txt_Serie.Visible = False
        lbl_Serie.Visible = False
    Else
        TrocaCorObjeto txtintNumSeq, True
        TrocaCorObjeto txtintNumeroProtocolo, True
        TrocaCorObjeto txtintLoteExecutivo, True
        chkbitDistribuicaoEletronica.Visible = False
        txt_NumDist.Visible = True
        lbl_NumDist.Visible = True
        txt_Serie.Visible = True
        lbl_Serie.Visible = True
    End If
End Sub

Private Sub tdb_Debitos_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
Dim intUtilizacao As Byte
    If ColIndex = 1 Then
        'intUtilizacao = Left(Value, 1)
        intUtilizacao = tdb_Debitos.Columns("intUtilizacao").Value
        Value = gstrFormataInscricao(Right(CStr(Value), gintRetornaTamanhoMascara(intUtilizacao)), CInt(intUtilizacao))
    End If
End Sub

Private Sub tdb_Debitos_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 1
End Sub

Private Sub tdb_Envolvidos_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
End Sub

Private Sub tdb_Envolvidos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim adoResultado As ADODB.Recordset
Dim strSQL As String
Dim strAux As String

    With tdb_Envolvidos
        If (Not .EOF And Not .BOF) Then
            
            Set gobjBanco = New clsBanco
            
            strSQL = strQueryEnvolvidos
            strSQL = strSQL & " AND EE.pkId = " & .Columns(0).Value
            
            If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                If Not adoResultado.EOF Then
                    With adoResultado
                        
                        txt_strEnvolvidoNome = gstrENulo(.Fields("strEnvolvidoNome").Value)
                        
                        txt_strEnvolvidoCPFCNPJ = gstrCGCCPFFormatado(gstrENulo(.Fields("strEnvolvidoCNPJCPF").Value))
                                                
                        txt_strEnvolvidoIdentidade = gstrENulo(.Fields("strEnvolvidoIdentidade").Value)
                        txt_strEnvolvidoNomeLogNotif = gstrENulo(.Fields("strEnvolvidoNomeLogNotif").Value)
                        txt_strEnvolvidoNumLogNotif = gstrENulo(.Fields("strEnvolvidoNumLogNotif").Value)
                        txt_strEnvolvidoComplLogNotif = gstrENulo(.Fields("strEnvolvidoComplLogNotif").Value)
                        txt_strEnvolvidoCidadeLogNotif = gstrENulo(.Fields("strEnvolvidoCidadeLogNotif").Value)
                        txt_strEnvolvidoCEPLogNotif = gstrENulo(.Fields("intEnvolvidoCEPLogNotif").Value)
                        txt_strEnvolvidoUFLogNotif = gstrENulo(.Fields("strEnvolvidoUFLogNotif").Value)
                        txt_strEnvolvidoBairroLogNotif = gstrENulo(.Fields("strEnvolvidoBairroLogNotif").Value)
                        txt_strEnvolvidoTituloLogNotif = gstrENulo(.Fields("strEnvolvidoTituloLogNotif").Value)
                        txt_strEnvolvidoTipoLogNotif = gstrENulo(.Fields("strEnvolvidoTipoLogNotif").Value)
                        
                    End With
                End If
            End If
        End If
    End With
End Sub


Private Sub tdb_Debitos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 
 With tdb_Debitos
        If (Not .EOF And Not .BOF) Then
            
            If tdb_Debitos.Columns(0) <> "" Then
                LeDaTabelaParaObj "", tdb_Parcelas, strQueryParcelas(tdb_Debitos.Columns(0))
            End If
        End If
    End With
    
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    tdb_Lista_RowColChange 0, 0
End Sub

Private Sub tdb_Lista_FilterChange()
    gblnFilraCampos tdb_Lista
    With tdb_Lista
       If Not .BOF And Not .EOF Then
           If .Bookmark = 1 Then
               tdb_Lista_RowColChange 0, 0
          End If
       End If
    End With
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Lista
        If (Not .EOF And Not .BOF) Then
            If Not mblnPrimeiraVez Then
                
                Screen.MousePointer = vbHourglass
                
                LimpaGrids
                
                txtPKId = .Columns("PKID").Value
                
                LeDaTabelaParaObj gstrExecutivo, Me
                
                txtstrExecutadoCNPJCPF = gstrCGCCPFFormatado(gstrENulo(txtstrExecutadoCNPJCPF))
                            
                LeDaTabelaParaObj "", tdb_Debitos, strQueryDebitos
                
                LeDaTabelaParaObj "", tdb_Envolvidos, strQueryEnvolvidos
                
                mblnAlterando = True
            Else
                mblnPrimeiraVez = False
            End If
        End If
    End With
    
End Sub

Private Sub tdb_Parcelas_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Select Case ColIndex
        Case 3, 4, 5, 6, 7, 8
            Value = gstrConvVrDoSql(Value)
    End Select
End Sub

Private Sub tdb_Parcelas_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 1
End Sub

Private Sub txt_strEnvolvidoIdentidade_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoIdentidade
End Sub

Private Sub txt_strEnvolvidoIdentidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoIdentidade, False
End Sub

Private Sub txt_strEnvolvidoCPFCNPJ_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoCPFCNPJ
End Sub

Private Sub txt_strEnvolvidoCPFCNPJ_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoCPFCNPJ, False
End Sub

Private Sub txt_strEnvolvidoNome_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoNome
End Sub

Private Sub txt_strEnvolvidoNome_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoNome, False
End Sub

Private Sub txt_strEnvolvidoTipoLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoTipoLogNotif
End Sub

Private Sub txt_strEnvolvidoTipoLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoTipoLogNotif, False
End Sub

Private Sub txt_strEnvolvidoTituloLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoTituloLogNotif
End Sub

Private Sub txt_strEnvolvidoTituloLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoTituloLogNotif, False
End Sub

Private Sub txt_strEnvolvidoUFLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoUFLogNotif
End Sub

Private Sub txt_strEnvolvidoUFLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoUFLogNotif, False
End Sub

Private Sub txt_strEnvolvidoCidadeLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoCidadeLogNotif
End Sub

Private Sub txt_strEnvolvidoCidadeLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoCidadeLogNotif
End Sub

Private Sub txt_strEnvolvidoNumLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoNumLogNotif
End Sub


Private Sub txt_strEnvolvidoNumLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strEnvolvidoNumLogNotif, False
End Sub

Private Sub txt_strEnvolvidoComplLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoComplLogNotif
End Sub

Private Sub txt_strEnvolvidoComplLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoComplLogNotif, False
End Sub

Private Sub txt_strEnvolvidoNomeLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoNomeLogNotif
End Sub


Private Sub txt_strEnvolvidoNomeLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoNomeLogNotif, False
End Sub

Private Sub txt_strEnvolvidoBairroLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoBairroLogNotif
End Sub

Private Sub txt_strEnvolvidoBairroLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoBairroLogNotif, False
End Sub

Private Sub txt_strEnvolvidoCEPLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 2
    MarcaCampo txt_strEnvolvidoCEPLogNotif
End Sub

Private Sub txt_strEnvolvidoCEPLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEnvolvidoCEPLogNotif, False
End Sub

Private Sub txtdtmDtDistribuidor_LostFocus()
    txtdtmDtDistribuidor = gstrDataFormatada(txtdtmDtDistribuidor)
End Sub

Private Sub txtintLoteExecutivo_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtintLoteExecutivo
End Sub

Private Sub txtintNumeroProtocolo_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtintNumeroProtocolo
End Sub

Private Sub txtintNumSeq_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtintNumSeq
End Sub

Private Sub txtstrExecutadoBairroNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
    MarcaCampo txtstrExecutadoBairroNotif
End Sub


Private Sub txtstrExecutadoTpLogNotif_GotFocus()
    MarcaCampo txtstrExecutadoTpLogNotif
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
End Sub

Private Sub txtstrExecutadoTpLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoTpLogNotif, False
End Sub

Private Sub txtstrExecutadoTitLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
    MarcaCampo txtstrExecutadoTitLogNotif
End Sub

Private Sub txtstrExecutadoTitLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoTpLogNotif, False
End Sub

Private Sub txtstrExecutadoUfNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
    MarcaCampo txtstrExecutadoUfNotif
End Sub

Private Sub txtstrExecutadoUfNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoUfNotif
End Sub

Private Sub txt_tabtdb_Envolvidos_GotFocus()
    
    If tab_3dPasta.Tab = 0 Then
        tdb_Envolvidos.SetFocus
    ElseIf tab_3dPasta.Tab = 2 Then
        AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 0
        txtintNumSeq.SetFocus
    End If

End Sub

Private Sub txt_tabtxtstrExecutadoIdentidade_GotFocus()
    
    If tab_3dExecutado.Tab = 0 Then
        txtstrExecutadoTpLogNotif.SetFocus
    ElseIf tab_3dExecutado.Tab = 1 Then
        txtstrExecutadoIdentidade.SetFocus
        AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 0
    End If
    
End Sub

Private Sub txt_tabtxtstrExecutadoNomeLogNotif_GotFocus()
    
    If tab_3dExecutado.Tab = 1 Then
        txt_tabtxtstrExecutadoIdentidade.SetFocus
    ElseIf tab_3dExecutado.Tab = 0 Then
        txtstrExecutadoNomeLogNotif.SetFocus
        AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
    End If
    
End Sub

Private Sub txtdtmDtDistribuidor_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtdtmDtDistribuidor
End Sub

Private Sub txtdtmDtDistribuidor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtDistribuidor, False
End Sub

Private Sub txtintExecutadoCepNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
    MarcaCampo txtintExecutadoCepNotif
End Sub

Private Sub txtintFolhasDistribuidor_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtintFolhasDistribuidor
End Sub

Private Sub txtintFolhasOficio_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtintFolhasOficio
End Sub

Private Sub txtintLivroDistribuidor_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtintLivroDistribuidor
End Sub

Private Sub txtintLivroOficio_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtintLivroOficio
End Sub

Private Sub txtintNumOficio_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtintNumOficio
End Sub

Private Sub txtintVara_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtintVara
End Sub

Private Sub txtstrExecutadoCidNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
    MarcaCampo txtstrExecutadoCidNotif
End Sub

Private Sub txtstrExecutadoCNPJCPF_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 0
    MarcaCampo txtstrExecutadoCNPJCPF
End Sub

Private Sub txtstrExecutadoComplNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
    MarcaCampo txtstrExecutadoComplNotif
End Sub

Private Sub txtstrExecutadoComplNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoComplNotif, False
End Sub

Private Sub txtstrExecutadoIdentidade_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 0
    MarcaCampo txtstrExecutadoIdentidade
End Sub

Private Sub txtstrExecutadoNomeLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
    MarcaCampo txtstrExecutadoNomeLogNotif
End Sub

Private Sub txtstrExecutadoNomeLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoNomeLogNotif, False
End Sub

Private Sub txtstrExecutadoNumLogNotif_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 1
    MarcaCampo txtstrExecutadoNumLogNotif
End Sub

Private Sub txtstrExecutadoNumLogNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrExecutadoNumLogNotif, False
End Sub

Private Sub txtintFolhasDistribuidor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintFolhasDistribuidor, False
End Sub

Private Sub txtintFolhasOficio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintFolhasOficio, False
End Sub

Private Sub txtintLivroDistribuidor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintLivroDistribuidor, False
End Sub

Private Sub txtintLivroOficio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintLivroOficio, False
End Sub

Private Sub txtintLoteExecutivo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintLoteExecutivo, False
End Sub

Private Sub txtintNumeroProtocolo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumeroProtocolo, False
End Sub

Private Sub txtintNumOficio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumOficio, False
End Sub

Private Sub txtintNumSeq_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumSeq, False
End Sub

Private Sub txtintVara_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintVara, False
End Sub

Private Sub txtstrExecutadoBairroNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoBairroNotif, False
End Sub

Private Sub txtintExecutadoCepNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtintExecutadoCepNotif, False
End Sub

Private Sub txtstrExecutadoCidNotif_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoCidNotif, False
End Sub

Private Sub txtstrExecutadoCNPJCPF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoCNPJCPF, False
End Sub

Private Sub txtstrExecutadoIdentidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoIdentidade, False
End Sub

Private Sub txtstrExecutadoNome_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0, tab_3dExecutado, 0
    MarcaCampo txtstrExecutadoNome
End Sub

Private Sub txtstrExecutadoNome_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrExecutadoNome, False
End Sub

Private Sub txtstrNumDistribuidor_Change()
    txt_NumDist = txtstrNumDistribuidor
End Sub

Private Sub txtstrNumDistribuidor_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtstrNumDistribuidor
End Sub

Private Sub txtstrNumDistribuidor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNumDistribuidor, False
End Sub

Private Sub txtstrSerieDistribuidor_Change()
    txt_Serie = txtstrSerieDistribuidor
End Sub

Private Sub txtstrSerieDistribuidor_GotFocus()
    AtivaPastaDeObjeto tab_3dPasta, 0
    MarcaCampo txtstrSerieDistribuidor
End Sub

Private Sub txtstrSerieDistribuidor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSerieDistribuidor, False
End Sub

Private Function strQueryDebitos() As String
Dim strSQL

    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " DA.pkId, "
    'strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "LA.intUtilizacao") & strCONCAT & "'-'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "LA.strInscricao") & " strInscricao, "
    strSQL = strSQL & " LA.strInscricao strInscricao, "
    strSQL = strSQL & " LA.strComposicaoDaReceita, "
    strSQL = strSQL & " LA.IntExercicio, "
    strSQL = strSQL & " LA.intUtilizacao "
    strSQL = strSQL & " FROM " & gstrDativa & " DA, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " LA.pkId = DA.intLancamentoAlfa "
    strSQL = strSQL & " AND DA.intExecutivo = " & txtPKId
    
    strQueryDebitos = strSQL
    
End Function


Private Function strQueryParcelas(intDativa As Long) As String
Dim strSQL

    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " EP.intParcela, "
    strSQL = strSQL & " EP.Dtmdtvencimento, "
    strSQL = strSQL & " MO.Strabreviatura, "
    strSQL = strSQL & " EP.DblVlOriginal, "
    strSQL = strSQL & " EP.DblVlPrincipal, "
    strSQL = strSQL & " EP.DblVlCorrecao, "
    strSQL = strSQL & " EP.DblVlMulta, "
    strSQL = strSQL & " EP.DblVlJuros, "
    strSQL = strSQL & " EP.dblVlTotal "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrExecutivoParcela & " EP, "
    strSQL = strSQL & gstrMoedas & " MO "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " EP.intMoeda " & strOUTJSQLServer & "= MO.pkId " & strOUTJOracle
    strSQL = strSQL & " AND EP.intDativa = " & intDativa
    strSQL = strSQL & " ORDER BY intParcela "
    
    strQueryParcelas = strSQL
    
End Function

Private Function strQueryEnvolvidos() As String
Dim strSQL

    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " * "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrExecutivoEnvolvidos & " EE "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " EE.intExecutivo = " & txtPKId
        
    strQueryEnvolvidos = strSQL
    
End Function

Public Sub LimpaGrids()

    Set tdb_Parcelas.DataSource = Nothing
    Set tdb_Debitos.DataSource = Nothing
    Set tdb_Envolvidos.DataSource = Nothing
    
    txt_strEnvolvidoNome = ""
    txt_strEnvolvidoCPFCNPJ = ""
    txt_strEnvolvidoIdentidade = ""
    txt_strEnvolvidoNomeLogNotif = ""
    txt_strEnvolvidoNumLogNotif = ""
    txt_strEnvolvidoComplLogNotif = ""
    txt_strEnvolvidoCidadeLogNotif = ""
    txt_strEnvolvidoCEPLogNotif = ""
    txt_strEnvolvidoBairroLogNotif = ""
    txt_strEnvolvidoUFLogNotif = ""
    txt_strEnvolvidoTituloLogNotif = ""
    txt_strEnvolvidoTipoLogNotif = ""
    
End Sub

