VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadEmissaoValidadeDeDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão e Validade de Documentos"
   ClientHeight    =   7080
   ClientLeft      =   2460
   ClientTop       =   1920
   ClientWidth     =   8985
   HelpContextID   =   12
   Icon            =   "CadEmissaoValidadeDeDocumentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8985
   Begin VB.TextBox txtPKId 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   420
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   795
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6855
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Emissão e Validade de Documentos"
      TabPicture(0)   =   "CadEmissaoValidadeDeDocumentos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_Documentos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Inscricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Endereco"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_Data"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Bla"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fra_Bla 
         Height          =   1755
         Left            =   120
         TabIndex        =   43
         Top             =   1020
         Width           =   8445
         Begin VB.TextBox txtintContribuinte 
            Height          =   315
            Left            =   7530
            TabIndex        =   50
            Top             =   570
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txt_intContribuinte 
            Height          =   315
            Left            =   2010
            MaxLength       =   50
            TabIndex        =   49
            Top             =   570
            Width           =   5445
         End
         Begin VB.TextBox txt_CNPJCPF 
            Height          =   285
            Left            =   2010
            MaxLength       =   50
            TabIndex        =   8
            Top             =   960
            Width           =   2505
         End
         Begin VB.TextBox txtstrNumeroDoProcesso 
            Height          =   285
            Left            =   5820
            MaxLength       =   20
            TabIndex        =   10
            Top             =   1350
            Width           =   1635
         End
         Begin MSDataListLib.DataCombo dbcintDocumentosEmitidos 
            Height          =   315
            Left            =   2010
            TabIndex        =   9
            Top             =   1320
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcstrInscricao 
            Height          =   315
            Left            =   2010
            TabIndex        =   7
            Top             =   210
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lblintContribuinte 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte"
            Height          =   195
            Left            =   1035
            TabIndex        =   48
            Top             =   675
            Width           =   840
         End
         Begin VB.Label lblintDocumentosEmitidos 
            AutoSize        =   -1  'True
            Caption         =   "Documentos emitidos"
            Height          =   195
            Left            =   330
            TabIndex        =   47
            Top             =   1425
            Width           =   1515
         End
         Begin VB.Label lblstrInscricao 
            Alignment       =   1  'Right Justify
            Caption         =   "Inscrição"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   285
            Width           =   1755
         End
         Begin VB.Label lbl_CNPJCPF 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CPF / CNPJ"
            Height          =   195
            Left            =   975
            TabIndex        =   45
            Top             =   1035
            Width           =   870
         End
         Begin VB.Label lblstrNumeroDoProcesso 
            AutoSize        =   -1  'True
            Caption         =   "Nº do Processo"
            Height          =   195
            Left            =   4620
            TabIndex        =   44
            Top             =   1425
            Width           =   1110
         End
      End
      Begin VB.Frame fra_Data 
         Caption         =   "Datas"
         Height          =   1005
         Left            =   120
         TabIndex        =   36
         Top             =   4230
         Width           =   8445
         Begin VB.TextBox txtdtmEntrega 
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
            Left            =   7290
            MaxLength       =   10
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtdtmEmissao 
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
            Left            =   3060
            MaxLength       =   10
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtdtmConcessao 
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
            Left            =   5280
            MaxLength       =   10
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtdtmSolicitacao 
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
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtdtmValidade 
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
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   23
            Top             =   600
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dbcintOcorrencia 
            Height          =   315
            Left            =   5280
            TabIndex        =   24
            Top             =   570
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lbldtmEmissao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Left            =   2370
            TabIndex        =   42
            Top             =   330
            Width           =   585
         End
         Begin VB.Label lbldtmValidade 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Validade"
            Height          =   195
            Left            =   300
            TabIndex        =   41
            Top             =   690
            Width           =   615
         End
         Begin VB.Label lbldtmSolicitacao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Solicitação"
            Height          =   195
            Left            =   135
            TabIndex        =   40
            Top             =   330
            Width           =   780
         End
         Begin VB.Label lbldtmConcessao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Concessão"
            Height          =   195
            Left            =   4350
            TabIndex        =   39
            Top             =   330
            Width           =   795
         End
         Begin VB.Label lblintOcorrencia 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ocorrências"
            Height          =   195
            Left            =   4290
            TabIndex        =   38
            Top             =   690
            Width           =   855
         End
         Begin VB.Label lbldtmEntrega 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Entrega"
            Height          =   195
            Left            =   6600
            TabIndex        =   37
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   "Endereço de Correspondência"
         Height          =   1395
         Left            =   120
         TabIndex        =   27
         Top             =   2820
         Width           =   8445
         Begin VB.TextBox txt_UF 
            Height          =   285
            Left            =   6120
            MaxLength       =   2
            TabIndex        =   17
            Top             =   960
            Width           =   510
         End
         Begin VB.TextBox txt_Municipio 
            Height          =   285
            Left            =   5130
            MaxLength       =   50
            TabIndex        =   15
            Top             =   600
            Width           =   3165
         End
         Begin VB.TextBox txt_Cep 
            Height          =   285
            Left            =   7200
            MaxLength       =   9
            TabIndex        =   18
            Top             =   960
            Width           =   1080
         End
         Begin VB.TextBox txt_Complemento 
            Height          =   285
            Left            =   7260
            MaxLength       =   20
            TabIndex        =   13
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox txt_Numero 
            Height          =   285
            Left            =   5790
            MaxLength       =   8
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txt_Logradouro 
            Height          =   285
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   11
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox txt_Bairro 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   14
            Top             =   600
            Width           =   3105
         End
         Begin VB.TextBox txt_Distrito 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   16
            Top             =   960
            Width           =   3525
         End
         Begin VB.Label lblintCepC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   6795
            TabIndex        =   35
            Top             =   1035
            Width           =   285
         End
         Begin VB.Label lblintUFC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   5760
            TabIndex        =   34
            Top             =   1035
            Width           =   210
         End
         Begin VB.Label lblstrComplementoC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6720
            TabIndex        =   33
            Top             =   315
            Width           =   480
         End
         Begin VB.Label lblintNumeroC 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5520
            TabIndex        =   32
            Top             =   315
            Width           =   180
         End
         Begin VB.Label lblintLogradouroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   315
            Width           =   810
         End
         Begin VB.Label lblintBairroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   525
            TabIndex        =   30
            Top             =   660
            Width           =   405
         End
         Begin VB.Label lblintMunicipioC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   4290
            TabIndex        =   29
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblstrDistritoC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   450
            TabIndex        =   28
            Top             =   1020
            Width           =   480
         End
      End
      Begin VB.Frame fra_Inscricao 
         Height          =   645
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   8445
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Urbano"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Rural"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   1
            Left            =   1770
            TabIndex        =   3
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Econômico"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   2
            Left            =   3270
            TabIndex        =   4
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Contribuição de Melhorias"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   3
            Left            =   4470
            TabIndex        =   5
            Top             =   270
            Width           =   2205
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Receitas Diversas"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   4
            Left            =   6690
            TabIndex        =   6
            Top             =   270
            Width           =   1605
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Documentos 
         Height          =   1395
         Left            =   150
         TabIndex        =   25
         Top             =   5310
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   2461
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Documentos do Contribuinte"
         Columns(1).DataField=   "strDescricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Numero do Processo"
         Columns(2).DataField=   "strNumeroDoProcesso"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "intcontribuinte"
         Columns(3).DataField=   "intContribuinte"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2487"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2408"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=11245"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=11165"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3122"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3043"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
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
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         AnimateWindow   =   3
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=44"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=47"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=24,.parent=43"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=44"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=45"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=47"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=43"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=44"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=45"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=47"
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
End
Attribute VB_Name = "frmCadEmissaoValidadeDeDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando               As Boolean
Dim mobjAux                     As Object
Dim mblnSelecionou              As Boolean
Dim mblnPrimeiraVez             As Boolean
Dim intIndiceOPT                As Integer
Dim adoResultado                As ADODB.Recordset
Dim bytOrdenacao                As Byte
Dim blnOrdenacaoAsc             As Boolean

Private Sub dbcintDocumentosEmitidos_Click(Area As Integer)
   DropDownDataCombo dbcintDocumentosEmitidos, Me, Area
End Sub

Private Sub dbcintDocumentosEmitidos_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintDocumentosEmitidos, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrencia_Click(Area As Integer)
   DropDownDataCombo dbcintOcorrencia, Me, Area
End Sub

Private Sub dbcintOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcstrInscricao_Change()
    If dbcstrInscricao.MatchedWithList Then
        MostraDadosContribuinte (dbcstrInscricao.BoundText)
    Else
        LimpaCorrespondencia
    End If
    
End Sub

Private Sub dbcstrInscricao_Click(Area As Integer)
    DropDownDataCombo dbcstrInscricao, Me, Area
    If Area = 2 And dbcstrInscricao.MatchedWithList Then
        LeDaTabelaParaObj "", tdb_Documentos, strQuery
        If intIndiceOPT < 2 Then
            ProprietarioImobiliario intIndiceOPT
        Else
            If intIndiceOPT = 2 Then
                ProprietarioEconomico
            Else
                If intIndiceOPT = 3 Then
                    ProprietarioImobiliario 0
                Else
                    txtintContribuinte = dbcstrInscricao.BoundText
                    MostraDadosContribuinte (dbcstrInscricao.BoundText)
                End If
            End If
        End If
    End If
End Sub

Private Sub dbcstrInscricao_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcstrInscricao, Me, , KeyCode, Shift
End Sub

Private Sub tdb_Documentos_HeadClick(ByVal ColIndex As Integer)
    blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
   
    bytOrdenacao = ColIndex: MantemForm gstrRefresh
End Sub

Private Sub txt_intContribuinte_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_intContribuinte
End Sub

Private Sub dbcintDocumentosEmitidos_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintDocumentosEmitidos
End Sub

Private Sub dbcintOcorrencia_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", dbcintOcorrencia
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 632
    VirificaGradeListView Me
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    dbcstrInscricao.Tag = "SELECT intContribuinte, " & _
                          gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao " & _
                          "FROM " & gstrImobiliario & " WHERE intContribuinte is not null;strInscricao"
    LeDaTabelaParaObj gstrDocumentoEmitido, dbcintDocumentosEmitidos
    LeDaTabelaParaObj gstrOcorrencia, dbcintOcorrencia, strQuerryEntrega
    optbitTipoDeInscricao_Click (0)
    TrocaCorObjeto txt_intContribuinte, True
    TrocaCorObjeto txt_CNPJCPF, True
    TrocaCorObjeto txt_Bairro, True
    TrocaCorObjeto txt_Cep, True
    TrocaCorObjeto txt_Complemento, True
    TrocaCorObjeto txt_Distrito, True
    TrocaCorObjeto txt_Logradouro, True
    TrocaCorObjeto txt_Municipio, True
    TrocaCorObjeto txt_Numero, True
    TrocaCorObjeto txt_UF, True
    
    VerificaObjParaAplicar mobjAux
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub dbcstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcstrInscricao
End Sub

Private Sub optbitTipoDeInscricao_KeyPress(Index As Integer, KeyAscii As Integer)
CaracterValido KeyAscii, "", optbitTipoDeInscricao(Index)
End Sub

Private Sub tdb_Documentos_Click()
    mblnPrimeiraVez = True
    With tdb_Documentos
        If Not .BOF And Not .EOF Then
            If .Bookmark = 1 Then
                tdb_Documentos_RowColChange 0, 0
            End If
        End If
    End With
End Sub

Sub tdb_Documentos_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Documentos_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Documentos
End Sub

Private Sub tdb_Documentos_KeyPress(KeyAscii As Integer)
    If tdb_Documentos.Col = 1 Then
        CaracterValido KeyAscii, "A", tdb_Documentos
    ElseIf tdb_Documentos.Col = 2 Then
        CaracterValido KeyAscii, "N", tdb_Documentos
    End If
End Sub

Private Sub tdb_Documentos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Documentos
        If Not .EOF And Not .BOF Then

            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value
                
                LeDaTabelaParaObj gstrEmissaoValidade, Me
                
                PreencherListaDeOpcoes dbcstrInscricao, .Columns("intContribuinte").Value
                dbcstrInscricao.BoundText = .Columns("intContribuinte").Value
                
                gCorLinhaSelecionada tdb_Documentos

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
            End If

        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSql As String
Dim i As Integer

    For i = 0 To 4
        If optbitTipoDeInscricao(i).Value = True Then
            intIndiceOPT = i
            i = 4
        End If
    Next
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaCorrespondencia
    End If
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If blnDadosOk = False Then
            Exit Sub
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
        mblnPrimeiraVez = False
    End If
    
If strModoOperacao = gstrPreencherLista Then

  Select Case intIndiceOPT
        Case 0
            dbcstrInscricao.Tag = "SELECT intContribuinte, " & _
                                  gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao " & _
                                  "FROM " & gstrImobiliario & " WHERE intContribuinte is not null;strInscricao"
            PreencherListaDeOpcoes dbcstrInscricao
        Case 1
            dbcstrInscricao.Tag = "SELECT intContribuinte, " & _
                                  gstrRIGHT("strInscricaoAnterior", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoAnterior " & _
                                  "FROM " & gstrImobiliarioRural & " WHERE intContribuinte IS NOT NULL;strIncricaoAnterior"
            PreencherListaDeOpcoes dbcstrInscricao
        Case 2
            dbcstrInscricao.Tag = "SELECT intContribuinte, " & _
                                  gstrRIGHT("strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricaoCadastral " & _
                                  "FROM " & gstrEconomico & ";strInscricaoCadastral"
            PreencherListaDeOpcoes dbcstrInscricao
        Case 3
            strSql = ""
            strSql = strSql & "SELECT DISTINCT IMO.intContribuinte, "
            strSql = strSql & gstrRIGHT("IMO.strInscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao "
            strSql = strSql & "FROM "
            strSql = strSql & gstrContribuicaoMelhoria & " CM, " & gstrImobiliario & " IMO "
            strSql = strSql & " WHERE IMO.PKId = CM.intImobiliario "
            dbcstrInscricao.Tag = strSql & ";strIncricao"
            PreencherListaDeOpcoes dbcstrInscricao
        Case 4

            dbcstrInscricao.Tag = "SELECT PKId, strNome FROM " & gstrContribuinte & " ORDER BY strNome " & ";strNome"
            PreencherListaDeOpcoes dbcstrInscricao
    End Select

    Exit Sub
End If
    
    
    ToolBarGeral strModoOperacao, gstrEmissaoValidade, mblnAlterando, tdb_Documentos, Me, mobjAux, strQuery
    mblnAlterando = False
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    optbitTipoDeInscricao(intIndiceOPT).Value = True
'    optbitTipoDeInscricao_Click intIndiceOPT
End Sub

Private Function strQuerryEntrega() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao FROM "
    strSql = strSql & gstrOcorrencia
    strSql = strSql & " WHERE intUtilizacaodaOcorrencia = 3 "
    strSql = strSql & " ORDER BY strDescricao"
strQuerryEntrega = strSql
End Function

Private Function strQuery() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT EV.PKId, DE.strDescricao, EV.strNumeroDoProcesso, EV.intContribuinte "
    strSql = strSql & " FROM " & gstrEmissaoValidade & " EV,"
    strSql = strSql & gstrDocumentoEmitido & " DE "
    strSql = strSql & " WHERE EV.intDocumentosEmitidos = DE.PKId "
    ' strSql = strSql & " AND EV.intContribuinte = '" & dbcstrInscricao.BoundText
    strSql = strSql & " AND EV.bitTipoDeInscricao = " & intIndiceOPT
    
    Select Case bytOrdenacao
        Case Is = 1
            strSql = strSql & " ORDER BY DE.strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSql = strSql & " ORDER BY EV.strNumeroDoProcesso" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSql
    
End Function



Private Sub MostraDadosContribuinte(intBound As Long)
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT CO.strBairroC, CO.strLogradouroC, CO.intNumeroC,"
    strSql = strSql & " CO.strComplementoC , CO.intCEPC, CO.strDistritoC, "
    strSql = strSql & " CO.strCNPJCPF, CD.strDescricao, UF.strSigla, CO.strNome "
    strSql = strSql & " FROM "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrCidade & " CD, "
    strSql = strSql & gstrUF & " UF "
    strSql = strSql & "WHERE intMunicipioC = CD.PKId "
    strSql = strSql & " AND intUFC = UF.PKId "
    strSql = strSql & " AND CO.PKId = " & intBound

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not adoResultado.EOF Then
                txt_Bairro = gstrVerificaCampoNulo(!strBairroC)
                txt_Cep = gstrVerificaCampoNulo(!intCepC)
                txt_Complemento = gstrVerificaCampoNulo(!strComplementoC)
                txt_CNPJCPF = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!StrCnpjCpf))
                txt_Distrito = gstrVerificaCampoNulo(!strDistritoC)
                txt_Logradouro = gstrVerificaCampoNulo(!strLogradouroC)
                txt_Municipio = gstrVerificaCampoNulo(!strDescricao)
                txt_Numero = gstrVerificaCampoNulo(!intNumeroC)
                txt_UF = gstrVerificaCampoNulo(!strSigla)
                txt_intContribuinte = gstrVerificaCampoNulo(!STRNOME)
            End If
        End With
    End If
End Sub
Private Sub ProprietarioEconomico()

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset

    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT CO.PKId AS CodigoContribuinte FROM "
'    strSQL = strSQL & gstrEconomico & " AS A, "
    strSql = strSql & gstrEconomico & " A, "
'    strSQL = strSQL & gstrContribuinte & " AS CO "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = A.intContribuinte "
    strSql = strSql & " AND A.intContribuinte = '" & dbcstrInscricao.BoundText & "'"

    LimpaCorrespondencia
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txtintContribuinte.Text = !CodigoContribuinte
                MostraDadosContribuinte Val(txtintContribuinte)
            End With
        End If
    End If

End Sub

Private Sub ProprietarioImobiliario(Index As Integer)

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT CO.PKId AS CodigoContribuinte "
    strSql = strSql & " FROM "
    If Index = 1 Then
'        strSQL = strSQL & gstrImobiliarioRural & " AS A, "
        strSql = strSql & gstrImobiliarioRural & " A, "
    Else
'        strSQL = strSQL & gstrImobiliario & " AS A, "
        strSql = strSql & gstrImobiliario & " A, "
    End If
'    strSQL = strSQL & gstrContribuinte & " AS CO "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = A.intContribuinte "
    strSql = strSql & " AND A.intContribuinte = '" & dbcstrInscricao.BoundText & "'"

    LimpaCorrespondencia
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txtintContribuinte.Text = !CodigoContribuinte
                MostraDadosContribuinte Val(txtintContribuinte)
            End With
        End If
    End If
End Sub

Private Sub ProprietarioContribuicao()

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset

    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT CO.PKId AS CodigoContribuinte FROM "
'    strSQL = strSQL & gstrEconomico & " AS A, "
    strSql = strSql & gstrEconomico & " A, "
'    strSQL = strSQL & gstrContribuinte & " AS CO "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = A.intContribuinte "
    strSql = strSql & " AND A.intContribuinte = '" & dbcstrInscricao.BoundText & "'"

    LimpaCorrespondencia
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txtintContribuinte.Text = !CodigoContribuinte
                MostraDadosContribuinte Val(txtintContribuinte)
            End With
        End If
    End If
End Sub
Sub LimpaCorrespondencia()
    txt_intContribuinte.Text = ""
    txt_Bairro = ""
    txt_Cep = ""
    txt_Complemento = ""
    txt_CNPJCPF = ""
    txt_Distrito = ""
    txt_Logradouro = ""
    txt_Municipio = ""
    txt_Numero = ""
    txt_UF = ""
End Sub

Private Sub optbitTipoDeInscricao_Click(Index As Integer)
    Dim strSql As String
    
    
    Set dbcstrInscricao.RowSource = Nothing
    dbcstrInscricao.Text = ""
    
    intIndiceOPT = Index
    strSql = ""
    lblstrInscricao.Caption = "Inscrição"
    If Index = 4 Then
        lblstrInscricao.Caption = "Código do Contribuinte"
    End If
    dbcstrInscricao.Text = ""
    txtintContribuinte.Text = ""
    txt_intContribuinte.Text = ""
    dbcintDocumentosEmitidos.Text = ""
    txtstrNumeroDoProcesso.Text = ""
    txtdtmSolicitacao.Text = ""
    txtdtmEmissao.Text = ""
    txtdtmConcessao.Text = ""
    txtdtmEntrega.Text = ""
    txtdtmValidade.Text = ""
    dbcintOcorrencia.Text = ""
    LimpaCorrespondencia
    Set tdb_Documentos.DataSource = Nothing
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Function blnDadosOk() As Boolean
blnDadosOk = False
    
    If txtdtmSolicitacao = "" Then
        ExibeMensagem "A data de solicitação tem que ser digitada."
        txtdtmSolicitacao.SetFocus
        Exit Function
    Else
        If gblnDataValida(txtdtmSolicitacao) = False Then
            ExibeMensagem "A data de solicitação não é válida."
            txtdtmSolicitacao.SetFocus
            Exit Function
        End If
    End If
    If txtdtmEmissao = "" Then
        ExibeMensagem "A data de emissão tem que ser digitada."
        txtdtmEmissao.SetFocus
        Exit Function
    Else
        If gblnDataValida(txtdtmEmissao) = False Then
            ExibeMensagem "A data de emissão não é válida."
            txtdtmEmissao.SetFocus
            Exit Function
        End If
    End If
    If txtdtmConcessao = "" Then
        ExibeMensagem "A data de concessão tem que ser digitada."
        txtdtmConcessao.SetFocus
        Exit Function
    Else
        If gblnDataValida(txtdtmConcessao) = False Then
            ExibeMensagem "A data de concessão não é válida."
            txtdtmConcessao.SetFocus
            Exit Function
        End If
    End If
    If txtdtmEntrega = "" Then
        ExibeMensagem "A data de entrega tem que ser digitada."
        txtdtmEntrega.SetFocus
        Exit Function
    Else
        If gblnDataValida(txtdtmEntrega) = False Then
            ExibeMensagem "A data de entrega não é válida."
            txtdtmEntrega.SetFocus
            Exit Function
        End If
    End If
    If txtdtmValidade = "" Then
        ExibeMensagem "A data de validade tem que ser digitada."
        txtdtmValidade.SetFocus
        Exit Function
    Else
        If gblnDataValida(txtdtmValidade) = False Then
            ExibeMensagem "A data de validade não é válida."
            txtdtmValidade.SetFocus
            Exit Function
        End If
    End If
    If dbcstrInscricao.Text = "" Then
        ExibeMensagem " " & lblstrInscricao.Caption & " deve ser selecionado."
        dbcstrInscricao.SetFocus
        Exit Function
    End If
    If txtstrNumeroDoProcesso = "" Then
        ExibeMensagem "O número do processo tem que ser digitado."
        txtstrNumeroDoProcesso.SetFocus
        Exit Function
    End If
    If dbcstrInscricao = "" Then
        ExibeMensagem "Selecione uma inscrição."
        dbcstrInscricao.SetFocus
        Exit Function
    End If
    If dbcintDocumentosEmitidos = "" Then
        ExibeMensagem "Selecione um documento."
        dbcintDocumentosEmitidos.SetFocus
        Exit Function
    End If
    If CVDate(txtdtmEntrega) < CVDate(txtdtmEmissao) Then
        ExibeMensagem "A data de entrega tem que ser maior do que a data de emissão."
        txtdtmEntrega.SetFocus
        Exit Function
    End If
    If CVDate(txtdtmValidade) < CVDate(txtdtmEmissao) Then
        ExibeMensagem "A data de validade tem que ser maior do que a data de emissão."
        txtdtmValidade.SetFocus
        Exit Function
    End If
    If CVDate(txtdtmValidade) < CVDate(txtdtmSolicitacao) Then
        ExibeMensagem "A data de validade tem que ser maior do que a data de solicitação."
        txtdtmValidade.SetFocus
        Exit Function
    End If
    If CVDate(txtdtmEmissao) < CVDate(txtdtmSolicitacao) Then
        ExibeMensagem "A data de emissão tem que ser maior do que a data de solicitação."
        txtdtmEmissao.SetFocus
        Exit Function
    End If
    If CVDate(txtdtmValidade) < CVDate(txtdtmEntrega) Then
        ExibeMensagem "A data de Validade tem que ser maior do que a data de entrega."
        txtdtmValidade.SetFocus
        Exit Function
    End If
    If CVDate(txtdtmValidade) < CVDate(txtdtmConcessao) Then
        ExibeMensagem "A data de Validade tem que ser maior do que a data de concessão."
        txtdtmValidade.SetFocus
        Exit Function
    End If
    If CVDate(txtdtmEntrega) < CVDate(txtdtmConcessao) Then
        ExibeMensagem "A data de Entrega tem que ser maior do que a data de concessão."
        txtdtmEntrega.SetFocus
        Exit Function
    End If
    If CVDate(txtdtmEmissao) < CVDate(txtdtmConcessao) Then
        ExibeMensagem "A data de Concessao tem que ser maior do que a data de emissão."
        txtdtmEmissao.SetFocus
        Exit Function
    End If

blnDadosOk = True
End Function

Private Sub txt_Bairro_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Bairro
End Sub

Private Sub txt_Cep_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Cep
End Sub

Private Sub txt_Complemento_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Complemento
End Sub

Private Sub txt_CNPJCPF_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txt_CNPJCPF
End Sub

Private Sub txt_Distrito_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Distrito

End Sub

Private Sub txt_Logradouro_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Logradouro
End Sub


Private Sub txt_Numero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_Numero

End Sub

Private Sub txt_UF_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_UF
End Sub

Private Sub txtdtmConcessao_LostFocus()
    txtdtmConcessao.Text = gstrDataFormatada(txtdtmConcessao.Text)
End Sub

Private Sub txtdtmEmissao_LostFocus()
    txtdtmEmissao.Text = gstrDataFormatada(txtdtmEmissao.Text)
End Sub

Private Sub txtdtmEntrega_LostFocus()
    txtdtmEntrega.Text = gstrDataFormatada(txtdtmEntrega.Text)
End Sub

Private Sub txtdtmSolicitacao_GotFocus()
    MarcaCampo txtdtmSolicitacao
End Sub

Private Sub txtdtmSolicitacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmSolicitacao
End Sub

Private Sub txtdtmEmissao_GotFocus()
    MarcaCampo txtdtmEmissao
End Sub

Private Sub txtdtmEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmEmissao
End Sub

Private Sub txtdtmConcessao_GotFocus()
    MarcaCampo txtdtmConcessao
End Sub

Private Sub txtdtmConcessao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmConcessao
End Sub

Private Sub txtdtmEntrega_GotFocus()
    MarcaCampo txtdtmEntrega
End Sub

Private Sub txtdtmEntrega_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmEntrega
End Sub

Private Sub txtdtmSolicitacao_LostFocus()
    txtdtmSolicitacao.Text = gstrDataFormatada(txtdtmSolicitacao.Text)
End Sub

Private Sub txtdtmValidade_GotFocus()
    MarcaCampo txtdtmValidade
End Sub

Private Sub txtdtmValidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmValidade
End Sub

Private Sub txtdtmValidade_LostFocus()
    txtdtmValidade.Text = gstrDataFormatada(txtdtmValidade.Text)
End Sub

'0912

Private Sub txtstrNumeroDoProcesso_GotFocus()
    MarcaCampo txtstrNumeroDoProcesso
End Sub

Private Sub txtstrNumeroDoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNumeroDoProcesso
End Sub
