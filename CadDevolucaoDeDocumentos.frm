VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadDevolucaoDeDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devolução de Documentos"
   ClientHeight    =   7035
   ClientLeft      =   1620
   ClientTop       =   2385
   ClientWidth     =   8910
   HelpContextID   =   11
   Icon            =   "CadDevolucaoDeDocumentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8910
   Begin VB.TextBox txtPKId 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   420
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "txtPKId"
      Top             =   1380
      Visible         =   0   'False
      Width           =   795
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6855
      Left            =   105
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Devolução de Documentos"
      TabPicture(0)   =   "CadDevolucaoDeDocumentos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_DocumentosDevolvidos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Bla"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Data"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_Endereco"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Inscricao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fra_Inscricao 
         Height          =   645
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   8445
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Receitas Diversas"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   4
            Left            =   6690
            TabIndex        =   5
            Top             =   270
            Width           =   1605
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Contribuição de Melhorias"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   3
            Left            =   4470
            TabIndex        =   4
            Top             =   270
            Width           =   2205
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Econômico"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   2
            Left            =   3270
            TabIndex        =   3
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Rural"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   1
            Left            =   1770
            TabIndex        =   2
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   270
            Value           =   -1  'True
            Width           =   1605
         End
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   "Endereço de Correspondência"
         Height          =   1395
         Left            =   120
         TabIndex        =   30
         Top             =   2820
         Width           =   8445
         Begin VB.TextBox txt_Distrito 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   15
            Top             =   960
            Width           =   3525
         End
         Begin VB.TextBox txt_Bairro 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   13
            Top             =   600
            Width           =   3105
         End
         Begin VB.TextBox txt_Logradouro 
            Height          =   285
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   10
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox txt_Numero 
            Height          =   285
            Left            =   5790
            MaxLength       =   8
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txt_Complemento 
            Height          =   285
            Left            =   7260
            MaxLength       =   20
            TabIndex        =   12
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox txt_Cep 
            Height          =   285
            Left            =   7230
            MaxLength       =   9
            TabIndex        =   17
            Top             =   960
            Width           =   1080
         End
         Begin VB.TextBox txt_Municipio 
            Height          =   285
            Left            =   5145
            MaxLength       =   50
            TabIndex        =   14
            Top             =   600
            Width           =   3165
         End
         Begin VB.TextBox txt_UF 
            Height          =   285
            Left            =   6120
            MaxLength       =   2
            TabIndex        =   16
            Top             =   960
            Width           =   510
         End
         Begin VB.Label lblstrDistritoC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   450
            TabIndex        =   38
            Top             =   1050
            Width           =   480
         End
         Begin VB.Label lblintMunicipioC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   4290
            TabIndex        =   37
            Top             =   690
            Width           =   705
         End
         Begin VB.Label lblintBairroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   525
            TabIndex        =   36
            Top             =   690
            Width           =   405
         End
         Begin VB.Label lblintLogradouroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lblintNumeroC 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5520
            TabIndex        =   34
            Top             =   330
            Width           =   180
         End
         Begin VB.Label lblstrComplementoC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6720
            TabIndex        =   33
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lblintUFC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   5760
            TabIndex        =   32
            Top             =   1065
            Width           =   210
         End
         Begin VB.Label lblintCepC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   6795
            TabIndex        =   31
            Top             =   1065
            Width           =   285
         End
      End
      Begin VB.Frame fra_Data 
         Height          =   645
         Left            =   120
         TabIndex        =   27
         Top             =   4230
         Width           =   8445
         Begin VB.TextBox txtdtmDevolucao 
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
            Left            =   2790
            MaxLength       =   10
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dbcintOcorrencia 
            Height          =   315
            Left            =   5325
            TabIndex        =   19
            Top             =   210
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblintOcorrencia 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ocorrências"
            Height          =   195
            Left            =   4335
            TabIndex        =   29
            Top             =   315
            Width           =   855
         End
         Begin VB.Label lbldtmDevolucao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data de devolução do documento"
            Height          =   195
            Left            =   210
            TabIndex        =   28
            Top             =   315
            Width           =   2430
         End
      End
      Begin VB.Frame fra_Bla 
         Height          =   1755
         Left            =   120
         TabIndex        =   22
         Top             =   1020
         Width           =   8445
         Begin VB.TextBox txt_CNPJCPF 
            Height          =   285
            Left            =   2010
            MaxLength       =   50
            TabIndex        =   8
            Top             =   960
            Width           =   2505
         End
         Begin MSDataListLib.DataCombo dbcintDocumentosEmitidos 
            Height          =   315
            Left            =   2010
            TabIndex        =   9
            Top             =   1320
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintContribuinte 
            Height          =   315
            Left            =   2010
            TabIndex        =   7
            Top             =   570
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskstrInscricao 
            Height          =   285
            Left            =   2010
            TabIndex        =   6
            Top             =   210
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin VB.Label lbl_CNPJCPF 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CPF / CNPJ"
            Height          =   195
            Left            =   975
            TabIndex        =   26
            Top             =   1035
            Width           =   870
         End
         Begin VB.Label lblstrInscricao 
            Alignment       =   1  'Right Justify
            Caption         =   "Inscrição"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   285
            Width           =   1755
         End
         Begin VB.Label lblintDocumentosEmitidos 
            AutoSize        =   -1  'True
            Caption         =   "Documentos emitidos"
            Height          =   195
            Left            =   330
            TabIndex        =   24
            Top             =   1425
            Width           =   1515
         End
         Begin VB.Label lblintContribuinte 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte"
            Height          =   195
            Left            =   1035
            TabIndex        =   23
            Top             =   675
            Width           =   840
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_DocumentosDevolvidos 
         Height          =   1695
         Left            =   120
         TabIndex        =   20
         Top             =   4980
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   2990
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
         Columns(1).Caption=   "Documentos já controlados para o contribuinte"
         Columns(1).DataField=   "strDocumento"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data de Devolução"
         Columns(2).DataField=   "dtmDevolucao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Ocorrência"
         Columns(3).DataField=   "strOcorrencia"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=6482"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6403"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2699"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2619"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=5159"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=5080"
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
Attribute VB_Name = "frmCadDevolucaoDeDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim mblnAlterando                   As Boolean
Dim mobjAux                         As Object
Dim mblnSelecionou                  As Boolean
Dim mblnPrimeiraVez                 As Boolean
Dim intIndiceOPT                    As Integer
Dim adoResultado                    As ADODB.Recordset

Private Sub dbcintContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinte_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintContribuinte
End Sub

Private Sub dbcintDocumentosEmitidos_Click(Area As Integer)
   DropDownDataCombo dbcintDocumentosEmitidos, Me, Area
End Sub

Private Sub dbcintDocumentosEmitidos_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintDocumentosEmitidos, Me, , KeyCode, Shift
End Sub

Private Sub dbcintDocumentosEmitidos_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", dbcintDocumentosEmitidos
End Sub

Private Sub dbcintOcorrencia_Click(Area As Integer)
   DropDownDataCombo dbcintOcorrencia, Me, Area
End Sub

Private Sub dbcintOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrencia_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", dbcintOcorrencia
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 633
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
    
    dbcintContribuinte.Tag = "SELECT DISTINCT IMO.intContribuinte, CON.strNome FROM " & gstrImobiliario & " IMO, " & gstrContribuinte & " CON WHERE CON.PKId = IMO.intContribuinte ORDER BY CON.strNome;CON.strNome"
    
    LeDaTabelaParaObj gstrOcorrencia, dbcintOcorrencia, strQuerryEntrega
    TrocaCorObjeto mskstrInscricao, True
    
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
    
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricao
End Sub

Private Sub tdb_DocumentosDevolvidos_Click()
    mblnPrimeiraVez = True
    'tdb_DocumentosDevolvidos_RowColChange 0, 0
End Sub

Sub tdb_DocumentosDevolvidos_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_DocumentosDevolvidos_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_DocumentosDevolvidos
End Sub

Private Sub tdb_DocumentosDevolvidos_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_DocumentosDevolvidos, ColIndex
End Sub

Private Sub tdb_DocumentosDevolvidos_KeyPress(KeyAscii As Integer)
    If tdb_DocumentosDevolvidos.Col = 2 Then
        CaracterValido KeyAscii, "D", tdb_DocumentosDevolvidos
    Else
        CaracterValido KeyAscii, "A", tdb_DocumentosDevolvidos
    End If
End Sub

Private Sub tdb_DocumentosDevolvidos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_DocumentosDevolvidos
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value

                LeDaTabelaParaObj gstrDevolucao, Me
                If dbcintContribuinte.BoundText <> "" Then
                    MostraDadosContribuinte (dbcintContribuinte.BoundText)
                End If
                'dbcintContribuinte_Click 2
                gCorLinhaSelecionada tdb_DocumentosDevolvidos

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                'mblnPrimeiraVez = False
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
Dim strSQL      As String
Dim strInscricao As String
Dim intIndice As Integer
Dim intSelecionado As Integer
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If blnDadosOk = False Then
            Exit Sub
        End If
    End If
    
    For intIndice = 0 To 4
        If optbitTipoDeInscricao(intIndice).Value Then
            strSQL = strQuery(intIndice)
            intSelecionado = intIndice
        End If
    Next

    If UCase(strModoOperacao) = UCase(gstrSalvar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
        mblnPrimeiraVez = False
    End If
    
    If strModoOperacao = gstrPreencherLista Then
        Select Case intSelecionado
            Case 0
                strSQL = "SELECT DISTINCT IMO.intContribuinte, CON.strNome FROM " & gstrImobiliario & " IMO, " & gstrContribuinte & " CON WHERE CON.PKId = IMO.intContribuinte ORDER BY CON.strNome"
            Case 1
                strSQL = "SELECT DISTINCT IMORU.intContribuinte, CON.strNome FROM " & gstrImobiliarioRural & " IMORU, " & gstrContribuinte & " CON WHERE CON.PKId = IMORU.intContribuinte ORDER BY CON.strNome"
            Case 2
                strSQL = "SELECT DISTINCT ECO.intContribuinte, CON.strNome FROM " & gstrEconomico & " ECO, " & gstrContribuinte & " CON WHERE CON.PKId = ECO.intContribuinte ORDER BY CON.strNome"
            Case 3
                strSQL = "SELECT DISTINCT IMO.intContribuinte, CON.strNome FROM " & gstrContribuicaoMelhoria & " CM, " & gstrImobiliario & " IMO, " & gstrContribuinte & " CON WHERE IMO.PKId = CM.intImobiliario  AND CON.PKId = IMO.intContribuinte ORDER BY CON.strNome"
            Case 4
                strSQL = "SELECT DISTINCT REC.intContribuinte, CON.strNome FROM " & gstrReceitaDiversa & " REC, " & gstrContribuinte & " CON WHERE CON.PKId = REC.intContribuinte ORDER BY CON.strNome"
        End Select
        strSQL = strSQL & ";CON.strNome"
        dbcintContribuinte.Tag = strSQL
        PreencherListaDeOpcoes dbcintContribuinte
        Exit Sub
    End If
    
    strInscricao = mskstrInscricao.Text
    ToolBarGeral strModoOperacao, gstrDevolucao, mblnAlterando, tdb_DocumentosDevolvidos, Me, mobjAux, strSQL
    mskstrInscricao.Text = strInscricao
    mblnAlterando = False
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar

End Sub

Private Function strQuerryEntrega() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao FROM "
    strSQL = strSQL & gstrOcorrencia
    strSQL = strSQL & " WHERE intUtilizacaodaOcorrencia = 3 "
    strSQL = strSQL & " ORDER BY strDescricao"
strQuerryEntrega = strSQL
End Function

Private Function strQuery(intTipoDeInscricao As Integer) As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT EV.PKId, DE.strDescricao AS strDocumento, EV.dtmDevolucao, OCO.strDescricao AS strOcorrencia "
'    strSQL = strSQL & " FROM " & gstrDevolucao & " AS EV,"
    strSQL = strSQL & " FROM " & gstrDevolucao & " EV,"
    strSQL = strSQL & gstrDocumentoEmitido & " DE, "
    strSQL = strSQL & gstrOcorrencia & " OCO "
    strSQL = strSQL & " WHERE EV.intDocumentosEmitidos = DE.PKId "
    strSQL = strSQL & " AND OCO.PKId = EV.intOcorrencia "
    If dbcintContribuinte.MatchedWithList Then
        strSQL = strSQL & " AND EV.intContribuinte = " & dbcintContribuinte.BoundText
    End If
    strSQL = strSQL & " AND EV.bitTipoDeInscricao = " & intTipoDeInscricao

    strSQL = strSQL & " ORDER BY DE.strDescricao"

    strQuery = strSQL
    
End Function

Private Function strQueryDocuementos(lngCodContribuinte As Long) As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT A.PKId, A.strDescricao FROM "
    strSQL = strSQL & gstrDocumentoEmitido & " A, "
    strSQL = strSQL & gstrEmissaoValidade & " B "
    strSQL = strSQL & " WHERE A.PKId = B.intDocumentosEmitidos"
    strSQL = strSQL & " AND B.intContribuinte = " & lngCodContribuinte
    
    strQueryDocuementos = strSQL
    
End Function

Private Sub dbcintContribuinte_Click(Area As Integer)
Dim intIndice As Integer
Dim strTabela As String
Dim strCampo As String
Dim strSQL As String
Dim ADOTemp As ADODB.Recordset
    
    DropDownDataCombo dbcintContribuinte, Me, Area
    
    If Area = 2 Then
        If dbcintContribuinte.BoundText <> "" Then
            mblnPrimeiraVez = False
            If MostraDadosContribuinte(dbcintContribuinte.BoundText) = False Then
                LimpaCorrespondencia
            End If
        End If
        If lblstrInscricao.Caption = "Código do Contribuinte" Then
            mskstrInscricao = dbcintContribuinte.BoundText
        Else
            For intIndice = 0 To 3
                If optbitTipoDeInscricao(intIndice).Value Then
                    Select Case intIndice
                        Case 0, 3
                            strTabela = gstrImobiliario
                            strCampo = "strInscricao"
                        Case 1
                            strTabela = gstrImobiliarioRural
                            strCampo = "strInscricao"
                        Case 2
                            strTabela = gstrEconomico
                            strCampo = "strInscricaoCadastral"
                    End Select
                    strSQL = "SELECT DISTINCT  " & gstrRIGHT(strCampo, gintRetornaTamanhoMascara(IIf(strTabela = gstrEconomico, TYP_ECONOMICA, TYP_IMOBILIARIA))) & " " & strCampo & " FROM " & strTabela & " WHERE intContribuinte = " & dbcintContribuinte.BoundText
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strSQL, 5, ADOTemp) Then
                        VerificaMascaraInscricao intIndice
                        mskstrInscricao.Text = gstrENulo(ADOTemp.Fields(strCampo).Value)
                    End If
                    dbcintDocumentosEmitidos.BoundText = 0
                    dbcintOcorrencia.BoundText = 0
                    txtdtmDevolucao = ""
                    mblnPrimeiraVez = False
                    mblnAlterando = False
                    LeDaTabelaParaObj "", tdb_DocumentosDevolvidos, strQuery(intIndice)
                    
                End If
            Next
        End If
        LeDaTabelaParaObj gstrDocumentoEmitido, dbcintDocumentosEmitidos, strQueryDocuementos(dbcintContribuinte.BoundText)
    End If
End Sub
Private Function MostraDadosContribuinte(intBound As Long) As Boolean
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT CO.strBairroC, CO.strLogradouroC, CO.intNumeroC,"
    strSQL = strSQL & " CO.strComplementoC , CO.intCEPC, CO.strDistritoC, "
    strSQL = strSQL & " CO.strCNPJCPF, CD.strDescricao, UF.strSigla "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrContribuinte & " CO, "
    strSQL = strSQL & gstrCidade & " CD, "
    strSQL = strSQL & gstrUF & " UF "
    strSQL = strSQL & "WHERE intMunicipioC = CD.PKId "
    strSQL = strSQL & " AND intUFC = UF.PKId "
    strSQL = strSQL & " AND CO.PKId = " & intBound
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                txt_Bairro = gstrVerificaCampoNulo(!strBairroC)
                txt_Cep = gstrVerificaCampoNulo(!intCepC)
                txt_Complemento = gstrVerificaCampoNulo(!strComplementoC)
                txt_CNPJCPF = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!StrCnpjCpf))
                txt_Distrito = gstrVerificaCampoNulo(!strDistritoC)
                txt_Logradouro = gstrVerificaCampoNulo(!strLogradouroC)
                txt_Municipio = gstrVerificaCampoNulo(!strDescricao)
                txt_Numero = gstrVerificaCampoNulo(!intNumeroC)
                txt_UF = gstrVerificaCampoNulo(!strSigla)
                MostraDadosContribuinte = True
                .MoveNext
            Loop
        End With
    End If
End Function

Sub LimpaCorrespondencia()
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
    Dim intIndice As Integer
    Dim strSQL As String
    
    Set dbcintContribuinte.RowSource = Nothing
    dbcintContribuinte.Text = ""
    
    optbitTipoDeInscricao(Index).CausesValidation = True
    For intIndice = 0 To 4
        If intIndice <> Index Then
            optbitTipoDeInscricao(intIndice).CausesValidation = False
        End If
    Next
     
    Set tdb_DocumentosDevolvidos.DataSource = Nothing
    mblnPrimeiraVez = False
    mblnAlterando = False
    
    VerificaMascaraInscricao Index
    
    'dbcintContribuinte.BoundText = 0
    mskstrInscricao = ""
    LimpaCorrespondencia
    dbcintDocumentosEmitidos.BoundText = 0
    dbcintOcorrencia.BoundText = 0
    txtdtmDevolucao = ""
    
    If optbitTipoDeInscricao(Index) Then
       If Index <> 4 Then
            lblstrInscricao.Caption = "Inscrição"
       Else
            lblstrInscricao.Caption = "Código do Contribuinte"
       End If
    End If
    
End Sub

Private Function blnMostraContribuinte(intCodigo As Variant) As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT COUNT(*) as Contador FROM " & gstrContribuinte
    strSQL = strSQL & " WHERE PKId = " & intCodigo
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If adoResultado!Contador <= 0 Then
                blnMostraContribuinte = False
                Exit Function
            Else
                dbcintContribuinte.BoundText = intCodigo
                If MostraDadosContribuinte(dbcintContribuinte.BoundText) = False Then LimpaCorrespondencia
                blnMostraContribuinte = True
                Exit Function
            End If
        End If
    End If
End Function

Private Sub mskstrInscricao_LostFocus()
    If lblstrInscricao.Caption = "Código do Contribuinte" Then
       If mskstrInscricao.Text <> "" Then
           If blnMostraContribuinte(mskstrInscricao.Text) = False Then
              ExibeMensagem "Contribuinte não encontrado." & Chr(13) & "Digite outro código ou selecione um contribuinte."
              mskstrInscricao = ""
              Exit Sub
           End If
        End If
     ElseIf lblstrInscricao.Caption = "Inscrição" Then
        If mskstrInscricao <> "" Then
            If blnDuplicataInscricao(mskstrInscricao.Text) = False Then
                ExibeMensagem "Inscrição não encontrada"
                mskstrInscricao = ""
                Exit Sub
            End If
        End If
    End If
End Sub

Function blnDuplicataInscricao(strInscricao As String) As Boolean
Dim strSQL       As String
    If strInscricao = "" Then
        blnDuplicataInscricao = False
        Exit Function
    End If
    
    strInscricao = String(gintLenInscricao - Len(strInscricao), "0") & strInscricao
    
    If intIndiceOPT = 0 Or intIndiceOPT = 3 Then
        strSQL = ""
        strSQL = strSQL & "SELECT count(*) as Contador, intContribuinte Codigo FROM " & gstrImobiliario
        strSQL = strSQL & " WHERE strInscricaoAnterior = '" & strInscricao & "'"
        strSQL = strSQL & " GROUP BY intContribuinte "
    ElseIf intIndiceOPT = 1 Then
        strSQL = ""
        strSQL = strSQL & "SELECT count(*) as Contador, intContribuinte Codigo FROM " & gstrImobiliarioRural
        strSQL = strSQL & " WHERE strInscricaoAnterior = '" & strInscricao & "'"
        strSQL = strSQL & " GROUP BY intContribuinte "
    ElseIf intIndiceOPT = 2 Then
        strSQL = ""
        strSQL = strSQL & "SELECT count(*) as Contador, intContribuinte Codigo FROM " & gstrEconomico
        strSQL = strSQL & " WHERE strInscricaoCadastral = '" & strInscricao & "'"
        strSQL = strSQL & " GROUP BY intContribuinte "
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If adoResultado!Contador <= 0 Then
                blnDuplicataInscricao = False
                Exit Function
            Else
                dbcintContribuinte.BoundText = adoResultado!Codigo
                MostraDadosContribuinte (dbcintContribuinte.BoundText)
                blnDuplicataInscricao = True
                Exit Function
            End If
        End If
    End If
End Function

Sub VerificaMascaraInscricao(Index As Integer)
Dim strSQL       As String
Dim strMascara   As String
strMascara = ""
    If Index = 0 Or Index = 3 Then
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM " & gstrCampoDeInscricao & " "
        strSQL = strSQL & " WHERE intTipoDeInscricao = " & TYP_IMOBILIARIA
    ElseIf Index = 1 Then
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM " & gstrCampoDeInscricao & " "
        strSQL = strSQL & " WHERE intTipoDeInscricao = " & TYP_ECONOMICA
    ElseIf Index = 2 Then
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM " & gstrCampoDeInscricao & " "
        strSQL = strSQL & " WHERE intTipoDeInscricao = " & TYP_DIVIDA_ATIVA
    ElseIf Index = 4 Then
        strSQL = ""
        mskstrInscricao.Mask = "#########"
        Exit Sub
    End If
        strSQL = strSQL & " ORDER BY intSequencia"

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
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

    If txtdtmDevolucao = "" Then
        ExibeMensagem "A data de devolução tem que ser digitada."
        txtdtmDevolucao.SetFocus
        Exit Function
    Else
        If gblnDataValida(txtdtmDevolucao.Text) = False Then
            ExibeMensagem "A data de devolução não é válida."
            txtdtmDevolucao.SetFocus
            Exit Function
        End If
    End If
    If mskstrInscricao = "" Then
        ExibeMensagem " " & lblstrInscricao.Caption & " tem que ser digitado."
        If mskstrInscricao.Enabled Then mskstrInscricao.SetFocus
        Exit Function
    End If
    If dbcintContribuinte = "" Then
        ExibeMensagem "Selecione um contribuinte."
        dbcintContribuinte.SetFocus
        Exit Function
    End If
    If dbcintDocumentosEmitidos = "" Then
        ExibeMensagem "Selecione um documento."
        dbcintDocumentosEmitidos.SetFocus
        Exit Function
    End If
    
blnDadosOk = True
End Function

Private Sub mskstrInscricao_GotFocus()
    MarcaCampo mskstrInscricao
End Sub

Private Sub txt_Municipio_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Municipio
End Sub

Private Sub txt_Numero_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txt_Numero
End Sub

Private Sub txt_UF_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_UF
End Sub

Private Sub txtdtmDevolucao_GotFocus()
    MarcaCampo txtdtmDevolucao
End Sub

Private Sub txtdtmDevolucao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDevolucao
End Sub

Private Sub txtdtmDevolucao_LostFocus()
    txtdtmDevolucao.Text = gstrDataFormatada(txtdtmDevolucao.Text)
End Sub

Private Sub txt_Bairro_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Bairro
End Sub

Private Sub txt_Cep_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "E", txt_Cep
End Sub

Private Sub txt_Complemento_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Complemento
End Sub

Private Sub txt_CNPJCPF_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_CNPJCPF
End Sub

Private Sub txt_Distrito_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Distrito

End Sub

Private Sub txt_Logradouro_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_Logradouro
End Sub

