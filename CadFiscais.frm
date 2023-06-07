VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadFiscais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fiscais"
   ClientHeight    =   6630
   ClientLeft      =   1725
   ClientTop       =   2430
   ClientWidth     =   9225
   HelpContextID   =   45
   Icon            =   "CadFiscais.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9225
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6480
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   11430
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Fiscais"
      TabPicture(0)   =   "CadFiscais.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_CNPJCPF"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintContribuinte"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_PKId"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcintContribuinte"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_Fiscais"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt_CNPJCPF"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmd_Contribuinte"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPKId"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tab_3DCorrespondencia"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_TipoFiscal"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkblnFiscalProdutividade"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.CheckBox chkblnFiscalProdutividade 
         Caption         =   "Produtividade em Folha"
         Height          =   375
         Left            =   6210
         TabIndex        =   2
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Frame fra_TipoFiscal 
         Caption         =   "Tipo de Fiscal"
         Height          =   1095
         Left            =   6210
         TabIndex        =   28
         Top             =   480
         Width           =   2775
         Begin VB.OptionButton optbytTipoFiscal 
            Caption         =   "Fiscal de Posturas"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton optbytTipoFiscal 
            Caption         =   "Fiscal de Tributos"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1815
         End
      End
      Begin TabDlg.SSTab tab_3DCorrespondencia 
         Height          =   2055
         Left            =   90
         TabIndex        =   11
         Top             =   1920
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   3625
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Endereço de correspondência"
         TabPicture(0)   =   "CadFiscais.frx":105E
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
         Tab(0).Control(8)=   "lbl_strLogradouroC"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lbl_intNumeroC"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lbl_strcomplementoC"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lbl_strBairroC"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "lbl_strMunicipioC"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "lbl_strUFC"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "lbl_intCepC"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "lbl_strDistritoC"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         Begin VB.Label lbl_strDistritoC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1080
            TabIndex        =   27
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label lbl_intCepC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6270
            TabIndex        =   26
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lbl_strUFC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4770
            TabIndex        =   25
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lbl_strMunicipioC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1080
            TabIndex        =   24
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label lbl_strBairroC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1080
            TabIndex        =   23
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label lbl_strcomplementoC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7530
            TabIndex        =   22
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label lbl_intNumeroC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5970
            TabIndex        =   21
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lbl_strLogradouroC 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1080
            TabIndex        =   20
            Top             =   480
            Width           =   4515
         End
         Begin VB.Label lblintCepC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   5865
            TabIndex        =   19
            Top             =   1290
            Width           =   285
         End
         Begin VB.Label lblintUFC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   4470
            TabIndex        =   18
            Top             =   1290
            Width           =   210
         End
         Begin VB.Label lblstrComplementoC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6960
            TabIndex        =   17
            Top             =   570
            Width           =   480
         End
         Begin VB.Label lblintNumeroC 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5700
            TabIndex        =   16
            Top             =   570
            Width           =   180
         End
         Begin VB.Label lblintLogradouroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   570
            Width           =   810
         End
         Begin VB.Label lblintBairroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   510
            TabIndex        =   14
            Top             =   900
            Width           =   405
         End
         Begin VB.Label lblintMunicipioC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   225
            TabIndex        =   13
            Top             =   1280
            Width           =   705
         End
         Begin VB.Label lblstrDistritoC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   1650
            Width           =   480
         End
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   450
         Width           =   1170
      End
      Begin VB.CommandButton cmd_Contribuinte 
         Height          =   315
         Left            =   5610
         Picture         =   "CadFiscais.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Ativa Cadastro de Contribuintes"
         Top             =   810
         Width           =   360
      End
      Begin VB.TextBox txt_CNPJCPF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   19
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   1600
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Fiscais 
         Height          =   2205
         Left            =   120
         TabIndex        =   3
         Top             =   4110
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   3889
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "Codigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nome"
         Columns(2).DataField=   "strNome"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tipo de Fiscal"
         Columns(3).DataField=   "strTipoFiscal"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2619"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2540"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=9155"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=9075"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3201"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33"
         _StyleDefs(63)  =   "Named:id=41:RecordSelector"
         _StyleDefs(64)  =   ":id=41,.parent=34"
         _StyleDefs(65)  =   "Named:id=42:FilterBar"
         _StyleDefs(66)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintContribuinte 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   810
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lbl_PKId 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   555
         TabIndex        =   10
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lblintContribuinte 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   630
         TabIndex        =   9
         Top             =   930
         Width           =   420
      End
      Begin VB.Label lbl_CNPJCPF 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ / CPF"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1290
         Width           =   870
      End
   End
   Begin VB.Menu mnu_TipoComunicacao 
      Caption         =   "mnu_TipoComunicacao"
      Visible         =   0   'False
      Begin VB.Menu mnu_Deletar 
         Caption         =   "Deletar"
      End
      Begin VB.Menu mnu_Traco 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Lista 
         Caption         =   "Lista"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmCadFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim mblnAlterando    As Boolean
Dim mobjAux          As Object
Dim oList            As Object
    
Dim mblnSelecionou   As Boolean
Dim mblnPrimeiraVez  As Boolean

Private Sub dbcintContribuinte_Click(Area As Integer)

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    DropDownDataCombo dbcintContribuinte, Me, Area
    
    If Area = 2 Then
        If dbcintContribuinte.BoundText = "" Then
            Exit Sub
        End If
        strSQL = ""
        strSQL = strSQL & "Select strCNPJCPF, strLogradouroC, strBairroC, intMunicipioC, IntNumeroC, "
        strSQL = strSQL & "strComplementoC, intCepC, strDistritoC, intUFC, strSigla, strdescricao "
        strSQL = strSQL & "From " & gstrContribuinte & " A, " & gstrUF & " B, " & gstrCidade & " C "
'        strSql = strSql & " Where B.Pkid =* A.intUFC "
        strSQL = strSQL & " Where B.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " A.intUFC "
'        strSql = strSql & " AND C.Pkid =* A.intMunicipioC "
        strSQL = strSQL & " AND C.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " A.intMunicipioC "
        strSQL = strSQL & " AND A.PKId = " & gstrItemData(dbcintContribuinte)
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                With adoResultado
                    txt_CNPJCPF = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!StrCnpjCpf))
                    lbl_strLogradouroC = gstrVerificaCampoNulo(!strLogradouroC)
                    lbl_strBairroC = gstrVerificaCampoNulo(!strBairroC)
                    lbl_strMunicipioC = gstrVerificaCampoNulo(!strDescricao)
                    lbl_strUFC = gstrVerificaCampoNulo(!strSigla)
                    lbl_intNumeroC = gstrVerificaCampoNulo(!intNumeroC)
                    lbl_strComplementoC = gstrVerificaCampoNulo(!strComplementoC)
                    lbl_intCepC = gstrCEPFormatado(gstrVerificaCampoNulo(!intCepC))
                    lbl_strDistritoC = gstrVerificaCampoNulo(!strDistritoC)
                End With
            Else
                txt_CNPJCPF = ""
                lbl_strLogradouroC = ""
                lbl_strBairroC = ""
                lbl_strMunicipioC = ""
                lbl_intCepC = ""
                lbl_intNumeroC = ""
                lbl_strComplementoC = ""
                lbl_intCepC = ""
                lbl_strDistritoC = ""
            End If
        End If
    End If
End Sub

Private Sub cmd_Contribuinte_Click()
    ChamaFormCadastro frmCadContribuinte, dbcintContribuinte
End Sub

Private Sub dbcintContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 607
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

Private Sub Form_Load()
    dbcintContribuinte.Tag = strQueryDataComboContribuinte & ";strNome"
    VerificaListaAutomatica gstrFiscais, tdb_Fiscais, strQueryFiscais
    PreencheMenuPopup
    VerificaObjParaAplicar mobjAux
End Sub

Private Function strQueryDataComboContribuinte()
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strNome "
    strSQL = strSQL & "FROM " & gstrContribuinte & " "
    strSQL = strSQL & "ORDER BY strNome"
    strQueryDataComboContribuinte = strSQL
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_Fiscais_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_Fiscais_FilterChange()
    gblnFilraCampos tdb_Fiscais
    With tdb_Fiscais
       If Not .BOF And Not .EOF Then
           If .Bookmark = 1 Then
               tdb_Fiscais_RowColChange 0, 0
          End If
       End If
    End With
End Sub

Private Sub tdb_Fiscais_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Fiscais, ColIndex
End Sub

Private Sub tdb_Fiscais_KeyPress(KeyAscii As Integer)
    Select Case tdb_Fiscais.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_Fiscais
        Case Else
            CaracterValido KeyAscii, "A", tdb_Fiscais
    End Select
End Sub

Private Sub tdb_Fiscais_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Fiscais
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKID.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrFiscais, Me
                dbcintContribuinte_Click 2
                 gCorLinhaSelecionada tdb_Fiscais
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
    Dim strSQL As String
    Dim intCodFiscais As Integer
    
    strSQL = ""
    
    If mblnAlterando Then
        intCodFiscais = tdb_Fiscais.Columns("PKID").Value
    Else
        intCodFiscais = IsNull(txtPKID)
    End If
    
    strSQL = strQueryFiscais
    
    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If
    
    Select Case UCase(strModoOperacao)
        Case "NOVO"
            LimpaObjeto Me, mblnAlterando
            NovoFiscal
            
        Case "SALVAR"
            If blnDadosOk Then
                If ToolBarGeral(strModoOperacao, gstrFiscais, mblnAlterando, tdb_Fiscais, Me, mobjAux, strSQL) Then
                    NovoFiscal
                End If
                
            End If
            
        Case "DELETAR"
            If ToolBarGeral(strModoOperacao, gstrFiscais, mblnAlterando, tdb_Fiscais, Me, mobjAux, strSQL) Then
                NovoFiscal
            End If
        Case Is = UCase(gstrImprimir)
            ImprimeRelatorio rptCadFiscais, strQueryRelatorio
        Case Else
            ToolBarGeral strModoOperacao, gstrFiscais, mblnAlterando, tdb_Fiscais, Me, mobjAux, strSQL
            
    End Select
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
End Sub

Private Function strQueryFiscais() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    strSQL = ""
'    strSql = strSql & "Select C.PKId, C.PKId Codigo, CO.strNome, CASE C.bytTipoFiscal "
    strSQL = strSQL & "Select C.PKId, C.PKId Codigo, CO.strNome, "
'    strSql = strSql & "WHEN 0 THEN 'Fiscal de Tributos' ELSE 'Fiscal de Posturas' END AS strTipoFiscal From " & gstrFiscais & " AS C, " & gstrContribuinte & " CO "
    strSQL = strSQL & gstrCASEWHEN("C.bytTipoFiscal", "0, 'Fiscal de Tributos'", "'Fiscal de Posturas'") & " AS strTipoFiscal From " & gstrFiscais & " C, " & gstrContribuinte & " CO "
    strSQL = strSQL & "Where C.intContribuinte = CO.PKId "
    strSQL = strSQL & "Order By C.PKId"
    strQueryFiscais = strSQL
End Function

Private Sub txt_CNPJCPF_GotFocus()
    MarcaCampo txt_CNPJCPF
End Sub

Private Sub txtPKId_GotFocus()
    MarcaCampo txtPKID
End Sub

Sub PreencheMenuPopup()
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    Dim intI         As Integer
    
    On Error GoTo Err_Handle
    intI = 0
    
    strSQL = ""
    strSQL = strSQL & "Select TP.PKId, TP.strDescricao TipoComunicacao "
    strSQL = strSQL & "From " & gstrTipoDeComunicacao & " TP "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                intI = intI + 1
                Load mnu_Lista(intI)
                mnu_Lista(intI).Tag = !pkID
                .MoveNext
            Loop
        End With
    End If
    mnu_Lista(0).Visible = False
    
Err_Handle:
End Sub

Sub NovoFiscal()
    txtPKID = glngPegaProximaChave(gstrFiscais, "PKId")
    txt_CNPJCPF = ""
    lbl_strLogradouroC = ""
    lbl_strBairroC = ""
    lbl_strMunicipioC = ""
    lbl_intCepC = ""
    lbl_intNumeroC = ""
    lbl_strComplementoC = ""
    lbl_intCepC = ""
    lbl_strDistritoC = ""
    tab_3dPasta.Tab = 0
    
    chkblnFiscalProdutividade.Value = 0
    
    fra_TipoFiscal.Enabled = True
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar, gstrSalvar, gstrDeletar

End Sub

Private Function blnDadosOk() As Boolean
    blnDadosOk = True
End Function

Function strQueryRelatorio() As String
Dim strSQL As String
strSQL = ""
    strSQL = strSQL & ""
    strSQL = strSQL & "SELECT "
        strSQL = strSQL & "F.Pkid Codigo, "
        strSQL = strSQL & "C.Strnome Nome, "
        strSQL = strSQL & gstrCASEWHEN("F.bytTipoFiscal", "0, 'Fiscal de Tributos'", "'Fiscal de Posturas'") & " AS Tipo"
    strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrFiscais & " F, "
        strSQL = strSQL & gstrContribuinte & " C "
    strSQL = strSQL & " WHERE "
        strSQL = strSQL & "F.intContribuinte = C.Pkid"
strQueryRelatorio = strSQL
End Function
