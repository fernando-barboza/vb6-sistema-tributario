VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadCodigoOrcamentario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Código Orçamentário"
   ClientHeight    =   6030
   ClientLeft      =   2130
   ClientTop       =   2415
   ClientWidth     =   9525
   HelpContextID   =   11
   Icon            =   "CadCodigoOrcamentario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9525
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5820
      Left            =   120
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   120
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   10266
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Código Orçamentário"
      TabPicture(0)   =   "CadCodigoOrcamentario.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrCodigoOrcamentario"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtstrDescricao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtstrCodigoOrcamentario"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_Lista"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_Aplicavel"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frm_Deducao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtintExercicio"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox txtintExercicio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   35
         Top             =   390
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame frm_Deducao 
         Caption         =   " Deduzir "
         Height          =   2025
         Left            =   4710
         TabIndex        =   34
         Top             =   1110
         Width           =   4485
         Begin VB.CheckBox chkbytDeduzProjecaoAtuarial 
            Caption         =   "Projeção Atuarial"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1740
            Width           =   1635
         End
         Begin VB.CheckBox chkbytDeduzAlienacaoAtivo 
            Caption         =   "Alienação de ativos"
            Height          =   195
            Left            =   1860
            TabIndex        =   27
            Top             =   1490
            Width           =   2355
         End
         Begin VB.CheckBox chkbytDeduzPrivatizacao 
            Caption         =   "Privatização"
            Height          =   195
            Left            =   1860
            TabIndex        =   26
            Top             =   1246
            Width           =   2355
         End
         Begin VB.CheckBox chkbytDeduzPessoalInativo 
            Caption         =   "Pessoal inativo"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1240
            Width           =   1545
         End
         Begin VB.CheckBox chkbytDeduzPensionista 
            Caption         =   "Pensionista"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1490
            Width           =   1545
         End
         Begin VB.CheckBox chkbytDeduzOperacaoDeCreditoExter 
            Caption         =   "Operação de crédito externo"
            Height          =   195
            Left            =   1860
            TabIndex        =   23
            Top             =   484
            Width           =   2355
         End
         Begin VB.CheckBox chkbytDeduzRefinanciamentoDaDivid 
            Caption         =   "Refinanciamento"
            Height          =   195
            Left            =   1860
            TabIndex        =   24
            Top             =   728
            Width           =   1545
         End
         Begin VB.CheckBox chkbytDeduzOperacaoDeCreditoInter 
            Caption         =   "Operação de crédito interno"
            Height          =   195
            Index           =   0
            Left            =   1860
            TabIndex        =   22
            Top             =   240
            Width           =   2355
         End
         Begin VB.CheckBox chkbytDeduzPrevidenciaria 
            Caption         =   "Previdenciária"
            Height          =   225
            Left            =   1860
            TabIndex        =   25
            Top             =   972
            Width           =   1365
         End
         Begin VB.CheckBox chkbytDeduzEducacao 
            Caption         =   "Educação"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1545
         End
         Begin VB.CheckBox chkbytDeduzSaude 
            Caption         =   "Saúde"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   490
            Width           =   1545
         End
         Begin VB.CheckBox chkbytDeduzFundef 
            Caption         =   "Fundef"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   740
            Width           =   1545
         End
         Begin VB.CheckBox chkbytDeduzPessoal 
            Caption         =   "Pessoal ativo"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   990
            Width           =   1545
         End
      End
      Begin VB.Frame fra_Aplicavel 
         Caption         =   " Aplicar "
         Height          =   2025
         Left            =   120
         TabIndex        =   33
         Top             =   1110
         Width           =   4485
         Begin VB.CheckBox chkbytProjecaoAtuarial 
            Caption         =   "Projeção Atuarial"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1740
            Width           =   1545
         End
         Begin VB.CheckBox chkbytAlienacaoAtivo 
            Caption         =   "Alienação de ativos"
            Height          =   195
            Left            =   1890
            TabIndex        =   14
            Top             =   1495
            Width           =   2355
         End
         Begin VB.CheckBox chkbytPrivatizacao 
            Caption         =   "Privatização"
            Height          =   195
            Left            =   1890
            TabIndex        =   13
            Top             =   1250
            Width           =   2355
         End
         Begin VB.CheckBox chkbytPessoalInativo 
            Caption         =   "Pessoal inativo"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1250
            Width           =   1455
         End
         Begin VB.CheckBox chkbytPensionista 
            Caption         =   "Pensionista"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1495
            Width           =   1455
         End
         Begin VB.CheckBox chkbytOperacaoDeCreditoExterno 
            Caption         =   "Operação de crédito externo"
            Height          =   195
            Left            =   1890
            TabIndex        =   10
            Top             =   515
            Width           =   2355
         End
         Begin VB.CheckBox chkbytRefinanciamentoDaDivida 
            Caption         =   "Refinanciamento"
            Height          =   195
            Left            =   1890
            TabIndex        =   11
            Top             =   760
            Width           =   2355
         End
         Begin VB.CheckBox chkbytOperacaoDeCreditoInterno 
            Caption         =   "Operação de crédito Interno"
            Height          =   195
            Left            =   1890
            TabIndex        =   9
            Top             =   270
            Width           =   2355
         End
         Begin VB.CheckBox chkbytPrevidenciaria 
            Caption         =   "Previdenciária"
            Height          =   195
            Left            =   1890
            TabIndex        =   12
            Top             =   1005
            Width           =   2355
         End
         Begin VB.CheckBox chkbytPessoal 
            Caption         =   "Pessoal ativo"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   1005
            Width           =   1455
         End
         Begin VB.CheckBox chkbytFundef 
            Caption         =   "Fundef"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   760
            Width           =   1455
         End
         Begin VB.CheckBox chkbytSaude 
            Caption         =   "Saúde"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   515
            Width           =   1455
         End
         Begin VB.CheckBox chkbytEducacao 
            Caption         =   "Educação"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Width           =   1455
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2475
         Left            =   120
         TabIndex        =   28
         Top             =   3210
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   4366
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "strCodigoOrcamentario"
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2540"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2461"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=12912"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=12832"
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
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=48,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtstrCodigoOrcamentario 
         Height          =   285
         Left            =   930
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "1"
         Top             =   390
         Width           =   1545
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   930
         MaxLength       =   100
         TabIndex        =   1
         Top             =   750
         Width           =   8250
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   135
         TabIndex        =   32
         Top             =   825
         Width           =   720
      End
      Begin VB.Label lblstrCodigoOrcamentario 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   4440
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSComctlLib.ImageList img_Arquivo 
      Left            =   5550
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":105E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":10BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":111A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":11D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":1234
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":1292
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":12F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img_ArquivoD 
      Left            =   180
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":134E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":13AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":140A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":1468
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":14C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":1524
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":1582
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadCodigoOrcamentario.frx":15E0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadCodigoOrcamentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando           As Boolean
Dim mobjAux                 As Object
Dim mblnClickOk             As Boolean
Dim mblndesabilitaLost      As Boolean
Dim strDescricaoAtual       As String
Dim intFiltroExercicio   As Integer
Dim blnPrimeiraVez          As Boolean
Dim blnOrdenacaoAsc         As Boolean
Dim bytOrdenacao            As Byte

Public mIntCodSeguranca     As Integer

Private Sub chkbytAlienacaoAtivo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzAlienacaoAtivo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzEducacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzFundef_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzOperacaoDeCreditoExter_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzOperacaoDeCreditoInter_KeyPress(Index As Integer, KeyAscii As Integer)
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

Private Sub chkbytDeduzPrivatizacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzProjecaoAtuarial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzRefinanciamentoDaDivid_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDeduzSaude_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytEducacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytFundef_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytOperacaoDeCreditoExterno_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytOperacaoDeCreditoInterno_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytPensionista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytPessoal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytPessoalInativo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytPrevidenciaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytPrivatizacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytProjecaoAtuarial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytRefinanciamentoDaDivida_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytSaude_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = mIntCodSeguranca '760
    
    VirificaGradeListView Me
    
    HabilitaDesabilitaBotao1 mblnAlterando, gstrMnuArquivo, gstrDeletar
    
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrSalvar, gstrImprimir
    
End Sub

Private Function strQuery() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strCodigoOrcamentario, strDescricao, intExercicio FROM "
    strSQL = strSQL & gstrCodigoOrcamentario & " "
    strSQL = strSQL & " WHERE intExercicio = " & intFiltroExercicio & " "
    strSQL = strSQL & "ORDER BY strCodigoOrcamentario"
    strQuery = strSQL
End Function

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
    blnOrdenacaoAsc = True
    'Vamos verificar qual menu que chamou o form, para definirmos o filtro
    If gbytMenu = gbytMenuCadastro Then
        intFiltroExercicio = gintExercicio
    Else
        intFiltroExercicio = gintExercicio + 1
    End If

    txtintExercicio = intFiltroExercicio
        
    LeDaTabelaParaObj gstrCodigoOrcamentario, tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    blnPrimeiraVez = False
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub tdb_Lista_Click()
    blnPrimeiraVez = True
    '    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
    '        tdb_Lista_RowColChange 0, 0
    '    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Value = gvntFormatacaoEspecifica(Value, 2)
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    
    blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
    
    bytOrdenacao = ColIndex
    
    LeDaTabelaParaObj "", tdb_Lista, strQueryOrdenaGrid
    
End Sub

Private Function strQueryOrdenaGrid() As String
    Dim strSQL As String
    Dim strSqlComplemento As String
    
    strSQL = "SELECT * FROM " & gstrCodigoOrcamentario
    
    '    If Trim(txtstrCodigoOrcamentario.Text) <> "" Then
    '        strSqlComplemento = strSqlComplemento & " WHERE strCodigoOrcamentario Like '" & Replace(txtstrCodigoOrcamentario, ".", "") & "%'"
    '    End If
    
    If Trim(txtintExercicio) <> "" Then
        If Len(strSqlComplemento) > 0 Then
            strSqlComplemento = strSqlComplemento & " AND intExercicio Like " & txtintExercicio
        Else
            strSqlComplemento = strSqlComplemento & " WHERE intExercicio = " & txtintExercicio
        End If
    End If
    
    '    If Trim(txtstrdescricao) <> "" Then
    '        If Len(strSqlComplemento) > 0 Then
    '            strSqlComplemento = strSqlComplemento & " AND UPPER(strDescricao) Like '" & UCase(txtstrdescricao) & "%'"
    '        Else
    '            strSqlComplemento = strSqlComplemento & " WHERE UPPER(strDescricao) Like '" & UCase(txtstrdescricao) & "%'"
    '        End If
    '    End If
    
    strSQL = strSQL & strSqlComplemento
    
    Select Case bytOrdenacao
    Case Is = 1
        strSQL = strSQL & " ORDER BY strCodigoOrcamentario" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    Case Is = 2
        strSQL = strSQL & " ORDER BY strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQueryOrdenaGrid = strSQL
    
    
End Function

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk And blnPrimeiraVez Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrCodigoOrcamentario, Me
            txtstrCodigoOrcamentario = .Columns("strCodigoOrcamentario").Text
            gCorLinhaSelecionada tdb_Lista
            HabilitaDesabilitaBotao1 Not mobjAux Is Nothing, gstrMnuArquivo, gstrAplicar
            mblnAlterando = True
            strDescricaoAtual = txtstrDescricao.Text
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    'EnviaTeclaTab vbKeyReturn
    
    'ORC1557
    '    If strModoOperacao = gstrSalvar Then
    '        If Not blnDadosOk Then Exit Sub
    '    End If
    '
    '    If strModoOperacao = gstrNovo Then
    '        blnPrimeiraVez = False
    '        mblnAlterando = False
    '    End If
    'ORC1557
    
    Select Case UCase(strModoOperacao)
    Case UCase(gstrSalvar)
        If Not blnDadosOk Then Exit Sub
    Case UCase(gstrNovo)
        blnPrimeiraVez = False
        mblnAlterando = False
    Case UCase(gstrImprimir), UCase(gstrDeletar)
        If rptCodigoOrcamentario.Enabled Then Unload rptCodigoOrcamentario
        If gbytMenu = gbytMenuCadastro Then
            intFiltroExercicio = gintExercicio
        Else
            intFiltroExercicio = gintExercicio + 1
        End If
       
    End Select
        txtintExercicio = intFiltroExercicio
    ToolBarGeral strModoOperacao, gstrCodigoOrcamentario, mblnAlterando, tdb_Lista, Me, _
    mobjAux, strQuery, strQueryAplicar, rptCodigoOrcamentario, strQueryRelatorio
    
    'If strModoOperacao = gstrNovo Or strModoOperacao = gstrDeletar Then txtintExercicio = intFiltroExercicio
    
End Sub

Private Sub txtstrCodigoOrcamentario_Change()
    If txtstrCodigoOrcamentario.Text = "" Then
    mblnAlterando = False
    End If
End Sub

Private Sub txtstrCodigoOrcamentario_KeyPress(KeyAscii As Integer)
    
    mblndesabilitaLost = False
    'CaracterValido KeyAscii, "N", txtstrCodigoOrcamentario
    gstrLimitaCampoValor txtstrCodigoOrcamentario, KeyAscii, 10, 0
    
End Sub

Private Sub txtstrcodigoorcamentario_LostFocus()
    If mblndesabilitaLost Then
        mblndesabilitaLost = False
    Else
        txtstrCodigoOrcamentario = gvntFormatacaoEspecifica(txtstrCodigoOrcamentario)
    End If
End Sub

Function strQueryRelatorio() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT CO.strCodigoOrcamentario as Codigo, CO.strDescricao "
    strSQL = strSQL & "FROM " & gstrCodigoOrcamentario & " CO "
    strSQL = strSQL & " WHERE intExercicio = " & intFiltroExercicio & " "
    strSQL = strSQL & "ORDER BY CO.strCodigoOrcamentario"
    strQueryRelatorio = strSQL
End Function

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Function blnDadosOk() As Boolean
    Dim strWhereComplementar    As String
    'Dim strCodGrid              As String
    
    'Incluido ORC1557 para impedir inclusão de descricoes repetidas no mesmo exercicio
    If mblnAlterando Then
        strWhereComplementar = " AND PKID <> " & Me.txtPKId.Text
    Else
        strWhereComplementar = ""
    End If
    
    'Tratamento para evitar erro de cadastro quando digitado NULL num campo string
    If Trim(txtstrDescricao.Text) = "NULL" Then txtstrDescricao.Text = "NULLL"
    
    If txtstrCodigoOrcamentario.Text = "" Then
        ExibeMensagem "O campo Código Orçamentário deve ser preenchido!"
        txtstrCodigoOrcamentario.SetFocus
        Exit Function
    ElseIf txtstrDescricao.Text = "" Then
        ExibeMensagem "O campo Descrição deve ser preenchido!"
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    'If Not mblnAlterando Then
    '    If strDescricaoAtual <> txtstrDescricao Then
    '        If gblnExisteCodigo(2, gstrCodigoOrcamentario, "strDescricao", "'" & txtstrDescricao.Text & "'", "intExercicio", Str(intFiltroExercicio)) Then
    '            ExibeMensagem "A descrição digitada já se encontra cadastrada!"
    '            txtstrDescricao.SetFocus
    '            Exit Function
    '        End If
    '    End If
    'Else
    'strCodGrid = Mid(tdb_Lista.Columns(1).Value, 1, 1) & "." & Mid(tdb_Lista.Columns(1).Value, 2, 1) & "." & Mid(tdb_Lista.Columns(1).Value, 3, 1) & "." & Mid(tdb_Lista.Columns(1).Value, 4, 1) & "." & Mid(tdb_Lista.Columns(1).Value, 5, 2) & "." & Mid(tdb_Lista.Columns(1).Value, 7, 2)
    
    If Not mblnAlterando Then
        If gblnExisteCodigo(2, gstrCodigoOrcamentario, "strCodigoOrcamentario", "'" & gvntConvFormatoEspecificoParaSQL(txtstrCodigoOrcamentario.Text, 2) & "'", "intExercicio", Str(intFiltroExercicio)) Then
            ExibeMensagem "O código digitado já se encontra cadastrado!"
            txtstrCodigoOrcamentario.SetFocus
            Exit Function
        End If
    End If
    If gblnExisteCodigo(2, gstrCodigoOrcamentario, "strDescricao", "'" & txtstrDescricao.Text & "'", "intExercicio", Str(intFiltroExercicio), , , strWhereComplementar) Then
        ExibeMensagem "A descrição digitada já se encontra cadastrada!"
        txtstrCodigoOrcamentario.SetFocus
        Exit Function
        
    End If
    
    'End If
    blnDadosOk = True
End Function

Private Function strQueryAplicar() As String
    
    strQueryAplicar = " SELECT PKId, strDescricao FROM " & gstrCodigoOrcamentario
    strQueryAplicar = strQueryAplicar & " WHERE intExercicio = " & intFiltroExercicio
    
End Function

