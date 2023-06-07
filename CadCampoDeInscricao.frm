VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadCampoDeInscricao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Parâmetros"
   ClientHeight    =   5685
   ClientLeft      =   2430
   ClientTop       =   2085
   ClientWidth     =   8865
   HelpContextID   =   46
   Icon            =   "CadCampoDeInscricao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8865
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5460
      TabIndex        =   44
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5625
      Left            =   30
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   30
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   9922
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Campos da Inscrição"
      TabPicture(0)   =   "CadCampoDeInscricao.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintTipoDaInscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintTamanho"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblstrSeparador"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintSequencia"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblintCodigo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrDescricao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtintTamanho"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstrSeparador"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtintSequencia"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tdb_CadCampodaInscricao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DbcintTipoDeInscricao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtintCodigo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optbytSetorQuadra(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "optbytSetorQuadra(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Fórmulas Básicas"
      TabPicture(1)   =   "CadCampoDeInscricao.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_strNome"
      Tab(1).Control(1)=   "cmd_BuscaFormula"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dlg_BuscaFormula"
      Tab(1).Control(3)=   "txt_PKIdFormulaBasica"
      Tab(1).Control(4)=   "fra_aux"
      Tab(1).Control(5)=   "fra_Agregacao"
      Tab(1).Control(6)=   "fra_PalavraChave"
      Tab(1).Control(7)=   "txt_strDescricao"
      Tab(1).Control(8)=   "txt_intCodigo"
      Tab(1).Control(9)=   "cbo_bytTipoDeFormula"
      Tab(1).Control(10)=   "txt_strFormula"
      Tab(1).Control(11)=   "Label1"
      Tab(1).Control(12)=   "lbl_TipoDeFormula"
      Tab(1).Control(13)=   "lbl_Descricao"
      Tab(1).Control(14)=   "lbl_Codigo"
      Tab(1).ControlCount=   15
      Begin VB.OptionButton optbytSetorQuadra 
         Caption         =   "Quadra"
         Height          =   195
         Index           =   2
         Left            =   4845
         TabIndex        =   7
         Top             =   2310
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.OptionButton optbytSetorQuadra 
         Caption         =   "Setor"
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   6
         Top             =   2295
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox txt_strNome 
         Height          =   285
         Left            =   -73305
         MaxLength       =   100
         TabIndex        =   11
         Top             =   1230
         Width           =   4470
      End
      Begin VB.CommandButton cmd_BuscaFormula 
         Height          =   330
         Left            =   -66720
         Picture         =   "CadCampoDeInscricao.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de Descrições Gerais"
         Top             =   1530
         Width           =   360
      End
      Begin MSComDlg.CommonDialog dlg_BuscaFormula 
         Left            =   -66810
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txt_PKIdFormulaBasica 
         Height          =   285
         Left            =   -67935
         TabIndex        =   49
         Top             =   540
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame fra_aux 
         Caption         =   "Auxiliares"
         Height          =   1335
         Left            =   -69105
         TabIndex        =   30
         Top             =   4140
         Width           =   2775
         Begin VB.CommandButton cmd_Porcentagem 
            Caption         =   "%"
            Height          =   315
            Left            =   1800
            TabIndex        =   36
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Mod 
            Caption         =   "Mod"
            Height          =   315
            Left            =   960
            TabIndex        =   35
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Divisao 
            Caption         =   "/"
            Height          =   315
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Multiplicacao 
            Caption         =   "*"
            Height          =   315
            Left            =   1800
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_menos 
            Caption         =   "-"
            Height          =   315
            Left            =   960
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Mais 
            Caption         =   "+"
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fra_Agregacao 
         Caption         =   "Agregação"
         Height          =   1335
         Left            =   -71985
         TabIndex        =   24
         Top             =   4140
         Width           =   2775
         Begin VB.CommandButton cmd_Min 
            Caption         =   "MIN"
            Height          =   315
            Left            =   960
            TabIndex        =   29
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Max 
            Caption         =   "MAX"
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_AVG 
            Caption         =   "AVG"
            Height          =   315
            Left            =   1800
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Count 
            Caption         =   "COUNT"
            Height          =   315
            Left            =   960
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Sum 
            Caption         =   "SUM"
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fra_PalavraChave 
         Caption         =   "Palavras Chaves"
         Height          =   1335
         Left            =   -74865
         TabIndex        =   14
         Top             =   4140
         Width           =   2775
         Begin VB.CommandButton cmd_Update 
            Caption         =   "Update"
            Height          =   315
            Left            =   1800
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Create 
            Caption         =   "Create"
            Height          =   315
            Left            =   1800
            TabIndex        =   23
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmd_Group 
            Caption         =   "Group By"
            Height          =   315
            Left            =   960
            TabIndex        =   22
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmd_Order 
            Caption         =   "Order By"
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmd_Where 
            Caption         =   "Where"
            Height          =   315
            Left            =   1800
            TabIndex        =   20
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Insert 
            Caption         =   "Insert"
            Height          =   315
            Left            =   960
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_From 
            Caption         =   "From"
            Height          =   315
            Left            =   960
            TabIndex        =   19
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Delete 
            Caption         =   "Delete"
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_Select 
            Caption         =   "Select"
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox txt_strDescricao 
         Height          =   285
         Left            =   -73305
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1545
         Width           =   6540
      End
      Begin VB.TextBox txt_intCodigo 
         Height          =   285
         Left            =   -73305
         MaxLength       =   15
         TabIndex        =   10
         Top             =   900
         Width           =   1185
      End
      Begin VB.ComboBox cbo_bytTipoDeFormula 
         Height          =   315
         ItemData        =   "CadCampoDeInscricao.frx":1198
         Left            =   -73305
         List            =   "CadCampoDeInscricao.frx":119A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   4470
      End
      Begin VB.TextBox txtintCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1725
         MaxLength       =   15
         TabIndex        =   1
         Top             =   810
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo DbcintTipoDeInscricao 
         Height          =   315
         Left            =   1725
         TabIndex        =   0
         Top             =   405
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_CadCampodaInscricao 
         Height          =   2775
         Left            =   780
         TabIndex        =   8
         Top             =   2640
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   4895
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
         Columns(1).DataField=   "intCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Sequência"
         Columns(3).DataField=   "intSequencia"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "bytSetorQuadra"
         Columns(4).DataField=   "bytSetorQuadra"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=7699"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=7620"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=1667"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1588"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
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
         TabAction       =   1
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
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
      Begin VB.TextBox txtintSequencia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1725
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1170
         Width           =   465
      End
      Begin VB.TextBox txtstrSeparador 
         Height          =   285
         Left            =   1725
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2250
         Width           =   375
      End
      Begin VB.TextBox txtintTamanho 
         Height          =   285
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1890
         Width           =   1605
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1725
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1530
         Width           =   4035
      End
      Begin VB.TextBox txt_strFormula 
         Height          =   2190
         Left            =   -74865
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   1890
         Width           =   8535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   -73815
         TabIndex        =   50
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label lbl_TipoDeFormula 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Fórmula"
         Height          =   195
         Left            =   -74535
         TabIndex        =   48
         Top             =   615
         Width           =   1140
      End
      Begin VB.Label lbl_Descricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   -74115
         TabIndex        =   47
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl_Codigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   -73890
         TabIndex        =   46
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1140
         TabIndex        =   45
         Top             =   855
         Width           =   495
      End
      Begin VB.Label lblintSequencia 
         AutoSize        =   -1  'True
         Caption         =   "Sequência"
         Height          =   195
         Left            =   870
         TabIndex        =   43
         Top             =   1215
         Width           =   765
      End
      Begin VB.Label lblstrSeparador 
         Caption         =   "Separador"
         Height          =   195
         Left            =   900
         TabIndex        =   42
         Top             =   2295
         Width           =   735
      End
      Begin VB.Label lblintTamanho 
         AutoSize        =   -1  'True
         Caption         =   "Tamanho"
         Height          =   195
         Left            =   960
         TabIndex        =   41
         Top             =   1935
         Width           =   675
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   915
         TabIndex        =   40
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label lblintTipoDaInscricao 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   1320
         TabIndex        =   39
         Top             =   465
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmCadCampoDeInscricao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando               As Boolean
Dim mobjAux                     As Object
Dim mblnSelecionou              As Boolean
Dim mblnPrimeiraVez             As Boolean
Dim mblnClickOk                 As Boolean
Dim blnAlterandoFormulaBasica   As Boolean
Dim strCodigoAtual              As String
Dim strDescriAtual              As String
Dim intSeqAtual                 As Integer
Dim bytSetorQuadraAtual         As Byte
    
Private Sub cbo_bytTipoDeFormula_Click()
    BuscaFormulaBasica
End Sub

Private Sub cmd_BuscaFormula_Click()
    With dlg_BuscaFormula
        .DialogTitle = "Abrir arquivo com Fórmula"
        .DefaultExt = "*.*"
        .Filter = "*.*"
        .InitDir = App.Path
        .flags = &H4
        .Filename = ""
        .ShowOpen
        If .Filename <> "" Then
            BuscaFormula (.Filename)
        End If
    End With
End Sub

Private Sub BuscaFormula(strArquivoFormula As String)
    Dim strLinha As String
    Screen.MousePointer = 11
    On Error GoTo ErroNaAbertura
    Open strArquivoFormula For Input As #1
        txt_strFormula.SetFocus
        txt_strFormula.Text = ""
        While Not EOF(1)
            Line Input #1, strLinha
            txt_strFormula.Text = txt_strFormula.Text & strLinha & Chr(13) & Chr(10)
        Wend
    Close #1
    On Error GoTo 0
ErroNaAbertura:
    Screen.MousePointer = vbDefault
End Sub

Private Sub LimpaFormulaBasica(Optional blnCombox As Boolean)
    If blnCombox Then
        cbo_bytTipoDeFormula.ListIndex = -1
    End If
    txt_strNome.Text = ""
    txt_intCodigo.Text = ""
    txt_strDescricao.Text = ""
    txt_strFormula.Text = ""
End Sub

Private Sub BuscaFormulaBasica()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco

    strSQL = ""
    strSQL = strSQL & " SELECT * FROM " & gstrFormulaBasica
    strSQL = strSQL & " WHERE bytTipoDeFormula = " & gstrItemData(cbo_bytTipoDeFormula)
    
    LimpaFormulaBasica
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                blnAlterandoFormulaBasica = True
                txt_PKIdFormulaBasica.Text = gstrENulo(!Pkid)
                txt_intCodigo.Text = gstrENulo(!intCodigo)
                txt_strNome.Text = gstrENulo(!STRNOME)
                txt_strDescricao.Text = gstrENulo(!strDescricao)
                BuscaFormulaCatalogada
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
            End With
        Else
            blnAlterandoFormulaBasica = False
            HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
        End If
    End If
    
End Sub
Private Sub BuscaFormulaCatalogada()

'******************************************************************************************
' Data: 14/03/2003
' Alteração: - Adaptação da leitura da estrutura das stored procedures no Banco.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    txt_strFormula.Text = ""
        
    If (bytDBType = EDatabases.SQLServer) Then
        strSQL = "EXECUTE sp_helpText '" & txt_strNome.Text & "'"
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        strSQL = "SELECT TEXT text FROM ALL_SOURCE WHERE UPPER(NAME) = '" & UCase(txt_strNome.Text) & "' AND "
        strSQL = strSQL & "TYPE = 'PROCEDURE' ORDER BY LINE"
    
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        
        If (Not (adoResultado.EOF)) And (bytDBType = EDatabases.Oracle) Then
            txt_strFormula.Text = txt_strFormula.Text & "CREATE OR REPLACE "
        End If
        
        Do While Not adoResultado.EOF
            
            If (bytDBType = EDatabases.SQLServer) Then
                txt_strFormula.Text = txt_strFormula.Text & gstrENulo(adoResultado("text"))
            
            ElseIf (bytDBType = EDatabases.Oracle) Then
                txt_strFormula.Text = txt_strFormula.Text & Replace(gstrENulo(adoResultado("text")), Chr(10), "") & vbCrLf
            
            End If
            
            adoResultado.MoveNext
        Loop
    End If
End Sub



Private Sub DbcintTipoDeInscricao_Click(Area As Integer)
    If Area = 2 Then
        'If mblnGuardaUltimo = False Then
            'mlngUltimo = Val(DbcintTipoDeInscricao.BoundText)
            'Limpa_Controles Me, True, False, False, False, False
        'End If
        mblnPrimeiraVez = False
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
        LeDaTabelaParaObj gstrCampoDeInscricao, tdb_CadCampodaInscricao, strQueryCampos
        'VerificaListaAutomatica gstrCampoDeInscricao, tdb_CadCampodaInscricao, strQueryCampos
        LimpaCampos
    ElseIf Area = 1 Then
        DropDownDataCombo DbcintTipoDeInscricao, Me, Area
    End If
End Sub

Private Sub DbcintTipoDeInscricao_GotFocus()
    MarcaCampo DbcintTipoDeInscricao
End Sub

Private Sub DbcintTipoDeInscricao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo DbcintTipoDeInscricao, Me, , KeyCode, Shift
End Sub

Private Sub DbcintTipoDeInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", DbcintTipoDeInscricao
End Sub

Private Sub DbcintTipoDeInscricao_LostFocus()
    If DbcintTipoDeInscricao.MatchedWithList Then
        LeDaTabelaParaObj gstrCampoDeInscricao, tdb_CadCampodaInscricao, strQueryCampos
        LimpaCampos
        If UCase(LTrim(RTrim(DbcintTipoDeInscricao.Text))) = "IMOBILIÁRIO URBANO" Then
            optbytSetorQuadra(1).Visible = True
            optbytSetorQuadra(2).Visible = True
        Else
            optbytSetorQuadra(1).Visible = False
            optbytSetorQuadra(2).Visible = False
        End If
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 602
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
    VerificaListaAutomatica gstrCampoDeInscricao, tdb_CadCampodaInscricao, strQuery
    'VerificaListaAutomatica gstrTipoDeInscricao, DbcintTipoDeInscricao
    VerificaObjParaAplicar mobjAux
    'VerificaListaAutomatica gstrTipoDeInscricao, DbcintTipoDeInscricao
    
    DbcintTipoDeInscricao.Tag = strQueryTipoDeInscricao & ";strNomeDaInscricao"
    
    With cbo_bytTipoDeFormula
        .AddItem "VALOR VENAL DO TERRENO"
        .ItemData(.NewIndex) = 1
        .AddItem "VALOR VENAL DA CONSTRUÇÃO"
        .ItemData(.NewIndex) = 2
        .AddItem "VALOR VENAL DO IMÓVEL"
        .ItemData(.NewIndex) = 3
        .AddItem "FRAÇÃO IDEAL"
        .ItemData(.NewIndex) = 4
        .AddItem "TESTADA IDEAL"
        .ItemData(.NewIndex) = 5
        .AddItem "PROFUNDIDADE"
        .ItemData(.NewIndex) = 6
        .AddItem "PONTUAÇÃO"
        .ItemData(.NewIndex) = 7
        .AddItem "ÁREAS"
        .ItemData(.NewIndex) = 8
        .AddItem "PRODUTIVIDADE FISCAL"
        .ItemData(.NewIndex) = 9
        .AddItem "JUROS"
        .ItemData(.NewIndex) = 10
        .AddItem "MULTA"
        .ItemData(.NewIndex) = 11
        .AddItem "CORREÇÃO MONETÁRIA"
        .ItemData(.NewIndex) = 12
        '========= Sandro
        .AddItem "Cálculo da Topografia"
        .ItemData(.NewIndex) = 13
        .AddItem "Cálculo da Situação"
        .ItemData(.NewIndex) = 14
        .AddItem "Cálculo da Pedologia"
        .ItemData(.NewIndex) = 15
        .AddItem "Cálculo do MT² do Terreno"
        .ItemData(.NewIndex) = 16
        .AddItem "Area do Terreno"
        .ItemData(.NewIndex) = 17
        .AddItem "Area Construída"
        .ItemData(.NewIndex) = 18
        .AddItem "Area Total Construída"
        .ItemData(.NewIndex) = 19
        .AddItem "Cálculo da Testada Principal"
        .ItemData(.NewIndex) = 20
        .AddItem "Cálculo do MT² de Construção"
        .ItemData(.NewIndex) = 21
        .AddItem "Cálculo do CAT"
        .ItemData(.NewIndex) = 22
        .AddItem "Cálculo Fatores de Correção"
        .ItemData(.NewIndex) = 23
        .AddItem "Cálculo IPTU"
        .ItemData(.NewIndex) = 24
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnAlterandoFormulaBasica = False
End Sub


Private Sub optbytSetorQuadra_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    optbytSetorQuadra(Index).Value = IIf(optbytSetorQuadra(Index).Value, False, True)
End Sub

Private Sub tdb_CadCampodaInscricao_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_CadCampodaInscricao) = 1 Then
        tdb_CadCampodaInscricao_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_CadCampodaInscricao_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_CadCampodaInscricao
End Sub

Private Function strQuery() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, intCodigo, strDescricao FROM "
    strSQL = strSQL & gstrCampoDeInscricao & " ORDER BY strDescricao"

    strQuery = strSQL
End Function

Private Sub tdb_CadCampodaInscricao_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_CadCampodaInscricao, ColIndex
    mblnPrimeiraVez = False
    mblnClickOk = False
    mblnPrimeiraVez = False
    mblnSelecionou = False
End Sub

Private Sub tdb_CadCampodaInscricao_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_CadCampodaInscricao_KeyPress(KeyAscii As Integer)
    Select Case tdb_CadCampodaInscricao.Col
        Case Is <> 0
            CaracterValido KeyAscii, "A", tdb_CadCampodaInscricao
    End Select
End Sub

Private Sub tdb_CadCampodaInscricao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_CadCampodaInscricao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_CadCampodaInscricao
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtpkID.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrCampoDeInscricao, Me
                gCorLinhaSelecionada tdb_CadCampodaInscricao
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar, gstrSalvar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                strCodigoAtual = tdb_CadCampodaInscricao.Columns("Código").Text
                strDescriAtual = tdb_CadCampodaInscricao.Columns("Descrição").Text
                intSeqAtual = tdb_CadCampodaInscricao.Columns("Sequência").Text
                bytSetorQuadraAtual = Val(tdb_CadCampodaInscricao.Columns("bytSetorQuadra"))
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark As Variant
    'Dim strSql As String
    
    If tab_3dPasta.Tab = 0 Then
        'mblnGuardaUltimo = False
        
    '    strSql = strQuery
        Select Case UCase(strModoOperacao)
            Case Is = UCase(gstrPreencherLista)
                PreencherListaDeOpcoes Me.ActiveControl
                Exit Sub
            Case Is = UCase(gstrSalvar)
                If Not blnDadosOK Then Exit Sub
                mblnPrimeiraVez = False
                ToolBarGeral strModoOperacao, gstrCampoDeInscricao, mblnAlterando, tdb_CadCampodaInscricao, _
                        Me, mobjAux, strQuery, strQuery, rptCampoDeInscricao, strQueryRelatorio
                LeDaTabelaParaObj gstrCampoDeInscricao, tdb_CadCampodaInscricao, strQueryCampos
                LimpaCampos
                Exit Sub
            Case Is = UCase(gstrDeletar)
                mblnPrimeiraVez = False
            Case Is <> UCase(gstrFechar)
                If Not IsEmpty(varBookMark) Then
                    If UCase(strModoOperacao) = "DELETAR" Then
                        tdb_CadCampodaInscricao.MoveFirst
                    Else
                        tdb_CadCampodaInscricao.Bookmark = varBookMark
                    End If
                    If mblnAlterando Then
                        tdb_CadCampodaInscricao_RowColChange 0, 0
                    End If
                End If
                'DbcintTipoDeInscricao.BoundText = mlngUltimo
                'DbcintTipoDeInscricao_Click 2
                'mblnGuardaUltimo = False
        End Select
        
        If UCase(strModoOperacao) = gstrRefresh Then
            LeDaTabelaParaObj "", tdb_CadCampodaInscricao, strQueryCampos
            Exit Sub
        End If
        
        ToolBarGeral strModoOperacao, gstrCampoDeInscricao, mblnAlterando, tdb_CadCampodaInscricao, _
                Me, mobjAux, strQuery, strQuery, rptCampoDeInscricao, strQueryRelatorio
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
        
        LeDaTabelaParaObj gstrCampoDeInscricao, tdb_CadCampodaInscricao, strQueryCampos
        
        
        'o Item Abaixo nao foi deletado pois o mesmo apresenta alteraçoes alem das que estao descritas nas tarefas
        
    ElseIf tab_3dPasta.Tab = 1 Then
        
        Select Case strModoOperacao
            Case gstrNovo
                blnAlterandoFormulaBasica = False
                LimpaFormulaBasica True
                cbo_bytTipoDeFormula.SetFocus
            Case gstrSalvar
                If DadosFormulaBasicaOK Then
                    SalvaFormulaBasica
                End If
            Case gstrDeletar
                If blnAlterandoFormulaBasica Then
                    If DeletaFormulaBasica Then
                        LimpaFormulaBasica
                    End If
                End If
            Case "FECHAR"
                Unload Me
        End Select
        
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        mblnSelecionou = False
        mblnPrimeiraVez = False
        mblnAlterando = False
    End If
    
End Sub

Private Function DadosFormulaBasicaOK() As Boolean
    If cbo_bytTipoDeFormula.Text = "" Then
        ExibeMensagem "Selecione o tipo de Fórmula"
        cbo_bytTipoDeFormula.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intCodigo.Text) = "" Then
        ExibeMensagem "O Código da Fórmula é obrigatório"
        txt_intCodigo.SetFocus
        Exit Function
    End If
    
    If Trim(txt_strNome.Text) = "" Then
        ExibeMensagem "O Nome da Fórmula é obrigatório"
        txt_strNome.SetFocus
        Exit Function
    End If
    
    If Trim(txt_strFormula) = "" Then
        ExibeMensagem "A Fórmula é obrigatória"
        txt_strFormula.SetFocus
        Exit Function
    End If
    
    DadosFormulaBasicaOK = True
    
End Function

Private Sub blnProcedimentoOK(strFormula As String)

'******************************************************************************************
' Data: 11/03/2003
' Alteração: - Adaptação da instrução de exclusão de procedure ao Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strMsg As String
Dim strSQL As String

On Error GoTo err_blnProcedimentoOK


    If (bytDBType = EDatabases.SQLServer) Then
        strSQL = " IF EXISTS (SELECT NAME FROM SYSOBJECTS " & _
                 " WHERE NAME = '" & txt_strNome & "' AND TYPE = 'P')" & _
                 " DROP PROCEDURE " & txt_strNome
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & "DECLARE "
        strSQL = strSQL & "varSQL VARCHAR2(100); "
        strSQL = strSQL & "numCursor NUMBER; "
        strSQL = strSQL & "numReturn NUMBER; "
        strSQL = strSQL & "excNoPriveleges EXCEPTION; "
        strSQL = strSQL & "PRAGMA EXCEPTION_INIT (excNoPriveleges, -20040); "
        strSQL = strSQL & "BEGIN "
        strSQL = strSQL & "SELECT COUNT(*) INTO numReturn FROM ALL_OBJECTS "
        strSQL = strSQL & "WHERE UPPER(OBJECT_NAME) = '" & UCase(txt_strNome) & "' AND OBJECT_TYPE = 'PROCEDURE';"
        strSQL = strSQL & " IF numReturn > 0 THEN "
        strSQL = strSQL & " varSQL := 'DROP PROCEDURE " & txt_strNome & "'; "
        strSQL = strSQL & " numCursor := DBMS_SQL.OPEN_CURSOR; "
        strSQL = strSQL & " DBMS_SQL.PARSE(numCursor, varSQL, DBMS_SQL.V7); "
        strSQL = strSQL & " numReturn := DBMS_SQL.EXECUTE(numCursor); "
        strSQL = strSQL & " DBMS_SQL.CLOSE_CURSOR(numCursor); "
        strSQL = strSQL & " END IF;"
        strSQL = strSQL & " END;"
    
    End If

Set gobjBanco = New clsBanco
gobjBanco.Execute strSQL

gcncADOMain.BeginTrans


'strFormula = Replace(strFormula, Chr(207), "'")
gcncADOMain.Execute txt_strFormula, , adCmdText
gcncADOMain.CommitTrans

MsgBox "Procedimento efetuado com sucesso!"

Exit Sub
err_blnProcedimentoOK:

gcncADOMain.RollbackTrans
strMsg = "Ocorreu um erro na gravação do procedimento no Banco."
ExibeDetalheErro strMsg

End Sub

Private Sub SalvaFormulaBasica()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    
    If blnAlterandoFormulaBasica Then
        If gblnExclusaoGravacaoOk("A", " da Fórmula Básica " & txt_intCodigo.Text) Then
            strSQL = "UPDATE " & gstrFormulaBasica & _
                     " SET bytTipoDeFormula = " & gstrItemData(cbo_bytTipoDeFormula) & _
                     ", strNome = '" & txt_strNome.Text & _
                     "', intCodigo = " & txt_intCodigo.Text & _
                     ", strDescricao = '" & txt_strDescricao.Text & _
                     "', strFormula = '" & Replace(txt_strFormula, "'", Chr(207))
'                     "', dtmDtAtualizacao = GetDate() "
            strSQL = strSQL & "', dtmDtAtualizacao = " & strGETDATE & _
                     ", lngCodUsr = " & glngCodUsr & _
                     " WHERE PKID = " & txt_PKIdFormulaBasica.Text
        End If
    Else
        If gblnExclusaoGravacaoOk("I", " da Fórmula Básica " & txt_intCodigo.Text) Then
            strSQL = "INSERT INTO " & gstrFormulaBasica & _
                    " (bytTipoDeFormula, strNome, intCodigo, strDescricao, strFormula, dtmDtAtualizacao, lngCodUsr) " & _
                    " VALUES (" & gstrItemData(cbo_bytTipoDeFormula) & _
                    ",'" & txt_strNome.Text & _
                    "', " & txt_intCodigo.Text & ", '" & _
                    txt_strDescricao.Text & "', '"
'                    Replace(txt_strFormula, "'", Chr(207)) & "', GetDate() , " & glngCodUsr & ")"
            strSQL = strSQL & Replace(txt_strFormula, "'", Chr(207)) & "', " & strGETDATE & " , " & glngCodUsr & ")"
            blnAlterandoFormulaBasica = True
        End If
    End If
    txt_strFormula = Replace(txt_strFormula, Chr(207), "'")
    If strSQL <> "" Then
        Set gobjBanco = New clsBanco
        gobjBanco.Execute (strSQL)
        If MsgBox("Deseja atualizar o procedimento armazenado?", vbYesNo, "Tributário") = vbYes Then
            blnProcedimentoOK (txt_strFormula)
        End If
    End If
End Sub

Private Function DeletaFormulaBasica() As Boolean
    Dim strSQL As String
  
    If gblnExclusaoGravacaoOk("E", " da Fórmula Básica " & txt_intCodigo.Text) Then
        strSQL = "DELETE " & gstrFormulaBasica & " WHERE PKID = " & txt_PKIdFormulaBasica.Text
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSQL) Then
            blnAlterandoFormulaBasica = False
            DeletaFormulaBasica = True
        End If
    End If
    
End Function

Private Sub txt_intCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intCodigo
End Sub

Private Sub txt_strDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strDescricao
End Sub

Private Sub txt_strNome_GotFocus()
    MarcaCampo txt_strNome
End Sub

Private Sub txt_strNome_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strNome
End Sub

Private Sub txtintCodigo_GotFocus()
    If DbcintTipoDeInscricao.MatchedWithList Then
        gstrProximoCodigo txtintCodigo, gstrCampoDeInscricao, "intCodigo", gintCodSeguranca, "intTipoDeInscricao", DbcintTipoDeInscricao.BoundText
    Else
        ExibeMensagem "É necessário selecionar um Tipo primeiro."
        DbcintTipoDeInscricao.SetFocus
    End If
    MarcaCampo txtintCodigo
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrSeparador_GotFocus()
    MarcaCampo txtstrSeparador
End Sub

Private Sub txtstrSeparador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSeparador
End Sub

Private Sub txtintSequencia_GotFocus()
    MarcaCampo txtintSequencia
End Sub

Private Sub txtintSequencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintSequencia
End Sub

Private Sub txtintTamanho_GotFocus()
    MarcaCampo txtintTamanho
End Sub

Private Sub txtintTamanho_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintTamanho
End Sub

Private Function strQueryCampos() As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT  PKId, intCodigo, strDescricao, intSequencia, bytSetorQuadra "
    strSQL = strSQL & " FROM " & gstrCampoDeInscricao & " "
    strSQL = strSQL & " WHERE intTipoDeInscricao = '" & DbcintTipoDeInscricao.BoundText & "'"
    strSQL = strSQL & " ORDER BY intSequencia,intCodigo "

    strQueryCampos = strSQL
End Function

Function strQueryRelatorio() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT CI.*, TI.strNomeDaInscricao "
    strSQL = strSQL & "FROM " & gstrCampoDeInscricao & " CI, "
    strSQL = strSQL & gstrTipoDeInscricao & " TI "
    If mblnAlterando = True Then
        strSQL = strSQL & " WHERE CI.PKId = " & Val(txtpkID) & " and CI.intTipoDeInscricao = TI.PKId "
        Else
        strSQL = strSQL & " WHERE CI.intTipoDeInscricao = TI.PKId "
    End If
    strSQL = strSQL & " ORDER BY CI.intSequencia "
    strQueryRelatorio = strSQL
End Function

Private Sub cmd_AVG_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " AVG() "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Count_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " COUNT() "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Create_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " CREATE "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Delete_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " DELETE "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Divisao_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " / "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_From_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " FROM "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Group_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " GROUP BY "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Insert_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " INSERT "
    txt_strFormula = strAux
    
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Mais_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " + "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Max_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " MAX() "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_menos_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " - "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Min_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " MIN() "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Mod_Click()

'******************************************************************************************
' Data: 06/05/2003
' Alteração: - Alteração da string de comando MOD para o Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strAux As String
    strAux = txt_strFormula
    
'    strAux = strAux & " Mod "
    If bytDBType = EDatabases.SQLServer Then
        strAux = strAux & " Mod "
    ElseIf bytDBType = EDatabases.Oracle Then
        strAux = strAux & " Mod( , ) "
    End If
    
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Multiplicacao_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " * "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Order_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " ORDER BY "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Porcentagem_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " % "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Select_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " SELECT "
    txt_strFormula = strAux
    
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Sum_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " SUM() "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Update_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " UPDATE "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Sub cmd_Where_Click()
    Dim strAux As String
    strAux = txt_strFormula
    strAux = strAux & " WHERE "
    txt_strFormula = strAux
    txt_strFormula.SetFocus
    txt_strFormula.SelStart = Len(txt_strFormula.Text) + 1
End Sub

Private Function strQueryTipoDeInscricao() As String
Dim strSQL As String
strSQL = "SELECT Pkid, strNomeDaInscricao"
strSQL = strSQL & " FROM "
strSQL = strSQL & gstrTipoDeInscricao
strSQL = strSQL & " ORDER BY Byttipodeinscricao"
strQueryTipoDeInscricao = strSQL

End Function


Private Function blnDadosOK() As Boolean

blnDadosOK = False

If Not DbcintTipoDeInscricao.MatchedWithList Then
    ExibeMensagem "Selecione um Tipo Válido."
    DbcintTipoDeInscricao.SetFocus
    Exit Function
End If

If txtintCodigo = "" Then
    ExibeMensagem "O Código deve ser preenchido."
    txtintCodigo.SetFocus
    Exit Function
End If

If txtintSequencia.Text = "" Then
    ExibeMensagem "A Sequência deve ser preenchida."
    txtintSequencia.SetFocus
    Exit Function
End If

If txtstrDescricao.Text = "" Then
    ExibeMensagem "A Descrição deve ser preenchida."
    txtstrDescricao.SetFocus
    Exit Function
End If
    
If txtintTamanho.Text = "" Then
    ExibeMensagem "O Tamanho deve ser preenchido."
    txtintTamanho.SetFocus
    Exit Function
End If
    
If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.Text)) Then

ProximoCodigo:
        Dim strCodigo As String
        If gblnExisteCodigo(2, gstrCampoDeInscricao, "intCodigo", "'" & txtintCodigo.Text & "'", "intTipoDeInscricao", "'" & DbcintTipoDeInscricao.BoundText & "'") Then
            strCodigo = (gstrProximoCodigo(txtintCodigo, gstrCampoDeInscricao, "intCodigo", gintCodSeguranca, "intTipoDeInscricao", "'" & DbcintTipoDeInscricao.BoundText, , True))
            If MsgBox("O código informado já se encontra cadastrado. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                txtintCodigo.SetFocus
                Exit Function
            Else
                txtintCodigo.Text = strCodigo
                GoTo ProximoCodigo
            End If
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescriAtual)) Then
            
        If gblnExisteCodigo(2, gstrCampoDeInscricao, "strDescricao", "'" & txtstrDescricao.Text & "'", "intTipoDeInscricao", "'" & DbcintTipoDeInscricao.BoundText & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtintSequencia.Text) <> UCase$(intSeqAtual)) Then
        If gblnExisteCodigo(2, gstrCampoDeInscricao, "intSequencia", "'" & txtintSequencia.Text & "'", "intTipoDeInscricao", "'" & DbcintTipoDeInscricao.BoundText & "'") Then
            ExibeMensagem "A Sequência informada já se encontra cadastrada."
            txtintSequencia.SetFocus
            Exit Function
        End If
    End If
    
    
    If UCase(LTrim(RTrim(DbcintTipoDeInscricao.Text))) = "IMOBILIÁRIO URBANO" Then
        Dim bytValorSetorQuadra As Byte
        Dim intCont             As Integer

        For intCont = optbytSetorQuadra.LBound To optbytSetorQuadra.UBound
            If optbytSetorQuadra(intCont).Value Then
                bytValorSetorQuadra = intCont
            End If
        Next intCont
    
        If Not mblnAlterando Or (mblnAlterando And bytValorSetorQuadra <> bytSetorQuadraAtual And bytValorSetorQuadra <> 0) Then
            If gblnExisteCodigo(2, gstrCampoDeInscricao, "bytSetorQuadra", "'" & bytValorSetorQuadra & "'", "intTipoDeInscricao", "'" & DbcintTipoDeInscricao.BoundText & "'") Then
                ExibeMensagem "Já existe " & IIf(bytValorSetorQuadra = 1, "um Setor.", "uma Quadra.")
                optbytSetorQuadra(bytValorSetorQuadra).SetFocus
                Exit Function
            End If
        End If
   End If


blnDadosOK = True
        
End Function

Private Sub LimpaCampos()

txtpkID.Text = ""
txtintCodigo.Text = ""
txtintSequencia.Text = ""
txtstrDescricao.Text = ""
txtintTamanho.Text = ""
txtstrSeparador.Text = ""

End Sub
