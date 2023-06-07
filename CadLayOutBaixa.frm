VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadLayOutBaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Layout de Baixa"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "CadLayOutBaixa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8070
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5325
      Left            =   150
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   180
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   9393
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Layout de Baixa"
      TabPicture(0)   =   "CadLayOutBaixa.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_LayOut"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Arquivo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtPKId"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Colunas"
      TabPicture(1)   =   "CadLayOutBaixa.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_Coluna"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "tdb_Colunas"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   1035
         Left            =   -74820
         TabIndex        =   23
         Top             =   420
         Width           =   7425
         Begin VB.TextBox txt_Codigo 
            Height          =   285
            Left            =   1665
            MaxLength       =   15
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txt_Descricao 
            Height          =   285
            Left            =   1665
            MaxLength       =   50
            TabIndex        =   24
            Top             =   570
            Width           =   5535
         End
         Begin VB.Label lbl_Codigo 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   1095
            TabIndex        =   27
            Top             =   285
            Width           =   495
         End
         Begin VB.Label lbl_Descricao 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   870
            TabIndex        =   26
            Top             =   615
            Width           =   720
         End
      End
      Begin VB.Frame fra_Coluna 
         Caption         =   "Dados das Colunas"
         Height          =   1785
         Left            =   -74820
         TabIndex        =   18
         Top             =   1500
         Width           =   7425
         Begin VB.ComboBox cbo_intDescricao 
            Height          =   315
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txt_intPosicaoColuna 
            Height          =   285
            Left            =   1665
            MaxLength       =   4
            TabIndex        =   5
            Top             =   630
            Width           =   555
         End
         Begin VB.TextBox txt_intTamanhoCampo 
            Height          =   285
            Left            =   3750
            MaxLength       =   15
            TabIndex        =   6
            Top             =   630
            Width           =   555
         End
         Begin VB.ComboBox cmb_bytTipoDado 
            Height          =   315
            ItemData        =   "CadLayOutBaixa.frx":107A
            Left            =   1665
            List            =   "CadLayOutBaixa.frx":107C
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   990
            Width           =   2655
         End
         Begin VB.CheckBox chk_blnContemVirgula 
            Caption         =   "Contém Vírgula"
            Height          =   195
            Left            =   4380
            TabIndex        =   8
            Top             =   1110
            Width           =   1395
         End
         Begin VB.TextBox txt_bytPosicaoDaVirgula 
            Height          =   285
            Left            =   1665
            MaxLength       =   2
            TabIndex        =   9
            Top             =   1380
            Width           =   555
         End
         Begin VB.Label lbl_DescricaoColuna 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   870
            TabIndex        =   31
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl_intPosicaoColuna 
            AutoSize        =   -1  'True
            Caption         =   "Posição"
            Height          =   195
            Left            =   1020
            TabIndex        =   22
            Top             =   675
            Width           =   570
         End
         Begin VB.Label lbl_intTamanhoCampo 
            AutoSize        =   -1  'True
            Caption         =   "Tamanho"
            Height          =   195
            Left            =   3015
            TabIndex        =   21
            Top             =   675
            Width           =   675
         End
         Begin VB.Label lbl_bytTipoDado 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Dados"
            Height          =   195
            Left            =   540
            TabIndex        =   20
            Top             =   1050
            Width           =   1050
         End
         Begin VB.Label lbl_bytPosicaoDaVirgula 
            AutoSize        =   -1  'True
            Caption         =   "Casas Decimais"
            Height          =   195
            Left            =   465
            TabIndex        =   19
            Top             =   1425
            Width           =   1125
         End
      End
      Begin VB.TextBox txtPKId 
         Height          =   285
         Left            =   2340
         MaxLength       =   15
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame fra_Arquivo 
         Caption         =   "Dados do Arquivo"
         Height          =   1635
         Left            =   180
         TabIndex        =   11
         Top             =   420
         Width           =   7455
         Begin VB.CheckBox chk_blnPularHeader 
            Caption         =   "Ignora Header"
            Height          =   195
            Left            =   2490
            TabIndex        =   29
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CheckBox chk_blnPularTrailer 
            Caption         =   "Ignora Trailer"
            Height          =   195
            Left            =   4320
            TabIndex        =   28
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtstrDescricao 
            Height          =   285
            Left            =   1665
            MaxLength       =   50
            TabIndex        =   1
            Top             =   570
            Width           =   5535
         End
         Begin VB.TextBox txtstrCodigo 
            Height          =   285
            Left            =   1665
            MaxLength       =   15
            TabIndex        =   0
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtintTamanhodaLinha 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1665
            MaxLength       =   4
            TabIndex        =   2
            Top             =   900
            Width           =   555
         End
         Begin VB.TextBox txtstrSeparadorColuna 
            Height          =   285
            Left            =   1665
            MaxLength       =   1
            TabIndex        =   3
            Top             =   1230
            Width           =   555
         End
         Begin VB.Label lblstrDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   870
            TabIndex        =   16
            Top             =   615
            Width           =   720
         End
         Begin VB.Label lblstrCodigo 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   1095
            TabIndex        =   15
            Top             =   285
            Width           =   495
         End
         Begin VB.Label lblintTamanhodaLinha 
            AutoSize        =   -1  'True
            Caption         =   "Caracteres por Linha"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   945
            Width           =   1470
         End
         Begin VB.Label lblstrSeparadorColuna 
            AutoSize        =   -1  'True
            Caption         =   "Dígito Separador"
            Height          =   195
            Left            =   375
            TabIndex        =   13
            Top             =   1275
            Width           =   1215
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_LayOut 
         Height          =   2895
         Left            =   180
         TabIndex        =   4
         Top             =   2220
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5106
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
         Columns(1).DataField=   "strCodigo"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2990"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2910"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=9657"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=9578"
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
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid tdb_Colunas 
         Height          =   1725
         Left            =   -74850
         TabIndex        =   10
         Top             =   3390
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   3043
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
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "intDescricaoColuna"
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Posição"
         Columns(2).DataField=   "intPosicaoCampo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tamanho"
         Columns(3).DataField=   "intTamanhoCampo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Tipo de Dados"
         Columns(4).DataField=   "TipoDados"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=3889"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3810"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1879"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1799"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1905"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1826"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=8890"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=8811"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
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
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "frmCadLayOutBaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim mblnPrimeiraVez         As Boolean
Dim mblnAlterando           As Boolean
Dim mblnAlterandoColunas    As Boolean
Dim vetDescricao(8)         As String

Private Sub chk_blnContemVirgula_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_blnContemVirgula
End Sub

Private Sub chk_blnPularHeader_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_blnPularHeader
End Sub

Private Sub chk_blnPularTrailer_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_blnPularTrailer
End Sub

Private Sub cmb_bytTipoDado_Click()
    If cmb_bytTipoDado.ListIndex >= 0 Then
        If cmb_bytTipoDado.ItemData(cmb_bytTipoDado.ListIndex) = 0 Then
            chk_blnContemVirgula.Enabled = True
            TrocaCorObjeto txt_bytPosicaoDaVirgula, False
        Else
            chk_blnContemVirgula.Enabled = False
            chk_blnContemVirgula.Value = 0
            TrocaCorObjeto txt_bytPosicaoDaVirgula, True
        End If
    End If
End Sub

Private Sub cmb_bytTipoDado_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cmb_bytTipoDado
End Sub

Private Sub Form_Load()
    CarregaDescricaoColuna
    
    cmb_bytTipoDado.AddItem "Numérico"
    cmb_bytTipoDado.ItemData(cmb_bytTipoDado.NewIndex) = 0
    cmb_bytTipoDado.AddItem "Alfa Numérico"
    cmb_bytTipoDado.ItemData(cmb_bytTipoDado.NewIndex) = 1
    cmb_bytTipoDado.AddItem "Data"
    cmb_bytTipoDado.ItemData(cmb_bytTipoDado.NewIndex) = 2
    
    chk_blnContemVirgula.Enabled = False
    TrocaCorObjeto txt_bytPosicaoDaVirgula, True
    
    LeDaTabelaParaObj "", tdb_LayOut, strQuery

    txt_Codigo.Enabled = False
    TrocaCorObjeto txt_Codigo, True
    txt_Descricao.Enabled = False
    TrocaCorObjeto txt_Descricao, True
    
    mblnAlterandoColunas = False
    tab_3dPasta.TabEnabled(1) = False
    
End Sub

Public Sub MantemForm(strModoOperacao As String)
    
    Select Case strModoOperacao
        Case gstrNovo
            If tab_3dPasta.Tab = 0 Then
                LimpaCampos
                mblnAlterando = False
                mblnAlterandoColunas = False
                tab_3dPasta.TabEnabled(1) = False
            ElseIf tab_3dPasta.Tab = 1 Then
                LimpaCamposColuna
                mblnAlterandoColunas = False
            End If
            
        Case gstrSalvar
            If blnCamposOK Then
                If blnCamposDoisOK Then
                    If tab_3dPasta.Tab = 0 Then
                        If mblnAlterando = False Then
                            InluirLayOut
                            mblnPrimeiraVez = False
                            txtPKId = ""
                            LeDaTabelaParaObj "", tdb_LayOut, strQuery
                            LimpaCampos
                        Else
                            AlteraLayout
                            txt_intPosicaoColuna.SetFocus
                            mblnPrimeiraVez = False
                            LeDaTabelaParaObj "", tdb_LayOut, strQuery
                            LimpaCampos
                        End If
                    Else
                        If mblnAlterandoColunas = True Then
                            AlteraColuna
                            LeDaTabelaParaObj "", tdb_Colunas, strQuerryColuna & " ORDER BY intPosicaoCampo"
                            LimpaColuna
                        Else
                            IncluiColuna
                            LeDaTabelaParaObj "", tdb_Colunas, strQuerryColuna
                            DoEvents
                            LimpaColuna
                            mblnAlterandoColunas = False
                        End If
                    End If
                End If
            End If
            
        Case gstrDeletar
            If tab_3dPasta.Tab = 0 Then
                If ExcluiLayOut Then
                    LimpaCampos
                    mblnAlterando = False
                    txtPKId = ""
                    LeDaTabelaParaObj "", tdb_LayOut, strQuery
                End If
            ElseIf tab_3dPasta.Tab = 1 Then
                If ExcluiColuna Then
                    LimpaColuna
                    mblnAlterandoColunas = False
                    LeDaTabelaParaObj "", tdb_Colunas, strQuerryColuna
                End If
            End If
            
        Case gstrFechar
            Unload Me
            
        Case gstrImprimir
            Exit Sub
    End Select
    
End Sub

Private Sub LimpaCamposColuna()
    
    txt_intPosicaoColuna = ""
    txt_intTamanhoCampo = ""
    cmb_bytTipoDado.ListIndex = -1
    txt_bytPosicaoDaVirgula = ""
    chk_blnContemVirgula.Value = 0
    chk_blnContemVirgula.Enabled = False
    chk_blnContemVirgula.Value = 0
    TrocaCorObjeto txt_bytPosicaoDaVirgula, True
    txt_intPosicaoColuna.SetFocus
    cbo_intDescricao.ListIndex = -1
    
End Sub

Private Sub LimpaCampos()
    
    txtstrCodigo = ""
    txtstrDescricao = ""
    txtintTamanhodaLinha = ""
    txtstrSeparadorColuna = ""
    txt_intPosicaoColuna = ""
    txt_intTamanhoCampo = ""
    cmb_bytTipoDado.ListIndex = -1
    txt_bytPosicaoDaVirgula = ""
    chk_blnContemVirgula.Value = 0
    chk_blnPularHeader.Value = 0
    chk_blnPularTrailer.Value = 0
    chk_blnContemVirgula.Enabled = False
    chk_blnContemVirgula.Value = 0
    TrocaCorObjeto txt_bytPosicaoDaVirgula, True
    txtstrCodigo.SetFocus
    cbo_intDescricao.ListIndex = -1
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mblnPrimeiraVez = False
    mblnAlterando = False
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    Select Case tab_3dPasta.Tab
        Case 0
        
        Case 1
            txt_Codigo.Text = txtstrCodigo.Text
            txt_Descricao.Text = txtstrDescricao.Text
    End Select
End Sub

Private Sub tdb_Colunas_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Select Case ColIndex
        Case 1
            Value = vetDescricao(Value)
    End Select
End Sub

Private Sub tdb_LayOut_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_LayOut_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_LayOut
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKId = .Columns(0).Value
                mblnAlterando = True
                MontaCampos
                tab_3dPasta.TabEnabled(1) = True
                gCorLinhaSelecionada tdb_LayOut
            End If
        Else
            tab_3dPasta.TabEnabled(1) = False
        End If
    End With
End Sub

Private Sub tdb_Colunas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Colunas
        If Not .EOF And Not .BOF Then
            MontaCamposColunas
            mblnAlterandoColunas = True
            gCorLinhaSelecionada tdb_Colunas
        End If
    End With
End Sub

Private Sub MontaCamposColunas()
Dim strSql As String
Dim adoRec As ADODB.Recordset

    strSql = strQuerryColuna & " AND PKId = " & tdb_Colunas.Columns(0).Value & " ORDER BY PKId "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        With adoRec
            If Not (.BOF And .EOF) Then
                txt_intPosicaoColuna = gstrENulo(!intPosicaoCampo)
                txt_intTamanhoCampo = gstrENulo(!intTamanhoCampo)
                cmb_bytTipoDado.ListIndex = gstrENulo(!bytTipoDado)
                txt_bytPosicaoDaVirgula = gstrENulo(!bytPosicaoDaVirgula)
                chk_blnContemVirgula.Value = Abs(!blnContemVirgula)
                cbo_intDescricao.ListIndex = gintIndiceCBO(cbo_intDescricao, !intDescricaoColuna)
                End If
        End With
    End If
End Sub

Private Sub MontaCampos()
Dim strSql As String
Dim adoRec As ADODB.Recordset

    strSql = strQuery

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        With adoRec
            If Not (.BOF And .EOF) Then
                txtstrCodigo = gstrENulo(!strCodigo)
                txtstrDescricao = gstrENulo(!strDescricao)
                txtintTamanhodaLinha = gstrENulo(!intTamanhodaLinha)
                txtstrSeparadorColuna = gstrENulo(!strSeparadorColuna)
                chk_blnPularHeader.Value = Abs(!blnPularHeader)
                chk_blnPularTrailer.Value = Abs(!blnPularTrailer)
                LeDaTabelaParaObj "", tdb_Colunas, strQuerryColuna
            End If
        End With
    End If
End Sub
'
Private Function strQuerryColuna() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT " & gstrLayoutColuna & ".*, "
'    strSql = strSql & " Case bytTipoDado WHEN 0 THEN 'Numérico' WHEN 1 THEN 'Alfa Numérico' ELSE 'Data' END AS TipoDados "
    strSql = strSql & gstrCASEWHEN("bytTipoDado", "0, 'Numérico', 1, 'Alfa Numérico'", "'Data'") & " AS TipoDados "
    strSql = strSql & " FROM "
    strSql = strSql & gstrLayoutColuna
    If txtPKId <> "" Then
        strSql = strSql & " WHERE intDescricaoLayOut = " & Val(txtPKId)
    End If
strQuerryColuna = strSql
End Function

Private Function strQuery() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT * FROM "
    strSql = strSql & gstrDescricaoLayout & " A "
    If txtPKId <> "" Then
        strSql = strSql & " WHERE A.PKId = " & Val(txtPKId)
    End If
    strSql = strSql & " ORDER BY strCodigo "
strQuery = strSql
End Function

Private Function blnCamposOK() As Boolean
blnCamposOK = False
    If Trim(txtstrCodigo) = "" Then
        ExibeMensagem "O campo " & lblstrCodigo.Caption & " não pode ser nulo!"
        tab_3dPasta.Tab = 0
        txtstrCodigo.SetFocus
        Exit Function
    End If
    If Trim(txtstrDescricao) = "" Then
        ExibeMensagem "O campo " & lblstrDescricao.Caption & " não pode ser nulo!"
        tab_3dPasta.Tab = 0
        txtstrDescricao.SetFocus
        Exit Function
    End If
    If Trim(txtintTamanhodaLinha) = "" Then
        ExibeMensagem "O campo " & lblintTamanhodaLinha.Caption & " não pode ser nulo!"
        tab_3dPasta.Tab = 0
        txtintTamanhodaLinha.SetFocus
        Exit Function
    End If
    If Trim(txtstrSeparadorColuna) = "" Then
        ExibeMensagem "O campo " & lblstrSeparadorColuna.Caption & " não pode ser nulo!"
        tab_3dPasta.Tab = 0
        txtstrSeparadorColuna.SetFocus
        Exit Function
    End If
blnCamposOK = True
End Function

Private Function blnCamposDoisOK() As Boolean
blnCamposDoisOK = False
    If tab_3dPasta.Tab = 1 Then
        If Trim(txt_intPosicaoColuna) = "" Then
            ExibeMensagem "O campo " & lbl_intPosicaoColuna.Caption & " não pode ser nulo!"
            txt_intPosicaoColuna.SetFocus
            Exit Function
        End If
        If Trim(txt_intTamanhoCampo) = "" Then
            ExibeMensagem "O campo " & lbl_intTamanhoCampo.Caption & " não pode ser nulo!"
            txt_intTamanhoCampo.SetFocus
            Exit Function
        End If
        If cmb_bytTipoDado.ListIndex < 0 Then
            ExibeMensagem "O campo " & lbl_bytTipoDado.Caption & " não pode ser nulo!"
            cmb_bytTipoDado.SetFocus
            Exit Function
        End If
    End If
blnCamposDoisOK = True
End Function

Private Function ExcluiLayOut() As Boolean
Dim strSql As String
    If tab_3dPasta.Tab = 0 Then
    
        If Not gblnExclusaoGravacaoOk("E", "Confirma exclusão do Lay Out e seus relacionamentos?", True) Then Exit Function
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        strSql = ""
    
        strSql = strSql & "DELETE  FROM "
        strSql = strSql & gstrLayoutColuna
        strSql = strSql & " WHERE "
        strSql = strSql & " intDescricaoLayOut = " & Val(txtPKId)
        
        If gobjBanco.Execute(strSql) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaCommitTrans
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        
        strSql = ""
    
        strSql = strSql & "DELETE  FROM "
        strSql = strSql & gstrDescricaoLayout
        strSql = strSql & " WHERE "
        strSql = strSql & " PKId = " & Val(txtPKId)
        
        If gobjBanco.Execute(strSql) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaCommitTrans
            ExcluiLayOut = True
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
        
    End If
End Function

Private Function ExcluiColuna() As Boolean
Dim strSql As String
    If tab_3dPasta.Tab = 1 Then
    
        If Not gblnExclusaoGravacaoOk("E", , False) Then Exit Function
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        
        strSql = ""
        
        strSql = strSql & "DELETE  FROM "
        strSql = strSql & gstrLayoutColuna
        strSql = strSql & " WHERE "
        strSql = strSql & " intDescricaoLayOut = " & Val(txtPKId)
        strSql = strSql & " AND PKId = " & Val(tdb_Colunas.Columns(0).Value)
        
        If gobjBanco.Execute(strSql) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaCommitTrans
            ExcluiColuna = True
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
        
    End If
End Function

Private Function IncluiColuna() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String

    If tab_3dPasta.Tab = 1 Then
        
        If Not blnVerificaColunaExistente(False) Then Exit Function
        
        If Not gblnExclusaoGravacaoOk("I", , False) Then Exit Function
            
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        strSql = ""
    
        strSql = strSql & " INSERT INTO " & gstrLayoutColuna
        strSql = strSql & " (intDescricaoLayout, intPosicaoCampo, intTamanhoCampo, bytTipoDado, blnContemVirgula, "
        strSql = strSql & "bytPosicaoDaVirgula, intDescricaoColuna, dtmDtAtualizacao, lngCodUsr ) "
        strSql = strSql & " VALUES ( " & Val(txtPKId)
        strSql = strSql & " ,'" & gstrConvVrParaSql(txt_intPosicaoColuna) & "'"
        strSql = strSql & ",'" & gstrConvVrParaSql(txt_intTamanhoCampo) & "'"
        strSql = strSql & ",'" & gstrConvVrParaSql(cmb_bytTipoDado.ListIndex) & "'"
        strSql = strSql & ",'" & gstrConvVrParaSql(chk_blnContemVirgula) & "'"
        If Trim(txt_bytPosicaoDaVirgula) <> "" Then
            strSql = strSql & "," & gstrConvVrParaSql(txt_bytPosicaoDaVirgula)
        Else
            strSql = strSql & ", NULL"
        End If
        strSql = strSql & ", " & gstrItemData(cbo_intDescricao)
'        strSql = strSql & ", GETDATE()"
        strSql = strSql & ", " & strGETDATE
        strSql = strSql & ",'" & gstrConvVrParaSql(glngCodUsr) & "' )"
        
        If gobjBanco.Execute(strSql) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaCommitTrans
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
    End If

End Function

Private Function AlteraColuna() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String

    If tab_3dPasta.Tab = 1 Then
        
        If Not blnVerificaColunaExistente(True) Then Exit Function
        
        If Not gblnExclusaoGravacaoOk("A", , False) Then Exit Function
    
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        
        strSql = ""
        strSql = strSql & " UPDATE " & gstrLayoutColuna & " SET"
        strSql = strSql & " intDescricaoLayout = '" & Val(txtPKId) & "',"
        strSql = strSql & " intPosicaoCampo = '" & gstrConvVrParaSql(txt_intPosicaoColuna) & "',"
        strSql = strSql & " intTamanhoCampo = '" & gstrConvVrParaSql(txt_intTamanhoCampo) & "',"
        strSql = strSql & " bytTipoDado = '" & gstrConvVrParaSql(cmb_bytTipoDado.ListIndex) & "',"
        strSql = strSql & " blnContemVirgula = '" & gstrConvVrParaSql(chk_blnContemVirgula) & "',"
        If Trim(txt_bytPosicaoDaVirgula) <> "" Then
            strSql = strSql & " bytPosicaoDaVirgula = '" & gstrConvVrParaSql(txt_bytPosicaoDaVirgula) & "',"
        Else
            strSql = strSql & " bytPosicaoDaVirgula = " & "NULL" & ","
        End If
        strSql = strSql & " intDescricaoColuna = '" & gstrItemData(cbo_intDescricao) & "',"
'        strSql = strSql & " dtmDtAtualizacao = " & "GETDATE()" & ", "
        strSql = strSql & " dtmDtAtualizacao = " & strGETDATE & ", "
        strSql = strSql & " lngCodUsr = '" & gstrConvVrParaSql(glngCodUsr) & "'"
        
        If txtPKId <> "" Then
            strSql = strSql & " WHERE intDescricaoLayout = " & Val(txtPKId)
        End If
        If tdb_Colunas.Columns(0).Value <> "" Then
            strSql = strSql & " AND PKId = " & tdb_Colunas.Columns(0).Value
        End If
        
        If gobjBanco.Execute(strSql) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaCommitTrans
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
    End If
    
End Function

Private Function AlteraLayout() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    
    If tab_3dPasta.Tab = 0 Then
        
        If Not gblnExclusaoGravacaoOk("A", , False) Then Exit Function
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
    
        strSql = ""
        
        strSql = strSql & " UPDATE " & gstrDescricaoLayout & " SET"
        strSql = strSql & " strCodigo = '" & gstrConvVrParaSql(txtstrCodigo) & "',"
        strSql = strSql & " strDescricao = '" & gstrConvVrParaSql(txtstrDescricao) & "',"
        strSql = strSql & " intTamanhodaLinha = '" & gstrConvVrParaSql(txtintTamanhodaLinha) & "',"
        If Trim(txtstrSeparadorColuna) <> "" Then
            strSql = strSql & " strSeparadorColuna = '" & txtstrSeparadorColuna & "',"
        Else
            strSql = strSql & " strSeparadorColuna = " & "NULL" & ","
        End If
        
        
        strSql = strSql & " blnPularHeader = '" & gstrConvVrParaSql(chk_blnPularHeader) & "',"
        strSql = strSql & " blnPularTrailer = '" & gstrConvVrParaSql(chk_blnPularTrailer) & "',"
'        strSql = strSql & " dtmDtAtualizacao = " & "GETDATE()" & ","
        strSql = strSql & " dtmDtAtualizacao = " & strGETDATE & ","
        strSql = strSql & " lngCodUsr = '" & gstrConvVrParaSql(glngCodUsr) & "'"
        
        strSql = strSql & " WHERE PKId = " & Val(txtPKId)
        
        If gobjBanco.Execute(strSql) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaCommitTrans
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
    End If
End Function

Private Function InluirLayOut() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    
    If tab_3dPasta.Tab = 0 Then
        
        If Not gblnExclusaoGravacaoOk("I", , False) Then Exit Function
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
    
        strSql = ""
        
        strSql = strSql & " INSERT INTO " & gstrDescricaoLayout
        strSql = strSql & " (strCodigo, strDescricao, intTamanhodaLinha, strSeparadorColuna,  blnPularHeader, blnPularTrailer, dtmDtAtualizacao, lngCodUsr )"
        strSql = strSql & " VALUES ("
        strSql = strSql & "'" & gstrConvVrParaSql(txtstrCodigo) & "'"
        strSql = strSql & ",'" & gstrConvVrParaSql(txtstrDescricao) & "'"
        strSql = strSql & ",'" & gstrConvVrParaSql(txtintTamanhodaLinha) & "'"
        If Trim(txtstrSeparadorColuna) <> "" Then
            strSql = strSql & ",'" & txtstrSeparadorColuna & "'"
        Else
            strSql = strSql & ", NULL"
        End If
        strSql = strSql & ",'" & gstrConvVrParaSql(chk_blnPularHeader) & "'"
        strSql = strSql & ",'" & gstrConvVrParaSql(chk_blnPularTrailer) & "'"
'        strSql = strSql & "," & " GETDATE()"
        strSql = strSql & "," & strGETDATE
        strSql = strSql & ",'" & gstrConvVrParaSql(glngCodUsr) & "'"
        strSql = strSql & ")"
        
        If gobjBanco.Execute(strSql) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaCommitTrans
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
    End If
End Function

Private Sub LimpaColuna()

    txt_intPosicaoColuna = ""
    txt_intTamanhoCampo = ""
    cmb_bytTipoDado.ListIndex = -1
    txt_bytPosicaoDaVirgula = ""
    chk_blnContemVirgula.Value = 0
    chk_blnContemVirgula.Enabled = False
    chk_blnContemVirgula.Value = 0
    TrocaCorObjeto txt_bytPosicaoDaVirgula, True
    txt_intPosicaoColuna.SetFocus
    cbo_intDescricao.ListIndex = -1
    
End Sub

Private Sub txt_bytPosicaoDaVirgula_GotFocus()
    MarcaCampo txt_bytPosicaoDaVirgula
End Sub

Private Sub txt_bytPosicaoDaVirgula_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_bytPosicaoDaVirgula
End Sub

Private Sub txt_intPosicaoColuna_GotFocus()
    MarcaCampo txt_intPosicaoColuna
End Sub

Private Sub txt_intPosicaoColuna_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intPosicaoColuna
End Sub

Private Sub txt_intTamanhoCampo_GotFocus()
    MarcaCampo txt_intTamanhoCampo
End Sub

Private Sub txt_intTamanhoCampo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intTamanhoCampo
End Sub

Private Sub txtintTamanhodaLinha_GotFocus()
    MarcaCampo txtintTamanhodaLinha
End Sub

Private Sub txtintTamanhodaLinha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintTamanhodaLinha
End Sub

Private Sub txtstrCodigo_GotFocus()
    MarcaCampo txtstrCodigo
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCodigo
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrSeparadorColuna_GotFocus()
    MarcaCampo txtstrSeparadorColuna
End Sub

Private Sub txtstrSeparadorColuna_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSeparadorColuna
End Sub

Private Sub CarregaDescricaoColuna()
    cbo_intDescricao.AddItem "Código da Parcela"
    cbo_intDescricao.ItemData(cbo_intDescricao.NewIndex) = 1
    vetDescricao(1) = "Código da Parcela"
        
    cbo_intDescricao.AddItem "Data Limite para Desconto"
    cbo_intDescricao.ItemData(cbo_intDescricao.NewIndex) = 2
    vetDescricao(2) = "Data Limite para Desconto"
                
    cbo_intDescricao.AddItem "Juros"
    cbo_intDescricao.ItemData(cbo_intDescricao.NewIndex) = 3
    vetDescricao(3) = "Juros"
    
    cbo_intDescricao.AddItem "Multa"
    cbo_intDescricao.ItemData(cbo_intDescricao.NewIndex) = 4
    vetDescricao(4) = "Multa"
    
    cbo_intDescricao.AddItem "Correção"
    cbo_intDescricao.ItemData(cbo_intDescricao.NewIndex) = 5
    vetDescricao(5) = "Correção"
    
    cbo_intDescricao.AddItem "Desconto"
    cbo_intDescricao.ItemData(cbo_intDescricao.NewIndex) = 6
    vetDescricao(6) = "Desconto"
    
    cbo_intDescricao.AddItem "Total Pago"
    cbo_intDescricao.ItemData(cbo_intDescricao.NewIndex) = 7
    vetDescricao(7) = "Total Pago"
        
    cbo_intDescricao.AddItem "Data de Pagamento"
    cbo_intDescricao.ItemData(cbo_intDescricao.NewIndex) = 8
    vetDescricao(8) = "Data de Pagamento"
    
End Sub

Private Function blnVerificaColunaExistente(blnAlterando As Boolean) As Boolean
    Dim adoResultado As ADODB.Recordset
    Dim strSql       As String
    
    On Error GoTo err_blnVerificaColunaExistente
    
    strSql = ""
    strSql = strSql & "SELECT * FROM " & gstrLayoutColuna & " "
    strSql = strSql & "WHERE intDescricaoLayOut = " & Val(txtPKId) & " "
    strSql = strSql & "AND intDescricaoColuna = " & gstrItemData(cbo_intDescricao)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not (.BOF And .EOF) Then
                Select Case blnAlterando
                    Case True
                        If !PKId <> tdb_Colunas.Columns(0).Value Then
                            ExibeMensagem "Coluna já cadastrada."
                            Exit Function
                        End If
                    Case False
                        ExibeMensagem "Coluna já cadastrada."
                        Exit Function
                End Select
            End If
        End With
    End If
    
    blnVerificaColunaExistente = True
err_blnVerificaColunaExistente:
End Function

