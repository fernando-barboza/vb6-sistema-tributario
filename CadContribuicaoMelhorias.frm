VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadContribuicaoMelhorias 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contribuição de Melhorias"
   ClientHeight    =   6735
   ClientLeft      =   210
   ClientTop       =   2220
   ClientWidth     =   8595
   HelpContextID   =   108
   Icon            =   "CadContribuicaoMelhorias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8595
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5700
      TabIndex        =   43
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4155
      Left            =   60
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   60
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   7329
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Contribuição de Melhorias"
      TabPicture(0)   =   "CadContribuicaoMelhorias.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintTabelaDeEdital"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_strInscricaoCadastral"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dbc_strInscricao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_DadosImovel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "msk_strInscricaoCadastral"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintTabelaDeEdital"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmd_TabelaDeEdital"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtintImobiliario"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_SelecaoImoveis"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chk_Selecao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Melhoria"
      TabPicture(1)   =   "CadContribuicaoMelhorias.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Melhorias"
      Tab(1).Control(1)=   "fra_Edital"
      Tab(1).Control(2)=   "fra_Promissario"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Áreas / Tributos"
      TabPicture(2)   =   "CadContribuicaoMelhorias.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblintArea"
      Tab(2).Control(1)=   "lblintTestada"
      Tab(2).Control(2)=   "lbl_Composicao"
      Tab(2).Control(3)=   "dbc_intComposicaoDaReceita"
      Tab(2).Control(4)=   "tdb_Tributos"
      Tab(2).Control(5)=   "tdd_Tributos"
      Tab(2).Control(6)=   "tdb_Testada"
      Tab(2).Control(7)=   "tdb_Area"
      Tab(2).Control(8)=   "dbcintTestada"
      Tab(2).Control(9)=   "dbcintArea"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Histórico"
      TabPicture(3)   =   "CadContribuicaoMelhorias.frx":1096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "img_Aux"
      Tab(3).Control(1)=   "lvw_Historico"
      Tab(3).Control(2)=   "ssp_TipoComunicacao"
      Tab(3).Control(3)=   "txt_Historico"
      Tab(3).ControlCount=   4
      Begin VB.TextBox txt_Historico 
         Height          =   1440
         Left            =   -74940
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   390
         Width           =   8340
      End
      Begin VB.CheckBox chk_Selecao 
         Caption         =   "Selecionar imóveis"
         Height          =   195
         Left            =   4935
         TabIndex        =   2
         Top             =   810
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Frame fra_SelecaoImoveis 
         Caption         =   " Imóveis "
         Height          =   3255
         Left            =   3510
         TabIndex        =   20
         Top             =   3870
         Visible         =   0   'False
         Width           =   8655
         Begin TrueOleDBGrid70.TDBGrid tdb_Imoveis 
            Height          =   2805
            Left            =   165
            TabIndex        =   40
            Top             =   270
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   4948
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
            Columns(1).Caption=   "Inscrição Cadastral"
            Columns(1).DataField=   "strInscricaoAnterior"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Proprietário"
            Columns(2).DataField=   "strNome"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "CNPJ / CPF"
            Columns(3).DataField=   "strCNPJCPF"
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
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=3387"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3307"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=7329"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=7250"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=3493"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3413"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DataMode        =   4
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTips        =   1
            CellTipsWidth   =   0
            MultiSelect     =   2
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
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
      End
      Begin VB.TextBox txtintImobiliario 
         Height          =   285
         Left            =   6795
         TabIndex        =   3
         Top             =   810
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame fra_Promissario 
         Caption         =   " Promissário "
         Height          =   645
         Left            =   -74820
         TabIndex        =   68
         Top             =   1380
         Width           =   8205
         Begin VB.TextBox txt_CNPJCPFPromissario 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6540
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   27
            Top             =   240
            Width           =   1545
         End
         Begin VB.TextBox txt_Promissario 
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   4995
         End
         Begin VB.Label lbl_CPFPromissario 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   5700
            TabIndex        =   70
            Top             =   300
            Width           =   780
         End
         Begin VB.Label lbl_Promissario 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   300
            Width           =   420
         End
      End
      Begin VB.Frame fra_Edital 
         Caption         =   " Dados do Edital "
         Height          =   945
         Left            =   -74820
         TabIndex        =   44
         Top             =   390
         Width           =   8205
         Begin VB.TextBox txt_CustoDaParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            MaxLength       =   12
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   525
            Width           =   1125
         End
         Begin VB.Frame fra_Tipo 
            Enabled         =   0   'False
            Height          =   705
            Left            =   6000
            TabIndex        =   45
            Top             =   150
            Width           =   1650
            Begin VB.OptionButton opt_Tipo 
               Caption         =   "Obra"
               Height          =   195
               Index           =   1
               Left            =   420
               TabIndex        =   46
               Top             =   420
               Width           =   765
            End
            Begin VB.OptionButton opt_Tipo 
               Caption         =   "Serviço"
               Height          =   195
               Index           =   0
               Left            =   420
               TabIndex        =   25
               Top             =   180
               Width           =   915
            End
            Begin VB.Label lbl_Tipo 
               AutoSize        =   -1  'True
               Caption         =   " Tipo "
               Height          =   195
               Left            =   90
               TabIndex        =   67
               Top             =   0
               Width           =   405
            End
         End
         Begin VB.TextBox txt_DataDeInicio 
            Height          =   285
            Left            =   1470
            Locked          =   -1  'True
            MaxLength       =   12
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   210
            Width           =   1125
         End
         Begin VB.TextBox txt_DataDeTermino 
            Height          =   285
            Left            =   4215
            Locked          =   -1  'True
            MaxLength       =   12
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   1125
         End
         Begin VB.TextBox txt_CustoDeTerceiros 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4215
            Locked          =   -1  'True
            MaxLength       =   12
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   540
            Width           =   1125
         End
         Begin VB.Label lbl_Valor 
            AutoSize        =   -1  'True
            Caption         =   "Custo da Parcela"
            Height          =   195
            Left            =   150
            TabIndex        =   50
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl_DataInicio 
            AutoSize        =   -1  'True
            Caption         =   "Data de Início"
            Height          =   195
            Left            =   345
            TabIndex        =   49
            Top             =   330
            Width           =   1020
         End
         Begin VB.Label lbl_DataTermino 
            AutoSize        =   -1  'True
            Caption         =   "Data de Término"
            Height          =   195
            Left            =   2940
            TabIndex        =   48
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label lbl_CustoDeTerceiros 
            AutoSize        =   -1  'True
            Caption         =   "Custo de Terceiros"
            Height          =   195
            Left            =   2805
            TabIndex        =   47
            Top             =   630
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmd_TabelaDeEdital 
         Height          =   315
         Left            =   8040
         Picture         =   "CadContribuicaoMelhorias.frx":10B2
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "617"
         ToolTipText     =   "Ativa Cadastro de Editais"
         Top             =   375
         Width           =   360
      End
      Begin MSDataListLib.DataCombo dbcintTabelaDeEdital 
         Height          =   315
         Left            =   1755
         TabIndex        =   0
         Top             =   375
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSMask.MaskEdBox msk_strInscricaoCadastral 
         Height          =   240
         Left            =   1785
         TabIndex        =   1
         Top             =   765
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin VB.Frame fra_Melhorias 
         Caption         =   " Melhorias "
         Height          =   1935
         Left            =   -74820
         TabIndex        =   71
         Top             =   2070
         Width           =   7875
         Begin MSComctlLib.ListView lvw_Melhoria 
            Height          =   1290
            Left            =   1950
            TabIndex        =   29
            Top             =   570
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   2275
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Equipamentos"
               Object.Width           =   7937
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbcintSecaoDeLogradouro 
            Height          =   315
            Left            =   1950
            TabIndex        =   28
            Top             =   180
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblintSecaoDeLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Seção de Logradouro"
            Height          =   195
            Left            =   255
            TabIndex        =   73
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label lbl_Melhorias 
            AutoSize        =   -1  'True
            Caption         =   "Melhorias da Seção"
            Height          =   195
            Left            =   390
            TabIndex        =   72
            Top             =   615
            Width           =   1410
         End
      End
      Begin VB.Frame fra_DadosImovel 
         Caption         =   " Dados do Imóvel "
         Height          =   2985
         Left            =   150
         TabIndex        =   51
         Top             =   1080
         Width           =   8310
         Begin VB.TextBox txt_Ocorrrencia 
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   2460
            Width           =   2445
         End
         Begin VB.TextBox txt_Bairro 
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   2133
            Width           =   2190
         End
         Begin VB.TextBox txt_Logradouro 
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1810
            Width           =   6435
         End
         Begin VB.TextBox txt_Contribuinte 
            Height          =   285
            Left            =   2385
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1487
            Width           =   3255
         End
         Begin VB.TextBox txt_CNPJCPF 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6585
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1487
            Width           =   1365
         End
         Begin VB.TextBox txt_Area 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6555
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   2133
            Width           =   1395
         End
         Begin VB.TextBox txt_Desmembramento 
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   5
            Top             =   518
            Width           =   2130
         End
         Begin VB.TextBox txt_Loteamento 
            Height          =   285
            Left            =   5820
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   6
            Top             =   518
            Width           =   2130
         End
         Begin VB.TextBox txt_PKIdContribuinte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   1487
            Width           =   885
         End
         Begin VB.TextBox txt_Cep 
            Height          =   285
            Left            =   4230
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   2133
            Width           =   1005
         End
         Begin VB.TextBox txt_PKIdImobiliario 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   195
            Width           =   2130
         End
         Begin VB.TextBox txt_Escritura 
            Height          =   285
            Left            =   5820
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   10
            Top             =   1164
            Width           =   2130
         End
         Begin VB.TextBox txt_Habitese 
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   9
            Top             =   1164
            Width           =   2130
         End
         Begin VB.TextBox txt_Lote 
            Height          =   285
            Left            =   5820
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   8
            Top             =   841
            Width           =   2130
         End
         Begin VB.TextBox txt_Quadra 
            Height          =   285
            Left            =   1515
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Top             =   841
            Width           =   2130
         End
         Begin VB.Label lbl_CNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   5685
            TabIndex        =   65
            Top             =   1530
            Width           =   780
         End
         Begin VB.Label lbl_Area 
            AutoSize        =   -1  'True
            Caption         =   "Área do Terreno"
            Height          =   195
            Left            =   5325
            TabIndex        =   64
            Top             =   2175
            Width           =   1155
         End
         Begin VB.Label lbl_Ocorrencias 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrências"
            Height          =   195
            Left            =   555
            TabIndex        =   63
            Top             =   2490
            Width           =   855
         End
         Begin VB.Label lbl_Desmembramento 
            AutoSize        =   -1  'True
            Caption         =   "Desmembramento"
            Height          =   195
            Left            =   135
            TabIndex        =   62
            Top             =   570
            Width           =   1275
         End
         Begin VB.Label lbl_Loteamento 
            AutoSize        =   -1  'True
            Caption         =   "Loteamento"
            Height          =   195
            Left            =   4875
            TabIndex        =   61
            Top             =   570
            Width           =   840
         End
         Begin VB.Label lbl_Lote 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            Height          =   195
            Left            =   5400
            TabIndex        =   60
            Top             =   900
            Width           =   315
         End
         Begin VB.Label lbl_Escritura 
            AutoSize        =   -1  'True
            Caption         =   "Matrícula"
            Height          =   195
            Left            =   5040
            TabIndex        =   59
            Top             =   1215
            Width           =   675
         End
         Begin VB.Label lbl_Bairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   1005
            TabIndex        =   58
            Top             =   2175
            Width           =   405
         End
         Begin VB.Label lbl_Logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   600
            TabIndex        =   57
            Top             =   1845
            Width           =   810
         End
         Begin VB.Label lbl_Cep 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   3825
            TabIndex        =   56
            Top             =   2175
            Width           =   285
         End
         Begin VB.Label lbl_Proprietario 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   615
            TabIndex        =   55
            Top             =   1530
            Width           =   795
         End
         Begin VB.Label lbl_CodigoReduzido 
            AutoSize        =   -1  'True
            Caption         =   "Código Reduzido"
            Height          =   195
            Left            =   195
            TabIndex        =   54
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lbl_Habitese 
            AutoSize        =   -1  'True
            Caption         =   "Habite-se"
            Height          =   195
            Left            =   735
            TabIndex        =   53
            Top             =   1215
            Width           =   675
         End
         Begin VB.Label lbl_Quadra 
            AutoSize        =   -1  'True
            Caption         =   "Quadra"
            Height          =   195
            Left            =   885
            TabIndex        =   52
            Top             =   900
            Width           =   525
         End
      End
      Begin MSDataListLib.DataCombo dbcintArea 
         Height          =   315
         Left            =   -74145
         TabIndex        =   30
         Top             =   390
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo dbcintTestada 
         Height          =   315
         Left            =   -74610
         TabIndex        =   32
         Top             =   2280
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Area 
         Height          =   1125
         Left            =   -74595
         TabIndex        =   31
         Top             =   765
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   1984
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Medida da Área"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "Standard"
         Columns(0).EditMaskUpdate=   -1  'True
         Columns(0).EditMaskRight=   -1  'True
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "N° de Pavimentos"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "N° de Edificações"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Última Reforma"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   11059392
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3440"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3360"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3784"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3704"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=260"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3572"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3493"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=260"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2884"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2805"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=260"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         DeadAreaBackColor=   -2147483644
         RowDividerColor =   11059392
         RowSubDividerColor=   11059392
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
         _StyleDefs(18)  =   "Splits(0).Style:id=67,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=69,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=71,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=86,.parent=67"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=68,.alignment=0"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=69"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=71"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=102,.parent=67"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=68,.alignment=0"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=69"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=71"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=106,.parent=67"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=103,.parent=68,.alignment=0"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=104,.parent=69"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=105,.parent=71"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=110,.parent=67"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=107,.parent=68,.alignment=0"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=108,.parent=69"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=109,.parent=71"
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
      Begin TrueOleDBGrid70.TDBGrid tdb_Testada 
         Height          =   1125
         Left            =   -74610
         TabIndex        =   34
         Top             =   2655
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   1984
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Medida da Testada"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "Standard"
         Columns(0).EditMaskUpdate=   -1  'True
         Columns(0).EditMaskRight=   -1  'True
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   11059392
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4815"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4736"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTips        =   1
         CellTipsWidth   =   0
         DeadAreaBackColor=   -2147483644
         RowDividerColor =   11059392
         RowSubDividerColor=   11059392
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
         _StyleDefs(18)  =   "Splits(0).Style:id=67,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=69,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=71,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=86,.parent=67"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=68,.alignment=0"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=69"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=71"
         _StyleDefs(34)  =   "Named:id=33:Normal"
         _StyleDefs(35)  =   ":id=33,.parent=0"
         _StyleDefs(36)  =   "Named:id=34:Heading"
         _StyleDefs(37)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(38)  =   ":id=34,.wraptext=-1"
         _StyleDefs(39)  =   "Named:id=35:Footing"
         _StyleDefs(40)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(41)  =   "Named:id=36:Selected"
         _StyleDefs(42)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(43)  =   "Named:id=37:Caption"
         _StyleDefs(44)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(45)  =   "Named:id=38:HighlightRow"
         _StyleDefs(46)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=39:EvenRow"
         _StyleDefs(48)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(49)  =   "Named:id=40:OddRow"
         _StyleDefs(50)  =   ":id=40,.parent=33"
         _StyleDefs(51)  =   "Named:id=41:RecordSelector"
         _StyleDefs(52)  =   ":id=41,.parent=34"
         _StyleDefs(53)  =   "Named:id=42:FilterBar"
         _StyleDefs(54)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBDropDown tdd_Tributos 
         Height          =   885
         Left            =   -71235
         TabIndex        =   36
         Top             =   2940
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   1561
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PkId"
         Columns(0).DataField=   "PkId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "strDescricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Sigla"
         Columns(2).DataField=   "strSigla"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=4392"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4313"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1191"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1111"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits.Count    =   1
         AllowRowSizing  =   0   'False
         Appearance      =   1
         BorderStyle     =   1
         ColumnHeaders   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowDividerStyle =   2
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   0   'False
         ListField       =   "strDescricao"
         DataField       =   "PkId"
         IntegralHeight  =   0   'False
         FetchRowStyle   =   0   'False
         AlternatingRowStyle=   0   'False
         DataMember      =   ""
         ColumnFooters   =   0   'False
         FootLines       =   1
         DeadAreaBackColor=   12632256
         ValueTranslate  =   -1  'True
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid tdb_Tributos 
         Height          =   1125
         Left            =   -71235
         TabIndex        =   35
         Top             =   2655
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1984
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Taxas"
         Columns(0).DataField=   ""
         Columns(0).DropDown=   "tdd_Tributos"
         Columns(0).DropDown.vt=   8
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=5689"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5609"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(0).AutoDropDown=1"
         Splits(0)._ColumnProps(6)=   "Column(0).DropDownList=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AutoCompletion=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
         DataMode        =   4
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(34)  =   "Named:id=33:Normal"
         _StyleDefs(35)  =   ":id=33,.parent=0"
         _StyleDefs(36)  =   "Named:id=34:Heading"
         _StyleDefs(37)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(38)  =   ":id=34,.wraptext=-1"
         _StyleDefs(39)  =   "Named:id=35:Footing"
         _StyleDefs(40)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(41)  =   "Named:id=36:Selected"
         _StyleDefs(42)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(43)  =   "Named:id=37:Caption"
         _StyleDefs(44)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(45)  =   "Named:id=38:HighlightRow"
         _StyleDefs(46)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=39:EvenRow"
         _StyleDefs(48)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(49)  =   "Named:id=40:OddRow"
         _StyleDefs(50)  =   ":id=40,.parent=33"
         _StyleDefs(51)  =   "Named:id=41:RecordSelector"
         _StyleDefs(52)  =   ":id=41,.parent=34"
         _StyleDefs(53)  =   "Named:id=42:FilterBar"
         _StyleDefs(54)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
         Height          =   315
         Left            =   -71235
         TabIndex        =   33
         Top             =   2280
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin Threed.SSPanel ssp_TipoComunicacao 
         Height          =   390
         Left            =   -74160
         TabIndex        =   77
         Top             =   1965
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1984
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComctlLib.Toolbar tlb_Historico 
            Height          =   330
            Left            =   30
            TabIndex        =   78
            Top             =   30
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "img_Aux"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Novo"
                  Object.ToolTipText     =   "Novo"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Salvar"
                  Object.ToolTipText     =   "Adicionar / Alterar"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Deletar"
                  Object.ToolTipText     =   "Remover"
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ListView lvw_Historico 
         Height          =   1545
         Left            =   -74940
         TabIndex        =   38
         Top             =   2460
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   2725
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ImageList img_Aux 
         Left            =   -72930
         Top             =   1830
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadContribuicaoMelhorias.frx":143C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadContribuicaoMelhorias.frx":159C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadContribuicaoMelhorias.frx":16F8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbc_strInscricao 
         Height          =   315
         Left            =   1755
         TabIndex        =   79
         Top             =   735
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lbl_Composicao 
         AutoSize        =   -1  'True
         Caption         =   "Composição  Rec."
         Height          =   195
         Left            =   -71235
         TabIndex        =   76
         Top             =   2100
         Width           =   1305
      End
      Begin VB.Label lblintTestada 
         AutoSize        =   -1  'True
         Caption         =   "Testada"
         Height          =   195
         Left            =   -74610
         TabIndex        =   75
         Top             =   2100
         Width           =   585
      End
      Begin VB.Label lblintArea 
         AutoSize        =   -1  'True
         Caption         =   "Área"
         Height          =   195
         Left            =   -74595
         TabIndex        =   74
         Top             =   450
         Width           =   330
      End
      Begin VB.Label lbl_strInscricaoCadastral 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   300
         TabIndex        =   66
         Top             =   750
         Width           =   1350
      End
      Begin VB.Label lblintTabelaDeEdital 
         AutoSize        =   -1  'True
         Caption         =   "Edital"
         Height          =   195
         Left            =   1260
         TabIndex        =   41
         Top             =   435
         Width           =   390
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   2415
      Left            =   60
      TabIndex        =   19
      Top             =   4260
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   4260
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
      Columns(1).Caption=   "Inscrição Cadastral"
      Columns(1).DataField=   "strInscricaoAnterior"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Proprietário"
      Columns(2).DataField=   "strNome"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "PKId_Imobiliario"
      Columns(3).DataField=   "PKId_Imobiliario"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "PKId_SecaoLogradouro"
      Columns(4).DataField=   "PKId_SecaoLogradouro"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3625"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3545"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=9128"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=9049"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
End
Attribute VB_Name = "frmCadContribuicaoMelhorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando                As Boolean
Dim mblnAlterandoH               As Boolean
Dim mobjAux                      As Object
Dim mblnClickOk                  As Boolean
Dim oList                        As Object
Dim mblnSelecionou               As Boolean
Dim mblnPrimeiraVez              As Boolean
    
Dim X                            As New XArrayDB 'Grid Seleção de Imóveis
Dim X1                           As New XArrayDB
Dim A                            As New XArrayDB 'Grid Tributos
Dim B                            As New XArrayDB 'DropDown Tributos

Private Sub dbc_intComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicaoDaReceita
End Sub

Private Sub dbc_strInscricao_Change()
    If dbc_strInscricao.MatchedWithList Then
        msk_strInscricaoCadastral = dbc_strInscricao.Text
    End If
End Sub

Private Sub dbcintArea_Click(Area As Integer)
    DropDownDataCombo dbcintArea, Me, Area
    With dbcintArea
        If Area = 2 Then
            If .MatchedWithList And Trim(txtintImobiliario.Text) <> "" Then
                MontaArray 3
            End If
        End If
    End With
End Sub

Private Sub dbcintArea_GotFocus()
    tab_3dPasta.Tab = 2
End Sub

Private Sub dbcintArea_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintArea, Me, , KeyCode, Shift
End Sub

Private Sub dbcintArea_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintArea
End Sub

Private Sub dbcintSecaoDeLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintSecaoDeLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTabelaDeEdital_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTabelaDeEdital, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTestada_Click(Area As Integer)
    DropDownDataCombo dbcintTestada, Me, Area
    With dbcintTestada
        If Area = 2 Then
            If .MatchedWithList And Trim(txtintImobiliario.Text) <> "" Then
                MontaArray 4
            End If
        End If
    End With
End Sub

Private Sub dbcintTestada_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTestada, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTestada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTestada
End Sub

Private Sub chk_Selecao_Click()
    If chk_Selecao.Value = 1 Then
        msk_strInscricaoCadastral.Enabled = False
        msk_strInscricaoCadastral.BackColor = &HC0C0C0
        fra_SelecaoImoveis.Visible = True
    Else
        msk_strInscricaoCadastral.Enabled = True
        msk_strInscricaoCadastral.BackColor = &H80000005
        fra_SelecaoImoveis.Visible = False
    End If
End Sub

Private Sub chk_Selecao_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub chk_Selecao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_Selecao
End Sub

Private Sub cmd_TabelaDeEdital_Click()
    ChamaFormCadastro frmCadTabelaDeEditais, dbcintTabelaDeEdital, "PKId, strNomeDoEdital"
End Sub

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
    With dbc_intComposicaoDaReceita
        If Area = 2 Then
            If .MatchedWithList Then
                MontaArray 1, .BoundText
            End If
        End If
    End With
End Sub

Private Sub dbc_intComposicaoDaReceita_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dbc_intComposicaoDaReceita.ToolTipText = dbc_intComposicaoDaReceita.Text
End Sub

Private Sub dbcintSecaoDeLogradouro_Click(Area As Integer)
    DropDownDataCombo dbcintSecaoDeLogradouro, Me, Area
    If Area = 2 Then
        CarregaMelhorias
    End If
End Sub

Sub CarregaMelhorias()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    lvw_Melhoria.ListItems.Clear
    
    If dbcintSecaoDeLogradouro.BoundText = "" Then
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & "SELECT MP.PKId, MP.strNomeDoMelhoramento AS Melhoramento "
    strSql = strSql & "FROM " & gstrMelhoramentoPublico & " MP, "
    strSql = strSql & gstrMelhoramentoDaSecaoDeLogradouro & " MS "
    strSql = strSql & "WHERE MP.PKId = MS.intMelhoramento "
    strSql = strSql & "AND MS.intSecaoDeLogradouro = " & dbcintSecaoDeLogradouro.BoundText

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                Set oList = lvw_Melhoria.ListItems.Add(, , !Melhoramento)
                oList.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub

Function blnGravaMelhorias(blnAlterando As Boolean, intContribuicao As Integer) As Boolean
    Dim i      As Integer
    Dim strSql As String
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    If blnAlterando Then
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrMelhoriaContribuicaoMelhoria & " "
        strSql = strSql & "WHERE intContribuicao = " & intContribuicao
        gobjBanco.Execute strSql
    End If
    
    For i = 1 To lvw_Melhoria.ListItems.Count
        strSql = ""
        If lvw_Melhoria.ListItems(i).Checked = True Then
            strSql = strSql & "INSERT INTO " & gstrMelhoriaContribuicaoMelhoria & " "
            strSql = strSql & "(intContribuicao, intMelhoria"
            strSql = strSql & ") VALUES ("
            strSql = strSql & intContribuicao & ", "
            strSql = strSql & lvw_Melhoria.ListItems(i).Tag & ")"
            
            If Not gobjBanco.Execute(strSql) Then
                gobjBanco.ExecutaRollbackTrans
                ExibeMensagem "Ocorreu um erro ao gravar as melhorias. Os dados não foram gravados."
                Exit Function
            End If
        End If
    Next
    
    gobjBanco.ExecutaCommitTrans
    Set gobjBanco = Nothing
    blnGravaMelhorias = True
    NovaContribuicao
End Function

Sub DeletaMelhorias(intContribuicao As Integer)
    Dim strSql As String

    strSql = ""
    strSql = strSql & "Delete From " & gstrMelhoriaContribuicaoMelhoria & " "
    strSql = strSql & "Where intContribuicao = " & intContribuicao
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
End Sub

Sub MarcaMelhorias(intTag As Integer)
    Dim i As Integer
    For i = 1 To lvw_Melhoria.ListItems.Count
        If lvw_Melhoria.ListItems(i).Tag = intTag Then
            lvw_Melhoria.ListItems(i).Checked = True
        End If
    Next
End Sub

Function SelecionaMelhorias(intContribuicao As Integer)
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "SELECT * "
    strSql = strSql & "FROM " & gstrMelhoriaContribuicaoMelhoria & " "
    strSql = strSql & "WHERE intContribuicao = " & intContribuicao
             
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                MarcaMelhorias (!intMelhoria)
                .MoveNext
            Loop
        End With
    End If
End Function

Private Sub dbcintSecaoDeLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintSecaoDeLogradouro
End Sub

Private Sub dbcintTabelaDeEdital_Click(Area As Integer)
    DropDownDataCombo dbcintTabelaDeEdital, Me, Area
    If dbcintTabelaDeEdital.MatchedWithList And Area = 2 Then
        VerificaListaAutomatica gstrContribuicaoMelhoria, tdb_Lista, strQuery
        LeDaTabelaParaObj gstrSecaoLogradouro, dbcintSecaoDeLogradouro, strQuerySecao
        NovaContribuicao
        CarregaDadosEdital
    End If
End Sub

Private Sub CarregaDadosEdital()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    txt_DataDeInicio = ""
    txt_DataDeTermino = ""
    txt_CustoDaParcela = ""
    txt_CustoDeTerceiros = ""
    opt_Tipo(0).Value = False
    opt_Tipo(1).Value = False
    
    If dbcintTabelaDeEdital.MatchedWithList = False Then Exit Sub

    strSql = ""
    strSql = strSql & "SELECT dtmDataDeInicio, dtmDataDeTermino, dblCustoDaParcela, "
    strSql = strSql & "dblCustoDeTerceiros, bytTipo "
    strSql = strSql & "FROM " & gstrTabelaDeEdital & " "
    strSql = strSql & "WHERE PKID = " & dbcintTabelaDeEdital.BoundText
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                txt_DataDeInicio = gstrDataFormatada(!dtmDataDeInicio)
                txt_DataDeTermino = gstrDataFormatada(!dtmDataDeTermino)
                txt_CustoDaParcela = gvntConvVrDoSql(!dblCustoDaParcela)
                txt_CustoDeTerceiros = gvntConvVrDoSql(!dblCustoDeTerceiros)
                opt_Tipo(!bytTipo).Value = True
            End If
        End With
        adoResultado.Close
        Set adoResultado = Nothing
    End If
    Set gobjBanco = Nothing
End Sub

Private Sub dbcintTabelaDeEdital_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintTabelaDeEdital_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTabelaDeEdital
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 738
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
    mblnPrimeiraVez = True
    PreencheComboInscricao
    
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
    Dim strSql As String

    If MDIMenu.Tag = "Ouvidoria" Then
        cmd_TabelaDeEdital.Enabled = False
    End If
    MontaColumnHeaders
    VerificaMascaraInscricao
    LeDaTabelaParaObj gstrTipoDeArea, dbcintArea, strQueryTipoArea
    LeDaTabelaParaObj gstrTipoDeTestada, dbcintTestada, strQueryTipoTestada
    LeDaTabelaParaObj "", dbcintTabelaDeEdital, "SELECT PKId, strNomeDoEdital FROM " & gstrTabelaDeEdital & " ORDER BY strNomeDoEdital"
    
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita
    strSql = strSql & " WHERE intUtilizacao = 1 " 'Imobiliarias
    LeDaTabelaParaObj "", dbc_intComposicaoDaReceita, strSql
End Sub

Private Function strQuery() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT CO.PKId, CO.strNome, "
    strSql = strSql & gstrRIGHT("IM.strInscricaoAnterior", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoAnterior, "
    strSql = strSql & "IM.PKId AS PKId_Imobiliario,SL.PKId as PKId_SecaoLogradouro "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTabelaDeEdital & " TE, "
    strSql = strSql & gstrSecaoLogradouro & " SL, "
    strSql = strSql & gstrEditalSecaoLogradouro & " TES, "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrImobiliario & " IM "
    strSql = strSql & "WHERE TE.PKId = TES.intTabelaDeEdital "
    strSql = strSql & "AND SL.PKId = TES.intSecaoDeLogradouro "
    strSql = strSql & "AND CO.PKId = IM.intContribuinte AND "
    strSql = strSql & "SL.PKId = IM.intSecoes "
    strSql = strSql & "AND TE.PKId = " & dbcintTabelaDeEdital.BoundText
    strSql = strSql & "ORDER BY IM.strInscricaoAnterior "
    strQuery = strSql
End Function

Private Function strQueryContribuicoesEdital() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT CM.PKId, "
    strSql = strSql & gstrRIGHT("IM.strInscricaoAnterior", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoAnterior, "
    strSql = strSql & "CO.strNome AS Contribuinte, CO.strCNPJCPF "
    strSql = strSql & "FROM " & gstrContribuicaoMelhoria & " CM, "
    strSql = strSql & gstrImobiliario & " IM, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & "WHERE CM.intImobiliario = IM.PKId "
    strSql = strSql & "AND IM.intContribuinte = CO.PKId "
    strSql = strSql & "AND CM.intTabelaDeEdital = " & dbcintTabelaDeEdital.BoundText & " "
    strSql = strSql & "ORDER BY Contribuinte"
    strQueryContribuicoesEdital = strSql
End Function

Private Function strQueryTipoArea() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNomeDaArea "
    strSql = strSql & "FROM " & gstrTipoDeArea & " "
    strSql = strSql & "WHERE bytPassivadeCM = 1 "
    strSql = strSql & "ORDER BY strNomeDaArea"
    strQueryTipoArea = strSql
End Function

Private Function strQueryTipoTestada() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNomeDaTestada "
    strSql = strSql & "FROM " & gstrTipoDeTestada & " "
    strSql = strSql & "WHERE bytPassivadeCM = 1 "
    strSql = strSql & "ORDER BY strNomeDaTestada"
    strQueryTipoTestada = strSql
End Function

Private Sub lvw_Historico_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txt_Historico = lvw_Historico.SelectedItem.Text
    mblnAlterandoH = True
End Sub

Private Sub lvw_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_Historico
End Sub

Private Sub lvw_Melhoria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_Melhoria
End Sub

Private Sub msk_strInscricaoCadastral_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            CarregaDadosImovel True, msk_strInscricaoCadastral
    End Select
    CaracterValido KeyAscii, "A", msk_strInscricaoCadastral
End Sub

Private Sub opt_Tipo_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", opt_Tipo(Index)
End Sub

Private Sub tdb_Area_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Area
End Sub

Private Sub tdb_Imoveis_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub tdb_Imoveis_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Imoveis
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Lista
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If Not (.EOF And Not .BOF) And mblnClickOk Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                mblnClickOk = False
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                 txtPKID = Val(.Columns(0).Text)
                 
                 
    '            LeDaTabelaParaObj gstrContribuicaoMelhoria, Me
    '            dbcintSecaoDeLogradouro_Click 2
                gCorLinhaSelecionada tdb_Lista
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                
                If Val(txtPKID.Text) <> 0 Then
                    CarregaHistoricos txtPKID
                    CarregaDadosImovel False
                    
                    CarregaMelhorias
                    SelecionaMelhorias txtPKID
                End If
                MontaArray 0
                dbcintArea.BoundText = ""
                LimpaGridArea
                dbcintTestada.BoundText = ""
                LimpaGridTestada
                
                CarregaDadosEdital
                
                msk_strInscricaoCadastral.Enabled = True
                msk_strInscricaoCadastral.BackColor = &H80000005
                'chk_Selecao.Visible = False
                fra_SelecaoImoveis.Visible = False
            End If
        End If
    End With
End Sub

Private Sub tdb_Lista_FilterChange()
    gblnFilraCampos tdb_Lista
End Sub

Private Function blnCamposOK() As Boolean
    
    Dim i As Integer
    
    blnCamposOK = False
    
    If Not dbcintTabelaDeEdital.MatchedWithList Then
        ExibeMensagem "Selecione um Edital"
        tab_3dPasta.Tab = 0
        dbcintTabelaDeEdital.SetFocus
        Exit Function
    End If
    
    For i = 1 To lvw_Melhoria.ListItems.Count
        If lvw_Melhoria.ListItems(i).Checked = True Then
            Exit For
        End If
    Next i
    
    If i > lvw_Melhoria.ListItems.Count Then
        ExibeMensagem "Selecione Melhoria(s) da Seção."
        tab_3dPasta.Tab = 1
        dbcintSecaoDeLogradouro.SetFocus
        Exit Function
    End If
    
    'Área da seção
    If X.Count(1) = 0 Or dbcintArea.Text = "" Then
        ExibeMensagem "Selecione área(s) da Seção."
        tab_3dPasta.Tab = 2
        tdb_Area.SetFocus
        Exit Function
    End If
    
    'Testada da seção
    If X.Count(1) = 0 Then
        ExibeMensagem "Selecione Testada(s) da Seção."
        tab_3dPasta.Tab = 2
        tdb_Testada.SetFocus
        Exit Function
    End If
    
    'Tributos
    If X.Count(1) = 0 Then
        ExibeMensagem "Selecione a Taxa para Composição da Receita."
        tab_3dPasta.Tab = 2
        tdb_Tributos.SetFocus
        Exit Function
    End If
    
    blnCamposOK = True
End Function

Private Function GravaContribuicao() As Boolean

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'        strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim adoRec As ADODB.Recordset
    Dim strSql As String
    
    GravaContribuicao = False
    
    If blnCamposOK = True Then
     
        Set adoRec = tdb_Lista.DataSource
        If adoRec.EOF And adoRec.BOF Then
            ExibeMensagem "Não existem imóveis cadastrados para este edital"
            Exit Function
        End If
        
        With adoRec
            .MoveFirst
            strSql = ""
            
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
            
            Do While Not .EOF
                strSql = strSql & " INSERT INTO " & gstrContribuicaoMelhoria
                strSql = strSql & " (intTabelaDeEdital, intImobiliario, intArea, intTestada, intSecaoDeLogradouro, "
                strSql = strSql & " dtmDtAtualizacao, lngCodUsr) VALUES ("
                
                strSql = strSql & dbcintTabelaDeEdital.BoundText
                strSql = strSql & ", " & adoRec!PKId_Imobiliario
                strSql = strSql & ", " & dbcintArea.BoundText
                If dbcintTestada.BoundText <> "" Then
                    strSql = strSql & ", " & dbcintTestada.BoundText
                Else
                    strSql = strSql & ", " & "Null"
                End If
                strSql = strSql & ", " & adoRec!PKId_SecaoLogradouro
'                strSql = strSql & ", GETDATE() "
                strSql = strSql & ", " & strGETDATE
                strSql = strSql & ", " & glngCodUsr
                strSql = strSql & ")"
            
                strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
            
                .MoveNext
            Loop
        End With
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSql) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaCommitTrans
            
            GravaContribuicao = True
            
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
    NovaContribuicao
    End If
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim intContribuicao As Integer
    Dim varBookMark     As Variant
    Dim strSql          As String
    Dim blnAlterando    As Boolean
    
    Select Case UCase(strModoOperacao)
        Case UCase(gstrNovo)
            LimpaObjeto Me, mblnAlterando
            NovaContribuicao
            
        Case UCase(gstrSalvar)
            If blnDadosOk Then
                If chk_Selecao.Value = 0 Then
                    blnAlterando = mblnAlterando
                    If GravaContribuicao Then
                        If blnAlterando Then
                            intContribuicao = Val(txtPKID)
                        Else: intContribuicao = glngPegaUltimaChave(gstrContribuicaoMelhoria, "PKId")
                        End If
                        If blnGravaTributos(mblnAlterando, intContribuicao) Then
                            If blnGravaHistoricos(mblnAlterando, intContribuicao) Then
                                If blnGravaMelhorias(mblnAlterando, intContribuicao) Then
                                End If
                            End If
                        End If
                        mblnPrimeiraVez = False
                    End If
                Else
                    GravaVariosImoveis
                End If
            End If
        Case UCase(gstrImprimir)
'            If blnDadosOK Then
                ImprimeRelatorio rptContribuicaoMelhoria, strQueryRelatorio
'            End If
        Case UCase(gstrDeletar)
            If blnDeletaContribuicao Then
                NovaContribuicao
            End If
            
        Case gstrPreencherLista
            PreencherListaDeOpcoes Me.ActiveControl
            
        Case UCase(gstrFechar)
            Unload Me
    End Select
End Sub

Private Function blnDeletaContribuicao() As Boolean
    Dim strSql As String

    If MsgBox("Confirma exclusão do registro de '" & "' ?", vbQuestion + vbYesNo) = vbYes Then
        DeletaHistoricos Val(txtPKID)
        
        strSql = ""
        strSql = strSql & "Delete From " & gstrContribuicaoMelhoria & " "
        strSql = strSql & "Where PKId = " & Val(txtPKID)
        
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSql
        
        DeletaTributos Val(txtPKID)
        DeletaHistoricos Val(txtPKID)
        DeletaMelhorias Val(txtPKID)
        
        VerificaListaAutomatica gstrContribuicaoMelhoria, tdb_Lista, strQuery
    End If
    blnDeletaContribuicao = True
End Function

Private Sub NovaContribuicao()
    LimpaGrid
    LimpaDadosImovel
    
    txt_Historico = ""
    lvw_Historico.ListItems.Clear
    lvw_Melhoria.ListItems.Clear
    
    txt_DataDeInicio = ""
    txt_DataDeTermino = ""
    txt_CustoDaParcela = ""
    txt_CustoDeTerceiros = ""
    opt_Tipo(0).Value = False
    opt_Tipo(1).Value = False
    
    mblnAlterando = False
    mblnAlterandoH = False
    
    msk_strInscricaoCadastral.Enabled = True
    msk_strInscricaoCadastral.BackColor = &H80000005
    fra_SelecaoImoveis.Visible = False
    'chk_Selecao.Visible = True
    chk_Selecao.Value = 0
    
    X.Clear
    dbcintArea.Text = ""
    dbcintTestada.Text = ""
    dbcintSecaoDeLogradouro.BoundText = ""
    dbcintTabelaDeEdital.SetFocus
    tab_3dPasta.Tab = 0
End Sub

Private Function blnDadosOk() As Boolean
    If chk_Selecao.Value = 0 Then
        If msk_strInscricaoCadastral.ClipText = "" Then
            ExibeMensagem "A inscrição cadastral do imóvel tem que ser informada."
            msk_strInscricaoCadastral.SetFocus
            Exit Function
        End If
        If Not gblnExisteValorNaTabela(gstrImobiliario, "strInscricaoAnterior", String(gintLenInscricao - Len(Trim(msk_strInscricaoCadastral)), "0") & Trim(msk_strInscricaoCadastral)) Then
            ExibeMensagem "Imóvel não cadastrado."
            msk_strInscricaoCadastral.SetFocus
            Exit Function
        End If
    End If
    If gblnDataValida(txt_DataDeInicio) = False Then
        ExibeMensagem "A data de início não é válida."
        txt_DataDeInicio.SetFocus
        Exit Function
    End If
    If gblnDataValida(txt_DataDeTermino) = False Then
        ExibeMensagem "A data de término não é válida."
        txt_DataDeTermino.SetFocus
        Exit Function
    End If
    blnDadosOk = True
End Function

Private Sub tdb_Testada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Testada
End Sub

Private Sub tdb_Tributos_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Tributos
End Sub

Private Sub tlb_Historico_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case gstrSalvar
            If Trim(txt_Historico) = "" Then Exit Sub
            If mblnAlterandoH Then
                lvw_Historico.SelectedItem.Text = txt_Historico
            Else
                lvw_Historico.ListItems.Add , , txt_Historico
            End If
        Case gstrNovo
            txt_Historico.SetFocus
        Case gstrDeletar
            If lvw_Historico.ListItems.Count = 0 Then Exit Sub
            If lvw_Historico.SelectedItem.Selected Then
                lvw_Historico.ListItems.Remove (lvw_Historico.SelectedItem.Index)
            End If
    End Select
    If lvw_Historico.ListItems.Count <> 0 Then
        lvw_Historico.SelectedItem.Selected = False
    End If
    mblnAlterandoH = False
    txt_Historico = ""
    txt_Historico.SetFocus
End Sub

Private Sub txt_CNPJCPFPromissario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_CNPJCPFPromissario
End Sub

Private Sub txt_CustoDaParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_CustoDaParcela
End Sub

Private Sub txt_CustoDeTerceiros_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_CustoDeTerceiros
End Sub

Private Sub txt_DataDeInicio_GotFocus()
    tab_3dPasta.Tab = 1
End Sub

Private Sub txt_DataDeInicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataDeInicio
End Sub

Private Sub txt_DataDeTermino_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataDeTermino
End Sub

Private Sub txt_Historico_GotFocus()
    MarcaCampo txt_Historico
    tab_3dPasta.Tab = 3
End Sub

Private Sub txt_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Historico
End Sub

Sub MontaColumnHeaders()
    With lvw_Historico
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Histórico", 7000
    End With
End Sub

Private Function blnGravaHistoricos(blnAlterando As Boolean, intCodContribuicao As Integer) As Boolean
    Dim strSql As String
    Dim intI   As Integer
    
    On Error GoTo err_GravaHistoricos
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    If blnAlterando Then
        strSql = ""
        strSql = strSql & "Delete From " & gstrHistoricoContribuicaoMelhoria & " "
        strSql = strSql & "Where intContribuicao = " & intCodContribuicao
        
        gobjBanco.Execute strSql
    End If
    
    With lvw_Historico
        For intI = 1 To .ListItems.Count
            strSql = ""
            strSql = strSql & "Insert Into " & gstrHistoricoContribuicaoMelhoria & " "
            strSql = strSql & "(intContribuicao, strDescricao "
            strSql = strSql & ") Values ("
            strSql = strSql & intCodContribuicao & ",' "
            strSql = strSql & .ListItems(intI).Text & "' "
            strSql = strSql & ")"
            If Not gobjBanco.Execute(strSql) Then
                gobjBanco.ExecutaRollbackTrans
                ExibeMensagem "Ocorreu um erro ao gravar os históricos. Os dados não foram gravados."
                Exit Function
            End If
        Next
    End With

    gobjBanco.ExecutaCommitTrans
    Set gobjBanco = Nothing
    
    blnGravaHistoricos = True
    NovaContribuicao
Exit Function
err_GravaHistoricos:
    gobjBanco.ExecutaRollbackTrans
    Set gobjBanco = Nothing
    ExibeDetalheErro ""
    
End Function

Private Sub DeletaHistoricos(intCodContribuicao As Integer)
    Dim strSql As String
    strSql = ""
    strSql = strSql & "DELETE FROM " & gstrHistoricoContribuicaoMelhoria & " "
    strSql = strSql & "WHERE intContribuicao = " & intCodContribuicao
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
End Sub

Private Sub CarregaHistoricos(intCodContribuicao As Integer)
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    lvw_Historico.ListItems.Clear
    txt_Historico = ""
    
    strSql = ""
    strSql = strSql & "SELECT HI.strDescricao AS Historico "
    strSql = strSql & "FROM " & gstrHistoricoContribuicaoMelhoria & " HI "
    strSql = strSql & "WHERE HI.intContribuicao = " & intCodContribuicao
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set oList = lvw_Historico.ListItems.Add(, , Trim(!Historico))
                .MoveNext
            Loop
        End With
    End If
    If lvw_Historico.ListItems.Count <> 0 Then
        lvw_Historico.SelectedItem.Selected = False
    End If
End Sub

Function blnGravaTributos(blnAlterando As Boolean, intContribuicao As Integer) As Boolean

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'        strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim i      As Integer

    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans

    If blnAlterando Then
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrTributoContribuicaoMelhoria & " "
        strSql = strSql & "WHERE intContribuicao = " & intContribuicao
        gobjBanco.Execute strSql
    End If
    
    tdb_Tributos.MoveFirst
    
    For i = 0 To A.Count(1) - 1
        If A(i, 0) <> "" And Not IsNull(A(i, 0)) And A(i, 0) <> Empty Then
            strSql = ""
            strSql = strSql & "INSERT INTO " & gstrTributoContribuicaoMelhoria & " "
            strSql = strSql & "(intContribuicao, intReceita, "
            strSql = strSql & "dtmDtAtualizacao, lngCodUsr"
            strSql = strSql & ") Values ("
            strSql = strSql & intContribuicao & ", "
            strSql = strSql & A(i, 0) & ", "
'            strSql = strSql & "getdate()" & ", "
            strSql = strSql & strGETDATE & ", "
            strSql = strSql & glngCodUsr
            strSql = strSql & ")"
    
            If Not gobjBanco.Execute(strSql, False) Then
                gobjBanco.ExecutaRollbackTrans
                ExibeMensagem "Ocorreu um erro ao gravar os tributos. Os dados não foram gravados."
                Exit Function
            End If
        End If
    Next i
    gobjBanco.ExecutaCommitTrans
    Set gobjBanco = Nothing
    blnGravaTributos = True
    
End Function

Sub DeletaTributos(intContribuicao As Integer)
    Dim strSql As String

    strSql = ""
    strSql = strSql & "Delete From " & gstrTributoContribuicaoMelhoria & " "
    strSql = strSql & "Where intContribuicao = " & intContribuicao
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
End Sub

Private Sub MontaArray(intFlag As Integer, Optional lngComposicaoDaReceita As Long)
    Dim varAux       As Variant
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    On Error GoTo Err_Handle
    
    Select Case intFlag
        Case 0  'Grid Tributos
            Set A = New XArrayDB
            A.Clear
            
            strSql = ""
            strSql = strSql & "SELECT intReceita "
            strSql = strSql & "FROM " & gstrTributoContribuicaoMelhoria & " "
            strSql = strSql & "WHERE intContribuicao = " & Val(txtPKID)

            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strSql, 5, adoResultado
            With adoResultado
                If Not .EOF Then
                    A.ReDim 0, .RecordCount - 1, 0, 0
                    Do While Not .EOF
                        varAux = !intReceita
                        A(.AbsolutePosition - 1, 0) = varAux
                        .MoveNext
                    Loop
                Else
                    A.ReDim 0, 0, 0, 0
                    A(0, 0) = ""
                End If
            End With
            Set tdb_Tributos.Array = A
            tdb_Tributos.Rebind
            tdb_Tributos.Refresh
    
        Case 1 'DropDown Tributos
            strSql = ""
            strSql = strSql & " SELECT A.PKId, A.strDescricao, A.strSigla FROM "
            strSql = strSql & gstrReceita & " A,"
            strSql = strSql & gstrValorCompoRec & " B"
            strSql = strSql & " WHERE A.PKId = B.intReceita "
            strSql = strSql & " AND B.intComposicaoDaReceita = " & lngComposicaoDaReceita

            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strSql, 5, adoResultado
            With adoResultado
                If Not .EOF Then
                    B.ReDim 0, .RecordCount - 1, 0, 2
                    Do While Not .EOF
                        varAux = !Pkid
                        B(.AbsolutePosition - 1, 0) = varAux

                        varAux = !strDescricao
                        B(.AbsolutePosition - 1, 1) = varAux

                        varAux = !strSigla
                        B(.AbsolutePosition - 1, 2) = varAux

                        .MoveNext
                    Loop
                Else
                    B.ReDim 0, 0, 0, 2
                    B(0, 0) = ""
                    B(0, 1) = ""
                    B(0, 2) = ""
                End If
            End With
            Set tdd_Tributos.Array = B
            tdd_Tributos.Rebind
            tdd_Tributos.Refresh
            
        Case 2
            Set X = New XArrayDB
            X.Clear
            
            strSql = ""
            strSql = strSql & "SELECT IM.PKId, CO.strNome, CO.strCNPJCPF, "
            strSql = strSql & gstrRIGHT("IM.strInscricaoAnterior", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoAnterior, "
            strSql = strSql & "FROM " & gstrImobiliario & " IM, " & gstrContribuinte & " CO "
            strSql = strSql & "WHERE IM.intContribuinte = CO.PKId "
            strSql = strSql & "ORDER BY strNome"
            
            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strSql, 5, adoResultado
            With adoResultado
                If Not .EOF Then
                    X.ReDim 0, .RecordCount - 1, 0, 3
                    Do While Not .EOF
                        varAux = !Pkid
                        X(.AbsolutePosition - 1, 0) = varAux
                    
                        varAux = !strInscricaoAnterior
                        X(.AbsolutePosition - 1, 1) = varAux

                        varAux = !STRNOME
                        X(.AbsolutePosition - 1, 2) = varAux

                        varAux = gstrCGCCPFFormatado(!StrCnpjCpf)
                        X(.AbsolutePosition - 1, 3) = varAux

                        .MoveNext
                    Loop
                Else
                    X.ReDim 0, 0, 0, 3
                    X(0, 0) = ""
                    X(0, 1) = ""
                    X(0, 2) = ""
                    X(0, 3) = ""
                End If
            End With
            Set tdb_Imoveis.Array = X
            tdb_Imoveis.Rebind
            tdb_Imoveis.Refresh
            
        Case 3
            Set X = New XArrayDB
            X.Clear
            
            strSql = ""
            strSql = strSql & "SELECT intMedidaDaArea, "
            strSql = strSql & " intNPavimento, intNEdificacao, dtmUltimaReforma"
            strSql = strSql & " FROM " & gstrAreaImobiliario
            strSql = strSql & " WHERE intImobiliario = " & txtintImobiliario.Text
            'strSql = strSql & " AND intArea = " & dbcintArea.BoundText
            strSql = strSql & " AND intTipoDeArea = " & dbcintArea.BoundText
            
            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strSql, 5, adoResultado
            With adoResultado
                If Not .EOF Then
                    X.ReDim 0, .RecordCount - 1, 0, 3
                    Do While Not .EOF
                        varAux = !intMedidaDaArea
                        X(.AbsolutePosition - 1, 0) = varAux
                    
                        varAux = !intNPavimento
                        X(.AbsolutePosition - 1, 1) = varAux

                        varAux = !intNEdificacao
                        X(.AbsolutePosition - 1, 2) = varAux

                        varAux = !dtmUltimaReforma
                        X(.AbsolutePosition - 1, 3) = varAux

                        .MoveNext
                    Loop
                Else
                    X.ReDim 0, 0, 0, 3
                    X(0, 0) = ""
                    X(0, 1) = ""
                    X(0, 2) = ""
                    X(0, 3) = ""
                End If
            End With
            Set tdb_Area.Array = X
            tdb_Area.Rebind
            tdb_Area.Refresh
            
        Case 4
            Set X1 = New XArrayDB
            X1.Clear
            
            strSql = ""
            
            strSql = ""
            strSql = strSql & "SELECT strMedidaDaTestada "
            strSql = strSql & " FROM " & gstrTestadaImobiliario
            strSql = strSql & " WHERE intImobiliario = " & txtintImobiliario.Text
            'strSql = strSql & " AND intTestada = " & dbcintTestada.BoundText
            strSql = strSql & " AND intTipoDeTestada = '" & dbcintTestada.BoundText & "'"
            
            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strSql, 5, adoResultado
            With adoResultado
                If Not .EOF Then
                    X1.ReDim 0, .RecordCount - 1, 0, 0
                    Do While Not .EOF
                        varAux = !strMedidaDaTestada
                        X1(.AbsolutePosition - 1, 0) = varAux
                        .MoveNext
                    Loop
                Else
                    X1.ReDim 0, 0, 0, 0
                    X1(0, 0) = ""
                End If
            End With
            Set tdb_Testada.Array = X1
            tdb_Testada.Rebind
            tdb_Testada.Refresh
                               
    End Select
Err_Handle:
End Sub

Private Sub LimpaGrid()
    Set A = New XArrayDB 'Grid Tributos
    
    A.Clear
    A.ReDim 0, 0, 0, 0
    
    Set tdb_Tributos.Array = A
    tdb_Tributos.Rebind
    tdb_Tributos.Refresh
End Sub

Private Sub LimpaGridArea()
    Set A = New XArrayDB
    
    A.Clear
    A.ReDim 0, 0, 0, 3
    
    Set tdb_Area.Array = A
    tdb_Area.Rebind
    tdb_Area.Refresh
End Sub

Private Sub LimpaGridTestada()
    Set A = New XArrayDB
    
    A.Clear
    A.ReDim 0, 0, 0, 0
    
    Set tdb_Testada.Array = A
    tdb_Testada.Rebind
    tdb_Testada.Refresh
End Sub


Sub VerificaMascaraInscricao()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
    
    strMascara = ""
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & TYP_IMOBILIARIA
    strSql = strSql & "Order By intSequencia"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    msk_strInscricaoCadastral.Mask = strMascara
End Sub

Sub CarregaDadosImovel(Optional blnInscricao As Boolean = True, Optional strInscricao As String)

'******************************************************************************************
' Data: 08/04/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
'            Foi mantida a forma antiga para o SQL Server pois não era possível o
'            deslocamento completo devido à incompatibilidade entres os bancos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/05/2003
' Alteração: - Alterado o nome do atributo strArea da tabela tblImobiliario para dblArea.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    On Error GoTo err_CarregaDadosImovel
    
    LimpaDadosImovel
    
    If blnInscricao And Trim(strInscricao) = "" Then
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & "SELECT IM.PKId, IM.strDesmembramento, IM.strQuadra, IM.strHabitese, "
    strSql = strSql & "LT.Strnome as strLoteamento, IM.strLote, IM.strMatricula, IM.intContribuinte, "
    strSql = strSql & gstrRIGHT("IM.strInscricaoAnterior", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoAnterior, "
    strSql = strSql & "CO.strNome AS Contribuinte, CO.strCNPJCPF, "
    strSql = strSql & "LO.strDescricao AS Logradouro, IM.intNumero, IM.strComplemento, "
'    strSql = strSql & "BA.strDescricao AS Bairro, IM.intCEP AS Cep, IM.strArea AS Area, "
    strSql = strSql & "BA.strDescricao AS Bairro, IM.intCEP AS Cep, IM.dblArea AS Area, "
    'strSql = strSql & "CR.PKId AS Composicao, OC.strDescricao AS Ocorrrencia, "
    strSql = strSql & "OC.strDescricao AS Ocorrrencia, "
    strSql = strSql & "TP.strDescricao AS TipoLogradouro, "
    strSql = strSql & "TT.strDescricao AS TituloLogradouro, "
    strSql = strSql & "COP.strNome AS Promissario, COP.strCNPJCPF AS CNPJCPFPromissario, A.PKId AS SecaoLogradouro "
    strSql = strSql & "FROM " & gstrSecaoLogradouro & " A, " & gstrTabelaDeEdital & " B, " & gstrEditalSecaoLogradouro & " ESL, "
    
    If (bytDBType = EDatabases.SQLServer) Then
        strSql = strSql & "(((((((" & gstrImobiliario & " IM "
        strSql = strSql & "LEFT JOIN " & gstrContribuinte & " CO ON IM.intContribuinte = CO.PKId) "
        strSql = strSql & "LEFT JOIN " & gstrOcorrencia & " OC ON IM.intOcorrrencia = OC.PKId) "
        'strSql = strSql & "LEFT JOIN " & gstrComposicaoDaReceita & " CR ON IM.intComposicao = CR.PKId) "
        strSql = strSql & "LEFT JOIN " & gstrBairro & " BA ON IM.intBairro = BA.PKId) "
        strSql = strSql & "LEFT JOIN " & gstrLogradouro & " LO ON IM.intLogradouro = LO.PKId) "
        strSql = strSql & "LEFT JOIN " & gstrTipoLogradouro & " TP ON LO.intTipoLogradouro = TP.PKId) "
        strSql = strSql & "LEFT JOIN " & gstrTituloLogradouro & " TT ON LO.intTituloLogradouro = TT.PKId) "
        strSql = strSql & "LEFT JOIN " & gstrContribuinte & " COP ON IM.intPromissario = COP.PKId "
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        strSql = strSql & gstrImobiliario & " IM, "
        strSql = strSql & gstrContribuinte & " CO, "
        strSql = strSql & gstrOcorrencia & " OC, "
        'strSql = strSql & gstrComposicaoDaReceita & " CR, "
        strSql = strSql & gstrBairro & " BA, "
        strSql = strSql & gstrLogradouro & " LO, "
        strSql = strSql & gstrTipoLogradouro & " TP, "
        strSql = strSql & gstrTituloLogradouro & " TT, "
        strSql = strSql & gstrContribuinte & " COP, "
        strSql = strSql & gstrLoteamento & " LT "
    
    End If
    
    strSql = strSql & " WHERE A.PKId = ESL.intSecaoDeLogradouro "
    strSql = strSql & " AND B.PKId = ESL.intTabelaDeEdital "
    strSql = strSql & " AND A.PKId = IM.intSecoes "
    
    If (bytDBType = EDatabases.Oracle) Then
        strSql = strSql & " AND IM.intContribuinte = CO.PKId " & strOUTJOracle
        strSql = strSql & " AND IM.intOcorrrencia = OC.PKId " & strOUTJOracle
        'strSql = strSql & " AND IM.intComposicao = CR.PKId " & strOUTJOracle
        strSql = strSql & " AND IM.intBairro = BA.PKId " & strOUTJOracle
        strSql = strSql & " AND IM.intLogradouro = LO.PKId " & strOUTJOracle
        strSql = strSql & " AND LO.intTipoLogradouro = TP.PKId " & strOUTJOracle
        strSql = strSql & " AND LO.intTituloLogradouro = TT.PKId " & strOUTJOracle
        strSql = strSql & " AND IM.intPromissario = COP.PKId " & strOUTJOracle
        strSql = strSql & " AND IM.Intloteamento = LT.Pkid"
    
    End If
    
    If blnInscricao Then
        strSql = strSql & " AND IM.strInscricaoAnterior = '" & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & Trim(strInscricao) & "'"
    Else
        strSql = strSql & " AND CO.PKId = " & Val(txtPKID)
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                txtintImobiliario = !Pkid
                txt_PKIDIMobiliario = !Pkid
                msk_strInscricaoCadastral = !strInscricaoAnterior
                txt_Desmembramento = gstrVerificaCampoNulo(!strDesmembramento)
                txt_Quadra = gstrVerificaCampoNulo(!strQuadra)
                txt_Habitese = gstrVerificaCampoNulo(!strHabitese)
                txt_Loteamento = gstrVerificaCampoNulo(!strLoteamento)
                txt_Lote = gstrVerificaCampoNulo(!strLote)
                txt_Escritura = gstrVerificaCampoNulo(!strMatricula)
                txt_PKIdContribuinte = !intContribuinte
                txt_Contribuinte = !Contribuinte
                txt_CNPJCPF = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!StrCnpjCpf))
                txt_Logradouro = gstrEnderecoConcatenado(!Logradouro, !TipoLogradouro, !INTNUMERO, !STRCOMPLEMENTO, , !TituloLogradouro)
                txt_Bairro = gstrVerificaCampoNulo(!Bairro)
                txt_Cep = gstrCEPFormatado(!CEP)
                txt_Area = gstrVerificaCampoNulo(!Area)
                txt_Ocorrrencia = gstrVerificaCampoNulo(!Ocorrrencia)
                txt_Promissario = gstrVerificaCampoNulo(!Promissario)
                txt_CNPJCPFPromissario = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!CNPJCPFPromissario))
                dbcintSecaoDeLogradouro.BoundText = !SecaoLogradouro
            Else
                ExibeMensagem "Imóvel não encontrado."
            End If
        End With
        adoResultado.Close
        Set adoResultado = Nothing
        Set gobjBanco = Nothing
    End If
    
Exit Sub
err_CarregaDadosImovel:
    ExibeDetalheErro ""
End Sub

Sub LimpaDadosImovel()
    txt_PKIDIMobiliario = ""
    msk_strInscricaoCadastral = ""
    txt_Desmembramento = ""
    txt_Quadra = ""
    txt_Habitese = ""
    txt_Loteamento = ""
    txt_Lote = ""
    txt_Escritura = ""
    txt_PKIdContribuinte = ""
    txt_Contribuinte = ""
    txt_CNPJCPF = ""
    txt_Logradouro = ""
    txt_Bairro = ""
    txt_Cep = ""
    txt_Area = ""
    dbc_intComposicaoDaReceita = ""
    txt_Ocorrrencia = ""
    txt_Promissario = ""
    txt_CNPJCPFPromissario = ""
End Sub

Private Function strQuerySecao() As String

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela variável
'            gstrISNULL.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'        pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 08/04/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
'            Foi mantida a forma antiga para o SQL Server pois não era possível o
'            deslocamento completo devido à incompatibilidade entres os bancos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT SL.PKId , "
'    strSql = strSql & "ISNULL(SL.strInscricaoCadastral, '') + ' - ' + RTRIM(LTRIM(ISNULL(TL.strSigla, '') + ' ' + ISNULL(U.strDescricao,'') + ' ' + L.strDescricao)) AS Logradouro "
    strSql = strSql & gstrISNULL("SL.strInscricaoCadastral", "''") & strCONCAT & " ' - ' " & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & strCONCAT & " ' ' " & strCONCAT & gstrISNULL("U.strDescricao", "''") & strCONCAT & " ' ' " & strCONCAT & " L.strDescricao)) AS Logradouro "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTabelaDeEdital & " TE, "
    strSql = strSql & gstrEditalSecaoLogradouro & " ESL, "
    
    If (bytDBType = EDatabases.SQLServer) Then
        strSql = strSql & " ((" & gstrSecaoLogradouro & " SL "
        strSql = strSql & "LEFT JOIN " & gstrLogradouro & " L ON SL.intLogradouro = L.PKId) "
        strSql = strSql & "LEFT JOIN " & gstrTituloLogradouro & " U ON L.intTituloLogradouro = U.PKId) "
        strSql = strSql & "LEFT JOIN " & gstrTipoLogradouro & " TL ON L.intTipoLogradouro = TL.PKId "
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        strSql = strSql & gstrSecaoLogradouro & " SL, "
        strSql = strSql & gstrLogradouro & " L, "
        strSql = strSql & gstrTituloLogradouro & " U, "
        strSql = strSql & gstrTipoLogradouro & " TL "
    
    End If
    
    strSql = strSql & "WHERE "
    strSql = strSql & " SL.PKID = ESL.intSecaoDeLogradouro "
    strSql = strSql & " AND TE.PKID = ESL.intTabelaDeEdital "
    
    If (bytDBType = EDatabases.Oracle) Then
        strSql = strSql & " AND SL.intLogradouro = L.PKId " & strOUTJOracle
        strSql = strSql & " AND L.intTituloLogradouro = U.PKId " & strOUTJOracle
        strSql = strSql & " AND L.intTipoLogradouro = TL.PKId " & strOUTJOracle
    
    End If
    
    strSql = strSql & " AND TE.PKId = " & dbcintTabelaDeEdital.BoundText
    strSql = strSql & " ORDER BY Logradouro"
    strQuerySecao = strSql
End Function

Sub GravaVariosImoveis()
    Dim i               As Integer
    Dim j               As Integer
    Dim intContribuicao As Integer
    
    On Error GoTo err_GravaVariosImoveis
    
    tdb_Imoveis.MoveFirst
    
    For j = 0 To tdb_Imoveis.SelBookmarks.Count - 1
        i = tdb_Imoveis.SelBookmarks(j)
        txtintImobiliario = X(i, 0)
        mblnAlterando = False
        If SalvarGeral(gstrContribuicaoMelhoria, "I", Me, tdb_Lista, strQuery, False) Then
            intContribuicao = glngPegaUltimaChave(gstrContribuicaoMelhoria, "PKId")
            If blnGravaTributos(mblnAlterando, intContribuicao) Then
                If blnGravaHistoricos(mblnAlterando, intContribuicao) Then
                    If blnGravaMelhorias(mblnAlterando, intContribuicao) Then
                    End If
                End If
            End If
        End If
    Next j
    
    LimpaObjeto Me, mblnAlterando
    NovaContribuicao
    
    Exit Sub
err_GravaVariosImoveis:
End Sub

Private Sub txt_Promissario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Promissario
End Sub

Private Function strQueryRelatorio() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT TE.strNomeDoEdital, "
    strSql = strSql & gstrRIGHT("IM.strInscricaoAnterior", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoAnterior, "
    strSql = strSql & "CT.strNome, TE.dblCustoDaParcela, TE.dblCustoDeTerceiros "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTabelaDeEdital & " TE, "
    strSql = strSql & gstrImobiliario & " IM, "
    strSql = strSql & gstrContribuinte & " CT "
    strSql = strSql & "WHERE TE.PKId = IM.PKId "
    strSql = strSql & "AND IM.intContribuinte = CT.PKId "
    strQueryRelatorio = strSql
End Function


Private Sub PreencheComboInscricao()
Dim strSql As String

strSql = "SELECT Pkid, "
strSql = strSql & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao "
strSql = strSql & "FROM "
strSql = strSql & gstrImobiliario
strSql = strSql & " ORDER BY " & gstrCONVERT(cdt_numeric, "strInscricao")

LeDaTabelaParaObj "", dbc_strInscricao, strSql

End Sub
