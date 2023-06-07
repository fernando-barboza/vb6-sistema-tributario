VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadImobiliario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imobiliário Urbano"
   ClientHeight    =   7770
   ClientLeft      =   2460
   ClientTop       =   2265
   ClientWidth     =   10440
   HelpContextID   =   8
   Icon            =   "frmCadImobiliario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   518
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   696
   Begin VB.TextBox txtPKId 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   211
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   870
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   7695
      Left            =   45
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   30
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      TabsPerRow      =   11
      TabHeight       =   529
      TabCaption(0)   =   "Urbano"
      TabPicture(0)   =   "frmCadImobiliario.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_Lista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Contribuinte"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Promissário"
      TabPicture(1)   =   "frmCadImobiliario.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_Inscricao"
      Tab(1).Control(1)=   "txt_Proprietario"
      Tab(1).Control(2)=   "fra_Promissario"
      Tab(1).Control(3)=   "lbl_InscricaoCadastral1"
      Tab(1).Control(4)=   "lbl_Proprietario7"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Envolvidos"
      TabPicture(2)   =   "frmCadImobiliario.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt_Proprietario7"
      Tab(2).Control(1)=   "txt_Inscricao7"
      Tab(2).Control(2)=   "fra_Envolvido"
      Tab(2).Control(3)=   "lvw_Envolvidos"
      Tab(2).Control(4)=   "Label2"
      Tab(2).Control(5)=   "Label1"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Áreas"
      TabPicture(3)   =   "frmCadImobiliario.frx":1096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Geral"
      TabPicture(4)   =   "frmCadImobiliario.frx":10B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txt_Proprietario2"
      Tab(4).Control(1)=   "txt_Inscricao2"
      Tab(4).Control(2)=   "lvw_Caracteristica(0)"
      Tab(4).Control(3)=   "lvw_Detalhe(0)"
      Tab(4).Control(4)=   "lbl_Proprietario2"
      Tab(4).Control(5)=   "lbl_InscricaoCadastral3"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Terreno"
      TabPicture(5)   =   "frmCadImobiliario.frx":10CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txt_dblArea"
      Tab(5).Control(1)=   "fra_FatorTerreno"
      Tab(5).Control(2)=   "fra_Testada"
      Tab(5).Control(3)=   "txt_Proprietario3"
      Tab(5).Control(4)=   "txt_Inscricao3"
      Tab(5).Control(5)=   "lvw_Caracteristica(1)"
      Tab(5).Control(6)=   "lvw_Detalhe(1)"
      Tab(5).Control(7)=   "lbl_strArea"
      Tab(5).Control(8)=   "lbl_Proprietario3"
      Tab(5).Control(9)=   "lbl_InscricaoCadastral4"
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Construção"
      TabPicture(6)   =   "frmCadImobiliario.frx":10EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txt_TotPredial"
      Tab(6).Control(1)=   "fra_CaracteristicaPredio"
      Tab(6).Control(2)=   "fra_Area"
      Tab(6).Control(3)=   "txt_Proprietario4"
      Tab(6).Control(4)=   "txt_Inscricao4"
      Tab(6).Control(5)=   "lvw_Caracteristica(2)"
      Tab(6).Control(6)=   "lvw_Detalhe(2)"
      Tab(6).Control(7)=   "dbc_intResumoTipoPadrao"
      Tab(6).Control(8)=   "lbl_TotPredial"
      Tab(6).Control(9)=   "lbl_Resumo"
      Tab(6).Control(10)=   "lbl_Proprietario4"
      Tab(6).Control(11)=   "lbl_InscricaoCadastral5"
      Tab(6).ControlCount=   12
      TabCaption(7)   =   "Equipamentos"
      TabPicture(7)   =   "frmCadImobiliario.frx":1106
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fra_Secoes"
      Tab(7).Control(1)=   "txt_Proprietario5"
      Tab(7).Control(2)=   "txt_Inscricao5"
      Tab(7).Control(3)=   "lbl_Proprietario5"
      Tab(7).Control(4)=   "lbl_InscricaoCadastral6"
      Tab(7).ControlCount=   5
      TabCaption(8)   =   "Históricos"
      TabPicture(8)   =   "frmCadImobiliario.frx":1122
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "fra_Historico"
      Tab(8).Control(1)=   "txt_Proprietario6"
      Tab(8).Control(2)=   "txt_Inscricao6"
      Tab(8).Control(3)=   "fra_Englobado"
      Tab(8).Control(4)=   "lbl_Proprietario6(0)"
      Tab(8).Control(5)=   "lbl_InscricaoCadastral7(0)"
      Tab(8).ControlCount=   6
      TabCaption(9)   =   "Doc/Proc"
      TabPicture(9)   =   "frmCadImobiliario.frx":113E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "fra_DocumentosProcessos"
      Tab(9).Control(1)=   "txt_Inscricao8"
      Tab(9).Control(2)=   "txt_Proprietario8"
      Tab(9).Control(3)=   "lbl_InscricaoCadastral8(1)"
      Tab(9).Control(4)=   "lbl_Proprietario8(1)"
      Tab(9).ControlCount=   5
      TabCaption(10)  =   "Fichas Cadastrais"
      TabPicture(10)  =   "frmCadImobiliario.frx":115A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lbl_CaminhoFicha"
      Tab(10).Control(1)=   "cmd_FichaAnterior"
      Tab(10).Control(2)=   "cmd_FichaPosterior"
      Tab(10).Control(3)=   "rtb_FichaCadastral"
      Tab(10).ControlCount=   4
      Begin RichTextLib.RichTextBox rtb_FichaCadastral 
         Height          =   6345
         Left            =   -74880
         TabIndex        =   224
         Top             =   450
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   11192
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmCadImobiliario.frx":1176
      End
      Begin VB.CommandButton cmd_FichaPosterior 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -65280
         TabIndex        =   103
         Top             =   6870
         Width           =   435
      End
      Begin VB.CommandButton cmd_FichaAnterior 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -65730
         TabIndex        =   102
         Top             =   6870
         Width           =   435
      End
      Begin VB.TextBox txt_TotPredial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66585
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   220
         Top             =   420
         Width           =   1125
      End
      Begin VB.TextBox txt_dblArea 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66585
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   218
         Top             =   420
         Width           =   1125
      End
      Begin VB.Frame fra_CaracteristicaPredio 
         Caption         =   "Caracterísitica do Prédio"
         Height          =   2475
         Left            =   -74280
         TabIndex        =   212
         Top             =   5130
         Width           =   8880
         Begin MSComctlLib.ListView lvw_CaracPredio 
            Height          =   1695
            Left            =   60
            TabIndex        =   82
            Top             =   720
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   2990
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Característica"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "intDetalheDaCaracteristica"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Detalhes"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Valor"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lbl_DescricaoPredios 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   213
            Top             =   330
            Width           =   75
         End
      End
      Begin VB.Frame fra_FatorTerreno 
         Caption         =   "Fatores do Terreno"
         Height          =   2565
         Left            =   -74310
         TabIndex        =   210
         Top             =   5040
         Width           =   8910
         Begin MSComctlLib.ListView lvw_FatorTerreno 
            Height          =   2205
            Left            =   90
            TabIndex        =   76
            Top             =   210
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   3889
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Característica"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "intDetalheDaCaracteristica"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Detalhes"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Valor"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fra_Area 
         Caption         =   "Áreas"
         Height          =   1650
         Left            =   -74325
         TabIndex        =   187
         Top             =   1155
         Width           =   8910
         Begin TrueOleDBGrid70.TDBDropDown tdd_CategoriaConstrucao 
            Height          =   945
            Left            =   4560
            TabIndex        =   189
            Top             =   600
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   1667
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
            Columns(1).Caption=   "Categoria de Construção"
            Columns(1).DataField=   "strDescricao"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8811"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8731"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   "strDescricao"
            DataField       =   "strDescricao"
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
            ValueTranslate  =   -1  'True
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bgcolor=&H80000005&,.fgcolor=&H80000008&"
            _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(20)  =   ":id=8,.fgcolor=&H8000000E&"
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
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
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
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(58)  =   "Named:id=39:EvenRow"
            _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(60)  =   "Named:id=40:OddRow"
            _StyleDefs(61)  =   ":id=40,.parent=33"
            _StyleDefs(62)  =   "Named:id=41:RecordSelector"
            _StyleDefs(63)  =   ":id=41,.parent=34"
            _StyleDefs(64)  =   "Named:id=42:FilterBar"
            _StyleDefs(65)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBDropDown tdd_Area 
            Height          =   945
            Left            =   1515
            TabIndex        =   188
            Top             =   570
            Width           =   5880
            _ExtentX        =   10372
            _ExtentY        =   1667
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
            Columns(1).Caption=   "Tipo de Área"
            Columns(1).DataField=   "strNomeDaArea"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8811"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8731"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   "strNomeDaArea"
            DataField       =   "PKId"
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
            ValueTranslate  =   -1  'True
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid grd_Area 
            Height          =   1365
            Left            =   180
            TabIndex        =   79
            Top             =   210
            Width           =   8580
            _ExtentX        =   15134
            _ExtentY        =   2408
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
            Columns(1).Caption=   "N° da Edificação"
            Columns(1).DataField=   ""
            Columns(1).DataWidth=   3
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   1
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tipo de Área"
            Columns(2).DataField=   "intTipoDeArea"
            Columns(2).DropDown=   "tdd_Area"
            Columns(2).DropDown.vt=   8
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Medida da Área"
            Columns(3).DataField=   ""
            Columns(3).NumberFormat=   "Standard"
            Columns(3).EditMaskUpdate=   -1  'True
            Columns(3).EditMaskRight=   -1  'True
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fração Ideal"
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Categoria da Construção"
            Columns(5).DataField=   ""
            Columns(5).DropDown=   "tdd_CategoriaConstrucao"
            Columns(5).DropDown.vt=   8
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "PkidCategoriaConstrucao"
            Columns(6).DataField=   "PkidCategoriaConstrucao"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "N° de Pavimentos"
            Columns(7).DataField=   ""
            Columns(7).DataWidth=   5
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Última Reforma"
            Columns(8).DataField=   ""
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Quartos"
            Columns(9).DataField=   ""
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Salas"
            Columns(10).DataField=   ""
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Cozinhas"
            Columns(11).DataField=   ""
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Banheiros"
            Columns(12).DataField=   ""
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Andar"
            Columns(13).DataField=   ""
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Elevadores"
            Columns(14).DataField=   ""
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "Suites"
            Columns(15).DataField=   ""
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "Vagas na Garagem"
            Columns(16).DataField=   ""
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   20
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "Hotelaria"
            Columns(17).DataField=   ""
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   18
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=18"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2408"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2328"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8452"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=4022"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3942"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=260"
            Splits(0)._ColumnProps(18)=   "Column(2).Button=1"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(2).AutoDropDown=1"
            Splits(0)._ColumnProps(21)=   "Column(2).AutoCompletion=1"
            Splits(0)._ColumnProps(22)=   "Column(3).Width=2381"
            Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2302"
            Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=258"
            Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(28)=   "Column(4).Width=2249"
            Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=2170"
            Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(33)=   "Column(5).Width=4577"
            Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=4498"
            Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(5).AutoCompletion=1"
            Splits(0)._ColumnProps(39)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(43)=   "Column(6).AllowSizing=0"
            Splits(0)._ColumnProps(44)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(45)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(46)=   "Column(7).Width=2461"
            Splits(0)._ColumnProps(47)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._WidthInPix=2381"
            Splits(0)._ColumnProps(49)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(50)=   "Column(7)._ColStyle=260"
            Splits(0)._ColumnProps(51)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(52)=   "Column(8).Width=2328"
            Splits(0)._ColumnProps(53)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._WidthInPix=2249"
            Splits(0)._ColumnProps(55)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(56)=   "Column(8)._ColStyle=260"
            Splits(0)._ColumnProps(57)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(58)=   "Column(9).Width=1244"
            Splits(0)._ColumnProps(59)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._WidthInPix=1164"
            Splits(0)._ColumnProps(61)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(62)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(63)=   "Column(10).Width=1111"
            Splits(0)._ColumnProps(64)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(65)=   "Column(10)._WidthInPix=1032"
            Splits(0)._ColumnProps(66)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(68)=   "Column(11).Width=1508"
            Splits(0)._ColumnProps(69)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(70)=   "Column(11)._WidthInPix=1429"
            Splits(0)._ColumnProps(71)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(73)=   "Column(12).Width=1535"
            Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=1455"
            Splits(0)._ColumnProps(76)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(77)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(78)=   "Column(13).Width=1270"
            Splits(0)._ColumnProps(79)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(80)=   "Column(13)._WidthInPix=1191"
            Splits(0)._ColumnProps(81)=   "Column(13)._EditAlways=0"
            Splits(0)._ColumnProps(82)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(83)=   "Column(14).Width=1693"
            Splits(0)._ColumnProps(84)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(85)=   "Column(14)._WidthInPix=1614"
            Splits(0)._ColumnProps(86)=   "Column(14)._EditAlways=0"
            Splits(0)._ColumnProps(87)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(88)=   "Column(15).Width=1005"
            Splits(0)._ColumnProps(89)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(90)=   "Column(15)._WidthInPix=926"
            Splits(0)._ColumnProps(91)=   "Column(15)._EditAlways=0"
            Splits(0)._ColumnProps(92)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(93)=   "Column(16).Width=2699"
            Splits(0)._ColumnProps(94)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(95)=   "Column(16)._WidthInPix=2619"
            Splits(0)._ColumnProps(96)=   "Column(16)._EditAlways=0"
            Splits(0)._ColumnProps(97)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(98)=   "Column(17).Width=1455"
            Splits(0)._ColumnProps(99)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(100)=   "Column(17)._WidthInPix=1376"
            Splits(0)._ColumnProps(101)=   "Column(17)._EditAlways=0"
            Splits(0)._ColumnProps(102)=   "Column(17)._ColStyle=1"
            Splits(0)._ColumnProps(103)=   "Column(17).Order=18"
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
            FootLines       =   0
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
            _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=69,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=72,.parent=6,.fgcolor=&H8000000E&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=71,.parent=7,.bgcolor=&H8000000D&,.fgcolor=&H80000005&"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8,.bgcolor=&H8000000D&"
            _StyleDefs(32)  =   ":id=73,.fgcolor=&H80000014&"
            _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
            _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
            _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
            _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=16,.parent=67"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=68"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=69"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=71"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=106,.parent=67,.bgcolor=&H8000000F&,.locked=-1"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=68,.alignment=0"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=69"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=71,.bgcolor=&H8000000F&"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=82,.parent=67"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=68,.alignment=0"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=69"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=71"
            _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=86,.parent=67,.alignment=1"
            _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=83,.parent=68,.alignment=0"
            _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=84,.parent=69"
            _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=85,.parent=71"
            _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=20,.parent=67"
            _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=68"
            _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=69"
            _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=71"
            _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=24,.parent=67"
            _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=68"
            _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=69"
            _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=71"
            _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=28,.parent=67"
            _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=68"
            _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=69"
            _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=71"
            _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=102,.parent=67"
            _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=99,.parent=68,.alignment=0"
            _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=100,.parent=69"
            _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=101,.parent=71"
            _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=110,.parent=67"
            _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=107,.parent=68,.alignment=0"
            _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=108,.parent=69"
            _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=109,.parent=71"
            _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=32,.parent=67"
            _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=68"
            _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=69"
            _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=71"
            _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=50,.parent=67"
            _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=68"
            _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=69"
            _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=71"
            _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=54,.parent=67"
            _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=51,.parent=68"
            _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=52,.parent=69"
            _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=53,.parent=71"
            _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=58,.parent=67"
            _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=55,.parent=68"
            _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=56,.parent=69"
            _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=57,.parent=71"
            _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=46,.parent=67"
            _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=43,.parent=68"
            _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=44,.parent=69"
            _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=45,.parent=71"
            _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=66,.parent=67"
            _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=63,.parent=68"
            _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=64,.parent=69"
            _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=65,.parent=71"
            _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=90,.parent=67"
            _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=87,.parent=68"
            _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=88,.parent=69"
            _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=89,.parent=71"
            _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=94,.parent=67"
            _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=91,.parent=68"
            _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=92,.parent=69"
            _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=93,.parent=71"
            _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=62,.parent=67,.alignment=2"
            _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=59,.parent=68"
            _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=60,.parent=69"
            _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=61,.parent=71"
            _StyleDefs(109) =   "Named:id=33:Normal"
            _StyleDefs(110) =   ":id=33,.parent=0"
            _StyleDefs(111) =   "Named:id=34:Heading"
            _StyleDefs(112) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(113) =   ":id=34,.wraptext=-1"
            _StyleDefs(114) =   "Named:id=35:Footing"
            _StyleDefs(115) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(116) =   "Named:id=36:Selected"
            _StyleDefs(117) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(118) =   "Named:id=37:Caption"
            _StyleDefs(119) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(120) =   "Named:id=38:HighlightRow"
            _StyleDefs(121) =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(122) =   "Named:id=39:EvenRow"
            _StyleDefs(123) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(124) =   "Named:id=40:OddRow"
            _StyleDefs(125) =   ":id=40,.parent=33"
            _StyleDefs(126) =   "Named:id=41:RecordSelector"
            _StyleDefs(127) =   ":id=41,.parent=34"
            _StyleDefs(128) =   "Named:id=42:FilterBar"
            _StyleDefs(129) =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_Testada 
         Caption         =   "Testada"
         Height          =   1770
         Left            =   -74310
         TabIndex        =   182
         Top             =   1200
         Width           =   8910
         Begin TrueOleDBGrid70.TDBDropDown tdd_FaceDeQuadra 
            Height          =   1155
            Left            =   3480
            TabIndex        =   184
            Top             =   495
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   2037
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Face de Quadra"
            Columns(0).DataField=   "strFaceDeQuadra"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "PkidFaceDeQuadra"
            Columns(1).DataField=   "PkidFaceDeQuadra"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   ""
            DataField       =   ""
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
            ValueTranslate  =   0   'False
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
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
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(20)  =   ":id=8,.fgcolor=&H8000000E&"
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
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(58)  =   "Named:id=39:EvenRow"
            _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(60)  =   "Named:id=40:OddRow"
            _StyleDefs(61)  =   ":id=40,.parent=33"
            _StyleDefs(62)  =   "Named:id=41:RecordSelector"
            _StyleDefs(63)  =   ":id=41,.parent=34"
            _StyleDefs(64)  =   "Named:id=42:FilterBar"
            _StyleDefs(65)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBDropDown tdd_Testada 
            Height          =   1155
            Left            =   825
            TabIndex        =   183
            Top             =   510
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   2037
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
            Columns(1).Caption=   "Tipo de Testada"
            Columns(1).DataField=   "strNomeDaTestada"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "bytPrincipal"
            Columns(2).DataField=   "bytPrincipal"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7408"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7329"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   "strNomeDaTestada"
            DataField       =   "PKId"
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
            ValueTranslate  =   -1  'True
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
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
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(20)  =   ":id=8,.fgcolor=&H8000000E&"
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
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(49)  =   "Named:id=33:Normal"
            _StyleDefs(50)  =   ":id=33,.parent=0"
            _StyleDefs(51)  =   "Named:id=34:Heading"
            _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   ":id=34,.wraptext=-1"
            _StyleDefs(54)  =   "Named:id=35:Footing"
            _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   "Named:id=36:Selected"
            _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=37:Caption"
            _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(60)  =   "Named:id=38:HighlightRow"
            _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(62)  =   "Named:id=39:EvenRow"
            _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(64)  =   "Named:id=40:OddRow"
            _StyleDefs(65)  =   ":id=40,.parent=33"
            _StyleDefs(66)  =   "Named:id=41:RecordSelector"
            _StyleDefs(67)  =   ":id=41,.parent=34"
            _StyleDefs(68)  =   "Named:id=42:FilterBar"
            _StyleDefs(69)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid grd_Testada 
            Height          =   1425
            Left            =   165
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   240
            Width           =   8610
            _ExtentX        =   15187
            _ExtentY        =   2514
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
            Columns(1).Caption=   "Tipo de Testada"
            Columns(1).DataField=   "strNomeDaTestada"
            Columns(1).DropDown=   "tdd_Testada"
            Columns(1).DropDown.vt=   8
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Medida da Testada"
            Columns(2).DataField=   ""
            Columns(2).DataWidth=   6
            Columns(2).NumberFormat=   "Standard"
            Columns(2).EditMaskUpdate=   -1  'True
            Columns(2).EditMaskRight=   -1  'True
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Face de Quadra"
            Columns(3).DataField=   "strFaceDeQuadra"
            Columns(3).DropDown=   "tdd_FaceDeQuadra"
            Columns(3).DropDown.vt=   8
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "PkidFaceDeQuadra"
            Columns(4).DataField=   "PkidFaceDeQuadra"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "bytPrincipal"
            Columns(5).DataField=   "bytPrincipal"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "intTipoDeTestada"
            Columns(6).DataField=   "intTipoDeTestada"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2275"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=4604"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4524"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=260"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(1).AutoCompletion=1"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=2858"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2778"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=258"
            Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(21)=   "Column(3).Width=7038"
            Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=6959"
            Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=260"
            Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(27)=   "Column(3).AutoCompletion=1"
            Splits(0)._ColumnProps(28)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=260"
            Splits(0)._ColumnProps(33)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(35)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(38)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=260"
            Splits(0)._ColumnProps(40)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(42)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(46)=   "Column(6).AllowSizing=0"
            Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=260"
            Splits(0)._ColumnProps(48)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
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
            TabAction       =   2
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
            _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.alignment=0"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=69,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=71,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=67"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=68"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=69"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=71"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=82,.parent=67"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=68,.alignment=0"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=69"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=71"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=86,.parent=67,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=83,.parent=68,.alignment=0"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=84,.parent=69"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=85,.parent=71"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=102,.parent=67,.locked=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=99,.parent=68,.alignment=0"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=100,.parent=69"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=101,.parent=71"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=20,.parent=67"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=68"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=69"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=71"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=24,.parent=67"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=68"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=69"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=71"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=67"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=68"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=69"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=71"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_DocumentosProcessos 
         Height          =   4740
         Left            =   -73845
         TabIndex        =   203
         Top             =   1395
         Width           =   7905
         Begin VB.TextBox txt_strCodigo 
            CausesValidation=   0   'False
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   285
            HideSelection   =   0   'False
            Left            =   1965
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   96
            Top             =   900
            Width           =   825
         End
         Begin VB.TextBox txt_intExercicio 
            CausesValidation=   0   'False
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   2820
            MaxLength       =   4
            TabIndex        =   97
            Top             =   900
            Width           =   465
         End
         Begin VB.TextBox txt_bitDigito 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   3300
            MaxLength       =   2
            TabIndex        =   98
            Top             =   900
            Width           =   285
         End
         Begin VB.TextBox txt_strObservacoes 
            Height          =   795
            Left            =   1965
            MaxLength       =   250
            TabIndex        =   100
            Top             =   1320
            Width           =   5685
         End
         Begin VB.TextBox txt_PkidDocProc 
            Height          =   300
            Left            =   6885
            TabIndex        =   95
            Top             =   450
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txt_DtmDataDocProc 
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
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   99
            Top             =   900
            Width           =   975
         End
         Begin VB.CommandButton cmd_TipoDocumentoProcesso 
            Height          =   315
            Left            =   6075
            Picture         =   "frmCadImobiliario.frx":11F8
            Style           =   1  'Graphical
            TabIndex        =   205
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Documentos"
            Top             =   435
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbc_intTiposDocumentosProcesso 
            Height          =   315
            Left            =   1965
            TabIndex        =   94
            Top             =   450
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_DocumentosProcessos 
            Height          =   2295
            Left            =   210
            TabIndex        =   101
            Top             =   2265
            Width           =   7440
            _ExtentX        =   13123
            _ExtentY        =   4048
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
            Columns(1).Caption=   "Documento/Processo"
            Columns(1).DataField=   "DocumentoProcesso"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Processo"
            Columns(2).DataField=   "Processo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Data do Processo"
            Columns(3).DataField=   "DataProc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Observações"
            Columns(4).DataField=   "strObservacoes"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "PkidDocumentoProcesso"
            Columns(5).DataField=   "PkidDocumentoProcesso"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "PkidProcesso"
            Columns(6).DataField=   "PkidProcesso"
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
            Splits(0)._ColumnProps(8)=   "Column(1).Width=7382"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7303"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2963"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2884"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=2381"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2302"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=5874"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=5794"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(38)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
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
            _StyleDefs(24)  =   "Splits(0).Style:id=51,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=60,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=52,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=53,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=54,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=56,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=55,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=57,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=58,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=59,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=61,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=62,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=51"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=52"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=53"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=55"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=51"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=52"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=53"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=55"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=51"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=52"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=53"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=55"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=51"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=52"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=53"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=55"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=24,.parent=51"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=21,.parent=52"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=22,.parent=53"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=23,.parent=55"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=16,.parent=51"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=52"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=53"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=55"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=20,.parent=51"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=52"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=53"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=55"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lbl_Observação 
            AutoSize        =   -1  'True
            Caption         =   "Observações"
            Height          =   195
            Left            =   930
            TabIndex        =   208
            Top             =   1335
            Width           =   945
         End
         Begin VB.Label lbl_DocumentosProcessos 
            AutoSize        =   -1  'True
            Caption         =   "Documentos"
            Height          =   195
            Left            =   990
            TabIndex        =   204
            Top             =   495
            Width           =   900
         End
         Begin VB.Label lbl_Processo 
            AutoSize        =   -1  'True
            Caption         =   "Processo"
            Height          =   195
            Left            =   1215
            TabIndex        =   206
            Top             =   945
            Width           =   660
         End
         Begin VB.Label lbl_DataDocumentoProcesso 
            Caption         =   "Data"
            Height          =   240
            Left            =   4485
            TabIndex        =   207
            Top             =   945
            Width           =   465
         End
      End
      Begin VB.TextBox txt_Inscricao8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72675
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   92
         Top             =   495
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario8 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72675
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   93
         Top             =   855
         Width           =   4650
      End
      Begin VB.TextBox txt_Proprietario7 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   55
         Top             =   855
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   54
         Top             =   495
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario2 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72675
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   70
         Top             =   855
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72675
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   69
         Top             =   495
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario3 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   181
         Top             =   855
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   179
         Top             =   495
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario4 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   78
         Top             =   855
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   77
         Top             =   495
         Width           =   2130
      End
      Begin VB.Frame fra_Secoes 
         Height          =   3720
         Left            =   -73875
         TabIndex        =   192
         Top             =   1245
         Width           =   8070
         Begin MSComctlLib.ListView lvw_Melhoria 
            Height          =   2850
            Left            =   1980
            TabIndex        =   86
            Top             =   630
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   5027
            View            =   3
            LabelEdit       =   1
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
         Begin MSDataListLib.DataCombo dbc_intFaceDeQuadra 
            Height          =   315
            Left            =   2010
            TabIndex        =   85
            Top             =   240
            Width           =   5805
            _ExtentX        =   10239
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_FaceDeQuadra 
            AutoSize        =   -1  'True
            Caption         =   "Face de Quadra"
            Height          =   195
            Left            =   315
            TabIndex        =   193
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label lbl_BLABLEVLIDB 
            AutoSize        =   -1  'True
            Caption         =   "Equipamentos"
            Height          =   195
            Left            =   450
            TabIndex        =   194
            Top             =   705
            Width           =   1005
         End
      End
      Begin VB.TextBox txt_Proprietario5 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   84
         Top             =   855
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72615
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   83
         Top             =   495
         Width           =   2130
      End
      Begin VB.Frame fra_Historico 
         Caption         =   "Históricos"
         Height          =   3450
         Left            =   -73815
         TabIndex        =   197
         Top             =   1230
         Width           =   8025
         Begin VB.TextBox txt_Historico 
            Height          =   1260
            Left            =   150
            MaxLength       =   4000
            MultiLine       =   -1  'True
            TabIndex        =   89
            Top             =   225
            Width           =   7740
         End
         Begin MSComctlLib.Toolbar tlb_Historico 
            Height          =   330
            Left            =   135
            TabIndex        =   198
            Top             =   1605
            Width           =   1050
            _ExtentX        =   1852
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
         Begin MSComctlLib.ListView lvw_Historico 
            Height          =   1260
            Left            =   120
            TabIndex        =   90
            Top             =   2070
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   2223
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
            Left            =   1590
            Top             =   1530
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
                  Picture         =   "frmCadImobiliario.frx":1582
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCadImobiliario.frx":16E2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCadImobiliario.frx":183E
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txt_Proprietario6 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72585
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   88
         Top             =   855
         Width           =   4650
      End
      Begin VB.TextBox txt_Inscricao6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72585
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   87
         Top             =   495
         Width           =   2130
      End
      Begin VB.Frame fra_Englobado 
         Height          =   705
         Left            =   -73830
         TabIndex        =   199
         Top             =   4710
         Width           =   8025
         Begin VB.CheckBox chkbitEnglobado 
            Caption         =   "Englobado"
            Height          =   195
            Left            =   210
            TabIndex        =   91
            Top             =   360
            Width           =   1095
         End
         Begin MSMask.MaskEdBox mskstrInscricaoEnglobada 
            Height          =   285
            Left            =   5520
            TabIndex        =   104
            Top             =   270
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin VB.Label lblstrInscricaoEnglobado 
            AutoSize        =   -1  'True
            Caption         =   "O imóvel sera englobado à inscrição"
            Height          =   195
            Left            =   2910
            TabIndex        =   200
            Top             =   360
            Width           =   2565
         End
      End
      Begin VB.TextBox txt_Inscricao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72645
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         Top             =   495
         Width           =   2130
      End
      Begin VB.TextBox txt_Proprietario 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72645
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   41
         Top             =   855
         Width           =   4650
      End
      Begin VB.Frame fra_Contribuinte 
         Height          =   5625
         Left            =   90
         TabIndex        =   105
         Top             =   315
         Width           =   10050
         Begin VB.TextBox txtstrSubDivisaoFiscal 
            Height          =   285
            Left            =   4635
            MaxLength       =   3
            TabIndex        =   8
            Top             =   1260
            Width           =   465
         End
         Begin VB.TextBox txtdtmdtcancelamento 
            Height          =   285
            Left            =   2205
            MaxLength       =   20
            TabIndex        =   7
            Top             =   1260
            Width           =   975
         End
         Begin VB.CommandButton cmd_Logradouro 
            Height          =   300
            Left            =   5925
            Picture         =   "frmCadImobiliario.frx":199A
            Style           =   1  'Graphical
            TabIndex        =   214
            TabStop         =   0   'False
            Tag             =   "584"
            ToolTipText     =   "Ativa Cadastro de Logradouro"
            Top             =   1905
            Width           =   330
         End
         Begin VB.TextBox txt_intUF 
            Height          =   285
            Left            =   4545
            MaxLength       =   9
            TabIndex        =   16
            Top             =   2280
            Width           =   405
         End
         Begin VB.TextBox txt_intBairro 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   15
            Top             =   2280
            Width           =   2370
         End
         Begin VB.TextBox txtintfolha 
            Height          =   285
            Left            =   8340
            MaxLength       =   6
            TabIndex        =   39
            Top             =   5190
            Width           =   960
         End
         Begin VB.TextBox txtintlivro 
            Height          =   285
            Left            =   6645
            MaxLength       =   6
            TabIndex        =   38
            Top             =   5190
            Width           =   990
         End
         Begin VB.TextBox txtstrCartorio 
            Height          =   285
            Left            =   5520
            MaxLength       =   40
            TabIndex        =   35
            Top             =   4830
            Width           =   3780
         End
         Begin VB.TextBox txtdtmdtescritura 
            Height          =   285
            Left            =   5040
            MaxLength       =   10
            TabIndex        =   37
            Top             =   5190
            Width           =   1020
         End
         Begin VB.TextBox txtdtmdtmatricula 
            Height          =   285
            Left            =   2190
            MaxLength       =   10
            TabIndex        =   36
            Top             =   5190
            Width           =   1020
         End
         Begin VB.CommandButton cmd_isencao 
            Height          =   300
            Left            =   9000
            Picture         =   "frmCadImobiliario.frx":1D24
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Isenção e Imunidade"
            Top             =   1260
            Width           =   330
         End
         Begin VB.CheckBox chk_intIsencao 
            Caption         =   "Isenção / Imunidade"
            Height          =   195
            Left            =   7170
            TabIndex        =   10
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtintCepC 
            Height          =   285
            Left            =   8220
            MaxLength       =   9
            TabIndex        =   32
            Top             =   3975
            Width           =   1080
         End
         Begin VB.TextBox txtstrComplementoC 
            Height          =   285
            Left            =   3435
            MaxLength       =   20
            TabIndex        =   28
            Top             =   3660
            Width           =   1260
         End
         Begin VB.TextBox txtintNumeroC 
            Height          =   285
            Left            =   1680
            MaxLength       =   8
            TabIndex        =   27
            Top             =   3660
            Width           =   855
         End
         Begin VB.TextBox txtstrBairroC 
            Height          =   285
            Left            =   6225
            MaxLength       =   50
            TabIndex        =   29
            Top             =   3645
            Width           =   3075
         End
         Begin VB.CommandButton cmd_MunicipioC 
            Height          =   300
            Left            =   5265
            Picture         =   "frmCadImobiliario.frx":200A
            Style           =   1  'Graphical
            TabIndex        =   135
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Municipio"
            Top             =   3975
            Width           =   330
         End
         Begin VB.TextBox txtstrDistritoC 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   33
            Top             =   4320
            Width           =   3525
         End
         Begin VB.CommandButton cmd_TituloLogradouro 
            Height          =   300
            Left            =   4305
            Picture         =   "frmCadImobiliario.frx":22F0
            Style           =   1  'Graphical
            TabIndex        =   130
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Título de Logradouro"
            Top             =   3315
            Width           =   330
         End
         Begin VB.CommandButton cmd_TipoLogradouro 
            Height          =   300
            Left            =   2415
            Picture         =   "frmCadImobiliario.frx":25D6
            Style           =   1  'Graphical
            TabIndex        =   129
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Tipo deLogradouro"
            Top             =   3315
            Width           =   330
         End
         Begin VB.TextBox txtintCodigoLogradouro 
            Height          =   315
            Left            =   4695
            MaxLength       =   8
            TabIndex        =   25
            Top             =   3315
            Width           =   735
         End
         Begin VB.TextBox txtdblArea 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8190
            MaxLength       =   20
            TabIndex        =   22
            Top             =   2625
            Width           =   1125
         End
         Begin VB.TextBox txtintNumero 
            Height          =   285
            Left            =   6615
            MaxLength       =   10
            TabIndex        =   13
            Top             =   1920
            Width           =   825
         End
         Begin VB.TextBox txtintCep 
            Height          =   285
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   17
            Top             =   2280
            Width           =   1005
         End
         Begin VB.TextBox txtstrComplemento 
            Height          =   285
            Left            =   8040
            MaxLength       =   30
            TabIndex        =   14
            Top             =   1920
            Width           =   1275
         End
         Begin VB.TextBox txtstrEmissao 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   8895
            MaxLength       =   3
            TabIndex        =   2
            Top             =   180
            Width           =   420
         End
         Begin VB.TextBox txtstrSequenciaDeFace 
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
            Left            =   8760
            MaxLength       =   2
            TabIndex        =   18
            Top             =   2265
            Width           =   555
         End
         Begin VB.CommandButton cmd_intLoteamento 
            Height          =   315
            Left            =   3105
            Picture         =   "frmCadImobiliario.frx":28BC
            Style           =   1  'Graphical
            TabIndex        =   123
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Loteamento"
            Top             =   2625
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintLoteamento 
            Height          =   315
            Left            =   1680
            TabIndex        =   19
            Top             =   2625
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.CommandButton cmd_intContribuinte 
            Height          =   315
            Left            =   6630
            Picture         =   "frmCadImobiliario.frx":2C46
            Style           =   1  'Graphical
            TabIndex        =   111
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro Único"
            Top             =   900
            Width           =   360
         End
         Begin VB.CheckBox chkbytEdificado 
            Caption         =   "Edificado"
            Height          =   195
            Left            =   6060
            TabIndex        =   9
            Top             =   1320
            Width           =   1005
         End
         Begin VB.TextBox txt_strCNPJCPF 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   7950
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   915
            Width           =   1365
         End
         Begin VB.TextBox txt_PKIdContribuinte 
            BackColor       =   &H80000016&
            Height          =   315
            Left            =   2205
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   900
            Width           =   885
         End
         Begin VB.TextBox txtstrmatricula 
            Height          =   285
            Left            =   1695
            MaxLength       =   15
            TabIndex        =   34
            Top             =   4815
            Width           =   2130
         End
         Begin VB.TextBox txtstrLote 
            Height          =   285
            Left            =   3930
            MaxLength       =   20
            TabIndex        =   20
            Top             =   2625
            Width           =   1020
         End
         Begin VB.TextBox txtstrQuadra 
            Height          =   285
            Left            =   5760
            MaxLength       =   20
            TabIndex        =   21
            Top             =   2625
            Width           =   1155
         End
         Begin MSDataListLib.DataCombo dbcintContribuinte 
            Height          =   315
            Left            =   3090
            TabIndex        =   5
            Top             =   900
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskstrInscricaoAnterior 
            Height          =   300
            Left            =   5670
            TabIndex        =   1
            Top             =   180
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo dbcintLogradouro 
            Height          =   315
            Left            =   1680
            TabIndex        =   12
            Top             =   1920
            Width           =   4230
            _ExtentX        =   7461
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintMunicipioC 
            Height          =   315
            Left            =   1680
            TabIndex        =   30
            Top             =   3975
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintUFC 
            Height          =   315
            Left            =   6240
            TabIndex        =   31
            Top             =   3975
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTipoLogradouro 
            Height          =   315
            Left            =   1680
            TabIndex        =   23
            Top             =   3315
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTituloLogradouro 
            Height          =   315
            Left            =   2865
            TabIndex        =   24
            Top             =   3315
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcstrLogradouroC 
            Height          =   315
            Left            =   5475
            TabIndex        =   26
            Top             =   3315
            Width           =   3840
            _ExtentX        =   6773
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskstrInscricao 
            Height          =   300
            Left            =   2205
            TabIndex        =   0
            Top             =   180
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskstrInscricaoAuxiliar 
            Height          =   300
            Left            =   2205
            TabIndex        =   3
            Top             =   540
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   25
            PromptChar      =   " "
         End
         Begin VB.Label lbl_strSubDivisaoFiscal 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Divisão Fiscal"
            Height          =   195
            Left            =   3270
            TabIndex        =   222
            Top             =   1365
            Width           =   1305
         End
         Begin VB.Label lblstrInscricaoAuxiliar 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Auxiliar"
            Height          =   195
            Left            =   990
            TabIndex        =   109
            Top             =   630
            Width           =   1185
         End
         Begin VB.Label lbl_DtCancelamento 
            AutoSize        =   -1  'True
            Caption         =   "Cancelamento"
            Height          =   195
            Left            =   1110
            TabIndex        =   215
            Top             =   1365
            Width           =   1020
         End
         Begin VB.Label lblintfolha 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Folha"
            Height          =   195
            Left            =   7860
            TabIndex        =   144
            Top             =   5280
            Width           =   390
         End
         Begin VB.Label lblintlivro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Livro"
            Height          =   195
            Left            =   6240
            TabIndex        =   143
            Top             =   5280
            Width           =   345
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Cartório"
            Height          =   195
            Left            =   4920
            TabIndex        =   140
            Top             =   4905
            Width           =   540
         End
         Begin VB.Label lbldtmdtescritura 
            AutoSize        =   -1  'True
            Caption         =   "Data de Escritura"
            Height          =   195
            Left            =   3750
            TabIndex        =   142
            Top             =   5295
            Width           =   1230
         End
         Begin VB.Label lbldtmdtmatricula 
            AutoSize        =   -1  'True
            Caption         =   "Data de Matrícula"
            Height          =   195
            Left            =   825
            TabIndex        =   141
            Top             =   5280
            Width           =   1290
         End
         Begin VB.Label lblstrLoteamento 
            AutoSize        =   -1  'True
            Caption         =   "Loteamento"
            Height          =   195
            Left            =   780
            TabIndex        =   122
            Top             =   2730
            Width           =   840
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   810
            TabIndex        =   128
            Top             =   3390
            Width           =   810
         End
         Begin VB.Line lne_5 
            BorderColor     =   &H8000000B&
            DrawMode        =   14  'Copy Pen
            X1              =   780
            X2              =   9285
            Y1              =   4725
            Y2              =   4725
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   7875
            TabIndex        =   137
            Top             =   4065
            Width           =   285
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   5895
            TabIndex        =   136
            Top             =   4095
            Width           =   210
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   2835
            TabIndex        =   132
            Top             =   3735
            Width           =   480
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   1440
            TabIndex        =   131
            Top             =   3735
            Width           =   180
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   5670
            TabIndex        =   133
            Top             =   3735
            Width           =   405
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   915
            TabIndex        =   134
            Top             =   4095
            Width           =   705
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   1140
            TabIndex        =   138
            Top             =   4410
            Width           =   480
         End
         Begin VB.Line lne_4 
            BorderColor     =   &H8000000B&
            DrawMode        =   14  'Copy Pen
            X1              =   3030
            X2              =   9300
            Y1              =   3165
            Y2              =   3165
         End
         Begin VB.Line lne_3 
            BorderColor     =   &H8000000B&
            DrawMode        =   14  'Copy Pen
            X1              =   765
            X2              =   1110
            Y1              =   3165
            Y2              =   3165
         End
         Begin VB.Label lblEnderecoDeNotificacao 
            AutoSize        =   -1  'True
            Caption         =   "Endereço de Notificação"
            Height          =   195
            Left            =   1155
            TabIndex        =   127
            Top             =   3030
            Width           =   1770
         End
         Begin VB.Label lbl_UF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   4275
            TabIndex        =   119
            Top             =   2355
            Width           =   210
         End
         Begin VB.Label lblstrArea 
            AutoSize        =   -1  'True
            Caption         =   "Área do Terreno"
            Height          =   195
            Left            =   7005
            TabIndex        =   126
            Top             =   2730
            Width           =   1155
         End
         Begin VB.Label lblintCep 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Left            =   5370
            TabIndex        =   120
            Top             =   2370
            Width           =   315
         End
         Begin VB.Label lblstrComplemento 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   7530
            TabIndex        =   117
            Top             =   1995
            Width           =   480
         End
         Begin VB.Label lblintNumero 
            AutoSize        =   -1  'True
            Caption         =   "N°"
            Height          =   195
            Left            =   6360
            TabIndex        =   116
            Top             =   1995
            Width           =   180
         End
         Begin VB.Label lblintBairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   1215
            TabIndex        =   118
            Top             =   2370
            Width           =   405
         End
         Begin VB.Label lblintLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   810
            TabIndex        =   114
            Top             =   1995
            Width           =   810
         End
         Begin VB.Line lne_2 
            BorderColor     =   &H8000000B&
            DrawMode        =   14  'Copy Pen
            X1              =   2880
            X2              =   9315
            Y1              =   1770
            Y2              =   1770
         End
         Begin VB.Line lne_1 
            BorderColor     =   &H8000000B&
            DrawMode        =   14  'Copy Pen
            X1              =   765
            X2              =   1095
            Y1              =   1770
            Y2              =   1770
         End
         Begin VB.Label lblEnderecoDoImobiliario 
            AutoSize        =   -1  'True
            Caption         =   "Endereço do Imobiliário"
            Height          =   195
            Left            =   1155
            TabIndex        =   113
            Top             =   1650
            Width           =   1650
         End
         Begin VB.Label lblstrInscricao 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   810
            TabIndex        =   106
            Top             =   285
            Width           =   1350
         End
         Begin VB.Label lbl_Emissao 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Index           =   0
            Left            =   8235
            TabIndex        =   108
            Top             =   270
            Width           =   585
         End
         Begin VB.Label lblSequenciaDeFase 
            AutoSize        =   -1  'True
            Caption         =   "Sequência de face"
            Height          =   195
            Left            =   7350
            TabIndex        =   121
            Top             =   2370
            Width           =   1350
         End
         Begin VB.Label lblstrInscricaoAnterior 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Anterior"
            Height          =   195
            Left            =   4395
            TabIndex        =   107
            Top             =   270
            Width           =   1230
         End
         Begin VB.Label lblstrCNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   7110
            TabIndex        =   112
            Top             =   1005
            Width           =   780
         End
         Begin VB.Label lblstrLote 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            Height          =   195
            Left            =   3555
            TabIndex        =   124
            Top             =   2730
            Width           =   315
         End
         Begin VB.Label lblstrEscritura 
            AutoSize        =   -1  'True
            Caption         =   "Matrícula"
            Height          =   195
            Left            =   810
            TabIndex        =   139
            Top             =   4890
            Width           =   675
         End
         Begin VB.Label lblintContribunte 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código/Proprietário"
            Height          =   195
            Left            =   780
            TabIndex        =   110
            Top             =   1005
            Width           =   1365
         End
         Begin VB.Label lblstrQuadra 
            AutoSize        =   -1  'True
            Caption         =   "Quadra"
            Height          =   195
            Left            =   5205
            TabIndex        =   125
            Top             =   2730
            Width           =   525
         End
      End
      Begin VB.Frame fra_Promissario 
         Height          =   2640
         Left            =   -73905
         TabIndex        =   147
         Top             =   1215
         Width           =   8085
         Begin VB.Frame fra_Endereco 
            Caption         =   "Endereço de Correspondência"
            Height          =   1395
            Left            =   120
            TabIndex        =   151
            Top             =   1110
            Width           =   7845
            Begin VB.TextBox txt_UF 
               Height          =   285
               Left            =   5160
               MaxLength       =   2
               TabIndex        =   52
               Top             =   960
               Width           =   810
            End
            Begin VB.TextBox txt_Municipio 
               Height          =   285
               Left            =   5160
               MaxLength       =   50
               TabIndex        =   50
               Top             =   600
               Width           =   2595
            End
            Begin VB.TextBox txt_Cep 
               Height          =   285
               Left            =   6675
               MaxLength       =   9
               TabIndex        =   53
               Top             =   960
               Width           =   1080
            End
            Begin VB.TextBox txt_Complemento 
               Height          =   285
               Left            =   6750
               MaxLength       =   20
               TabIndex        =   115
               Top             =   240
               Width           =   990
            End
            Begin VB.TextBox txt_Numero 
               Height          =   285
               Left            =   5160
               MaxLength       =   8
               TabIndex        =   47
               Top             =   240
               Width           =   795
            End
            Begin VB.TextBox txt_Logradouro 
               Height          =   285
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   46
               Top             =   240
               Width           =   3225
            End
            Begin VB.TextBox txt_Bairro 
               Height          =   285
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   49
               Top             =   600
               Width           =   3225
            End
            Begin VB.TextBox txt_Distrito 
               Height          =   285
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   51
               Top             =   960
               Width           =   3225
            End
            Begin VB.Label lblintCepC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   6330
               TabIndex        =   159
               Top             =   1050
               Width           =   285
            End
            Begin VB.Label lblintUFC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   4860
               TabIndex        =   158
               Top             =   1035
               Width           =   210
            End
            Begin VB.Label lblstrComplementoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6240
               TabIndex        =   154
               Top             =   330
               Width           =   480
            End
            Begin VB.Label lblintNumeroC 
               AutoSize        =   -1  'True
               Caption         =   "Nº"
               Height          =   195
               Left            =   4890
               TabIndex        =   153
               Top             =   330
               Width           =   180
            End
            Begin VB.Label lblintLogradouroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   210
               TabIndex        =   152
               Top             =   330
               Width           =   810
            End
            Begin VB.Label lblintBairroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   615
               TabIndex        =   155
               Top             =   690
               Width           =   405
            End
            Begin VB.Label lblintMunicipioC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   4380
               TabIndex        =   156
               Top             =   690
               Width           =   705
            End
            Begin VB.Label lblstrDistritoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               Height          =   195
               Left            =   540
               TabIndex        =   157
               Top             =   1050
               Width           =   480
            End
         End
         Begin VB.CommandButton cmd_Contribuinte 
            Height          =   315
            Left            =   6075
            Picture         =   "frmCadImobiliario.frx":2FD0
            Style           =   1  'Graphical
            TabIndex        =   149
            TabStop         =   0   'False
            ToolTipText     =   "Ativa cadastro de Promissários"
            Top             =   300
            Width           =   360
         End
         Begin VB.OptionButton optbytNaturezaJuridica 
            BackColor       =   &H80000004&
            Caption         =   "Pessoa Jurídica"
            Height          =   240
            Index           =   1
            Left            =   4590
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   690
            Width           =   1500
         End
         Begin VB.OptionButton optbytNaturezaJuridica 
            BackColor       =   &H80000004&
            Caption         =   "Pessoa Física"
            Height          =   315
            Index           =   0
            Left            =   3225
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   660
            Width           =   1365
         End
         Begin VB.TextBox txt_strCNPJCPFP 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1215
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   660
            Width           =   1845
         End
         Begin MSDataListLib.DataCombo dbcintPromissario 
            Height          =   315
            Left            =   1215
            TabIndex        =   42
            Top             =   300
            Width           =   4860
            _ExtentX        =   8573
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin VB.Label lblstrCNPJCPFP 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   330
            TabIndex        =   150
            Top             =   735
            Width           =   780
         End
         Begin VB.Label lblintPromissario 
            AutoSize        =   -1  'True
            Caption         =   "Promissário"
            Height          =   195
            Left            =   345
            TabIndex        =   148
            Top             =   375
            Width           =   795
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   1590
         Left            =   135
         TabIndex        =   209
         Top             =   6000
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   2805
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
         Columns(1).Caption=   "Inscrição Cadastral"
         Columns(1).DataField=   "strInscricao"
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Contribuinte"
         Columns(2).DataField=   "strNome"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "CNPJ / CPF"
         Columns(3).DataField=   "strCNPJCPF"
         Columns(3).NumberFormat=   "FormatText Event"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "intSecoes"
         Columns(4).DataField=   "intSecoes"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2672"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2593"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=8573"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=8493"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=3651"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3572"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(29)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(30)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
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
      Begin MSComctlLib.ListView lvw_Caracteristica 
         Height          =   1815
         Index           =   2
         Left            =   -74295
         TabIndex        =   80
         Top             =   3270
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Característica"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Detalhe 
         Height          =   1815
         Index           =   2
         Left            =   -69870
         TabIndex        =   81
         Top             =   3270
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Detalhes"
            Object.Width           =   6368
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Caracteristica 
         Height          =   4575
         Index           =   0
         Left            =   -74040
         TabIndex        =   71
         Top             =   1425
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   8070
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Característica"
            Object.Width           =   6368
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Detalhe 
         Height          =   4575
         Index           =   0
         Left            =   -69930
         TabIndex        =   72
         Top             =   1425
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Detalhes"
            Object.Width           =   6368
         EndProperty
      End
      Begin VB.Frame fra_Envolvido 
         Height          =   2625
         Left            =   -74205
         TabIndex        =   162
         Top             =   1275
         Width           =   8655
         Begin VB.Frame fra_EnderecoEnv 
            Caption         =   "Endereço de Correspondência"
            Height          =   1395
            Left            =   105
            TabIndex        =   166
            Top             =   1125
            Width           =   8370
            Begin VB.TextBox txt_DistritoEnv 
               Height          =   285
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   66
               Top             =   960
               Width           =   3225
            End
            Begin VB.TextBox txt_BairroEnv 
               Height          =   285
               Left            =   1080
               MaxLength       =   50
               TabIndex        =   64
               Top             =   600
               Width           =   3225
            End
            Begin VB.TextBox txt_LogradouroEnv 
               Height          =   285
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   61
               Top             =   240
               Width           =   3225
            End
            Begin VB.TextBox txt_NumeroEnv 
               Height          =   285
               Left            =   5130
               MaxLength       =   8
               TabIndex        =   62
               Top             =   240
               Width           =   795
            End
            Begin VB.TextBox txt_ComplementoEnv 
               Height          =   285
               Left            =   6750
               MaxLength       =   20
               TabIndex        =   63
               Top             =   240
               Width           =   990
            End
            Begin VB.TextBox txt_CepEnv 
               Height          =   285
               Left            =   6675
               MaxLength       =   9
               TabIndex        =   68
               Top             =   960
               Width           =   1080
            End
            Begin VB.TextBox txt_MunicipioEnv 
               Height          =   285
               Left            =   5130
               MaxLength       =   50
               TabIndex        =   65
               Top             =   600
               Width           =   2625
            End
            Begin VB.TextBox txt_UFEnv 
               Height          =   285
               Left            =   5130
               MaxLength       =   2
               TabIndex        =   67
               Top             =   960
               Width           =   810
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               Height          =   195
               Left            =   540
               TabIndex        =   172
               Top             =   1050
               Width           =   480
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   4380
               TabIndex        =   171
               Top             =   690
               Width           =   705
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   615
               TabIndex        =   170
               Top             =   690
               Width           =   405
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   210
               TabIndex        =   167
               Top             =   330
               Width           =   810
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Nº"
               Height          =   195
               Left            =   4890
               TabIndex        =   168
               Top             =   330
               Width           =   180
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6240
               TabIndex        =   169
               Top             =   330
               Width           =   480
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   4860
               TabIndex        =   173
               Top             =   1035
               Width           =   210
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   6330
               TabIndex        =   174
               Top             =   1050
               Width           =   285
            End
         End
         Begin VB.OptionButton opt_Proprietario 
            Caption         =   "Promissário"
            Height          =   195
            Index           =   1
            Left            =   4575
            TabIndex        =   60
            Top             =   705
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt_Proprietario 
            Caption         =   "Proprietário"
            Height          =   195
            Index           =   0
            Left            =   3300
            TabIndex        =   59
            Top             =   705
            Width           =   1215
         End
         Begin VB.TextBox txt_PKIdContribuinte2 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   225
            Width           =   885
         End
         Begin VB.TextBox txt_strCNPJCPFEnv 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   675
            Width           =   1365
         End
         Begin VB.CommandButton cmd_intContribuinte2 
            Height          =   315
            Left            =   6660
            Picture         =   "frmCadImobiliario.frx":335A
            Style           =   1  'Graphical
            TabIndex        =   164
            TabStop         =   0   'False
            ToolTipText     =   "Ativa  Cadastro Único"
            Top             =   225
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbc_intContribuinte 
            Height          =   315
            Left            =   2505
            TabIndex        =   57
            Top             =   225
            Width           =   4170
            _ExtentX        =   7355
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin VB.Label lbl_Envolvido 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código/Envolvido"
            Height          =   195
            Left            =   255
            TabIndex        =   163
            Top             =   330
            Width           =   1275
         End
         Begin VB.Label lbl_CPFCNPJ 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   255
            TabIndex        =   165
            Top             =   750
            Width           =   780
         End
      End
      Begin MSComctlLib.ListView lvw_Envolvidos 
         Height          =   2235
         Left            =   -74220
         TabIndex        =   175
         Top             =   4065
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Contribuinte"
            Object.Width           =   9878
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CPF/CNPJ"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Vinculo"
            Object.Width           =   2293
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Caracteristica 
         Height          =   1995
         Index           =   1
         Left            =   -74295
         TabIndex        =   74
         Top             =   3015
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   3519
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Característica"
            Object.Width           =   7497
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Detalhe 
         Height          =   1995
         Index           =   1
         Left            =   -69795
         TabIndex        =   75
         Top             =   3015
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   3519
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Detalhes"
            Object.Width           =   7629
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbc_intResumoTipoPadrao 
         Height          =   315
         Left            =   -72150
         TabIndex        =   216
         Top             =   2880
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lbl_CaminhoFicha 
         Height          =   345
         Left            =   -74790
         TabIndex        =   223
         Top             =   6960
         Width           =   8955
      End
      Begin VB.Label lbl_TotPredial 
         AutoSize        =   -1  'True
         Caption         =   "Total de Área Predial"
         Height          =   195
         Left            =   -68130
         TabIndex        =   221
         Top             =   525
         Width           =   1485
      End
      Begin VB.Label lbl_strArea 
         AutoSize        =   -1  'True
         Caption         =   "Área do Terreno"
         Height          =   195
         Left            =   -67770
         TabIndex        =   219
         Top             =   525
         Width           =   1155
      End
      Begin VB.Label lbl_Resumo 
         AutoSize        =   -1  'True
         Caption         =   "Características para Resumo"
         Height          =   195
         Left            =   -74250
         TabIndex        =   217
         Top             =   2970
         Width           =   2055
      End
      Begin VB.Label lbl_InscricaoCadastral8 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Index           =   1
         Left            =   -74085
         TabIndex        =   201
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label lbl_Proprietario8 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Index           =   1
         Left            =   -73530
         TabIndex        =   202
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -73470
         TabIndex        =   161
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74025
         TabIndex        =   160
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label lbl_Proprietario2 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -73530
         TabIndex        =   177
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lbl_InscricaoCadastral3 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74085
         TabIndex        =   176
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label lbl_Proprietario3 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -73470
         TabIndex        =   180
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lbl_InscricaoCadastral4 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74025
         TabIndex        =   178
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label lbl_Proprietario4 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -73470
         TabIndex        =   186
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lbl_InscricaoCadastral5 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74025
         TabIndex        =   185
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label lbl_Proprietario5 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -73470
         TabIndex        =   191
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lbl_InscricaoCadastral6 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74025
         TabIndex        =   190
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label lbl_Proprietario6 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Index           =   0
         Left            =   -73440
         TabIndex        =   196
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lbl_InscricaoCadastral7 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Index           =   0
         Left            =   -73995
         TabIndex        =   195
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label lbl_InscricaoCadastral1 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   -74055
         TabIndex        =   145
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label lbl_Proprietario7 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário"
         Height          =   195
         Left            =   -73500
         TabIndex        =   146
         Top             =   960
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCadImobiliario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando      As Boolean
Dim mblnAlterandoList  As Boolean
Dim mblnPrimeiraVez    As Boolean
Dim mcboAux            As ComboBox
Dim adoResultado       As ADODB.Recordset
Dim strsql             As String
Dim mblnSelecionou     As Boolean
Dim intMaxPKId         As Long
Dim oList              As Object
Dim objList            As Object
Dim objList1           As Object
Dim objListApoio       As Object
Dim adoRec             As ADODB.Recordset
Dim adoTdb             As ADODB.Recordset
Dim x                  As XArrayDB
Dim Y                  As XArrayDB
Dim Z                  As XArrayDB
Dim A                  As XArrayDB
Dim j                  As XArrayDB
Dim aFichasCadastrais  As XArrayDB
Dim intIndiceFichas    As Integer
Dim intCodImobiliario  As Integer
Dim dblEdificacao      As Double
Dim mobjAux            As Object
Dim dblTerreno         As Double
Dim PKId_Temporario    As Double
Dim strEnvolvidos()    As String

'Guarda o tipo do imóvel de acordo com o tab selecionado
'1 = Imobiliário Geral
'2 = Imobiliário Terreno
'3 = Imobiliário Construção

Dim intCaractImobil    As Integer
Dim intCodContribuinte As Integer
Dim tdbGridSelecionada As TrueOleDBGrid70.TDBGrid

Dim mblnClick          As Boolean
Dim strInscricaoA      As String
Dim intRow             As Integer
Dim dblValorProfundidade   As Double
Dim blnGleba           As Boolean
Dim intCategoriaConstrucao  As Long
Dim blnGridCategoria    As Boolean
Dim blnCarregando       As Boolean
    
Dim blnEmTransacao      As Boolean 'Indica quando ja esta com BeginTransaction - para o caso de gravar as caracteristicas
Dim bytbaselimpa        As Byte 'Indica se a base esta limpa (banco vazio)
Dim bytTamanhoMascara   As Byte 'Armazena o tamanho de caracteres da mascara

Private Sub cmd_FichaAnterior_Click()
    intIndiceFichas = intIndiceFichas - 1
    MovimentaFichas intIndiceFichas
End Sub

Private Sub cmd_FichaPosterior_Click()
    intIndiceFichas = intIndiceFichas + 1
    MovimentaFichas intIndiceFichas
End Sub

Private Sub cmd_intContribuinte_Click()
    ChamaFormCadastro frmCadContribuinte, dbcintContribuinte
End Sub

Private Sub cmd_intContribuinte2_Click()
    ChamaFormCadastro frmCadContribuinte, dbc_intContribuinte
End Sub

Private Sub cmd_intLoteamento_Click()
    ChamaFormCadastro frmCadLoteamentos, dbcintLoteamento
End Sub

Private Sub cmd_isencao_Click()
    frmCadIsencaoImunidade.strFormulario = Me.Name
    frmCadIsencaoImunidade.intPkid = tdb_Lista.Columns("pkid").Value
    PreencherListaDeOpcoes frmCadIsencaoImunidade.dbcintIdentificacao, frmCadIsencaoImunidade.intPkid
    frmCadIsencaoImunidade.mblnPrimeiraVez = True
    frmCadIsencaoImunidade.MantemForm gstrLocalizar
    ChamaFormCadastro frmCadIsencaoImunidade, cmd_isencao
End Sub

Private Sub cmd_Logradouro_Click()
    CarregaForm frmCadLogradouro, dbcintLogradouro
End Sub

Private Sub cmd_MunicipioC_Click()
    ChamaFormCadastro frmCadCidade, dbcintMunicipioC
End Sub

Private Sub cmd_TipoDocumentoProcesso_Click()
    ChamaFormCadastro frmCadDocumentos, dbc_intTiposDocumentosProcesso
End Sub

Private Sub cmd_TipoLogradouro_Click()
    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouro
End Sub

Private Sub cmd_TituloLogradouro_Click()
    CarregaForm frmCadTituloLogradouro, dbcintTituloLogradouro
End Sub

Private Sub dbc_intContribuinte_Click(Area As Integer)
Dim strsql As String
Dim adoResultado As ADODB.Recordset
   
'   DropDownDataCombo dbc_intContribuinte, Me, Area
   
'   PreencherListaDeOpcoes dbc_intContribuinte
      
   If Area = 2 Then
       
       If dbc_intContribuinte.Locked Then Exit Sub
       
       If dbc_intContribuinte.BoundText <> "" Then
           strsql = ""
            strsql = strsql & "SELECT strCodigoAnterior, PKId, " & gstrISNULL("strCNPJCPF", "0") & " AS strCNPJCPF FROM " & gstrContribuinte & " "
            strsql = strsql & "WHERE PKId = " & dbc_intContribuinte.BoundText
            Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
                    txt_strCNPJCPFEnv = gstrCGCCPFFormatado(adoResultado!StrCnpjCpf)
                    txt_PKIdContribuinte2 = gstrVerificaCampoNulo(adoResultado!strCodigoAnterior, False)
                    MostraDadosEnvolvidos CLng(dbc_intContribuinte.BoundText)
                End If
        End If
   
   End If
   
End Sub

Private Sub dbc_intContribuinte_GotFocus()
    tab_3DPasta.Tab = 2
End Sub

Private Sub dbc_intResumoTipoPadrao_Change()
    If dbc_intResumoTipoPadrao.MatchedWithList = True And Val(dbc_intResumoTipoPadrao.BoundText) > 0 Then
        dbc_intResumoTipoPadrao.Enabled = False
        GravaCaracteristicaResumo
        dbc_intResumoTipoPadrao.Enabled = True
        If dbc_intResumoTipoPadrao.Enabled Then dbc_intResumoTipoPadrao.SetFocus
    End If
End Sub

Private Sub dbc_intTiposDocumentosProcesso_Click(Area As Integer)
    DropDownDataCombo dbc_intTiposDocumentosProcesso, Me, Area
End Sub

Private Sub dbc_intTiposDocumentosProcesso_GotFocus()
    MarcaCampo dbc_intTiposDocumentosProcesso
End Sub

Private Sub dbc_intTiposDocumentosProcesso_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intTiposDocumentosProcesso, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intTiposDocumentosProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intTiposDocumentosProcesso
End Sub

Private Sub dbcintContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinte_LostFocus()
    strPreencheContribuinteTabs
End Sub

Private Sub dbcintLogradouro_Change()
    If dbcintLogradouro.MatchedWithList Then
        LogradouroCep Val(dbcintLogradouro.BoundText), txt_intBairro, False, , txt_intUF, txtintCep, , False
    End If
End Sub

Private Sub dbcintLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLoteamento_Click(Area As Integer)
    DropDownDataCombo dbcintLoteamento, Me, Area
End Sub

Private Sub dbcintLoteamento_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLoteamento, Me, , KeyCode, Shift
End Sub

Private Sub dbcintMunicipioC_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintMunicipioC, Me, Area
End Sub

Private Sub dbcintMunicipioC_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintMunicipioC, Me, , KeyCode, Shift
End Sub

Private Sub dbcintMunicipioC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintPromissario_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintPromissario, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intFaceDeQuadra_Click(Area As Integer)
    DropDownDataCombo dbc_intFaceDeQuadra, Me, Area
    If Area = 2 And dbc_intFaceDeQuadra.MatchedWithList = True Then
        CarregaMelhoria
    End If
End Sub

Private Sub dbc_intFaceDeQuadra_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intFaceDeQuadra, Me, , KeyCode, Shift
End Sub

Private Sub chkbitEnglobado_Click()
    If chkbitEnglobado = 1 Then
        mskstrInscricaoEnglobada.Enabled = True
        TrocaCorObjeto mskstrInscricaoEnglobada, False
    Else
        mskstrInscricaoEnglobada.Text = ""
        mskstrInscricaoEnglobada.Enabled = False
        TrocaCorObjeto mskstrInscricaoEnglobada, True
    End If
End Sub

Private Sub chkbytEdificado_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkbytEdificado
End Sub
Private Sub cmd_Contribuinte_Click()
    ChamaFormCadastro frmCadContribuinte, dbcintPromissario
End Sub

Private Sub dbcintLogradouro_Click(Area As Integer)
On Error Resume Next
    DropDownDataCombo dbcintLogradouro, Me, Area
End Sub

Private Sub dbcintLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintLogradouro
End Sub

Private Sub dbcintPromissario_GotFocus()
    tab_3DPasta.Tab = 1
End Sub

Private Sub dbcintTipoLogradouro_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintTipoLogradouro, Me, Area
End Sub

Private Sub dbcintTipoLogradouro_GotFocus()
    MarcaCampo dbcintTipoLogradouro
End Sub

Private Sub dbcintTipoLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTipoLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTipoLogradouro
End Sub

Private Sub dbcintTituloLogradouro_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintTituloLogradouro, Me, Area
End Sub

Private Sub dbcintTituloLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTituloLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTituloLogradouro
End Sub

Private Sub dbcintUFC_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintUFC, Me, Area
End Sub

Private Sub dbcintUFC_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUFC, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUFC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcstrLogradouroC_GotFocus()
    MarcaCampo dbcstrLogradouroC
End Sub

Private Sub dbcstrLogradouroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcstrLogradouroC
End Sub

Private Sub Form_Activate()
    If MDIMenu.Tag = "Ouvidoria" Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir
    End If
    gintCodSeguranca = 735
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
    
    If App.ProductName = "Tributario" Then
        If Trim(txtdtmdtcancelamento) <> "" Then
            MDIMenu.actBarra.Bands(gstrBtnArquivo).Tools.Item(20).ToolTipText = "Reativar"
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCancelarReativar
        Else
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCancelarReativar
            MDIMenu.actBarra.Bands(gstrBtnArquivo).Tools.Item(20).ToolTipText = "Cancelar"
        End If
    End If
    
    tab_3dPasta_Click tab_3DPasta.Tab + 1
    
End Sub

Private Sub dbcintPromissario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If dbcintPromissario.MatchedWithList Then
            dbcintPromissario_Click 2
        Else
            dbcintPromissario.BoundText = ""
        End If
        CaracterValido KeyAscii, "A", dbcintPromissario
    End If
End Sub

Private Sub dbcintContribuinte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If dbcintContribuinte.MatchedWithList Then
            dbcintContribuinte_Click 2
        Else
            LimpaGrid
            dbcintContribuinte.BoundText = ""
        End If
        CaracterValido KeyAscii, "A", dbcintContribuinte
    End If
End Sub

Private Sub Form_Deactivate()

    If App.ProductName = "Tributario" Then
        MDIMenu.actBarra.Bands(gstrBtnArquivo).Tools.Item(20).ToolTipText = ""
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelarReativar
    End If
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar

End Sub

Private Sub Form_Load()
   
   VerificaObjParaAplicar mobjAux
   
   bytTamanhoMascara = 0
   'Nino - Desabilita o check de Edificado, conforme pedido em pendência.
   TrocaCorObjeto chkbytEdificado, True

   dbcintTipoLogradouro.Tag = gstrQueryDataComboTipoLogradouro & ";strSigla"
   dbcintTituloLogradouro.Tag = gstrQueryDataComboTituloLogradouro & ";strDescricao"
   dbcintMunicipioC.Tag = gstrQueryDataComboMunicipio & ";strDescricao"
   dbcintUFC.Tag = gstrQueryDataComboUF & ";strSigla"
   TrocaCorObjeto dbcintMunicipioC, True
   TrocaCorObjeto dbcintUFC, True

   tab_3DPasta.TabVisible(3) = False
   
   'TrocaCorObjeto dbcintLogradouro, True
   TrocaCorObjeto txt_intBairro, True
   'TrocaCorObjeto txtintCep, True
   TrocaCorObjeto txt_intUF, True
    
    If MDIMenu.Tag = "Ouvidoria" Then
        cmd_Contribuinte.Enabled = False
    End If
    
    mblnAlterando = False
    MontaColumnHeaders
    mskstrInscricaoEnglobada.Enabled = False
    TrocaCorObjeto mskstrInscricaoEnglobada, True
    
   'TAB GERAL
    dbcintLoteamento.Tag = strQuerryLoteamento & ";strNome"
'   dbcintComposicao.Tag = strQuerryComposicao & ";strDescricao"
'   dbcintOcorrrencia.Tag = strQueryTabelaOcorrencia & ";strDescricao"
    dbcintContribuinte.Tag = strQueryDataComboContribuinte & ";strNome"
    dbcintLogradouro.Tag = strQueryLogradouro(True) & ";L.strDescricao"
    dbcstrLogradouroC.Tag = gstrQueryLogradouro & ";L.strDescricao"
    
   'TAB ENVOLVIDOS
    dbc_intContribuinte.Tag = strQueryDataComboContribuinte & ";strNome"
'    dbc_intResumoTipoPadrao.Tag = strQueryResumoTipoPadrao & ";strCodigo"
    VerificaMascaraInscricao
    'LeDaTabelaParaObj gstrImobiliario, tdb_Lista, strQueryListView
    'MontaArray
   'TAB PROMISSARIO
    dbcintPromissario.Tag = strQueryDataComboContribuinte & ";strNome"
    PKId_Temporario = 0
'    tab_3DPasta.TabEnabled(6) = False
'    PreencheFaceDeQuadra
    txt_Bairro.Enabled = False
    TrocaCorObjeto txt_Bairro, True
    txt_Cep.Enabled = False
    TrocaCorObjeto txt_Cep, True
    txt_Complemento.Enabled = False
    TrocaCorObjeto txt_Complemento, True
    txt_Distrito.Enabled = False
    TrocaCorObjeto txt_Distrito, True
    txt_Logradouro.Enabled = False
    TrocaCorObjeto txt_Logradouro, True
    txt_Municipio.Enabled = False
    TrocaCorObjeto txt_Municipio, True
    txt_Numero.Enabled = False
    TrocaCorObjeto txt_Numero, True
    txt_UF.Enabled = False
    TrocaCorObjeto txt_UF, True
    
    txt_BairroEnv.Enabled = False
    TrocaCorObjeto txt_BairroEnv, True
    txt_CepEnv.Enabled = False
    TrocaCorObjeto txt_CepEnv, True
    txt_ComplementoEnv.Enabled = False
    TrocaCorObjeto txt_ComplementoEnv, True
    txt_DistritoEnv.Enabled = False
    TrocaCorObjeto txt_DistritoEnv, True
    txt_LogradouroEnv.Enabled = False
    TrocaCorObjeto txt_LogradouroEnv, True
    txt_MunicipioEnv.Enabled = False
    TrocaCorObjeto txt_MunicipioEnv, True
    txt_NumeroEnv.Enabled = False
    TrocaCorObjeto txt_NumeroEnv, True
    txt_UFEnv.Enabled = False
    TrocaCorObjeto txt_UFEnv, True
    TrocaCorObjeto txtdtmdtcancelamento, True
    
    PrrencheGRDCategoriaConstrucao
    
    PreencheGRD
    PreencheGRD2
                    
    MontaDropDownAreaCM
    'PreencheFaceDeQuadra
    PreencheTestada False

    'Tab Doc/Proc
    TrocaCorObjeto txt_DtmDataDocProc, True
    TrocaCorObjeto chk_intIsencao, True

End Sub

Private Function strQueryDataComboContribuinte()
Dim strsql As String
    strsql = ""
    strsql = strsql & "SELECT PKId, strNome "
    strsql = strsql & "FROM " & gstrContribuinte & " "
    strsql = strsql & "ORDER BY strNome"
    strQueryDataComboContribuinte = strsql
End Function

Private Function ContadorDetalhe() As Boolean
Dim strsql As String
        
    strsql = ""
    strsql = strsql & " SELECT COUNT(*) as Contador FROM "
    strsql = strsql & gstrCaracteristicaDoImovel
    strsql = strsql & " WHERE intCodigoImobiliario = " & Val(PKId_Temporario)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If adoResultado!Contador <> 0 Then
                ContadorDetalhe = True
                Exit Function
            End If
        End If
    End If
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      
    If MDIMenu.Tag = "Ouvidoria" Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
    End If
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaRollbackTrans
    
    If App.ProductName = "Tributario" Then
        MDIMenu.actBarra.Bands(gstrBtnArquivo).Tools.Item(20).ToolTipText = ""
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelarReativar
    End If
    
    If PKId_Temporario = 0 Then
        Exit Sub
    End If
    If ContadorDetalhe = True Then
        DeletaTemporario
        DoEvents
    Else
        gobjBanco.ExecutaBeginTrans
        gobjBanco.ExecutaCommitTrans
    End If
End Sub
 
Private Sub DeletaTemporario()
Dim strsql As String
        
    strsql = ""
    strsql = strsql & " DELETE "
    strsql = strsql & " FROM " & gstrCaracteristicaDoImovel
    strsql = strsql & " WHERE "
    strsql = strsql & " intCodigoImobiliario = " & Val(PKId_Temporario)
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql
    PKId_Temporario = 0

End Sub

Private Sub grd_Area_AfterColEdit(ByVal ColIndex As Integer)
    Select Case ColIndex
      Case 2
        If grd_Area.Columns(ColIndex).Text = "" Then
            grd_Area.Columns(ColIndex).Value = ""
        End If
      Case 3
        grd_Area.Columns("Medida da Área").Value = gstrConvVrDoSql(grd_Area.Columns("Medida da Área").Value, 2)
      Case 4
        grd_Area.Columns("Fração Ideal").Text = gstrConvVrDoSql(grd_Area.Columns("Fração Ideal").Text, 6)
    End Select
    
End Sub

Private Sub grd_Area_LostFocus()
    Set tdbGridSelecionada = Nothing
End Sub

Private Sub grd_Area_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With grd_Area
        If Not .EOF And Not .BOF Then
            If Len(.Columns("Pkid").Value) > 0 Then
                CarregaCaracteristica
                lvw_Caracteristica_ItemClick 2, lvw_Caracteristica(2).SelectedItem
                If Val(grd_Area.Columns("PkidCategoriaConstrucao").Value) > 0 Then
                    LeDaTabelaParaObj "", dbc_intResumoTipoPadrao, strQueryResumoTipoPadrao(grd_Area.Columns("PkidCategoriaConstrucao").Value)
                End If
                CarregaConstrucao
            Else
                lvw_Caracteristica(intCaractImobil).ListItems.Clear
                lvw_Detalhe(intCaractImobil).ListItems.Clear
                lvw_CaracPredio.ListItems.Clear
                lbl_DescricaoPredios = ""
            End If
            Set tdbGridSelecionada = grd_Area
        End If
    End With
    
End Sub

Private Sub grd_Testada_ButtonClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 1
            intRow = grd_Testada.Row
            If grd_Testada.Row = 0 Then
                PreencheTestada True
            Else
                PreencheTestada False
            End If
        Case 3
            If Trim(mskstrInscricao.Text) <> "" Then
                PreencheFaceDeQuadra
            End If
    End Select
 
End Sub

Private Sub grd_Testada_Click()
    Set tdbGridSelecionada = grd_Testada
End Sub

Private Sub grd_Testada_LostFocus()
    Set tdbGridSelecionada = Nothing
End Sub

Private Sub grd_Testada_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If grd_Testada.Row <> LastRow And LastRow <> 0 Then
        If x(LastRow, 6) = "" And x(LastRow, 2) = "" And _
            x(LastRow, 3) = "" Then
            x.DeleteRows (LastRow)
            Set grd_Testada.Array = x
            grd_Testada.ReBind
            grd_Testada.Refresh
        End If
    End If
    gCorLinhaSelecionada grd_Testada
End Sub

Private Sub lvw_Caracteristica_Click(Index As Integer)
    If intCaractImobil = 2 Then
        If Not Len(grd_Area.Columns("Pkid").Value) > 0 Then
           ExibeMensagem "É necessário selecionar uma Área."
           grd_Area.SetFocus
        End If
    End If
End Sub

Private Sub lvw_Caracteristica_GotFocus(Index As Integer)
    Select Case Index
        Case 0
            tab_3DPasta.Tab = 4
        Case 1
            tab_3DPasta.Tab = 5
        Case 2
            tab_3DPasta.Tab = 6
        Case 3
            tab_3DPasta.Tab = 7
    End Select
End Sub

Private Sub lvw_Caracteristica_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If tab_3DPasta.Tab = 6 And Not Len(grd_Area.Columns("Pkid").Value) > 0 Then Exit Sub
    If lvw_Caracteristica(Index).ListItems.Count > 0 Then
        CarregaDetalhes
        If mblnAlterando = True Then
            SelecionaDetalhe
        Else
            SelecionaDetalheApoio
        End If
    End If
    
End Sub

Private Sub lvw_Caracteristica_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_Caracteristica
End Sub

Private Sub lvw_Detalhe_Click(Index As Integer)
Dim i As Integer
    
    If intCaractImobil = 2 Then
        If Not Len(grd_Area.Columns("Pkid").Value) > 0 Then
            ExibeMensagem "É necessário selecionar uma Área."
            grd_Area.SetFocus
            For i = 1 To lvw_Detalhe(Index).ListItems.Count
                lvw_Detalhe(Index).ListItems(i).Checked = False
            Next
        End If
    ElseIf intCaractImobil = 1 Then
        If Trim(txtPKId.Text) <> "0" And Trim(txtPKId.Text) <> "" Then CarregaFatorTerreno
        
        For i = 1 To lvw_Detalhe(Index).ListItems.Count
            If lvw_Detalhe(Index).ListItems(i).Checked Then
                AtualizaFatorTerreno (CLng(lvw_Detalhe(Index).ListItems(i).Tag))
                Exit For
            End If
        Next
    End If
    
End Sub

Private Sub lvw_Detalhe_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        lvw_Detalhe_MouseDown Index, 0, 0, 0, 0
    End If
End Sub

Private Sub lvw_Detalhe_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_Detalhe
End Sub

Private Sub lvw_Detalhe_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        lvw_Detalhe_MouseUp Index, 0, 0, 0, 0
        lvw_Detalhe_Click Index
    End If
End Sub

Private Sub lvw_Detalhe_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer
    For i = 1 To lvw_Detalhe(Index).ListItems.Count
        lvw_Detalhe(Index).ListItems(i).Checked = False
    Next
End Sub

Private Sub lvw_Detalhe_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim intAux       As Long
Dim intEdificado As Integer

    If intCaractImobil = 2 And Not Len(grd_Area.Columns("Pkid").Value) > 0 Then Exit Sub
    If Trim(txtdtmdtcancelamento) <> "" Then Exit Sub
    
    'Caso nao tenha salvo o imobiliario vamos forcar a salva, para que possa ser salva as caracteristicas
    If Val(txtPKId) = 0 Then
        
        If blnDadosOK(gstrSalvar, True) Then
        
            intEdificado = chkbytEdificado.Value
            Set gobjBanco = New clsBanco
    
            gobjBanco.ExecutaBeginTrans
            
            If SalvarGeral(gstrImobiliario, IIf(mblnAlterando, "A", "I"), Me, tdb_Lista, strQueryListView(gstrSalvar), False, False, True) Then
            
            'If ToolBarGeral(gstrSalvar, gstrImobiliario, mblnAlterando, tdb_Lista, Me, mobjAux, strQueryListView(gstrSalvar), , , , False) Then   'Mudei
                
                If intAux = 0 Then
                    intAux = PegaMaxPKId
                End If

                GravaHistoricos intAux, False
                GravaValores intAux, intEdificado, False
                GravaValores2 intAux, False
                GravaEnvolvidos intAux, False
                
                txtPKId = intAux
                
                mblnPrimeiraVez = False
                mblnAlterando = True
                
                blnEmTransacao = True
                
            End If
            
            Screen.MousePointer = vbDefault
            
        End If
        
    End If
    
    If Val(txtPKId) <> 0 Then
        GravaDetalhe CLng(txtPKId)
        CarregaConstrucao
    End If
    
End Sub

Function SelecionaDetalheApoio()

'******************************************************************************************
' Data: 09/03/2003
' Alteração: - Alteração do nome do atributo intCodigoDetalheDaCaracteristica por
'            intCodigoDetalheDaCaracteristi a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/03/2003
' Alteração: - Alteração do nome do atributo intCodigoUtilizacaoDaTabelaDeValor por
'            intCodigoUtilizacaoDaTabelaDeV a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strsql As String
    strsql = ""
'    strSql = strSql & " SELECT D.intCodigoDetalheDaCaracteristica Detalhe "
    strsql = strsql & " SELECT D.intCodigoDetalheDaCaracteristi Detalhe "
    strsql = strsql & " FROM " & gstrImobiliario & " A," & gstrCaracteristicaGeral & " B,"
    strsql = strsql & gstrUtilizacaoDaTabelaDeValor & " C, " & gstrCaracteristicaDoImovel & " D"
    strsql = strsql & " WHERE D.intCodigoImobiliario = " & Val(PKId_Temporario)
    strsql = strsql & " AND D.intCodigoCaracteristicaGeral = B.PKId"
'    strSql = strSql & " AND D.intCodigoUtilizacaoDaTabelaDeValor = C.PKId"
    strsql = strsql & " AND B.intUtilizacaoDaCaracteristica = C.PKId" 'D.intCodigoUtilizacaoDaTabelaDeV = C.PKId Alterado Rafael 21/10/04
    strsql = strsql & " AND D.intCodigoCaracteristicaGeral = " & lvw_Caracteristica(intCaractImobil).SelectedItem.Tag
    strsql = strsql & " AND D.intArea = '" & Val(grd_Area.Columns("Pkid").Value) & "'"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                Call MarcaDetalhe(!Detalhe)
                .MoveNext
            Loop
        End With
    End If
End Function

Private Sub DadosExcedentes()
    strPreencheContribuinteTabs
    lvw_Detalhe(0).ListItems.Clear
    lvw_Detalhe(1).ListItems.Clear
    lvw_Detalhe(2).ListItems.Clear
End Sub

Private Sub lvw_Envolvidos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    LimpaEnvolvidos
    PreencherListaDeOpcoes dbc_intContribuinte, Item.Tag
'    dbc_intContribuinte.BoundText = Item.Tag
'    dbc_intContribuinte.Text = Item.Text
    dbc_intContribuinte_Click 2
    opt_Proprietario(0).Value = strEnvolvidos(1, Item.Index - 1)
    mblnAlterandoList = True
    
End Sub

Private Sub lvw_Melhoria_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    GravaEquipamento Val(txtPKId)
End Sub

Private Sub mskstrInscricaoAuxiliar_GotFocus()
    MarcaCampo mskstrInscricaoAuxiliar
End Sub

Private Sub mskstrInscricaoAuxiliar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricaoAuxiliar
End Sub

Private Sub mskstrInscricaoAuxiliar_LostFocus()
    If Len(Trim$(mskstrInscricaoAuxiliar.Text)) > 0 Then
        mskstrInscricaoAuxiliar.Text = String(25 - Len(mskstrInscricaoAuxiliar.Text), "0") & mskstrInscricaoAuxiliar.Text
    End If
End Sub

Private Sub tdb_DocumentosProcessos_HeadClick(ByVal ColIndex As Integer)
   gOrdenaGrid tdb_DocumentosProcessos, ColIndex
End Sub

Private Sub tdb_DocumentosProcessos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tdb_DocumentosProcessos.EOF Then
        txt_PkidDocProc.Text = tdb_DocumentosProcessos.Columns("Pkid")
        PreencherListaDeOpcoes dbc_intTiposDocumentosProcesso, tdb_DocumentosProcessos.Columns("PkidDocumentoProcesso")
        PreencheProcesso Val(tdb_DocumentosProcessos.Columns("PkidProcesso"))
        PreencheDtProcesso
        txt_strObservacoes.Text = tdb_DocumentosProcessos.Columns("Observações")
    End If
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
Dim strInscricao As String
    On Error Resume Next
    Select Case ColIndex
        Case 1
            strInscricao = Value
            Value = gstrFormataInscricao(strInscricao, TYP_IMOBILIARIA)
        Case 4
            Value = gstrCGCCPFFormatado(CStr(Value))
    End Select
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    mblnClick = False
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim lngSecaoLogradouro As Long
    Dim i As Integer
    txt_TotPredial = ""
    With tdb_Lista
        If Not .EOF And Not .BOF Then
            If mblnClick Then
                If mblnPrimeiraVez Then
                    Screen.MousePointer = vbHourglass
                    lbl_DescricaoPredios = ""
                    mblnClick = False
                    Limpa_Controles Me, False, False, False, False, True
                    dbcstrLogradouroC.BoundText = ""
                    
                    mblnSelecionou = True
                    mblnAlterando = True
                    dbcintLogradouro_Click 2
                    
                    Set gobjBanco = New clsBanco
                    
                    gobjBanco.ExecutaRollbackTrans
                    
                    blnCarregando = True
                    LeDaTabelaParaObj gstrImobiliario, Me, "SELECT IM.*, " & gstrRIGHT("IM.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, " & gstrRIGHT("IM.strInscricaoAnterior", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoAnterior FROM " & gstrImobiliario & " IM WHERE Pkid = " & .Columns("PKID").Value
                    PreencheLogradouroC Val(txtPKId)
                    txt_strCNPJCPF.Text = tdb_Lista.Columns("strCnpjCpf").Value

                    blnCarregando = False
                    
                    If Val(txtPKId.Text) = 0 Then Exit Sub
                    
                    ' TIMTIM - 22/04/2003
                    mskstrInscricao.Tag = mskstrInscricao.Text
                                                                                        
                    PreencheCkIsencao
                                                                                        
                    gCorLinhaSelecionada tdb_Lista
                    '=============

                    lngSecaoLogradouro = Val(.Columns("intSecoes").Value)
                    dbcintLogradouro_Click 2
                    'dbc_intFaceDeQuadra.BoundText = lngSecaoLogradouro
                    'dbc_intFaceDeQuadra_Click 2
                    
                    CarregaEnvolvidos
                    CarregaHistoricos Val(txtPKId)
                    'Cláudio
                    PreencheGRD
                    PreencheGRD2
                    
                    SelecionaEquipamento Val(txtPKId)
                    DadosExcedentes
                    
                'Cláudio - Vamos carregar a guia de Doc/Proc(grid, combo etc...)
                    'Carrega a Combo de Documentos Processo
                    
                    dbc_intTiposDocumentosProcesso.Tag = PreencheComboDocProc & ";strDescricao"
                    LeDaTabelaParaObj gstrDocumentos, dbc_intTiposDocumentosProcesso, PreencheComboDocProc
                    
                    'Carrega a Combo de Processos(Como se trata de grandes quantidades de registros, optei por nao prenchela somente
                    'a propriedade tag para que possa ser usada em uma busca direta
'                    dbc_intProcesso.Tag = strPreencheComboProcesso & ";strCodigo"
                    
                    'Carrega o Grid de Documentos/Processos
                    
                    PreencheGRDDocProc
                    CarregaFatorTerreno
                    If txtPKId.Text <> "" Then
                       CarregaConstrucao
                    End If
                    '=============
                    
                    CarregaFichasCadastrais
                    
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                    If mobjAux Is Nothing Then
                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                    End If
                    
                    If App.ProductName = "Tributario" Then
                        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCancelarReativar
                        If Trim(txtdtmdtcancelamento) <> "" Then
                            MDIMenu.actBarra.Bands(gstrBtnArquivo).Tools.Item(20).ToolTipText = "Reativar"
                            HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
                        Else
                            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
                            MDIMenu.actBarra.Bands(gstrBtnArquivo).Tools.Item(20).ToolTipText = "Cancelar"
                        End If
                    End If
                    
                    Screen.MousePointer = vbDefault
                End If
            End If
        End If
    End With
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    mblnClick = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrImobiliario
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    mblnClick = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Lista
End Sub

Function MarcaDetalhe(intTag As Integer)
Dim i As Integer
    For i = 1 To lvw_Detalhe(intCaractImobil).ListItems.Count
        If lvw_Detalhe(intCaractImobil).ListItems(i).Tag = intTag Then
            lvw_Detalhe(intCaractImobil).ListItems(i).Checked = True
            lvw_Detalhe(intCaractImobil).ListItems(i).Selected = True
        End If
    Next
End Function

Function SelecionaDetalhe()

'******************************************************************************************
' Data: 09/03/2003
' Alteração: - Alteração do nome do atributo intCodigoDetalheDaCaracteristica por
'            intCodigoDetalheDaCaracteristi a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/03/2003
' Alteração: - Alteração do nome do atributo intCodigoUtilizacaoDaTabelaDeValor por
'            intCodigoUtilizacaoDaTabelaDeV a fim de manter a compatibilidade com o Banco
'            de Dados.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strsql As String
    strsql = ""
    strsql = strsql & " SELECT D.intCodigoDetalheDaCaracteristi Detalhe "
    strsql = strsql & " FROM " & gstrImobiliario & " A,"
    strsql = strsql & gstrUtilizacaoDaTabelaDeValor & " C, " & gstrCaracteristicaDoImovel & " D"
    strsql = strsql & " WHERE A.PKId = D.intCodigoImobiliario "
    strsql = strsql & " AND D.intCodigoUtilizacaoDaTabelaDeV = C.PKId"
    strsql = strsql & " AND D.intCodigoCaracteristicaGeral = " & lvw_Caracteristica(intCaractImobil).SelectedItem.Tag
    strsql = strsql & " AND D.intCodigoImobiliario = " & Val(txtPKId)
    'Cláudio
    
    If intCaractImobil = 2 Then
        strsql = strsql & " AND D.intArea = '" & Val(grd_Area.Columns("Pkid").Value) & "'"
    End If
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                Call MarcaDetalhe(!Detalhe)
                .MoveNext
            Loop
        End With
    End If
End Function

'$$$$$$$$$$$$$$$ Query para montar detalhes $$$$$$$$$$$$$$$
Private Function CarregaDetalhes()
    Dim strsql As String
    
    strsql = ""
    strsql = strsql & " SELECT PKId, strNomeDoDetalhe Detalhe"
    strsql = strsql & " FROM " & gstrDetalheDaCaracteristica
    'Apenas para característica selecionada no grid de características
    strsql = strsql & " WHERE intCaracteristica = " & lvw_Caracteristica(intCaractImobil).SelectedItem.Tag & " Order By Intcodigododetalhe"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            lvw_Detalhe(intCaractImobil).ListItems.Clear
            Do While .EOF = False
                Set objList1 = lvw_Detalhe(intCaractImobil).ListItems.Add(, , !Detalhe)
                objList1.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Function

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricao
End Sub

Private Sub mskstrInscricao_LostFocus()
Dim bytFor As Byte
Dim blnFor As Boolean
Dim strInsc As String
  
  '****
  'Zerar Campos em FaceDeQuadras
  For bytFor = 1 To 24
    If Mid(Trim(mskstrInscricao.FormattedText), bytFor, 1) <> "." Then
       strInsc = strInsc & Mid(Trim(mskstrInscricao.FormattedText), bytFor, 1)
    Else
       If blnFor = True Then
          Exit For
       End If
       strInsc = strInsc & Mid(Trim(mskstrInscricao.FormattedText), bytFor, 1)
       blnFor = True
    End If
  Next
  
  If strInsc <> strInscricaoA And strInscricaoA <> ".   " Then
     strInscricaoA = strInsc
'     MontaArray2
  Else
     If strInscricaoA = "" Then
        strInscricaoA = strInsc
     End If
  End If
    
  If blnDuplicataInscricao(mskstrInscricao.Text) = False Then
      Exit Sub
  End If
  strPreencheContribuinteTabs
End Sub

Function blnDuplicataInscricao(strInscricao As String) As Boolean
Dim strsql      As String
Dim strSqlAux   As String
Dim INT_PKIDI   As Long
    If strInscricao = "" Then
        blnDuplicataInscricao = False
        Exit Function
    End If
    strsql = ""
    strsql = strsql & "SELECT count(*) as Contador FROM " & gstrImobiliario & " WHERE strInscricaoAnterior = '" & String(gintLenInscricao - Len(strInscricao), "0") & strInscricao & "'"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If adoResultado!Contador <= 0 Then
                blnDuplicataInscricao = False
                Exit Function
                Else
                strSqlAux = ""
                strSqlAux = strSqlAux & "SELECT PKId as PP FROM " & gstrImobiliario & " WHERE strInscricaoAnterior = '" & String(gintLenInscricao - Len(strInscricao), "0") & strInscricao & "'"
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSqlAux, 5, adoResultado) Then
                    If Not adoResultado.EOF Then
                        INT_PKIDI = adoResultado!PP
                        blnDuplicataInscricao = True
                    End If 'Mudei
                End If
            End If
        End If
    End If
End Function

Function PegaMaxPKId() As Long
    Dim strsql As String
        
    strsql = ""
    strsql = strsql & "SELECT MAX(PKId) as PKId "
    strsql = strsql & " FROM " & gstrImobiliario
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            PegaMaxPKId = Val(gstrENulo(adoResultado!Pkid))
        End If
    End If
    
End Function

Sub MontaColumnHeaders()
    With lvw_Historico
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Historico", 7700
    End With
End Sub
 
 'TAB GERAL
Private Sub dbcintContribuinte_Change()
    If dbcintContribuinte.MatchedWithList And blnCarregando = False Then
       dbcintContribuinte_Click 2
    End If
End Sub

Private Sub dbcintContribuinte_Click(Area As Integer)
Dim strsql As String
Dim adoResultado As ADODB.Recordset
   
    DropDownDataCombo dbcintContribuinte, Me, Area
      
    If Area = 2 Then
       If dbcintContribuinte.Locked Then Exit Sub
       CarregaContribuinte
   End If

End Sub

Private Sub CarregaContribuinte()
Dim strsql As String
    
    If dbcintContribuinte.BoundText = "" Then
       txt_strCNPJCPF.Text = ""
       Exit Sub
    End If
    
    strsql = ""
    strsql = strsql & "SELECT strCodigoAnterior, PKId, intLogradouro, intNumero, strComplemento, intBairro, intUf, intCep, " & gstrISNULL("strCNPJCPF", "0") & " AS strCNPJCPF FROM " & gstrContribuinte & " "
    strsql = strsql & "WHERE PKId = " & dbcintContribuinte.BoundText
    Set gobjBanco = New clsBanco
        
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        txt_strCNPJCPF = gstrCGCCPFFormatado(adoResultado!StrCnpjCpf)
        txt_PKIdContribuinte = adoResultado!strCodigoAnterior & Space(0)
        txtintNumero = IIf(txtintNumero = "", adoResultado!INTNUMERO & Space(0), txtintNumero)
        txtstrComplemento = adoResultado!STRCOMPLEMENTO & Space(0)
        PreencherListaDeOpcoes dbcintLogradouro, gstrVerificaCampoNulo(adoResultado!intLogradouro)
        txtintCep = gstrCEPFormatado(adoResultado!INTCEP & Space(0))
    End If
    
End Sub

Private Function strQueryTabelaOcorrencia() As String
Dim strsql As String
    
    strsql = ""
    strsql = strsql & "Select PKId, strDescricao "
    strsql = strsql & "From " & gstrOcorrencia & " "
    strsql = strsql & "Where intUtilizacaoDaOcorrencia = 6 "
    
    strQueryTabelaOcorrencia = strsql
    
End Function

Sub VerificaMascaraInscricao()
Dim strsql As String
Dim adoResultado As ADODB.Recordset
Dim strMascara   As String
    
    strMascara = ""
    strsql = ""
    strsql = strsql & "Select * From " & gstrCampoDeInscricao & " "
    strsql = strsql & "Where intTipoDeInscricao = " & TYP_IMOBILIARIA
    strsql = strsql & "Order By intSequencia"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                bytTamanhoMascara = bytTamanhoMascara + !intTamanho
                .MoveNext
            Loop
        End With
    End If
    
    mskstrInscricao.Mask = strMascara
    mskstrInscricaoEnglobada.Mask = strMascara
    mskstrInscricaoAnterior.Mask = strMascara
    
End Sub

Private Function strQuerryLoteamento()
Dim strsql As String
    strsql = ""
    strsql = strsql & " SELECT PKId, strNome"
    strsql = strsql & " FROM " & gstrLoteamento
    strsql = strsql & " ORDER BY strNome"
    strQuerryLoteamento = strsql
End Function

Function strQuerryComposicao()
Dim strsql As String
    strsql = ""
    strsql = strsql & "SELECT PKId, strDescricao "
    strsql = strsql & " FROM " & gstrComposicaoDaReceita
    strsql = strsql & " WHERE intUtilizacao = " & TYP_IMOBILIARIA
    strsql = strsql & " ORDER BY strDescricao "
    strQuerryComposicao = strsql
End Function

Function strQueryListView(Optional strModoOperacao As String) As String
Dim strsql As String

    strsql = ""
    strsql = strsql & "SELECT IM.PKId as UM,IM.PKId ," & gstrRIGHT("IM.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao," & gstrRIGHT("IM.strInscricaoAnterior", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoAnterior,"
    strsql = strsql & " CO.strNome, CO.strCNPJCPF, IM.intSecoes"
    strsql = strsql & " FROM "
    strsql = strsql & gstrImobiliario & " IM, "
    strsql = strsql & gstrContribuinte & " CO "
    strsql = strsql & " WHERE IM.intContribuinte " & strOUTJSQLServer & "= CO.PKId" & strOUTJOracle & " "
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If Not mblnAlterando Then
            strsql = strsql & " AND IM.Pkid = " & glngPegaUltimaChave(gstrImobiliario, "Pkid") + 1
        Else
            strsql = strsql & " AND IM.Pkid = " & txtPKId.Text
        End If
    End If
    
    strsql = strsql & " ORDER BY strInscricao "
    
    strQueryListView = strsql
    
End Function

'Function blnDuplicataInscricaoEnglobada(strInscricao As String) As Boolean
'Dim strSql      As String
'Dim strSqlAux   As String
'Dim INT_PKIDI   As Integer
'    If strInscricao = "" Then
'        blnDuplicataInscricaoEnglobada = False
'        Exit Function
'    End If
'    strSql = ""
'    strSql = strSql & "SELECT count(*) as Contador FROM " & gstrImobiliario & " WHERE strInscricaoAnterior = '" & strInscricao & "'"
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        If Not adoResultado.EOF Then
'            If adoResultado!Contador <= 0 Then
'                blnDuplicataInscricaoEnglobada = False
'                Exit Function
'            End If
'        End If
'    End If
'    blnDuplicataInscricaoEnglobada = True
'End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strsql                  As String
Dim intAux                  As Long
Dim i                       As Integer
Dim intEdificado            As Integer
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        mskstrInscricaoAuxiliar_LostFocus
    End If
    
    If UCase(strModoOperacao) = UCase(gstrAplicar) Then
        ToolBarGeral strModoOperacao, gstrImobiliario, mblnAlterando, tdb_Lista, Me, mobjAux, "", strQueryAplicar
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = gstrCancelarReativar Then
        If mblnAlterando = True And Trim(txtPKId) <> "" Then
            If gblnExclusaoGravacaoOk("SALVAR", "Deseja realmente " & IIf(Trim(txtdtmdtcancelamento) <> "", "Reativar", "Cancelar") & " esta Inscrição") Then
                Screen.MousePointer = vbArrow
                GravaCancelamento
                Screen.MousePointer = vbDefault
            End If
        End If
            Exit Sub
    End If
        
    If strModoOperacao = gstrImprimir Then
        ImprimeRelatorio rptcadImobiliario, gstrStoredProcedure("sp_cadimobiliario", Val(txtPKId), True)
        
    ElseIf UCase(strModoOperacao) = UCase(gstrLocalizar) Then
        mblnClick = True
        mblnPrimeiraVez = True
        LocalizarImobiliario
        Screen.MousePointer = vbDefault
        Exit Sub
        
    ElseIf UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
        If UCase(Me.ActiveControl.Name) = "DBCINTCONTRIBUINTE" And IsNumeric(dbcintContribuinte.Text) Then
            strsql = ""
            strsql = strsql & " SELECT PKId, strNome FROM " & gstrContribuinte
            strsql = strsql & " WHERE strCodigoAnterior LIKE '" & dbcintContribuinte.Text & "%'"
            
            LeDaTabelaParaObj gstrContribuinte, dbcintContribuinte, strsql
        ElseIf UCase(Me.ActiveControl.Name) = "DBC_INTRESUMOTIPOPADRAO" Then
            Exit Sub
        ElseIf UCase(Me.ActiveControl.Name) = "DBCINTLOGRADOURO" Then
            dbcintLogradouro.Tag = strQueryLogradouro(False) & ";L.strDescricao"
            PreencherListaDeOpcoes Me.ActiveControl
            dbcintLogradouro.Tag = strQueryLogradouro(True) & ";L.strDescricao"
        Else
            PreencherListaDeOpcoes Me.ActiveControl
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        If tab_3DPasta.Tab = 2 Then
            LimpaEnvolvidos True
            Exit Sub
        ElseIf tab_3DPasta.Tab = 9 Then 'Doc/Proc
            LimpaTabDocProc False
            Exit Sub
        Else
            If PKId_Temporario <> 0 Then
                strsql = ""
                strsql = strsql & " DELETE "
                strsql = strsql & " FROM " & gstrCaracteristicaDoImovelRural
                strsql = strsql & " WHERE "
                strsql = strsql & " intCodigoImobiliario = " & Val(PKId_Temporario)
                Set gobjBanco = New clsBanco
                gobjBanco.Execute strsql
                PKId_Temporario = 0
                strsql = ""
            End If
            strLimpaContribuintesTabs
            For i = 1 To lvw_Detalhe(0).ListItems.Count
                lvw_Detalhe(0).ListItems(i).Checked = False
                lvw_Detalhe(0).ListItems(1).Selected = True
            Next
            i = 0
            For i = 1 To lvw_Detalhe(1).ListItems.Count
                lvw_Detalhe(1).ListItems(i).Checked = False
                lvw_Detalhe(1).ListItems(1).Selected = True
            Next
            i = 0
            For i = 1 To lvw_Detalhe(2).ListItems.Count
                lvw_Detalhe(2).ListItems(i).Checked = False
                lvw_Detalhe(2).ListItems(1).Selected = True
            Next
            LimpaTabDocProc True  'Cláudio - Doc/Proc
            LimpaGrid2
            Set dbc_intResumoTipoPadrao.RowSource = Nothing
            dbc_intResumoTipoPadrao.Text = ""
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelarReativar
        End If
    End If
    
    If UCase(strModoOperacao) = "SALVAR" Then
        If blnDadosOK(strModoOperacao) Then
            If tab_3DPasta.Tab = 9 Then 'Cláudio - Doc/Proc
                If gblnExclusaoGravacaoOk("A") Then
                    If txt_PkidDocProc.Text <> "" Then 'Alteração
                        AlteraDocProc (txt_PkidDocProc)
                        LimpaTabDocProc False
                        PreencheGRDDocProc
                    Else
                        GravaDocProc
                        LimpaTabDocProc False
                        PreencheGRDDocProc
                    End If
                End If
                'LimpaObjeto Me
                Exit Sub
            Else
                If mskstrInscricaoEnglobada.Text <> "" Then
                    'If blnDuplicataInscricaoEnglobada(mskstrInscricaoEnglobada.Text) = False Then
                    '    ExibeMensagem "Imóvel a ser englobado não existe."
                    '    mskstrInscricaoEnglobada.SetFocus
                    '    Screen.MousePointer = vbDefault
                    '    Exit Sub
                    'End If
                    If mskstrInscricaoEnglobada.Text = mskstrInscricao.Text Then
                        ExibeMensagem "A inscrição englobada não pode ser igual à inscrição deste imóvel."
                        mskstrInscricaoEnglobada.SetFocus
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
            End If
        Else
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        
        ' TIMTIM - 22/04/2003
        If UCase(strModoOperacao) = UCase(gstrNovo) Or UCase(strModoOperacao) = UCase(gstrLimpar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
            mskstrInscricao.Tag = Space$(0)
        End If
        
    End If
    
    intAux = 0
    If mblnAlterando Then
        intAux = Val(txtPKId) 'Mudei
    End If
    
    If strModoOperacao = gstrDeletar Then
    
        If Trim(txtdtmdtcancelamento) <> "" Then
            ExibeMensagem "Registro não pode ser excluído pois esta cancelado."
            Exit Sub
        End If
        
        If tab_3DPasta.Tab = 6 Then
            
            If tdbGridSelecionada Is Nothing Then
                ExibeMensagem "Não há Registro Selecionado a ser Excluído."
                Exit Sub
            ElseIf tdbGridSelecionada.Columns("Pkid") = "" Then
                ExibeMensagem "Não há Registro Selecionado a ser Excluído."
                Exit Sub
            End If
            
            If gblnExclusaoGravacaoOk(strModoOperacao, "Confirma Exclusão") Then
                If InStr(UCase(tdbGridSelecionada.Name), "AREA") > 0 Then
                    DeletaValores intAux, tdbGridSelecionada.Columns("Pkid")
                    PreencheGRD
                ElseIf InStr(UCase(tdbGridSelecionada.Name), "TESTADA") > 0 Then
                    DeletaValores2 intAux, tdbGridSelecionada.Columns("Pkid")
                    PreencheFaceDeQuadra
                    PreencheGRD2
                End If
                chkbytEdificado.Value = RetornaEdificacao
                'LimpaObjeto Me
                'txt_intBairro.Text = ""
                'txtintCep.text = ""
                'txt_intUf.Text = ""
            End If
            
            Exit Sub
            
        ElseIf tab_3DPasta.Tab = 9 Then
            If Len(txt_PkidDocProc) > 0 Then
                If gblnExclusaoGravacaoOk("E") Then
                    DeletaDocProc (txt_PkidDocProc)
                    LimpaTabDocProc False
                    PreencheGRDDocProc
                End If
                LimpaObjeto Me
                txt_intBairro.Text = ""
                txtintCep.Text = ""
                txt_intUF.Text = ""
            End If
            Exit Sub
        Else
            'Nao vamos excluir a inscricao, somente predios e documentos
            Exit Sub
            'Set gobjBanco = New clsBanco
            'gobjBanco.ExecutaBeginTrans
            'DeletaHistoricos intAux
            'DeletaValores intAux
            'DeletaValores2 intAux
            'DeletaEquipamento intAux
        End If
    End If
    
    If strModoOperacao = gstrIncluirItem Then
        IncluiAlteraItemLista mblnAlterando
    ElseIf strModoOperacao = gstrExcluirItem Then
        ExcluiItemLista
    End If
    
    intEdificado = RetornaEdificacao
        
    'Caso ainda nao tenha sido iniciada uma transacao no listview das caracteristicas
    If Not blnEmTransacao Then
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
    End If
    
    If ToolBarGeral(strModoOperacao, gstrImobiliario, mblnAlterando, tdb_Lista, Me, mobjAux, strQueryListView(strModoOperacao), , , , False) Then  'Mudei
        strLimpaContribuintesTabs False
        If intAux = 0 Then
            intAux = PegaMaxPKId
        End If
        
        If strModoOperacao = gstrSalvar Then
            GravaHistoricos intAux
            GravaValores intAux, intEdificado
            GravaValores2 intAux
            GravaEnvolvidos intAux
        ElseIf strModoOperacao = gstrDeletar Then
            DeletaHistoricos intAux
            DeletaValores intAux
            DeletaValores2 intAux
            DeletaEquipamento intAux
            DeletaEnvolvidos intAux
        End If
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelarReativar
        PKId_Temporario = 0
        tab_3DPasta.Tab = 0
        mblnPrimeiraVez = False
        mblnAlterando = False
        If strModoOperacao = gstrSalvar Or strModoOperacao = gstrDeletar Then
            gobjBanco.ExecutaCommitTrans
            HabilitaDesabilitaBotao1 strModoOperacao = gstrSalvar, gstrMnuArquivo, gstrDeletar
        Else
            gobjBanco.ExecutaRollbackTrans
        End If
        LimpaObjeto Me
        txt_intBairro.Text = ""
        txtintCep.Text = ""
        txt_intUF.Text = ""
        lvw_FatorTerreno.ListItems.Clear
        lvw_CaracPredio.ListItems.Clear
        lbl_DescricaoPredios = ""
        blnEmTransacao = False
    Else
        If (strModoOperacao = gstrSalvar Or strModoOperacao = gstrDeletar) And mblnAlterando = True Then
            gobjBanco.ExecutaRollbackTrans
        End If
        grd_Testada.Refresh
    End If
    
    If strModoOperacao = gstrNovo And mblnAlterando = False Then
        dbcintContribuinte.Locked = False
    End If
    
    If strModoOperacao = gstrNovo Then
        lvw_Historico.ListItems.Clear
        txt_Historico = ""
        mblnAlterando = False
        dbc_intFaceDeQuadra.BoundText = ""
        lvw_Melhoria.ListItems.Clear
        txt_TotPredial = ""
        mskstrInscricao.SetFocus
    End If
    Screen.MousePointer = 0
End Sub

Private Sub mskstrInscricaoEnglobada_GotFocus()
    MarcaCampo mskstrInscricaoEnglobada
End Sub

Private Sub mskstrInscricaoEnglobada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricaoEnglobada
End Sub

Private Sub tdd_CategoriaConstrucao_DropDownClose()
    
    If tdd_CategoriaConstrucao.Row = -1 Then
        ExibeMensagem "Selecione uma Categoria de Construção válida."
        grd_Area.Columns("PkidCategoriaConstrucao").Value = ""
        grd_Area.Columns("Categoria da Construção").Value = ""
    Else
        If mblnAlterando Then
            If grd_Area.Columns(0).Value <> "" Then
                If MsgBox("Todas as caracterísitcas dessa área serão excluidas." & Chr(13) & "Você deseja realmente trocar a categoria da construção?", vbQuestion + vbYesNo) = vbYes Then
                    DeletaCaracteristicasDoImovel
                    grd_Area.Columns("PkidCategoriaConstrucao").Value = tdd_CategoriaConstrucao.Columns("Pkid").Value
                    LeDaTabelaParaObj "", dbc_intResumoTipoPadrao, strQueryResumoTipoPadrao(tdd_CategoriaConstrucao.Columns("Pkid").Value)
                    lvw_CaracPredio.ListItems.Clear
                    lvw_Detalhe(2).ListItems.Clear
                    lbl_DescricaoPredios = ""
                    CarregaCaracteristica
                    blnGridCategoria = False
                Else
                    MontaArray
                End If
            Else
                grd_Area.Columns("PkidCategoriaConstrucao").Value = tdd_CategoriaConstrucao.Columns("Pkid").Value
                LeDaTabelaParaObj "", dbc_intResumoTipoPadrao, strQueryResumoTipoPadrao(tdd_CategoriaConstrucao.Columns("Pkid").Value)
            End If
        Else
            grd_Area.Columns("PkidCategoriaConstrucao").Value = tdd_CategoriaConstrucao.Columns("Pkid").Value
            LeDaTabelaParaObj "", dbc_intResumoTipoPadrao, strQueryResumoTipoPadrao(tdd_CategoriaConstrucao.Columns("Pkid").Value)
        End If
    End If

End Sub

Private Sub tdd_Testada_DropDownClose()
    If tdd_Testada.Row <> -1 Then
       grd_Testada.Columns(5) = tdd_Testada.Columns(2).Value
       grd_Testada.Columns(6) = tdd_Testada.Columns(0).Value
       If tdd_Testada.Columns("Tipo de Testada").Text <> "" Then
          'grd_Testada.Update
          VerificaTestada tdd_Testada.Columns(0).Value, grd_Testada.Row
       End If
    End If
    grd_Testada.SetFocus
    grd_Testada.Row = intRow
    gCorLinhaSelecionada grd_Testada
End Sub

Private Sub tlb_Historico_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim alterando As Boolean
    If lvw_Historico.ListItems.Count <> 0 Then
        If lvw_Historico.SelectedItem.Selected Then
            alterando = True
            Else
            alterando = False
        End If
    End If
    Select Case UCase(Button.Key)
        Case gstrSalvar
            If txt_Historico = "" Then Exit Sub
            If alterando Then
                lvw_Historico.SelectedItem.Text = txt_Historico
            Else
                lvw_Historico.ListItems.Add , , txt_Historico
            End If
            txt_Historico = ""
        Case gstrNovo
            txt_Historico = ""
            txt_Historico.SetFocus
        Case gstrDeletar
            If txt_Historico = "" Then Exit Sub
            If lvw_Historico.SelectedItem.Selected Then
                lvw_Historico.ListItems.Remove (lvw_Historico.SelectedItem.Index)
                txt_Historico = ""
            End If
    End Select
    If lvw_Historico.ListItems.Count <> 0 Then
        lvw_Historico.SelectedItem.Selected = False
    End If
End Sub

Private Sub txt_bitDigito_GotFocus()
    MarcaCampo txt_bitDigito
End Sub

Private Sub txt_bitDigito_LostFocus()
    If VerificaEmpenhoProcesso(Trim(txt_strCodigo), Val(txt_bitDigito), Val(txt_intExercicio)) Then
        PreencheDtProcesso
    End If
End Sub

Private Sub txt_intExercicio_GotFocus()
    If txt_intExercicio = "" Then
        txt_intExercicio.Text = Year(gstrDataDoSistema())
    End If
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_strCodigo_GotFocus()
     MarcaCampo txt_strCodigo
End Sub

Private Sub txtdblArea_Change()
    txt_dblArea = txtdblArea
End Sub

Private Sub txtdblArea_GotFocus()
    MarcaCampo txtdblArea
End Sub

Private Sub txtdtmdtescritura_GotFocus()
    MarcaCampo txtdtmdtescritura
End Sub

Private Sub txtdtmdtescritura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdtescritura
End Sub

Private Sub txtdtmdtescritura_LostFocus()
    txtdtmdtescritura = gstrDataFormatada(txtdtmdtescritura)
End Sub

Private Sub txtdtmDtMatricula_GotFocus()
    MarcaCampo txtdtmdtmatricula
End Sub

Private Sub txtdtmDtMatricula_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdtmatricula
End Sub

Private Sub txtdtmDtMatricula_LostFocus()
    txtdtmdtmatricula = gstrDataFormatada(txtdtmdtmatricula)
End Sub

Private Sub txtintCepC_GotFocus()
    MarcaCampo txtintCepC
End Sub

Private Sub txtintCepC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepC
End Sub

Private Sub txtintCepC_LostFocus()
    txtintCepC = gstrCEPFormatado(txtintCepC)
    CepLogradouro txtintCepC, dbcstrLogradouroC, txtstrBairroC, dbcintMunicipioC, dbcintUFC, dbcintTipoLogradouro, dbcintTituloLogradouro, , True, False, True, True, True, True
End Sub

Private Sub txtintCodigoLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintFolha_GotFocus()
    MarcaCampo txtintfolha
End Sub

Private Sub txtintfolha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintfolha
End Sub

Private Sub txtintLivro_GotFocus()
    MarcaCampo txtintlivro
End Sub

Private Sub txtintlivro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintlivro
End Sub

Private Sub txtintNumeroC_GotFocus()
    MarcaCampo txtintNumeroC
End Sub

Private Sub txtintNumeroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumeroC
End Sub

Private Sub txtstrBairroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrBairroC
End Sub

Private Sub txtstrCartorio_GotFocus()
    MarcaCampo txtstrCartorio
End Sub

Private Sub txtstrCartorio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCartorio
End Sub

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub txtstrComplementoC_GotFocus()
    MarcaCampo txtstrComplementoC
End Sub

Private Sub txtstrComplementoC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplementoC
End Sub

'Private Sub txtstrDesmembramento_KeyPress(KeyAscii As Integer)
'   CaracterValido KeyAscii, "N", txtstrDesmembramento
'End Sub

Private Sub txtstrEmissao_GotFocus()
    MarcaCampo txtstrEmissao
End Sub

Private Sub txtstrEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrEmissao
End Sub

Private Sub txtstrmatricula_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrmatricula
End Sub

'Private Sub txtstrHabitese_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "N", txtstrHabitese
'End Sub

Private Sub mskstrInscricaoAnterior_GotFocus()
    MarcaCampo mskstrInscricaoAnterior
    'tab_3DPasta.Tab = 0
End Sub

Private Sub mskstrInscricaoAnterior_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricaoAnterior
End Sub

Private Sub txtintNumero_GotFocus()
    MarcaCampo txtintNumero
    'tab_3DPasta.Tab = 0
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
    'tab_3DPasta.Tab = 0
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
    'tab_3DPasta.Tab = 0
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub

Private Sub txtintCEP_LostFocus()
    If txtintCep.Text = "" Then
        Exit Sub
    End If
    
    txtintCep = gstrCEPFormatado(txtintCep)
    CepLogradouro txtintCep, dbcintLogradouro, txt_intBairro, , txt_intUF, , , , True, False, False, False, False, False
    'If gblnCepValido(txtintCep, dbcintLogradouro) = False Then
    '    MsgBox "CEP inválido para o logradouro cadastrado "
    'End If
End Sub

Private Sub txtstrLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLote
End Sub

Private Sub txtstrLote_GotFocus()
    MarcaCampo txtstrLote
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumero
End Sub

Private Sub txtstrQuadra_GotFocus()
    MarcaCampo txtstrQuadra
End Sub

Private Sub txtstrMatricula_GotFocus()
    MarcaCampo txtstrmatricula
End Sub

'TAB PROMISSARIO
Function strEncheCamposPromissario()

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela variável
'            gstrISNULL.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strsql As String
    Dim adoResultado As ADODB.Recordset
    
    If dbcintPromissario.MatchedWithList Then
        strsql = ""
        strsql = strsql & " SELECT bytNaturezaJuridica, " & gstrISNULL("strCNPJCPF", "0") & " as strCNPJCPFP "
        strsql = strsql & " FROM " & gstrContribuinte & " "
        strsql = strsql & " WHERE PKId = " & dbcintPromissario.BoundText
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    optbytNaturezaJuridica(!bytNaturezaJuridica) = True
                    txt_strCNPJCPFP = gstrCGCCPFFormatado(!strCNPJCPFP)
                    .MoveNext
                Loop
            End With
        End If
    End If
End Function

Private Sub dbcintPromissario_Click(Area As Integer)
   DropDownDataCombo dbcintPromissario, Me, Area
   If Area = 2 Then
      MostraDadosContribuinte (dbcintPromissario.BoundText)
      strEncheCamposPromissario
   End If
End Sub

Private Function MostraDadosContribuinte(intBound As Long) As Boolean
Dim strsql As String
On Error Resume Next
    
    strsql = ""
    strsql = strsql & "SELECT CO.strBairroC,"
    strsql = strsql & " TL.strSigla " & strCONCAT & "' '" & strCONCAT
    strsql = strsql & " TTL.strSigla " & strCONCAT & "' '" & strCONCAT
    strsql = strsql & " CO.strLogradouroC AS strLogradouroC,"
    strsql = strsql & " CO.intNumeroC,"
    strsql = strsql & " CO.strComplementoC,"
    strsql = strsql & " CO.intCEPC,"
    strsql = strsql & " CO.strDistritoC,"
    strsql = strsql & " CD.strDescricao,"
    strsql = strsql & " UF.strSigla"
    strsql = strsql & " FROM "
    strsql = strsql & gstrContribuinte & " CO, "
    strsql = strsql & gstrCidade & " CD, "
    strsql = strsql & gstrTipoLogradouro & " TL, "
    strsql = strsql & gstrTituloLogradouro & " TTL, "
    strsql = strsql & gstrUF & " UF"
    strsql = strsql & " WHERE intMunicipioC = CD.PKId  AND"
    strsql = strsql & " intUFC = UF.PKId AND"
    strsql = strsql & " TL.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CO.intTipoLogradouro AND"
    strsql = strsql & " TTL.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CO.intTituloLogradouro AND"
    strsql = strsql & " CO.PKId = " & intBound
      
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                txt_Bairro = gstrVerificaCampoNulo(!strBairroC)
                txt_Cep = gstrCEPFormatado(gstrVerificaCampoNulo(!intcepc))
                txt_Complemento = gstrVerificaCampoNulo(!strComplementoC)
                txt_Distrito = gstrVerificaCampoNulo(!strDistritoC)
                txt_Logradouro = gstrVerificaCampoNulo(!strlogradouroc)
                txt_Municipio = gstrVerificaCampoNulo(!strDescricao)
                txt_Numero = gstrVerificaCampoNulo(!intNumeroC)
                txt_UF = gstrVerificaCampoNulo(!strsigla)
                MostraDadosContribuinte = True
                .MoveNext
            Loop
            strEncheCamposPromissario
        End With
    End If
End Function

Private Sub dbcintPromissario_Validate(Cancel As Boolean)
    strEncheCamposPromissario
End Sub

'TAB ÁREAS E HISTÓRICOS

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
Dim intCodImobiliario As Integer
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    
    If PreviousTab <> tab_3DPasta.Tab Then
        Select Case tab_3DPasta.Tab
            Case 0
            
            Case 1
                If dbcintPromissario.BoundText <> "" Then
                    MostraDadosContribuinte (dbcintPromissario.BoundText)
                End If
            Case 2
                If dbc_intContribuinte.BoundText <> "" Then
                    dbc_intContribuinte_Click 2
                End If
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
            Case 3
                grd_Area.SetFocus
                Me.Refresh
            Case 4
                '1 = Imobiliário Geral
                If mblnAlterando = False And PKId_Temporario = 0 Then
                   PKId_Temporario = Timer * 100
                End If
                intCaractImobil = 0
                CarregaCaracteristica
                SelecionaDetalhe
            Case 5
                '2 = Imobiliário Terreno
                If mblnAlterando = False And PKId_Temporario = 0 Then
                   PKId_Temporario = Timer * 100
                End If
                intCaractImobil = 1
                CarregaCaracteristica
                If bytbaselimpa = 0 Then
                    lvw_Caracteristica_ItemClick 1, lvw_Caracteristica.Item(1).ListItems(1)
                End If
                'CarregaDetalhes
                               
            Case 6
                '3 = Imobiliário Construção
                If mblnAlterando = False And PKId_Temporario = 0 Then
                   PKId_Temporario = Timer * 100
                End If
                If Val(grd_Area.Columns("Pkid").Value) > 0 Then
                    LeDaTabelaParaObj "", dbc_intResumoTipoPadrao, strQueryResumoTipoPadrao(grd_Area.Columns("PkidCategoriaConstrucao").Value)
                End If
                intCaractImobil = 2
                'lvw_Caracteristica(intCaractImobil).ListItems.Clear
                lvw_Detalhe(intCaractImobil).ListItems.Clear
                'CarregaCaracteristica
                
            Case 7
'                dbc_intFaceDeQuadra.ListField = ""
'                dbc_intFaceDeQuadra.Text = ""
'                dbc_intFaceDeQuadra.Tag = strQueryFaceDeQuadra & ";LO.strDescricao"
                dbc_intFaceDeQuadra.HelpContextID = 0
                
                dbc_intFaceDeQuadra.ListField = ""
                dbc_intFaceDeQuadra.Text = ""
                dbc_intFaceDeQuadra.Tag = strQueryFaceDeQuadra & ";LO.strDescricao"
                If txtPKId.Text <> "" Then
                    dbc_intFaceDeQuadra.HelpContextID = 1
                    LeDaTabelaParaObj "", dbc_intFaceDeQuadra, Left(dbc_intFaceDeQuadra.Tag, Len(dbc_intFaceDeQuadra.Tag) - 90) & " and lo.pkid = (select intlogradouro from " & gstrImobiliario & " where pkid = " & txtPKId.Text & ")"
                
                    If dbc_intFaceDeQuadra.MatchedWithList = True Then
                        CarregaMelhoria
                    End If
                End If
            End Select
                
         
        'Call lvw_Caracteristica_ItemClick(0, )
    End If
    If Trim(txtdtmdtcancelamento) <> "" Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    End If
    
End Sub

Function GravaDetalhe(lngPkid As Long)
Dim i      As Integer
Dim strsql As String
    
    On Error GoTo err_GravaValores
    
    If bytDBType = Oracle Then
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
    End If
    
    strsql = ""
    strsql = strsql & " DELETE "
    strsql = strsql & " FROM " & gstrCaracteristicaDoImovel
    strsql = strsql & " WHERE "
    'Código do imobiliário
    strsql = strsql & " intCodigoImobiliario = " & Val(lngPkid)
    If intCaractImobil = 2 Then
        strsql = strsql & " AND intArea = '" & grd_Area.Columns("Pkid").Value & "'"
    End If
    'Código da utilização
    strsql = strsql & " AND intCodigoUtilizacaoDaTabelaDeV = " & (intCaractImobil + 1)
    'Código da Característica geral
    strsql = strsql & " AND intCodigoCaracteristicaGeral = " & lvw_Caracteristica(intCaractImobil).SelectedItem.Tag
    Set gobjBanco = New clsBanco
    If Not gobjBanco.Execute(strsql, False) Then
        'gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If

    For i = 1 To lvw_Detalhe(intCaractImobil).ListItems.Count
        strsql = ""
        If lvw_Detalhe(intCaractImobil).ListItems(i).Checked = True Then
            strsql = ""
            strsql = strsql & " INSERT INTO " & gstrCaracteristicaDoImovel & "(intCodigoImobiliario,intCodigoCaracteristicaGeral,"
            strsql = strsql & " intCodigoDetalheDaCaracteristi,intCodigoUtilizacaoDaTabelaDeV " & IIf(intCaractImobil = 2, ", intArea", "") & ") VALUES ("
            'Código do imobiliário
            strsql = strsql & Trim(lngPkid)
            'Código da Característica geral
            strsql = strsql & "," & lvw_Caracteristica(intCaractImobil).SelectedItem.Tag
            'Código do detalhe da característica
            strsql = strsql & "," & lvw_Detalhe(intCaractImobil).ListItems(i).Tag
            'Código da utilização
            strsql = strsql & "," & (intCaractImobil + 1)
            If intCaractImobil = 2 Then
                strsql = strsql & ",'" & grd_Area.Columns("Pkid").Value & "'"
            End If
            strsql = strsql & ")"
            Set gobjBanco = New clsBanco
            If Not gobjBanco.Execute(strsql, False) Then
                'gobjBanco.ExecutaRollbackTrans
                Exit Function
            End If
        End If
    Next
    
    Screen.MousePointer = vbDefault
    
    'gobjBanco.ExecutaCommitTrans
    Exit Function
    
err_GravaValores:
    gobjBanco.ExecutaRollbackTrans
End Function

Private Sub CarregaCaracteristica()
    Dim strsql As String
    
    If intCaractImobil = 2 Then
        strsql = ""
        strsql = strsql & " SELECT PKId,intcodigodacaracteristica, strNomeDaCaracteristica Caracteristica"
        strsql = strsql & " FROM " & gstrCaracteristicaGeral
        '1 = Imobiliário Geral
        '2 = Imobiliário Terreno
        '3 = Imobiliário Construção
        strsql = strsql & " WHERE intUtilizacaoDaCaracteristica = " & (intCaractImobil + 1)
        
        'Vamos filtrar a Categoria Construcao
        If intCaractImobil = 2 Then
            strsql = strsql & " AND intCategoriaConstrucao = " & IIf(grd_Area.Columns("PkidCategoriaConstrucao").Value <> "", grd_Area.Columns("PkidCategoriaConstrucao").Value, 0)
        End If
        strsql = strsql & " ORDER BY intcodigodacaracteristica "
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
            
            With adoResultado
                If Not .EOF Then
                    lvw_Caracteristica(intCaractImobil).ListItems.Clear
                    Do While .EOF = False
                        Set objList1 = lvw_Caracteristica(intCaractImobil).ListItems.Add(, , gstrENulo(!intcodigodacaracteristica))
                        objList1.SubItems(1) = gstrENulo(!Caracteristica)
                        objList1.Tag = !Pkid
                        .MoveNext
                    Loop
                    lvw_Caracteristica(intCaractImobil).Refresh
                    lvw_Caracteristica(intCaractImobil).SetFocus
                    lvw_Caracteristica(intCaractImobil).ListItems(1).Selected = True
                    If Not intCaractImobil = 2 Then CarregaDetalhes
                Else
                    lvw_Caracteristica(intCaractImobil).ListItems.Clear
                    lvw_Detalhe(intCaractImobil).ListItems.Clear
                End If
                
            End With
        End If
    Else
        strsql = ""
        strsql = strsql & " SELECT PKId, strNomeDaCaracteristica Caracteristica"
        strsql = strsql & " FROM " & gstrCaracteristicaGeral
        '1 = Imobiliário Geral
        '2 = Imobiliário Terreno
        '3 = Imobiliário Construção
        strsql = strsql & " WHERE intUtilizacaoDaCaracteristica = " & (intCaractImobil + 1)
        
        'Vamos filtrar a Categoria Construcao
        If intCaractImobil = 2 Then
            strsql = strsql & " AND intCategoriaConstrucao = " & IIf(grd_Area.Columns("PkidCategoriaConstrucao").Value <> "", grd_Area.Columns("PkidCategoriaConstrucao").Value, 0)
        End If
        strsql = strsql & " ORDER BY strNomeDaCaracteristica "
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
            
            With adoResultado
                If Not .EOF Then
                    lvw_Caracteristica(intCaractImobil).ListItems.Clear
                    Do While .EOF = False
                        Set objList1 = lvw_Caracteristica(intCaractImobil).ListItems.Add(, , !Caracteristica)
    
                        objList1.Tag = !Pkid
                        .MoveNext
                    Loop
                    lvw_Caracteristica(intCaractImobil).Refresh
                    lvw_Caracteristica(intCaractImobil).SetFocus
                    lvw_Caracteristica(intCaractImobil).ListItems(1).Selected = True
                    If Not intCaractImobil = 2 Then CarregaDetalhes
                Else
                    lvw_Caracteristica(intCaractImobil).ListItems.Clear
                    lvw_Detalhe(intCaractImobil).ListItems.Clear
                End If
                If adoResultado.BOF Then
                    bytbaselimpa = 1
                Else
                    bytbaselimpa = 0
                End If
                
            End With
        End If
    End If
    
DoEvents
'lvw_Detalhe(intCaractImobil).ListItems(1).Selected = True
End Sub

Function PreencheGRD()
Dim strsql As String
    LimpaGrid
'    If Me.ActiveControl.Name = "lvw_Lista" Then
'        strSql = strQuerryGrid
'        Else
    MontaArray

'        strSql = ""
'        strSql = strSql & "SELECT strNomeDaArea "
'        strSql = strSql & "FROM " & gstrTipoDeArea & " "
'        If chkbytEdificado.Value = 0 Then
'            strSql = strSql & " WHERE bytPassivaDeCM = 1"
'        End If
'
'        Set gobjBanco = New clsBanco
'        gobjBanco.CriaADO strSql, 5, adoTdb
'
'        Y.ReDim 0, adoTdb.RecordCount - 1, 0, 0
'        Dim varAux As Variant
'        Do While Not adoTdb.EOF
'
'            varAux = adoTdb!strNomeDaArea
'            Y(adoTdb.AbsolutePosition - 1, 0) = varAux
'
'            adoTdb.MoveNext
'        Loop
'        Set tdd_Area.Array = Y
'        tdd_Area.ReBind
'        tdd_Area.Refresh

End Function

Private Function MontaDropDownAreaCM()
Dim strsql As String
    LimpaGridArea

        strsql = ""
        strsql = strsql & "SELECT PKId, strNomeDaArea "
        strsql = strsql & "FROM " & gstrTipoDeArea & " "
'        If chkbytEdificado.Value = 0 Then
'            strSql = strSql & " WHERE bytPassivaDeCM = 1"
'        End If
        
        LeDaTabelaParaObj "", tdd_Area, strsql

End Function

Private Sub txtdblArea_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblArea
End Sub

Private Sub txtdblArea_LostFocus()
    txtdblArea = gvntConvVrDoSql(txtdblArea)
End Sub

Function strQuerryGrid() As String
Dim strsql As String
    strsql = ""
    strsql = strsql & "SELECT VL.intX, VL.strY FROM "
    strsql = strsql & gstrValorCompoRec & " VL, "
    strsql = strsql & gstrComposicaoDaReceita & " CP "
    strsql = strsql & " WHERE CP.intCodigo = VL.intComposicao "
    strsql = strsql & " AND CP.PKId = " & txtPKId 'Mudei
    strQuerryGrid = strsql
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftDown, AltDown, CtrlDown
    Select Case KeyCode
        Case vbKeyEscape
            If Not IsNull(tdd_Area.SelectedItem) Then
                grd_Area.SelStart = Len(grd_Area.Text)
            End If
            SendKeys "{RIGHT}"
            Exit Sub
    End Select
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
End Sub

Private Sub grd_Area_KeyPress(KeyAscii As Integer)
    
    Select Case grd_Area.Col
    Case 2
        CaracterValido KeyAscii, "A", grd_Area
    Case 3
        CaracterValido KeyAscii, "V", grd_Area
        
    Case 4
        CaracterValido KeyAscii, "V", grd_Area
    Case 7
        CaracterValido KeyAscii, "N", grd_Area
    Case 8
        CaracterValido KeyAscii, "D", grd_Area
    Case 9, 10, 11, 12, 13, 15, 16
        CaracterValido KeyAscii, "N", grd_Area
    End Select
End Sub

Private Sub MontaArray()
    Dim varAux As Variant
    Dim dblTotPredial As Double

    strsql = ""
    strsql = strsql & "SELECT AI.Pkid, AI.intNEdificacao, AI.intTipoDeArea, AI.intMedidaDaArea, AI.dblFracaoIdeal,"
    strsql = strsql & " CC.strDescricao strCategoriaConstrucao,"
    strsql = strsql & " AI.intCategoriaConstrucao PkidCategoriaConstrucao,"
    strsql = strsql & " AI.intNPavimento, AI.dtmUltimaReforma, CC.pkid as intCategoriaConstrucao, "
    strsql = strsql & "AI.intQuarto, "
    strsql = strsql & "AI.intSala, "
    strsql = strsql & "AI.intCozinha, "
    strsql = strsql & "AI.intBanheiro, "
    strsql = strsql & "AI.intAndar, "
    strsql = strsql & "AI.intElevador, "
    strsql = strsql & "AI.intSuite, "
    strsql = strsql & "AI.intGaragem, "
    strsql = strsql & "AI.intServicoHotelaria "
    strsql = strsql & " FROM " & gstrAreaImobiliario & " AI, "
    strsql = strsql & gstrCategoriaConstrucao & " CC"
    
    'If mblnAlterando = True Then
        strsql = strsql & " WHERE AI.intImobiliario = '" & txtPKId & "'" 'Mudei
        strsql = strsql & " AND CC.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " AI.intCategoriaConstrucao"
    'End If
    
    dblTotPredial = 0
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoRec) Then
        If Not adoRec.EOF Then
            intCategoriaConstrucao = IIf(IsNull(adoRec!intCategoriaConstrucao), "0", adoRec!intCategoriaConstrucao) 'Vamos pegar o ID para poder preencher a caracterisitca do terreno
            Set A = New XArrayDB
            A.Clear
            With adoRec
                If Not .EOF And mblnAlterando Then
                    A.ReDim 0, .RecordCount - 1, 0, 17
                    Do While Not .EOF
                        varAux = .Fields(0)
                        A(.AbsolutePosition - 1, 0) = varAux
                        varAux = .Fields(1)
                        A(.AbsolutePosition - 1, 1) = varAux
                        varAux = .Fields(2)
                        A(.AbsolutePosition - 1, 2) = varAux
                        dblTotPredial = dblTotPredial + CDbl(gstrConvVrDoSql(gstrENulo(.Fields(3)), , , True))
                        varAux = .Fields(3)
                        A(.AbsolutePosition - 1, 3) = varAux
                        varAux = .Fields(4)
                        A(.AbsolutePosition - 1, 4) = varAux
                        varAux = .Fields(5)
                        A(.AbsolutePosition - 1, 5) = varAux
                        varAux = .Fields(6)
                        A(.AbsolutePosition - 1, 6) = varAux
                        varAux = .Fields(7)
                        A(.AbsolutePosition - 1, 7) = varAux
                        varAux = .Fields(8)
                        A(.AbsolutePosition - 1, 8) = varAux
                        
                        varAux = gstrENulo(.Fields(10))
                        A(.AbsolutePosition - 1, 9) = varAux
                        varAux = gstrENulo(.Fields(11))
                        A(.AbsolutePosition - 1, 10) = varAux
                        varAux = gstrENulo(.Fields(12))
                        A(.AbsolutePosition - 1, 11) = varAux
                        varAux = gstrENulo(.Fields(13))
                        A(.AbsolutePosition - 1, 12) = varAux
                        varAux = gstrENulo(.Fields(14))
                        A(.AbsolutePosition - 1, 13) = varAux
                        varAux = gstrENulo(.Fields(15))
                        A(.AbsolutePosition - 1, 14) = varAux
                        varAux = gstrENulo(.Fields(16))
                        A(.AbsolutePosition - 1, 15) = varAux
                        varAux = gstrENulo(.Fields(17))
                        A(.AbsolutePosition - 1, 16) = varAux
                        varAux = gstrENulo(.Fields(18))
                        A(.AbsolutePosition - 1, 17) = varAux
                        
                        .MoveNext
                    Loop
                    chkbytEdificado.Value = 1
                Else
                    A.ReDim 0, 0, 0, 17
                    A(0, 0) = ""
                    A(0, 1) = ""
                    A(0, 2) = ""
                    A(0, 3) = ""
                    A(0, 4) = ""
                    A(0, 5) = ""
                    A(0, 6) = ""
                    A(0, 7) = ""
                    A(0, 8) = ""
                    A(0, 9) = ""
                    A(0, 11) = ""
                    A(0, 12) = ""
                    A(0, 13) = ""
                    A(0, 14) = ""
                    A(0, 15) = ""
                    A(0, 16) = ""
                    A(0, 17) = ""
                    
                    chkbytEdificado.Value = 0
                End If
            End With
            txt_TotPredial = gstrConvVrDoSql(dblTotPredial)
            Set grd_Area.Array = A
            grd_Area.ReBind
            grd_Area.Refresh
        End If
    End If
End Sub

Private Sub DeletaValores(intCodImobiliario As Long, Optional lngPkid As Long = 0)
    Dim strsql As String
'Deleta os detalhes do Prédios
    strsql = ""
    strsql = strsql & IIf(EDatabases.Oracle, "Begin ", "")
    
    strsql = strsql & "DELETE FROM " & gstrCaracteristicaDoImovel & " "
    strsql = strsql & "WHERE  INTCODIGOIMOBILIARIO = " & intCodImobiliario
    strsql = strsql & IIf(lngPkid > 0, " AND Intarea = " & lngPkid & "; ", "; ")

'Deleta a área dos Prédios
    strsql = strsql & "DELETE FROM " & gstrAreaImobiliario & " "
    strsql = strsql & "WHERE  intImobiliario = " & intCodImobiliario
    strsql = strsql & IIf(lngPkid > 0, " AND Pkid = " & lngPkid & ";", "; ")
    strsql = strsql & IIf(EDatabases.Oracle, "End; ", "")
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql

    LimpaGrid
End Sub

Private Sub GravaValores(intCodImobiliario As Long, intEdificado As Integer, Optional blnLimparControles As Boolean = True)
    Dim strsql As String
    Dim strMsg As String
    Dim i      As Integer

On Error GoTo err_GravaValores
    If intEdificado = 0 Then 'Não edificado
        Exit Sub
    End If
    
    Set gobjBanco = New clsBanco
    
    grd_Area.MoveFirst

    For i = 0 To A.Count(1) - 1
        If A.Value(i, 0) = "" Then
            strsql = ""
            strsql = "INSERT INTO " & gstrAreaImobiliario & " "
            strsql = strsql & "(intImobiliario, intTipoDeArea, intMedidaDaArea, dblFracaoIdeal, "
            strsql = strsql & " intCategoriaConstrucao,"
            strsql = strsql & " intNPavimento,dtmUltimaReforma, "
            strsql = strsql & " intQuarto , intSala, intCozinha, intBanheiro, intAndar, intElevador, intSuite, intGaragem, intServicoHotelaria,"
            strsql = strsql & " dtmDtAtualizacao, lngCodUsr, intNEdificacao"
            strsql = strsql & ") Values ("
            strsql = strsql & intCodImobiliario & ", "
            strsql = strsql & "'" & A(i, 2) & "', "
            strsql = strsql & gstrConvVrParaSql(A(i, 3)) & ", "
            strsql = strsql & gstrConvVrParaSql(A(i, 4)) & ", "
            strsql = strsql & A(i, 6) & ", "
            strsql = strsql & Val(gstrENulo(A(i, 7), , True)) & ", "
            strsql = strsql & gstrConvDtParaSql(A(i, 8)) & ", "
            strsql = strsql & Val(gstrENulo(A(i, 9))) & ", "
            strsql = strsql & Val(gstrENulo(A(i, 10))) & ", "
            strsql = strsql & Val(gstrENulo(A(i, 11))) & ", "
            strsql = strsql & Val(gstrENulo(A(i, 12))) & ", "
            strsql = strsql & Val(gstrENulo(A(i, 13))) & ", "
            strsql = strsql & Val(gstrENulo(A(i, 14))) & ", "
            strsql = strsql & Val(gstrENulo(A(i, 15))) & ", "
            strsql = strsql & Val(gstrENulo(A(i, 16))) & ", "
            strsql = strsql & Abs(Val(gstrENulo(A(i, 17)))) & ", "
            strsql = strsql & strGETDATE & ", "
            strsql = strsql & glngCodUsr
            
            If A.Value(i, 1) = "" Or A.Value(i, 1) = Null Then
                strsql = strsql & ", " & i + 1
            Else
                strsql = strsql & ", " & A.Value(i, 1)
            End If
            strsql = strsql & ")"
            
            If Not gobjBanco.Execute(strsql) Then
                'gobjBanco.ExecutaRollbackTrans
                Exit Sub
            End If
        Else
            strsql = ""
            strsql = "UPDATE " & gstrAreaImobiliario
            strsql = strsql & " SET intImobiliario = " & intCodImobiliario
            strsql = strsql & " , intTipoDeArea = " & A(i, 2)
            strsql = strsql & " , intMedidaDaArea = " & gstrConvVrParaSql(A(i, 3))
            strsql = strsql & " , dblFracaoIdeal = " & gstrConvVrParaSql(A(i, 4))
            strsql = strsql & " , intCategoriaConstrucao = " & Val(gstrENulo(A(i, 6)))
            strsql = strsql & " , intNPavimento = " & Val(gstrENulo(A(i, 7), , True))
            strsql = strsql & " ,dtmUltimaReforma = " & gstrConvDtParaSql(A(i, 8))
            
            strsql = strsql & " ,intQuarto =" & Val(gstrENulo(A(i, 9)))
            strsql = strsql & " ,intSala =" & Val(gstrENulo(A(i, 10)))
            strsql = strsql & " ,intCozinha =" & Val(gstrENulo(A(i, 11)))
            strsql = strsql & " ,intBanheiro =" & Val(gstrENulo(A(i, 12)))
            strsql = strsql & " ,intAndar =" & Val(gstrENulo(A(i, 13)))
            strsql = strsql & " ,intElevador =" & Val(gstrENulo(A(i, 14)))
            strsql = strsql & " ,intSuite =" & Val(gstrENulo(A(i, 15)))
            strsql = strsql & " ,intGaragem =" & Val(gstrENulo(A(i, 16)))
            strsql = strsql & " ,intServicoHotelaria =" & Val(gstrENulo(A(i, 17)))

            
            strsql = strsql & " ,dtmDtAtualizacao = " & strGETDATE
            strsql = strsql & " , lngCodUsr = " & glngCodUsr
            strsql = strsql & " , intNEdificacao = "
            
            If A.Value(i, 1) = "" Or A.Value(i, 1) = Null Then
                strsql = strsql & UltimoNumeroEdificio(A.Value(i, 0))
            Else
                strsql = strsql & A.Value(i, 1)
            End If

            strsql = strsql & " WHERE Pkid = " & A.Value(i, 0)
            
            If Not gobjBanco.Execute(strsql) Then
                'gobjBanco.ExecutaRollbackTrans
                Exit Sub
            End If
        
        End If
    Next i
    
    'gobjBanco.ExecutaCommitTrans
    
    If blnLimparControles Then LimpaGrid

Exit Sub
err_GravaValores:
    gobjBanco.ExecutaRollbackTrans
End Sub

Private Sub LimpaGridArea()
    Set j = New XArrayDB
    j.Clear

    Set tdd_Area.Array = j
    tdd_Area.ReBind
    tdd_Area.Refresh
End Sub

Private Sub LimpaGrid()
    Set A = New XArrayDB

    A.Clear
    A.ReDim 0, 0, 0, 17

    Set grd_Area.Array = A
    grd_Area.ReBind
    grd_Area.Refresh
End Sub

Private Function GridAreaOK() As Boolean
Dim intLinha As Integer
Dim strMsg As String
    
    strMsg = ""
    grd_Area.Update
    
    If grd_Area.ApproxCount > 0 Then
    
        Set A = grd_Area.Array
        If A.Count(1) = 0 Then
            strMsg = "O imóvel tem que ter no mínimo uma área."
        ElseIf A.Count(1) = 1 Then
            If Trim(A.Value(0, 2)) = "" And Trim(A.Value(0, 3)) = "" And Trim(A.Value(0, 5)) = "" Then
                strMsg = "O imóvel tem que ter no mínimo uma área."
            End If
        End If
        If strMsg = "" Then
            For intLinha = 0 To A.Count(1) - 1
                If Trim(A.Value(intLinha, 2)) = "" Or Trim(A.Value(intLinha, 3)) = "" Or Trim(A.Value(intLinha, 5)) = "" Or IsNull(A.Value(intLinha, 5)) Then
                    strMsg = "Dados incompletos para uma área."
                End If
            Next
        End If
    Else
        strMsg = "O imóvel tem que ter no mínimo uma área."
    End If
    
    If strMsg <> "" Then
        tab_3DPasta.Tab = 6
        grd_Area.SetFocus
        ExibeMensagem strMsg
        Exit Function
    End If
    
    chkbytEdificado.Value = 1
    GridAreaOK = True

End Function

Private Function GridTestadaOK() As Boolean
    Dim intLinha As Integer
    Dim intCol As Integer
    Dim strMsg As String
    Dim strMsgPrincipal As String
    
    strMsg = ""
'    grd_Testada.Update
    
    If grd_Testada.ApproxCount > 0 Then
        
        Set x = grd_Testada.Array
        If x.Count(1) = 0 Then
            strMsg = "O imóvel tem que ter no mínimo uma testada."
        ElseIf x.Count(1) = 1 Then
            'If Trim(X.Value(0, 1)) = "" And Trim(X.Value(0, 2)) = "" And Trim(X.Value(0, 3)) = "" Then
            If Trim(x.Value(0, 1)) = "" And Trim(x.Value(0, 2)) = "" Then
                strMsg = "O imóvel tem que ter no mínimo uma testada."
            End If
        End If
    Else
        strMsg = "O imóvel tem que ter no mínimo uma testada."
    End If
    
    If strMsg = "" Then
       grd_Testada.MoveFirst
       For intLinha = 0 To x.Count(1) - 1
           'If intLinha = 0 Then
           '   If X(intLinha, 5) <> 1 Then
           '      strMsgPrincipal = "A primeira testada deve ser a Principal."
           '      ExibeMensagem strMsgPrincipal
           '      grd_Testada.Refresh
           '      Exit Function
           '   End If
           'End If
           If (x.Value(intLinha, 1) = "" Or x.Value(intLinha, 2) = "" Or _
           x.Value(intLinha, 3) = "") And intLinha <> x.Count(1) Then
              grd_Testada.Row = intLinha
              If grd_Testada.Columns(1) = "" Or grd_Testada.Columns(2) = "" Or _
              grd_Testada.Columns(3) = "" Then
                 strMsg = "Dados incompletos para uma testada."
                 Exit For
              End If
           End If
       Next
    End If
    
    If strMsg <> "" Then
        tab_3DPasta.Tab = 5
        grd_Testada.SetFocus
        ExibeMensagem strMsg
        grd_Testada.Refresh
        Exit Function
    End If
    GridTestadaOK = True
End Function

Private Function blnDadosOK(strOperacao As String, Optional blnIncluindoCaracteristica As Boolean = False) As Boolean
    
If UCase(strOperacao) = "SALVAR" Then
    If tab_3DPasta.Tab = 9 Then 'Cláudio - Doc/Proc'
        If Not dbc_intTiposDocumentosProcesso.MatchedWithList Then
            ExibeMensagem "Selecione um Documento/Processo válido."
            dbc_intTiposDocumentosProcesso.SetFocus
            Exit Function
        End If
         
        If Len(Trim(txt_strCodigo)) > 0 Or Len(Trim(txt_bitDigito)) > 0 Or Len(Trim(txt_intExercicio)) > 0 Then
            If Not VerificaEmpenhoProcesso(txt_strCodigo, txt_bitDigito, txt_intExercicio) Then
                ExibeMensagem "Processo não localizado."
                If txt_strCodigo.Enabled Then txt_strCodigo.SetFocus
                Exit Function
            End If
        End If
    Else
        If mskstrInscricao.ClipText = "" Then
            MsgBox "O número da Inscrição Cadastral tem que ser digitado."
            tab_3DPasta.Tab = 0
            mskstrInscricao.SetFocus
            blnDadosOK = False
            Exit Function
        End If
        
        If Len(Trim(mskstrInscricao.ClipText)) < bytTamanhoMascara Then
            ExibeMensagem "Inscrição Cadastral inválido."
            tab_3DPasta.Tab = 0
            mskstrInscricao.SetFocus
            blnDadosOK = False
            Exit Function
        End If
        If txtstrEmissao.Text = "" Then
            ExibeMensagem "É necessário informar uma Emissão."
            tab_3DPasta.Tab = 0
            txtstrEmissao.SetFocus
            blnDadosOK = False
            Exit Function
        End If
        If dbcintLogradouro.BoundText = "" Then
            MsgBox "O Logradouro do endereço imobiliário tem que ser selecionado."
            tab_3DPasta.Tab = 0
            dbcintLogradouro.SetFocus
            blnDadosOK = False
            Exit Function
        End If
        
        If Not mblnAlterando Then
            If blnVerificaLogradouro(dbcintLogradouro.BoundText) Then
                tab_3DPasta.Tab = 0
                MsgBox "O logradouro selecionado está cancelado."
                dbcintLogradouro.SetFocus
                Exit Function
            End If
        End If
        
        If Len(Trim(txtintNumero.Text)) = 0 Then
            MsgBox "O Número do Logradouro do endereço imobiliário tem que ser digitado."
            tab_3DPasta.Tab = 0
            txtintNumero.SetFocus
            blnDadosOK = False
            Exit Function
        End If
        If Len(Trim(txtintCep.Text)) = 0 Then
            MsgBox "O CEP do Logradouro do endereço imobiliário tem que ser digitado."
            tab_3DPasta.Tab = 0
            txtintCep.SetFocus
            blnDadosOK = False
            Exit Function
        End If
        If txtdblArea.Text = "" Then
            MsgBox "A Área do Terreno tem que ser digitada."
            tab_3DPasta.Tab = 0
            txtdblArea.SetFocus
            blnDadosOK = False
            Exit Function
        End If
        If dbcintContribuinte.BoundText = "" Then
            MsgBox "O contribuinte tem que ser selecionado."
            tab_3DPasta.Tab = 0
            dbcintContribuinte.SetFocus
            blnDadosOK = False
            Exit Function
        End If
        If (mskstrInscricao.Text <> mskstrInscricao.Tag) Then
            If gblnExisteCodigo(1, gstrImobiliario, "strInscricaoAnterior", "'" & String(gintLenInscricao - Len(mskstrInscricao.Text), "0") & mskstrInscricao.Text & "'") Then
                tab_3DPasta.Tab = 0
                ExibeMensagem "A inscrição cadastral digitada já se encontra cadastrada!"
                mskstrInscricao.SetFocus
                Exit Function
            End If
        End If
        
        If Not GridTestadaOK Then
            Exit Function
        End If
        
        'If Not GridAreaOK Then
        '    Exit Function
        'End If
            
        If blnIncluindoCaracteristica = False Then
            If mblnAlterando = False And blnEmTransacao = False Then
                tab_3DPasta.Tab = 5
                grd_Testada.SetFocus
                ExibeMensagem "É necessário selecionar detalhes para as características."
                Exit Function
            Else
                If blnDetalhesDaCaracterisiticas(2) = False Then 'Terreno
                    Exit Function
                End If
            '    If blnDetalhesDaCaracterisiticas(3) = False Then 'Construção
            '       Exit Function
            '    End If
            End If
            
            If Not VerificaEdificacao Then Exit Function
            
        End If
        
        If Not blnVerificaFaceDeQuadra Then Exit Function
    End If
End If

blnDadosOK = True

End Function

Private Sub lvw_Historico_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_Historico
        txt_Historico = .SelectedItem.Text
    End With
End Sub

Private Sub GravaHistoricos(intCodImobiliario As Long, Optional blnLimparControles As Boolean = True)
    Dim strsql As String
    Dim intI   As Integer

    On Error GoTo err_GravaHistoricos
    strsql = ""
    strsql = strsql & "Delete From " & gstrHistoricoImobiliario & " "
    strsql = strsql & "Where intImobiliario = " & intCodImobiliario

    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql

    Set gobjBanco = New clsBanco
    'gobjBanco.ExecutaBeginTrans

    With lvw_Historico
        For intI = 1 To .ListItems.Count
            strsql = ""
            strsql = strsql & "Insert Into " & gstrHistoricoImobiliario & " "
            strsql = strsql & "(intImobiliario, strDescricao "
            strsql = strsql & ") Values ("
            strsql = strsql & intCodImobiliario & ",'"
            strsql = strsql & .ListItems(intI).Text & "' "
            strsql = strsql & ")"
            Set gobjBanco = New clsBanco
            If Not gobjBanco.Execute(strsql) Then
                gobjBanco.ExecutaRollbackTrans
                Exit Sub
            End If
        Next
    End With
    
    If blnLimparControles Then
        lvw_Historico.ListItems.Clear
        txt_Historico = ""
    End If

    'gobjBanco.ExecutaCommitTrans
    
    Exit Sub
err_GravaHistoricos:
        gobjBanco.ExecutaRollbackTrans
        
End Sub

Private Sub DeletaHistoricos(intCodImobiliario As Long)
    Dim strsql As String
    
    strsql = ""
    strsql = strsql & "Delete From " & gstrHistoricoImobiliario & " "
    strsql = strsql & "Where intImobiliario = " & intCodImobiliario

    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql
End Sub

Private Function CarregaHistoricos(intCodImobiliario As Long)
    Dim strsql       As String
    Dim adoResultado As ADODB.Recordset

    lvw_Historico.ListItems.Clear
    txt_Historico = ""

    strsql = ""
    strsql = strsql & "Select HI.strDescricao Historico "
    strsql = strsql & "From " & gstrHistoricoImobiliario & " HI "
    strsql = strsql & "Where HI.intImobiliario = " & txtPKId 'Mudei
    'CarregaHistoricos =
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
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
End Function

'Private Sub txtdblValorEdificacao_LostFocus()
'dblEdificacao = 0
'dblTerreno = 0
'    If txtdblValorEdificacao = "Null" Then
'       txtdblValorEdificacao = 0
'    End If
'    If txtdblValorTerreno = "Null" Then
'       txtdblValorTerreno = 0
'    End If
'    If txtdblValorEdificacao = "" Then
'        dblEdificacao = 0
'        Else
'        dblEdificacao = CDbl(txtdblValorEdificacao)
'    End If
'    If txtdblValorTerreno = "" Then
'        dblTerreno = 0
'        Else
'        dblTerreno = CDbl(txtdblValorTerreno)
'    End If
'    txtdblValorImovel = (dblEdificacao) + (dblTerreno)
'    txtdblValorEdificacao = gvntConvVrDoSql(txtdblValorEdificacao)
'End Sub

'Private Sub txtdblValorImovel_Change()
'    If txtdblValorImovel = "" Then
'        Exit Sub
'    End If
'    txtdblValorImovel = gvntConvVrDoSql(txtdblValorImovel)
'End Sub

'Private Sub txtdblValorITBI_Change()
'    If txtdblValorITBI = "" Then
'        Exit Sub
'    End If
'    txtdblValorITBI = gvntConvVrDoSql(txtdblValorITBI)
'End Sub

'Private Sub txtdblValorTerreno_LostFocus()
'dblEdificacao = 0
'dblTerreno = 0
    'If txtdblValorEdificacao = "Null" Then
    '   txtdblValorEdificacao = 0
    'End If
 '   If txtdblValorTerreno = "Null" Then
 '      txtdblValorTerreno = 0
 '   End If
    'If txtdblValorEdificacao = "" Then
    '    dblEdificacao = 0
    '    Else
    '    dblEdificacao = CDbl(txtdblValorEdificacao)
    'End If
  '  If txtdblValorTerreno = "" Then
  '      dblTerreno = 0
  '      Else
  '      dblTerreno = CDbl(txtdblValorTerreno)
  '  End If
  '  txtdblValorImovel = (dblEdificacao) + (dblTerreno)
  '  txtdblValorTerreno = gvntConvVrDoSql(txtdblValorTerreno)
'End Sub


'Private Sub txtdblValorTerreno_GotFocus()
'    MarcaCampo txtdblValorTerreno
'End Sub

'Private Sub txtdblValorTerreno_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "V", txtdblValorTerreno
'End Sub

'Private Sub txtdblValorEdificacao_GotFocus()
'    MarcaCampo txtdblValorEdificacao
'End Sub

'Private Sub txtdblValorEdificacao_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "V", txtdblValorEdificacao
'End Sub

'Private Sub txtdblValorITBI_GotFocus()
'    MarcaCampo txtdblValorITBI
'End Sub

'Private Sub txtdblValorITBI_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "V", txtdblValorITBI
'End Sub

'Private Sub txtdblValorImovel_LostFocus()
'    txtdblValorImovel = gvntConvVrDoSql(txtdblValorImovel)
'End Sub


'Tab EQUIPAMENTOS

'no clique do  combo  CarregaMelhoria
Sub CarregaMelhoria()
Dim strsql As String
Dim i As Integer

    strsql = ""
    strsql = strsql & "Select MP.PKId as PKID3, MP.strNomeDoMelhoramento AS Melhoramento "
    strsql = strsql & " From " & gstrMelhoramentoPublico & " MP,"
    strsql = strsql & gstrMelhoramentoDaSecaoDeLogradouro & " MS "
    strsql = strsql & " Where MP.PKId = MS.intMelhoramento And MS.intFaceDeQuadra = '" & dbc_intFaceDeQuadra.BoundText & "'"

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            lvw_Melhoria.ListItems.Clear
            Do While .EOF = False
                Set objList1 = lvw_Melhoria.ListItems.Add(, , !Melhoramento)
                objList1.Tag = !Pkid3
                .MoveNext
            Loop
        End With
        For i = 1 To lvw_Melhoria.ListItems.Count
            If MelhoriaChecada(lvw_Melhoria.ListItems(i).Tag) Then
                lvw_Melhoria.ListItems(i).Checked = True
                lvw_Melhoria.ListItems(i).Selected = True
            Else
                lvw_Melhoria.ListItems(i).Checked = False
                lvw_Melhoria.ListItems(i).Selected = False
            End If
        Next
    End If
End Sub

Private Function strQueryFaceDeQuadra() As String

Dim strsql       As String
    
    strsql = "SELECT FQ.Pkid, LTRIM(RTRIM(FQ.strSetor))" & strCONCAT
    strsql = strsql & "'.'" & strCONCAT & " LTRIM(RTRIM(FQ.strQuadra))" & strCONCAT
    strsql = strsql & "'.'" & strCONCAT & " LTRIM(RTRIM(FQ.strSequenciaDeFace))" & strCONCAT
    strsql = strsql & "' - '" & strCONCAT & gstrISNULL("TL.strSigla", "''") & strCONCAT & "' '"
    strsql = strsql & strCONCAT & gstrISNULL("TTL.strSigla", "''") & strCONCAT & "''" & strCONCAT & "' '"
    'strSql = strSql & " LTRIM(RTrim(LO.strDescricao)) As strFaceDeQuadra"
    strsql = strsql & strCONCAT & " LTRIM(RTrim(LO.strDescricao)) " & strCONCAT
    strsql = strsql & " ' - ' " & strCONCAT & " LTRIM(RTrim(BA.strDescricao)) As strFaceDeQuadra "
    strsql = strsql & " FROM "
    strsql = strsql & gstrFaceDeQuadra & " FQ, "
    strsql = strsql & gstrTipoLogradouro & " TL, "
    strsql = strsql & gstrTituloLogradouro & " TTL, "
    strsql = strsql & gstrBairro & " BA, "
    'strSql = strSql & gstrLogradouro & " LO, "
    strsql = strsql & gstrLogradouro & " LO "
    'strSql = strSql & gstrTestadaImobiliario & " TI"
    strsql = strsql & " WHERE FQ.intLogradouro = LO.Pkid AND"
    strsql = strsql & " LO.intTipoLogradouro " & strOUTJSQLServer & "=" & " TL.Pkid " & strOUTJOracle & " AND"
    strsql = strsql & " LO.intTituloLogradouro " & strOUTJSQLServer & "=" & " TTL.Pkid " & strOUTJOracle
    strsql = strsql & " AND LO.intbairro " & strOUTJSQLServer & "=" & " BA.Pkid " & strOUTJOracle
    'strSql = strSql & " AND TI.intFaceDeQuadra = FQ.Pkid"
    
    If Len(Trim$(txtstrSequenciaDeFace.Text)) > 0 Then
        strsql = strsql & " AND FQ.strSequenciaDeFace = '" & txtstrSequenciaDeFace.Text & "'"
    End If
    
    strsql = strsql & " AND LO.Dtmdtexclusao is null "
    strsql = strsql & " ORDER BY FQ.strsetor, FQ.strquadra, FQ.strsequenciadeface, LO.strDescricao"
    
    strQueryFaceDeQuadra = strsql
    
End Function

Private Sub GravaEquipamento(intCodImobiliario As Long)
Dim i      As Integer
Dim strsql As String
    
    strsql = ""
    strsql = "Delete From " & gstrEquipamentoImobiliario
    strsql = strsql & " Where intImobiliario = '" & txtPKId & "'"
    strsql = strsql & " AND intFaceDeQuadra = '" & dbc_intFaceDeQuadra.BoundText & "'"
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql

    For i = 1 To lvw_Melhoria.ListItems.Count
        strsql = ""
        If lvw_Melhoria.ListItems(i).Checked = True Then
            strsql = "Insert Into " & gstrEquipamentoImobiliario & " (intImobiliario, intMelhoria, intFaceDeQuadra) Values ('" & txtPKId & "', "
            strsql = strsql & lvw_Melhoria.ListItems(i).Tag & "," & "'" & dbc_intFaceDeQuadra.BoundText & "'" & ")"
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strsql
        End If
    Next
    'lvw_Melhoria.ListItems.Clear
End Sub

Private Sub DeletaEquipamento(intCodImobiliario As Long)
Dim strsql As String
    
    strsql = ""
    strsql = "Delete From " & gstrEquipamentoImobiliario _
             & " Where intImobiliario = " & intCodImobiliario
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql
    lvw_Melhoria.ListItems.Clear
End Sub

Function MarcaEquipamento(intTag As Integer)
Dim i As Integer
    For i = 1 To lvw_Melhoria.ListItems.Count
        If lvw_Melhoria.ListItems(i).Tag = intTag Then
            lvw_Melhoria.ListItems(i).Checked = True
        End If
    Next
End Function

Function SelecionaEquipamento(intCodImobiliario As Long)
Dim strsql As String
    strsql = ""
    strsql = "Select * From " & gstrEquipamentoImobiliario & " " _
             & "Where intImobiliario = " & intCodImobiliario
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                Call MarcaEquipamento(!intMelhoria)
                .MoveNext
            Loop
        End With
    End If
End Function

'GRID TESTADA
Function PreencheGRD2()

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
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strsql As String
    LimpaGrid2

        strsql = ""
        strsql = strsql & "SELECT TI.Pkid, TI.intTipoDeTestada, TI.strMedidaDaTestada, "
        strsql = strsql & " FQ.Pkid PkidFaceDeQuadra, "
        strsql = strsql & " LTRIM(RTRIM(FQ.strSetor))" & strCONCAT
        strsql = strsql & "'.'" & strCONCAT & " LTRIM(RTRIM(FQ.strQuadra))" & strCONCAT
        strsql = strsql & "'.'" & strCONCAT & " LTRIM(RTRIM(FQ.strSequenciaDeFace))" & strCONCAT
        strsql = strsql & "' - '" & strCONCAT & gstrISNULL("TL.strSigla", "''") & strCONCAT & "''"
        strsql = strsql & strCONCAT & gstrISNULL("TTL.strSigla", "''") & strCONCAT & "''" & strCONCAT
        strsql = strsql & " LTRIM(RTrim(LO.strDescricao))" & strCONCAT & "' - '" & strCONCAT & " LTRIM(RTrim(BA.strDescricao)) " & strCONCAT & " '('" & strCONCAT & gstrISNULL(gstrCONVERT(CDT_VARCHAR, "VT.dblValor"), "''") & strCONCAT & "')' As strFaceDeQuadra,"
        strsql = strsql & " TT.bytPrincipal, TT.strNomeDaTestada "
                
        If bytDBType = Oracle Then
            strsql = strsql & " FROM " & gstrTestadaImobiliario & " TI, "
            strsql = strsql & gstrFaceDeQuadra & " FQ, "
            strsql = strsql & gstrHistoricoFaceDeQuadra & " HFQ, "
            strsql = strsql & gstrValorMetroTerreno & " VT, "
            strsql = strsql & gstrTipoLogradouro & " TL, "
            strsql = strsql & gstrTituloLogradouro & " TTL, "
            strsql = strsql & gstrLogradouro & " LO, "
            strsql = strsql & gstrTipoDeTestada & " TT, "
            strsql = strsql & gstrBairro & " BA"
            strsql = strsql & " WHERE FQ.intLogradouro = LO.Pkid AND"
            strsql = strsql & " FQ.PKID " & strOUTJSQLServer & "=" & " HFQ.INTFACEDEQUADRA " & strOUTJOracle & " AND"
            strsql = strsql & " HFQ.INTVALORMETROTERRENO " & strOUTJSQLServer & "=" & " VT.PKID " & strOUTJOracle & " AND"
            strsql = strsql & " LO.intTipoLogradouro " & strOUTJSQLServer & "=" & " TL.Pkid " & strOUTJOracle & " AND"
            strsql = strsql & " LO.intTituloLogradouro " & strOUTJSQLServer & "=" & " TTL.Pkid " & strOUTJOracle
            strsql = strsql & " AND LO.intBairro " & strOUTJSQLServer & "=" & " BA.Pkid " & strOUTJOracle
            strsql = strsql & " AND TI.intImobiliario = '" & Val(txtPKId) & "'" 'Mudei"
            strsql = strsql & " AND VT.intExercicio = " & gintExercicio
    '        strSql = strSql & " AND LO.Dtmdtexclusao is null "
            strsql = strsql & " AND FQ.Pkid = TI.intFaceDeQuadra"
            strsql = strsql & " AND TI.intTipoDeTestada = TT.Pkid"
        Else
            strsql = strsql & " FROM " & gstrTipoDeTestada & " TT INNER JOIN "
            strsql = strsql & gstrTestadaImobiliario & " TI ON TT.PKId = TI.intTipoDeTestada LEFT OUTER JOIN "
            strsql = strsql & gstrLogradouro & " LO INNER JOIN "
            strsql = strsql & gstrFaceDeQuadra & " FQ ON LO.PKId = FQ.intLogradouro LEFT OUTER JOIN "
            strsql = strsql & gstrTipoLogradouro & " TL ON LO.intTipoLogradouro = TL.PKId LEFT OUTER JOIN "
            strsql = strsql & gstrTituloLogradouro & " TTL ON LO.intTituloLogradouro = TTL.PKId LEFT OUTER JOIN "
            strsql = strsql & gstrBairro & " BA ON LO.intBairro = BA.PKId ON TI.intFaceDeQuadra = FQ.PKId LEFT OUTER JOIN "
            strsql = strsql & gstrHistoricoFaceDeQuadra & " HFQ ON HFQ.INTFACEDEQUADRA = FQ.PKID LEFT OUTER JOIN "
            strsql = strsql & gstrValorMetroTerreno & " VT ON VT.PKID = HFQ.INTVALORMETROTERRENO"
            strsql = strsql & " WHERE TI.intImobiliario = " & Val(txtPKId)
            strsql = strsql & " AND VT.intExercicio = " & gintExercicio
        End If
        
        strsql = strsql & " Order By TT.bytPrincipal Desc"
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strsql, 5, adoRec) Then
            MontaArray2
        End If

        'PreencheFaceDeQuadra
        
End Function

Private Sub PreencheFaceDeQuadra()
Dim strsql As String
Dim i As Byte
Dim strSetor As String
Dim strQuadra As String
  For i = 1 To 24
    If Mid(Trim(mskstrInscricao.FormattedText), i, 1) <> "." Then
       strSetor = Trim(strSetor & Mid(mskstrInscricao.FormattedText, i, 1))
    Else
       Exit For
    End If
  Next
  
  For i = i + 1 To 24
    If Mid(Trim(mskstrInscricao.FormattedText), i, 1) <> "." Then
       strQuadra = Trim(strQuadra & Mid(mskstrInscricao.FormattedText, i, 1))
    Else
       Exit For
    End If
  Next
  
  strsql = "SELECT FQ.Pkid PkidFaceDeQuadra, LTRIM(RTRIM(FQ.strSetor))" & strCONCAT
  strsql = strsql & "'.'" & strCONCAT & " LTRIM(RTRIM(FQ.strQuadra))" & strCONCAT
  strsql = strsql & "'.'" & strCONCAT & " LTRIM(RTRIM(FQ.strSequenciaDeFace))" & strCONCAT
  strsql = strsql & "' - '" & strCONCAT & gstrISNULL("TL.strSigla", "''") & strCONCAT & "''"
  strsql = strsql & strCONCAT & gstrISNULL("TTL.strSigla", "''") & strCONCAT & "''" & strCONCAT
  strsql = strsql & " LTRIM(RTrim(LO.strDescricao)) " & strCONCAT & " '('" & strCONCAT & gstrISNULL(gstrCONVERT(CDT_VARCHAR, "VT.dblValor"), "''") & strCONCAT & "')' As strFaceDeQuadra"
  strsql = strsql & " FROM "
  
  If bytDBType = Oracle Then
    strsql = strsql & gstrFaceDeQuadra & " FQ, "
    strsql = strsql & gstrHistoricoFaceDeQuadra & " HFQ, "
    strsql = strsql & gstrValorMetroTerreno & " VT, "
    strsql = strsql & gstrTipoLogradouro & " TL, "
    strsql = strsql & gstrTituloLogradouro & " TTL, "
    strsql = strsql & gstrLogradouro & " LO "
    strsql = strsql & " WHERE FQ.intLogradouro = LO.Pkid AND "
    strsql = strsql & " LO.intTipoLogradouro " & strOUTJSQLServer & "=" & " TL.Pkid " & strOUTJOracle & " AND "
    strsql = strsql & " LO.intTituloLogradouro " & strOUTJSQLServer & "=" & " TTL.Pkid " & strOUTJOracle & " AND "
    strsql = strsql & " FQ.PKID " & strOUTJSQLServer & "=" & " HFQ.INTFACEDEQUADRA " & strOUTJOracle & " AND"
    strsql = strsql & " HFQ.INTVALORMETROTERRENO " & strOUTJSQLServer & "=" & " VT.PKID " & strOUTJOracle & " AND"
  Else
    strsql = strsql & gstrFaceDeQuadra & " FQ LEFT JOIN "
    strsql = strsql & gstrHistoricoFaceDeQuadra & " HFQ ON FQ.PKID = HFQ.INTFACEDEQUADRA LEFT JOIN "
    strsql = strsql & gstrValorMetroTerreno & " VT ON HFQ.INTVALORMETROTERRENO = VT.PKID INNER JOIN "
    strsql = strsql & gstrLogradouro & " LO ON FQ.intLogradouro = LO.Pkid LEFT JOIN "
    strsql = strsql & gstrTipoLogradouro & " TL ON LO.intTipoLogradouro = TL.Pkid LEFT JOIN "
    strsql = strsql & gstrTituloLogradouro & " TTL ON LO.intTituloLogradouro = TTL.Pkid "
    strsql = strsql & " WHERE "
  End If
  
  strsql = strsql & " LO.Dtmdtexclusao is null"
  strsql = strsql & " AND VT.intExercicio = " & gintExercicio
  strsql = strsql & " AND FQ.strSetor='" & strSetor & "'"
  strsql = strsql & " AND FQ.strQuadra='" & strQuadra & "'"
  strsql = strsql & " ORDER BY LO.strDescricao"

  LeDaTabelaParaObj "", tdd_FaceDeQuadra, strsql

'Set gobjBanco = New clsBanco
'If Not gobjBanco.CriaADO(strSql, 5, adoTdb) Then
'    Exit Sub
'End If
        
'Set Z = New XArrayDB
'If Not adoTdb.EOF Then
    
'    Z.ReDim 0, adoTdb.RecordCount - 1, 0, 1
'    Dim varAux2 As Variant
'    Do While Not adoTdb.EOF
'        varAux2 = adoTdb!strFaceDeQuadra
'        Z(adoTdb.AbsolutePosition - 1, 0) = varAux2
'
'        varAux2 = adoTdb!PkidFaceDeQuadra
'        Z(adoTdb.AbsolutePosition - 1, 1) = varAux2
'
'        adoTdb.MoveNext
'    Loop
'Else
'    Z.ReDim 0, 0, 0, 1
'    Z(0, 0) = ""
'    Z(0, 1) = ""
'End If
'
'Set tdd_FaceDeQuadra.Array = Z
'tdd_FaceDeQuadra.ReBind
'tdd_FaceDeQuadra.Refresh

End Sub

Private Function PreencheTestada(blnPrincipal As Boolean)
Dim strsql As String

    strsql = ""
    strsql = strsql & "SELECT PKId, strNomeDaTestada, bytPrincipal "
    strsql = strsql & "FROM " & gstrTipoDeTestada & " "
    
    If blnPrincipal = True Then
        strsql = strsql & "WHERE bytPrincipal = 1"
    Else
        strsql = strsql & "WHERE bytPrincipal = 0"
    End If
        
    LeDaTabelaParaObj "", tdd_Testada, strsql
    
    DoEvents
    
End Function

Private Sub grd_Testada_KeyPress(KeyAscii As Integer)
Select Case grd_Testada.Col
    Case 1, 3
        CaracterValido KeyAscii, "A", grd_Testada
    Case 2
        CaracterValido KeyAscii, "V", grd_Testada
End Select
End Sub

Private Sub tdd_FaceDeQuadra_DropDownClose()
    Dim intRow As Integer
    On Error GoTo Err_Handle
    If Not IsNull(tdd_FaceDeQuadra.SelectedItem) Then
        grd_Testada.Columns(3) = tdd_FaceDeQuadra.Columns(0)
        grd_Testada.Columns(4) = tdd_FaceDeQuadra.Columns(1)
    Else
        grd_Testada.Columns(3) = ""
        grd_Testada.Columns(4) = 0
    End If
    Exit Sub
Err_Handle:
End Sub


Private Sub MontaArray2()

    Dim varAux As Variant

    Set x = New XArrayDB
    x.Clear
    With adoRec
        
        If Not .EOF And mblnAlterando Then
            
            x.ReDim 0, .RecordCount - 1, 0, 6
            Do While Not .EOF
                If !bytPrincipal = 1 Then
                    dblValorProfundidade = CDbl(gstrENulo(!strMedidaDaTestada))
                End If
                varAux = .Fields(0)
                x(.AbsolutePosition - 1, 0) = varAux
                varAux = .Fields(6)
                x(.AbsolutePosition - 1, 1) = varAux
                varAux = gstrConvVrDoSql(.Fields(2), 2)
                x(.AbsolutePosition - 1, 2) = varAux
                varAux = .Fields(4)
                x(.AbsolutePosition - 1, 3) = varAux
                varAux = .Fields(3)
                x(.AbsolutePosition - 1, 4) = varAux
                varAux = .Fields(5)
                x(.AbsolutePosition - 1, 5) = varAux
                varAux = .Fields(1)
                x(.AbsolutePosition - 1, 6) = varAux
                .MoveNext
            Loop
        Else
            x.ReDim 0, 0, 0, 6
            x(0, 0) = ""
            x(0, 1) = ""
            x(0, 2) = ""
            x(0, 3) = ""
            x(0, 4) = ""
            x(0, 5) = ""
            x(0, 6) = ""

        End If
    End With

    Set grd_Testada.Array = x
    grd_Testada.ReBind
    grd_Testada.Refresh
End Sub

Private Sub DeletaValores2(intCodImobiliario As Long, Optional lngPkid As Long = 0)
    Dim strsql As String
    
    strsql = ""
    strsql = strsql & "DELETE FROM " & gstrTestadaImobiliario & " "
    strsql = strsql & "WHERE  intImobiliario = " & intCodImobiliario
    strsql = strsql & IIf(lngPkid > 0, " AND Pkid = " & lngPkid, "")
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql

    LimpaGrid2
    
End Sub

Private Sub GravaValores2(intCodImobiliario As Long, Optional blnLimparControles As Boolean = True)

'******************************************************************************************
' Data: 04/04/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'        strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strsql     As String
    Dim strMsg     As String
    Dim i          As Integer

    On Error GoTo err_GravaValores2
    
    Set gobjBanco = New clsBanco
    'gobjBanco.ExecutaBeginTrans
    strsql = ""
    strsql = strsql & "DELETE FROM " & gstrTestadaImobiliario & " "
    strsql = strsql & "WHERE  intImobiliario = '" & txtPKId & "'"
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql

    grd_Testada.MoveFirst
    
    For i = 0 To x.Count(1) - 1
        
        If x.Value(i, 1) <> Space$(0) Then
        
            strsql = ""
            strsql = strsql & "INSERT INTO " & gstrTestadaImobiliario
            strsql = strsql & " (intImobiliario, intTipoDeTestada, strMedidaDaTestada, "
            strsql = strsql & " intFaceDeQuadra, "
            strsql = strsql & " dtmDtAtualizacao, lngCodUsr "
            strsql = strsql & ") Values ("
            strsql = strsql & intCodImobiliario & ", "
            strsql = strsql & "'" & x(i, 6) & "', "
            strsql = strsql & "'" & IIf(bytDBType = SQLServer, gstrConvVrParaSql(x(i, 2)), x(i, 2)) & "', "
            strsql = strsql & "'" & x(i, 4) & "', "
            strsql = strsql & strGETDATE & ", "
            strsql = strsql & glngCodUsr
            strsql = strsql & ")"

            If Not gobjBanco.Execute(strsql, False) Then
                'gobjBanco.ExecutaRollbackTrans
                Exit Sub
            End If
        
        End If
    Next i
    
    'gobjBanco.ExecutaCommitTrans
    
    If blnLimparControles Then LimpaGrid2

Exit Sub
err_GravaValores2:
    gobjBanco.ExecutaRollbackTrans
End Sub

Private Sub LimpaGrid2()
    Set x = New XArrayDB
    Set Z = New XArrayDB

    x.Clear
    x.ReDim 0, 0, 0, 6
    Z.Clear

    Set grd_Testada.Array = x
    grd_Testada.ReBind
    grd_Testada.Refresh
    
    Set tdd_FaceDeQuadra.Array = Z
    tdd_FaceDeQuadra.ReBind
    tdd_FaceDeQuadra.Refresh

End Sub

Sub strLimpaContribuintesTabs(Optional blpLimpaGrids As Boolean = True)
    txt_Inscricao = ""
    txt_Proprietario = ""
    txt_Inscricao2 = ""
    txt_Proprietario2 = ""
    txt_Inscricao3 = ""
    txt_Proprietario3 = ""
    txt_Inscricao4 = ""
    txt_Proprietario4 = ""
    txt_Inscricao5 = ""
    txt_Proprietario5 = ""
    txt_Inscricao6 = ""
    txt_Proprietario6 = ""
    txt_Inscricao7 = ""
    txt_Proprietario7 = ""
    txt_Inscricao8 = ""
    txt_Proprietario8 = ""
    txt_Bairro = ""
    txt_Cep = ""
    txt_Complemento = ""
    txt_Distrito = ""
    txt_Logradouro = ""
    txt_Municipio = ""
    txt_Numero = ""
    txt_UF = ""
    txt_PKIdContribuinte = ""
    txt_strCNPJCPF.Text = ""
    If blpLimpaGrids Then
       LimpaGrid
       'PreencheFaceDeQuadra
    End If
    
End Sub

Sub strPreencheContribuinteTabs()
    txt_Inscricao = gstrFormataInscricao(gstrVerificaCampoNulo(mskstrInscricao.ClipText), TYP_IMOBILIARIA)
    txt_Proprietario = gstrVerificaCampoNulo(dbcintContribuinte.Text)
    txt_Inscricao2 = gstrFormataInscricao(gstrVerificaCampoNulo(mskstrInscricao.ClipText), TYP_IMOBILIARIA)
    txt_Proprietario2 = gstrVerificaCampoNulo(dbcintContribuinte.Text)
    txt_Inscricao3 = gstrFormataInscricao(gstrVerificaCampoNulo(mskstrInscricao.ClipText), TYP_IMOBILIARIA)
    txt_Proprietario3 = gstrVerificaCampoNulo(dbcintContribuinte.Text)
    txt_Inscricao4 = gstrFormataInscricao(gstrVerificaCampoNulo(mskstrInscricao.ClipText), TYP_IMOBILIARIA)
    txt_Proprietario4 = gstrVerificaCampoNulo(dbcintContribuinte.Text)
    txt_Inscricao5 = gstrFormataInscricao(gstrVerificaCampoNulo(mskstrInscricao.ClipText), TYP_IMOBILIARIA)
    txt_Proprietario5 = gstrVerificaCampoNulo(dbcintContribuinte.Text)
    txt_Inscricao6 = gstrFormataInscricao(gstrVerificaCampoNulo(mskstrInscricao.ClipText), TYP_IMOBILIARIA)
    txt_Proprietario6 = gstrVerificaCampoNulo(dbcintContribuinte.Text)
    txt_Inscricao7 = gstrFormataInscricao(gstrVerificaCampoNulo(mskstrInscricao.ClipText), TYP_IMOBILIARIA)
    txt_Proprietario7 = gstrVerificaCampoNulo(dbcintContribuinte.Text)
    txt_Inscricao8 = gstrFormataInscricao(gstrVerificaCampoNulo(mskstrInscricao.ClipText), TYP_IMOBILIARIA)
    txt_Proprietario8 = gstrVerificaCampoNulo(dbcintContribuinte.Text)
End Sub

Private Sub txtstrQuadra_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrQuadra
End Sub

Private Sub LocalizarImobiliario()
Dim strsql As String
Dim strCondicao As String
Dim strValor As String
Dim strCampo As String
Dim i As Integer

    strCondicao = ""

    With Me
    
    For i = 0 To .Controls.Count - 1
        
        
        If Not TypeOf .Controls(i) Is Label Then 'Elimina os Label's da pesquisa
            'Elimina objetos indesejáveis
            
            If UCase(.Controls(i).Name) <> "TXTPKId" _
            And UCase(Left(.Controls(i).Name, 3)) <> "IMG" _
            And UCase(Left(.Controls(i).Name, 3)) <> "LVW" _
            And UCase(Left(.Controls(i).Name, 3)) <> "TLB" _
            And UCase(Left(.Controls(i).Name, 3)) <> "TDD" _
            And UCase(Left(.Controls(i).Name, 3)) <> "GRD" _
            And UCase(Left(.Controls(i).Name, 3)) <> "ACR" _
            And UCase(Left(.Controls(i).Name, 3)) <> "CHK" _
            And UCase(.Controls(i).Name) <> UCase("txt_PKIdContribuinte") _
            And UCase(.Controls(i).Name) <> UCase("txt_strCNPJCPF") _
            And InStr(1, UCase(.Controls(i).Name), UCase("_Inscricao")) = 0 _
            And InStr(1, UCase(.Controls(i).Name), UCase("_Proprietario")) = 0 _
            And UCase(.Controls(i).Name) <> UCase("txt_strCNPJCPFP") _
            And InStr(1, UCase(.Controls(i).Name), UCase("txt_")) = 0 _
            And UCase(.Controls(i).Name) <> UCase("optbytNaturezaJuridica") Then
            
                If Not (TypeOf .Controls(i) Is OptionButton) Or .Controls(i) = True Then 'Elimina OptionButton desmarcado
                    If TypeOf .Controls(i) Is TextBox Then
                        If Trim(.Controls(i).Text) <> "" Then
                        
                            If InStr(1, .Controls(i).Name, "Cep") > 0 Then
                                strValor = gstrValorSemMascara(Trim(.Controls(i).Text))
                            Else
                                strValor = Trim(.Controls(i).Text)
                            End If
                            
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, "_", strCampo) > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            strCampo = "IM." & strCampo
                            
                            If strCampo = "IM.strEmissao" Then strValor = String(gintLenEmissao - Len(strValor), "0") & strValor
                            
                            If InStr(1, "%", strValor) > 0 Then
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & strValor & "'"
                                Else
                                    strCondicao = strCampo & " LIKE '" & strValor & "'"
                                End If
                            ElseIf InStr(1, UCase(.Controls(i).Name), "DTM") > 0 Then
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & gstrConvDtParaSql(strValor)
                                Else
                                    strCondicao = strCampo & " = " & gstrConvDtParaSql(strValor)
                                End If
                            ElseIf InStr(1, UCase(.Controls(i).Name), "DBL") > 0 Then
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & gstrConvVrParaSql(strValor)
                                Else
                                    strCondicao = strCampo & " = " & gstrConvVrParaSql(strValor)
                                End If
                            ElseIf InStr(1, UCase(.Controls(i).Name), "INT") > 0 Then
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                                Else
                                    strCondicao = strCampo & " = " & strValor
                                End If
                            Else
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & strValor & "%'"
                                Else
                                    strCondicao = strCampo & " LIKE '" & strValor & "%'"
                                End If
                            End If
                        End If
                    'If TypeOf .Controls(i) Is TextBox Then
                    ElseIf TypeOf .Controls(i) Is OptionButton Then
                        strValor = .Controls(i).Index
                        strCampo = Trim(.Controls(i).Name)
                        If InStr(1, "_", strCampo) > 0 Then
                            strCampo = Mid(strCampo, 5, Len(strCampo))
                        Else
                            strCampo = Mid(strCampo, 4, Len(strCampo))
                        End If
                        strCampo = "IM." & strCampo
                        
                        If strCondicao <> "" Then
                            strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                        Else
                            strCondicao = strCampo & " = " & strValor
                        End If
                    'ElseIf TypeOf .Controls(i) Is OptionButton Then
                    ElseIf TypeOf .Controls(i) Is CheckBox Then
                        If .Controls(i).Value = 1 Then
                            strValor = .Controls(i).Value
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, "_", strCampo) > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            strCampo = "IM." & strCampo
                            
                            If strCondicao <> "" Then
                                strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                            Else
                                strCondicao = strCampo & " = " & strValor
                            End If
                        End If
                    'ElseIf TypeOf .Controls(i) Is CheckBox Then
                    ElseIf TypeOf .Controls(i) Is DataCombo Then
                        If .Controls(i).Name <> "dbc_intContribuinte" Then
                            If .Controls(i).MatchedWithList Then
                                strValor = .Controls(i).BoundText
                                strCampo = Trim(.Controls(i).Name)
                                
                                If InStr(1, strCampo, "_") > 0 Then
                                    strCampo = Mid(strCampo, 5, Len(strCampo))
                                Else
                                    strCampo = Mid(strCampo, 4, Len(strCampo))
                                End If
                                strCampo = "IM." & strCampo
                                
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                                Else
                                    strCondicao = strCampo & " = " & strValor
                                End If
                            End If
                        End If
                    ElseIf TypeOf .Controls(i) Is MaskEdBox Then
                        If Trim(.Controls(i).ClipText) <> "" Then
                            strValor = .Controls(i).ClipText
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, "_", strCampo) > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            strCampo = "IM." & strCampo
                            
                            'If strCampo = "IM.strInscricao" Then strValor = String(gintLenInscricao - Len(strValor), "0") & strValor
                            If strCondicao <> "" Then
                                If Trim(.Controls(i).Name) = "mskstrInscricaoAuxiliar" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " LIKE '%" & strValor & "'"
                                Else
                                    'strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & strValor & "%'"
                                    strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & UCase(String(gintLenInscricao - gintRetornaTamanhoMascara(TYP_IMOBILIARIA), "0") & strValor) & "%'"
                                    'strCondicao = strCondicao & " AND " & strCampo & " = '" & strValor & "'"
                                End If
                                
                                                               
                            Else
                                If Trim(.Controls(i).Name) = "mskstrInscricaoAuxiliar" Then
                                    strCondicao = strCampo & " LIKE '%" & strValor & "'"
                                Else
                                    'strCondicao = strCampo & " LIKE '" & strValor & "%'"
                                    strCondicao = strCampo & " LIKE '" & UCase(String(gintLenInscricao - gintRetornaTamanhoMascara(TYP_IMOBILIARIA), "0") & strValor) & "%'"
                                    'strCondicao = strCampo & " = '" & strValor & "'"
                                End If
                            End If
                            
                        End If
                    End If 'If TypeOf .Controls(i) Is TextBox Then
                End If
            End If 'If Not (TypeOf .Controls(I) Is OptionButton) Or .Controls(I) = True Then
        End If 'If Not TypeOf .Controls(I) Is Label Then
    Next i
    
    End With

    strsql = ""
    If strCondicao <> "" Then
        strsql = strsql & Left(strQueryListView, InStr(1, strQueryListView, " ORDER BY")) & "AND " & strCondicao & " ORDER BY IM.strInscricao"
    Else
        strsql = strQueryListView
    End If
    
    'Verifica se o Contribuinte está preenchido mas não está na lista
    If Trim(dbcintContribuinte.Text) <> "" And Not dbcintContribuinte.MatchedWithList Then
       strsql = Left(strsql, InStr(1, strsql, " ORDER BY")) & " AND CO.strNome LIKE '" & UCase(Trim(dbcintContribuinte.Text)) & "%' " & Right(strsql, Len(strsql) - InStr(1, strsql, " ORDER BY"))
       If bytDBType = SQLServer Then strsql = Replace(strsql, "*", "")
    End If
    
    LeDaTabelaParaObj gstrImobiliario, tdb_Lista, strsql

End Sub

Private Function blnEnvolvidosOk() As Boolean

Dim intFor          As Integer

    If dbc_intContribuinte.Text = "" Or Not dbc_intContribuinte.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo ""Código/Envolvido"""
        dbc_intContribuinte.SetFocus
        Exit Function
    ElseIf opt_Proprietario(0).Value = False And opt_Proprietario(1).Value = False Then
        ExibeMensagem "Preencha corretamente o campo ""Proprietário/Promissário"""
        opt_Proprietario(0).SetFocus
        Exit Function
    
    End If

    If mblnAlterandoList Then
        If UCase$(lvw_Envolvidos.SelectedItem.Text) = UCase$(dbc_intContribuinte.Text) Then
            blnEnvolvidosOk = True
            Exit Function
        End If
    End If
    
    For intFor = 1 To lvw_Envolvidos.ListItems.Count
        If UCase$(lvw_Envolvidos.ListItems(intFor).Text) = UCase$(dbc_intContribuinte.Text) Then
            ExibeMensagem "O contribuinte selecionado já se encontra na lista!"
            dbc_intContribuinte.SetFocus
            Exit Function
        End If
    Next
    
    blnEnvolvidosOk = True

End Function

Private Sub IncluiAlteraItemLista(blnAlterando As Boolean)
    
Dim objLista                    As Object
Dim intFor                      As Integer
    
    If blnEnvolvidosOk Then
        If mblnAlterandoList Then
            lvw_Envolvidos.ListItems(lvw_Envolvidos.SelectedItem.Index).Text = dbc_intContribuinte.Text
            lvw_Envolvidos.ListItems(lvw_Envolvidos.SelectedItem.Index).Tag = dbc_intContribuinte.BoundText
            lvw_Envolvidos.SelectedItem.SubItems(1) = txt_strCNPJCPFEnv
            lvw_Envolvidos.SelectedItem.SubItems(2) = IIf(opt_Proprietario(0).Value, "Proprietário", "Promissário")
            
            strEnvolvidos(0, lvw_Envolvidos.SelectedItem.Index - 1) = dbc_intContribuinte.BoundText
            strEnvolvidos(1, lvw_Envolvidos.SelectedItem.Index - 1) = IIf(opt_Proprietario(0).Value, 1, 0)
            
        Else
            Set objLista = lvw_Envolvidos.ListItems.Add(, , dbc_intContribuinte)
            objLista.Tag = dbc_intContribuinte.BoundText
            objLista.SubItems(1) = txt_strCNPJCPFEnv
            objLista.SubItems(2) = IIf(opt_Proprietario(0).Value, "Proprietário", "Promissário")
            
            ReDim Preserve strEnvolvidos(2, lvw_Envolvidos.ListItems.Count - 1)
        
            strEnvolvidos(0, lvw_Envolvidos.ListItems.Count - 1) = dbc_intContribuinte.BoundText
            strEnvolvidos(1, lvw_Envolvidos.ListItems.Count - 1) = IIf(opt_Proprietario(0).Value, 1, 0)
            
        End If
    
        LimpaEnvolvidos True
        
    End If
End Sub

Private Sub ExcluiItemLista()
Dim objLista        As Object
Dim intFor          As Integer
    
    For intFor = lvw_Envolvidos.SelectedItem.Index To lvw_Envolvidos.ListItems.Count - 1
        strEnvolvidos(0, lvw_Envolvidos.SelectedItem.Index - 1) = strEnvolvidos(0, lvw_Envolvidos.SelectedItem.Index)
        strEnvolvidos(1, lvw_Envolvidos.SelectedItem.Index - 1) = strEnvolvidos(1, lvw_Envolvidos.SelectedItem.Index)
    Next
    
    lvw_Envolvidos.ListItems.Remove lvw_Envolvidos.SelectedItem.Index
    
    
    ReDim Preserve strEnvolvidos(2, Abs(lvw_Envolvidos.ListItems.Count - 1))
    
    LimpaEnvolvidos True
    
End Sub

Private Sub LimpaEnvolvidos(Optional SetFocus As Boolean = False)
    
    txt_BairroEnv.Text = ""
    txt_CepEnv.Text = ""
    txt_ComplementoEnv.Text = ""
    txt_DistritoEnv.Text = ""
    txt_LogradouroEnv.Text = ""
    txt_MunicipioEnv.Text = ""
    txt_NumeroEnv.Text = ""
    txt_UFEnv.Text = ""

    dbc_intContribuinte.Text = ""
    opt_Proprietario(0).Value = False
    opt_Proprietario(1).Value = False
    txt_PKIdContribuinte2 = ""
    txt_strCNPJCPFEnv.Text = ""
    
    If SetFocus Then dbc_intContribuinte.SetFocus
    
    mblnAlterandoList = False
    
End Sub

Private Function strQueryEnvolvidos(intAux As Long) As String
Dim strsql      As String
Dim intFor      As Integer
    
    strsql = IIf(bytDBType = Oracle, "Begin ", "")
    
    For intFor = 0 To lvw_Envolvidos.ListItems.Count - 1
        strsql = strsql & "INSERT INTO " & gstrImobiliarioProprietarios
        strsql = strsql & " (intImovel, intContribuinte, bitProprietario, lngCodUsr, dtmDtAtualizacao) VALUES "
        strsql = strsql & " (" & intAux & ", "
        strsql = strsql & strEnvolvidos(0, intFor) & ", " & strEnvolvidos(1, intFor) & ", "
        strsql = strsql & glngCodUsr & ", " & strGETDATE & ")" & Chr(13)
        strsql = strsql & IIf(bytDBType = Oracle, ";", "")
    Next
    strsql = strsql & IIf(bytDBType = Oracle, "End;", "")
    
    strQueryEnvolvidos = strsql
    
End Function

Private Sub GravaEnvolvidos(intAux As Long, Optional blnLimparControles As Boolean = True)
Dim strsql As String
    
    DeletaEnvolvidos intAux

    If lvw_Envolvidos.ListItems.Count = 0 Then Exit Sub

    gobjBanco.Execute strQueryEnvolvidos(intAux)
    
    If blnLimparControles Then lvw_Envolvidos.ListItems.Clear
    
End Sub

Private Sub DeletaEnvolvidos(intAux As Long)
Dim strsql As String
    
    strsql = ""
    strsql = "DELETE FROM " & gstrImobiliarioProprietarios _
              & " WHERE intImovel = " & intAux
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute strsql

End Sub

Private Sub CarregaEnvolvidos()
Dim strsql       As String
Dim adoResultado As ADODB.Recordset
Dim objLista        As Object
    
    lvw_Envolvidos.ListItems.Clear
    LimpaEnvolvidos
    
    If txtPKId.Text = "" Then Exit Sub
    
    strsql = ""
    strsql = strsql & "SELECT  "
    strsql = strsql & "IP.bitProprietario, CO.PKId, CO.strNome, CO.strCNPJCPF FROM "
    strsql = strsql & gstrImobiliarioProprietarios & " IP INNER JOIN "
    strsql = strsql & gstrContribuinte & " CO ON IP.intContribuinte = CO.PKId "
    strsql = strsql & "WHERE IP.intImovel= " & txtPKId
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            ReDim Preserve strEnvolvidos(2, .RecordCount)
            Do While Not .EOF
                Set objLista = lvw_Envolvidos.ListItems.Add(, , .Fields("strNome"))
                objLista.Tag = .Fields("PKId")
                objLista.SubItems(1) = gstrENulo(.Fields("strCNPJCPF"))
                objLista.SubItems(2) = IIf(.Fields("bitProprietario"), "Proprietário", "Promissário")
                        
                strEnvolvidos(0, .AbsolutePosition - 1) = .Fields("PKId")
                strEnvolvidos(1, .AbsolutePosition - 1) = IIf(.Fields("bitProprietario") = True, 1, 0)
                .MoveNext
            Loop
        End With
    End If

    If lvw_Envolvidos.ListItems.Count <> 0 Then
        lvw_Envolvidos.SelectedItem.Selected = False
    End If
    
    Set adoResultado = Nothing
    
End Sub

Private Function MostraDadosEnvolvidos(intBound As Long) As Boolean
Dim strsql As String
On Error Resume Next
    
    strsql = ""
    strsql = strsql & "SELECT CO.strBairroC,"
    strsql = strsql & " TL.strSigla " & strCONCAT & "' '" & strCONCAT
    strsql = strsql & " TTL.strSigla " & strCONCAT & "' '" & strCONCAT
    strsql = strsql & " CO.strLogradouroC AS strLogradouroC,"
    strsql = strsql & " CO.intNumeroC,"
    strsql = strsql & " CO.strComplementoC,"
    strsql = strsql & " CO.intCEPC,"
    strsql = strsql & " CO.strDistritoC,"
    strsql = strsql & " CD.strDescricao,"
    strsql = strsql & " UF.strSigla"
    strsql = strsql & " FROM "
    strsql = strsql & gstrContribuinte & " CO, "
    strsql = strsql & gstrCidade & " CD, "
    strsql = strsql & gstrTipoLogradouro & " TL, "
    strsql = strsql & gstrTituloLogradouro & " TTL, "
    strsql = strsql & gstrUF & " UF"
    strsql = strsql & " WHERE intMunicipioC = CD.PKId  AND"
    strsql = strsql & " intUFC = UF.PKId AND"
    strsql = strsql & " TL.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CO.intTipoLogradouro AND"
    strsql = strsql & " TTL.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CO.intTituloLogradouro AND"
    strsql = strsql & " CO.PKId = " & intBound
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                txt_BairroEnv = gstrVerificaCampoNulo(!strBairroC)
                txt_CepEnv = gstrCEPFormatado(gstrVerificaCampoNulo(!intcepc))
                txt_ComplementoEnv = gstrVerificaCampoNulo(!strComplementoC)
                txt_DistritoEnv = gstrVerificaCampoNulo(!strDistritoC)
                txt_LogradouroEnv = gstrVerificaCampoNulo(!strlogradouroc)
                txt_MunicipioEnv = gstrVerificaCampoNulo(!strDescricao)
                txt_NumeroEnv = gstrVerificaCampoNulo(!intNumeroC)
                txt_UFEnv = gstrVerificaCampoNulo(!strsigla)
                MostraDadosEnvolvidos = True
                .MoveNext
            Loop
        End With
    End If
End Function

Private Sub txtstrSequenciaDeFace_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSequenciaDeFace
End Sub

Private Sub VerificaTestada(strTestada As String, intLinha As Integer)
Dim XTestada    As XArrayDB
Dim intFor      As Integer
    
    Set XTestada = New XArrayDB
    
    XTestada.ReDim 0, 0, 0, 6
    XTestada(0, 0) = ""
    XTestada(0, 1) = ""
    XTestada(0, 2) = ""
    XTestada(0, 3) = ""
    XTestada(0, 4) = ""
    XTestada(0, 5) = ""
    XTestada(0, 6) = ""
    XTestada.Clear
        
    Set XTestada = grd_Testada.Array
    
    For intFor = 0 To XTestada.Count(1) - 1
        If intFor <> intLinha Then
            If XTestada(intFor, 1) = strTestada Then
                ExibeMensagem "Não é possível inserir um mesmo Tipo de Testada."
                grd_Testada.Row = intLinha
                grd_Testada.Columns("Tipo de Testada").Text = ""
                If grd_Testada.Columns(2) = "" And grd_Testada.Columns(3) = "" Then
                   grd_Testada.Delete
                End If
                grd_Testada.Refresh
                grd_Testada.SetFocus
                Exit Sub
            End If
        End If
    Next
    
    Set grd_Testada.Array = XTestada
    grd_Testada.Refresh
    grd_Testada.Update
    
End Sub

Private Function UltimoNumeroEdificio(intImobiliario As Long) As Long
Dim strsql         As String
Dim adoResultado   As ADODB.Recordset

    strsql = "SELECT MAX(intNEdificacao) AS UltimoNumero"
    strsql = strsql & " FROM "
    strsql = strsql & gstrAreaImobiliario
    strsql = strsql & " WHERE intImobiliario = " & intImobiliario
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            UltimoNumeroEdificio = Val(gstrENulo(adoResultado!UltimoNumero)) + 1
        End If
    End If
    
End Function

Private Function blnExisteFaceQuadra() As Boolean
Dim strsql As String
Dim adoResultado As ADODB.Recordset

    strsql = "SELECT Pkid "
    strsql = strsql & " FROM "
    strsql = strsql & gstrCampoDeInscricao
    strsql = strsql & " WHERE bytSetorQuadra IN (1,2) AND intTipoDeInscricao = " & TYP_IMOBILIARIA 'Cláudio - Quadra e Setor cadastrado (Face de quadra existente)

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If adoResultado.RecordCount = 2 Then
                blnExisteFaceQuadra = True
            Else
                blnExisteFaceQuadra = False
            End If
        Else
            blnExisteFaceQuadra = False
        End If
    End If

End Function

Private Function PreencheComboDocProc() As String
Dim strsql As String

    strsql = "SELECT Pkid,"
    strsql = strsql & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao"
    strsql = strsql & " FROM "
    strsql = strsql & gstrDocumentos
    strsql = strsql & " ORDER BY strDescricao"
    
    PreencheComboDocProc = strsql

End Function

Private Function PreencheDtProcesso() As String
Dim strsql          As String
Dim adoResultado    As ADODB.Recordset

    strsql = "SELECT dtmdtData"
    strsql = strsql & " FROM "
    strsql = strsql & gstrProtocolizacaoProcesso
    strsql = strsql & " WHERE strCodigo = '" & Trim(txt_strCodigo.Text) & "' AND"
    strsql = strsql & " bitDigito = " & Val(txt_bitDigito.Text) & " AND"
    strsql = strsql & " intExercicio = " & Val(txt_intExercicio.Text)
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_DtmDataDocProc = gstrDataFormatada(gstrENulo(adoResultado!dtmDtData))
        Else
            txt_DtmDataDocProc = ""
        End If
    End If

End Function

Private Sub GravaDocProc()
Dim strsql As String
    strsql = "INSERT INTO "
    strsql = strsql & gstrDocumentosImobiliario
    strsql = strsql & " (intImobiliario, intTiposDocumentoProcesso,"
    strsql = strsql & " strCodigo, bitDigito, intExercicio, dtmdtDataProc, strObservacoes, dtmDtAtualizacao, lngCodUsr)"
    strsql = strsql & " VALUES("
    strsql = strsql & "'" & txtPKId & "',"
    strsql = strsql & "'" & dbc_intTiposDocumentosProcesso.BoundText & "',"
    strsql = strsql & "'" & Trim(txt_strCodigo.Text) & "', "
    strsql = strsql & Val(txt_bitDigito.Text) & ", "
    strsql = strsql & Val(txt_intExercicio.Text) & ", "
    strsql = strsql & gstrConvDtParaSql(txt_DtmDataDocProc) & ","
    strsql = strsql & "'" & txt_strObservacoes.Text & "',"
    strsql = strsql & strGETDATE & ", "
    strsql = strsql & "'" & glngCodUsr & "'"
    strsql = strsql & ")"
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute (strsql)

End Sub

Private Sub LimpaTabDocProc(blnLimpaGrdDocumentos As Boolean)

    txt_PkidDocProc.Text = ""
    
    dbc_intTiposDocumentosProcesso.ListField = ""
    dbc_intTiposDocumentosProcesso.Text = ""
    
    'dbc_intProcesso.ListField = ""
    'dbc_intProcesso.Text = ""
    txt_strCodigo.Text = ""
    txt_bitDigito.Text = ""
    txt_intExercicio.Text = ""
    
    txt_DtmDataDocProc.Text = ""
    
    txt_strObservacoes.Text = ""
    
    dbc_intTiposDocumentosProcesso.SetFocus
    
    If blnLimpaGrdDocumentos Then Set tdb_DocumentosProcessos.DataSource = Nothing
    
End Sub

Private Sub PreencheGRDDocProc()
Dim strsql As String

    strsql = "SELECT DI.Pkid,"
    strsql = strsql & " TP.strDescricao DocumentoProcesso,"
    strsql = strsql & " PP.strCodigo " & strCONCAT & "'/'" & strCONCAT
    strsql = strsql & gstrCONVERT(CDT_NVARCHAR, " PP.intExercicio ") & strCONCAT & "'-'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, " PP.bitDigito") & " AS Processo,"
    strsql = strsql & " PP.dtmDtData DataProc,"
    strsql = strsql & " DI.strObservacoes,"
    strsql = strsql & " TP.Pkid PkidDocumentoProcesso,"
    strsql = strsql & " PP.Pkid PkidProcesso"
    strsql = strsql & " FROM "
    strsql = strsql & gstrDocumentosImobiliario & " DI, "
    strsql = strsql & gstrDocumentos & " TP, "
    strsql = strsql & gstrProtocolizacaoProcesso & " PP"
    strsql = strsql & " WHERE intImobiliario = '" & txtPKId & "' AND"
    strsql = strsql & " DI.intTiposDocumentoProcesso " & strOUTJSQLServer & "=" & " TP.Pkid " & strOUTJOracle & " AND"
    strsql = strsql & " DI.strCodigo " & strOUTJSQLServer & "=" & " PP.strCodigo " & strOUTJOracle & " AND"
    strsql = strsql & " DI.bitDigito " & strOUTJSQLServer & "=" & " PP.bitDigito " & strOUTJOracle & " AND"
    strsql = strsql & " DI.intExercicio " & strOUTJSQLServer & "=" & " PP.intExercicio" & strOUTJOracle

    strsql = strsql & " ORDER BY DocumentoProcesso "

    LeDaTabelaParaObj "", tdb_DocumentosProcessos, strsql

End Sub

Private Sub AlteraDocProc(lngPkid As Long)
Dim strsql As String

    strsql = "UPDATE "
    strsql = strsql & gstrDocumentosImobiliario
    strsql = strsql & " SET intTiposDocumentoProcesso = " & dbc_intTiposDocumentosProcesso.BoundText & ","
    'strSQL = strSQL & " intProcesso = " & dbc_intProcesso.BoundText & ","
    strsql = strsql & " strCodigo = '" & Trim(txt_strCodigo.Text) & "',"
    strsql = strsql & " bitDigito = " & Val(txt_bitDigito.Text) & ","
    strsql = strsql & " intExercicio = " & Val(txt_intExercicio.Text) & ","
    strsql = strsql & " dtmDtDataProc = " & gstrConvDtParaSql(txt_DtmDataDocProc.Text) & ","
    strsql = strsql & " strObservacoes = '" & txt_strObservacoes.Text & "',"
    strsql = strsql & " dtmDtAtualizacao = " & strGETDATE & ","
    strsql = strsql & " lngCodUsr = " & glngCodUsr
    strsql = strsql & " WHERE Pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute (strsql)
    
End Sub

Private Sub DeletaDocProc(lngPkid)
Dim strsql As String
    
    strsql = "DELETE FROM "
    strsql = strsql & gstrDocumentosImobiliario
    strsql = strsql & " WHERE Pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute (strsql)

End Sub

Private Function MelhoriaChecada(intMelhoria As Integer) As Boolean
Dim strsql          As String
Dim adoResultado    As ADODB.Recordset

    strsql = "SELECT Pkid"
    strsql = strsql & " FROM "
    strsql = strsql & gstrEquipamentoImobiliario
    strsql = strsql & " WHERE intImobiliario = " & txtPKId & " AND"
    strsql = strsql & " intMelhoria = " & intMelhoria
    strsql = strsql & " AND intFaceDeQuadra = '" & dbc_intFaceDeQuadra.BoundText & "'"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            MelhoriaChecada = True
        Else
            MelhoriaChecada = False
        End If
    End If
    
End Function

Private Function blnVerificaFaceDeQuadra() As Boolean
Dim strsql          As String
Dim adoResultado    As ADODB.Recordset
    
On Error GoTo Erro_blnVerificaFaceDeQuadra
    
    'Caso nao esteja preenchido a seq. de quadra nao vamos validar
    If Len(Trim(txtstrSequenciaDeFace.Text)) > 0 Then
    
        blnVerificaFaceDeQuadra = False

        If blnExisteFaceQuadra Then
                
            strsql = "SELECT FQ.PKID FROM " & gstrFaceDeQuadra & " FQ, " & gstrHistoricoFaceDeQuadra & " HFQ"
            strsql = strsql & " WHERE FQ.strSetor = '" & Mid(mskstrInscricao.FormattedText, (Val(Mid(strPosicaoTamanho(1), 1, 1))), Val(Mid(strPosicaoTamanho(1), 3, 1))) & "' AND "       'Quadra (bytSetorQuadra = 1)
            strsql = strsql & "FQ.strQuadra = '" & Mid(mskstrInscricao.FormattedText, (Val(Val(Mid(strPosicaoTamanho(1), 1, 1)) + Mid(strPosicaoTamanho(2), 1, 1))), Mid(strPosicaoTamanho(2), Val(Mid(strPosicaoTamanho(2), 1, 1)) + 1, Val(Mid(strPosicaoTamanho(2), 3, 1)))) & "' AND "    'Setor (bytSetorQuadra = 2)
            strsql = strsql & "FQ.strSequenciaDeFace = '" & Format$(txtstrSequenciaDeFace.Text, "00") & "' AND "
            strsql = strsql & "FQ.intLogradouro = " & gstrENulo(dbcintLogradouro.BoundText, , True) & " AND "
            strsql = strsql & "HFQ.intFaceDeQuadra " & strOUTJOracle & "=" & strOUTJSQLServer & " FQ.Pkid AND "
            strsql = strsql & "HFQ.intExercicio = " & Year(gstrDataDoSistema)
                
            Set gobjBanco = New clsBanco
                
            If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
                If adoResultado.EOF Then
                    ExibeMensagem "Face de quadra não cadastrada!"
                    Exit Function
                End If
            End If
        Else
            ExibeMensagem "Não existe nenhuma Face de Quadra Cadastrada."
            Exit Function
        End If
    
    End If
    
    blnVerificaFaceDeQuadra = True
Exit Function

Erro_blnVerificaFaceDeQuadra:
    blnVerificaFaceDeQuadra = True
    
End Function

Private Sub txtstrSequenciaDeFace_LostFocus()
    
    If Trim(mskstrInscricao) <> "" And Trim(txtstrSequenciaDeFace) <> "" And dbcintLogradouro.MatchedWithList = True Then
        If Not blnVerificaFaceDeQuadra Then
            If mskstrInscricao.Enabled = True Then
                mskstrInscricao.SetFocus
            End If
        End If
    ElseIf Trim(mskstrInscricao) = "" Then
        ExibeMensagem "Face de quadra inválida inscrição não preenchida."
        If mskstrInscricao.Enabled = True Then
            mskstrInscricao.SetFocus
        End If
        Exit Sub
    ElseIf Trim(txtstrSequenciaDeFace) = "" Then
        ExibeMensagem "Face de quadra inválida sequência não preenchida."
        Exit Sub
    ElseIf dbcintLogradouro.MatchedWithList = False Then
        ExibeMensagem "Face de quadra inválida logradouro não preenchido."
        If dbcintLogradouro.Enabled = True Then
            dbcintLogradouro.SetFocus
        End If
        Exit Sub
    End If
    
End Sub

Private Sub PrrencheGRDCategoriaConstrucao()
Dim strsql As String

    strsql = "SELECT PKid, strDescricao"
    strsql = strsql & " FROM "
    strsql = strsql & gstrCategoriaConstrucao
    strsql = strsql & " WHERE intUtilizacaoTabelaValor = 3 " 'Imobiliario Construcao
    strsql = strsql & " ORDER BY strDescricao"
    
    LeDaTabelaParaObj "", tdd_CategoriaConstrucao, strsql

End Sub

Private Sub PreencheProcesso(lngPkidProcesso As Long)
Dim strsql          As String
Dim adoResultado    As ADODB.Recordset


    strsql = "SELECT PP.strCodigo,"
    strsql = strsql & " PP.bitDigito,"
    strsql = strsql & " PP.intExercicio"
    strsql = strsql & " FROM "
    strsql = strsql & gstrProtocolizacaoProcesso & " PP"
    strsql = strsql & " WHERE"
    strsql = strsql & " PKid = " & Val(lngPkidProcesso)
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_strCodigo = gstrENulo(adoResultado!strCodigo)
            txt_bitDigito = gstrENulo(adoResultado!bitDigito)
            txt_intExercicio = gstrENulo(adoResultado!intExercicio)
        Else
            txt_strCodigo = ""
            txt_bitDigito = ""
            txt_intExercicio = ""
        End If
    End If

End Sub

Public Sub PreencheCkIsencao()
Dim strsql          As String
Dim adoResult    As ADODB.Recordset

    strsql = strsql & "Select "
    strsql = strsql & "IM.strinscricao, "
    strsql = strsql & "IP.Bytposicao, "
    strsql = strsql & "IP.Bytcancelamento "
    strsql = strsql & "From "
    strsql = strsql & "tblimobiliario IM, "
    strsql = strsql & "tblisencaoimunidade I, "
    strsql = strsql & "tblisencaoperiodo IP "
    strsql = strsql & "Where "
    strsql = strsql & "Im.Pkid = I.intidentificacao and "
    strsql = strsql & "I.BITTIPODEINSCRICAO = 0 and "
    strsql = strsql & "i.pkid = Ip.Intisencaoimunidade and "
    If bytDBType = Oracle Then
        strsql = strsql & "(TO_NUMBER(TO_CHAR(IP.DTMINICIAL, 'yyyy')) <= " & Year(gstrDataDoSistema) & " And "
        strsql = strsql & "TO_NUMBER(TO_CHAR(IP.DTMFINAL, 'yyyy')) >= " & Year(gstrDataDoSistema) & ")" & " AND "
    Else
        strsql = strsql & "Year(IP.DTMINICIAL) <= " & Year(gstrDataDoSistema) & " And "
        strsql = strsql & "Year(IP.DTMFINAL) >= " & Year(gstrDataDoSistema) & " AND "
    End If
    strsql = strsql & "IM.Pkid = '" & txtPKId & "' "
    strsql = strsql & " Group By "
    strsql = strsql & "IM.strinscricao, "
    strsql = strsql & "IP.Bytposicao, "
    strsql = strsql & "IP.Bytcancelamento"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strsql, 5, adoResult) Then
        If Not adoResult.EOF Then
            If adoResult!Bytcancelamento = 1 Then
                chk_intIsencao.Value = vbUnchecked
            Else
                If adoResult!bytPosicao <> 0 Then
                    chk_intIsencao.Value = vbUnchecked
                Else
                    chk_intIsencao.Value = vbChecked
                End If
            End If
        Else
            chk_intIsencao.Value = vbUnchecked
       End If
    End If

End Sub

Private Sub CarregaFatorTerreno()
Dim strsql      As String
Dim dblValor    As Double
Dim dblValorGleba As Double
    
    lvw_FatorTerreno.ListItems.Clear
    
    'Vamos pegar o fator de profundidade
    If dblValorProfundidade > 0 Then
        dblValor = IIf(txtdblArea <= "", 0, txtdblArea) / dblValorProfundidade
    Else
        dblValor = 0
    End If
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "dblfaixainicial, "
    strsql = strsql & "dblfaixafinal, "
    strsql = strsql & "dblValor AS Indice "
    strsql = strsql & "From "
    strsql = strsql & gstrValorDaFaixa
    strsql = strsql & " Where "
    strsql = strsql & "intfaixadevalores = " & FATOR_ZONEAMENTO
    
    Set adoResultado = Nothing
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Do While .EOF = False
                    If dblValor >= gstrENulo(!dblFaixaInicial) And dblValor <= gstrENulo(!dblFaixaFinal) Then
                        Set objList1 = lvw_FatorTerreno.ListItems.Add(, , "Fator Profundidade")
                        objList1.Tag = ""
                        objList1.SubItems(1) = ""
                        objList1.SubItems(2) = dblValor
                        objList1.SubItems(3) = gstrENulo(!Indice)
                        Exit Do
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
    Set adoResultado = Nothing
    Set gobjBanco = Nothing
    
    blnGleba = False
    dblValorGleba = IIf(txtdblArea <= "", 0, txtdblArea)
    
    If dblValorGleba >= CDbl("10000") Then
        'Vamos pegar o fator de Gleba
        strsql = ""
        strsql = strsql & "Select "
        strsql = strsql & "dblfaixainicial, "
        strsql = strsql & "dblfaixafinal, "
        strsql = strsql & "dblValor AS Indice "
        strsql = strsql & "From "
        strsql = strsql & gstrValorDaFaixa
        strsql = strsql & " Where "
        strsql = strsql & "intfaixadevalores = " & FATOR_SITUACAO
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
            With adoResultado
                If Not .EOF Then
                    Do While .EOF = False
                        If dblValorGleba >= gstrENulo(!dblFaixaInicial) And dblValorGleba <= gstrENulo(!dblFaixaFinal) Then
                            Set objList1 = lvw_FatorTerreno.ListItems.Add(, , "Fator Gleba")
                            objList1.Tag = ""
                            objList1.SubItems(1) = ""
                            objList1.SubItems(2) = dblValorGleba
                            objList1.SubItems(3) = gstrENulo(!Indice)
                            blnGleba = True
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        End If
        Set adoResultado = Nothing
        Set gobjBanco = Nothing
    End If
    
    'Vamos pegar as caracteristicas e seus fatores
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "CG.PKId As intCaracteristicaGeral, "
    strsql = strsql & "CG.strNomeDaCaracteristica As Caracteristica, "
    strsql = strsql & "D.intCodigoDetalheDaCaracteristi As IntDetalheCaracteristica, "
    strsql = strsql & "B.STRNOMEDODETALHE As Detalhe, "
    strsql = strsql & "E.DBLVALOR As Valor "
    strsql = strsql & "From "
    strsql = strsql & gstrCaracteristicaGeral & " CG, "
    strsql = strsql & gstrImobiliario & " A, "
    strsql = strsql & gstrDetalheDaCaracteristica & " B, "
    strsql = strsql & gstrUtilizacaoDaTabelaDeValor & " C, "
    strsql = strsql & gstrCaracteristicaDoImovel & " D, "
    strsql = strsql & gstrTabelaDeValor & " E "
    strsql = strsql & "Where "
    'strSQL = strSQL & "CG.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & "D.Intcodigocaracteristicageral  AND "
    strsql = strsql & "CG.Pkid = D.Intcodigocaracteristicageral  AND "
    strsql = strsql & "A.PKId  = D.intCodigoImobiliario " & strOUTJOracle & " AND "
    strsql = strsql & "B.pkid  " & strOUTJOracle & "= D.Intcodigodetalhedacaracteristi   AND "
    strsql = strsql & "C.PKId  " & strOUTJOracle & "=" & strOUTJSQLServer & " CG.intUtilizacaoDaCaracteristica AND " 'D.intCodigoUtilizacaoDaTabelaDeV  Alterado Rafael 21/10/04
    strsql = strsql & "E.Pkid  " & strOUTJOracle & "=" & strOUTJSQLServer & " B.Inttabeladevalores AND "
    strsql = strsql & "CG.intUtilizacaoDaCaracteristica = 2 AND "
    strsql = strsql & "A.Pkid = " & txtPKId
    strsql = strsql & " ORDER BY CG.strNomeDaCaracteristica"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Do While .EOF = False
                    Set objList1 = lvw_FatorTerreno.ListItems.Add(, , gstrENulo(!Caracteristica))
                    objList1.Tag = gstrENulo(!intCaracteristicaGeral)
                    objList1.SubItems(1) = gstrENulo(!IntDetalheCaracteristica)
                    objList1.SubItems(2) = gstrENulo(!Detalhe)
                    objList1.SubItems(3) = gstrConvVrDoSql(gstrENulo(!Valor))
                    .MoveNext
                Loop
                lvw_FatorTerreno.Refresh
            Else
                lvw_FatorTerreno.ListItems.Clear
            End If
        End With
    End If
    Set adoResultado = Nothing
    
    DoEvents
End Sub

Private Sub AtualizaFatorTerreno(IntDetalheCaracteristica As Long)
Dim strsql As String
Dim intFor As Integer
    
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "A.intCaracteristica , "
    strsql = strsql & "A.Pkid as IntDetalheCaracteristica, "
    strsql = strsql & "A.STRNOMEDODETALHE AS Detalhe, "
    strsql = strsql & "B.DBLVALOR "
    strsql = strsql & "From "
    strsql = strsql & gstrDetalheDaCaracteristica & " A, "
    strsql = strsql & gstrTabelaDeValor & " B "
    strsql = strsql & "Where "
    strsql = strsql & "B.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " A.INTTABELADEVALORES AND "
    strsql = strsql & "A.Pkid = " & IntDetalheCaracteristica
    
    Set gobjBanco = New clsBanco
    Set adoResultado = Nothing
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            For intFor = IIf(blnGleba, 3, 2) To lvw_FatorTerreno.ListItems.Count
                If lvw_FatorTerreno.ListItems(intFor).Tag = gstrENulo(adoResultado!intCaracteristica) Then
                    lvw_FatorTerreno.ListItems(intFor).SubItems(1) = gstrENulo(adoResultado!IntDetalheCaracteristica)
                    lvw_FatorTerreno.ListItems(intFor).SubItems(2) = gstrENulo(adoResultado!Detalhe)
                    lvw_FatorTerreno.ListItems(intFor).SubItems(3) = gstrConvVrDoSql(gstrENulo(adoResultado!dblValor))
                    Exit For
                End If
            Next
        End If
    End If
    
End Sub

Private Sub CarregaConstrucao()
    Dim strsql              As String
    Dim dblTotalConstrucao  As String
    Dim blnFatores          As Boolean
    
    dblTotalConstrucao = 0
    lvw_CaracPredio.ListItems.Clear
    blnFatores = False
    
    'Vamos pegar as caracteristicas e seus fatores
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "CG.PKId As intCaracteristicaGeral, "
    strsql = strsql & "CG.strNomeDaCaracteristica As Caracteristica, "
    strsql = strsql & "D.intCodigoDetalheDaCaracteristi As IntDetalheCaracteristica, "
    strsql = strsql & "B.STRNOMEDODETALHE As Detalhe, "
    strsql = strsql & "E.DBLVALOR As Valor "
    strsql = strsql & "From "
    strsql = strsql & gstrCaracteristicaGeral & " CG, "
    strsql = strsql & gstrImobiliario & " A, "
    strsql = strsql & gstrDetalheDaCaracteristica & " B, "
    strsql = strsql & gstrUtilizacaoDaTabelaDeValor & " C, "
    strsql = strsql & gstrCaracteristicaDoImovel & " D, "
    strsql = strsql & gstrTabelaDeValor & " E "
    strsql = strsql & "Where "
    'strSQL = strSQL & "CG.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & "D.Intcodigocaracteristicageral  AND "
    strsql = strsql & "CG.Pkid = D.Intcodigocaracteristicageral  AND "
    strsql = strsql & "A.PKId  = D.intCodigoImobiliario  AND "
    strsql = strsql & "B.pkid  " & strOUTJOracle & "= D.Intcodigodetalhedacaracteristi  AND "
    strsql = strsql & "C.PKId  " & strOUTJOracle & "=" & strOUTJSQLServer & " CG.intUtilizacaoDaCaracteristica  AND " 'D.intCodigoUtilizacaoDaTabelaDeV  Alterado Rafael 21/10/04
    strsql = strsql & "E.Pkid  " & strOUTJOracle & "=" & strOUTJSQLServer & " B.Inttabeladevalores  AND "
    strsql = strsql & "CG.intUtilizacaoDaCaracteristica = 3  AND "
    strsql = strsql & "D.Intarea = " & IIf(grd_Area.Columns("Pkid").Value <> "", grd_Area.Columns("Pkid").Value, 0) & " AND "
    strsql = strsql & "A.Pkid = " & txtPKId
    strsql = strsql & " ORDER BY CG.intcodigodacaracteristica"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Do While .EOF = False
                    Set objList1 = lvw_CaracPredio.ListItems.Add(, , gstrENulo(!Caracteristica))
                    objList1.Tag = gstrENulo(!intCaracteristicaGeral)
                    objList1.SubItems(1) = gstrENulo(!IntDetalheCaracteristica)
                    objList1.SubItems(2) = gstrENulo(!Detalhe)
                    objList1.SubItems(3) = gstrConvVrDoSql(gstrENulo(!Valor))
                    dblTotalConstrucao = dblTotalConstrucao + CDbl(gstrConvVrDoSql(IIf(IsNull(!Valor), "0", !Valor)))
                    .MoveNext
                    blnFatores = True
                Loop
                lvw_CaracPredio.Refresh
            Else
                lvw_CaracPredio.ListItems.Clear
                lbl_DescricaoPredios = ""
            End If
        End With
    End If
    Set adoResultado = Nothing
    
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "FP.Pkid, "
    strsql = strsql & "FP.Strdescricao as DescricaoFaixaPontos, "
    strsql = strsql & "FP.STRCODIGO, "
    strsql = strsql & "FP.DBLPONTOINICIAL, "
    strsql = strsql & "FP.DBLPONTOFINAL, "
    strsql = strsql & "CC.strDescricao as DescricaoConstrucao, "
    strsql = strsql & "EVM2.DBLVALOR as dblValorM2, "
    strsql = strsql & "ME.STRABREVIATURA as strMoeda "
    strsql = strsql & "From "
    strsql = strsql & gstrCategoriaConstrucao & " CC, "
    strsql = strsql & gstrFaixaPontosPredio & " FP, "
    strsql = strsql & gstrExercicioValorM2Predio & " EVM2, "
    strsql = strsql & gstrMoedas & " ME "
    strsql = strsql & "Where "
    strsql = strsql & "CC.Pkid = FP.Intcategoriaconstrucao and "
    strsql = strsql & "FP.Pkid = EVM2.Intfaixapontospredio and "
    strsql = strsql & "ME.Pkid = EVM2.Intmoeda and "
    strsql = strsql & "CC.pkid =  " & IIf(grd_Area.Columns("PkidCategoriaConstrucao").Value = "", 0, grd_Area.Columns("PkidCategoriaConstrucao").Value)
    strsql = strsql & " AND EVM2.intexercicio = " & Year(gstrDataDoSistema)
    strsql = strsql & " Order By "
    strsql = strsql & "FP.Pkid "
    
    If blnFatores = True Then
        Set adoResultado = Nothing
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
            With adoResultado
                If Not .EOF Then
                    Do While .EOF = False
                        If CDbl(dblTotalConstrucao) >= gstrENulo(!DBLPONTOINICIAL) And CDbl(dblTotalConstrucao) <= gstrENulo(!DBLPONTOFINAL) Then
                            lbl_DescricaoPredios = gstrENulo(!DescricaoConstrucao) & " - " & gstrENulo(!DescricaoFaixaPontos) & " / Pontuação: " & dblTotalConstrucao & " / Valor M²: " & gstrENulo(!dblValorM2) & " / Moeda: " & gstrENulo(!strMoeda)
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        End If
        Set adoResultado = Nothing
        Set gobjBanco = Nothing
    End If
    DoEvents
    
End Sub

Private Sub DeletaCaracteristicasDoImovel()
Dim strsql As String
        
    strsql = ""
    strsql = strsql & " DELETE "
    strsql = strsql & " FROM " & gstrCaracteristicaDoImovel
    strsql = strsql & " WHERE "
    'Código do imobiliário
    strsql = strsql & " intCodigoImobiliario = " & Val(txtPKId)
    'Código da utilização
    strsql = strsql & " AND intCodigoUtilizacaoDaTabelaDeV = 3"
    strsql = strsql & " AND intArea = " & grd_Area.Columns(0).Value
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strsql
    
End Sub

Private Function blnDetalhesDaCaracterisiticas(IntUtilizacaoValor As Long) As Boolean
Dim strsql  As String
    
    blnDetalhesDaCaracterisiticas = True
    
    'Vamos pegar todas as Características
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "Pkid, Strnomedacaracteristica "
    strsql = strsql & "From "
    strsql = strsql & gstrCaracteristicaGeral & " "
    strsql = strsql & "Where "
    strsql = strsql & "Intutilizacaodacaracteristica = " & IntUtilizacaoValor & " "
    If IntUtilizacaoValor = 3 Then
        strsql = strsql & " AND INTCATEGORIACONSTRUCAO = " & grd_Area.Columns("PkidCategoriaConstrucao").Value & " "
    End If
    strsql = strsql & "Order By strNomeDaCaracteristica"
    Set adoResultado = New ADODB.Recordset
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Do While Not .EOF
                    'Vamos verificar se esta checado pelo menos um detalhe de cada característica
                    strsql = ""
                    strsql = strsql & "Select "
                    strsql = strsql & "CI.Pkid "
                    strsql = strsql & "From "
                    strsql = strsql & gstrCaracteristicaDoImovel & " CI"
                    If IntUtilizacaoValor = 3 Then
                        strsql = strsql & ", " & gstrCaracteristicaGeral & " CG"
                    End If
                    strsql = strsql & " Where "
                    If IntUtilizacaoValor = 3 Then
                        'Tipo Construcão
                        strsql = strsql & "CG.Pkid = CI.Intcodigocaracteristicageral AND "
                        strsql = strsql & "CG.INTCATEGORIACONSTRUCAO = " & grd_Area.Columns("PkidCategoriaConstrucao").Value & " AND "
                    End If
                    'Imobiliario
                    strsql = strsql & "CI.INTCODIGOIMOBILIARIO = " & txtPKId & " AND "
                    'Terreno / Prédio
                    strsql = strsql & "CI.Intcodigoutilizacaodatabeladev = " & IntUtilizacaoValor & " AND "
                    'Caracteristica Geral
                    strsql = strsql & "CI.Intcodigocaracteristicageral = " & gstrENulo(!Pkid) & " AND "
                    'Detalhe
                    strsql = strsql & "Not CI.Intcodigodetalhedacaracteristi is null"
                    Set adoRec = New ADODB.Recordset
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strsql, 5, adoRec) Then
                        If adoRec.EOF Then
                            ExibeMensagem "É necessário selecionar detalhe para característica " & gstrENulo(!strNomeDaCaracteristica) & "."
                            blnDetalhesDaCaracterisiticas = False
                            If IntUtilizacaoValor = 3 Then
                                tab_3DPasta.Tab = 6
                            Else
                                tab_3DPasta.Tab = 5
                            End If
                            Exit Do
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If

End Function

Private Function strQueryLogradouro(blnConsulta As Boolean) As String
    Dim strsql  As String
     
    strsql = ""
    
    strsql = strsql & "SELECT L.PKId, "
    strsql = strsql & " RTRIM(LTRIM(L.strDescricao)) " & strCONCAT & gstrISNULL("TL.strSigla", "''", "', '") & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & _
             strCONCAT & gstrISNULL("U.strDescricao", "' '", "', '") & strCONCAT & gstrISNULL("U.strDescricao", "''") & ")) " & strCONCAT & "' ( '" & strCONCAT & gstrISNULL("BA.strDescricao", "''") & strCONCAT & "' ) '" & " AS Logradouro "
    strsql = strsql & "FROM " & gstrLogradouro & " L, "
    strsql = strsql & gstrTituloLogradouro & " U, "
    strsql = strsql & gstrTipoLogradouro & " TL, "
    strsql = strsql & gstrBairro & " BA "
    strsql = strsql & " WHERE L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle
    If Not blnConsulta Then
        strsql = strsql & " AND L.Dtmdtexclusao is null "
    End If
    strsql = strsql & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
    strsql = strsql & " AND L.intBairro " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle
    strsql = strsql & " ORDER BY L.strDescricao "
    
    strQueryLogradouro = strsql
        
End Function

Private Sub GravaCancelamento()
    Dim strsql      As String
    Dim blnCancel   As Boolean
    
    strsql = ""
    strsql = strsql & "Update " & gstrImobiliario & " Set dtmdtcancelamento = "
    
    If Trim(txtdtmdtcancelamento) = "" Then
        strsql = strsql & gstrConvDtParaSql(gstrDataDoSistema)
        blnCancel = True
    Else
        strsql = strsql & "Null"
        blnCancel = False
    End If
     
    strsql = strsql & " Where Pkid = " & txtPKId
    
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    If gobjBanco.Execute(strsql) Then
        gobjBanco.ExecutaCommitTrans
        
        
        If blnCancel Then
            txtdtmdtcancelamento = gstrDataFormatada(gstrDataDoSistema)
            HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
            
            If App.ProductName = "Tributario" Then
                MDIMenu.actBarra.Bands(gstrBtnArquivo).Tools.Item(20).ToolTipText = "Reativar"
            End If
        Else
            txtdtmdtcancelamento = ""
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
            
            If App.ProductName = "Tributario" Then
                MDIMenu.actBarra.Bands(gstrBtnArquivo).Tools.Item(20).ToolTipText = "Cancelar"
            End If
        End If
    Else
        ExibeMensagem "Não foi possível gravar o cancelamento"
        gobjBanco.ExecutaRollbackTrans
    End If
    
End Sub

Private Function VerificaEdificacao() As Boolean
    Dim xTerreno    As XArrayDB
    Dim intFor      As Integer

    VerificaEdificacao = False
    
    Set xTerreno = New XArrayDB
    xTerreno.ReDim 0, 0, 0, 8
    xTerreno(0, 0) = ""
    xTerreno(0, 1) = ""
    xTerreno(0, 2) = ""
    xTerreno(0, 3) = ""
    xTerreno(0, 4) = ""
    xTerreno(0, 5) = ""
    xTerreno(0, 6) = ""
    xTerreno(0, 7) = ""
    xTerreno(0, 8) = ""
    xTerreno.Clear
    grd_Area.Update
    
    Set xTerreno = grd_Area.Array
    For intFor = 0 To xTerreno.Count(1) - 1
        If xTerreno(intFor, 2) <> "" And xTerreno(intFor, 3) <> "" And xTerreno(intFor, 4) <> "" And _
           xTerreno(intFor, 5) <> "" And xTerreno(intFor, 6) <> "" And xTerreno(intFor, 7) <> "" And _
           xTerreno(intFor, 8) <> "" Then 'Verifica se todas as colunas estão preenchidas
           If gblnDataValida(xTerreno(intFor, 8)) = False Then 'Se estiver valida a data
                ExibeMensagem "O valor da coluna última reforma na guia de Construção está incorreto."
                Exit Function
           End If
           
        Else
            If xTerreno(intFor, 2) = "" And xTerreno(intFor, 3) = "" And xTerreno(intFor, 4) = "" And _
               Trim(xTerreno(intFor, 5)) = "" And xTerreno(intFor, 7) = "" And xTerreno(intFor, 8) = "" Then
               'Esse filtro verifica se o usuario preencheu a coluna mas apagou posteriormente
               'Nesse caso o texto é apagado mas o ID da coluna não
               
            Else
                If xTerreno(intFor, 2) = "" Then
                    ExibeMensagem "É necessário preencher a coluna tipo de área na guia de Construção."
                    Exit Function
                ElseIf xTerreno(intFor, 3) = "" Then
                    ExibeMensagem "É necessário preencher a coluna medida da área na guia de Construção."
                    Exit Function
                ElseIf xTerreno(intFor, 4) = "" Then
                    ExibeMensagem "É necessário preencher a coluna fração ideal na guia de Construção."
                    Exit Function
                ElseIf xTerreno(intFor, 6) = "" Then
                    ExibeMensagem "É necessário preencher a coluna categoria da construção na guia de Construção."
                    Exit Function
                ElseIf xTerreno(intFor, 7) = "" Then
                    ExibeMensagem "É necessário preencher a coluna nº de pavimentos na guia de Construção."
                    Exit Function
                ElseIf xTerreno(intFor, 8) = "" Then
                    ExibeMensagem "É necessário preencher a coluna última reforma na guia de Construção."
                    Exit Function
                End If
            End If
        End If
    Next
    VerificaEdificacao = True
End Function

Private Function RetornaEdificacao() As Byte
Dim xTerreno    As XArrayDB
Dim intFor      As Integer
    
    Set xTerreno = New XArrayDB
    xTerreno.ReDim 0, 0, 0, 8
    xTerreno(0, 0) = ""
    xTerreno(0, 1) = ""
    xTerreno(0, 2) = ""
    xTerreno(0, 3) = ""
    xTerreno(0, 4) = ""
    xTerreno(0, 5) = ""
    xTerreno(0, 6) = ""
    xTerreno(0, 7) = ""
    xTerreno(0, 8) = ""
    xTerreno.Clear
    grd_Area.Update
    
    Set xTerreno = grd_Area.Array
    For intFor = 0 To xTerreno.Count(1) - 1
        If xTerreno(intFor, 2) <> "" And xTerreno(intFor, 3) <> "" And xTerreno(intFor, 4) <> "" And _
           xTerreno(intFor, 5) <> "" And xTerreno(intFor, 6) <> "" And xTerreno(intFor, 7) <> "" And _
           xTerreno(intFor, 8) <> "" Then
           RetornaEdificacao = 1
        Else
            RetornaEdificacao = 0
        End If
    Next
End Function

Private Function strQueryAplicar() As String
    
    Dim strsql As String
    
    strsql = "SELECT Pkid,  " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao "
    strsql = strsql & "FROM " & gstrImobiliario
    strsql = strsql & " ORDER BY strInscricao "
    
    strQueryAplicar = strsql
    
End Function

Private Function strQueryResumoTipoPadrao(strPKId As String) As String
    Dim strsql As String
    
    strsql = "Select Pkid, strcodigo From " & gstrResumoTipoPadrao & " Where Intcategoriaconstrucao = " & strPKId & " Order By strCodigo"
    strQueryResumoTipoPadrao = strsql
    
End Function

Private Sub GravaCaracteristicaResumo()
    Dim strsql      As String
    Dim intAux       As Long
    Dim intEdificado As Integer
    Screen.MousePointer = vbArrow
    
    If Trim(txtdtmdtcancelamento) <> "" Then
        ExibeMensagem "Não foi possível gravar características para resumo. Imóvel está cancelado."
        Screen.MousePointer = vbDefault
        Exit Sub
    ElseIf Val(grd_Area.Columns("Pkid").Value) <= 0 Then
        Exit Sub
    End If
    
    'Caso não tenha salvo o imobiliário vamos forçar a salva, para que possa ser salva as características
    If Val(txtPKId) = 0 Then
        If blnDadosOK(gstrSalvar, True) Then
            intEdificado = chkbytEdificado.Value
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            If SalvarGeral(gstrImobiliario, IIf(mblnAlterando, "A", "I"), Me, tdb_Lista, strQueryListView(gstrSalvar), False, False, True) Then
                If intAux = 0 Then
                    intAux = PegaMaxPKId
                End If
                GravaHistoricos intAux, False
                GravaValores intAux, intEdificado, False
                GravaValores2 intAux, False
                GravaEnvolvidos intAux, False
                
                txtPKId = intAux
                
                mblnPrimeiraVez = False
                mblnAlterando = True
                
                blnEmTransacao = True
            Else
                gobjBanco.ExecutaRollbackTrans
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    Else
        intAux = txtPKId
    End If
    

    'Vamos deletar todas as caracterisitcas do prédio selecionado
    strsql = ""
    strsql = IIf(bytDBType = Oracle, "Begin ", "")
    strsql = strsql & " DELETE "
    strsql = strsql & " FROM " & gstrCaracteristicaDoImovel
    strsql = strsql & " WHERE "
    strsql = strsql & " intCodigoImobiliario = " & Val(intAux)                  'Código do imobiliário
    strsql = strsql & " AND intArea = '" & grd_Area.Columns("Pkid").Value & "'" 'Código do Prédio
    strsql = strsql & " AND intCodigoUtilizacaoDaTabelaDeV = 3 "                'Código da utilização Imobiliário Construção
    strsql = strsql & IIf(bytDBType = Oracle, ";", " ")
    
    'Vamos inserir o cadastro simplificado das carcacterísticas do prédio selecionado
    strsql = strsql & "Insert Into " & gstrCaracteristicaDoImovel & "( "
    strsql = strsql & "intcodigoimobiliario, "
    strsql = strsql & "intcodigocaracteristicageral, "
    strsql = strsql & "intcodigodetalhedacaracteristi, "
    strsql = strsql & "intcodigoutilizacaodatabeladev, "
    strsql = strsql & "intarea) "
    strsql = strsql & "(Select "
    strsql = strsql & intAux & " As intcodigoimobiliario, "             'Código do imobiliário
    strsql = strsql & "CG.Pkid As intcodigocaracteristicageral, "       'Código da Característica geral
    strsql = strsql & "DC.pkid As intcodigodetalhedacaracteristi, "     'Código do Detalhe da característica
    strsql = strsql & "3 As intcodigoutilizacaodatabeladev, "           'Código da utilização
    strsql = strsql & grd_Area.Columns("Pkid").Value & " As intarea "   'Código do Prédio
    strsql = strsql & "From "
    strsql = strsql & gstrResumoTipoPadrao & " RP, "
    strsql = strsql & gstrResumoTipoPadraoCarac & " RPC, "
    strsql = strsql & gstrCategoriaConstrucao & " CC, "
    strsql = strsql & gstrCaracteristicaGeral & " CG, "
    strsql = strsql & gstrDetalheDaCaracteristica & " DC "
    strsql = strsql & "Where "
    strsql = strsql & "RP.Pkid = RPC.INTRESUMOTIPOPADRAO AND "
    strsql = strsql & "CG.Pkid = RPC.INTCARACTERISTICAGERAL AND "
    strsql = strsql & "CC.Pkid = CG.Intcategoriaconstrucao AND "
    strsql = strsql & "DC.Pkid = RPC.INTDETALHEDACARACTERISTICA AND "
    strsql = strsql & "RP.pkid = " & dbc_intResumoTipoPadrao.BoundText & " )"
    strsql = strsql & IIf(bytDBType = Oracle, ";", " ")
    strsql = strsql & IIf(bytDBType = Oracle, "End;", "")
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    If Not gobjBanco.Execute(strsql) Then
        ExibeMensagem "Não foi possível gravar características para resumo"
        gobjBanco.ExecutaRollbackTrans
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        CarregaConstrucao
    End If
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Function blnVerificaLogradouro(lngIntLogradouro) As Boolean
    Dim strsql          As String
    Dim adoResultado    As ADODB.Recordset

    blnVerificaLogradouro = False
    strsql = ""
    strsql = strsql & "SELECT "
    strsql = strsql & "L.Pkid "
    strsql = strsql & "FROM "
    strsql = strsql & gstrLogradouro & " L "
    strsql = strsql & "WHERE "
    strsql = strsql & "L.Pkid = " & dbcintLogradouro.BoundText & " And "
    strsql = strsql & "not L.Dtmdtexclusao is null "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            blnVerificaLogradouro = True
        End If
    End If

End Function

Private Sub PreencheLogradouroC(lngPkid)
    Dim strsql          As String
    Dim adoResultado    As ADODB.Recordset

    strsql = ""
    strsql = strsql & "SELECT "
    strsql = strsql & "strlogradouroc "
    strsql = strsql & "FROM "
    strsql = strsql & gstrImobiliario & "  "
    strsql = strsql & "WHERE "
    strsql = strsql & "Pkid = " & lngPkid

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            dbcstrLogradouroC = gstrENulo(adoResultado!strlogradouroc)
        End If
    End If

End Sub

Private Sub CarregaFichasCadastrais()
Dim strsql       As String
Dim adoResultado As New ADODB.Recordset
    
    strsql = "SELECT intImobiliario, strCaminhoFicha FROM tblImobiliarioFichasCadastrais WHERE intImobiliario = " & tdb_Lista.Columns("PKID").Value
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
        
            intIndiceFichas = 0
            
            Set aFichasCadastrais = New XArrayDB
            aFichasCadastrais.Clear
            aFichasCadastrais.ReDim 0, adoResultado.RecordCount - 1, 0, 1
            
            Do While Not adoResultado.EOF
                aFichasCadastrais(adoResultado.AbsolutePosition - 1, 0) = adoResultado("intImobiliario").Value
                aFichasCadastrais(adoResultado.AbsolutePosition - 1, 1) = gstrDirDocumentos & "Documentos\Fichas" & Trim(adoResultado("strCaminhoFicha").Value) & ".rtf"
                adoResultado.MoveNext
            Loop
            
            MovimentaFichas intIndiceFichas
            
        Else
            cmd_FichaAnterior.Enabled = False
            cmd_FichaPosterior.Enabled = False
            'acr_FichaCadastral.Visible = False
            rtb_FichaCadastral.Text = ""
            lbl_CaminhoFicha.Caption = ""
        End If
    End If
    
End Sub

Private Sub MovimentaFichas(intIndice As Integer)
        
    On Error GoTo NaoExiste
    
    cmd_FichaAnterior.Enabled = True
    cmd_FichaPosterior.Enabled = True
    
    lbl_CaminhoFicha.Caption = ""
    
    If Dir(aFichasCadastrais(intIndice, 1), vbArchive) <> "" Then
        'acr_FichaCadastral.LoadFile aFichasCadastrais(intIndice, 1)
        rtb_FichaCadastral.filename = aFichasCadastrais(intIndice, 1)
        lbl_CaminhoFicha.Caption = aFichasCadastrais(intIndice, 1)
        'acr_FichaCadastral.Visible = True
    Else
        lbl_CaminhoFicha.Caption = "Arquivo não encontrado. (" & aFichasCadastrais(intIndice, 1) & ")"
        rtb_FichaCadastral.Text = ""
        'acr_FichaCadastral.Visible = False
    End If
    
    If intIndice = aFichasCadastrais.UpperBound(1) Then
        cmd_FichaPosterior.Enabled = False
    End If
    
    If intIndice = aFichasCadastrais.LowerBound(1) Then
        cmd_FichaAnterior.Enabled = False
    End If
    
    Exit Sub
    
NaoExiste:
    cmd_FichaAnterior.Enabled = False
    cmd_FichaPosterior.Enabled = False
    
    rtb_FichaCadastral.Text = ""
    lbl_CaminhoFicha.Caption = "Caminho não encontrado. (" & aFichasCadastrais(intIndice, 1) & ")"
End Sub
