VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadContribuinte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contribuintes"
   ClientHeight    =   7635
   ClientLeft      =   1410
   ClientTop       =   2460
   ClientWidth     =   9480
   HelpContextID   =   109
   Icon            =   "CadContribuinte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9480
   Begin TabDlg.SSTab tab_3DDadosGerais 
      Height          =   4725
      HelpContextID   =   13
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8334
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Dados Gerais"
      TabPicture(0)   =   "CadContribuinte.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrCNPJCPF"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrIdentidade"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrTituloEleitoral"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbldtmDataNascimento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblstrCarteiraTrabalho"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrInscricaoEstadual"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblstrNomeFantasia"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblstrNome"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbldtmDataCadastro"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_Codigo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblintCodigo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "mskstrPIS"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dbcstrNomeFantasia"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dbcstrNome"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkblnResidenteNoMunicipio"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtstrIdentidade"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtstrTituloEleitoral"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtdtmDataNascimento"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtstrCarteiraTrabalho"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtstrInscricaoEstadual"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "dtpdtmDataCadastro"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "mskstrCNPJCPF"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "fra_Linha1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "fra_Linha2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtPKId"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkblnInativo"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtintCodigo"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cbobytNaturezaJuridica"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "fra_Tipo"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Residência / Estabelecido"
      TabPicture(1)   =   "CadContribuinte.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_Codigo1"
      Tab(1).Control(1)=   "lbl_Nome2"
      Tab(1).Control(2)=   "tab_3DCorrespondencia"
      Tab(1).Control(3)=   "txt_Codigo1"
      Tab(1).Control(4)=   "txt_Nome1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Comunicação"
      TabPicture(2)   =   "CadContribuinte.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl_TipoComunicacao"
      Tab(2).Control(1)=   "lbl_DescricaoConteudo"
      Tab(2).Control(2)=   "lbl_Codigo2"
      Tab(2).Control(3)=   "lbl_Nome5"
      Tab(2).Control(4)=   "tlb_TipoComunicacao"
      Tab(2).Control(5)=   "img_Aux"
      Tab(2).Control(6)=   "lvw_TipoComunicacao"
      Tab(2).Control(7)=   "txt_Conteudo"
      Tab(2).Control(8)=   "txt_DescricaoConteudo"
      Tab(2).Control(9)=   "txt_Codigo2"
      Tab(2).Control(10)=   "txt_Nome2"
      Tab(2).Control(11)=   "cmd_Down"
      Tab(2).Control(12)=   "cmd_Up"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Contas Bancárias "
      TabPicture(3)   =   "CadContribuinte.frx":1096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lbl_Conta"
      Tab(3).Control(1)=   "lbl_Banco"
      Tab(3).Control(2)=   "lbl_Agencia"
      Tab(3).Control(3)=   "lbl_Codigo5"
      Tab(3).Control(4)=   "lbl_Nome3"
      Tab(3).Control(5)=   "lbl_DV"
      Tab(3).Control(6)=   "lvw_Contas"
      Tab(3).Control(7)=   "txt_Conta"
      Tab(3).Control(8)=   "txt_Codigo3"
      Tab(3).Control(9)=   "txt_Nome3"
      Tab(3).Control(10)=   "txt_DigitoVerificador"
      Tab(3).Control(11)=   "txt_CodBanco"
      Tab(3).Control(12)=   "txt_CodAgencia"
      Tab(3).Control(13)=   "txt_Banco"
      Tab(3).Control(14)=   "txt_Agencia"
      Tab(3).Control(15)=   "fra_Publica_debito"
      Tab(3).Control(16)=   "cmd_ContasBancarias"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "Histórico"
      TabPicture(4)   =   "CadContribuinte.frx":10B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lbl_Codigo4"
      Tab(4).Control(1)=   "lbl_Nome4"
      Tab(4).Control(2)=   "tdb_Historico"
      Tab(4).Control(3)=   "txt_Codigo4"
      Tab(4).Control(4)=   "txt_Nome4"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Sócios"
      TabPicture(5)   =   "CadContribuinte.frx":10CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fra_Socios"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Aplicações"
      TabPicture(6)   =   "CadContribuinte.frx":10EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lvw_Aplicacoes"
      Tab(6).ControlCount=   1
      Begin VB.Frame fra_Tipo 
         Height          =   735
         Left            =   7080
         TabIndex        =   14
         Top             =   1320
         Width           =   2145
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Fornecedor"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   16
            Top             =   450
            Width           =   1635
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Credor"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   180
            Value           =   -1  'True
            Width           =   1665
         End
      End
      Begin VB.ComboBox cbobytNaturezaJuridica 
         Height          =   315
         ItemData        =   "CadContribuinte.frx":1106
         Left            =   1605
         List            =   "CadContribuinte.frx":1108
         TabIndex        =   7
         Top             =   1020
         Width           =   3795
      End
      Begin VB.TextBox txtintCodigo 
         Height          =   285
         Left            =   4155
         MaxLength       =   10
         TabIndex        =   4
         Top             =   690
         Width           =   1485
      End
      Begin VB.CommandButton cmd_ContasBancarias 
         Height          =   300
         Left            =   -68220
         Picture         =   "CadContribuinte.frx":110A
         Style           =   1  'Graphical
         TabIndex        =   118
         TabStop         =   0   'False
         Tag             =   "193"
         ToolTipText     =   "Clique para cadastrar objetivo"
         Top             =   1650
         Width           =   330
      End
      Begin VB.CheckBox chkblnInativo 
         Caption         =   "Inativo"
         Height          =   210
         Left            =   8460
         TabIndex        =   5
         Top             =   765
         Width           =   810
      End
      Begin MSComctlLib.ListView lvw_Aplicacoes 
         Height          =   3570
         Left            =   -74700
         TabIndex        =   141
         Top             =   690
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   6297
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "img_Check"
         SmallIcons      =   "img_Check"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrição"
            Object.Width           =   15223
         EndProperty
      End
      Begin VB.Frame fra_Socios 
         Caption         =   " Sócios "
         Height          =   3045
         Left            =   -74880
         TabIndex        =   134
         Top             =   360
         Width           =   9165
         Begin VB.TextBox txt_TotalDeCotas 
            Height          =   285
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   137
            Top             =   2640
            Width           =   1545
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Socios 
            Height          =   2265
            Left            =   120
            TabIndex        =   135
            Top             =   270
            Width           =   8925
            _ExtentX        =   15743
            _ExtentY        =   3995
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nome do Sócio"
            Columns(0).DataField=   ""
            Columns(0).DropDown=   "tdd_Socios"
            Columns(0).DropDown.vt=   8
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "CNPJ / CPF"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Número de Cotas"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=9102"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=9022"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(0).AutoDropDown=1"
            Splits(0)._ColumnProps(8)=   "Column(0).DropDownList=1"
            Splits(0)._ColumnProps(9)=   "Column(0).AutoCompletion=1"
            Splits(0)._ColumnProps(10)=   "Column(1).Width=3678"
            Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=3598"
            Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
            Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(17)=   "Column(2).Width=2752"
            Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2672"
            Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(2).AllowFocus=0"
            Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
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
            DirectionAfterEnter=   0
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
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
            _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
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
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.locked=-1"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
            _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2,.wraptext=0"
            _StyleDefs(60)  =   "Named:id=38:HighlightRow"
            _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(62)  =   "Named:id=39:EvenRow"
            _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(64)  =   "Named:id=40:OddRow"
            _StyleDefs(65)  =   ":id=40,.parent=33"
            _StyleDefs(66)  =   "Named:id=41:RecordSelector"
            _StyleDefs(67)  =   ":id=41,.parent=34"
            _StyleDefs(68)  =   "Named:id=42:FilterBar"
            _StyleDefs(69)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lbl_Cotas 
            AutoSize        =   -1  'True
            Caption         =   "Total de cotas"
            Height          =   195
            Left            =   6390
            TabIndex        =   136
            Top             =   2730
            Width           =   1020
         End
      End
      Begin VB.TextBox txtPKId 
         Height          =   285
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   690
         Width           =   1605
      End
      Begin VB.CommandButton cmd_Up 
         Enabled         =   0   'False
         Height          =   585
         Left            =   -66270
         Picture         =   "CadContribuinte.frx":1494
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   2235
         Width           =   525
      End
      Begin VB.CommandButton cmd_Down 
         Enabled         =   0   'False
         Height          =   585
         Left            =   -66270
         Picture         =   "CadContribuinte.frx":179E
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   2820
         Width           =   525
      End
      Begin VB.Frame fra_Publica_debito 
         Enabled         =   0   'False
         Height          =   885
         Left            =   -73440
         TabIndex        =   123
         Top             =   2400
         Width           =   5865
         Begin VB.TextBox txt_dtmDebito 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3660
            Locked          =   -1  'True
            TabIndex        =   127
            Top             =   450
            Width           =   1185
         End
         Begin VB.CheckBox chk_ContaPublica 
            Caption         =   "Conta Pública"
            Height          =   195
            Left            =   210
            TabIndex        =   124
            Top             =   210
            Width           =   1365
         End
         Begin VB.CheckBox chk_DebitoAutomatico 
            Caption         =   "Débito Automático"
            Height          =   195
            Left            =   210
            TabIndex        =   125
            Top             =   540
            Width           =   1725
         End
         Begin VB.Label lbl_dtmDebito 
            AutoSize        =   -1  'True
            Caption         =   "Data Início Débito"
            Height          =   195
            Left            =   2250
            TabIndex        =   126
            Top             =   540
            Width           =   1305
         End
      End
      Begin VB.TextBox txt_Agencia 
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
         Left            =   -72780
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   1650
         Width           =   4500
      End
      Begin VB.TextBox txt_Banco 
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
         Left            =   -72780
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   1230
         Width           =   4500
      End
      Begin VB.TextBox txt_CodAgencia 
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
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   1650
         Width           =   660
      End
      Begin VB.TextBox txt_CodBanco 
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
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   1230
         Width           =   660
      End
      Begin VB.TextBox txt_DigitoVerificador 
         Height          =   285
         Left            =   -70860
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   2040
         Width           =   525
      End
      Begin VB.Frame fra_Linha2 
         Height          =   75
         Left            =   90
         TabIndex        =   144
         Top             =   2880
         Width           =   6675
      End
      Begin VB.Frame fra_Linha1 
         Height          =   75
         Left            =   90
         TabIndex        =   142
         Top             =   2070
         Width           =   6675
      End
      Begin VB.TextBox txt_Nome4 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -74025
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   132
         Top             =   960
         Width           =   4905
      End
      Begin VB.TextBox txt_Codigo4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
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
         Left            =   -74025
         Locked          =   -1  'True
         MaxLength       =   19
         MultiLine       =   -1  'True
         TabIndex        =   130
         Top             =   570
         Width           =   1215
      End
      Begin VB.TextBox txt_Nome3 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   111
         Top             =   840
         Width           =   5145
      End
      Begin VB.TextBox txt_Codigo3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
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
         Left            =   -73440
         Locked          =   -1  'True
         MaxLength       =   19
         MultiLine       =   -1  'True
         TabIndex        =   109
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox txt_Nome2 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -72690
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   102
         Top             =   990
         Width           =   4905
      End
      Begin VB.TextBox txt_Codigo2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
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
         Left            =   -72690
         Locked          =   -1  'True
         MaxLength       =   19
         MultiLine       =   -1  'True
         TabIndex        =   100
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txt_Nome1 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   33
         Top             =   1020
         Width           =   4905
      End
      Begin VB.TextBox txt_Codigo1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
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
         Left            =   -73290
         Locked          =   -1  'True
         MaxLength       =   19
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   630
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskstrCNPJCPF 
         Height          =   285
         Left            =   1605
         TabIndex        =   11
         Top             =   1755
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "0"
         PromptChar      =   " "
      End
      Begin MSComCtl2.DTPicker dtpdtmDataCadastro 
         Height          =   285
         Left            =   8010
         TabIndex        =   139
         Top             =   2550
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Format          =   69926913
         CurrentDate     =   36930
      End
      Begin VB.TextBox txt_Conta 
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   120
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txt_DescricaoConteudo 
         Height          =   285
         Left            =   -72690
         MaxLength       =   50
         TabIndex        =   106
         Top             =   1710
         Width           =   3945
      End
      Begin VB.TextBox txt_Conteudo 
         Height          =   285
         Left            =   -72690
         MaxLength       =   50
         TabIndex        =   104
         Top             =   1365
         Width           =   3945
      End
      Begin VB.TextBox txtstrInscricaoEstadual 
         Height          =   285
         Left            =   1605
         MaxLength       =   20
         TabIndex        =   21
         Top             =   2550
         Width           =   1605
      End
      Begin VB.TextBox txtstrCarteiraTrabalho 
         Height          =   285
         Left            =   1605
         MaxLength       =   20
         TabIndex        =   29
         Top             =   4020
         Width           =   1605
      End
      Begin VB.TextBox txtdtmDataNascimento 
         Height          =   285
         Left            =   1605
         TabIndex        =   27
         Top             =   3690
         Width           =   1605
      End
      Begin VB.TextBox txtstrTituloEleitoral 
         Height          =   285
         Left            =   1605
         MaxLength       =   20
         TabIndex        =   25
         Top             =   3360
         Width           =   1605
      End
      Begin VB.TextBox txtstrIdentidade 
         Height          =   285
         Left            =   1605
         MaxLength       =   20
         TabIndex        =   23
         Top             =   3030
         Width           =   1605
      End
      Begin VB.CheckBox chkblnResidenteNoMunicipio 
         Caption         =   "Residente no município"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   7080
         TabIndex        =   19
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin TabDlg.SSTab tab_3DCorrespondencia 
         Height          =   2445
         Left            =   -74910
         TabIndex        =   34
         Top             =   1560
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   4313
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Endereço residencial"
         TabPicture(0)   =   "CadContribuinte.frx":1AA8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblintMunicipio"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblintBairro"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblintLogradouro"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblintNumero"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblstrComplemento"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblintUF"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblintCep"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "dbcintLogradouro"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtintNumero"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtstrComplemento"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtintCep"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txt_strBairro"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txt_strMunicipio"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txt_strUF"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "cmd_Logradouro"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "Endereço de correspondência"
         TabPicture(1)   =   "CadContribuinte.frx":1AC4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblintCepC"
         Tab(1).Control(1)=   "lblintUFC"
         Tab(1).Control(2)=   "lblstrComplementoC"
         Tab(1).Control(3)=   "lblintNumeroC"
         Tab(1).Control(4)=   "lblintLogradouroC"
         Tab(1).Control(5)=   "lblintBairroC"
         Tab(1).Control(6)=   "lblintMunicipioC"
         Tab(1).Control(7)=   "lblstrDistritoC"
         Tab(1).Control(8)=   "dbcstrLogradouroC"
         Tab(1).Control(9)=   "dbcintTituloLogradouro"
         Tab(1).Control(10)=   "dbcintTipoLogradouro"
         Tab(1).Control(11)=   "dbcintUFC"
         Tab(1).Control(12)=   "dbcintMunicipioC"
         Tab(1).Control(13)=   "txtintCepC"
         Tab(1).Control(14)=   "txtstrComplementoC"
         Tab(1).Control(15)=   "txtintNumeroC"
         Tab(1).Control(16)=   "txtstrBairroC"
         Tab(1).Control(17)=   "txtstrDistritoC"
         Tab(1).Control(18)=   "txtintCodigoLogradouro"
         Tab(1).Control(19)=   "cmd_TipoLogradouro"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "cmd_TituloLogradouro"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "cmd_MunicipioC"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).ControlCount=   22
         TabCaption(2)   =   "Domicílio Fiscal"
         TabPicture(2)   =   "CadContribuinte.frx":1AE0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblintCepD"
         Tab(2).Control(1)=   "lblintUFD"
         Tab(2).Control(2)=   "lblstrComplementoD"
         Tab(2).Control(3)=   "lblintNumeroD"
         Tab(2).Control(4)=   "lblintLogradouroD"
         Tab(2).Control(5)=   "lblintBairroD"
         Tab(2).Control(6)=   "lblintMunicipioD"
         Tab(2).Control(7)=   "lblstrSetorD"
         Tab(2).Control(8)=   "lblstrQuadraD"
         Tab(2).Control(9)=   "lblstrLoteD"
         Tab(2).Control(10)=   "lblstrDistritoD"
         Tab(2).Control(11)=   "dbcintTituloLogradouroD"
         Tab(2).Control(12)=   "dbcintTipoLogradouroD"
         Tab(2).Control(13)=   "dbcintUFD"
         Tab(2).Control(14)=   "txtintCepD"
         Tab(2).Control(15)=   "txtstrComplementoD"
         Tab(2).Control(16)=   "txtintNumeroD"
         Tab(2).Control(17)=   "txtstrBairroD"
         Tab(2).Control(18)=   "txtstrMunicipioD"
         Tab(2).Control(19)=   "txtstrSetorD"
         Tab(2).Control(20)=   "txtstrQuadraD"
         Tab(2).Control(21)=   "txtstrLoteD"
         Tab(2).Control(22)=   "txtintCodigoLogradouroD"
         Tab(2).Control(23)=   "txtstrDistritoD"
         Tab(2).Control(24)=   "txtstrLogradouroD"
         Tab(2).Control(25)=   "cmd_TipoLogradouroD"
         Tab(2).Control(25).Enabled=   0   'False
         Tab(2).Control(26)=   "cmd_TituloLogradouroD"
         Tab(2).Control(26).Enabled=   0   'False
         Tab(2).ControlCount=   27
         Begin VB.CommandButton cmd_TituloLogradouroD 
            Height          =   300
            Left            =   -71310
            Picture         =   "CadContribuinte.frx":1AFC
            Style           =   1  'Graphical
            TabIndex        =   76
            TabStop         =   0   'False
            Tag             =   "193"
            ToolTipText     =   "Clique para cadastrar objetivo"
            Top             =   540
            Width           =   330
         End
         Begin VB.CommandButton cmd_TipoLogradouroD 
            Height          =   300
            Left            =   -73110
            Picture         =   "CadContribuinte.frx":1E86
            Style           =   1  'Graphical
            TabIndex        =   74
            TabStop         =   0   'False
            Tag             =   "193"
            ToolTipText     =   "Clique para cadastrar objetivo"
            Top             =   540
            Width           =   330
         End
         Begin VB.CommandButton cmd_MunicipioC 
            Height          =   300
            Left            =   -70290
            Picture         =   "CadContribuinte.frx":2210
            Style           =   1  'Graphical
            TabIndex        =   67
            TabStop         =   0   'False
            Tag             =   "193"
            ToolTipText     =   "Clique para cadastrar objetivo"
            Top             =   1530
            Width           =   330
         End
         Begin VB.CommandButton cmd_TituloLogradouro 
            Height          =   300
            Left            =   -71250
            Picture         =   "CadContribuinte.frx":259A
            Style           =   1  'Graphical
            TabIndex        =   56
            TabStop         =   0   'False
            Tag             =   "193"
            ToolTipText     =   "Clique para cadastrar objetivo"
            Top             =   750
            Width           =   330
         End
         Begin VB.CommandButton cmd_TipoLogradouro 
            Height          =   300
            Left            =   -73080
            Picture         =   "CadContribuinte.frx":2924
            Style           =   1  'Graphical
            TabIndex        =   54
            TabStop         =   0   'False
            Tag             =   "193"
            ToolTipText     =   "Clique para cadastrar objetivo"
            Top             =   750
            Width           =   330
         End
         Begin VB.CommandButton cmd_Logradouro 
            Height          =   300
            Left            =   5820
            Picture         =   "CadContribuinte.frx":2CAE
            Style           =   1  'Graphical
            TabIndex        =   37
            TabStop         =   0   'False
            Tag             =   "193"
            ToolTipText     =   "Clique para cadastrar objetivo"
            Top             =   570
            Width           =   330
         End
         Begin VB.TextBox txtstrLogradouroD 
            Height          =   285
            Left            =   -70140
            TabIndex        =   78
            Top             =   540
            Width           =   3840
         End
         Begin VB.TextBox txt_strUF 
            Height          =   315
            Left            =   5535
            TabIndex        =   47
            Top             =   1395
            Width           =   705
         End
         Begin VB.TextBox txt_strMunicipio 
            Height          =   315
            Left            =   1065
            TabIndex        =   45
            Top             =   1380
            Width           =   3975
         End
         Begin VB.TextBox txt_strBairro 
            Height          =   315
            Left            =   1065
            TabIndex        =   43
            Top             =   975
            Width           =   3975
         End
         Begin VB.TextBox txtstrDistritoD 
            Height          =   285
            Left            =   -73860
            MaxLength       =   50
            TabIndex        =   98
            Top             =   1950
            Width           =   3525
         End
         Begin VB.TextBox txtintCodigoLogradouroD 
            Height          =   285
            Left            =   -70950
            MaxLength       =   8
            TabIndex        =   77
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtintCodigoLogradouro 
            Height          =   315
            Left            =   -70860
            MaxLength       =   8
            TabIndex        =   57
            Top             =   750
            Width           =   735
         End
         Begin VB.TextBox txtstrDistritoC 
            Height          =   285
            Left            =   -73860
            MaxLength       =   50
            TabIndex        =   71
            Top             =   1950
            Width           =   3525
         End
         Begin VB.TextBox txtstrLoteD 
            Height          =   285
            Left            =   -67800
            MaxLength       =   20
            TabIndex        =   96
            Top             =   1605
            Width           =   1500
         End
         Begin VB.TextBox txtstrQuadraD 
            Height          =   285
            Left            =   -70920
            MaxLength       =   20
            TabIndex        =   94
            Top             =   1605
            Width           =   1500
         End
         Begin VB.TextBox txtstrSetorD 
            Height          =   285
            Left            =   -73860
            MaxLength       =   20
            TabIndex        =   92
            Top             =   1605
            Width           =   1500
         End
         Begin VB.TextBox txtstrBairroC 
            Height          =   285
            Left            =   -69375
            MaxLength       =   50
            TabIndex        =   64
            Top             =   1170
            Width           =   3375
         End
         Begin VB.TextBox txtstrMunicipioD 
            Height          =   285
            Left            =   -73860
            MaxLength       =   50
            TabIndex        =   86
            Top             =   1260
            Width           =   3525
         End
         Begin VB.TextBox txtstrBairroD 
            Height          =   285
            Left            =   -69525
            MaxLength       =   50
            TabIndex        =   84
            Top             =   900
            Width           =   3225
         End
         Begin VB.TextBox txtintCep 
            Height          =   315
            Left            =   6840
            MaxLength       =   9
            TabIndex        =   49
            Top             =   1395
            Width           =   1080
         End
         Begin VB.TextBox txtstrComplemento 
            Height          =   285
            Left            =   7800
            MaxLength       =   20
            TabIndex        =   41
            Top             =   600
            Width           =   1350
         End
         Begin VB.TextBox txtintNumero 
            Height          =   285
            Left            =   6420
            MaxLength       =   8
            TabIndex        =   39
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtintNumeroC 
            Height          =   285
            Left            =   -73860
            MaxLength       =   8
            TabIndex        =   60
            Top             =   1170
            Width           =   855
         End
         Begin VB.TextBox txtstrComplementoC 
            Height          =   285
            Left            =   -72180
            MaxLength       =   20
            TabIndex        =   62
            Top             =   1170
            Width           =   1260
         End
         Begin VB.TextBox txtintCepC 
            Height          =   285
            Left            =   -73860
            MaxLength       =   9
            TabIndex        =   51
            Top             =   375
            Width           =   1080
         End
         Begin VB.TextBox txtintNumeroD 
            Height          =   285
            Left            =   -73860
            MaxLength       =   8
            TabIndex        =   80
            Top             =   900
            Width           =   915
         End
         Begin VB.TextBox txtstrComplementoD 
            Height          =   285
            Left            =   -71910
            MaxLength       =   20
            TabIndex        =   82
            Top             =   900
            Width           =   1320
         End
         Begin VB.TextBox txtintCepD 
            Height          =   285
            Left            =   -67380
            MaxLength       =   9
            TabIndex        =   90
            Top             =   1260
            Width           =   1080
         End
         Begin MSDataListLib.DataCombo dbcintLogradouro 
            Height          =   315
            Left            =   1050
            TabIndex        =   36
            Top             =   570
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintMunicipioC 
            Height          =   315
            Left            =   -73860
            TabIndex        =   66
            Top             =   1530
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintUFC 
            Height          =   315
            Left            =   -69375
            TabIndex        =   69
            Top             =   1545
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintUFD 
            Height          =   315
            Left            =   -69525
            TabIndex        =   88
            Top             =   1215
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTipoLogradouro 
            Height          =   315
            Left            =   -73860
            TabIndex        =   53
            Top             =   750
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTituloLogradouro 
            Height          =   315
            Left            =   -72720
            TabIndex        =   55
            Top             =   750
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTipoLogradouroD 
            Height          =   315
            Left            =   -73860
            TabIndex        =   73
            Top             =   525
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTituloLogradouroD 
            Height          =   315
            Left            =   -72750
            TabIndex        =   75
            Top             =   525
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcstrLogradouroC 
            Height          =   315
            Left            =   -70110
            TabIndex        =   58
            Top             =   750
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label lblstrDistritoD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   -74520
            TabIndex        =   97
            Top             =   2025
            Width           =   480
         End
         Begin VB.Label lblstrDistritoC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   -74520
            TabIndex        =   70
            Top             =   2040
            Width           =   480
         End
         Begin VB.Label lblstrLoteD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            Height          =   195
            Left            =   -68295
            TabIndex        =   95
            Top             =   1695
            Width           =   315
         End
         Begin VB.Label lblstrQuadraD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Quadra"
            Height          =   195
            Left            =   -71595
            TabIndex        =   93
            Top             =   1695
            Width           =   525
         End
         Begin VB.Label lblstrSetorD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Setor"
            Height          =   195
            Left            =   -74415
            TabIndex        =   91
            Top             =   1695
            Width           =   375
         End
         Begin VB.Label lblintCep 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   6495
            TabIndex        =   48
            Top             =   1485
            Width           =   285
         End
         Begin VB.Label lblintUF 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   5130
            TabIndex        =   46
            Top             =   1485
            Width           =   210
         End
         Begin VB.Label lblstrComplemento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   7320
            TabIndex        =   40
            Top             =   690
            Width           =   480
         End
         Begin VB.Label lblintNumero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   6210
            TabIndex        =   38
            Top             =   690
            Width           =   180
         End
         Begin VB.Label lblintLogradouro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   690
            Width           =   810
         End
         Begin VB.Label lblintBairro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   540
            TabIndex        =   42
            Top             =   1080
            Width           =   405
         End
         Begin VB.Label lblintMunicipio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   255
            TabIndex        =   44
            Top             =   1485
            Width           =   705
         End
         Begin VB.Label lblintMunicipioC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   -74745
            TabIndex        =   65
            Top             =   1665
            Width           =   705
         End
         Begin VB.Label lblintBairroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   -69930
            TabIndex        =   63
            Top             =   1260
            Width           =   405
         End
         Begin VB.Label lblintLogradouroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   -74790
            TabIndex        =   52
            Top             =   840
            Width           =   810
         End
         Begin VB.Label lblintNumeroC 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   -74190
            TabIndex        =   59
            Top             =   1260
            Width           =   180
         End
         Begin VB.Label lblstrComplementoC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   -72780
            TabIndex        =   61
            Top             =   1260
            Width           =   480
         End
         Begin VB.Label lblintUFC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   -69720
            TabIndex        =   68
            Top             =   1650
            Width           =   210
         End
         Begin VB.Label lblintCepC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   -74205
            TabIndex        =   50
            Top             =   465
            Width           =   285
         End
         Begin VB.Label lblintMunicipioD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   -74745
            TabIndex        =   85
            Top             =   1350
            Width           =   705
         End
         Begin VB.Label lblintBairroD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   -70050
            TabIndex        =   83
            Top             =   960
            Width           =   405
         End
         Begin VB.Label lblintLogradouroD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   -74850
            TabIndex        =   72
            Top             =   660
            Width           =   810
         End
         Begin VB.Label lblintNumeroD 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   -74220
            TabIndex        =   79
            Top             =   990
            Width           =   180
         End
         Begin VB.Label lblstrComplementoD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   -72510
            TabIndex        =   81
            Top             =   990
            Width           =   480
         End
         Begin VB.Label lblintUFD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   -69840
            TabIndex        =   87
            Top             =   1350
            Width           =   210
         End
         Begin VB.Label lblintCepD 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   -67815
            TabIndex        =   89
            Top             =   1350
            Width           =   285
         End
      End
      Begin MSComctlLib.ListView lvw_TipoComunicacao 
         Height          =   2040
         Left            =   -74580
         TabIndex        =   107
         Top             =   2220
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3598
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrição"
            Object.Width           =   52917
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Contas 
         Height          =   1260
         Left            =   -74910
         TabIndex        =   128
         Top             =   3390
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   2223
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrição"
            Object.Width           =   52917
         EndProperty
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Historico 
         Height          =   2295
         Left            =   -74940
         TabIndex        =   133
         Top             =   1470
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4048
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKID"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "strCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data / Hora"
         Columns(2).DataField=   "dtmDataHora"
         Columns(2).NumberFormat=   "General Date"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tipo de Transação"
         Columns(3).DataField=   "strTransacao"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nome do Sistema"
         Columns(4).DataField=   "strNomeSistema"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Valor"
         Columns(5).DataField=   "dblValor"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1984"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1984"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1905"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=2619"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2540"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=4842"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4763"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=4604"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=4524"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=2064"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1984"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
         MultiSelect     =   0
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
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H0&"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
      Begin MSDataListLib.DataCombo dbcstrNome 
         Height          =   315
         Left            =   1605
         TabIndex        =   9
         Top             =   1395
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcstrNomeFantasia 
         Height          =   315
         Left            =   1605
         TabIndex        =   18
         Top             =   2190
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComctlLib.ImageList img_Aux 
         Left            =   -66300
         Top             =   840
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
               Picture         =   "CadContribuinte.frx":3038
               Key             =   "Novo"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadContribuinte.frx":3192
               Key             =   "Salvar"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadContribuinte.frx":32EC
               Key             =   "Deletar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb_TipoComunicacao 
         Height          =   330
         Left            =   -68640
         TabIndex        =   146
         Top             =   1350
         Width           =   1035
         _ExtentX        =   1826
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
      Begin MSMask.MaskEdBox mskstrPIS 
         Height          =   285
         Left            =   5160
         TabIndex        =   13
         Top             =   1755
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "0"
         Mask            =   "###\.#####\.##\.#\ "
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Natureza"
         Height          =   195
         Left            =   900
         TabIndex        =   6
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Referência"
         Height          =   195
         Left            =   3330
         TabIndex        =   3
         Top             =   750
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PIS\PASEP"
         Height          =   195
         Left            =   4170
         TabIndex        =   12
         Top             =   1830
         Width           =   855
      End
      Begin VB.Label lbl_DV 
         AutoSize        =   -1  'True
         Caption         =   "DV"
         Height          =   195
         Left            =   -71190
         TabIndex        =   121
         Top             =   2130
         Width           =   225
      End
      Begin VB.Label lbl_Nome4 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   -74565
         TabIndex        =   131
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label lbl_Codigo4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   -74640
         TabIndex        =   129
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lbl_Nome3 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   -73980
         TabIndex        =   110
         Top             =   900
         Width           =   420
      End
      Begin VB.Label lbl_Codigo5 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   -74055
         TabIndex        =   108
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lbl_Nome5 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   -73245
         TabIndex        =   101
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label lbl_Codigo2 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   -73320
         TabIndex        =   99
         Top             =   690
         Width           =   495
      End
      Begin VB.Label lbl_Nome2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   -73905
         TabIndex        =   32
         Top             =   1110
         Width           =   420
      End
      Begin VB.Label lbl_Codigo1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   -73980
         TabIndex        =   30
         Top             =   750
         Width           =   495
      End
      Begin VB.Label lbl_Codigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1050
         TabIndex        =   1
         Top             =   780
         Width           =   495
      End
      Begin VB.Label lbldtmDataCadastro 
         AutoSize        =   -1  'True
         Caption         =   "Data de cadastro"
         Height          =   195
         Left            =   6750
         TabIndex        =   138
         Top             =   2640
         Width           =   1260
      End
      Begin VB.Label lbl_Agencia 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   -74145
         TabIndex        =   115
         Top             =   1740
         Width           =   585
      End
      Begin VB.Label lbl_Banco 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   -74025
         TabIndex        =   112
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lbl_Conta 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         Height          =   195
         Left            =   -73980
         TabIndex        =   119
         Top             =   2130
         Width           =   420
      End
      Begin VB.Label lblstrNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   1125
         TabIndex        =   8
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label lbl_DescricaoConteudo 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   -73545
         TabIndex        =   105
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label lbl_TipoComunicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   -73140
         TabIndex        =   103
         Top             =   1455
         Width           =   315
      End
      Begin VB.Label lblstrNomeFantasia 
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   2310
         Width           =   1065
      End
      Begin VB.Label lblstrInscricaoEstadual 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estadual"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   1305
      End
      Begin VB.Label lblstrCarteiraTrabalho 
         AutoSize        =   -1  'True
         Caption         =   "Carteira de Trabalho"
         Height          =   195
         Left            =   105
         TabIndex        =   28
         Top             =   4110
         Width           =   1440
      End
      Begin VB.Label lbldtmDataNascimento 
         AutoSize        =   -1  'True
         Caption         =   "Data de Nascimento"
         Height          =   195
         Left            =   90
         TabIndex        =   26
         Top             =   3780
         Width           =   1455
      End
      Begin VB.Label lblstrTituloEleitoral 
         AutoSize        =   -1  'True
         Caption         =   "Título Eleitoral"
         Height          =   195
         Left            =   525
         TabIndex        =   24
         Top             =   3450
         Width           =   1020
      End
      Begin VB.Label lblstrIdentidade 
         AutoSize        =   -1  'True
         Caption         =   "Identidade"
         Height          =   195
         Left            =   795
         TabIndex        =   22
         Top             =   3120
         Width           =   750
      End
      Begin VB.Label lblstrCNPJCPF 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ / CPF"
         Height          =   195
         Left            =   675
         TabIndex        =   10
         Top             =   1830
         Width           =   870
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   2745
      Left            =   60
      TabIndex        =   140
      Top             =   4800
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   4842
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
      Columns(1).Caption=   "Código"
      Columns(1).DataField=   "PKId"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nome"
      Columns(2).DataField=   "strNome"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   16
      Columns(3)._MaxComboItems=   5
      Columns(3).ValueItems(0)._DefaultItem=   0
      Columns(3).ValueItems(0).Value=   "0"
      Columns(3).ValueItems(0).Value.vt=   8
      Columns(3).ValueItems(0).DisplayValue=   "0"
      Columns(3).ValueItems(0).DisplayValue.vt=   8
      Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems.Count=   1
      Columns(3).Caption=   "CPF / CNPJ"
      Columns(3).DataField=   "strCNPJCPF"
      Columns(3).NumberFormat=   "FormatText Event"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Inativo"
      Columns(4).DataField=   "strInativo"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Nome Fantasia"
      Columns(5).DataField=   "strNomeFantasia"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1984"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2037"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1958"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=9499"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=9419"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=4128"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=4048"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=1402"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1323"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
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
      MultiSelect     =   0
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
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
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin VB.Menu mnu_TipoComunicacao 
      Caption         =   "mnuTipoComunicacao"
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
Attribute VB_Name = "frmCadContribuinte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public intCodSeguranca       As Integer
    
    Dim mblnAlterando            As Boolean
    Dim mobjAux                  As Object
    Dim oList                    As Object
    Dim mblnClickOk              As Boolean
    
    Dim strNomeAtual             As String
    Dim strCnpjCpfAtual          As String
    
    Dim X                        As XArrayDB     'Grid Contribuintes
    Dim mblnselecionou           As Boolean
    Dim mblnPrimeiraVez          As Boolean
    Dim e                        As New XArrayDB 'Grid Sócios
    
    Dim mblnActivate             As Boolean
    
    Dim bytOrdenacao             As Byte
    Dim blnOrdenacaoAsc          As Boolean
    Dim intUsuario               As Long
    Dim strCep                   As String
    
    Dim blnEstadoInativo         As Boolean
    
    Dim intCodigoAtual           As Long
        
    Dim intMunicipioEmpresa      As Long
    Dim intCepEmpresa            As Long
    
Private Sub converte_complemento()
'Rotina que verifica se o campo intnumeroc tem caracter se tiver
'pega o valor dele joga para o campo strcomplementoc e deixa o intnumeroc nulo
On Error GoTo Trataerro
   Dim strSql          As String
   Dim adoResultado    As ADODB.Recordset
   Dim strNumero As String
   Dim ncount As Integer
  
    Screen.MousePointer = vbHourglass
    
    strSql = ""
    strSql = "Select pkid,intnumeroc,strcomplementoc from " & gstrContribuinte
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            
            Do While Not .EOF
                strNumero = ""
                
                If IsNull(!intNumeroC) = False Then
                    If Not IsNumeric(Trim((!intNumeroC))) Then
                            
                            strNumero = IIf(IsNull(!strComplementoC) = True, "", Trim(!strComplementoC)) & Space(1) & Trim(!intNumeroC)
                            gobjBanco.ExecutaBeginTrans
                            strSql = "UPDATE " & gstrContribuinte & " set intnumeroc = null  ,strcomplementoC = '" & Trim(strNumero) & "'  where pkid = " & !Pkid
                            
                            If Not gobjBanco.Execute(strSql) Then
                                gobjBanco.ExecutaRollbackTrans
                                 Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                    End If
                End If
                adoResultado.MoveNext
            Loop
       End With
    End If
    gobjBanco.ExecutaCommitTrans
    Screen.MousePointer = vbDefault

Trataerro:
        gobjBanco.ExecutaRollbackTrans
        Screen.MousePointer = vbDefault
End Sub


Private Sub cbobytNaturezaJuridica_Click()
    
    Select Case cbobytNaturezaJuridica.ListIndex
        
        Case 0  'Física
            chkblnResidenteNoMunicipio.Caption = "Residente no município"
            HabilitaDesabilitaObjeto chkblnResidenteNoMunicipio, True
            HabilitaDesabilitaObjeto txtstrIdentidade, True
            HabilitaDesabilitaObjeto txtstrTituloEleitoral, True
            HabilitaDesabilitaObjeto txtdtmDataNascimento, True
            HabilitaDesabilitaObjeto txtstrCarteiraTrabalho, True
            HabilitaDesabilitaObjeto txtstrInscricaoEstadual, False
            HabilitaDesabilitaObjeto dbcstrNomeFantasia, False
            HabilitaDesabilitaObjeto mskstrPIS, True
            mskstrPIS.Mask = "###\.#####\.##\.# "
            'If mblnAlterando = False Then
                HabilitaDesabilitaObjeto mskstrCNPJCPF, True
            'End If
            mskstrCNPJCPF.Mask = "###\.###\.###\-##"
            tab_3DDadosGerais.TabEnabled(1) = False
            tab_3DDadosGerais.TabEnabled(5) = False
            dbcstrNomeFantasia.Text = ""
            tab_3DDadosGerais.TabEnabled(1) = True
            tab_3DDadosGerais.TabEnabled(2) = True
            tab_3DDadosGerais.TabEnabled(3) = True
            tab_3DDadosGerais.TabEnabled(4) = True
            tab_3DDadosGerais.TabEnabled(6) = True
            txtintCepC.Text = ""
            
        Case 1 'Jurídica
            chkblnResidenteNoMunicipio.Caption = "Estabelecido no município"
            HabilitaDesabilitaObjeto chkblnResidenteNoMunicipio, True
            HabilitaDesabilitaObjeto txtstrIdentidade, False
            HabilitaDesabilitaObjeto txtstrTituloEleitoral, False
            HabilitaDesabilitaObjeto txtdtmDataNascimento, False
            HabilitaDesabilitaObjeto txtstrCarteiraTrabalho, False
            HabilitaDesabilitaObjeto txtstrInscricaoEstadual, True
            HabilitaDesabilitaObjeto dbcstrNomeFantasia, True
            HabilitaDesabilitaObjeto mskstrPIS, False
            'If mblnAlterando = False Then
                HabilitaDesabilitaObjeto mskstrCNPJCPF, True
            'End If
            mskstrCNPJCPF.Mask = "##\.###\.###\/####\-##"
            DoEvents
            tab_3DDadosGerais.TabEnabled(1) = True
            tab_3DDadosGerais.TabEnabled(5) = True
            dbcstrNomeFantasia.Text = dbcstrNome.Text
            tab_3DDadosGerais.TabEnabled(1) = True
            tab_3DDadosGerais.TabEnabled(2) = True
            tab_3DDadosGerais.TabEnabled(3) = True
            tab_3DDadosGerais.TabEnabled(4) = True
            tab_3DDadosGerais.TabEnabled(6) = True
            txtintCepC.Text = ""
            
        Case Is > 1 'Todos outros
            chkblnResidenteNoMunicipio.Caption = "Estabelecido no município"
            chkblnResidenteNoMunicipio.Value = vbUnchecked
            HabilitaDesabilitaObjeto chkblnResidenteNoMunicipio, False
            HabilitaDesabilitaObjeto txtstrIdentidade, False
            HabilitaDesabilitaObjeto txtstrTituloEleitoral, False
            HabilitaDesabilitaObjeto txtdtmDataNascimento, False
            HabilitaDesabilitaObjeto txtstrCarteiraTrabalho, False
            HabilitaDesabilitaObjeto txtstrInscricaoEstadual, False
            HabilitaDesabilitaObjeto dbcstrNomeFantasia, False
            HabilitaDesabilitaObjeto mskstrPIS, False
            HabilitaDesabilitaObjeto mskstrCNPJCPF, False
            DoEvents
            tab_3DDadosGerais.TabEnabled(1) = False
            tab_3DDadosGerais.TabEnabled(5) = False
            tab_3DDadosGerais.TabEnabled(1) = False
            tab_3DDadosGerais.TabEnabled(2) = False
            tab_3DDadosGerais.TabEnabled(3) = False
            tab_3DDadosGerais.TabEnabled(4) = False
            tab_3DDadosGerais.TabEnabled(6) = False
            
            PreencherListaDeOpcoes dbcintMunicipioC, intMunicipioEmpresa
            txtintCepC.Text = intCepEmpresa
            
    End Select
    
    HabilitaDesabilitaObjeto dbcstrNome, True
    HabilitaDesabilitaObjeto dtpdtmDataCadastro, True
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar

End Sub

Private Sub dbcstrLogradouroC_Click(Area As Integer)
Dim strSql As String
Dim adoResultado As ADODB.Recordset
    
    On Error GoTo Trataerro
    
    If (Area = 2 Or Area = 1) And dbcstrLogradouroC.BoundText <> "" And IsNumeric(dbcstrLogradouroC.BoundText) Then
       strSql = strSql & "SELECT "
       strSql = strSql & "TL.strDescricao strTituloLogradouro, "
       strSql = strSql & "TP.strSigla strTipoLogradouro, "
       strSql = strSql & "BA.strDescricao strBairro, "
       strSql = strSql & "MU.strDescricao strMunicipio, "
       strSql = strSql & "UF.strSigla strUF, "
       strSql = strSql & "LO.intCEP intCEP "
       
       'Feito por causa de problemas de join no sql
       If bytDBType = EDatabases.Oracle Then
           strSql = strSql & "FROM "
           strSql = strSql & gstrLogradouro & " LO, "
           strSql = strSql & gstrTituloLogradouro & " TL, "
           strSql = strSql & gstrTipoLogradouro & " TP, "
           strSql = strSql & gstrBairro & " BA, "
           strSql = strSql & gstrCidade & " MU, "
           strSql = strSql & gstrUF & " UF "
            
           strSql = strSql & "WHERE "
           strSql = strSql & "TL.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LO.intTituloLogradouro AND "
           strSql = strSql & "TP.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LO.intTipoLogradouro AND "
           strSql = strSql & "BA.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LO.intBairro AND "
           strSql = strSql & "MU.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " BA.intMunicipio AND "
           strSql = strSql & "UF.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " MU.intUF AND "
           strSql = strSql & "LO.pkID = " & dbcstrLogradouroC.BoundText
       Else
       
           strSql = strSql & " FROM " & gstrCidade & " MU LEFT OUTER JOIN "
           strSql = strSql & gstrUF & " UF ON MU.intUF = UF.PKId RIGHT OUTER JOIN "
           strSql = strSql & gstrBairro & " BA ON MU.PKId = BA.intMunicipio RIGHT OUTER JOIN "
           strSql = strSql & gstrLogradouro & " LO LEFT OUTER JOIN "
           strSql = strSql & gstrTipoLogradouro & " TP ON LO.intTipoLogradouro = TP.PKId LEFT OUTER JOIN "
           strSql = strSql & gstrTituloLogradouro & " TL ON LO.intTituloLogradouro = TL.PKId ON BA.PKId = LO.intBairro "
           
           strSql = strSql & "WHERE "
           strSql = strSql & "LO.pkID = " & dbcstrLogradouroC.BoundText
       End If
       
      Set adoResultado = New ADODB.Recordset
       Set gobjBanco = New clsBanco
       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
          If Not adoResultado.EOF Then
             If Not IsNull(adoResultado("strTituloLogradouro")) Then
                dbcintTituloLogradouro.Text = adoResultado("strTituloLogradouro")
                PreencherListaDeOpcoes dbcintTituloLogradouro
                dbcintTituloLogradouro.Text = adoResultado("strTituloLogradouro")
             Else
                dbcintTituloLogradouro.Text = ""
                dbcintTituloLogradouro.BoundText = ""
             End If
             
             If Not IsNull(adoResultado("strTipoLogradouro")) Then
                dbcintTipoLogradouro.Text = adoResultado("strTipoLogradouro")
                PreencherListaDeOpcoes dbcintTipoLogradouro
                dbcintTipoLogradouro.Text = adoResultado("strTipoLogradouro")
             Else
                dbcintTipoLogradouro.Text = ""
                dbcintTipoLogradouro.BoundText = ""
             End If
             
             If Not IsNull(adoResultado("strMunicipio")) Then
                dbcintMunicipioC.Text = adoResultado("strMunicipio")
                PreencherListaDeOpcoes dbcintMunicipioC
                dbcintMunicipioC.Text = adoResultado("strMunicipio")
             Else
                dbcintMunicipioC.Text = ""
                dbcintMunicipioC.BoundText = ""
             End If
             
             If Not IsNull(adoResultado("strUF")) Then
                dbcintUFC.Text = adoResultado("strUF")
                PreencherListaDeOpcoes dbcintUFC
                dbcintUFC.Text = adoResultado("strUF")
             Else
                dbcintUFC.Text = ""
                dbcintUFC.BoundText = ""
             End If
             
             txtstrBairroC.Text = gstrENulo(adoResultado("strBairro"))
             
             txtintCepC.Text = gstrENulo(adoResultado("intCEP"))
             txtintCepC.Text = gstrCEPFormatado(txtintCepC.Text)
          End If
       End If
    End If
    
    txtintCodigoLogradouro.Text = ""
    txtstrDistritoC.Text = ""
    txtintNumeroC.Text = ""
    txtstrComplementoC.Text = ""
    
Exit Sub
Trataerro:
    
    If Err.Number = 7 Then 'Out of Memory (dbcstrLogradouroC.BoundText)
       Exit Sub
    Else
       ExibeMensagem Err.Number & " - " & Err.Description
    End If
    
End Sub

Private Sub dbcstrNome_Click(Area As Integer)
    If Area = 0 Then
        DropDownDataCombo dbcstrNome, Me, Area
    End If
End Sub

Private Sub dbcstrNome_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcstrNome, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouro_Change()
    If dbcintLogradouro.MatchedWithList Then
        LogradouroCep dbcintLogradouro.BoundText, txt_strBairro, , txt_strMunicipio, txt_strUF, txtintCep
    End If
End Sub

Private Sub dbcintLogradouro_Click(Area As Integer)
   If Area = 0 Then DropDownDataCombo dbcintLogradouro, Me, Area
    If dbcintLogradouro.MatchedWithList Then
        LogradouroCep dbcintLogradouro.BoundText, txt_strBairro, , txt_strMunicipio, txt_strUF, txtintCep
    End If

End Sub

Private Sub dbcintLogradouro_GotFocus()
    AjustaToolBar
    tab_3DDadosGerais.Tab = 1
End Sub

Private Sub dbcintLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
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

Private Sub dbcintTipoLogradouro_Click(Area As Integer)
   If Area = 0 Then DropDownDataCombo dbcintTipoLogradouro, Me, Area
End Sub

Private Sub dbcintTipoLogradouro_GotFocus()
    tab_3DDadosGerais.Tab = 1
End Sub

Private Sub dbcintTipoLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTipoLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoLogradouroD_Click(Area As Integer)
   If Area = 0 Then DropDownDataCombo dbcintTipoLogradouroD, Me, Area
End Sub

Private Sub dbcintTipoLogradouroD_GotFocus()
    tab_3DDadosGerais.Tab = 1
End Sub

Private Sub dbcintTipoLogradouroD_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTipoLogradouroD, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouro_Click(Area As Integer)
   If Area = 0 Then DropDownDataCombo dbcintTituloLogradouro, Me, Area
End Sub

Private Sub dbcintTituloLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTituloLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouroD_Click(Area As Integer)
   If Area = 0 Then DropDownDataCombo dbcintTituloLogradouroD, Me, Area
End Sub

Private Sub dbcintTituloLogradouroD_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTipoLogradouroD, Me, , KeyCode, Shift
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

Private Sub dbcintUFD_Click(Area As Integer)
   If Area = 0 Then DropDownDataCombo dbcintUFD, Me, Area
End Sub

Private Sub dbcintUFD_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUFD, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUFD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chk_ContaPublica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chk_DebitoAutomatico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkblnResidenteNoMunicipio_Click()
    If chkblnResidenteNoMunicipio.Value = 0 Then
        tab_3DCorrespondencia.TabEnabled(0) = False
        tab_3DCorrespondencia.Tab = 1
        
        dbcintLogradouro.BoundText = ""
        txt_strBairro.Text = ""
        txtintNumero = ""
        txtstrComplemento = ""
        txtintCep = ""
    Else
        tab_3DCorrespondencia.TabEnabled(0) = True
        tab_3DCorrespondencia.Tab = 0
    End If
End Sub

Private Sub chkblnResidenteNoMunicipio_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", chkblnResidenteNoMunicipio
End Sub

Private Sub cmd_ContasBancarias_Click()
    If Trim(txtPKId) = "" Then
        ExibeMensagem "O contribuinte tem que ser salvo."
        Exit Sub
    End If
    frmCadContasBancarias.Show
    'PreencherListaDeOpcoes frmCadContasBancarias.dbcintContribuinte, Val(txtPKId)
    'TrocaCorObjeto frmCadContasBancarias.dbcintContribuinte, True
    'frmCadContasBancarias.cmd_Contribuinte.Enabled = False
    frmCadContasBancarias.Visible = True
    'frmCadContasBancarias.RotinaAuxiliarClickCombo
End Sub

Private Sub dbcintTipoLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTipoLogradouro
End Sub

Private Sub dbcintTipoLogradouroD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTipoLogradouroD
End Sub

Private Sub dbcintTituloLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTituloLogradouro
End Sub

Private Sub dbcintTituloLogradouroD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTituloLogradouroD
End Sub

Private Sub cmd_TipoLogradouro_Click()
    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouro
End Sub

Private Sub cmd_TituloLogradouro_Click()
    CarregaForm frmCadTituloLogradouro, dbcintTituloLogradouro
End Sub

Private Sub cmd_TipoLogradouroD_Click()
    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouroD
End Sub

Private Sub cmd_TituloLogradouroD_Click()
    CarregaForm frmCadTituloLogradouro, dbcintTituloLogradouroD
End Sub

Private Sub cmd_Logradouro_Click()
    ChamaFormCadastro frmCadLogradouro, dbcintLogradouro
End Sub

Private Sub cmd_Down_Click()
    If lvw_TipoComunicacao.ListItems.Count <> 0 Then
        MoveItemNoListView lvw_TipoComunicacao, True
    End If
End Sub

Private Sub cmd_Up_Click()
    If lvw_TipoComunicacao.ListItems.Count <> 0 Then
        MoveItemNoListView lvw_TipoComunicacao, False
    End If
End Sub

Private Sub cmd_MunicipioC_Click()
    ChamaFormCadastro frmCadCidade, dbcintMunicipioC
End Sub


Private Sub dbcstrNomeFantasia_Click(Area As Integer)
    If Area = 0 Then
        DropDownDataCombo dbcstrNomeFantasia, Me, Area
    End If
End Sub

Private Sub dbcstrNomeFantasia_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcstrNomeFantasia, Me, , KeyCode, Shift
End Sub

Private Sub dtpdtmDataCadastro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dtpdtmDataCadastro
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = intCodSeguranca
    VirificaGradeListView Me
    If mblnselecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
    If Not mblnActivate Then VerificaObjParaAplicar mobjAux
    
    mblnActivate = True
    
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
Dim lngTempo As Long
Dim adoResultado As ADODB.Recordset

    intCodSeguranca = gintCodSeguranca
    Me.HelpContextID = intCodSeguranca

    bytOrdenacao = 2: blnOrdenacaoAsc = True
    tab_3DDadosGerais.TabVisible(3) = False
    lngTempo = Timer

    TrocaCorObjeto txtintCodigo, True

    lblintCodigo.Visible = False
    txtintCodigo.Visible = False

    Me.Icon = MDIMenu.Icon
    If MDIMenu.Tag = "Ouvidoria" Then
        'cmd_Bairro.Enabled = False
        'cmd_ContasBancarias.Enabled = False
        'cmd_Logradouro.Enabled = False
        'cmd_Municipio.Enabled = False
        'cmd_MunicipioC.Enabled = False
        'cmd_TipoLogradouro.Enabled = False
        'cmd_TipoLogradouroD.Enabled = False
        'cmd_TituloLogradouro.Enabled = False
        'cmd_TituloLogradouroD.Enabled = False
    ElseIf MDIMenu.Tag = "FROTA" Then
        Me.Caption = "Único de pessoas "
    ElseIf MDIMenu.Tag = "Protocolo" Then
        lblintCodigo.Visible = True
        txtintCodigo.Visible = True
        TrocaCorObjeto txtintCodigo, False
    End If
    
    fra_Tipo.Visible = UCase(App.ProductName) = "ORCAMENTARIO"
        
    TrocaCorObjeto txtPKId, True
    TrocaCorObjeto txt_strBairro, True
    TrocaCorObjeto txt_strMunicipio, True
    TrocaCorObjeto txt_strUF, True
    
    MontaColumnHeaders
   
    CarregaMunicipio
   
    'dbcintLogradouro.Tag = gstrQueryLogradouro(gstrBairro, "bytPertenceAoMunicipo = 1 AND L.intBairro = " & gstrBairro & ".PKId ") & ";L.strDescricao"
    'dbcintLogradouro.Tag = strQueryLogradouro(True, False) & ";L.strDescricao"
    'dbcstrLogradouroC.Tag = strQueryLogradouro(True, True) & ";L.strDescricao"

    dbcintMunicipioC.Tag = gstrQueryDataComboMunicipio & ";strDescricao"
    dbcintUFC.Tag = gstrQueryDataComboUF & ";strSigla"

    dbcintUFD.Tag = gstrQueryDataComboUF & ";strSigla"

    dbcintTipoLogradouro.Tag = gstrQueryDataComboTipoLogradouro & ";strSigla"
    dbcintTituloLogradouro.Tag = gstrQueryDataComboTituloLogradouro & ";strDescricao"
    dbcintLogradouro.Tag = strQueryLogradouro(False) & ";L.strDescricao"

    dbcintTipoLogradouroD.Tag = gstrQueryDataComboTipoLogradouro & ";strSigla"
    dbcintTituloLogradouroD.Tag = gstrQueryDataComboTituloLogradouro & ";strDescricao"

    dbcstrNome.Tag = "SELECT PKId, strNome FROM " & gstrContribuinte & " ORDER BY strNome " & ";strNome"
    dbcstrNomeFantasia.Tag = "SELECT PKId, strNomeFantasia FROM " & gstrContribuinte & " WHERE NOT strNomeFantasia IS NULL ORDER BY strNomeFantasia " & ";strNomeFantasia"

    'LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strQueryGridContribuinte
    'MontaArray
        
    PreencheMenuPopup
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    txtintCodigoLogradouro.Enabled = False
    TrocaCorObjeto txtintCodigoLogradouro, True
    txtintCodigoLogradouroD.Enabled = False
    TrocaCorObjeto txtintCodigoLogradouroD, True
    tab_3DDadosGerais.TabEnabled(5) = False
    tab_3DDadosGerais.TabEnabled(6) = False
    
   
    tab_3DCorrespondencia.TabEnabled(0) = True
    tab_3DCorrespondencia.Tab = 0
    
    PreencheListaModulos
    'converte_complemento
    SelecionaModuloAtual 0
    
    LeDadosEmpresa
    PreencheNatureza
    NovoContribuinte
    
End Sub

Private Function strQueryLogradouro(blnExcluidos As Boolean) As String
    Dim strSql  As String
     
    strSql = ""
    
    'If Not blnSemBairro Then 'Logradouro + Bairro
       strSql = strSql & "SELECT L.Pkid, "
       strSql = strSql & " RTRIM(LTRIM(L.strDescricao)) " & strCONCAT & gstrISNULL("TL.strSigla", "''", "', '") & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & _
                strCONCAT & gstrISNULL("U.strDescricao", "' '", "', '") & strCONCAT & gstrISNULL("U.strDescricao", "''") & ")) " & strCONCAT & "' ( '" & strCONCAT & gstrISNULL("BA.strDescricao", "''") & strCONCAT & "' ) '" & " AS Logradouro "
    'Else 'Logradouro
    '   strSql = strSql & "SELECT L.strDescricao, "
    '   strSql = strSql & " RTRIM(LTRIM(L.strDescricao)) " & strCONCAT & gstrISNULL("TL.strSigla", "''", "', '") & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & _
    '            strCONCAT & gstrISNULL("U.strDescricao", "' '", "', '") & strCONCAT & gstrISNULL("U.strDescricao", "''") & ")) AS Logradouro "
    'End If
    
    strSql = strSql & "FROM "
    strSql = strSql & gstrBairro & " BA, "
    strSql = strSql & gstrLogradouro & " L, "
    strSql = strSql & gstrTituloLogradouro & " U, "
    strSql = strSql & gstrTipoLogradouro & " TL "
    
    strSql = strSql & " WHERE L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle
    If Not blnExcluidos Then
        strSql = strSql & " AND L.Dtmdtexclusao IS NULL "
    End If
    strSql = strSql & " AND BA.bytPertenceAoMunicipo = 1 "
    strSql = strSql & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
    strSql = strSql & " AND L.intBairro " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle
    
    strSql = strSql & " ORDER BY L.strDescricao "
    
    strQueryLogradouro = strSql
        
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrSalvar
    mblnselecionou = False
    mblnPrimeiraVez = False
    mblnActivate = False
End Sub

Private Sub lvw_Contas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_Contas
        txt_CodBanco = .SelectedItem.Text
        txt_Banco = .SelectedItem.SubItems(1)
        txt_CodAgencia = .SelectedItem.SubItems(2)
        txt_Agencia = .SelectedItem.SubItems(3)
        txt_Conta = .SelectedItem.SubItems(4)
        txt_DigitoVerificador = .SelectedItem.SubItems(5)
        chk_ContaPublica.Value = gbytZeroOuUm(.SelectedItem.SubItems(6))
        chk_DebitoAutomatico.Value = gbytZeroOuUm(.SelectedItem.SubItems(7))
        txt_dtmDebito = .SelectedItem.SubItems(8)
    End With
End Sub

Private Sub lvw_Contas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub lvw_TipoComunicacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_TipoComunicacao
        lbl_TipoComunicacao = .SelectedItem.Text
        txt_Conteudo = .SelectedItem.SubItems(1)
        txt_DescricaoConteudo = .SelectedItem.SubItems(2)
    End With
End Sub

Private Sub lvw_TipoComunicacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub lvw_TipoComunicacao_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnu_TipoComunicacao
    End If
End Sub

Private Sub mnu_Deletar_Click()
    With lvw_TipoComunicacao
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem.Selected = False Then Exit Sub
        .ListItems.Remove .SelectedItem.Index
        txt_Conteudo = ""
        txt_DescricaoConteudo = ""
        lbl_TipoComunicacao = "Tipo"
    End With
End Sub

Private Sub mnu_Lista_Click(Index As Integer)

    With lvw_TipoComunicacao
        .Sorted = False
        Set oList = .ListItems.Add(, , mnu_Lista(Index).Caption)
        oList.SubItems(1) = ""
        oList.SubItems(2) = ""
        oList.Tag = mnu_Lista(Index).Tag
        .ListItems(.ListItems.Count).Selected = True
        .ListItems(.ListItems.Count).EnsureVisible
        lvw_TipoComunicacao_ItemClick .SelectedItem
        txt_Conteudo.SetFocus
    End With
End Sub

Private Sub mskstrCNPJCPF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrCNPJCPF
End Sub

Private Sub mskstrCNPJCPF_LostFocus()
    mskstrCNPJCPF = gstrCGCCPFFormatado(mskstrCNPJCPF.Text)
End Sub

Private Sub mskstrPIS_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrPIS
End Sub

Private Sub tab_3DCorrespondencia_Click(PreviousTab As Integer)

    If tab_3DDadosGerais.Tab <> 1 Then Exit Sub
        
    Select Case tab_3DCorrespondencia.Tab
        Case 0
            If dbcintLogradouro.Enabled Then dbcintLogradouro.SetFocus
        Case 1
            If mblnAlterando = False And tab_3DCorrespondencia.TabEnabled(0) = True Then
                
                If Len(dbcintLogradouro.BoundText) > 0 And Not dbcintLogradouro.MatchedWithList Then
                    ExibeMensagem "O Endereço residencial não é válido."
                    tab_3DCorrespondencia.Tab = 0
                    dbcintLogradouro.SetFocus
                    Exit Sub
                End If
                                                    
                If Len(dbcintLogradouro.BoundText) = 0 Or Len(dbcstrLogradouroC.BoundText) > 0 Then
                    If txtintCepC.Enabled Then txtintCepC.SetFocus
                    Exit Sub
                End If
                If MsgBox("O endereço de correspondência é o mesmo que o endereço residencial?", vbYesNo + vbQuestion) = vbYes Then
                    
                    CopiaLogradouro Val(dbcintLogradouro.BoundText), 1
                    
                    dbcintTipoLogradouro.Enabled = False
                    TrocaCorObjeto dbcintTipoLogradouro, True
                    dbcintTituloLogradouro.Enabled = False
                    TrocaCorObjeto dbcintTituloLogradouro, True
                    cmd_TipoLogradouro.Enabled = False
                    TrocaCorObjeto cmd_TipoLogradouro, True
                    cmd_TituloLogradouro.Enabled = False
                    TrocaCorObjeto cmd_TituloLogradouro, True
                    
                    'txtintCodigoLogradouro = dbcintLogradouro.BoundText
                    'dbcstrLogradouroC = dbcintLogradouro.Text
                    txtstrBairroC = txt_strBairro.Text
                    PreencherListaDeOpcoes dbcintMunicipioC, IIf(txt_strMunicipio.Tag = "", gintMunicipioEmpresa, txt_strMunicipio.Tag)
                    PreencherListaDeOpcoes dbcintUFC, IIf(txt_strUF.Tag = "", gintMunicipioEmpresa, txt_strUF.Tag)
                    txtintNumeroC = txtintNumero
                    txtstrComplementoC = txtstrComplemento
                    txtintCepC = txtintCep
                    If txtintCepC.Enabled Then txtintCepC.SetFocus
                End If
            Else
                If txtintCepC.Enabled Then txtintCepC.SetFocus
            End If
            
        Case 2
            If mblnAlterando = False And tab_3DCorrespondencia.TabEnabled(0) = True And txtintCodigoLogradouroD.Text <> dbcintLogradouro.BoundText Then
                
                If Len(dbcintLogradouro.BoundText) > 0 And Not dbcintLogradouro.MatchedWithList Then
                    ExibeMensagem "O Endereço residencial não é válido."
                    tab_3DCorrespondencia.Tab = 0
                    dbcintLogradouro.SetFocus
                    Exit Sub
                End If
                
                If Len(dbcintLogradouro.BoundText) = 0 Or Len(txtstrLogradouroD.Text) > 0 Then
                    If dbcintTipoLogradouroD.Enabled Then
                        dbcintTipoLogradouroD.SetFocus
                    Else
                        If txtstrLogradouroD.Enabled Then txtstrLogradouroD.SetFocus
                    End If
                    Exit Sub
                End If
                If MsgBox("O endereço do domicílio fiscal é o mesmo que o endereço residencial?", vbYesNo + vbQuestion) = vbYes Then
                    
                    CopiaLogradouro dbcintLogradouro.BoundText, 2
                    
                    dbcintTipoLogradouroD.Enabled = False
                    TrocaCorObjeto dbcintTipoLogradouroD, True
                    dbcintTituloLogradouroD.Enabled = False
                    TrocaCorObjeto dbcintTituloLogradouroD, True
                    cmd_TipoLogradouroD.Enabled = False
                    TrocaCorObjeto cmd_TipoLogradouroD, True
                    cmd_TituloLogradouroD.Enabled = False
                    TrocaCorObjeto cmd_TituloLogradouroD, True
                    TrocaCorObjeto txtstrDistritoD, True
                    
                    'txtintCodigoLogradouroD = dbcintLogradouro.BoundText
                    'txtstrLogradouroD = dbcintLogradouro.Text
                    txtstrBairroD = txt_strBairro.Text
                    txtstrMunicipioD = txt_strMunicipio.Text
                    PreencherListaDeOpcoes dbcintUFD, txt_strUF.Tag
                    txtintNumeroD = txtintNumero
                    txtstrComplementoD = txtstrComplemento
                    txtintCepD = txtintCep
                    txtstrDistritoD.Text = txtstrDistritoC.Text
                    If dbcintTipoLogradouroD.Enabled Then
                        dbcintTipoLogradouroD.SetFocus
                    Else
                        If txtstrLogradouroD.Enabled Then txtstrLogradouroD.SetFocus
                    End If

                End If
            Else
                If dbcintTipoLogradouroD.Enabled Then
                    dbcintTipoLogradouroD.SetFocus
                Else
                    If txtstrLogradouroD.Enabled Then txtstrLogradouroD.SetFocus
                End If
            End If
            
            If MDIMenu.Tag = "MATERIAL" Then
                If mblnAlterando = False And tab_3DCorrespondencia.TabEnabled(0) = False And txtstrLogradouroD.Text <> dbcstrLogradouroC.Text Then
                    
                    If MsgBox("O endereço do domicílio fiscal é o mesmo que o endereço de correspondência?", vbYesNo + vbQuestion) = vbYes Then
                        
                        dbcintTipoLogradouroD.Enabled = False
                        TrocaCorObjeto dbcintTipoLogradouroD, True
                        dbcintTituloLogradouroD.Enabled = False
                        TrocaCorObjeto dbcintTituloLogradouroD, True
                        cmd_TipoLogradouroD.Enabled = False
                        TrocaCorObjeto cmd_TipoLogradouroD, True
                        cmd_TituloLogradouroD.Enabled = False
                        TrocaCorObjeto cmd_TituloLogradouroD, True
                        TrocaCorObjeto txtstrDistritoD, True
                        
                        PreencherListaDeOpcoes dbcintTipoLogradouroD, dbcintTipoLogradouro.BoundText
                        PreencherListaDeOpcoes dbcintTituloLogradouroD, dbcintTituloLogradouro.BoundText
                        txtstrLogradouroD.Text = dbcstrLogradouroC.Text
                        txtstrBairroD.Text = txtstrBairroC.Text
                        txtstrMunicipioD.Text = dbcintMunicipioC.Text
                        PreencherListaDeOpcoes dbcintUFD, dbcintUFC.BoundText
                        txtintNumeroD.Text = txtintNumeroC.Text
                        txtstrComplementoD.Text = txtstrComplementoC.Text
                        txtintCepD.Text = txtintCepC.Text
                        txtstrDistritoD.Text = txtstrDistritoC.Text
                        If dbcintTipoLogradouroD.Enabled Then
                            dbcintTipoLogradouroD.SetFocus
                        Else
                            If txtstrLogradouroD.Enabled Then txtstrLogradouroD.SetFocus
                        End If
                        
                    End If
                End If
            End If
    End Select
    
End Sub


Private Sub tab_3DDadosGerais_Click(PreviousTab As Integer)
'    Select Case tab_3DDadosGerais.Tab
'        Case 1
'            If chkblnResidenteNoMunicipio.Value = 1 Then
'
'                '### GUSTAVO ###
'                'PreencherListaDeOpcoes dbcintMunicipio, gintMunicipioEmpresa
'                PreencherListaDeOpcoes dbcintMunicipioC, gintMunicipioEmpresa
'
'                'dbcintMunicipio.BoundText = gintMunicipioEmpresa
'                dbcintMunicipioC.BoundText = gintMunicipioEmpresa
'
'                TrocaCorObjeto dbcintMunicipioC, True
'                'cmd_Municipio.Enabled = False
'
'                'PreencherListaDeOpcoes dbcintUf, gintUFEmpresa
'                PreencherListaDeOpcoes dbcintUFC, IIf(txt_strUF.Tag = "", gintUFEmpresa, txt_strUF.Tag)
'
'                'dbcintUf.BoundText = gintUFEmpresa
'                dbcintUFC.BoundText = gintUFEmpresa
'                'TrocaCorObjeto dbcintUf, True
'                TrocaCorObjeto dbcintUFC, True
'                'cmd_MunicipioC.Enabled = False
'            Else
'                'TrocaCorObjeto dbcintMunicipio, False
'                TrocaCorObjeto dbcintMunicipioC, False
'                'TrocaCorObjeto dbcintUf, False
'                TrocaCorObjeto dbcintUFC, False
'            End If
'    End Select
End Sub

Private Sub tab_3DDadosGerais_GotFocus()
'    txt_Codigo1 = txtstrCodigoAnterior
'    txt_Nome1 = dbcstrNome
'    txt_Codigo2 = txtstrCodigoAnterior
'    txt_Nome2 = dbcstrNome
'    txt_Codigo3 = txtstrCodigoAnterior
'    txt_Nome3 = dbcstrNome
'    txt_Codigo4 = txtstrCodigoAnterior
'    txt_Nome4 = dbcstrNome
    txt_Codigo1 = txtPKId
    txt_Nome1 = dbcstrNome.Text
    txt_Codigo2 = txtPKId
    txt_Nome2 = dbcstrNome.Text
    txt_Codigo3 = txtPKId
    txt_Nome3 = dbcstrNome.Text
    txt_Codigo4 = txtPKId
    txt_Nome4 = dbcstrNome.Text
End Sub

Private Sub tdb_Historico_KeyPress(KeyAscii As Integer)

    Select Case tdb_Historico.Col
        Case 1
            CaracterValido KeyAscii, "N", tdb_Historico
        Case 2
            CaracterValido KeyAscii, "D", tdb_Historico
        Case 3, 4
            CaracterValido KeyAscii
        Case 5
            CaracterValido KeyAscii, "V", tdb_Historico
    End Select

End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
    'PreencheGridContribuinte
End Sub

Sub PreencheGridContribuinte()
    Dim Col    As TrueOleDBGrid70.Column
    Dim c      As Integer
    Dim tmp    As String
    Dim n      As Integer
    Dim strAux As String
    Dim strSql As String
    Dim adoTemp As ADODB.Recordset
    
    c = tdb_Lista.Col
    tdb_Lista.HoldFields
    
    tmp = ""
    
    strSql = ""
    strSql = strSql & " SELECT CO.PKID, CO.strNome, CO.strCNPJCPF, CO.bytNaturezaJuridica, CO.strNomeFantasia," & gstrCASEWHEN("CO.blnInativo", "0,'Não',1,'Sim'") & " strInativo "

    strSql = strSql & "FROM " & gstrContribuinte & " CO, " & gstrItens & " IT, " & gstrModuloContribuinte & " MC "
    strSql = strSql & "WHERE IT.PKId = MC.intItem AND "
    strSql = strSql & "MC.intContribuinte = CO.PKId AND IT.PKId =" & gintModulo

    strAux = strSql
    strAux = strAux & " ORDER BY strNome"
    
    strSql = strSql & " AND "
    
    
    For Each Col In tdb_Lista.Columns
        If Trim(Col.FilterText) <> "" Then
            n = n + 1
        
            If tmp <> "" Then
                tmp = tmp & " AND "
            End If
            
            Select Case UCase(Col.DataField)
                Case "PKID"
                    tmp = tmp & "CO." & Col.DataField & " = " & Col.FilterText
                Case "STRNOME"
                    tmp = tmp & "UPPER(CO." & Col.DataField & ") LIKE '" & UCase(Col.FilterText) & "%'"
                Case "STRCNPJCPF"
                    tmp = tmp & "CO." & Col.DataField & " LIKE '" & gstrValorSemMascara(Col.FilterText) & "%'"
                Case "STRINATIVO"
                    If UCase(Left(Col.FilterText, 1)) = "S" Then
                        tmp = tmp & "CO.blnInativo = 1 "
                    ElseIf UCase(Left(Col.FilterText, 1)) = "N" Then
                        tmp = tmp & "CO.blnInativo = 0 "
                    End If
                    
                    
            End Select
        End If
    Next
    
    If tmp <> "" Then
        strSql = strSql & tmp
        strSql = strSql & " ORDER BY strNome"
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoTemp) Then
            If Not adoTemp.EOF Then
                'MontaArray strSql
                LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strSql
            Else
                'MontaArray
                
                LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strQueryGridContribuinte
            End If
        End If
    Else
        'MontaArray
        LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strQueryGridContribuinte
    End If
    
    tdb_Lista.Col = c
    tdb_Lista.EditActive = True
    tdb_Lista.CurrentCellModified = True
    
    NovoContribuinte
   
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    
    'blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
    
    'bytOrdenacao = ColIndex
    
    'ToolBarGeral gstrRefresh, gstrContribuinte, mblnAlterando, tdb_Lista, Me, mobjAux, strQueryGridContribuinte
    
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnPrimeiraVez = True
    Select Case KeyCode
    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown
        mblnClickOk = True
    Case Else
        mblnClickOk = False
    End Select
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    
    Select Case tdb_Lista.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_Lista
        Case Else
            CaracterValido KeyAscii, "A", tdb_Lista
    End Select
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Integer
Dim strSql As String
Dim adoResultado As ADODB.Recordset

        
    On Error GoTo err_tdb_Lista_RowColChange
    With tdb_Lista
        If mblnClickOk Then
            If Not .EOF And Not .BOF Then
                If mblnPrimeiraVez Then
                    
                    If Trim(tdb_Lista.Columns("PKId").Value) = "" Then
                        Exit Sub
                    End If
                    mblnClickOk = False
                    mblnAlterando = True
                    mblnPrimeiraVez = False
                    'mskstrCNPJCPF.Mask = ""
                    HabilitaDesabilitaObjeto mskstrCNPJCPF
                    LimpaObjeto Me, True
                    txtPKId = Val(tdb_Lista.Columns("PKId").Value)
                                        
                    dbcstrLogradouroC.BoundText = ""
                    LeDaTabelaParaObj gstrContribuinte, Me
                    
                    strSql = "Select * from " & gstrContribuinte & " Where Pkid = " & txtPKId
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                        dbcstrNome.Text = gstrENulo(adoResultado("strNome").Value)
                        dbcstrNomeFantasia.Text = gstrENulo(adoResultado("strNomeFantasia").Value)
                    End If
                    
                    gCorLinhaSelecionada tdb_Lista
                    If mobjAux Is Nothing Then
                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                    End If
                    
                    tab_3DCorrespondencia.TabEnabled(0) = chkblnResidenteNoMunicipio.Value
                    
                    CarregaTipoComunicacao txtPKId
                    'CarregaContasBancarias txtPKId
                    CarregaHistorico txtPKId
                    PreencheListaModulos txtPKId
                    
                    'txt_Codigo = Format(txtPKId, "00000000")
                    If cbobytNaturezaJuridica.ListIndex < 2 Then
                        tab_3DDadosGerais.TabEnabled(1) = True
                        tab_3DDadosGerais.TabEnabled(2) = True
                        tab_3DDadosGerais.TabEnabled(3) = True
                        tab_3DDadosGerais.TabEnabled(4) = True
                        tab_3DDadosGerais.TabEnabled(6) = True
                    End If
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar, gstrDeletar
                    
                    Select Case cbobytNaturezaJuridica.ListIndex
                        Case 0
                            mskstrCNPJCPF.Mask = "###\.###\.###\-##"
                            tab_3DDadosGerais.TabEnabled(5) = False
                            
                        Case 1
                            mskstrCNPJCPF.Mask = "##\.###\.###\/####\-##"
                            tab_3DDadosGerais.TabEnabled(5) = True
                            
                            CarregaSocios
                     End Select
                
                    'tab_3DDadosGerais.Tab = 0
                    'mskstrCNPJCPF.Mask = ""
                    'mskstrCNPJCPF.Text = gstrCGCCPFFormatado(.Columns("strCNPJCPF"))
                    mskstrCNPJCPF.Text = .Columns("strCNPJCPF")
                    
                    If mskstrCNPJCPF.Text = "" And cbobytNaturezaJuridica.ListIndex < 2 Then TrocaCorObjeto mskstrCNPJCPF, False
                    
                    strNomeAtual = dbcstrNome.Text
                    'strCnpjCpfAtual = gstrValorSemMascara(mskstrCNPJCPF)
                    strCnpjCpfAtual = mskstrCNPJCPF.ClipText
                    
                    blnEstadoInativo = IIf(chkblnInativo.Value = 0, False, True)
            '        dbcstrNome.SetFocus
            
                    'Cláudio
                    If Not ContribuinteHabilitadoAplicacao(Val(txtPKId.Text)) Then
                        If MsgBox("Este " & frmCadContribuinte.Tag & _
                                " não foi habilitado para este Módulo. Deseja habilitá-lo?", vbYesNo + vbQuestion) = vbYes Then
                            
                            Set gobjBanco = New clsBanco
                            gobjBanco.Execute "INSERT INTO " & gstrModuloContribuinte & "(intContribuinte, intItem, dtmDtAtualizacao, lngCodUsr) VALUES (" & txtPKId.Text & ", " & gintModulo & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")", True
                            
                            Set gobjBanco = Nothing
                            LocalizarContribuinte
                            
                        End If
                    End If
                    intCodigoAtual = Val(txtintCodigo.Text)
                End If
            End If
        End If
    End With
    Exit Sub

err_tdb_Lista_RowColChange:
Resume Next

End Sub

Private Function blnDadosOk() As Boolean
    Dim strSql       As String
    Dim adoResultado As New ADODB.Recordset
    
    On Error GoTo err_blnDadosOK
    blnDadosOk = False

    If MDIMenu.Tag = "Protocolo" Then
        If Trim(txtintCodigo.Text) = "" Then
            ExibeMensagem "A referência tem que ser digitada."
            txtintCodigo.SetFocus
            Exit Function
        End If
        
        If Not mblnAlterando Or (mblnAlterando And intCodigoAtual <> Val(LTrim(RTrim(txtintCodigo.Text)))) Then
            If gblnExisteCodigo(1, gstrContribuinte, "intCodigo", txtintCodigo.Text) Then
                ExibeMensagem "Já existe um registro com a mesma referência informada."
                If txtintCodigo.Enabled Then txtintCodigo.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Trim(dbcstrNome.Text) = "" Then
        ExibeMensagem "O nome tem que ser digitado."
        dbcstrNome.SetFocus
        Exit Function
    End If
    
    If cbobytNaturezaJuridica.ListIndex = -1 Then
        ExibeMensagem "Selecione um Tipo de Natureza Jurídica."
        Exit Function
    End If
    
    'Caso seja natureza diferente de Fisica ou Juridica nenhum destes campos serao obrigatorios
    If cbobytNaturezaJuridica.ListIndex < 2 Then
        
        If txt_dtmDebito.Text <> "" Then
            If gblnDataValida(txt_dtmDebito.Text) = False Then
                ExibeMensagem "A data de débito não é válida."
                txt_dtmDebito.SetFocus
                Exit Function
            End If
        End If
            
        If txtdtmDataNascimento.Text <> "" Then
            If gblnDataValida(txtdtmDataNascimento.Text) = False Then
                ExibeMensagem "A data de nascimento não é válida."
                txtdtmDataNascimento.SetFocus
                Exit Function
            ElseIf CVDate(txtdtmDataNascimento.Text) > CVDate(gstrDataDoSistema) Then
                ExibeMensagem "A data de nascimento não pode ser maior que a da atual"
                txtdtmDataNascimento.SetFocus
                Exit Function
            End If
        End If
        
        'Cláudio - Obrigatoriedade do campo CNPJ/CPF, so em ouvidoria e selecionado pessoa fisica pode ser nulo.
        'Pendência Nº 93

        If CStr(mskstrCNPJCPF.Text) = "" Then
            If UCase(App.ProductName) = "OUVIDORIA" Then
                If cbobytNaturezaJuridica.ListIndex = 1 Then
                    ExibeMensagem "CNPJ / CPF deve ser preenchido."
                    mskstrCNPJCPF.SetFocus
                    Exit Function
                End If
            Else
                ExibeMensagem "CNPJ / CPF deve ser preenchido."
                mskstrCNPJCPF.SetFocus
                Exit Function
            End If
        End If
        
        If mskstrCNPJCPF.ClipText <> "" Then
             If cbobytNaturezaJuridica.ListIndex = 0 Then
                 If Not gblnCPFOk(mskstrCNPJCPF) Then
                     ExibeMensagem "CPF inválido."
                     mskstrCNPJCPF.SetFocus
                     Exit Function
                 End If
             Else
                 If Not gblnCGCOk(mskstrCNPJCPF) Then
                     ExibeMensagem "CNPJ / CPF inválido."
                     mskstrCNPJCPF.SetFocus
                     Exit Function
                 End If
            End If
            
            If Not mblnAlterando And chkblnInativo.Value = 0 Then
                If blnVerificaCPFCNPJAtivo Then
                    ExibeMensagem "CNPJ / CPF, já existentes para um " & gstrContribuinteTituloSg & " Ativo."
                    mskstrCNPJCPF.SetFocus
                    Exit Function
                End If
            End If
        End If
    
        If tab_3DCorrespondencia.TabEnabled(0) = True Then
            If dbcintLogradouro.BoundText = "" Then
                ExibeMensagem "O logradouro residencial tem que ser informado."
                dbcintLogradouro.SetFocus
                Exit Function
            ElseIf blnVerificaLogradouro(dbcintLogradouro.BoundText) Then
                MsgBox "O logradouro selecionado está cancelado."
                dbcintLogradouro.SetFocus
                Exit Function
            End If
            
        End If
        If dbcstrLogradouroC.Text = "" Then
            ExibeMensagem "O logradouro do endereço de correspondência tem que ser informado."
            dbcstrLogradouroC.SetFocus
            Exit Function
        End If
        
        If txtstrBairroC.Text = "" Then
            ExibeMensagem "O bairro do endereço de correspondência tem que ser informado."
            txtstrBairroC.SetFocus
            Exit Function
        End If
        
        If dbcintMunicipioC.BoundText = "" Then
            ExibeMensagem "O município do endereço de correspondência tem que ser informado."
            dbcintMunicipioC.SetFocus
            Exit Function
        End If
        
        If dbcintUFC.BoundText = "" Then
            ExibeMensagem "O UF do endereço de correspondência tem que ser informado."
            dbcintUFC.SetFocus
            Exit Function
        End If
        
        If txtintCepC.Text = "" Then
            ExibeMensagem "O CEP do endereço de correspondência tem que ser informado."
            txtintCepC.SetFocus
            Exit Function
        End If
        
        If Trim(dbcintTipoLogradouro.Text) <> "" And Not dbcintTipoLogradouro.MatchedWithList Then
            ExibeMensagem "O campo tipo do logradouro deve ser preenchido corretamente."
            dbcintTipoLogradouro.SetFocus
            Exit Function
        End If
        
        If Trim(dbcintTituloLogradouro.Text) <> "" And Not dbcintTituloLogradouro.MatchedWithList Then
            ExibeMensagem "O campo título do logradouro deve ser preenchido corretamente."
            dbcintTituloLogradouro.SetFocus
            Exit Function
        End If
    
        If (mblnAlterando And UCase$(mskstrCNPJCPF) <> UCase$(strCnpjCpfAtual)) Then
            
            strSql = "SELECT strCNPJCPF FROM " & gstrContribuinte & " WHERE strCNPJCPF = '" & gstrValorSemMascara(Trim(mskstrCNPJCPF)) & "'"
            strSql = strSql & " AND PKID <> " & Trim(txtPKId)
            strSql = strSql & " AND blnInativo = 0 "
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
                If adoResultado.RecordCount >= 1 Then
                    ExibeMensagem "CNPJ / CPF já cadastrado para outro contribuinte."
                    mskstrCNPJCPF.SetFocus
                    Exit Function
                End If
            End If
        ElseIf (mblnAlterando And (chkblnInativo.Value = 0)) Then
            'ElseIf (mblnAlterando And (blnEstadoInativo = True And chkblnInativo.Value = 0)) Then
            strSql = "SELECT strCNPJCPF FROM " & gstrContribuinte & " WHERE strCNPJCPF in ( '" & Trim(mskstrCNPJCPF) & "','" & Trim(mskstrCNPJCPF.FormattedText) & "')"
            strSql = strSql & " AND PKID <> " & Trim(txtPKId)
            strSql = strSql & " AND blnInativo = 0 "
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
                If adoResultado.RecordCount >= 1 Then
                    ExibeMensagem "CNPJ / CPF já cadastrado para outro " & gstrContribuinteTituloSg & " Ativo."
                    mskstrCNPJCPF.SetFocus
                    Exit Function
                End If
            End If
                    
        End If
    
    End If
    
    blnDadosOk = True
    
err_blnDadosOK:

End Function

Sub NovoContribuinte()
    Dim i As Integer

    mblnAlterando = False
    lvw_TipoComunicacao.ListItems.Clear
    lvw_Contas.ListItems.Clear
    'txt_Codigo = Format(glngPegaProximaChave(gstrContribuinte, "PKId"), "00000000")
    'HabilitaDesabilitaObjeto txt_Codigo, False
    HabilitaDesabilitaObjeto txtPKId, False
    txt_Conteudo = ""
    txt_DescricaoConteudo = ""
    lbl_TipoComunicacao = "Tipo"
    
    txt_CodBanco = ""
    txt_CodAgencia = ""
    txt_Banco = ""
    txt_Agencia = ""
    txt_Conta = ""
    txt_DigitoVerificador = ""
    chk_ContaPublica.Value = 0
    chk_DebitoAutomatico.Value = 0
    txt_dtmDebito = ""
    txt_strMunicipio = ""
    txt_strBairro = ""
    txt_strUF = ""
    txtintCep = ""
    
    dtpdtmDataCadastro = Date
    
    HabilitaDesabilitaObjeto mskstrCNPJCPF, False
    HabilitaDesabilitaObjeto dbcstrNome, True
    HabilitaDesabilitaObjeto chkblnResidenteNoMunicipio, True
    HabilitaDesabilitaObjeto dtpdtmDataCadastro, True
    HabilitaDesabilitaObjeto dbcstrNomeFantasia, True
    HabilitaDesabilitaObjeto txtstrInscricaoEstadual, True
    HabilitaDesabilitaObjeto txtstrIdentidade, True
    HabilitaDesabilitaObjeto txtstrTituloEleitoral, True
    HabilitaDesabilitaObjeto txtdtmDataNascimento, True
    HabilitaDesabilitaObjeto txtstrCarteiraTrabalho, True

'    tab_3DCorrespondencia.Tab = 1
    
'    tab_3DCorrespondencia.TabEnabled(0) = False

    
    tab_3DDadosGerais.TabEnabled(1) = False
    tab_3DDadosGerais.TabEnabled(2) = False
    tab_3DDadosGerais.TabEnabled(3) = False
    tab_3DDadosGerais.TabEnabled(4) = False
    tab_3DDadosGerais.TabEnabled(5) = False
    tab_3DDadosGerais.TabEnabled(6) = False
    chkblnResidenteNoMunicipio.Value = vbChecked
    
    Set e = New XArrayDB 'Limpa Grid Socios
    e.Clear
    e.ReDim 0, 0, 0, 2
    Set tdb_Socios.Array = e
    tdb_Socios.ReBind
    tdb_Socios.Refresh
    txt_TotalDeCotas = ""
    
    cbobytNaturezaJuridica.ListIndex = 0
    SelecionaModuloAtual 0
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrSalvar, gstrDeletar
    tab_3DCorrespondencia.Tab = 0
    tab_3DDadosGerais.Tab = 0
    If MDIMenu.Tag = "Protocolo" Then
        gstrProximoCodigo txtintCodigo, gstrContribuinte, "intCodigo", gintCodSeguranca
    End If
End Sub

Private Sub tlb_TipoComunicacao_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case UCase(Button.Key)
        Case gstrSalvar
            If lvw_TipoComunicacao.ListItems.Count = 0 Then Exit Sub
            If lvw_TipoComunicacao.SelectedItem.Selected = False Then Exit Sub
            lvw_TipoComunicacao.SelectedItem.Selected = False
            
        Case gstrNovo
            mnu_Deletar.Visible = False
            mnu_Traco.Visible = False
            PopupMenu mnu_TipoComunicacao
            mnu_Deletar.Visible = True
            mnu_Traco.Visible = True
            
        Case gstrDeletar
            If lvw_TipoComunicacao.ListItems.Count = 0 Then Exit Sub
            If lvw_TipoComunicacao.SelectedItem.Selected = False Then Exit Sub
            lvw_TipoComunicacao.ListItems.Remove lvw_TipoComunicacao.SelectedItem.Index
    End Select
    txt_Conteudo = ""
    txt_DescricaoConteudo = ""
End Sub

Private Sub txt_Agencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Agencia
End Sub

Private Sub txt_Banco_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Banco
End Sub

Private Sub txt_CodAgencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_CodBanco_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrCodigoAnterior_GotFocus()
    'tab_3DDadosGerais.Tab = 0
End Sub

Private Sub txtstrCodigoAnterior_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtPKId
End Sub

Private Sub txt_Codigo1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_Conta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Conta
End Sub

Private Sub txt_Conteudo_Change()
    If lvw_TipoComunicacao.ListItems.Count = 0 Then Exit Sub
    If lvw_TipoComunicacao.SelectedItem.Selected = False Then Exit Sub
    lvw_TipoComunicacao.SelectedItem.SubItems(1) = Trim(txt_Conteudo)
End Sub

Private Sub txt_Conteudo_GotFocus()
    MarcaCampo txt_Conteudo
End Sub

Private Sub txt_Conteudo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Conteudo
End Sub

Private Sub txt_DescricaoConteudo_Change()
    If lvw_TipoComunicacao.ListItems.Count = 0 Then Exit Sub
    If lvw_TipoComunicacao.SelectedItem.Selected = False Then Exit Sub
    lvw_TipoComunicacao.SelectedItem.SubItems(2) = Trim(txt_DescricaoConteudo)
End Sub

Private Sub txt_DescricaoConteudo_GotFocus()
    MarcaCampo txt_DescricaoConteudo
End Sub

Private Sub txt_DescricaoConteudo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_DescricaoConteudo
End Sub

Private Sub txt_DigitoVerificador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_DigitoVerificador
End Sub

Private Sub txt_dtmDebito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDebito
End Sub

Private Sub txt_Nome1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub

Private Sub txtintCEP_LostFocus()
    If Not dbcintLogradouro.MatchedWithList Then
        txtintNumero.Text = ""
        txtstrComplemento.Text = ""
        txtintCep = gstrCEPFormatado(txtintCep)
        CepLogradouro txtintCep, dbcintLogradouro, txt_strBairro, , , , , , True, False, True, True, True, True, True, False
        'AjustaToolBar
    Else
        tab_3DCorrespondencia.Tab = 1
        If txtintCepC.Enabled Then txtintCepC.SetFocus
    End If
End Sub

Private Sub AjustaToolBar()
    gintCodSeguranca = intCodSeguranca
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
    
    Select Case tab_3DCorrespondencia.Tab
        Case 0
            If dbcintLogradouro.Enabled Then
                dbcintLogradouro.SetFocus
            End If
        Case 1
            If dbcintTipoLogradouro.Enabled Then
                dbcintTipoLogradouro.SetFocus
            End If
        Case 2
            If txtstrLogradouroD.Enabled Then
                txtstrLogradouroD.SetFocus
            End If
    End Select
End Sub

Private Sub txtintCepC_GotFocus()
    MarcaCampo txtintCepC
    strCep = txtintCepC
End Sub

Private Sub txtintCepC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepC
End Sub

Private Sub txtintCepC_LostFocus()
    
        
    If Trim(txtintCepC) <> strCep Then 'And Trim(dbcstrLogradouroC.Text) = ""
        dbcintTipoLogradouro.Text = ""
        dbcintTituloLogradouro.Text = ""
        dbcstrLogradouroC = ""
        txtintCodigoLogradouro.Text = ""
        txtstrBairroC.Text = ""
        txtstrDistritoC.Text = ""
        txtintNumeroC.Text = ""
        txtstrComplementoC.Text = ""

        txtintCepC = gstrCEPFormatado(txtintCepC)
        
        CepLogradouro txtintCepC, dbcstrLogradouroC, txtstrBairroC, dbcintMunicipioC, dbcintUFC, dbcintTipoLogradouro, dbcintTituloLogradouro, , False, False, True, True, True, True
        AjustaToolBar
        
    End If
End Sub

Private Sub txtintCepD_GotFocus()
    MarcaCampo txtintCepD
End Sub

Private Sub txtintCepD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepD
End Sub

Private Sub txtintCepD_LostFocus()
    If Len(Trim(txtstrLogradouroD)) = 0 Then
        txtintCepD = gstrCEPFormatado(txtintCepD)
        CepLogradouro txtintCepD, txtstrLogradouroD, txtstrBairroD, txtstrMunicipioD, dbcintUFD, dbcintTipoLogradouroD, dbcintTituloLogradouroD, , False, False, False, True, True, True
        AjustaToolBar
    End If
End Sub

Private Sub mskstrCNPJCPF_GotFocus()
    'tab_3DDadosGerais.Tab = 0
    MarcaCampo mskstrCNPJCPF
End Sub

Sub MontaColumnHeaders()
    With lvw_TipoComunicacao
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Tipo", 2000
        .ColumnHeaders.Add 2, , "Conteúdo", 3000
        .ColumnHeaders.Add 3, , "Descrição", 3210
    End With
    With lvw_Contas
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "CodBanco", 0
        .ColumnHeaders.Add 2, , "Banco", 2700
        .ColumnHeaders.Add 3, , "codAgência", 0
        .ColumnHeaders.Add 4, , "Agência", 2000
        .ColumnHeaders.Add 5, , "Conta", 1500
        .ColumnHeaders.Add 6, , "DV", 500
        .ColumnHeaders.Add 7, , "Pública", 800
        .ColumnHeaders.Add 8, , "Débito Automático", 1500
        .ColumnHeaders.Add 9, , "Data Início Débito", 1500
    End With
'    With lvw_Historico
'        .ColumnHeaders.Clear
'        .ColumnHeaders.Add 1, , "Código", 1000
'        .ColumnHeaders.Add 2, , "Data / Hora", 2000
'        .ColumnHeaders.Add 3, , "Tipo da transação", 3000
'        .ColumnHeaders.Add 4, , "Valor", 1500
'    End With
End Sub

Private Sub HabilitaDesabilitaObjeto(mobjObjeto As Object, Optional blnFlag As Boolean)
    If Not blnFlag Then  'Desabilita
        If TypeOf mobjObjeto Is TextBox Then
            mobjObjeto.Text = ""
            mobjObjeto.BackColor = &HC0C0C0
        ElseIf TypeOf mobjObjeto Is MaskEdBox Then
            mobjObjeto.Mask = ""
            mobjObjeto.Text = ""
            mobjObjeto.BackColor = &HC0C0C0
        ElseIf TypeOf mobjObjeto Is DTPicker Then
            mobjObjeto = Format(Date, "dd/mm/yyyy")
        ElseIf TypeOf mobjObjeto Is CheckBox Then
            mobjObjeto.Value = 0
        End If
        mobjObjeto.Enabled = False
    Else    'Habilita
        If TypeOf mobjObjeto Is TextBox Then
'            mobjObjeto.Text = ""
            mobjObjeto.BackColor = &H80000005
        ElseIf TypeOf mobjObjeto Is MaskEdBox Then
            mobjObjeto.Mask = ""
            mobjObjeto.Text = ""
            mobjObjeto.BackColor = &H80000005
        ElseIf TypeOf mobjObjeto Is DTPicker Then
'            mobjObjeto.Date = Format(Date, "dd/mm/yyyy")
        ElseIf TypeOf mobjObjeto Is CheckBox Then
'            mobjObjeto.Value = 0
        End If
        mobjObjeto.Enabled = True
    End If
End Sub

Private Function blnVerificaCPFCNPJAtivo() As Boolean
    
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT pkid "
    strSql = strSql & " FROM " & gstrContribuinte
    strSql = strSql & " WHERE strCNPJCPF in('" & gstrValorSemMascara(mskstrCNPJCPF) & "','" & mskstrCNPJCPF.FormattedText & "') "
    strSql = strSql & " AND blnInativo = 0 "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then blnVerificaCPFCNPJAtivo = True
    End If

End Function

Sub CarregaTipoComunicacao(intCodContribuinte As Long)

'******************************************************************************************
' Data: 10/03/2003
' Alteração: - Alteração do comando SELECT devido a incompatibilidades de estrutura dos
'            outer joins entre o SQL Server e o Oracle. Os joins da cláusula FROM foram
'            substituídos por joins correspondentes na cláusula WHERE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    lvw_TipoComunicacao.ListItems.Clear
    lbl_TipoComunicacao = "Tipo"
    txt_Conteudo = ""
    txt_DescricaoConteudo = ""
    
    strSql = ""
    strSql = strSql & "Select TP.strDescricao TipoComunicacao, "
    strSql = strSql & "FC.intTipoDeComunicacao, FC.strDescricao, FC.strConteudo "
    strSql = strSql & "From " & gstrTipoDeComunicacao & " TP "
'    strSql = strSql & "Left Join " & gstrFormaDeComunicacao & " FC "
    strSql = strSql & ", " & gstrFormaDeComunicacao & " FC "
'    strSql = strSql & "On TP.PKId = FC.intTipoDeComunicacao "
    strSql = strSql & "Where FC.intContribuinte = " & intCodContribuinte & " "
    strSql = strSql & "And TP.PKId =" & strOUTJSQLServer & " FC.intTipoDeComunicacao " & strOUTJOracle
    strSql = strSql & "Order By FC.intSequencia"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set oList = lvw_TipoComunicacao.ListItems.Add(, , Trim(!TipoComunicacao))
                oList.SubItems(1) = gstrVerificaCampoNulo(!strConteudo)
                oList.SubItems(2) = gstrVerificaCampoNulo(!strDescricao)
                oList.Tag = gstrVerificaCampoNulo(!intTipoDeComunicacao)
                .MoveNext
            Loop
        End With
    End If
    If lvw_TipoComunicacao.ListItems.Count <> 0 Then
        lvw_TipoComunicacao.SelectedItem.Selected = False
    End If
End Sub

Sub CarregaContasBancarias(intCodContribuinte As Long)
    Dim strSql As String
    
    lvw_Contas.ListItems.Clear
    
    strSql = ""
    strSql = strSql & "Select CB.PKId, BO.intBanco, BO.strDescricao AS Banco, AG.intNumero, AG.strDescricao AS Agencia, CB.strConta, CB.strDigitoVerificador, CB.blnContaPublica, blnDebitoAutomatico, dtmDebito "
    strSql = strSql & "From " & gstrContaBancaria & " CB, "
    strSql = strSql & gstrBanco & " BO, "
    strSql = strSql & gstrAgencia & " AG "
    strSql = strSql & "Where intContribuinte = " & intCodContribuinte & " "
    strSql = strSql & "AND CB.intAgencia = AG.PKId "
    strSql = strSql & "AND CB.intBanco = BO.PKId "
    
    LeDaTabelaParaObj gstrContaBancaria, lvw_Contas, strSql
    
    If lvw_Contas.ListItems.Count <> 0 Then
        lvw_Contas.SelectedItem.Selected = False
    End If
End Sub

Sub CarregaHistorico(intCodContribuinte As Long)
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & " Select * "
    strSql = strSql & " FROM " & gstrHistoricoContribuinte
    strSql = strSql & " WHERE intContribuinte = " & txtPKId.Text
    strSql = strSql
    
    LeDaTabelaParaObj gstrHistoricoContribuinte, tdb_Historico, strSql
    
End Sub

Function blnGravaTipoComunicacao(intCodContribuinte As Long) As Boolean
    Dim strSql As String
    Dim intFor   As Integer
    
    DeletaTipoComunicacao intCodContribuinte
    
    With lvw_TipoComunicacao
        For intFor = 1 To .ListItems.Count
            strSql = ""
            strSql = strSql & "Insert Into " & gstrFormaDeComunicacao & " "
            strSql = strSql & "(intContribuinte, intTipoDeComunicacao, strConteudo, strDescricao, dtmdtAtualizacao, lngCodUsr, "
            strSql = strSql & "intSequencia) Values ("
            strSql = strSql & intCodContribuinte & ", "
            strSql = strSql & .ListItems(intFor).Tag & ", '"
            strSql = strSql & .ListItems(intFor).SubItems(1) & "', '"
            strSql = strSql & .ListItems(intFor).SubItems(2) & "', "
            strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSql = strSql & glngCodUsr & ", "
            strSql = strSql & intFor & ")"
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSql
        Next
    End With
    blnGravaTipoComunicacao = True
End Function

Sub DeletaTipoComunicacao(intCodContribuinte As Long)
    Dim strSql As String

    strSql = ""
    strSql = strSql & "Delete From " & gstrFormaDeComunicacao & " "
    strSql = strSql & "Where intContribuinte = " & intCodContribuinte
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
End Sub

Sub PreencheMenuPopup()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    Dim intI         As Integer
    
    On Error GoTo Err_Handle
    intI = 0
    
    strSql = ""
    strSql = strSql & "Select TP.PKId, TP.strDescricao TipoComunicacao "
    strSql = strSql & "From " & gstrTipoDeComunicacao & " TP "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                intI = intI + 1
                Load mnu_Lista(intI)
                mnu_Lista(intI).Caption = Trim(!TipoComunicacao)
                mnu_Lista(intI).Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
    mnu_Lista(0).Visible = False
    
Err_Handle:
End Sub

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
    gstrProximoCodigo txtintCodigo, gstrContribuinte, "intCodigo", gintCodSeguranca
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub txtintCodigoLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintCodigoLogradouroD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrBairroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrBairroC
End Sub

Private Sub txtstrBairroD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrCarteiraTrabalho_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCarteiraTrabalho
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
End Sub

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub txtdtmDataNascimento_GotFocus()
    MarcaCampo txtdtmDataNascimento
End Sub

Private Sub txtdtmDataNascimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataNascimento
End Sub

Private Sub txtdtmDataNascimento_LostFocus()
    txtdtmDataNascimento = gstrDataFormatada(txtdtmDataNascimento)
End Sub


Private Sub txtstrDistritoC_GotFocus()
    MarcaCampo txtstrDistritoC
End Sub

Private Sub txtstrDistritoC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDistritoC
End Sub

Private Sub txtstrDistritoC_LostFocus()
    tab_3DCorrespondencia.Tab = 2
    If dbcintTipoLogradouroD.Enabled Then
        dbcintTipoLogradouroD.SetFocus
    Else
        If txtstrLogradouroD.Enabled Then txtstrLogradouroD.SetFocus
    End If
End Sub

Private Sub txtstrIdentidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrIdentidade
End Sub

Private Sub txtstrInscricaoEstadual_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrInscricaoEstadual
End Sub

Private Sub dbcstrLogradouroC_GotFocus()
    MarcaCampo dbcstrLogradouroC
End Sub

Private Sub dbcstrLogradouroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcstrLogradouroC
End Sub

Private Sub txtstrLogradouroD_GotFocus()
    AjustaToolBar
    MarcaCampo txtstrLogradouroD
End Sub

Private Sub txtstrLogradouroD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLogradouroD
End Sub

Private Sub txtstrLoteD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrMunicipioD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcstrNome_GotFocus()
'    tab_3DDadosGerais.Tab = 0
    MarcaCampo dbcstrNome
End Sub

Private Sub dbcstrNome_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcstrNome
End Sub

Private Sub txtintNumero_GotFocus()
    MarcaCampo txtintNumero
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumero
End Sub

Private Sub txtstrComplementoC_GotFocus()
    MarcaCampo txtstrComplementoC
End Sub

Private Sub txtstrComplementoC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplementoC
End Sub

Private Sub txtintNumeroC_GotFocus()
    MarcaCampo txtintNumeroC
End Sub

Private Sub txtintNumeroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumeroC
End Sub

Private Sub txtstrComplementoD_GotFocus()
    MarcaCampo txtstrComplementoD
End Sub

Private Sub txtstrComplementoD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplementoD
End Sub

Private Sub txtintNumeroD_GotFocus()
    MarcaCampo txtintNumeroD
End Sub

Private Sub txtintNumeroD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumeroD
End Sub

Private Function strQueryGridContribuinte() As String
Dim strSql              As String
Dim strSQLModulos       As String

    strSql = ""
    strSQLModulos = "SELECT "
    'dbo.tblItem INNER JOIN
    'dbo.tblModuloContribuinte ON dbo.tblItem.PKId = dbo.tblModuloContribuinte.intItem INNER JOIN
    'dbo.tblContribuinte ON dbo.tblModuloContribuinte.intContribuinte = dbo.tblContribuinte.PKId


    strSql = strSql & "SELECT CO.PKId, CO.strNome, CO.strCNPJCPF, CO.strNomeFantasia "
    strSql = strSql & "FROM " & gstrContribuinte & " CO, " & gstrItens & " IT, " & gstrModuloContribuinte & " MC "
    strSql = strSql & "WHERE IT.PKId = MC.intItem AND "
    strSql = strSql & "MC.intContribuinte = CO.PKId AND IT.PKId =" & gintModulo
    Select Case bytOrdenacao
      
        Case Is = 1
            strSql = strSql & " ORDER BY CO.PKId" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSql = strSql & " ORDER BY CO.strNome" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strSql = strSql & " ORDER BY CO.strCNPJCPF" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 4
            strSql = strSql & " ORDER BY CO.strNomeFantasia" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
   
    
strQueryGridContribuinte = strSql



End Function

'Sub MontaArray(Optional strQuery As String)
'    Dim adoContribuintes As ADODB.Recordset
'    Dim varAux           As Variant
'    Dim strSql           As String
'
'
'    If Trim(strQuery) = "" Then
'        strSql = ""
'        strSql = strSql & "SELECT PKId, strNome, strCNPJCPF, bytNaturezaJuridica "
'        strSql = strSql & "FROM " & gstrContribuinte & " "
'        strSql = strSql & "ORDER BY strNome"
'        strQuery = strSql
'    End If
'
'
'    Set x = New XArrayDB
'
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strQuery, 5, adoContribuintes) Then
'        With adoContribuintes
'            If Not .EOF Then
'                x.ReDim 0, .RecordCount - 1, 0, 2
'                Do While Not .EOF
'                    varAux = Format(!PKId, "00000000")
'                    x(.AbsolutePosition - 1, 0) = varAux
'                    varAux = Trim(!strNome)
'                    x(.AbsolutePosition - 1, 1) = varAux
'                    varAux = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!strCNPJCPF), IIf(!bytNaturezaJuridica = 0, "PF", "PJ"))
'                    x(.AbsolutePosition - 1, 2) = varAux
'                    .MoveNext
'                Loop
'            End If
'        End With
'    End If
'
'    Set tdb_Lista.Array = x
'    tdb_Lista.ReBind
'    tdb_Lista.Refresh
'
'    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
'End Sub

Function blnDeletaContribuinte(intAuxPKId As Long) As Boolean
    Dim strSql As String
    If Trim(txtPKId) = "" Then
        Exit Function
    End If
    
    If VerificaEmpenho Then
       ExibeMensagem "Este contribuinte possui empenhos e não poderá ser excluído."
       Exit Function
    End If
    
    If MsgBox("Confirma a exclusão do contribuinte '" & tdb_Lista.Columns("strNome").Value & "' ?", vbYesNo + vbQuestion) = vbYes Then
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        
        DeletaTipoComunicacao intAuxPKId
    
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrContribuinte & " "
        strSql = strSql & "WHERE PKId = " & txtPKId
        
        If Not gobjBanco.Execute(strSql) Then
            gobjBanco.ExecutaRollbackTrans
        End If
        
        gobjBanco.ExecutaCommitTrans
        blnDeletaContribuinte = True
    End If
End Function

Private Function strQueryAplicar() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNome "
    strSql = strSql & "FROM " & gstrContribuinte
    strSql = strSql & " ORDER BY strNome "
strQueryAplicar = strSql
End Function

Private Function strQueryRelatorio() As String
    Dim strSql As String

    
    strSql = "SELECT "
    strSql = strSql & "CT.PKId,"
    strSql = strSql & gstrCASEWHEN("CT.blnInativo", "1,'Não'", "'Sim'") & " blnInativo ,"
    strSql = strSql & gstrCASEWHEN("CT.bytNaturezaJuridica", "1,'Jurídica'", "'Física'") & " bytNaturezaJuridica ,"
    strSql = strSql & "CT.strNome, "
    strSql = strSql & "CT.strCNPJCPF, "
    strSql = strSql & "CT.strNomeFantasia ,"
    strSql = strSql & gstrCASEWHEN("CT.blnResidenteNoMunicipio", "1,'Sim'", "'Não'") & "blnResidenteNoMunicipio,"
    strSql = strSql & "CT.strInscricaoEstadual,"
    strSql = strSql & "CT.dtmDataCadastro,"
    strSql = strSql & "CT.strIdentidade,"
    strSql = strSql & "CT.strTituloEleitoral,"
    strSql = strSql & "CT.dtmDataNascimento,"
    strSql = strSql & "CT.strCarteiraTrabalho,"
    strSql = strSql & "CT.CDC intCDC,"
    strSql = strSql & "FC.strconteudo, "
    strSql = strSql & "(SELECT TC.strDescricao FROM " & gstrTipoDeComunicacao & " TC WHERE TC.PKID = FC.intTipodeComunicacao ) strDescricao,"
    'campos do endereço Residencial
    strSql = strSql & "LO.strDescricao STRENDERECO,"
    strSql = strSql & "CT.INTNUMERO,"
    strSql = strSql & "CT.STRCOMPLEMENTO,"
    strSql = strSql & "BA.strDescricao as STRBAIRRO,"
    strSql = strSql & " (SELECT MU.strDescricao FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId) AS  STRMUNICIPIO,"
    strSql = strSql & " (SELECT UF.strSigla FROM " & gstrUF & " UF WHERE UF.PKId = (SELECT MU.intUF FROM " & gstrCidade & " MU WHERE BA.intMunicipio = MU.PKId )) as STRUF,"
    strSql = strSql & "LO.intCep INTCEP,"
    
    'campos do endereço de Correspondencia
    strSql = strSql & "TLC.strSigla intTipoLogradouro,"
    strSql = strSql & "TLLC.strSigla intTituloLogradouro,"
    strSql = strSql & "CT.intCodigoLogradouro,"
    strSql = strSql & "CT.strLogradouroC,"
    strSql = strSql & "CT.intNumeroC,"
    strSql = strSql & "CT.strComplementoC,"
    strSql = strSql & "CT.strBairroC,"
    strSql = strSql & "MPC.strDescricao intMunicipioC,"
    strSql = strSql & "UFC.strSigla intUFC,"
    strSql = strSql & "CT.intCepC,"
    strSql = strSql & "CT.strDistritoC,"
    
    'campos do endereço distrito fiscal
    strSql = strSql & "TLD.strSigla intTipoLogradouroD,"
    strSql = strSql & "TLLC.strSigla intTituloLogradouroD,"
    strSql = strSql & "CT.intCodigoLogradouroD,"
    strSql = strSql & "CT.strLogradouroD,"
    strSql = strSql & "CT.intNumeroD,"
    strSql = strSql & "CT.strComplementoD,"
    strSql = strSql & "CT.strBairroD,"
    strSql = strSql & "CT.strMunicipioD,"
    strSql = strSql & "UFD.strSigla  intUFD,"
    strSql = strSql & "CT.intCepD,"
    strSql = strSql & "CT.strSetorD,"
    strSql = strSql & "CT.strQuadraD,"
    strSql = strSql & "CT.strLoteD,"
    strSql = strSql & "CT.strDistritoD"
    strSql = strSql & " FROM "
    
    strSql = strSql & gstrContribuinte & " CT "
    strSql = strSql & "INNER JOIN " & gstrModuloContribuinte & " MC ON MC.intContribuinte = CT.PKID "
    strSql = strSql & "INNER JOIN " & gstrItens & " IT ON IT.PKID = MC.intItem "
    strSql = strSql & "LEFT OUTER JOIN " & gstrLogradouro & " LO ON LO.PKID = CT.intLogradouro "
    strSql = strSql & "LEFT OUTER JOIN " & gstrBairro & " BA ON BA.PKID = LO.intBairro "
    strSql = strSql & "LEFT OUTER JOIN " & gstrCidade & " MPC ON MPC.PKID = CT.intMunicipioC "
    strSql = strSql & "LEFT OUTER JOIN " & gstrUF & " UFC ON UFC.PKID = CT.intUFC "
    strSql = strSql & "LEFT OUTER JOIN " & gstrUF & " UFD ON UFD.PKID = CT.intUFD "
    strSql = strSql & "LEFT OUTER JOIN " & gstrFormaDeComunicacao & " FC ON FC.intContribuinte = CT.PKID "
    strSql = strSql & "LEFT OUTER JOIN " & gstrTipoLogradouro & " TLC ON TLC.PKID = CT.intTipoLogradouro "
    strSql = strSql & "LEFT OUTER JOIN " & gstrTituloLogradouro & " TLLC ON TLLC.PKID = CT.intTituloLogradouro "
    strSql = strSql & "LEFT OUTER JOIN " & gstrTipoLogradouro & " TLD ON TLD.PKID = CT.intTipoLogradouroD "
    strSql = strSql & "LEFT OUTER JOIN " & gstrTituloLogradouro & " TLLD ON TLLD.PKID = CT.intTituloLogradouroD "
    
    strSql = strSql & " WHERE "
'    strSql = strSql & "BA.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & " LO.intBairro "
'    strSql = strSql & " AND LO.PKID " & strOUTJOracle & "=" & strOUTJSQLServer & " CT.intLogradouro "
'    strSql = strSql & " AND LO.Dtmdtexclusao is null "
'    strSql = strSql & " AND UFC.PKID " & strOUTJOracle & "=" & strOUTJSQLServer & " CT.intUFC "
'    strSql = strSql & " AND MPC.PKID " & strOUTJOracle & "=" & strOUTJSQLServer & " CT.intMunicipioC "
'    strSql = strSql & " AND TLC.PKID " & strOUTJOracle & "=" & strOUTJSQLServer & " CT.intTipoLogradouro "
'    strSql = strSql & " AND TLLC.PKID " & strOUTJOracle & "=" & strOUTJSQLServer & " CT.intTituloLogradouro "
'    strSql = strSql & " AND TLD.PKID " & strOUTJOracle & "=" & strOUTJSQLServer & " CT.intTipoLogradouroD "
'    strSql = strSql & " AND TLLD.PKID " & strOUTJOracle & "=" & strOUTJSQLServer & " CT.intTituloLogradouroD "
'    strSql = strSql & " AND UFD.PKID " & strOUTJOracle & "=" & strOUTJSQLServer & " CT.intUFD "
'    strSql = strSql & " AND FC.intcontribuinte " & strOUTJOracle & "=" & strOUTJSQLServer & " CT.PKID "
'    strSql = strSql & " AND IT.PKId = MC.intItem AND"
'    strSql = strSql & " MC.intContribuinte = CT.Pkid AND"
    
    strSql = strSql & " IT.Pkid =" & gintModulo
    strSql = strSql & " AND CT.Pkid =" & txtPKId
    
    
        
    strQueryRelatorio = strSql
    
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark As Variant
    Dim strSql      As String
    Dim lngLinha    As Long
    Dim blnFlag     As Boolean
    Dim blnAlterando As Boolean
    
    If strModoOperacao = UCase("IMPRIMIR") Then
        If Trim(txtPKId) = "" Then
            ExibeMensagem "É necessário selecionar um " & gstrContribuinteTituloSg & " para imprimir."
            Exit Sub
        End If
        strSql = strQueryRelatorio
        rptListagemCredores.Caption = "Listagem de " & gstrContribuinteTituloPl
        ToolBarGeral strModoOperacao, gstrContribuinte, mblnAlterando, tdb_Lista, Me, mobjAux, strSql, , rptListagemCredores, strQueryRelatorio
        Exit Sub
    End If
    On Error GoTo err_MantemForm
    
    blnFlag = gblnListagemAutomatica
    gblnListagemAutomatica = False
    
    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If
    
    strSql = ""

    Dim intCodContribuinte As Long
                    
    Select Case UCase(strModoOperacao)
        Case UCase(gstrNovo)
            LimpaObjeto Me, mblnAlterando
            NovoContribuinte
            
        Case UCase(gstrImprimir)
            ImprimeContribuinte
            
        Case UCase(gstrSalvar)
            If Not mblnAlterando Then
                If gblnSistemaDemonstracao(gstrContribuinte, 50) Then
                    gblnListagemAutomatica = blnFlag
                    Exit Sub
                End If
            End If
                        
            If blnUsuarioOutroModulo And Not mblnAlterando Then
                If MDIMenu.actBarra.Bands(gstrMnuArquivo).Tools(gstrSalvar).Enabled Then
                    If MsgBox(gstrContribuinteTituloSg & " já cadastrado. Deseja habilitá-lo para este módulo?", vbYesNo + vbQuestion) = vbYes Then
                        
                        Set gobjBanco = New clsBanco
                        
                        gobjBanco.Execute "INSERT INTO " & gstrModuloContribuinte & "(intContribuinte, intItem) VALUES (" & intUsuario & ", " & gintModulo & ")", True
                        
                        gblnListagemAutomatica = blnFlag
                        
                        LimpaObjeto Me
                        
                        NovoContribuinte
                        
                        Set gobjBanco = Nothing
                        
                    Else
                        
                        LimpaObjeto Me
                        
                        NovoContribuinte
                        
                    End If
                Else
                    MsgBox gstrContribuinteTituloSg & " já cadastrado. Para habilitá-lo para este módulo, contacte o administrador do sistema."
                End If
            Else
           
                If blnDadosOk Then
                    blnAlterando = mblnAlterando
                    'dbcstrLogradouroC.Tag = gstrQueryLogradouro & ";L.strDescricao"
                    If ToolBarGeral(strModoOperacao, gstrContribuinte, mblnAlterando, tdb_Lista, Me, mobjAux, , strQueryAplicar, , , True) Then
                        gblnListagemAutomatica = blnFlag
                        
                        If blnAlterando Then
                            intCodContribuinte = tdb_Lista.Columns("PKId").Value
                        Else
                            intCodContribuinte = glngRetornaPkidTabelaPai("Seq" & gstrContribuinte, gstrContribuinte)
                        End If
                        
                        If blnGravaTipoComunicacao(intCodContribuinte) Then
                        End If
                        
                        If blnGravaAplicacoes(intCodContribuinte, blnAlterando) Then
                        End If
                        
                        
                        LeDaTabelaParaObj gstrContribuinte, tdb_Lista, "Select CO.*, " & gstrCASEWHEN("blnInativo", "0,'Não',1,'Sim'") & " strInativo From " & gstrContribuinte & " co Where pkid = " & intCodContribuinte
                        NovoContribuinte
                    End If
                    'dbcstrLogradouroC.Tag = ""
                End If
            End If
            
        Case UCase(gstrDeletar)
            If blnDeletaContribuinte(txtPKId) Then
                LimpaObjeto Me, mblnAlterando
                NovoContribuinte
                Set tdb_Lista.DataSource = Nothing
                tdb_Lista.Refresh
            End If
        Case UCase(gstrLocalizar)
        
            LocalizarContribuinte
            
        Case UCase(gstrPreencherLista)
            If Me.ActiveControl.Name = dbcintLogradouro.Name Then
                dbcintLogradouro.Tag = strQueryLogradouro(False) & ";L.strDescricao"
                PreencherListaDeOpcoes Me.ActiveControl
                dbcintLogradouro.Tag = strQueryLogradouro(True) & ";L.strDescricao"
                
            ElseIf Me.ActiveControl.Name = dbcstrLogradouroC.Name Then
                dbcstrLogradouroC.Tag = gstrQueryLogradouro(, , , True) & ";L.strDescricao"
                PreencherListaDeOpcoes Me.ActiveControl
                dbcstrLogradouroC.Tag = ""
            ElseIf Me.ActiveControl.Name = dbcstrNome.Name Then
                PreencheNomeContribuinte
            Else
                PreencherListaDeOpcoes Me.ActiveControl
            End If
        Case Else
            ToolBarGeral strModoOperacao, gstrContribuinte, mblnAlterando, tdb_Lista, Me, mobjAux, , strQueryAplicar

    End Select

    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
    dbcintTipoLogradouro.Enabled = True
    TrocaCorObjeto dbcintTipoLogradouro, False
    dbcintTituloLogradouro.Enabled = True
    TrocaCorObjeto dbcintTituloLogradouro, False
    dbcintTipoLogradouroD.Enabled = True
    TrocaCorObjeto dbcintTipoLogradouroD, False
    dbcintTituloLogradouroD.Enabled = True
    TrocaCorObjeto dbcintTituloLogradouroD, False
    cmd_TipoLogradouro.Enabled = True
    TrocaCorObjeto cmd_TipoLogradouro, False
    cmd_TituloLogradouro.Enabled = True
    TrocaCorObjeto cmd_TituloLogradouro, False
    cmd_TipoLogradouroD.Enabled = True
    TrocaCorObjeto cmd_TipoLogradouroD, False
    cmd_TituloLogradouroD.Enabled = True
    TrocaCorObjeto cmd_TituloLogradouroD, False
    
    gblnListagemAutomatica = blnFlag
    
err_MantemForm:
End Sub

Private Sub dbcstrNome_LostFocus()
    If dbcstrNomeFantasia.Enabled = True And Len(Trim(dbcstrNomeFantasia)) = 0 Then
        dbcstrNomeFantasia.Text = dbcstrNome.Text
    End If
End Sub

Private Sub dbcstrNomeFantasia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcstrNomeFantasia
End Sub

Private Sub txtstrQuadraD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrSetorD_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrTituloEleitoral_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrTituloEleitoral
End Sub

Private Sub ImprimeContribuinte()
    
'******************************************************************************************
' Data: 06/03/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função.
'            - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'           nativos
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 10/03/2003
' Alteração: - Alteração do comando SELECT devido a incompatibilidades de estrutura dos
'            outer joins entre o SQL Server e o Oracle. Os joins da cláusula FROM foram
'            substituídos por joins correspondentes na cláusula WHERE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 11/03/2003
' Alteração: - Retirada referências desnecessárias ao owner das tabelas no select.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql  As String
    
    strSql = ""
'    strSql = strSql & "SELECT   dbo.tblContribuinte.strNome, " & _
'                                "dbo.tblContribuinte.PKId, " & _
'                                "dbo.tblContribuinte.bytNaturezaJuridica, " & _
'                                "dbo.tblContribuinte.strCNPJCPF, " & _
'                                "dbo.tblContribuinte.blnResidenteNoMunicipio, " & _
'                                "dbo.tblContribuinte.dtmDataCadastro, "
    strSql = strSql & "SELECT   " & gstrContribuinte & ".strNome, " & _
                                gstrContribuinte & ".PKId, " & _
                                gstrContribuinte & ".bytNaturezaJuridica, " & _
                                gstrContribuinte & ".strCNPJCPF, " & _
                                gstrContribuinte & ".blnResidenteNoMunicipio, " & _
                                gstrContribuinte & ".dtmDataCadastro, "
'                                "CASE  bytNaturezaJuridica WHEN 0 THEN 'Física' " & _
'                                                         " WHEN 1 THEN 'Jurídica' " & _
'                                                         " WHEN 2 THEN 'SC' " & _
'                                                         " WHEN 3 THEN 'Outros' " & _
'                                "END AS StrNaturezaJuridica, "
    strSql = strSql & gstrCASEWHEN("bytNaturezaJuridica", _
                                    "0,'Física'," & _
                                    "1,'Jurídica'," & _
                                    "2,'SC'," & _
                                    "3,'Outros'") & _
                                " AS StrNaturezaJuridica, "
'                                "CASE blnResidenteNoMunicipio WHEN 1 THEN 'Sim' " & _
'                                                            " WHEN 2 THEN 'Não' " & _
'                                "END AS strResidenteNoMunicipio "
    strSql = strSql & gstrCASEWHEN("blnResidenteNoMunicipio", _
                                    "1,'Sim'," & _
                                    "2,'Não'") & _
                                " AS strResidenteNoMunicipio "
'    strSql = strSql & " FROM " & _
'                                "dbo.tblModuloContribuinte  INNER JOIN dbo.tblContribuinte " & _
'                      " ON " & _
'                                "dbo.tblModuloContribuinte.intContribuinte = dbo.tblContribuinte.PKId "
    strSql = strSql & " FROM " & gstrModuloContribuinte & ", " & _
                            gstrContribuinte
'    strSql = strSql & " WHERE (dbo.tblModuloContribuinte.intItem = "
    strSql = strSql & " WHERE " & gstrModuloContribuinte & ".intContribuinte = " & gstrContribuinte & ".PKId AND "
    strSql = strSql & " (" & gstrModuloContribuinte & ".intItem = "
'                              "(SELECT PKId From dbo.tblItem WHERE (strItem ='" & App.Title & "'"
    strSql = strSql & "(SELECT PKId From " & gstrItens & " WHERE (UPPER(strItem) ='" & UCase$(App.Title) & "'"
'                              ") AND (LEN(strCodItem) = 1)))"
    strSql = strSql & ") AND (" & strLen & "(strCodItem) = 1)))"
'                      " ORDER BY dbo.tblContribuinte.PKId, dbo.tblContribuinte.strNome"
    strSql = strSql & " ORDER BY " & gstrContribuinte & ".PKId, " & gstrContribuinte & ".strNome"
    ImprimeRelatorio rptCadContribuinte, strSql
End Sub

Private Sub LocalizarContribuinte()
Dim strSql As String
Dim strCondicao As String
Dim strValor As String
Dim strCampo As String

Dim i As Integer

strCondicao = ""

With Me
    For i = 0 To .Controls.Count - 1
        
        If Not TypeOf .Controls(i) Is Label Then 'Elimina os Label's da pesquisa
            'Elimina objetos indesejáveis
            If UCase(.Controls(i).Name) <> "TXT_CONTEUDO" And UCase(.Controls(i).Name) <> "TXT_DESCRICAOCONTEUDO" And _
            UCase(Left(.Controls(i).Name, 3)) <> "IMG" And UCase(.Controls(i).Name) <> "TXT_TOTALDECOTAS" And _
            UCase(Left(.Controls(i).Name, 3)) <> "LVW" And UCase(Left(.Controls(i).Name, 3)) <> "TLB" And _
            UCase(.Controls(i).Name) <> "CHK_CONTAPUBLICA" And UCase(.Controls(i).Name) <> "CHK_DEBITOAUTOMATICO" And _
            UCase(.Controls(i).Name) <> "TXT_DTMDEBITO" And InStr(1, UCase(.Controls(i).Name), "_NOME") = 0 And _
            InStr(1, UCase(.Controls(i).Name), "_CODIGO") = 0 And UCase(.Controls(i).Name) <> "TXT_STRUF" And UCase(.Controls(i).Name) <> "TXT_STRMUNICIPIO" And UCase(.Controls(i).Name) <> "TXT_STRBAIRRO" Then
                If Not (TypeOf .Controls(i) Is OptionButton) Or .Controls(i) = True Then 'Elimina OptionButton desmarcado
                    If TypeOf .Controls(i) Is TextBox Then
                        If Trim(.Controls(i).Text) <> "" Then
                            
                            If InStr(1, .Controls(i).Name, "Cep") > 0 Then
                                strValor = gstrValorSemMascara(Trim(.Controls(i).Text))
                            Else
                                strValor = Trim(.Controls(i).Text)
                            End If
                            
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, strCampo, "_") > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            
                            If UCase(strCampo) = "PKID" Then
                                strCampo = "CO." & strCampo
                            End If
                            
                            If UCase(Trim(strCampo)) = "INTCODIGO" Then
                                strCampo = "CO." & strCampo
                            End If
                            
                            If InStr(1, "%", strValor) > 0 Then
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND UPPER(" & strCampo & ") LIKE '" & UCase(strValor) & "'"
                                Else
                                    strCondicao = strCampo & " LIKE '" & strValor & "'"
                                End If
                            ElseIf InStr(1, UCase(.Controls(i).Name), "DTM") > 0 Then
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & gstrConvDtParaSql(strValor)
                                Else
                                    strCondicao = strCampo & " = " & gstrConvDtParaSql(strValor)
                                End If
                            Else
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND UPPER(" & strCampo & ") LIKE '" & UCase(strValor) & "%'"
                                Else
                                    strCondicao = "UPPER(" & strCampo & ") LIKE '" & UCase(strValor) & "%'"
                                End If
                            End If
                        End If
                    'If TypeOf .Controls(i) Is TextBox Then
                    ElseIf TypeOf .Controls(i) Is OptionButton Then
                        strValor = .Controls(i).Index
                        strCampo = Trim(.Controls(i).Name)
                        'Só será filtrado por este campo no modulo de Orcamentario
                        If strCampo = "optbytTipo" And UCase(App.ProductName) <> "ORCAMENTARIO" Then
                            GoTo ProximoCampo
                        Else
                            If InStr(1, strCampo, "_") > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            
                            If strCondicao <> "" Then
                                strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                            Else
                                strCondicao = strCampo & " = " & strValor
                            End If
                        End If
                    'ElseIf TypeOf .Controls(i) Is OptionButton Then
                    ElseIf TypeOf .Controls(i) Is CheckBox Then
                        If .Controls(i).Value = 1 Then
                            strValor = .Controls(i).Value
                            
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, strCampo, "_") > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            
                            If strCondicao <> "" Then
                                strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                            Else
                                strCondicao = strCampo & " = " & strValor
                            End If
                        End If
                    'ElseIf TypeOf .Controls(i) Is CheckBox Then
                    ElseIf TypeOf .Controls(i) Is DataCombo Then
                        If (.Controls(i).Name = "dbcstrNome" And Trim(dbcstrNome.Text) <> "") Or (.Controls(i).Name = "dbcstrNomeFantasia" And Trim(dbcstrNomeFantasia.Text) <> "") Then
                            strValor = .Controls(i).Text
                            strCampo = Trim(.Controls(i).Name)
                            
                            strCampo = Mid(strCampo, 4, Len(strCampo))
                            If strCondicao <> "" Then
                                strCondicao = strCondicao & " AND UPPER(" & strCampo & ") " & "LIKE '" & UCase(strValor) & "%'"
                            Else
                                strCondicao = "UPPER(" & strCampo & ")" & "LIKE '" & UCase(strValor) & "%'"
                            End If
                        Else
                            If .Controls(i).MatchedWithList Then
                            
                                strValor = .Controls(i).BoundText
                                strCampo = Trim(.Controls(i).Name)
                                
                                If InStr(1, strCampo, "_") > 0 Then
                                    strCampo = Mid(strCampo, 5, Len(strCampo))
                                Else
                                    strCampo = Mid(strCampo, 4, Len(strCampo))
                                End If
                                
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                                Else
                                    strCondicao = strCampo & " = " & strValor
                                End If
                            End If
                        End If
                    
                    ElseIf TypeOf .Controls(i) Is ComboBox Then
                        If .Controls(i).ListIndex > -1 Then
                        
                            strValor = .Controls(i).ListIndex
                            strCampo = Trim(.Controls(i).Name)
                            
                            If InStr(1, strCampo, "_") > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            
                            If strCondicao <> "" Then
                                strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                            Else
                                strCondicao = strCampo & " = " & strValor
                            End If
                        End If
                    End If 'If TypeOf .Controls(i) Is TextBox Then
ProximoCampo:
                End If
            End If 'If Not (TypeOf .Controls(I) Is OptionButton) Or .Controls(I) = True Then
        End If 'If Not TypeOf .Controls(I) Is Label Then
    Next i
End With

   If Me.mskstrCNPJCPF <> "" Then
        If strCondicao <> "" Then
            strCondicao = strCondicao & " AND strCNPJCPF = '" & gstrValorSemMascara(Me.mskstrCNPJCPF) & "' "
        Else
            strCondicao = "strCNPJCPF = " & gstrValorSemMascara(Me.mskstrCNPJCPF)
        End If
   End If



strSql = ""
If strCondicao <> "" Then
    strSql = strSql & " SELECT DISTINCT CO.* ," & gstrCASEWHEN("blnInativo", "0,'Não',1,'Sim'") & " strInativo "
    strSql = strSql & "FROM " & gstrContribuinte & " CO, " & gstrItens & " IT, " & gstrModuloContribuinte & " MC "
    strSql = strSql & "WHERE IT.PKId = MC.intItem AND "
    strSql = strSql & "MC.intContribuinte = CO.PKId /*AND IT.PKId =" & gintModulo & "*/ AND " & strCondicao

    'strSQL = strSQL & " SELECT * FROM " & gstrContribuinte & " WHERE " & strCondicao
Else
    strSql = strSql & " SELECT DISTINCT CO.* ," & gstrCASEWHEN("blnInativo", "0,'Não',1,'Sim'") & " strInativo "
    strSql = strSql & "FROM " & gstrContribuinte & " CO, " & gstrItens & " IT, " & gstrModuloContribuinte & " MC "
    strSql = strSql & "WHERE IT.PKId = MC.intItem AND "
    strSql = strSql & "MC.intContribuinte = CO.PKId " 'AND IT.PKId =" & gintModulo
    'strSQL = strSQL & " SELECT * FROM " & gstrContribuinte
End If

Select Case bytOrdenacao
    Case Is = 1: strSql = strSql & " ORDER BY CO.PKId"
    Case Is = 2: strSql = strSql & " ORDER BY CO.strNome"
    Case Is = 3: strSql = strSql & " ORDER BY CO.strCNPJCPF"
    Case Is = 3: strSql = strSql & " ORDER BY CO.strNomeFantasia"
    
End Select

LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strSql

End Sub

Private Sub CarregaSocios()
    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 26/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim varAux             As Variant
    Dim strSql             As String
    Dim adoResultado       As ADODB.Recordset
    Dim lngPKIdEconomico   As Long
    On Error GoTo Err_Handle
    
    strSql = "SELECT PKId FROM " & gstrEconomico & "  WHERE intContribuinte = " & txtPKId
    
    Set gobjBanco = New clsBanco
    gobjBanco.CriaADO strSql, 5, adoResultado
    With adoResultado
        If .EOF Then
            'Removido ref. Pendência COMP38 PEN_353
            'ExibeMensagem "Este contribuinte não foi cadastrado no Cadastro Econômico."
            tab_3DDadosGerais.TabEnabled(5) = False
            Exit Sub
        Else
            lngPKIdEconomico = !Pkid
        End If
     End With
        
    Set e = New XArrayDB
    e.Clear
'SE.intSocio
    strSql = ""
    strSql = strSql & "SELECT CO.strNome, CO.strCNPJCPF, "
'    strSql = strSql & "ISNULL(SE.intNumeroDeCotas, 0) Cotas "
'    strSql = strSql & strISNULL & "(SE.intNumeroDeCotas, 0) Cotas "
    strSql = strSql & gstrISNULL("SE.intNumeroDeCotas", "0") & " Cotas "
    strSql = strSql & "FROM " & gstrSocioEconomico & " SE, "
    strSql = strSql & gstrSocio & " SO, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & "WHERE SE.intSocio = SO.PKId "
    strSql = strSql & "AND SO.intContribuinte = CO.PKID "
    strSql = strSql & "AND SE.intCodEconomico = " & lngPKIdEconomico
    
    Set gobjBanco = New clsBanco
    gobjBanco.CriaADO strSql, 5, adoResultado
    With adoResultado
        If Not .EOF Then
            e.ReDim 0, .RecordCount - 1, 0, 2
            txt_TotalDeCotas = 0
            Do While Not .EOF
                varAux = !STRNOME
                e(.AbsolutePosition - 1, 0) = varAux
                
                varAux = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!StrCnpjCpf))
                e(.AbsolutePosition - 1, 1) = varAux

                varAux = !cotas
                e(.AbsolutePosition - 1, 2) = varAux
                txt_TotalDeCotas = txt_TotalDeCotas + Val(varAux)
                
                .MoveNext
            Loop
        Else
            e.ReDim 0, 0, 0, 2
            e(0, 0) = ""
            e(0, 1) = ""
            e(0, 2) = ""
        End If
    End With
    Set tdb_Socios.Array = e
    tdb_Socios.ReBind
    tdb_Socios.Refresh
    
    
    Set e = New XArrayDB
    e.Clear

    
Exit Sub
Err_Handle:
    ExibeDetalheErro ""
End Sub


Private Sub PreencheListaModulos(Optional intCodContribuinte As Long = -1)

'******************************************************************************************
' Data: 06/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim intFor             As Integer
Dim strSql             As String
Dim adoResultado       As ADODB.Recordset

    Set gobjBanco = New clsBanco
       
    If intCodContribuinte = -1 Then
        lvw_Aplicacoes.ListItems.Clear
        
'       strSql = "SELECT PKId, strItem FROM " & gstrItens & " WHERE Len(strCodItem) = 1"
        strSql = "SELECT PKId, strItem, strCodItem FROM " & gstrItens & " WHERE " & strLen & "(strCodItem) = 1"
        
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    Set oList = lvw_Aplicacoes.ListItems.Add(, Trim(.Fields("strCodItem")), Trim(.Fields("strItem")))
                    oList.Tag = .Fields("PKId")
                    .MoveNext
                Loop
            End With
        End If
    Else
        
        strSql = "SELECT intContribuinte, intItem FROM " & gstrModuloContribuinte & " WHERE intContribuinte=" & intCodContribuinte
        
        For intFor = 1 To lvw_Aplicacoes.ListItems.Count
            lvw_Aplicacoes.ListItems(intFor).Checked = False
        Next
        
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            
            While Not adoResultado.EOF
                For intFor = 1 To lvw_Aplicacoes.ListItems.Count
                    If lvw_Aplicacoes.ListItems(intFor).Tag = adoResultado("intItem") Then
                        lvw_Aplicacoes.ListItems(intFor).Checked = True
                    End If
                Next
                adoResultado.MoveNext
            Wend
        End If
    End If
    
    If lvw_Aplicacoes.ListItems.Count <> 0 Then
        lvw_Aplicacoes.SelectedItem.Selected = False
    End If
    
End Sub

Private Sub DeletaModulosUsuarios(intCodContribuinte As Long)
    Dim strSql As String

    strSql = ""
    strSql = strSql & "Delete From " & gstrModuloContribuinte & " "
    strSql = strSql & "Where intContribuinte = " & intCodContribuinte
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
    
End Sub

Function blnGravaAplicacoes(intCodContribuinte As Long, blnAlterando As Boolean) As Boolean
Dim adoResultado As New ADODB.Recordset
Dim strSql As String
Dim intI   As Integer
Dim bytNatureza As Byte

    DeletaModulosUsuarios intCodContribuinte
    
    strSql = "SELECT bytNaturezaJuridica FROM " & gstrContribuinte & " WHERE Pkid = " & intCodContribuinte
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            bytNatureza = adoResultado("bytNaturezaJuridica").Value
        End If
    End If
    
    If Not blnAlterando Then SelecionaModuloAtual bytNatureza
    
    With lvw_Aplicacoes
        For intI = 1 To .ListItems.Count
            If .ListItems(intI).Checked Then
                strSql = ""
                strSql = strSql & "Insert Into " & gstrModuloContribuinte & " "
                strSql = strSql & "(intContribuinte, intItem, dtmdtAtualizacao, lngCodUsr "
                strSql = strSql & ") Values ("
                strSql = strSql & intCodContribuinte & ", "
                strSql = strSql & .ListItems(intI).Tag & ", "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSql = strSql & glngCodUsr
                strSql = strSql & ")"
                Set gobjBanco = New clsBanco
                gobjBanco.Execute strSql
            End If
        Next
    End With
    
    blnGravaAplicacoes = True
    
End Function

Private Sub CopiaLogradouro(PkIDLogradouro As Integer, btpGuia As Byte)
Dim adoRec As ADODB.Recordset
Dim strSql As String

    On Error Resume Next

    Set gobjBanco = New clsBanco
    
    strSql = ""
    strSql = strSql & "SELECT A.PkID, A.strCodigo, A.strDescricao Logradouro, "
    strSql = strSql & "B.PkID IDTipo, B.strSigla Tipo,"
    strSql = strSql & "C.PkID IDTitulo, C.strDescricao Titulo"
    
    strSql = strSql & " FROM " & gstrLogradouro & " A,"
    strSql = strSql & gstrTipoLogradouro & " B, "
    strSql = strSql & gstrTituloLogradouro & " C "

    strSql = strSql & "WHERE B.PkID " & strOUTJOracle & "=" & strOUTJSQLServer & " A.intTipoLogradouro AND "
    strSql = strSql & "C.PkID " & strOUTJOracle & "=" & strOUTJSQLServer & " A.intTituloLogradouro AND "
    strSql = strSql & "A.PkID = " & PkIDLogradouro
    strSql = strSql & " AND A.Dtmdtexclusao is null "
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
            
            If btpGuia = 1 Then
                If Not (.BOF And .EOF) Then
            
                    dbcintTipoLogradouro.Text = gstrENulo(!Tipo)
                    dbcintTipoLogradouro.SetFocus: dbcintTipoLogradouro_Click 0
                    dbcintTipoLogradouro.BoundText = gstrENulo(!IDTipo)
                    
                    dbcintTituloLogradouro.Text = gstrENulo(!Titulo)
                    dbcintTituloLogradouro.SetFocus: dbcintTituloLogradouro_Click 0
                    dbcintTituloLogradouro.BoundText = gstrENulo(!IDTitulo)
                    
                    'Preenche dados do logradouro
                    txtintCodigoLogradouro = gstrENulo(!strCodigo)
                    dbcstrLogradouroC = gstrENulo(!Logradouro)

                Else
                    txtintCodigoLogradouro = Space$(0)
                    dbcstrLogradouroC = Space$(0)
                End If
            
            Else
                If Not (.BOF And .EOF) Then
                    dbcintTipoLogradouroD.Text = gstrENulo(!Tipo)
                    dbcintTipoLogradouroD.SetFocus: dbcintTipoLogradouroD_Click 0
                    dbcintTipoLogradouroD.BoundText = gstrENulo(!IDTipo)
                        
                    dbcintTituloLogradouroD.Text = gstrENulo(!Titulo)
                    dbcintTituloLogradouroD.SetFocus: dbcintTituloLogradouroD_Click 0
                    dbcintTituloLogradouroD.BoundText = gstrENulo(!IDTitulo)

                    'Preenche dados do logradouro
                    txtintCodigoLogradouroD = gstrENulo(!strCodigo)
                    txtstrLogradouroD = gstrENulo(!Logradouro)

                Else
                    txtintCodigoLogradouroD = Space$(0)
                    txtstrLogradouroD = Space$(0)
                End If
            End If
            
        End With
        
    End If
    
    Set gobjBanco = Nothing

End Sub

Private Sub SelecionaModuloAtual(bytNatureza As Byte)

Dim intFor              As Integer
Dim intForMod           As Integer
Dim strMods()           As String * 1

    Select Case UCase(App.ProductName)
        Case "TRIBUTARIO"
            ReDim Preserve strMods(3)
            strMods(0) = "J"
            strMods(1) = "M"
            strMods(2) = "H"
        Case "ORCAMENTARIO"
            If bytNatureza < 2 Then
                ReDim Preserve strMods(4)
                strMods(0) = "A"
                strMods(1) = "F"
                strMods(2) = "E"
                strMods(3) = "G"
            Else
                ReDim Preserve strMods(1)
                strMods(0) = "F"
            End If
        Case "FROTA"
            ReDim Preserve strMods(1)
            strMods(0) = "B"
        Case "RH"
            ReDim Preserve strMods(1)
            strMods(0) = "I"
        Case "LEGISLACAO"
            ReDim Preserve strMods(1)
            strMods(0) = "C"
        Case "OUVIDORIA"
            ReDim Preserve strMods(3)
            strMods(0) = "M"
            strMods(1) = "H"
            strMods(2) = "J"
        Case "COMPRAS"
            ReDim Preserve strMods(4)
            strMods(0) = "A"
            strMods(1) = "F"
            strMods(2) = "E"
            strMods(3) = "G"
        Case "PATRIMONIO"
            ReDim Preserve strMods(4)
            strMods(0) = "A"
            strMods(1) = "F"
            strMods(2) = "E"
            strMods(3) = "G"
        Case "MATERIAL"
            ReDim Preserve strMods(4)
            strMods(0) = "A"
            strMods(1) = "F"
            strMods(2) = "E"
            strMods(3) = "G"
        Case "PROTOCOLO"
            ReDim Preserve strMods(3)
            strMods(0) = "M"
            strMods(1) = "H"
            strMods(2) = "J"
        Case "SEGURANCA"
            ReDim Preserve strMods(1)
            strMods(0) = "L"
        Case "MENOR"
            ReDim Preserve strMods(1)
            strMods(0) = "D"
        Case "GERENCIAL"
            ReDim Preserve strMods(1)
            strMods(0) = "N"
    End Select

    With lvw_Aplicacoes
        For intFor = 1 To .ListItems.Count
            For intForMod = 0 To UBound(strMods) - 1
                If .ListItems(intFor).Key = strMods(intForMod) Then
                    .ListItems(intFor).Checked = True
                    Exit For
                Else
                    'Caso sejam outras naturezas, só vamos habilitar para orcamentario
                    If bytNatureza > 1 Then
                        .ListItems(intFor).Checked = False
                    End If
                End If
            Next
        Next
    End With
    
End Sub

Private Function blnUsuarioOutroModulo() As Boolean
    Dim adoResultado    As ADODB.Recordset
    Dim strSql          As String
    
    
    
    'strSql = "SELECT DISTINCT CO.PKId FROM " & gstrContribuinte & " CO, "
    'strSql = strSql & gstrModuloContribuinte & " MC, "
    'strSql = strSql & gstrItens & " IT"
    
    strSql = "SELECT DISTINCT CO.PKId "
    strSql = strSql & "FROM " & gstrContribuinte & " CO LEFT OUTER JOIN "
    strSql = strSql & gstrModuloContribuinte & " MC ON CO.PKId = MC.intContribuinte RIGHT OUTER JOIN "
    strSql = strSql & gstrItens & "  IT ON MC.intItem = IT.PKId "
    
    strSql = strSql & " WHERE "
    'strSql = strSql & "CO.PKId " & strOUTJSQLServer & "= MC.intContribuinte " & strOUTJOracle & " AND "
    'strSql = strSql & "IT.PKId " & strOUTJSQLServer & "= MC.intItem " & strOUTJOracle & " AND "
    strSql = strSql & " UPPER(CO.strNome) ='" & UCase(dbcstrNome.Text) & "' AND "
    
    If mskstrCNPJCPF.ClipText <> "" Then
        strSql = strSql & " strCNPJCPF ='" & mskstrCNPJCPF.ClipText & "' AND "
    Else
        If dbcintLogradouro.BoundText <> "" Then
            strSql = strSql & "CO.intLogradouro =" & dbcintLogradouro.BoundText & " AND "
            strSql = strSql & "CO.intNumero =" & Val(txtintNumero.Text) & " AND "
        Else
            strSql = strSql & "CO.strLogradouroC='" & dbcstrLogradouroC.Text & "' AND "
            strSql = strSql & "CO.intNumeroC=" & Val(txtintNumeroC.Text) & " AND "
        End If
    End If
    
    strSql = strSql & " MC.intItem <> " & gintModulo
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF Then
            blnUsuarioOutroModulo = False
        Else
            intUsuario = adoResultado("PKId")
            
            blnUsuarioOutroModulo = True
        End If
        
    End If
    
    Set gobjBanco = Nothing
    
End Function

Private Sub CarregaMunicipio()
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "SELECT MU.PKId AS intMunicipio, MU.strDescricao, UF.PKId AS intUF, UF.strSigla "
    strSql = strSql & "FROM "
    strSql = strSql & gstrCidade & " MU, "
    strSql = strSql & gstrUF & " UF "
    strSql = strSql & "WHERE MU.intUF = UF.PKId AND "
    strSql = strSql & "MU.PKId =" & gintMunicipioEmpresa
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                txt_strMunicipio.Text = !strDescricao
                txt_strMunicipio.Tag = !intMunicipio
                txt_strUF.Text = !strsigla
                txt_strUF.Tag = !intUf
            End If
        End With
    End If
    
    Set gobjBanco = Nothing
    
End Sub
Private Function VerificaEmpenho() As Boolean
   Dim strSql       As String
   Dim adoResultado As New ADODB.Recordset
   
   strSql = "SELECT intCredor FROM " & gstrEmpenho
   strSql = strSql & " WHERE intCredor = " & txtPKId
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         VerificaEmpenho = True
      End If
   End If

End Function

Private Function blnVerificaLogradouro(lngIntLogradouro) As Boolean
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset

    blnVerificaLogradouro = False
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "L.Pkid "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLogradouro & " L "
    strSql = strSql & "WHERE "
    strSql = strSql & "L.Pkid = " & dbcintLogradouro.BoundText & " And "
    strSql = strSql & "not L.Dtmdtexclusao is null "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            blnVerificaLogradouro = True
        End If
    End If

End Function

Private Sub PreencheNomeContribuinte()
    Dim strSql       As String
    Dim strWhere     As String
    Dim adoResultado As New ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & " PKId, "
    strSql = strSql & " strNome "
    strSql = strSql & "FROM "
    strSql = strSql & gstrContribuinte
    
    If dbcstrNome.Text <> "" Then
        strSql = strSql & " WHERE "
        strSql = strSql & "strNome like '" & dbcstrNome.Text & "%'"
    End If
    
    strSql = strSql & " ORDER BY "
    strSql = strSql & " strNome "
    
    LeDaTabelaParaObj "", dbcstrNome, strSql
    
End Sub

Private Sub PreencheNatureza()

    cbobytNaturezaJuridica.AddItem "Física"
    cbobytNaturezaJuridica.ItemData(cbobytNaturezaJuridica.NewIndex) = 0
    cbobytNaturezaJuridica.AddItem "Jurídica"
    cbobytNaturezaJuridica.ItemData(cbobytNaturezaJuridica.NewIndex) = 1
    If UCase(App.ProductName) = "ORCAMENTARIO" Then
        cbobytNaturezaJuridica.AddItem "Especial"
        cbobytNaturezaJuridica.ItemData(cbobytNaturezaJuridica.NewIndex) = 2
        cbobytNaturezaJuridica.AddItem "Resto a pagar (Exercícios anteriores)"
        cbobytNaturezaJuridica.ItemData(cbobytNaturezaJuridica.NewIndex) = 3
        cbobytNaturezaJuridica.AddItem "Sentenças judiciais (Natureza alimentar)"
        cbobytNaturezaJuridica.ItemData(cbobytNaturezaJuridica.NewIndex) = 4
        cbobytNaturezaJuridica.AddItem "Sentenças judiciais(Outros)"
        cbobytNaturezaJuridica.ItemData(cbobytNaturezaJuridica.NewIndex) = 5
        cbobytNaturezaJuridica.AddItem "Precatórios (Natureza alimentar)"
        cbobytNaturezaJuridica.ItemData(cbobytNaturezaJuridica.NewIndex) = 6
        cbobytNaturezaJuridica.AddItem "Precatórios (Outros)"
        cbobytNaturezaJuridica.ItemData(cbobytNaturezaJuridica.NewIndex) = 7
        cbobytNaturezaJuridica.AddItem "Outros"
        cbobytNaturezaJuridica.ItemData(cbobytNaturezaJuridica.NewIndex) = 8
    End If

End Sub

Private Sub LeDadosEmpresa()
Dim strSql      As String
Dim adoEmpresa  As ADODB.Recordset
    
    'Busca dados do Cadastro de Empresa
    strSql = ""
    strSql = strSql & "SELECT E.intCidade, E.intCep FROM "
    strSql = strSql & gstrEmpresa & " E, " & gstrCidade & " M "
    strSql = strSql & " WHERE M.PKId = E.intCidade"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoEmpresa) Then
        With adoEmpresa
            If .EOF = False Then
                intMunicipioEmpresa = !intCidade
                intCepEmpresa = !INTCEP
            End If
        End With
        adoEmpresa.Close
        Set adoEmpresa = Nothing
    End If
    
    Set gobjBanco = Nothing
    
End Sub

Private Function ContribuinteHabilitadoAplicacao(intContribuinte As Long) As Boolean
Dim strSql        As String
Dim adoResultado  As ADODB.Recordset

    ContribuinteHabilitadoAplicacao = False
    
    'Busca dados do Cadastro de Empresa
    strSql = ""
    strSql = strSql & "SELECT MC.intContribuinte FROM "
    strSql = strSql & gstrModuloContribuinte & " MC "
    strSql = strSql & " WHERE MC.intContribuinte = " & intContribuinte & " AND MC.intItem = " & gintModulo
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                ContribuinteHabilitadoAplicacao = True
            End If
        End With
        adoResultado.Close
        Set adoResultado = Nothing
    End If
    
    Set gobjBanco = Nothing

End Function
