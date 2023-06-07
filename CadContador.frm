VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadContador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Contadores"
   ClientHeight    =   7335
   ClientLeft      =   1065
   ClientTop       =   1905
   ClientWidth     =   9075
   HelpContextID   =   50
   Icon            =   "CadContador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9075
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   7140
      Left            =   150
      TabIndex        =   17
      Top             =   90
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   12594
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Contador"
      TabPicture(0)   =   "CadContador.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrCRC"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_PKId"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintContribuinte"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_CNPJCPF"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbointContribuinte"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdb_Contador"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrCRC"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPKId"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Contribuinte"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_Endereco"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_CNPJCPF"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Comunicações"
      TabPicture(1)   =   "CadContador.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_DescricaoConteudo"
      Tab(1).Control(1)=   "lbl_TipoComunicacao"
      Tab(1).Control(2)=   "img_Aux"
      Tab(1).Control(3)=   "ssp_TipoComunicacao"
      Tab(1).Control(4)=   "lvw_TipoComunicacao"
      Tab(1).Control(5)=   "txt_DescricaoConteudo"
      Tab(1).Control(6)=   "cmd_Up"
      Tab(1).Control(7)=   "cmd_Down"
      Tab(1).Control(8)=   "txt_Conteudo"
      Tab(1).ControlCount=   9
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
         TabIndex        =   13
         Top             =   1260
         Width           =   1600
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   " Endereço "
         Height          =   1575
         Left            =   180
         TabIndex        =   28
         Top             =   2010
         Width           =   8490
         Begin MSDataListLib.DataCombo dbcintUF 
            Height          =   315
            Left            =   7260
            TabIndex        =   10
            Top             =   1050
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.CommandButton cmd_TituloLogradouro 
            Height          =   315
            Left            =   3600
            Picture         =   "CadContador.frx":107A
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            Tag             =   "583"
            ToolTipText     =   "Ativa Cadastro de Logradouro"
            Top             =   270
            Width           =   360
         End
         Begin VB.CommandButton cmd_Logradouro 
            Height          =   315
            Left            =   1710
            Picture         =   "CadContador.frx":1198
            Style           =   1  'Graphical
            TabIndex        =   39
            TabStop         =   0   'False
            Tag             =   "582"
            ToolTipText     =   "Ativa Cadastro de Logradouro"
            Top             =   270
            Width           =   360
         End
         Begin VB.CommandButton cmd_Municipio 
            Height          =   330
            Left            =   5475
            Picture         =   "CadContador.frx":12B6
            Style           =   1  'Graphical
            TabIndex        =   38
            TabStop         =   0   'False
            Tag             =   "53"
            ToolTipText     =   "Ativa Cadastro de  Município"
            Top             =   1050
            Width           =   360
         End
         Begin VB.CommandButton cmd_UF 
            Height          =   315
            Left            =   8010
            Picture         =   "CadContador.frx":13D4
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            Tag             =   "1276"
            ToolTipText     =   "Ativa Cadastro de Unidade Federativa"
            Top             =   1050
            Width           =   390
         End
         Begin VB.TextBox txtstrBairro 
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
            Left            =   4260
            MaxLength       =   50
            TabIndex        =   7
            Top             =   690
            Width           =   2520
         End
         Begin VB.TextBox txtstrLogradouro 
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
            Left            =   4005
            MaxLength       =   50
            TabIndex        =   4
            Top             =   300
            Width           =   4335
         End
         Begin VB.TextBox txtintCEP 
            Height          =   285
            Left            =   7260
            MaxLength       =   9
            TabIndex        =   8
            Top             =   675
            Width           =   1080
         End
         Begin VB.TextBox txtstrComplemento 
            Height          =   285
            Left            =   2715
            MaxLength       =   20
            TabIndex        =   6
            Top             =   690
            Width           =   855
         End
         Begin VB.TextBox txtintNumero 
            Height          =   285
            Left            =   1050
            MaxLength       =   8
            TabIndex        =   5
            Top             =   690
            Width           =   855
         End
         Begin MSDataListLib.DataCombo dbcintTipoLogradouro 
            Height          =   315
            Left            =   1050
            TabIndex        =   2
            Top             =   285
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintMunicipio 
            Height          =   315
            Left            =   1050
            TabIndex        =   9
            Top             =   1065
            Width           =   4410
            _ExtentX        =   7779
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTituloLogradouro 
            Height          =   315
            Left            =   2100
            TabIndex        =   3
            Top             =   285
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label lblintNumero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   780
            TabIndex        =   37
            Top             =   780
            Width           =   180
         End
         Begin VB.Label lblintUF 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   6945
            TabIndex        =   36
            Top             =   1170
            Width           =   210
         End
         Begin VB.Label lblintCEP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   6885
            TabIndex        =   33
            Top             =   750
            Width           =   285
         End
         Begin VB.Label lblintComplemento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   2130
            TabIndex        =   32
            Top             =   780
            Width           =   480
         End
         Begin VB.Label lblintLogradouro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   150
            TabIndex        =   31
            Top             =   390
            Width           =   810
         End
         Begin VB.Label lblstrBairro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   3720
            TabIndex        =   30
            Top             =   780
            Width           =   405
         End
         Begin VB.Label lblintMunicipio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   255
            TabIndex        =   29
            Top             =   1185
            Width           =   705
         End
      End
      Begin VB.TextBox txt_Conteudo 
         Height          =   285
         Left            =   -73620
         MaxLength       =   50
         TabIndex        =   14
         Top             =   600
         Width           =   3945
      End
      Begin VB.CommandButton cmd_Down 
         Height          =   285
         Left            =   -66570
         Picture         =   "CadContador.frx":14F2
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Abaixo"
         Top             =   1725
         Width           =   300
      End
      Begin VB.CommandButton cmd_Up 
         Height          =   285
         Left            =   -66570
         Picture         =   "CadContador.frx":163C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Acima"
         Top             =   1425
         Width           =   300
      End
      Begin VB.TextBox txt_DescricaoConteudo 
         Height          =   285
         Left            =   -73620
         MaxLength       =   50
         TabIndex        =   15
         Top             =   945
         Width           =   3945
      End
      Begin VB.CommandButton cmd_Contribuinte 
         Height          =   315
         Left            =   6000
         Picture         =   "CadContador.frx":1786
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Ativa Cadastro de Contribuintes"
         Top             =   870
         Width           =   360
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   510
         Width           =   1170
      End
      Begin VB.TextBox txtstrCRC 
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
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1620
         Width           =   1600
      End
      Begin MSComctlLib.ListView lvw_TipoComunicacao 
         Height          =   2610
         Left            =   -74850
         TabIndex        =   16
         Top             =   1425
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   4604
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
      Begin Threed.SSPanel ssp_TipoComunicacao 
         Height          =   390
         Left            =   -69570
         TabIndex        =   24
         Top             =   855
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
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
         Begin MSComctlLib.Toolbar tlb_TipoComunicacao 
            Height          =   330
            Left            =   30
            TabIndex        =   27
            Top             =   30
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
      End
      Begin MSComctlLib.ImageList img_Aux 
         Left            =   -68160
         Top             =   690
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
               Picture         =   "CadContador.frx":18A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadContador.frx":1A04
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadContador.frx":1B60
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Contador 
         Height          =   3285
         Left            =   180
         TabIndex        =   11
         Top             =   3690
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   5794
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2434"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2355"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=11986"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=11906"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
      Begin MSDataListLib.DataCombo cbointContribuinte 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   870
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lbl_CNPJCPF 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ / CPF"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   1350
         Width           =   870
      End
      Begin VB.Label lbl_TipoComunicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   -74160
         TabIndex        =   26
         Top             =   690
         Width           =   315
      End
      Begin VB.Label lbl_DescricaoConteudo 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   -74565
         TabIndex        =   25
         Top             =   1035
         Width           =   720
      End
      Begin VB.Label lblintContribuinte 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   630
         TabIndex        =   21
         Top             =   990
         Width           =   420
      End
      Begin VB.Label lbl_PKId 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   555
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblstrCRC 
         AutoSize        =   -1  'True
         Caption         =   "CRC"
         Height          =   195
         Left            =   720
         TabIndex        =   18
         Top             =   1710
         Width           =   330
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
Attribute VB_Name = "frmCadContador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando             As Boolean
Dim mobjAux                   As Object
Dim oList                     As Object
Dim mblnSelecionou            As Boolean
Dim mblnPrimeiraVez           As Boolean
Dim mblnClickOk               As Boolean
Dim blnOrdenacaoAsc           As Boolean
Dim bytOrdenacao              As Byte


Private Sub cbointContribuinte_Click(Area As Integer)
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    DropDownDataCombo cbointContribuinte, Me, Area
    
    If Area = 2 Then
        If cbointContribuinte.BoundText = "" Then
            Exit Sub
        End If
        strSql = ""
        strSql = strSql & "Select strCNPJCPF "
        strSql = strSql & "From " & gstrContribuinte & " "
        strSql = strSql & "Where PKId = " & gstrItemData(cbointContribuinte)
    
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                With adoResultado
                    txt_CNPJCPF = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!StrCnpjCpf))
'                    cbo_Logradouro.ListIndex = !intLogradouro
'                    cbo_Bairro.ListIndex = !intBairro
'                    cbo_Municipio.ListIndex = !intMunicipio
'                    cbo_UF = !strUF
'                    txt_Numero = gstrVerificaCampoNulo(!intNumero)
'                    txt_Complemento = gstrVerificaCampoNulo(!strComplemento)
'                    txt_Cep = gstrCEPFormatado(gstrVerificaCampoNulo(!intCep))
                End With
            End If
        End If
    End If
End Sub

Private Sub cbointContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo cbointContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub cmd_Contribuinte_Click()
    ChamaFormCadastro frmCadContribuinte, cbointContribuinte
End Sub

Private Sub cmd_Logradouro_Click()
    ChamaFormCadastro frmCadTipoLogradouro, dbcintTipoLogradouro ', "PKId, strSigla"
End Sub

Private Sub cmd_Municipio_Click()
    ChamaFormCadastro frmCadCidade, dbcintMunicipio
End Sub

Private Sub cmd_TituloLogradouro_Click()
    ChamaFormCadastro frmCadTituloLogradouro, dbcintTituloLogradouro
End Sub



Private Sub cmd_UF_Click()
    ChamaFormCadastro frmCadUF, dbcintUF
End Sub

Private Sub dbcintMunicipio_Click(Area As Integer)
   DropDownDataCombo dbcintMunicipio, Me, Area
End Sub

Private Sub dbcintMunicipio_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintMunicipio, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoLogradouro_Click(Area As Integer)
   DropDownDataCombo dbcintTipoLogradouro, Me, Area
End Sub

Private Sub dbcintTipoLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTipoLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouro_Click(Area As Integer)
   DropDownDataCombo dbcintTituloLogradouro, Me, Area
End Sub

Private Sub dbcintTituloLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTituloLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUf_Click(Area As Integer)
   DropDownDataCombo dbcintUF, Me, Area
End Sub

Private Sub dbcintUf_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUF, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 626
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
    bytOrdenacao = 2: blnOrdenacaoAsc = True
    MontaColumnHeaders
    dbcintMunicipio.Tag = gstrQueryDataComboMunicipio & ";strDescricao"
    dbcintTituloLogradouro.Tag = gstrQueryDataComboTituloLogradouro & ";strDescricao"
    dbcintTipoLogradouro.Tag = gstrQueryDataComboTipoLogradouro & ";strSigla"
    dbcintUF.Tag = gstrQueryDataComboUF & ";strSigla"
    cbointContribuinte.Tag = strQueryDataComboContribuinte & ";strNome"
    'VerificaListaAutomatica gstrContador, tdb_Contador, strQueryContador
    'VerificaPermissoes Me  , Me.tlb_BarraFermta, Me.Tag
    PreencheMenuPopup
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub lvw_TipoComunicacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_TipoComunicacao
        lbl_TipoComunicacao = .SelectedItem.Text
        txt_Conteudo = .SelectedItem.SubItems(1)
        txt_DescricaoConteudo = .SelectedItem.SubItems(2)
    End With
End Sub

Private Sub lvw_TipoComunicacao_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnu_TipoComunicacao
    End If
End Sub

Private Sub tdb_Contador_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Contador) = 1 Then
        tdb_Contador_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Contador_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Contador
End Sub

Private Sub tdb_Contador_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Contador, ColIndex
End Sub

Private Sub tdb_Contador_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Contador_KeyPress(KeyAscii As Integer)
    Select Case tdb_Contador.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_Contador
        Case Else
            CaracterValido KeyAscii, "A", tdb_Contador
    End Select
End Sub

Private Sub tdb_Contador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Contador_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Contador
        If Not .EOF And Not .BOF And mblnClickOk = True Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrContador, Me
                cbointContribuinte_Click 2
                CarregaTipoComunicacao Val(txtPKId)
                 gCorLinhaSelecionada tdb_Contador
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
    Dim intCodContador As Integer
    
    strSql = ""
    
    If mblnAlterando Then
        intCodContador = tdb_Contador.Columns("PKID").Value
    Else
        intCodContador = IsNull(txtPKId)
    End If
    
    strSql = strQueryContador
    
    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If
    
    Select Case UCase(strModoOperacao)
        Case "NOVO"
            LimpaObjeto Me, mblnAlterando
            NovoContador
            
        Case "SALVAR"
            If blnDadosOk Then
                If ToolBarGeral(strModoOperacao, gstrContador, mblnAlterando, tdb_Contador, Me, mobjAux, strSql) Then
                    GravaTipoComunicacao intCodContador
                    NovoContador
                End If
                
            End If
            
        Case "DELETAR"
            If ToolBarGeral(strModoOperacao, gstrContador, mblnAlterando, tdb_Contador, Me, mobjAux, strSql) Then
                DeletaTipoComunicacao intCodContador
                NovoContador
            End If
        Case "IMPRIMIR"
            ImprimeRelatorio rptContador, strQueryRelatorio
        Case Else
            ToolBarGeral strModoOperacao, gstrContador, mblnAlterando, tdb_Contador, Me, mobjAux, strSql
            
    End Select
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
End Sub

Sub MontaColumnHeaders()
    With lvw_TipoComunicacao
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Tipo", 2000
        .ColumnHeaders.Add 2, , "Conteúdo", 3000
        .ColumnHeaders.Add 3, , "Descrição", 3000
    End With
End Sub

Private Function strQueryContador() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
    strSql = ""
    strSql = strSql & "Select C.PKId, C.PKId Codigo, CO.strNome "
    strSql = strSql & "From " & gstrContador & " C, " & gstrContribuinte & " CO "
    strSql = strSql & "Where C.intContribuinte = CO.PKId "
    
    Select Case bytOrdenacao
      Case Is = 1
         strSql = strSql & " ORDER BY C.PKId" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 2
         strSql = strSql & " ORDER BY CO.strNome" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQueryContador = strSql
End Function

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

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCEP
End Sub

Private Sub txt_CNPJCPF_GotFocus()
    MarcaCampo txt_CNPJCPF
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCEP
End Sub

Private Sub txtintCEP_LostFocus()
    txtintCEP = gstrCEPFormatado(txtintCEP)
    gblnCepValido txtintCEP, , dbcintMunicipio
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumero
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
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

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub txtstrBairro_GotFocus()
    MarcaCampo txtstrBairro
End Sub

Private Sub txtstrBairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrBairro
End Sub
 
Private Sub txtstrLogradouro_GotFocus()
    MarcaCampo txtstrLogradouro
End Sub

Private Sub txtstrLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLogradouro
End Sub

Private Sub txtintNumero_GotFocus()
    MarcaCampo txtintNumero
End Sub

Private Sub txtstrCRC_GotFocus()
    MarcaCampo txtstrCRC
End Sub

Private Sub txtstrCRC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCRC
End Sub

Private Sub txtPKId_GotFocus()
    MarcaCampo txtPKId
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

Sub GravaTipoComunicacao(intCodContador As Integer)
    Dim strSql As String
    Dim intI   As Integer
    
    DeletaTipoComunicacao intCodContador
    
    With lvw_TipoComunicacao
        For intI = 1 To .ListItems.Count
            strSql = ""
            strSql = strSql & "Insert Into " & gstrFormaDeComunicacaoContador & " "
            strSql = strSql & "(intContador, intTipoDeComunicacao, strConteudo, strDescricao, "
            strSql = strSql & "intSequencia) Values ("
            strSql = strSql & intCodContador & ", "
            strSql = strSql & .ListItems(intI).Tag & ", '"
            strSql = strSql & .ListItems(intI).SubItems(1) & "', '"
            strSql = strSql & .ListItems(intI).SubItems(2) & "', "
            strSql = strSql & intI & ")"
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSql
        Next
    End With
End Sub

Sub DeletaTipoComunicacao(intCodContador As Integer)
    Dim strSql As String
    strSql = ""
    strSql = strSql & "Delete From " & gstrFormaDeComunicacaoContador & " "
    strSql = strSql & "Where intContador = " & intCodContador
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
End Sub

Sub CarregaTipoComunicacao(intCodContador As Integer)

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
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
'    strSql = strSql & "Left Join " & gstrFormaDeComunicacaoContador & " FC "
    strSql = strSql & ", " & gstrFormaDeComunicacaoContador & " FC "
'    strSql = strSql & "On TP.PKId = FC.intTipoDeComunicacao "
    strSql = strSql & "Where FC.intContador = " & intCodContador & " "
    
    strSql = strSql & " And TP.PKId " & strOUTJOracle & strOUTJSQLServer & "= FC.intTipoDeComunicacao "
    
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

Sub NovoContador()
    lvw_TipoComunicacao.ListItems.Clear
    txtPKId = glngPegaProximaChave(gstrContador, "PKId")
    txt_Conteudo = ""
    txt_DescricaoConteudo = ""
    txt_CNPJCPF = ""
    dbcintUF.Text = ""
    txtintNumero = ""
    txtstrComplemento = ""
    txtintCEP = ""
    lbl_TipoComunicacao = "Tipo"
    tab_3dPasta.Tab = 0
End Sub

Private Function blnDadosOk() As Boolean
    blnDadosOk = True
End Function

Function strQueryRelatorio() As String
'RESPONSAVEL 'LEANDRO' 29/06/2004

Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT CT.PKId Codigo, CO.strNome Contribuinte, CT.strCRC, "
    strSql = strSql & "FC.strConteudo ContextoCom,TC.strDescricao TipoCom "
    strSql = strSql & "FROM " & gstrFormaDeComunicacaoContador & " FC,"
    strSql = strSql & gstrTipoDeComunicacao & " TC,"
    strSql = strSql & gstrContador & " CT,"
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE CO.PKID =" & strOUTJSQLServer & " CT.intContribuinte " & strOUTJOracle & " AND TC.PKID = FC.intTipodeComunicacao AND CT.PKId = FC.intContador "
strQueryRelatorio = strSql
End Function

Private Function strQueryDataComboContribuinte()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNome "
    strSql = strSql & "FROM " & gstrContribuinte & " "
    strSql = strSql & "ORDER BY strNome"
    strQueryDataComboContribuinte = strSql
End Function

