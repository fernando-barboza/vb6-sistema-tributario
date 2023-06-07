VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadAgenciaBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agências Bancárias"
   ClientHeight    =   5085
   ClientLeft      =   1815
   ClientTop       =   2025
   ClientWidth     =   8415
   HelpContextID   =   16
   Icon            =   "CadAgenciaBanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8415
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   3540
      TabIndex        =   31
      Top             =   -60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   4845
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   135
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8546
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   " Agência "
      TabPicture(0)   =   "CadAgenciaBanco.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintBanco"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdb_Agencia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frm_Endereco"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmd_Banco"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintBanco"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Comunicações"
      TabPicture(1)   =   "CadAgenciaBanco.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_TipoComunicacao"
      Tab(1).Control(1)=   "lbl_DescricaoConteudo"
      Tab(1).Control(2)=   "img_Aux"
      Tab(1).Control(3)=   "ssp_TipoComunicacao"
      Tab(1).Control(4)=   "lvw_TipoComunicacao"
      Tab(1).Control(5)=   "cmd_Down"
      Tab(1).Control(6)=   "cmd_Up"
      Tab(1).Control(7)=   "txt_DescricaoConteudo"
      Tab(1).Control(8)=   "txt_Conteudo"
      Tab(1).ControlCount=   9
      Begin VB.TextBox txt_Conteudo 
         Height          =   285
         Left            =   -73140
         MaxLength       =   50
         TabIndex        =   16
         Top             =   720
         Width           =   3945
      End
      Begin VB.TextBox txt_DescricaoConteudo 
         Height          =   285
         Left            =   -73140
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1020
         Width           =   3945
      End
      Begin VB.CommandButton cmd_Up 
         Height          =   285
         Left            =   -67470
         Picture         =   "CadAgenciaBanco.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Acima"
         Top             =   1425
         Width           =   300
      End
      Begin VB.CommandButton cmd_Down 
         Height          =   285
         Left            =   -67470
         Picture         =   "CadAgenciaBanco.frx":11C4
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Abaixo"
         Top             =   1725
         Width           =   300
      End
      Begin MSDataListLib.DataCombo dbcintBanco 
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Top             =   450
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.CommandButton cmd_Banco 
         Height          =   300
         Left            =   7530
         Picture         =   "CadAgenciaBanco.frx":130E
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Bancos"
         Top             =   465
         Width           =   405
      End
      Begin VB.Frame frm_Endereco 
         Caption         =   " Agência "
         Height          =   1845
         Left            =   90
         TabIndex        =   25
         Top             =   855
         Width           =   7935
         Begin VB.TextBox txtstrLogradouro 
            Height          =   285
            Left            =   3870
            MaxLength       =   60
            TabIndex        =   7
            Top             =   645
            Width           =   3945
         End
         Begin VB.CommandButton cmd_TituloLogradouro 
            Height          =   330
            Left            =   3510
            Picture         =   "CadAgenciaBanco.frx":142C
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Tag             =   "583"
            ToolTipText     =   "Ativa Cadastro Títulos de Logradouro"
            Top             =   600
            Width           =   360
         End
         Begin VB.CommandButton cmd_TipoLogradouro 
            Height          =   330
            Left            =   1710
            Picture         =   "CadAgenciaBanco.frx":154A
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Tag             =   "582"
            ToolTipText     =   "Ativa Cadastro Tipos de Logradouro"
            Top             =   600
            Width           =   360
         End
         Begin VB.CommandButton cmd_UF 
            Height          =   330
            Left            =   7455
            Picture         =   "CadAgenciaBanco.frx":1668
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "1276"
            ToolTipText     =   "Ativa Cadastro de Unidades Federativas"
            Top             =   1380
            Width           =   360
         End
         Begin VB.TextBox txtstrComplemento 
            Height          =   285
            Left            =   3210
            MaxLength       =   10
            TabIndex        =   9
            Top             =   1020
            Width           =   1125
         End
         Begin VB.TextBox txtstrBairro 
            Height          =   285
            Left            =   5040
            MaxLength       =   40
            TabIndex        =   10
            Top             =   1020
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Cidade 
            Height          =   330
            Left            =   6030
            Picture         =   "CadAgenciaBanco.frx":1786
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Tag             =   "53"
            ToolTipText     =   "Ativa Cadastro de Cidades"
            Top             =   1395
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintCidade 
            Height          =   315
            Left            =   2790
            TabIndex        =   12
            Top             =   1410
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.TextBox txtstrDescricao 
            Height          =   285
            Left            =   2145
            MaxLength       =   40
            TabIndex        =   3
            Top             =   240
            Width           =   5670
         End
         Begin VB.TextBox txtstrAgencia 
            Height          =   285
            Left            =   960
            MaxLength       =   4
            TabIndex        =   2
            Top             =   270
            Width           =   585
         End
         Begin VB.TextBox txtintNumero 
            Height          =   285
            Left            =   960
            MaxLength       =   10
            TabIndex        =   8
            Top             =   1035
            Width           =   1095
         End
         Begin VB.TextBox txtintCep 
            Height          =   285
            Left            =   960
            MaxLength       =   12
            TabIndex        =   11
            Top             =   1425
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dbcintUF 
            Height          =   315
            Left            =   6780
            TabIndex        =   14
            Top             =   1395
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTipoLogradouro 
            Height          =   315
            Left            =   960
            TabIndex        =   4
            Top             =   615
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTituloLogradouro 
            Height          =   315
            Left            =   2070
            TabIndex        =   5
            Top             =   630
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label lblintUF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   6450
            TabIndex        =   39
            Top             =   1515
            Width           =   210
         End
         Begin VB.Label lblintNumero 
            AutoSize        =   -1  'True
            Caption         =   "N°"
            Height          =   195
            Left            =   675
            TabIndex        =   38
            Top             =   1080
            Width           =   180
         End
         Begin VB.Label lblstrComplemento 
            AutoSize        =   -1  'True
            Caption         =   "Complemento"
            Height          =   195
            Left            =   2160
            TabIndex        =   37
            Top             =   1110
            Width           =   960
         End
         Begin VB.Label lblstrDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   1620
            TabIndex        =   33
            Top             =   330
            Width           =   420
         End
         Begin VB.Label lblstrAgencia 
            AutoSize        =   -1  'True
            Caption         =   "Agência"
            Height          =   195
            Left            =   300
            TabIndex        =   32
            Top             =   330
            Width           =   585
         End
         Begin VB.Label lblintTipoLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   75
            TabIndex        =   29
            Top             =   705
            Width           =   810
         End
         Begin VB.Label lblstrBairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   4530
            TabIndex        =   28
            Top             =   1095
            Width           =   405
         End
         Begin VB.Label lblintCidade 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   195
            Left            =   2160
            TabIndex        =   27
            Top             =   1500
            Width           =   495
         End
         Begin VB.Label lblintCEP 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   585
            TabIndex        =   26
            Top             =   1515
            Width           =   285
         End
      End
      Begin MSComctlLib.ListView lvw_TipoComunicacao 
         Height          =   3210
         Left            =   -74760
         TabIndex        =   22
         Top             =   1410
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   5662
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
         Left            =   -69060
         TabIndex        =   34
         Top             =   915
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
            TabIndex        =   18
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
         Left            =   -67470
         Top             =   2070
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
               Picture         =   "CadAgenciaBanco.frx":18A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadAgenciaBanco.frx":1A04
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CadAgenciaBanco.frx":1B60
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Agencia 
         Height          =   1875
         Left            =   120
         TabIndex        =   21
         Top             =   2820
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3307
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "pkid"
         Columns(0).DataField=   "pkid"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Banco"
         Columns(1).DataField=   "strBanco"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3175"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3096"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=5662"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5583"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=7726"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=7646"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_DescricaoConteudo 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   -74175
         TabIndex        =   36
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label lbl_TipoComunicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   -73770
         TabIndex        =   35
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lblintBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   510
         TabIndex        =   30
         Top             =   540
         Width           =   465
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
Attribute VB_Name = "frmCadAgenciaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando       As Boolean
    Dim mlngUltimo          As Long
    Dim mblnGuardaUltimo    As Boolean
    Dim mobjAux             As Object
    Dim oList               As Object
    Dim mblnselecionou      As Boolean
    Dim mblnPrimeiraVez     As Boolean
    Dim intMaxPKId          As Integer
    Dim strDescricaoAtual  As String
    Dim bytOrdenacao                As Byte
    Dim blnOrdenacaoAsc             As Boolean

Private Function strQuery() As String
    
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT AG.PKId, AG.strDescricao strDescricao, BC.strDescricao strBanco FROM "
    strSql = strSql & gstrBanco & " BC, "
    strSql = strSql & gstrAgencia & " AG "
    strSql = strSql & "WHERE BC.PKId = AG.intBanco "
    
'    If dbcintBanco.MatchedWithList Then
'        strSql = strSql & "AND BC.PKId = " & dbcintBanco.BoundText & " "
'    End If
    
    strSql = strSql & "ORDER BY AG.strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    strQuery = strSql
End Function

Private Sub cmd_Banco_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub cmd_Banco_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", cmd_Banco
End Sub

Private Sub cmd_TipoLogradouro_Click()
    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouro
End Sub

Private Sub cmd_TituloLogradouro_Click()
    CarregaForm frmCadTituloLogradouro, dbcintTituloLogradouro
End Sub

Private Sub dbcintBanco_Click(Area As Integer)
    DropDownDataCombo dbcintBanco, Me, Area
'    If Area = 2 And dbcintBanco.MatchedWithList Then
'        LeDaTabelaParaObj gstrAgencia, tdb_Agencia, strQuery
'    End If
End Sub

Private Sub dbcintBanco_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintBanco, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBanco_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintBanco
End Sub

Private Sub dbcintCidade_Click(Area As Integer)
Dim adoUF As ADODB.Recordset
Dim strSql As String
    
    DropDownDataCombo dbcintCidade, Me, Area
    
    If Trim(dbcintCidade.Text) <> "" Then
       strSql = ""
       strSql = strSql & "SELECT UF.pkID, UF.strSigla "
       strSql = strSql & "FROM " & gstrUF & " UF, "
       strSql = strSql & gstrCidade & " MU "
       strSql = strSql & "WHERE MU.pkID = " & dbcintCidade.BoundText & " AND "
       strSql = strSql & "UF.pkID = MU.intUF "
       Set gobjBanco = New clsBanco
       If gobjBanco.CriaADO(strSql, 5, adoUF) Then
          If Not adoUF.EOF Then
             dbcintUF.Text = adoUF(1)
             'DropDownDataCombo dbcintUF, Me, Area
             PreencherListaDeOpcoes dbcintUF
             dbcintUF.Text = adoUF(1)
          End If
       End If
    End If

End Sub

Private Sub dbcintCidade_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintCidade, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCidade_KeyPress(KeyAscii As Integer)
  CaracterValido KeyAscii, "A", dbcintCidade
End Sub

Private Sub dbcintTipoLogradouro_Click(Area As Integer)
    DropDownDataCombo dbcintTipoLogradouro, Me, Area
End Sub

Private Sub dbcintTipoLogradouro_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintTipoLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTipoLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoLogradouro_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintTipoLogradouro
End Sub

Private Sub dbcintTituloLogradouro_Click(Area As Integer)
    DropDownDataCombo dbcintTituloLogradouro, Me, Area
End Sub

Private Sub dbcintTituloLogradouro_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintTituloLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTituloLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouro_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintTituloLogradouro
End Sub

Private Sub dbcintUf_Click(Area As Integer)
    DropDownDataCombo dbcintUF, Me, Area
End Sub

Private Sub dbcintUF_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintUf_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUF, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "U", dbcintUF
End Sub

Private Sub cmd_Banco_Click()
    CarregaForm frmCadBanco, dbcintBanco
End Sub

Private Sub cmd_Cidade_Click()
    CarregaForm frmCadCidade, dbcintCidade
End Sub

Private Sub cmd_UF_Click()
    CarregaForm frmCadUF, dbcintUF
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 591
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
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Sub MontaColumnHeaders()
    With lvw_TipoComunicacao
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Tipo", 2000
        .ColumnHeaders.Add 2, , "Conteúdo", 3000
        .ColumnHeaders.Add 3, , "Descrição", 3000
    End With
End Sub

Function PegaMaxPKId()
Dim strSql As String
Dim adoResultado As ADODB.Recordset
        strSql = ""
        strSql = strSql & "SELECT MAX(PKId) as PKId "
        strSql = strSql & " FROM " & gstrAgencia
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
             intMaxPKId = adoResultado!Pkid
        End If
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSql As String

If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar

mblnGuardaUltimo = True
Dim intPkid As Integer
    
    If mblnAlterando Then
        intPkid = Val(tdb_Agencia.Columns(0))
    End If
    
    strSql = strQuery
    
    Select Case UCase(strModoOperacao)
        Case "NOVO"
            LimpaObjeto Me, mblnAlterando
            NovoTipo
            'tdb_Agencia.DataSource = Nothing
            mblnAlterando = False
            mblnGuardaUltimo = False
            
        Case "SALVAR"
            If blnDadosOk Then
                mblnPrimeiraVez = False
                '''
                If ToolBarGeral(strModoOperacao, gstrAgencia, mblnAlterando, tdb_Agencia, Me, mobjAux, strSql, strSql, rptAgenciaBanco, strQueryRelatorio) Then
                    If intPkid <> 0 Then
                        GravaTipoComunicacao intPkid
                        NovoTipo
                    Else
                        PegaMaxPKId
                        GravaTipoComunicacao intMaxPKId
                        NovoTipo
                    End If
                    mblnAlterando = False
                    mblnGuardaUltimo = False
                End If
            End If
            
        Case "DELETAR"
            mblnPrimeiraVez = False
            '''
            If ToolBarGeral(strModoOperacao, gstrAgencia, mblnAlterando, tdb_Agencia, Me, mobjAux, strSql, strSql, rptAgenciaBanco, strQueryRelatorio) Then
                DeletaTipoComunicacao intPkid
                NovoTipo
                mblnAlterando = False
                mblnGuardaUltimo = False
            End If
            
        Case "APLICAR"
            AplicarGeral Me, mobjAux, tdb_Agencia, gstrAgencia ', , strQueryAplicar
            
'        Case "GRADE"
'            MudaGradeListView Me
            
        Case "IMPRIMIR"
            ImprimeRelatorio rptAgenciaBanco, strQueryRelatorio
            
        Case "REFRESH"
            LeDaTabelaParaObj gstrAgencia, tdb_Agencia, strQuery
        
        Case gstrLocalizar, gstrPreencherLista
            ToolBarGeral strModoOperacao, gstrAgencia, mblnAlterando, tdb_Agencia, Me, mobjAux, strQuery
            
        Case "FECHAR"
            Unload Me
            
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnselecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub lvw_TipoComunicacao_GotFocus()
tab_3dPasta.Tab = 1
End Sub

Private Sub lvw_TipoComunicacao_KeyPress(KeyAscii As Integer)
  CaracterValido KeyAscii, " ", lvw_TipoComunicacao
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", tab_3dPasta
End Sub

Private Sub tdb_Agencia_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Agencia) = 1 Then
        tdb_Agencia_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Agencia_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Agencia_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Agencia
End Sub

Private Sub tdb_Agencia_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub tdb_Agencia_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Agencia, ColIndex
End Sub

Private Sub tdb_Agencia_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", tdb_Agencia
End Sub

Private Sub tdb_Agencia_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Agencia
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKID.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrAgencia, Me
                CarregaTipoComunicacao .Columns("PKID").Value
                 gCorLinhaSelecionada tdb_Agencia
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                strDescricaoAtual = tdb_Agencia.Columns("strDescricao").Value
                mblnselecionou = True
                mblnAlterando = True
            End If
        End If
    End With
End Sub

Private Sub txtstrAgencia_GotFocus()
    MarcaCampo txtstrAgencia
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrAgencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrAgencia
End Sub

Private Sub Form_Load()
    
    bytOrdenacao = 1: blnOrdenacaoAsc = True
    
    dbcintBanco.Tag = gstrQueryDataComboBanco & ";strDescricao"
    dbcintTipoLogradouro.Tag = gstrQueryDataComboTipoLogradouro & ";strSigla"
    dbcintTituloLogradouro.Tag = gstrQueryDataComboTituloLogradouro & ";strDescricao"
    dbcintCidade.Tag = gstrQueryDataComboMunicipio & ";strDescricao"
    dbcintUF.Tag = gstrQueryDataComboUF & ";strSigla"
    VerificaObjParaAplicar mobjAux
    MontaColumnHeaders
    PreencheMenuPopup
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCEP
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCEP
End Sub

Private Sub txtintCEP_LostFocus()
    txtintCEP = gstrCEPFormatado(txtintCEP)
    gblnCepValido txtintCEP, , dbcintCidade
    CepLogradouro txtintCEP, txtstrLogradouro, txtstrBairro, dbcintCidade, dbcintUF, dbcintTipoLogradouro, dbcintTituloLogradouro, False, False, False, True, True, True, True, False, False, ""
End Sub

Private Sub txtstrBairro_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrBairro_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txtstrBairro
End Sub

Private Sub txtstrComplemento_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub txtstrDescricao_Change()
        MarcaCampo txtintNumero
End Sub

Private Sub txtstrDescricao_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrLogradouro_GotFocus()
tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrLogradouro_KeyPress(KeyAscii As Integer)
      CaracterValido KeyAscii, "A", txtstrLogradouro
  
End Sub

Private Sub txtintNumero_GotFocus()
        MarcaCampo txtintNumero
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumero
End Sub

'----------------COMUNICAÇÃO-------------'

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
        'MoveItemNoListView lvw_TipoComunicacao, True
    End If
End Sub

Private Sub cmd_Up_Click()
    If lvw_TipoComunicacao.ListItems.Count <> 0 Then
        'MoveItemNoListView lvw_TipoComunicacao, False
    End If
End Sub

Sub GravaTipoComunicacao(txtPKID As Integer)
    Dim strSql As String
    Dim intI   As Integer
    
    DeletaTipoComunicacao Val(txtPKID)
    
    With lvw_TipoComunicacao
        For intI = 1 To .ListItems.Count
            strSql = ""
            strSql = strSql & "Insert Into " & gstrFormaDeComunicacaoAgencia & " "
            strSql = strSql & "(intAgencia, intTipoDeComunicacao, strConteudo, strDescricao, "
            strSql = strSql & "intSequencia) Values ("
            strSql = strSql & Val(txtPKID) & ", "
            strSql = strSql & .ListItems(intI).Tag & ", '"
            strSql = strSql & .ListItems(intI).SubItems(1) & "', '"
            strSql = strSql & .ListItems(intI).SubItems(2) & "', "
            strSql = strSql & intI & ")"
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSql
        Next
    End With
End Sub

Sub DeletaTipoComunicacao(txtPKID As Integer)
    Dim strSql As String

    strSql = ""
    strSql = strSql & "Delete From " & gstrFormaDeComunicacaoAgencia & " "
    strSql = strSql & "Where intAgencia = " & Val(txtPKID)
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
End Sub

Sub CarregaTipoComunicacao(txtPKID As Integer)

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
'    strSql = strSql & "Left Join " & gstrFormaDeComunicacaoAgencia & " FC "
    strSql = strSql & ", " & gstrFormaDeComunicacaoAgencia & " FC "
'    strSql = strSql & "On TP.PKId = FC.intTipoDeComunicacao "
    strSql = strSql & "Where FC.intAgencia = " & Val(txtPKID) & " "
    strSql = strSql & " AND TP.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & " FC.intTipoDeComunicacao "
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
    txt_Conteudo = ""
    txt_DescricaoConteudo = ""
    lbl_TipoComunicacao = "Tipo"
'    If tdb_Agencia.ListItems.Count <> 0 Then
'        tdb_Agencia.SelectedItem.Selected = False
'    End If
    tab_3dPasta.Tab = 0
End Sub

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    If dbcintBanco.MatchedWithList = False Then
        ExibeMensagem "O banco deve ser preenchido corretamente."
        dbcintBanco.SetFocus
        Exit Function
    ElseIf Trim(txtstrAgencia) = "" Then
        ExibeMensagem "A agência deve ser preenchida corretamente."
        txtstrAgencia.SetFocus
        Exit Function
    ElseIf Trim(txtstrDescricao) = "" Then
        ExibeMensagem "A descrição deve ser preenchida corretamente."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(Trim(txtstrDescricao.Text)) <> UCase$(Trim(strDescricaoAtual))) Then
        If gblnExisteCodigo(2, gstrAgencia, "intBanco", dbcintBanco.BoundText, "strDescricao", "'" & Trim(txtstrDescricao) & "'") Then
            ExibeMensagem "Já existe registro com essa descrição para esse banco."
            Exit Function
        End If
    End If
    
    If Trim(txtstrLogradouro) = "" Then
        ExibeMensagem "O campo descrição do logradouro deve ser informado."
        txtstrLogradouro.SetFocus
        Exit Function
    ElseIf Trim(txtstrBairro) = "" Then
        ExibeMensagem "O bairro deve ser informado."
        txtstrBairro.SetFocus
        Exit Function
    ElseIf Not dbcintCidade.MatchedWithList Then
        ExibeMensagem "A cidade deve ser informada."
        dbcintCidade.SetFocus
        Exit Function
    ElseIf Not dbcintUF.MatchedWithList Then
        ExibeMensagem "O campo UF deve ser informado."
        dbcintUF.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
End Function

Private Sub txt_Conteudo_Change()
    If lvw_TipoComunicacao.ListItems.Count = 0 Then Exit Sub
    If lvw_TipoComunicacao.SelectedItem.Selected = False Then Exit Sub
    lvw_TipoComunicacao.SelectedItem.SubItems(1) = Trim(txt_Conteudo)
End Sub

Private Sub txt_Conteudo_GotFocus()
  MarcaCampo txt_Conteudo
    tab_3dPasta.Tab = 1
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
    tab_3dPasta.Tab = 1
End Sub

Private Sub txt_DescricaoConteudo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_DescricaoConteudo
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

Sub NovoTipo()
    lvw_TipoComunicacao.ListItems.Clear
    'txtPKId = glngPegaProximaChave(gstrAgencia, "PKId")
    txt_Conteudo = ""
    txt_DescricaoConteudo = ""
    
    lbl_TipoComunicacao = "Tipo"
    tab_3dPasta.Tab = 0
End Sub

'----------------^^^^COMUNICAÇÃO^^^^--------------'

Function strQueryRelatorio() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT DISTINCT AG.strAgencia, AG.strDescricao, BC.strDescricao Banco, FG.strConteudo ContextoCom,"
    strSql = strSql & " TC.strDescricao TipoCom "
    strSql = strSql & " FROM " & gstrFormaDeComunicacaoAgencia & " FG, "
    strSql = strSql & gstrTipoDeComunicacao & " TC, "
    strSql = strSql & gstrBanco & " BC, "
    strSql = strSql & gstrAgencia & " AG "
    If Val(dbcintBanco.BoundText) <> 0 Then
        strSql = strSql & "WHERE FG.intAgencia = AG.PKId "
        strSql = strSql & "AND FG.intTipoDeComunicacao = TC.PKId "
        strSql = strSql & "AND AG.intBanco = BC.PKId "
        strSql = strSql & "AND BC.PKId = " & dbcintBanco.BoundText
    Else
        strSql = strSql & "WHERE BC.PKID = AG.intBanco "
        strSql = strSql & "AND TC.PKID = FG.intTipodeComunicacao "
        strSql = strSql & "AND AG.PKID = FG.intAgencia "
        strSql = strSql & "UNION "
        strSql = strSql & "SELECT AG.strAgencia, AG.strDescricao, BC.strDescricao Banco,"
        strSql = strSql & " '','' "
        strSql = strSql & "FROM "
        strSql = strSql & gstrBanco & " BC, "
        strSql = strSql & gstrAgencia & " AG "
        strSql = strSql & "WHERE BC.PKID = AG.intBanco "
        strSql = strSql & "AND AG.PKID not in (SELECT intAgencia FROM " & gstrFormaDeComunicacaoAgencia & " )"
    End If
    strSql = strSql & " ORDER BY Banco"
strQueryRelatorio = strSql
End Function
