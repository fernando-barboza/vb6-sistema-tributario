VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadCatalogoAssunto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogos de Assunto"
   ClientHeight    =   5835
   ClientLeft      =   2190
   ClientTop       =   2130
   ClientWidth     =   7935
   HelpContextID   =   104
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKId 
      Height          =   285
      Left            =   6090
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5655
      Left            =   105
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   90
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Catálogos de Assunto"
      TabPicture(0)   =   "CadCatalogoAssunto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintTipoAssunto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintTipoMaterialServico"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblstrCodCatalogoAssunto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintUnidadeCentroDeCusto"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_PrazoMinimo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_PrazoFinal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_PrazoEstimado"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_dtmDtCancelamento"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "tdb_catalogo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbcintTipoAssunto"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dbcintGrupoAssunto"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtstrDescricao"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtstrCodCatalogoAssunto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmd_Grupo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmd_Tipo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dbcintUnidadeCentroDeCusto"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmd_CentroCusto"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtIntPrazoMinimo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtIntPrazoMaximo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtIntPrazoPrevisto"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtdtmDtCancelamento"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Preço Público"
      TabPicture(1)   =   "CadCatalogoAssunto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Receitas"
      Tab(1).Control(1)=   "txtstrhistoricopadrao"
      Tab(1).Control(2)=   "lblstrhistoricopadrao"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtdtmDtCancelamento 
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
         Left            =   4575
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1150
         Width           =   1005
      End
      Begin VB.Frame fra_Receitas 
         Caption         =   "Receitas"
         Height          =   3585
         Left            =   -74850
         TabIndex        =   27
         Top             =   1740
         Width           =   7455
         Begin VB.CommandButton cmd_Receitas 
            Height          =   300
            Left            =   5115
            Picture         =   "CadCatalogoAssunto.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Tag             =   "439"
            ToolTipText     =   "Ativa Cadastro de Receitas"
            Top             =   390
            Width           =   330
         End
         Begin MSDataListLib.DataCombo dbc_intReceita 
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   390
            Width           =   4980
            _ExtentX        =   8784
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvw_Itens 
            Height          =   2505
            Left            =   90
            TabIndex        =   12
            Top             =   900
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   4419
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Pkid"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descrição da Receita"
               Object.Width           =   12347
            EndProperty
         End
      End
      Begin VB.TextBox txtstrhistoricopadrao 
         Height          =   885
         Left            =   -74070
         MaxLength       =   500
         TabIndex        =   10
         Top             =   780
         Width           =   6645
      End
      Begin VB.TextBox txtIntPrazoPrevisto 
         Height          =   285
         Left            =   6570
         MaxLength       =   4
         TabIndex        =   8
         Top             =   2430
         Width           =   1005
      End
      Begin VB.TextBox txtIntPrazoMaximo 
         Height          =   285
         Left            =   4335
         MaxLength       =   4
         TabIndex        =   7
         Top             =   2430
         Width           =   1005
      End
      Begin VB.TextBox txtIntPrazoMinimo 
         Height          =   285
         Left            =   2220
         MaxLength       =   4
         TabIndex        =   6
         Top             =   2430
         Width           =   1005
      End
      Begin VB.CommandButton cmd_CentroCusto 
         Height          =   300
         Left            =   7245
         Picture         =   "CadCatalogoAssunto.frx":03C2
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "439"
         ToolTipText     =   "Ativa o Cadastro de Unidade de Centro de Custo"
         Top             =   2040
         Width           =   330
      End
      Begin MSDataListLib.DataCombo dbcintUnidadeCentroDeCusto 
         Height          =   315
         Left            =   2220
         TabIndex        =   5
         Top             =   2040
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton cmd_Tipo 
         Height          =   300
         Left            =   7245
         Picture         =   "CadCatalogoAssunto.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "439"
         ToolTipText     =   "Ativa Cadastro de Tipo de Assunto"
         Top             =   780
         Width           =   330
      End
      Begin VB.CommandButton cmd_Grupo 
         Height          =   300
         Left            =   7245
         Picture         =   "CadCatalogoAssunto.frx":0AD6
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "438"
         ToolTipText     =   "Ativa Cadastro de Grupos de Assunto"
         Top             =   435
         Width           =   330
      End
      Begin VB.TextBox txtstrCodCatalogoAssunto 
         Height          =   285
         Left            =   2220
         MaxLength       =   10
         OLEDragMode     =   1  'Automatic
         TabIndex        =   2
         Top             =   1150
         Width           =   1005
      End
      Begin VB.TextBox txtstrDescricao 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   525
         Left            =   2220
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1470
         Width           =   5370
      End
      Begin MSDataListLib.DataCombo dbcintGrupoAssunto 
         Height          =   315
         Left            =   2220
         TabIndex        =   0
         Top             =   450
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintTipoAssunto 
         Height          =   315
         Left            =   2220
         TabIndex        =   1
         Top             =   795
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_catalogo 
         Height          =   2595
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4577
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKID"
         Columns(0).DataField=   "PKID"
         Columns(0).NumberFormat=   "FormatText Event"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "strCodCatalogoAssunto"
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1588"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1508"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=9843"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=9763"
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
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   3
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
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
      Begin VB.Label lbl_dtmDtCancelamento 
         AutoSize        =   -1  'True
         Caption         =   "Cancelamento"
         Height          =   195
         Left            =   3450
         TabIndex        =   29
         Top             =   1170
         Width           =   1020
      End
      Begin VB.Label lblstrhistoricopadrao 
         AutoSize        =   -1  'True
         Caption         =   "Histórico"
         Height          =   195
         Left            =   -74730
         TabIndex        =   26
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lbl_PrazoEstimado 
         AutoSize        =   -1  'True
         Caption         =   "Prazo Estimado"
         Height          =   195
         Left            =   5430
         TabIndex        =   25
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label lbl_PrazoFinal 
         AutoSize        =   -1  'True
         Caption         =   "Prazo máximo"
         Height          =   195
         Left            =   3330
         TabIndex        =   24
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label lbl_PrazoMinimo 
         AutoSize        =   -1  'True
         Caption         =   "Prazo mínimo"
         Height          =   195
         Left            =   1170
         TabIndex        =   23
         Top             =   2460
         Width           =   960
      End
      Begin VB.Label lblintUnidadeCentroDeCusto 
         AutoSize        =   -1  'True
         Caption         =   "Unidade de Centro de Custo"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2100
         Width           =   2010
      End
      Begin VB.Label lblstrCodCatalogoAssunto 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1650
         TabIndex        =   17
         Top             =   1170
         Width           =   495
      End
      Begin VB.Label lblintTipoMaterialServico 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Assunto"
         Height          =   195
         Left            =   990
         TabIndex        =   16
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblintTipoAssunto 
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Assunto"
         Height          =   195
         Left            =   870
         TabIndex        =   15
         Top             =   495
         Width           =   1275
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1425
         TabIndex        =   14
         Top             =   1500
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadCatalogoAssunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim mblnAlterando       As Boolean
   Dim mblnAlterandoAux    As Boolean
   Dim mobjAux             As Object
   Dim mblnSelecionou      As Boolean
   Dim mblnPrimeiraVez     As Boolean
   Dim mobjLista           As Object
   Dim mblnActivate        As Boolean
   Dim strCodigoAtual      As String
   Dim strDescricaoAtual   As String
   Dim strCodigo           As String
   Dim intPkid             As Long
   Dim bytOrdenacao        As Byte
   Dim blnOrdenacaoAsc     As Boolean


Private Sub cmd_CentroCusto_Click()
   CarregaForm frmCadLocais, dbcintUnidadeCentroDeCusto
End Sub

Private Sub cmd_Receitas_Click()
    CarregaForm frmCadReceita, dbc_intReceita
End Sub

Private Sub cmd_Tipo_Click()
    CarregaForm frmCadTipoAssunto, dbcintTipoAssunto, strQueryAplicar
End Sub
Private Sub dbcintUnidadeCentroDeCusto_Click(Area As Integer)
    DropDownDataCombo dbcintUnidadeCentroDeCusto, Me, Area
End Sub

Private Sub dbcintUnidadeCentroDeCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUnidadeCentroDeCusto, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUnidadeCentroDeCusto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintGrupoAssunto_Click(Area As Integer)
   If Area = 2 And dbcintGrupoAssunto.MatchedWithList Then
'       Set tdb_catalogo.DataSource = Nothing
'       tdb_catalogo.ReBind
'       tdb_catalogo.Refresh
       LeDaTabelaParaObj gstrTipoAssunto, dbcintTipoAssunto, strQueryAssunto
   ElseIf Area = 0 Then
      DropDownDataCombo dbcintGrupoAssunto, Me, Area
   End If
End Sub

Private Sub dbcintGrupoAssunto_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintGrupoAssunto, Me, , KeyCode, Shift
End Sub

Private Sub dbcintGrupoAssunto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintTipoAssunto_Click(Area As Integer)
   DropDownDataCombo dbcintTipoAssunto, Me, Area
   'If Area = 2 And dbcintTipoAssunto.MatchedWithList Then
   '    LeDaTabelaParaObj gstrCatalogoAssunto, tdb_catalogo, strQuery
   '    LeDaTabelaParaObj gstrLocais, dbcintUnidadeCentroDeCusto, strQueryLocais
   'End If
End Sub

Private Sub dbcintTipoAssunto_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTipoAssunto, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoAssunto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 450
    VirificaGradeListView Me
    
    If UCase(App.ProductName) = "TRIBUTARIO" Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar
    Else
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
        
        If Not mblnActivate Then VerificaObjParaAplicar mobjAux
        
        mblnActivate = True
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
    mblnAlterandoAux = False
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
    Me.Icon = MDIMenu.Icon
    dbcintGrupoAssunto.Tag = strQueryGrupoAssunto & ";A.strDescricao"
    dbcintTipoAssunto.Tag = strQueryTipoAssunto & ";A.strDescricao"
    dbcintUnidadeCentroDeCusto.Tag = strQueryLocais & ";A.strDescricao"
    dbc_intReceita.Tag = strQueryReceita & ";strDescricao"
    bytOrdenacao = 1: blnOrdenacaoAsc = True
    If UCase(App.ProductName) = "TRIBUTARIO" Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    mblnSelecionou = False
    mblnPrimeiraVez = False
    mblnActivate = False
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If UCase(App.ProductName) <> "TRIBUTARIO" Then
        If tab_3dPasta.Tab = 1 Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        End If
    End If
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Catalogo_Click()
    mblnPrimeiraVez = True
    mblnAlterando = True
    If glngQtdLinhaTDBGrid(tdb_catalogo) = 1 Then
        tdb_Catalogo_RowColChange 0, 0
    End If
End Sub

Sub tdb_Catalogo_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Catalogo_FilterChange()
    gblnFilraCampos tdb_catalogo
    mblnAlterando = False
End Sub

Private Sub tdb_catalogo_HeadClick(ByVal ColIndex As Integer)
    mblnAlterando = False
    gOrdenaGrid tdb_catalogo, ColIndex
End Sub

Private Sub tdb_catalogo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Catalogo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_catalogo
        If Not .EOF And Not .BOF And mblnAlterando Then
            txtPKId.Text = .Columns("PKID").Value
            intPkid = .Columns("PKID").Value
            If mblnPrimeiraVez Then
                LeDaTabelaParaObj gstrCatalogoAssunto, Me
                gCorLinhaSelecionada tdb_catalogo
                If UCase(App.ProductName) <> "TRIBUTARIO" Then
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                    If mobjAux Is Nothing Then
                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                    End If
                Else
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar
                End If
                mblnSelecionou = True
                mblnAlterando = True
                PreencheGridReceita
                strCodigoAtual = txtstrCodCatalogoAssunto.Text
                strDescricaoAtual = txtstrDescricao.Text
                
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark As Variant
    Dim strSql As String
    
    strSql = ""
    If strModoOperacao = gstrImprimir Then
        ToolBarGeral strModoOperacao, gstrCatalogoAssunto, mblnAlterando, tdb_catalogo, Me, _
                     mobjAux, strQuery, , rptCatalogoDeAssunto, strQueryRelatorio
        Exit Sub
    End If
    
    dbcintUnidadeCentroDeCusto.Tag = strQueryLocais & ";A.strDescricao"
    strSql = strQuery
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
        If UCase(strModoOperacao) = UCase(gstrDeletar) Then
        
            gobjBanco.ExecutaBeginTrans
            strSql = "DELETE FROM " & gstrCatalogoAssuntoReceita
            strSql = strSql & " WHERE "
            strSql = strSql & " IntCatalogoAssunto = " & Val(intPkid)
            gobjBanco.Execute strSql
            
            If ToolBarGeral(strModoOperacao, gstrCatalogoAssunto, mblnAlterando, tdb_catalogo, Me, mobjAux, strQuery, , rptCatalogoDeAssunto, strQueryRelatorio) Then
                If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
                gobjBanco.ExecutaCommitTrans
                intPkid = 0
                mblnAlterando = False
                Limpa_Controles Me, True, True, True, True, True
                tab_3dPasta.Tab = 0
                dbcintGrupoAssunto.SetFocus
            Else
                gobjBanco.ExecutaRollbackTrans
            End If
            Exit Sub
        Else
            If blnDadosOk Then
                mblnPrimeiraVez = False
                If mblnAlterando Then mblnAlterandoAux = True
                If ToolBarGeral(strModoOperacao, gstrCatalogoAssunto, mblnAlterando, tdb_catalogo, Me, mobjAux, strQuery, , rptCatalogoDeAssunto, strQueryRelatorio) Then
                    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
                    GravaReceita
                    Limpa_Controles Me, True, True, True, True, True
                    tab_3dPasta.Tab = 0
                    dbcintGrupoAssunto.SetFocus
                End If
            End If
            Exit Sub
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrLocalizar) Then
        
        'If dbcintGrupoAssunto.Text = "" And dbcintTipoAssunto.Text = "" Then
        '    LeDaTabelaParaObj gstrCatalogoAssunto, tdb_catalogo, strQueryGrid
        '    Exit Sub
        'Else
            LeDaTabelaParaObj gstrCatalogoAssunto, tdb_catalogo, strQuery
        '    Exit Sub
        'End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrRefresh) Then
       LeDaTabelaParaObj gstrCatalogoAssunto, tdb_catalogo, strQueryRefresh
       Exit Sub
    End If
    
    ToolBarGeral strModoOperacao, gstrCatalogoAssunto, mblnAlterando, tdb_catalogo, Me, _
                mobjAux, strQuery, , rptCatalogoDeAssunto, strQueryRelatorio
    
    
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
    If UCase(strModoOperacao) = UCase(gstrDeletar) Then
        tdb_catalogo.MoveFirst
    ElseIf UCase(strModoOperacao) = UCase(gstrNovo) Then
        If txtstrCodCatalogoAssunto.Enabled Then
            Limpa_Controles Me, True, True, True, True, True
            intPkid = 0
            dbcintGrupoAssunto.SetFocus
        End If
    End If

If UCase(strModoOperacao) = UCase(gstrIncluirItem) Then
    If blnDadosReceitaOk Then
        If dbc_intReceita.MatchedWithList Then
            IncluirItemNoGrid
        End If
    End If
ElseIf UCase(strModoOperacao) = UCase(gstrExcluirItem) Then
    ExcluirItemNoGrid
End If

If UCase(App.ProductName) = "TRIBUTARIO" Then
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar
End If
End Sub

Private Function strQuery() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "A.Pkid, "
    strSql = strSql & "A.strCodCatalogoAssunto, "
    strSql = strSql & "A.strDescricao "
    strSql = strSql & "From "
    strSql = strSql & gstrCatalogoAssunto & " A, "
    strSql = strSql & gstrGrupoAssunto & " B, "
    strSql = strSql & gstrTipoAssunto & " c "
    strSql = strSql & "Where "
    strSql = strSql & "A.IntGrupoAssunto = B.PKId  AND "
    strSql = strSql & "A.intTipoAssunto = c.Pkid "
    
    If dbcintGrupoAssunto.MatchedWithList Then
        strSql = strSql & " AND A.IntGrupoAssunto = '" & dbcintGrupoAssunto.BoundText & "'"
    End If
    If dbcintTipoAssunto.MatchedWithList Then
        strSql = strSql & " AND A.intTipoAssunto = '" & dbcintTipoAssunto.BoundText & "'"
    End If
    If Trim(txtstrDescricao.Text) <> "" Then
        strSql = strSql & " AND UPPER(A.strDescricao) LIKE '" & UCase(txtstrDescricao.Text) & "%'"
    End If
    If dbcintUnidadeCentroDeCusto.MatchedWithList Then
        strSql = strSql & " AND intUnidadeCentroDeCusto = '" & dbcintUnidadeCentroDeCusto.BoundText & "'"
    End If

    strSql = strSql & "Order By "
    strSql = strSql & "A.strDescricao "

    
    strQuery = strSql
End Function

Private Function strQueryAssunto() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT B.PKID, B.strDescricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrGrupoAssunto & " A ,"
    strSql = strSql & gstrTipoAssunto & " B "
    strSql = strSql & " WHERE "
    strSql = strSql & " B.intGrupoAssunto = A.PKID "
    strSql = strSql & " AND B.intGrupoAssunto = '" & dbcintGrupoAssunto.BoundText & "'"
    strSql = strSql & " ORDER BY B.strDescricao"
    strQueryAssunto = strSql
End Function

Private Sub cmd_Grupo_Click()
    CarregaForm frmCadGrupoAssunto, dbcintGrupoAssunto
End Sub

Private Function strQueryAplicar()
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKID, strdescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrCatalogoAssunto
    strQueryAplicar = strSql
End Function

Private Sub txtdtmDtCancelamento_GotFocus()
    MarcaCampo txtdtmdtcancelamento
End Sub

Private Sub txtdtmDtCancelamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdtcancelamento
End Sub

Private Sub txtdtmDtCancelamento_LostFocus()
    txtdtmdtcancelamento = gstrDataFormatada(txtdtmdtcancelamento)
End Sub

Private Sub txtIntPrazoMaximo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtIntPrazoMaximo
End Sub

Private Sub txtIntPrazoMinimo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtIntPrazoMinimo
End Sub

Private Sub txtIntPrazoPrevisto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtIntPrazoPrevisto
End Sub

Private Sub txtstrCodCatalogoAssunto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodCatalogoAssunto
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Function strQueryRelatorio() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " A.strCodCatalogoAssunto, A.strDescricao, B.strDescricao AS Grupo, C.strDescricao AS Tipo "
    
    strSql = strSql & " FROM "
    strSql = strSql & gstrCatalogoAssunto & " A,"
    strSql = strSql & gstrGrupoAssunto & " B,"
    strSql = strSql & gstrTipoAssunto & " C "
    
    strSql = strSql & " WHERE "
    strSql = strSql & " A.IntGrupoAssunto = B.PKId "
    strSql = strSql & " AND A.intTipoAssunto = C.PKId "
     
    strSql = strSql & " ORDER BY "
    strSql = strSql & " B.strDescricao, C.strDescricao, A.strDescricao "
strQueryRelatorio = strSql
End Function

Private Function strQueryLocais() As String

'******************************************************************************************
' Data: 07/03/2003
' Alteração: - Retirada a palavra chave "AS" das cláusulas FROM, pois o Oracle não permite
'            a utilização desta palavra chave nesta cláusula.
' Responsável: Everton Bianchini
'******************************************************************************************
'
'******************************************************************************************
' Data: 24/06/2003
' Alteração: - Retirada a tabela tblUnidadeCentroDeCusto, pois será usada tblLocais
' Responsável: Gustavo Monteiro
'******************************************************************************************
Dim strSql As String

strSql = ""
strSql = strSql & " SELECT A.PkId, A.strDescricao"
strSql = strSql & " FROM"
'strSql = strSql & " " & gstrUnidadeCentroDeCusto2 & " AS A"
'strSql = strSql & " " & gstrUnidadeCentroDeCusto2 & " A"
strSql = strSql & " " & gstrLocais & " A"

strQueryLocais = strSql

End Function


Private Sub txtstrCodCatalogoAssunto_GotFocus()
    gstrProximoCodigo txtstrCodCatalogoAssunto, gstrCatalogoAssunto, "strCodCatalogoAssunto", gintCodSeguranca, "intTipoAssunto", dbcintTipoAssunto.BoundText, , , "intGrupoAssunto", dbcintGrupoAssunto.BoundText
End Sub

Private Function blnDadosOk() As Boolean
    
    If dbcintGrupoAssunto.Text = "" Or Not dbcintGrupoAssunto.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo grupo de assunto!"
        dbcintGrupoAssunto.SetFocus
        Exit Function
    ElseIf dbcintTipoAssunto.Text = "" Or Not dbcintTipoAssunto.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo tipo de assunto!"
        dbcintTipoAssunto.SetFocus
        Exit Function
    ElseIf Trim(txtstrCodCatalogoAssunto.Text) = "" Then
        ExibeMensagem "Preencha corretamente o campo código!"
        txtstrCodCatalogoAssunto.SetFocus
        Exit Function
    ElseIf Trim(txtdtmdtcancelamento.Text) <> "" Then
        If Not gblnDataValida(txtdtmdtcancelamento, True) Then Exit Function
    ElseIf txtstrDescricao.Text = "" Then
        ExibeMensagem "Preencha corretamente o campo descrição!"
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    If Not dbcintUnidadeCentroDeCusto.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo unidade de centro de custo!"
        dbcintUnidadeCentroDeCusto.SetFocus
        Exit Function
    End If
    
    If txtIntPrazoMinimo <> "" And txtIntPrazoMaximo <> "" Then
      If txtIntPrazoMinimo <> "" Then
         If txtIntPrazoMaximo <> "" Then
            If txtIntPrazoMinimo > txtIntPrazoMaximo Then
               ExibeMensagem "O campo prazo máximo deve ser maior que o prazo mínimo!"
               txtIntPrazoMaximo.SetFocus
               Exit Function
            Else
               If txtIntPrazoPrevisto <> "" Then
                  If txtIntPrazoPrevisto < txtIntPrazoMinimo Or txtIntPrazoPrevisto > txtIntPrazoMaximo Then
                     ExibeMensagem "O campo prazo estimado deve ser maior ou igual ao prazo mínimo e menor ou igual ao prazo máximo!"
                     txtIntPrazoPrevisto.SetFocus
                     Exit Function
                  End If
               End If
            End If
         End If
      End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtstrCodCatalogoAssunto.Text)) Then

ProximoCodigo:

        If gblnExisteCodigo(2, gstrCatalogoAssunto, "strCodCatalogoAssunto", "'" & txtstrCodCatalogoAssunto.Text & "'", "intTipoAssunto", dbcintTipoAssunto.BoundText, "intgrupoAssunto", dbcintGrupoAssunto.BoundText) Then
            strCodigo = (gstrProximoCodigo(txtstrCodCatalogoAssunto, gstrCatalogoAssunto, "strCodCatalogoAssunto", gintCodSeguranca, "intTipoAssunto", dbcintTipoAssunto.BoundText, , True, "intGrupoAssunto", dbcintGrupoAssunto.BoundText))
            If MsgBox("O código informado já se encontra cadastrado. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                txtstrCodCatalogoAssunto.SetFocus
                Exit Function
            Else
                txtstrCodCatalogoAssunto.Text = strCodigo
                GoTo ProximoCodigo
            End If
        End If
        
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescricaoAtual)) Then
            
        If gblnExisteCodigo(1, gstrCatalogoAssunto, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
        
    blnDadosOk = True
    
End Function

Private Function strQueryGrupoAssunto() As String
Dim strSql As String
strSql = ""

strSql = "SELECT A.Pkid, A.strDescricao "
strSql = strSql & "FROM " & gstrGrupoAssunto & " A"
strSql = strSql & " ORDER BY A.strDescricao"
strQueryGrupoAssunto = strSql

End Function

Private Function strQueryTipoAssunto() As String

Dim strSql As String
strSql = ""
strSql = "SELECT A.Pkid, A.strDescricao"
strSql = strSql & " FROM " & gstrTipoAssunto & " A"
strSql = strSql & " ORDER BY A.strDescricao"
strQueryTipoAssunto = strSql

End Function

Private Function strQueryGrid() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "A.Pkid, "
    strSql = strSql & "A.strCodCatalogoAssunto, "
    strSql = strSql & "A.strDescricao "
    strSql = strSql & "From "
    strSql = strSql & gstrCatalogoAssunto & " A, "
    strSql = strSql & gstrGrupoAssunto & " B, "
    strSql = strSql & gstrTipoAssunto & " c "
    strSql = strSql & "Where "
    strSql = strSql & "A.IntGrupoAssunto = B.PKId  AND "
    strSql = strSql & "A.intTipoAssunto = c.Pkid "
    strSql = strSql & "Order By "
    strSql = strSql & "A.strDescricao "
    strQueryGrid = strSql
End Function

Private Function strQueryReceita() As String
    Dim strSql As String
    
    strSql = ""
    strSql = "SELECT Pkid, strDescricao "
    strSql = strSql & "FROM " & gstrReceita
    strSql = strSql & " ORDER BY strDescricao "
    
    strQueryReceita = strSql
    
End Function

Private Sub GravaReceita()
    Dim strSql  As String
    Dim Pkid As Long
    Dim intInd As Integer
    Dim blnStr As Boolean
    
    strSql = ""
    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    If mblnAlterandoAux And intPkid <> 0 Then
        Pkid = intPkid
        intPkid = 0
    Else
        Pkid = glngPegaUltimaChave(gstrCatalogoAssunto, "Pkid")
    End If
    
    If mblnAlterandoAux Then
        blnStr = True
        mblnAlterandoAux = False
        strSql = strSql & "DELETE FROM " & gstrCatalogoAssuntoReceita
        strSql = strSql & " WHERE "
        strSql = strSql & " IntCatalogoAssunto = " & Val(Pkid) & IIf(bytDBType = Oracle, ";", "")
    End If
    
    With lvw_Itens
        If .ListItems.Count >= 1 Then
            For intInd = 1 To .ListItems.Count
                blnStr = True
                strSql = strSql & "INSERT INTO "
                strSql = strSql & gstrCatalogoAssuntoReceita & " ("
                strSql = strSql & "IntCatalogoAssunto, "
                strSql = strSql & "intreceita, "
                strSql = strSql & "dtmDtAtualizacao, "
                strSql = strSql & "lngCodUsr) "
                strSql = strSql & "Values(" & Pkid & ","
                strSql = strSql & .ListItems(intInd).Text & ", "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSql = strSql & glngCodUsr & " "
                strSql = strSql & ")" & IIf(bytDBType = Oracle, ";", "")
            Next
        End If
    End With
    

    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
        
    Set gobjBanco = New clsBanco
    If blnStr Then gobjBanco.Execute strSql
    Set gobjBanco = Nothing
End Sub

Private Function blnDadosReceitaOk() As Boolean
    blnDadosReceitaOk = False
    If dbc_intReceita.MatchedWithList = False Then
        ExibeMensagem "É necessário selecionar uma receita."
    End If
    blnDadosReceitaOk = True
End Function

Private Function IncluirItemNoGrid()
    Dim intInd          As Integer
    With lvw_Itens
        For intInd = 1 To .ListItems.Count
            If dbc_intReceita.BoundText = .ListItems(intInd).Text Then
                ExibeMensagem "Não é possível incluir Receitas iguais."
                Exit Function
            End If
        Next
    End With
    Set mobjLista = lvw_Itens.ListItems.Add(, , dbc_intReceita.BoundText)
    mobjLista.SubItems(1) = dbc_intReceita.Text
End Function

Private Function ExcluirItemNoGrid()
    With lvw_Itens
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
End Function

Private Sub PreencheGridReceita()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = ""
    
    strSql = strSql & "Select "
    strSql = strSql & "R.Pkid, "
    strSql = strSql & "R.strDescricao "
    strSql = strSql & "From "
    strSql = strSql & gstrCatalogoAssunto & " CA, "
    strSql = strSql & gstrCatalogoAssuntoReceita & " CAR, "
    strSql = strSql & gstrReceita & " R "
    strSql = strSql & "Where "
    strSql = strSql & "CA.pkid = CAR.Intcatalogoassunto AND "
    strSql = strSql & "R.Pkid = CAR.Intreceita AND "
    strSql = strSql & "CA.Pkid = " & txtPKId
    strSql = strSql & " Order By R.strDescricao"
    lvw_Itens.ListItems.Clear
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        If adoResultado.RecordCount >= 1 Then
            Do While Not adoResultado.EOF
                Set mobjLista = lvw_Itens.ListItems.Add(, , gstrENulo(adoResultado!Pkid))
                mobjLista.SubItems(1) = gstrENulo(adoResultado!strDescricao)
                adoResultado.MoveNext
            Loop
        End If
    End If
    Set gobjBanco = Nothing
    

End Sub

Private Function strQueryRefresh() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "A.Pkid, "
    strSql = strSql & "A.strCodCatalogoAssunto, "
    strSql = strSql & "A.strDescricao "
    strSql = strSql & "From "
    strSql = strSql & gstrCatalogoAssunto & " A, "
    strSql = strSql & gstrGrupoAssunto & " B, "
    strSql = strSql & gstrTipoAssunto & " c "
    strSql = strSql & "Where "
    strSql = strSql & "A.IntGrupoAssunto = B.PKId  AND "
    strSql = strSql & "A.intTipoAssunto = c.Pkid "
    Select Case bytOrdenacao
        Case Is = 1
            strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_INT, "strCodCatalogoAssunto") & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSql = strSql & " ORDER BY strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
     
    strQueryRefresh = strSql
End Function


