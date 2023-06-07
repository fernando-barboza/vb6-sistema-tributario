VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadAtividadeEconomica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atividades Econômicas"
   ClientHeight    =   7110
   ClientLeft      =   2730
   ClientTop       =   2265
   ClientWidth     =   6345
   HelpContextID   =   137
   Icon            =   "frmCadAtividades.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   6345
   Tag             =   "153"
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   3660
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
      Height          =   2055
      Left            =   60
      TabIndex        =   7
      Top             =   5010
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   3625
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
      Columns(2).Caption=   "Atividade"
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
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1667"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1588"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=8625"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=8546"
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
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   4995
      Left            =   60
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   30
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   8811
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Atividades Econômicas"
      TabPicture(0)   =   "frmCadAtividades.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintUtilizacao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintCodigo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrObservacoes"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintSubGrupo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintGrupo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrDescricao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_intAtividadeBasica"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbcintatividadeBasica"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dbcintUtilizacao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dbcintSubGrupo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbcintGrupo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtintCodigo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtstrObservacoes"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmd_intSubGrupo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmd_intGrupo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtstrDescricao"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmd_intAtiviadeBasica"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Tributos"
      TabPicture(1)   =   "frmCadAtividades.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frm_Tributos"
      Tab(1).Control(1)=   "frm_TipoTributo"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmd_intAtiviadeBasica 
         Height          =   315
         Left            =   5685
         Picture         =   "frmCadAtividades.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de Atividade Básica"
         Top             =   2310
         Width           =   360
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   975
         MaxLength       =   60
         TabIndex        =   4
         Top             =   1995
         Width           =   5085
      End
      Begin VB.CommandButton cmd_intGrupo 
         Height          =   315
         Left            =   5685
         Picture         =   "frmCadAtividades.frx":1198
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de Grupos de Atividade"
         Top             =   795
         Width           =   360
      End
      Begin VB.CommandButton cmd_intSubGrupo 
         Height          =   315
         Left            =   5685
         Picture         =   "frmCadAtividades.frx":12B6
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de SubGrupos de Atividade"
         Top             =   1185
         Width           =   360
      End
      Begin VB.TextBox txtstrObservacoes 
         Height          =   1560
         Left            =   975
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2700
         Width           =   5085
      End
      Begin VB.TextBox txtintCodigo 
         Height          =   285
         Left            =   975
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1650
         Width           =   1305
      End
      Begin MSDataListLib.DataCombo dbcintGrupo 
         Height          =   315
         Left            =   975
         TabIndex        =   1
         Top             =   810
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintSubGrupo 
         Height          =   315
         Left            =   975
         TabIndex        =   2
         Top             =   1200
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintUtilizacao 
         Height          =   315
         Left            =   975
         TabIndex        =   0
         Top             =   420
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintatividadeBasica 
         Height          =   315
         Left            =   975
         TabIndex        =   5
         Top             =   2325
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Frame frm_TipoTributo 
         Caption         =   "Tipo de tributo"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   20
         Top             =   450
         Width           =   5925
         Begin VB.CommandButton cmd_intTributoTipo 
            Height          =   315
            Left            =   5490
            Picture         =   "frmCadAtividades.frx":13D4
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Ativa Cadastro de Tipos de Tributos"
            Top             =   300
            Width           =   360
         End
         Begin MSComctlLib.ListView lvw_Leis 
            Height          =   1380
            Left            =   90
            TabIndex        =   24
            Top             =   660
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   2434
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
               Text            =   "Pkid_Tipo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Tipo do Tributo"
               Object.Width           =   10125
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbc_intTributoTipo 
            Height          =   315
            Left            =   540
            TabIndex        =   22
            Top             =   300
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_intTipo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Frame frm_Tributos 
         Caption         =   "Tributos"
         Height          =   2145
         Left            =   -74910
         TabIndex        =   25
         Top             =   2730
         Width           =   5925
         Begin VB.CommandButton cmd_Tributos 
            Height          =   315
            Left            =   5475
            Picture         =   "frmCadAtividades.frx":14F2
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Ativa Cadastro de Tipos de Tributos"
            Top             =   270
            Width           =   360
         End
         Begin MSComctlLib.ListView lvw_Tributos 
            Height          =   1380
            Left            =   90
            TabIndex        =   29
            Top             =   630
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   2434
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Pkid_Tipo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Tributo"
               Object.Width           =   10107
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "int"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbc_intTributo 
            Height          =   315
            Left            =   645
            TabIndex        =   27
            Top             =   270
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_Tributos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tributo"
            Height          =   195
            Left            =   90
            TabIndex        =   26
            Top             =   330
            Width           =   495
         End
      End
      Begin VB.Label lbl_intAtividadeBasica 
         AutoSize        =   -1  'True
         Caption         =   "Ativ. Básica"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   2415
         Width           =   840
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Atividade"
         Height          =   195
         Left            =   270
         TabIndex        =   17
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label lblintGrupo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
         Height          =   195
         Left            =   495
         TabIndex        =   16
         Top             =   900
         Width           =   435
      End
      Begin VB.Label lblintSubGrupo 
         AutoSize        =   -1  'True
         Caption         =   "SubGrupo"
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lblstrObservacoes 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   2670
         Width           =   720
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   435
         TabIndex        =   13
         Top             =   1695
         Width           =   495
      End
      Begin VB.Label lblintUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   570
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmCadAtividadeEconomica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando      As Boolean
    Dim mblnAlterandoLista As Boolean
    Dim blnClickLista      As Boolean
    Dim mobjGeral          As Object
    Dim mlngUltimo         As Long
    Dim mlngUltimoG        As Long
    Dim mlngUltimoS        As Long
    Dim mobjAux            As Object
    Dim strDuplicataCodigo As String
    Dim mblnSelecionou     As Boolean
    Dim mblnPrimeiraVez    As Boolean
    Dim mobjLista          As Object
    Dim strCodigoAtual     As String
    Dim strDescricaoAtual  As String
    Dim strCodigo          As String
    Dim intID              As Double
    Dim mblnCarregando     As Boolean
    
    Dim vetTipo(3)         As String
    Dim adoRec             As ADODB.Recordset
    Dim adoTdb             As ADODB.Recordset
    Dim x                  As XArrayDB
    Dim y                  As New XArrayDB
    Dim Z                  As New XArrayDB
    
 ' TIMTIM - 11/02/2003 - Pendência nº 6
   Dim bytOrdenacao      As Byte
   Dim blnOrdenacaoAsc   As Boolean
   
Private Sub cmd_intGrupo_Click()
    If Not dbcintUtilizacao.MatchedWithList Then
        ExibeMensagem "A Utilização escolhida não é válida"
        mblnCarregando = True
        dbcintUtilizacao.SetFocus
        Exit Sub
    End If
    'ChamaFormCadastro
    CarregaForm frmCadGrupoDeAtividade, dbcintUtilizacao
    frmCadGrupoDeAtividade.dbcintUtilizacaoDaTabelaDeValor.Text = dbcintUtilizacao.Text
    frmCadGrupoDeAtividade.dbcintUtilizacaoDaTabelaDeValor.BoundText = dbcintUtilizacao.BoundText
    TrocaCorObjeto frmCadGrupoDeAtividade.dbcintUtilizacaoDaTabelaDeValor, True, True
    
End Sub

Private Sub cmd_intSubGrupo_Click()
    If Not dbcintUtilizacao.MatchedWithList Then
        ExibeMensagem "A Utilização escolhida não é válida"
        mblnCarregando = True
        dbcintUtilizacao.SetFocus
        Exit Sub
    Else
        If Not dbcintGrupo.MatchedWithList Then
            ExibeMensagem "O Grupo escolhido não é válido "
            dbcintGrupo.SetFocus
            Exit Sub
        End If
    End If
    CarregaForm frmCadSubGrupoDeAtividade, dbcintGrupo
    frmCadSubGrupoDeAtividade.dbcintCodigoDoGrupo.Text = dbcintGrupo.Text
    frmCadSubGrupoDeAtividade.dbcintCodigoDoGrupo.BoundText = dbcintGrupo.BoundText
    TrocaCorObjeto frmCadSubGrupoDeAtividade.dbcintCodigoDoGrupo, True, True
    
    
End Sub

Private Sub cmd_intTributoTipo_Click()
    CarregaForm frmCadTipoTributo, dbc_intTributoTipo
End Sub

Private Sub cmd_Tributos_Click()
    CarregaForm frmCadTributos, dbc_intTributo
End Sub

Private Sub dbc_intTributoTipo_Change()
    blnClickLista = False
End Sub

Private Sub dbc_intTributoTipo_GotFocus()
    If Not mblnCarregando Then
        tab_3DPasta.Tab = 1
    Else
        mblnCarregando = True
        dbcintUtilizacao.SetFocus
    End If
        mblnCarregando = False
End Sub

Private Sub dbcintatividadeBasica_Click(Area As Integer)
    tab_3DPasta.Tab = 0
End Sub

Private Sub dbcintGrupo_Click(Area As Integer)
    If Area = 2 And dbcintGrupo.MatchedWithList Then
        mlngUltimoG = dbcintGrupo.BoundText
    ElseIf Area = 0 Then
        DropDownDataCombo dbcintGrupo, Me, Area
    End If
End Sub

Private Sub dbcintGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintGrupo, Me, , KeyCode, Shift
End Sub

Private Sub dbcintGrupo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintGrupo
End Sub

Private Sub dbcintSubGrupo_Click(Area As Integer)
    If Area = 2 And dbcintSubGrupo.MatchedWithList Then
        mlngUltimoS = dbcintSubGrupo.BoundText
    ElseIf Area = 0 Then
        DropDownDataCombo dbcintSubGrupo, Me, Area
    End If
End Sub

Private Function strQueryAplicar() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId , strDescricao FROM "
    strSQL = strSQL & gstrAtividadeEC
    strQueryAplicar = strSQL
End Function

Private Sub dbcintSubGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintSubGrupo, Me, , KeyCode, Shift
End Sub

Private Sub dbcintSubGrupo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintSubGrupo
End Sub

Private Function strQueryGrupo() As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT A.PKId, A.strNomeDoGrupo "
    strSQL = strSQL & " FROM " & gstrGrupoDeAtividade & " A, "
    strSQL = strSQL & gstrUtilizacaoDaTabelaDeValor & " B "
    strSQL = strSQL & " WHERE B.PKId = A.intUtilizacaoDaTabelaDeValor "
    strSQL = strSQL & " AND bitIdentificador = 0 "
    
    If dbcintUtilizacao.MatchedWithList Then
        strSQL = strSQL & " AND B.PKId = " & dbcintUtilizacao.BoundText
    End If
    
    strQueryGrupo = strSQL
End Function

Private Sub dbcintUtilizacao_Click(Area As Integer)
    If Area = 2 And dbcintUtilizacao.MatchedWithList Then
        If mlngUltimo <> dbcintUtilizacao.BoundText Then
           dbcintGrupo.BoundText = ""
           dbcintSubGrupo.BoundText = ""
           Set dbcintSubGrupo.RowSource = Nothing
           Set dbcintGrupo.RowSource = Nothing
        End If
        mlngUltimo = dbcintUtilizacao.BoundText
    ElseIf Area = 0 Then
        DropDownDataCombo dbcintUtilizacao, Me, Area
    End If
End Sub

Private Sub dbcintUtilizacao_GotFocus()
    If Not mblnCarregando Then
        tab_3DPasta.Tab = 0
    End If
    mblnCarregando = False
End Sub

Private Sub dbcintUtilizacao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUtilizacao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUtilizacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintUtilizacao
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 615
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
    
    If tab_3DPasta.Tab = 1 Then HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
       Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
   bytOrdenacao = 1: blnOrdenacaoAsc = True
      
    mlngUltimo = 0
    mlngUltimoG = 0
    mlngUltimoS = 0
    mblnAlterando = False
    tab_3DPasta.TabEnabled(1) = False
    
    dbcintUtilizacao.Tag = strQueryDataComboUtilizacao & ";strNomeDaUtilizacao"
    dbcintGrupo.Tag = strQueryGrupo & ";A.strNomeDoGrupo"
    dbcintSubGrupo.Tag = strQuerySubGrupo & ";strNomeDoSubGrupo"
    dbcintatividadeBasica.Tag = "Select Pkid, " & gstrCONVERT(CDT_VARCHAR, "intcodigo") & strCONCAT & "' - '" & strCONCAT & " Ltrim(Rtrim(strDescricao)) As strDescricao From Tblatividadebasica" & ";strDescricao"
    PreencherListaDeOpcoes dbcintUtilizacao
    
    vetTipo(0) = "Percentual"
    vetTipo(1) = "Quantidade"
    vetTipo(2) = "Moeda"
    vetTipo(3) = "Fator"
    
    
    
    dbc_intTributoTipo.Tag = strQueryDataComboTributoTipo & ";strDescricao"
    dbc_intTributo.Tag = strQueryDataComboTributo & ";strDescricao"
      
    
    VerificaObjParaAplicar mobjAux
End Sub

Public Function strQueryDataComboUtilizacao()
    Dim strSQL As String
    '0 => Pessoa Física / Pessoa Jurídica
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strNomeDaUtilizacao "
    strSQL = strSQL & "FROM " & gstrUtilizacaoDaTabelaDeValor & " "
    strSQL = strSQL & "WHERE bitIdentificador = 0 "
    strSQL = strSQL & "ORDER BY strNomeDaUtilizacao"
    strQueryDataComboUtilizacao = strSQL
End Function

Public Function strQueryDataComboComposicaoReceita()
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrComposicaoDaReceita & " "
    strSQL = strSQL & "ORDER BY strDescricao"
    strQueryDataComboComposicaoReceita = strSQL
End Function

Public Function strQueryDataComboTributoTipo()
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrTributoTipo & " "
    strSQL = strSQL & "ORDER BY strDescricao"
    strQueryDataComboTributoTipo = strSQL
End Function

Public Function strQueryDataComboTributo()
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrTributo & " "
    strSQL = strSQL & "ORDER BY strDescricao"
    strQueryDataComboTributo = strSQL
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub lvw_Leis_Click()
    With lvw_Leis
       If .ListItems.Count > 0 Then
          blnClickLista = True
          dbc_intTributoTipo.Text = .SelectedItem.SubItems(1)
          PreencherListaDeOpcoes dbc_intTributoTipo, .SelectedItem.Tag
          mblnAlterandoLista = True
          TrocaCorObjeto lvw_Tributos, False
          TrocaCorObjeto cmd_Tributos, False
          TrocaCorObjeto dbc_intTributo, False
          PreencheGridTributoTrib
          If lvw_Tributos.ListItems.Count > 0 Then
            dbc_intTributo.Text = lvw_Tributos.SelectedItem.SubItems(1)
          End If
          If Not mblnCarregando Then
            dbc_intTributoTipo.SetFocus
            MarcaCampo dbc_intTributoTipo
          End If
       End If
    End With
End Sub

Private Sub lvw_Leis_KeyUp(KeyCode As Integer, Shift As Integer)
    mblnCarregando = True
    lvw_Leis_Click
    mblnCarregando = False
End Sub

Private Sub lvw_Tributos_Click()
    With lvw_Tributos
       If .ListItems.Count > 0 Then
          dbc_intTributo.Text = .SelectedItem.SubItems(1)
          PreencherListaDeOpcoes dbc_intTributo, .SelectedItem.Tag
          If Not mblnCarregando Then
            dbc_intTributo.SetFocus
            MarcaCampo dbc_intTributo
          End If
       End If
    End With

End Sub

Private Sub lvw_Tributos_GotFocus()
    tab_3DPasta.Tab = 1
End Sub

Private Sub lvw_Tributos_KeyUp(KeyCode As Integer, Shift As Integer)
    mblnCarregando = True
    lvw_Tributos_Click
    mblnCarregando = False
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3DPasta.Tab = 2 Then
       HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
       mblnCarregando = True
       dbcintUtilizacao.SetFocus
    Else
       HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
       mblnCarregando = True
       dbc_intTributoTipo.SetFocus
    End If
End Sub

Private Sub tdb_Atividades_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_Atividades_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Atividades
End Sub

Private Sub tdb_Atividades_HeadClick(ByVal ColIndex As Integer)
   blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
   bytOrdenacao = ColIndex
   gOrdenaGrid tdb_Atividades, ColIndex
End Sub

Private Sub tdb_Atividades_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Atividades
End Sub

Private Sub tdb_Atividades_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Atividades
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnCarregando = True
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrAtividadeEC, Me
                
                tab_3DPasta.TabEnabled(1) = True
                
                PreencheGridTributo
                LimpaListaAtividade
                lvw_Tributos.ListItems.Clear
                dbc_intTributo.Text = ""
                If lvw_Leis.ListItems.Count > 0 Then
                   lvw_Leis.ListItems.Item(1).Selected = True
                   lvw_Leis_Click
                End If
                   
                
                strCodigoAtual = txtintCodigo.Text
                strDescricaoAtual = txtstrDescricao.Text
                
                gCorLinhaSelecionada tdb_Atividades
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                
                mblnCarregando = False
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSQL          As String
Dim lngCodAtividade As Long
Dim blnAlterando    As Boolean
Dim blnExclusaoOK   As Boolean
        
    blnAlterando = mblnAlterando
    If blnAlterando Then
        lngCodAtividade = txtPKId
    End If
    
    strSQL = strQueryAtividadeSubGrupo
    
    If strModoOperacao = UCase("IMPRIMIR") Then
        ToolBarGeral strModoOperacao, gstrAtividadeEC, mblnAlterando, tdb_Atividades, Me, mobjAux, strSQL, , rptatividadeeconomica, strQueryAtividadeSubGrupo
        Exit Sub
    End If
   
    Select Case UCase(strModoOperacao)
        Case UCase(gstrNovo)
            If tab_3DPasta.Tab = 1 Then
               LimpaListaAtividade
               dbc_intTributoTipo.SetFocus
            Else
               NovaAtividade
               dbcintUtilizacao.SetFocus
            End If
        Case UCase(gstrSalvar)
            If blnDadosOk Then
                If ToolBarGeral(strModoOperacao, gstrAtividadeEC, mblnAlterando, tdb_Atividades, Me, mobjAux, strSQL, strQueryAplicar, , , False) Then
                    'GravaTributo
                    LimpaListaAtividade
                    NovaAtividade
                    LeDaTabelaParaObj "", tdb_Atividades, strSQL
                    dbcintUtilizacao.SetFocus
                End If
            End If
            
        Case UCase(gstrDeletar)
            If blnDeletaAtividade(lngCodAtividade) Then
               blnExclusaoOK = True
               LimpaListaAtividade
               NovaAtividade
               LeDaTabelaParaObj gstrAtividadeEC, tdb_Atividades, strQueryAtividadeSubGrupo
               dbcintUtilizacao.SetFocus
            End If
        
        Case UCase(gstrLocalizar), UCase(gstrPreencherLista)
            dbcintGrupo.Tag = strQueryGrupo & ";A.strNomeDoGrupo"
            dbcintSubGrupo.Tag = strQuerySubGrupo & ";strNomeDoSubGrupo"
            ToolBarGeral strModoOperacao, gstrAtividadeEC, mblnAlterando, tdb_Atividades, Me, mobjAux, strQueryAtividadeSubGrupo
            Exit Sub
        Case UCase(gstrFechar)
            Unload Me
    End Select
   
        
    If UCase(strModoOperacao) = gstrSalvar Or UCase(strModoOperacao) = gstrDeletar Then
        mblnPrimeiraVez = False
    End If
        
    If UCase(strModoOperacao) = UCase(gstrIncluirItem) Then
        If Me.ActiveControl.Name = dbc_intTributoTipo.Name Or Me.ActiveControl.Name = cmd_intTributoTipo.Name Then
            If blnDadosTributoTipoOk Then
                IncluirItemNoGrid
                Exit Sub
            End If
        ElseIf Me.ActiveControl.Name = dbc_intTributo.Name Or Me.ActiveControl.Name = cmd_Tributos.Name Then
            If blnDadosTributoOk Then
                IncluirItemNoGrid
                Exit Sub
            End If
        End If
        
    ElseIf UCase(strModoOperacao) = UCase(gstrExcluirItem) Then
        ExcluirItemNoGrid
        Exit Sub
    End If
        
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
End Sub

Private Function blnDadosOk() As Boolean
    Dim i As Integer
    Dim strSQL As String
    Dim adoRec As ADODB.Recordset
    
    If dbcintUtilizacao.MatchedWithList = False Then
        ExibeMensagem "O campo utilização tem que ser informado."
        dbcintUtilizacao.SetFocus
        Exit Function
    ElseIf dbcintGrupo.MatchedWithList = False Then
        ExibeMensagem "O campo grupo tem que ser informado."
        dbcintGrupo.SetFocus
        Exit Function
    ElseIf dbcintSubGrupo.MatchedWithList = False Then
        ExibeMensagem "O campo subGrupo tem que ser informado."
        dbcintSubGrupo.SetFocus
        Exit Function
    ElseIf Val(txtintCodigo) = 0 Then
        ExibeMensagem "O campo código tem que ser informado."
        txtintCodigo.SetFocus
        Exit Function
    ElseIf Trim(txtstrDescricao) = "" Then
        ExibeMensagem "O campo descrição tem que ser informado."
        txtstrDescricao.SetFocus
        Exit Function
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.Text)) Then
                
ProximoCodigo:
                
        If gblnExisteCodigo(1, gstrAtividadeEC, "intCodigo", "'" & txtintCodigo.Text & "'") Then
            strCodigo = (gstrProximoCodigo(txtintCodigo, gstrAtividadeEC, "intCodigo", gintCodSeguranca, , , , True))
            If Len(strCodigo) > 0 Then
                If MsgBox("O código informado já se encontra cadastrado. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                    txtintCodigo.SetFocus
                    Exit Function
                Else
                    txtintCodigo.Text = strCodigo
                    GoTo ProximoCodigo
                End If
            Else
                ExibeMensagem "O código informado já se encontra cadastrado."
                txtintCodigo.SetFocus
                Exit Function
            End If
        End If
    End If
                    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescricaoAtual)) Then
                
        If gblnExisteCodigo(1, gstrAtividadeEC, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrDescricao.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosOk = True

End Function

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Function strQueryAtividadeSubGrupo() As String
Dim strSQL  As String
   
   strSQL = ""
   
   strSQL = strSQL & "SELECT PKId, intCodigo, strDescricao "
   strSQL = strSQL & "FROM " & gstrAtividadeEC & " "

   Select Case bytOrdenacao
      
      Case Is = 1
         strSQL = strSQL & " ORDER BY intCodigo" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
         
      Case Is = 2
         strSQL = strSQL & " ORDER BY strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      
   End Select
   
   strQueryAtividadeSubGrupo = strSQL
   
End Function

Private Function strQuerySubGrupo() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strNomeDoSubGrupo "
    strSQL = strSQL & "FROM " & gstrSubGrupoDeAtividade & " "
    If dbcintGrupo.MatchedWithList Then
        strSQL = strSQL & "WHERE intCodigoDoGrupo = " & dbcintGrupo.BoundText
    End If
    strQuerySubGrupo = strSQL
End Function

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrObservacoes_GotFocus()
    tab_3DPasta.Tab = 0
End Sub

Private Sub txtstrObservacoes_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrObservacoes
End Sub

Private Function blnValorCadastrado(dblValor As Variant) As Boolean
    Dim i As Integer
    For i = 0 To y.Count(1) - 1
        If gvntConvVrDoSql(y(i, 1)) = gvntConvVrDoSql(dblValor) Then
            blnValorCadastrado = True
            Exit Function
        End If
    Next
    blnValorCadastrado = False
End Function

 Private Function blnDeletaAtividade(lngAtividade As Long) As Boolean
    Dim strSQL As String
    On Error GoTo err_blnDeletaAtividade
    If MsgBox("Confirma a exclusão da atividade de '" & Trim(txtstrDescricao) & "' e de todos os valores relacionados?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        blnDeletaAtividade = True
        Exit Function
    End If
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    strSQL = "DELETE FROM " & gstrAtividadeTributoTributo
    strSQL = strSQL & " WHERE intAtividadeTributo in (SELECT PKID FROM " & gstrAtividadeTributo
    strSQL = strSQL & " WHERE intAtividadeEc = " & lngAtividade & ")"
    If Not gobjBanco.Execute(strSQL) Then
        GoTo err_blnDeletaAtividade
    End If
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    strSQL = "DELETE FROM " & gstrAtividadeTributo
    strSQL = strSQL & " WHERE intAtividadeEc = " & lngAtividade
    If Not gobjBanco.Execute(strSQL) Then
        GoTo err_blnDeletaAtividade
    End If
    
    strSQL = "DELETE FROM " & gstrAtivEmpresaTributo & " "
    strSQL = strSQL & "WHERE intAtividadeDaEmpresa IN "
    strSQL = strSQL & "(SELECT DISTINCT pkID FROM " & gstrAtividadeDaEmpresa & " "
    strSQL = strSQL & "WHERE intAtividade = " & lngAtividade & ") "
    If Not gobjBanco.Execute(strSQL) Then
        GoTo err_blnDeletaAtividade
    End If
    
    strSQL = "DELETE FROM " & gstrAtividadeDaEmpresa
    strSQL = strSQL & " WHERE intAtividade = " & lngAtividade
    If Not gobjBanco.Execute(strSQL) Then
        GoTo err_blnDeletaAtividade
    End If
    
    strSQL = ""
    strSQL = strSQL & "DELETE FROM " & gstrAtividadeEC & " "
    strSQL = strSQL & "WHERE PKId = " & lngAtividade
    If Not gobjBanco.Execute(strSQL) Then
        GoTo err_blnDeletaAtividade
    End If
    gobjBanco.ExecutaCommitTrans
    blnDeletaAtividade = True
    Exit Function
err_blnDeletaAtividade:
    gobjBanco.ExecutaRollbackTrans
    ExibeMensagem "Não foi possível excluir a atividade."
End Function

Private Sub NovaAtividade()
    'LimpaObjeto Me
    dbcintUtilizacao.Text = ""
    dbcintGrupo.Text = ""
    dbcintSubGrupo.Text = ""
    
    
    txtPKId = ""
    txtintCodigo.Text = ""
    txtstrDescricao = ""
    txtstrObservacoes = ""
    tab_3DPasta.Tab = 0
    tab_3DPasta.TabEnabled(1) = False
    mblnAlterando = False
    mblnPrimeiraVez = False
    lvw_Leis.ListItems.Clear
    dbc_intTributoTipo = ""
    
    PreencheCodigo
End Sub

Private Sub PreencheCodigo()
Dim adoRec  As ADODB.Recordset

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT MAX (intCodigo) AS Codigo FROM " & gstrAtividadeEC, 10, adoRec) Then
        If Not (adoRec.EOF Or adoRec.BOF) Then txtintCodigo = gstrENulo(adoRec("Codigo") + 1)
    End If
    Set adoRec = Nothing: Set gobjBanco = Nothing
End Sub

Private Function blnDadosTributoTipoOk() As Boolean
    blnDadosTributoTipoOk = False
    If dbc_intTributoTipo.MatchedWithList = False Then
       ExibeMensagem "O Tipo de Tributo deve ser selecionado."
       dbc_intTributoTipo.SetFocus
       Exit Function
    End If
    
    blnDadosTributoTipoOk = True
End Function

Private Function blnDadosTributoOk() As Boolean
    blnDadosTributoOk = False
    If dbc_intTributo.MatchedWithList = False Then
       ExibeMensagem "O Tributo deve ser selecionado."
       dbc_intTributo.SetFocus
       Exit Function
    End If
    
    blnDadosTributoOk = True
End Function


Private Function blnDadosItens()
  
  If dbc_intTributoTipo.MatchedWithList = False Then
     ExibeMensagem "O Tipo do Tributo deve ser informado."
     dbc_intTributoTipo.SetFocus
     Exit Function
  End If
  blnDadosItens = True
End Function

Private Function IncluirItemNoGrid()
Dim intInd          As Integer
Dim strSQL          As String
Dim varAux

    If Me.ActiveControl.Name = dbc_intTributoTipo.Name Or Me.ActiveControl.Name = cmd_intTributoTipo.Name Then
        If blnDadosItens = False Then Exit Function
        With lvw_Leis
'            If mblnAlterandoLista Then
'               For intInd = 1 To .ListItems.Count
'                   If .SelectedItem.Index <> intInd And dbc_intTributoTipo.BoundText = .SelectedItem Then
'                       ExibeMensagem "Não é possível incluir tipos de Tributos iguais."
'                       dbc_intTributoTipo.SetFocus
'                       Exit Function
'                   End If
'               Next
'               varAux = .SelectedItem.Text
'               .SelectedItem.Text = dbc_intTributoTipo.BoundText
'               .SelectedItem.SubItems(1) = dbc_intTributoTipo.Text
'
'                strSQL = ""
'                strSQL = strSQL & "UPDATE "
'                strSQL = strSQL & gstrAtividadeTributo
'                strSQL = strSQL & " SET intAtividadeEc = " & txtPKId.Text & ", "
'                strSQL = strSQL & " intTributoTipo = " & dbc_intTributoTipo.BoundText
'                strSQL = strSQL & " WHERE  intAtividadeEc = " & txtPKId.Text
'                strSQL = strSQL & " AND intTributoTipo = " & varAux
'
'                Set gobjBanco = New clsBanco
'                gobjBanco.Execute strSQL
'
'            Else
               For intInd = 1 To .ListItems.Count
                   If dbc_intTributoTipo.BoundText = .ListItems(intInd).Tag Then
                      ExibeMensagem "Não é possível incluir tipos de Tributos iguais."
                      dbc_intTributoTipo.SetFocus
                      Exit Function
                   End If
               Next
               Set mobjLista = lvw_Leis.ListItems.Add(, , dbc_intTributoTipo.BoundText)
               mobjLista.SubItems(1) = dbc_intTributoTipo.Text
               
                strSQL = ""
                strSQL = strSQL & "INSERT INTO "
                strSQL = strSQL & gstrAtividadeTributo & " ("
                strSQL = strSQL & "intAtividadeEc, "
                strSQL = strSQL & "intTributoTipo, "
                strSQL = strSQL & "dtmDtAtualizacao, "
                strSQL = strSQL & "lngCodUsr) "
                strSQL = strSQL & "Values(" & txtPKId.Text & ","
                strSQL = strSQL & dbc_intTributoTipo.BoundText & ", "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema(, , False)) & ", "
                strSQL = strSQL & glngCodUsr & " "
                strSQL = strSQL & ")"
                
                Set gobjBanco = New clsBanco
                gobjBanco.Execute strSQL
                                
                PreencheGridTributo
                TrocaCorObjeto lvw_Tributos, True
                TrocaCorObjeto cmd_Tributos, True
                TrocaCorObjeto dbc_intTributo, True
                lvw_Tributos.ListItems.Clear
                dbc_intTributo = ""
        End With
        dbc_intTributoTipo.Text = ""
        mblnAlterandoLista = False
        LimpaListaAtividade
        dbc_intTributoTipo.SetFocus
        
    ElseIf Me.ActiveControl.Name = dbc_intTributo.Name Or Me.ActiveControl.Name = cmd_Tributos.Name Then
        With lvw_Tributos
'            If mblnAlterandoLista = 200 Then
'               For intInd = 1 To .ListItems.Count
'                   If .SelectedItem.Index <> intInd And dbc_intTributo.BoundText = .SelectedItem Then
'                       ExibeMensagem "Não é possível incluir Tributo iguais."
'                       dbc_intTributo.SetFocus
'                       Exit Function
'                   End If
'               Next
'               .SelectedItem.Text = dbc_intTributo.BoundText
'               .SelectedItem.SubItems(1) = dbc_intTributo.Text
'            Else
               For intInd = 1 To .ListItems.Count
                   If dbc_intTributo.BoundText = .ListItems(intInd).Tag Then
                      ExibeMensagem "Não é possível incluir Tributos iguais."
                      dbc_intTributo.SetFocus
                      Exit Function
                   End If
               Next
               Set mobjLista = .ListItems.Add(, , dbc_intTributo.BoundText)
               mobjLista.SubItems(1) = dbc_intTributo.Text
               
                strSQL = ""
                strSQL = strSQL & "INSERT INTO "
                strSQL = strSQL & gstrAtividadeTributoTributo & " ("
                strSQL = strSQL & "intAtividadeTributo, "
                strSQL = strSQL & "intTributo, "
                strSQL = strSQL & "dtmDtAtualizacao, "
                strSQL = strSQL & "lngCodUsr) "
                strSQL = strSQL & "Values(" & lvw_Leis.SelectedItem.Text & ","
                strSQL = strSQL & dbc_intTributo.BoundText & ", "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema(, , False)) & ", "
                strSQL = strSQL & glngCodUsr & " "
                strSQL = strSQL & ")"
                
                Set gobjBanco = New clsBanco
                gobjBanco.Execute strSQL
                
                PreencheGridTributoTrib
        End With
        dbc_intTributo.Text = ""
        mblnAlterandoLista = False
        LimpaListaAtividadeTrib
        dbc_intTributo.SetFocus
    End If
    
End Function

Private Function ExcluirItemNoGrid()
Dim strSQL As String


    If Me.ActiveControl.Name = dbc_intTributoTipo.Name Then
        With lvw_Leis
            
            strSQL = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
          
            strSQL = strSQL & " DELETE FROM " & gstrAtividadeTributoTributo
            strSQL = strSQL & " WHERE intAtividadeTributo = " & .SelectedItem.Text
            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", "")
            
            strSQL = strSQL & " DELETE FROM " & gstrAtividadeTributo
            strSQL = strSQL & " WHERE PKID = " & .SelectedItem.Text
           
            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; END; ", "")
            
            
            
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSQL
            
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
            
            TrocaCorObjeto lvw_Tributos, True
            TrocaCorObjeto cmd_Tributos, True
            TrocaCorObjeto dbc_intTributo, True
            lvw_Tributos.ListItems.Clear
            dbc_intTributo = ""
            
        End With
        mblnAlterandoLista = False
        dbc_intTributoTipo.Text = ""
        mblnAlterandoLista = False
        LimpaListaAtividade
        dbc_intTributoTipo.SetFocus
    ElseIf Me.ActiveControl.Name = dbc_intTributo.Name Then
        With lvw_Tributos
            
            strSQL = ""
            strSQL = strSQL & " DELETE FROM " & gstrAtividadeTributoTributo
            strSQL = strSQL & " WHERE PKID = " & .SelectedItem.Text
            
            
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSQL
        
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
        End With
        mblnAlterandoLista = False
        dbc_intTributo.Text = ""
        mblnAlterandoLista = False
        LimpaListaAtividadeTrib
        dbc_intTributo.SetFocus
    End If
End Function

Private Sub PreencheGridTributo()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT AT.PKID AttPKID, TT.pkID pkID_Tipo, TT.strDescricao TipoTributo "
    strSQL = strSQL & "FROM " & gstrAtividadeTributo & " AT, "
    strSQL = strSQL & gstrTributoTipo & " TT "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "AT.intAtividadeEc = " & txtPKId.Text & " AND "
    strSQL = strSQL & "TT.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " AT.intTributoTipo "
    strSQL = strSQL & "ORDER BY TT.strDescricao "
    
    lvw_Leis.ListItems.Clear
    dbc_intTributoTipo = ""
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If adoResultado.RecordCount >= 1 Then
            Do While Not adoResultado.EOF
                Set mobjLista = lvw_Leis.ListItems.Add(, , adoResultado!AttPKID)
                                                        ', gstrENulo(adoResultado!pkID_Tipo)
                mobjLista.Tag = gstrENulo(adoResultado!pkID_Tipo)
                mobjLista.SubItems(1) = gstrENulo(adoResultado!TipoTributo)
                adoResultado.MoveNext
            Loop
        End If
    End If
    
    Set gobjBanco = Nothing

End Sub

Private Sub PreencheGridTributoTrib()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT TT.PKID pkID_Att, TB.pkID pkID_Trib, AT.PKID pkID_AtTrib, TB.strDescricao strTributo "
    strSQL = strSQL & "FROM " & gstrAtividadeTributo & " AT, "
    strSQL = strSQL & gstrTributo & " TB, "
    strSQL = strSQL & gstrAtividadeTributoTributo & " TT "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "AT.PKID = " & lvw_Leis.SelectedItem.Text & " AND "
    strSQL = strSQL & "TB.PKID = TT.intTributo AND "
    strSQL = strSQL & "AT.PKID = TT.intAtividadeTributo "
    strSQL = strSQL & "ORDER BY TB.strDescricao "
    
    lvw_Tributos.ListItems.Clear
    dbc_intTributo = ""
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If adoResultado.RecordCount >= 1 Then
            Do While Not adoResultado.EOF
                Set mobjLista = lvw_Tributos.ListItems.Add(, , adoResultado!PKID_att)
                                                        ', gstrENulo(adoResultado!pkID_Tipo)
                mobjLista.Tag = gstrENulo(adoResultado!pkID_Trib)
                mobjLista.SubItems(1) = gstrENulo(adoResultado!strTributo)
                adoResultado.MoveNext
            Loop
        End If
    End If
    
    Set gobjBanco = Nothing

End Sub

Private Sub GravaTributo()
Dim strSQL  As String
Dim intInd As Integer
        
    If Len(txtPKId.Text) = 0 Then Exit Sub
    
    strSQL = ""
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    strSQL = strSQL & "DELETE FROM " & gstrAtividadeTributo
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " intAtividadeEc = " & Val(txtPKId.Text) & IIf(bytDBType = Oracle, ";", "")
    
    With lvw_Leis
        If .ListItems.Count >= 1 Then
            For intInd = 1 To .ListItems.Count
                strSQL = strSQL & "INSERT INTO "
                strSQL = strSQL & gstrAtividadeTributo & " ("
                strSQL = strSQL & "intAtividadeEc, "
                strSQL = strSQL & "intTributoTipo, "
                strSQL = strSQL & "dtmDtAtualizacao, "
                strSQL = strSQL & "lngCodUsr) "
                strSQL = strSQL & "Values(" & txtPKId.Text & ","
                strSQL = strSQL & .ListItems(intInd).Text & ", "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & glngCodUsr & " "
                strSQL = strSQL & ")" & IIf(bytDBType = Oracle, ";", "")
            Next
        End If
    End With
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
        
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute strSQL
    
    Set gobjBanco = Nothing
    
End Sub

Private Sub LimpaListaAtividade()
    Set dbc_intTributoTipo.RowSource = Nothing
    dbc_intTributoTipo.Text = ""
    mblnAlterandoLista = False
    TrocaCorObjeto lvw_Tributos, True
    TrocaCorObjeto cmd_Tributos, True
    TrocaCorObjeto dbc_intTributo, True
End Sub

Private Sub LimpaListaAtividadeTrib()
    Set dbc_intTributo.RowSource = Nothing
    dbc_intTributo.Text = ""
    mblnAlterandoLista = False
End Sub

