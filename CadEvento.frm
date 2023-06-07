VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadEvento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eventos"
   ClientHeight    =   6345
   ClientLeft      =   1170
   ClientTop       =   2250
   ClientWidth     =   9780
   HelpContextID   =   42
   Icon            =   "CadEvento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6195
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   10927
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Eventos"
      TabPicture(0)   =   "CadEvento.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrCodigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_strTipoEvento"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintExercicio"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_Lista"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtstrDescricao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrCodigo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tab_3dPastaConta"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbo_strTipodeEvento"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtintExercicio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.TextBox txtintExercicio 
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
         Left            =   8730
         MaxLength       =   4
         TabIndex        =   8
         Top             =   615
         Width           =   675
      End
      Begin VB.ComboBox cbo_strTipodeEvento 
         Height          =   315
         ItemData        =   "CadEvento.frx":105E
         Left            =   6255
         List            =   "CadEvento.frx":108C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2295
      End
      Begin TabDlg.SSTab tab_3dPastaConta 
         Height          =   2685
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   990
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   4736
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Debitar"
         TabPicture(0)   =   "CadEvento.frx":1170
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl_ContaDebito"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lvw_ContaDebito"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cbointContaDebito"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmd_ContaDebito"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cbostrContaDebito"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkContaGrupoDebitar"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Creditar"
         TabPicture(1)   =   "CadEvento.frx":118C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lbl_ContaCredito"
         Tab(1).Control(1)=   "lvw_ContaCredito"
         Tab(1).Control(2)=   "cmd_ContaCredito"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cbostrContaCredito"
         Tab(1).Control(4)=   "cbointContaCredito"
         Tab(1).Control(5)=   "chkContaGrupoCreditar"
         Tab(1).ControlCount=   6
         Begin VB.CheckBox chkContaGrupoCreditar 
            Caption         =   "Conta de Grupo"
            Height          =   225
            Left            =   -67230
            TabIndex        =   23
            Top             =   440
            Width           =   1425
         End
         Begin VB.CheckBox chkContaGrupoDebitar 
            Caption         =   "Conta de Grupo"
            Height          =   225
            Left            =   7770
            TabIndex        =   22
            Top             =   440
            Width           =   1425
         End
         Begin VB.ComboBox cbointContaCredito 
            Height          =   315
            Left            =   -74400
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   390
            Width           =   1575
         End
         Begin VB.ComboBox cbostrContaCredito 
            Height          =   315
            Left            =   -72870
            TabIndex        =   17
            Top             =   390
            Width           =   5145
         End
         Begin VB.CommandButton cmd_ContaCredito 
            Height          =   300
            Left            =   -67695
            Picture         =   "CadEvento.frx":11A8
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Tag             =   "322"
            ToolTipText     =   "Clique para cadastar conta"
            Top             =   390
            Width           =   330
         End
         Begin VB.ComboBox cbostrContaDebito 
            Height          =   315
            Left            =   2130
            TabIndex        =   12
            Top             =   390
            Width           =   5145
         End
         Begin VB.CommandButton cmd_ContaDebito 
            Height          =   300
            Left            =   7305
            Picture         =   "CadEvento.frx":1532
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Tag             =   "322"
            ToolTipText     =   "Clique para cadastar conta"
            Top             =   390
            Width           =   330
         End
         Begin VB.ComboBox cbointContaDebito 
            Height          =   315
            Left            =   600
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   390
            Width           =   1575
         End
         Begin MSComctlLib.ListView lvw_ContaDebito 
            Height          =   1785
            Left            =   60
            TabIndex        =   14
            Top             =   780
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   3149
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
               Text            =   "Conta"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descrição"
               Object.Width           =   10849
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Conta de Grupo"
               Object.Width           =   2293
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_ContaCredito 
            Height          =   1785
            Left            =   -74940
            TabIndex        =   19
            Top             =   780
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   3149
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
               Text            =   "Conta"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descrição"
               Object.Width           =   10848
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Conta de Grupo"
               Object.Width           =   2293
            EndProperty
         End
         Begin VB.Label lbl_ContaCredito 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   -74895
            TabIndex        =   15
            Top             =   510
            Width           =   420
         End
         Begin VB.Label lbl_ContaDebito 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   105
            TabIndex        =   10
            Top             =   510
            Width           =   420
         End
      End
      Begin VB.TextBox txtstrCodigo 
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
         Left            =   210
         MaxLength       =   10
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         TabIndex        =   2
         Top             =   615
         Width           =   870
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
         Height          =   285
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   4
         Top             =   615
         Width           =   4770
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2295
         Left            =   120
         TabIndex        =   20
         Top             =   3780
         Width           =   9345
         _ExtentX        =   16484
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "strCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tipo de Evento"
         Columns(3).DataField=   "intTipoEvento"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Exercício"
         Columns(4).DataField=   "intExercicio"
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
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2434"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2355"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=8096"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=8017"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=3678"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3598"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=1376"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1296"
         Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
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
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
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
      Begin VB.Label lblintExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   8730
         TabIndex        =   7
         Top             =   375
         Width           =   675
      End
      Begin VB.Label lbl_strTipoEvento 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Evento"
         Height          =   195
         Left            =   6270
         TabIndex        =   5
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   375
         Width           =   720
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   375
         Width           =   495
      End
   End
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   5925
      TabIndex        =   21
      Top             =   255
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSComctlLib.ImageList img_Arquivo 
      Left            =   4605
      Top             =   225
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
            Picture         =   "CadEvento.frx":18BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":191A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1978
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":19D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1B4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img_ArquivoD 
      Left            =   3765
      Top             =   225
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
            Picture         =   "CadEvento.frx":1BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadEvento.frx":1E3E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando               As Boolean
Dim mobjAux                     As Object
Dim mobjLista                   As Object
Dim mblnSelecionou              As Boolean
Dim mblnAlterandoContaDebito    As Boolean
Dim mblnAlterandoContaCredito   As Boolean
Dim mblnClickOk                 As Boolean
Dim mvtContaCredito()           As Integer
Dim mvtContaDebito()            As Integer
Dim mblnCarregaFormConta        As Boolean
Dim mblbCarregaPrimeiraVez      As Boolean
Dim mblnPrimeiraVez             As Boolean

Private Sub LeTabelaEvento()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT strCodigo, strDescricao, intTipoEvento, intExercicio "
    strSQL = strSQL & "FROM " & gstrEvento & " "
    strSQL = strSQL & "WHERE PKId = " & Val(txtPKId)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txtstrCodigo = Trim(!strCodigo)
                txtstrDescricao = Trim(!strDescricao)
                txtintExercicio = Space$(0) & Trim(!intExercicio)
                If IsNull(!intTipoEvento) Then
                    cbo_strTipodeEvento.ListIndex = -1
                Else
                    cbo_strTipodeEvento.ListIndex = Val(!intTipoEvento)
                End If
                
            End If
            .Close
        End With
    End If
    LeTabelaContaDoEvento gstrEventoContaContabilCredito, lvw_ContaCredito, Val(txtPKId)
    LeTabelaContaDoEvento gstrEventoContaContabilDebito, lvw_ContaDebito, Val(txtPKId)
End Sub

Private Sub cbo_strTipodeEvento_Click()
    If cbo_strTipodeEvento.ListIndex = 4 Or cbo_strTipodeEvento.ListIndex = 7 Or cbo_strTipodeEvento.ListIndex = 12 Then
        TrocaCorObjeto txtintExercicio, False
    Else
        txtintExercicio = ""
        TrocaCorObjeto txtintExercicio, True
    End If
End Sub

Private Sub cbointContaCredito_Click()
    cbostrContaCredito.ListIndex = gintIndiceCBO(cbostrContaCredito, _
                              gstrItemData(cbointContaCredito))
    'glIgualaContas cbointContaCredito, cbostrContaCredito, _
                   lvw_ContaCredito, mblnAlterandoContaCredito
End Sub

Private Sub cbointContaCredito_GotFocus()
    AtivaPastaDeObjeto tab_3dPastaConta, 1
End Sub

Private Sub cbointContaCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cbointContaCredito
End Sub

Private Sub cbointContaCredito_LostFocus()
    Dim i                   As Integer
    Dim strGuardaValorCombo As String

    If InStr(1, cbointContaCredito.Text, ".") = 0 Or cbointContaCredito.ListIndex = -1 Then
        cbointContaCredito.Text = Replace(cbointContaCredito.Text, ".", "")
        cbointContaCredito.Text = gvntFormatacaoEspecifica(cbointContaCredito.Text, 1)
        For i = 0 To cbointContaCredito.ListCount - 1
            If cbointContaCredito.Text = cbointContaCredito.list(i) Then
                cbointContaCredito.ListIndex = i
                Exit Sub
            End If
        Next
        strGuardaValorCombo = cbointContaCredito.Text
        cbostrContaCredito.ListIndex = -1
        cbointContaCredito.Text = strGuardaValorCombo
    End If
End Sub

Private Sub cbointContaDebito_Click()
       cbostrContaDebito.ListIndex = gintIndiceCBO(cbostrContaDebito, _
                              gstrItemData(cbointContaDebito))
    
    'glIgualaContas cbointContaDebito, cbostrContaDebito, _
                   lvw_ContaDebito, mblnAlterandoContaDebito
End Sub

Private Sub cbointContaDebito_GotFocus()
    AtivaPastaDeObjeto tab_3dPastaConta, 0
End Sub

Private Sub cbointContaDebito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cbointContaDebito
End Sub

Private Sub cbointContaDebito_LostFocus()
    Dim i                   As Integer
    Dim strGuardaValorCombo As String

    If InStr(1, cbointContaDebito.Text, ".") = 0 Or cbointContaDebito.ListIndex = -1 Then
        cbointContaDebito.Text = Replace(cbointContaDebito.Text, ".", "")
        cbointContaDebito.Text = gvntFormatacaoEspecifica(cbointContaDebito.Text, 1)
        For i = 0 To cbointContaDebito.ListCount - 1
            If cbointContaDebito.Text = cbointContaDebito.list(i) Then
                cbointContaDebito.ListIndex = i
                Exit Sub
            End If
        Next
        strGuardaValorCombo = cbointContaDebito.Text
        cbostrContaDebito.ListIndex = -1
        cbointContaDebito.Text = strGuardaValorCombo
    End If
End Sub

Private Sub cbostrContaCredito_Click()
    Dim tempIndice As String
    Dim tempIndice1 As Integer
                             
   If gintIndiceCBO(cbointContaCredito, gstrItemData(cbostrContaCredito)) = -1 Then
        tempIndice = gstrItemData(cbostrContaCredito)
        tempIndice1 = gstrItemData(cbostrContaDebito)
        
        LePlanoContaGeral cbointContaDebito, cbostrContaDebito, "PA"
        LePlanoContaGeral cbointContaCredito, cbostrContaCredito, "PA"
        
        cbostrContaCredito.ListIndex = gintIndiceCBO(cbostrContaCredito, tempIndice)
        cbostrContaDebito.ListIndex = gintIndiceCBO(cbostrContaDebito, tempIndice1)
   End If
   
   cbointContaCredito.ListIndex = gintIndiceCBO(cbointContaCredito, _
                           gstrItemData(cbostrContaCredito))
 
    'glIgualaContas cbostrContaCredito, cbointContaCredito, _
                   lvw_ContaCredito, mblnAlterandoContaCredito
End Sub

Private Sub cbostrContaCredito_GotFocus()
    If mblnCarregaFormConta = True Then
        mblnCarregaFormConta = False
        If cbostrContaCredito.ListIndex = -1 Then cbointContaCredito.ListIndex = -1
    End If
End Sub

Private Sub cbostrContaCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbostrContaDebito_Click()
    Dim tempIndice As Integer
    Dim tempIndice1 As Integer
                             
   If gintIndiceCBO(cbointContaDebito, gstrItemData(cbostrContaDebito)) = -1 Then
        tempIndice = gstrItemData(cbostrContaDebito)
        tempIndice1 = gstrItemData(cbostrContaCredito)
        
        LePlanoContaGeral cbointContaDebito, cbostrContaDebito, "PA"
        LePlanoContaGeral cbointContaCredito, cbostrContaCredito, "PA"
        
        cbostrContaDebito.ListIndex = gintIndiceCBO(cbostrContaDebito, tempIndice)
        cbostrContaCredito.ListIndex = gintIndiceCBO(cbostrContaCredito, tempIndice1)
        
   End If
   
   cbointContaDebito.ListIndex = gintIndiceCBO(cbointContaDebito, _
                           gstrItemData(cbostrContaDebito))
   
    'glIgualaContas cbostrContaDebito, cbointContaDebito, _
                   lvw_ContaDebito, mblnAlterandoContaDebito
End Sub

Private Sub cbostrContaDebito_GotFocus()
    If mblnCarregaFormConta = True Then
        mblnCarregaFormConta = False
        If cbointContaDebito.ListIndex = -1 Then cbostrContaDebito.ListIndex = -1
    End If
End Sub

Private Sub cbostrContaDebito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cbostrContaDebito
End Sub

Private Sub cmd_ContaCredito_Click()
    mblnCarregaFormConta = True
    CarregaForm frmCadPlanoConta, cbostrContaCredito, strQueryPlanoConta
'    CarregaForm frmCadPlanoConta, cbostrContaCredito
End Sub

Private Sub cmd_ContaDebito_Click()
    mblnCarregaFormConta = True
    CarregaForm frmCadPlanoConta, cbostrContaDebito, strQueryPlanoConta
    'CarregaForm frmCadPlanoConta, cbostrContaDebito
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 868
    VirificaGradeListView Me
    'If blnVerificaMovi Then
        'HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir, gstrIncluirItem, gstrExcluirItem
        'HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    'Else
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, _
                             gstrIncluirItem, gstrExcluirItem
    'End If
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

Private Sub lvw_ContaCredito_Click()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, _
                             gstrIncluirItem, gstrExcluirItem
End Sub

Private Sub lvw_ContaCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub lvw_ContaDebito_Click()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, _
                             gstrIncluirItem, gstrExcluirItem

End Sub

Private Sub lvw_ContaDebito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Lista
End Sub

Private Sub LimpaTela()
    txtstrCodigo = ""
    txtstrDescricao = ""
    txtintExercicio = ""
    cbointContaDebito.ListIndex = -1
    lvw_ContaDebito.ListItems.Clear
    cbointContaCredito.ListIndex = -1
    lvw_ContaCredito.ListItems.Clear
    proximoCodigoEvento
    If txtstrCodigo.Enabled Then txtstrCodigo.SetFocus
End Sub

Private Sub GravaEvento()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim intEvento       As Integer
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    If blnDadosOk Then
'        If gblnExclusaoGravacaoOk("I", "Gravação do evento " & Trim(txtstrDescricao)) Then
        If gblnExclusaoGravacaoOk(IIf(mblnAlterando, "A", "I"), "do evento") Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            strSQL = ""
            If mblnAlterando Then
                strSQL = strSQL & "UPDATE " & gstrEvento & " SET "
                strSQL = strSQL & "strCodigo = '" & Trim(txtstrCodigo) & "', "
                strSQL = strSQL & "strDescricao = '" & Trim(txtstrDescricao) & "', "
                strSQL = strSQL & "intExercicio = " & IIf(Trim(txtintExercicio) = "", "NULL", Trim(txtintExercicio)) & ", "
                strSQL = strSQL & "dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
                strSQL = strSQL & "WHERE PKId = " & Val(txtPKId)
                Set gobjBanco = New clsBanco
                If gobjBanco.Execute(strSQL) Then
                    intEvento = Val(txtPKId)
                Else
                    intEvento = 0
                End If
            Else
'                strSql = strSql & "sp_GravaEventoContabil "
'                strSql = strSql & "'" & Trim(txtstrCodigo) & "', "
'                strSql = strSql & "'" & Trim(txtstrDescricao) & " ', "
'                strSql = strSql & glngCodUsr
                strSQL = strSQL & gstrStoredProcedure("sp_GravaEventoContabil", _
                    "'" & Trim(txtstrCodigo) & "', " & _
                    "'" & Trim(txtstrDescricao) & " ', " & _
                    glngCodUsr & ", " & cbo_strTipodeEvento.ListIndex & ", " & IIf(Trim(txtintExercicio) = "", "NULL", txtintExercicio), True)
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    With adoResultado
                        If .EOF = False Then
                            intEvento = gstrENulo(!intEvento)
                        End If
                    End With
                End If
            End If
            If intEvento = 0 Then
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
            ElseIf (blnGravouConta(intEvento, lvw_ContaDebito, _
                                   gstrEventoContaContabilDebito, mvtContaDebito) And _
                    blnGravouConta(intEvento, lvw_ContaCredito, _
                                   gstrEventoContaContabilCredito, mvtContaCredito)) = False Then
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaCommitTrans
                LeDaTabelaParaObj "", tdb_Lista, strQuery
                'LimpaTela
                MantemForm gstrNovo
            End If
        End If
    End If
End Sub

Private Function blnGravouConta(intEvento As Integer, _
                                objLista As ListView, _
                                strTabela As String, _
                                mvtExcluido() As Integer) As Boolean

'******************************************************************************************
' Data: 12/06/2003
' Alteração: - Incluídos os nomes das colunas no comando INSERT.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim intInd      As Integer
    Dim strSQL      As String
    strSQL = ""
    strSQL = strSQL & "DELETE " & strTabela & " "
    strSQL = strSQL & "WHERE intEvento = " & intEvento
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSQL) = False Then
        Exit Function
    End If
    With objLista
        If .ListItems.Count = 0 Then
            blnGravouConta = True
        Else
            For intInd = 1 To .ListItems.Count
                strSQL = ""
'                strSQL = strSQL & "INSERT INTO " & strTabela & " VALUES ("
                strSQL = strSQL & "INSERT INTO " & strTabela
                
                strSQL = strSQL & " (intEvento, intContaContabil , bytContaGrupo , dtmDtAtualizacao, lngCodUsr) "
                strSQL = strSQL & " VALUES ("
                
                strSQL = strSQL & intEvento & ", " & .ListItems(intInd).Tag & ", "
                strSQL = strSQL & IIf(.ListItems(intInd).SubItems(2) = "Sim", 1, 0) & ", "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & glngCodUsr & ")"
                Set gobjBanco = New clsBanco
                If gobjBanco.Execute(strSQL) Then
                    blnGravouConta = True
                End If
            Next
        End If
    End With
End Function

Private Function blnDadosContaOk(cboConta As ComboBox) As Boolean
    If cboConta.ListIndex = -1 Then
        ExibeMensagem "A conta não foi informada corretamente."
        cboConta.SetFocus
    Else
        blnDadosContaOk = True
    End If
End Function

Private Function blnDadosOk() As Boolean
    If Trim(txtstrCodigo) = "" Then
        ExibeMensagem "O código não foi informado corretamente."
        txtstrCodigo.SetFocus
    ElseIf Trim(txtstrDescricao) = "" Then
        ExibeMensagem "A descrição não foi informada corretamente."
        txtstrDescricao.SetFocus
    ElseIf lvw_ContaDebito.ListItems.Count = 0 And lvw_ContaCredito.ListItems.Count = 0 Then
        ExibeMensagem "Não há nenhuma conta a ser cadastrada para este evento."
    ElseIf cbo_strTipodeEvento.ListIndex = -1 Then
        ExibeMensagem "É necessário informar o Tipo do Evento."
         cbo_strTipodeEvento.SetFocus
    ElseIf (cbo_strTipodeEvento.ListIndex = 4 Or cbo_strTipodeEvento.ListIndex = 7 Or cbo_strTipodeEvento.ListIndex = 12) And Trim(txtintExercicio.Text) = "" Then
        ExibeMensagem "É necessário informar o Exercício."
        txtintExercicio.SetFocus
     Else
        blnDadosOk = True
    End If
End Function

Private Sub IncluiItemNaLista(cboConta As ComboBox, _
                              cboDescricao As ComboBox, _
                              lvw_Lista As ListView, _
                              blnAlterando As Boolean, _
                              chkContaGrupo As CheckBox)
                              
    If blnDadosContaOk(cboDescricao) Then
       
       If cbo_strTipodeEvento.ListIndex = -1 Then
          ExibeMensagem "É necessário informar o tipo do evento antes de inserir este item."
          cbo_strTipodeEvento.SetFocus
          Exit Sub
       End If
       
       If Not VerificaItemNaLista(cbo_strTipodeEvento.ListIndex, lvw_Lista, cboConta) Then
          If blnAlterando Then
             lvw_Lista.SelectedItem.Text = cboConta.Text
             lvw_Lista.SelectedItem.SubItems(1) = cboDescricao.Text
             lvw_Lista.SelectedItem.SubItems(2) = IIf(chkContaGrupo.Value = 1, "Sim", "Não")
             lvw_Lista.Tag = gstrItemData(cboDescricao)
          Else
             Set mobjLista = lvw_Lista.ListItems.Add(, , cboConta.Text)
             mobjLista.SubItems(1) = cboDescricao.Text
             mobjLista.SubItems(2) = IIf(chkContaGrupo.Value = 1, "Sim", "Não")
             mobjLista.Tag = gstrItemData(cboDescricao)
          End If
       End If
       LimpaDados cboConta, cboDescricao, blnAlterando, chkContaGrupo
    End If
End Sub

Sub LimpaDados(cboDescricao As ComboBox, _
               cboConta As ComboBox, _
               blnAlterando As Boolean, _
               chkContaGrupo As CheckBox)
    cboDescricao.ListIndex = -1
    blnAlterando = False
    cboDescricao.SetFocus
    chkContaGrupo.Value = 0
End Sub

Private Sub VerificaListaAExcluir()
    Select Case tab_3dPastaConta.Tab
    Case 0 'Débito
        ExcluirItemDaLista lvw_ContaDebito, _
                           cbointContaDebito, _
                           mblnAlterandoContaDebito, _
                           mvtContaCredito()
    Case 1 'Crédito
        ExcluirItemDaLista lvw_ContaCredito, _
                           cbointContaCredito, _
                           mblnAlterandoContaCredito, _
                           mvtContaDebito()
    End Select
End Sub

Private Sub ExcluirItemDaLista(lvw_Lista As ListView, _
                               cboPrograma As ComboBox, _
                               blnAlterando As Boolean, _
                               mvtExcluido() As Integer)
    With lvw_Lista
        If .ListItems.Count > 0 Then
            ReDim Preserve mvtExcluido(UBound(mvtExcluido) + 1)
            mvtExcluido(UBound(mvtExcluido)) = .ListItems(.SelectedItem.Index).Tag
            .ListItems.Remove .SelectedItem.Index
            cboPrograma.ListIndex = -1
            blnAlterando = False
        End If
    End With
End Sub

Private Sub VerificaListaAIncluir()
    Select Case tab_3dPastaConta.Tab
    Case 0 'Débito
        IncluiItemNaLista cbointContaDebito, _
                          cbostrContaDebito, _
                          lvw_ContaDebito, _
                          mblnAlterandoContaDebito, _
                          chkContaGrupoDebitar
    Case 1 'Crédito
        IncluiItemNaLista cbointContaCredito, _
                          cbostrContaCredito, _
                          lvw_ContaCredito, _
                          mblnAlterandoContaCredito, _
                          chkContaGrupoCreditar
    End Select
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
    
        Case UCase(gstrNovo)
            ToolBarGeral strModoOperacao, gstrEvento, _
                         mblnAlterando, tdb_Lista, Me, _
                         mobjAux, strQuery, , , , True
            lvw_ContaCredito.ListItems.Clear
            lvw_ContaDebito.ListItems.Clear
            chkContaGrupoDebitar.Value = 0
            chkContaGrupoCreditar.Value = 0
            HabilitaDesabilitaCodigo True
            proximoCodigoEvento
            mblnAlterando = False
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrIncluirItem, gstrExcluirItem
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
        Case UCase(gstrSalvar)
            GravaEvento
            
        Case UCase(gstrIncluirItem)
            VerificaListaAIncluir
            
        Case UCase(gstrExcluirItem)
            VerificaListaAExcluir
            
        Case UCase(gstrDeletar)
            DeletaEvento
            
        Case UCase(gstrFechar)
            Unload Me
        
        Case UCase(gstrLocalizar)
                ToolBarGeral strModoOperacao, gstrEvento, _
                             mblnAlterando, tdb_Lista, Me, _
                             mobjAux, strQuery, , , , True
        
        Case UCase(gstrPreencherLista)
            
            If cbointContaDebito.Name = ActiveControl.Name Or cbostrContaDebito.Name = ActiveControl.Name Or _
               cbointContaCredito = ActiveControl.Name Or cbostrContaCredito.Name = ActiveControl.Name Then
               
                LePlanoContaGeral cbointContaDebito, cbostrContaDebito, "PA"
                LePlanoContaGeral cbointContaCredito, cbostrContaCredito, "PA"
            
                
            End If
            
        Case UCase(gstrAplicar)
            ToolBarGeral strModoOperacao, gstrEvento, _
                         mblnAlterando, tdb_Lista, Me, _
                         mobjAux, strQuery, , , , True
    End Select
    
End Sub

Private Function blnDeletouConta(strTabela As String) As Boolean
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "DELETE " & strTabela & " "
    strSQL = strSQL & "WHERE intEvento = " & Val(txtPKId)
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSQL) Then
        blnDeletouConta = True
    End If
End Function

Private Sub DeletaEvento()
    Dim strSQL  As String
    
    If EventoJaUsado = True Then
        ExibeMensagem "Este evento está sendo usado em outros lançamentos e não pode ser excluído."
        Exit Sub
    End If
    
    If gblnExclusaoGravacaoOk("E", "do evento " & Trim(tdb_Lista.Columns("strDescricao"))) Then
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        If blnDeletouConta(gstrEventoContaContabilCredito) And _
           blnDeletouConta(gstrEventoContaContabilDebito) Then
            strSQL = ""
            strSQL = strSQL & "DELETE " & gstrEvento & " "
            strSQL = strSQL & "WHERE PKId = " & Val(txtPKId)
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSQL) Then
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaCommitTrans
                VerificaListaAutomatica gstrEvento, tdb_Lista, strQuery
                LimpaTela
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
            End If
        Else
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaRollbackTrans
        End If
    End If
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    VerificaObjParaAplicar mobjAux
    VerificaListaAutomatica gstrEvento, tdb_Lista, strQuery
    LePlanoContaGeral cbointContaDebito, cbostrContaDebito, "PA"
    LePlanoContaGeral cbointContaCredito, cbostrContaCredito, "PA"
                       
    ReDim mvtContaCredito(0)
    ReDim mvtContaDebito(0)
    mblbCarregaPrimeiraVez = True
    MantemForm gstrNovo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblbCarregaPrimeiraVez = False
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            If mblnPrimeiraVez Then
                mblnClickOk = False
                txtPKId = .Columns("PKID").Value
                gCorLinhaSelecionada tdb_Lista
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnAlterando = True
                LeTabelaEvento
'                If blnVerificaMovi Then
'                    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir, gstrIncluirItem, gstrExcluirItem
'                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
'                End If
                
                HabilitaDesabilitaCodigo False
            End If
        End If
    End With
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
    If txtintExercicio = "" Then txtintExercicio = Year(gstrDataDoSistema)
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintExercicio_LostFocus()
   If Len(txtintExercicio) > 0 Then txtintExercicio = Format("01/01/" & txtintExercicio, "yyyy")
End Sub

Private Sub txtstrCodigo_LostFocus()
   BuscaCodigoEvento
   'If mblnAlterando Then
    'If blnVerificaMovi Then
    '     HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir, gstrIncluirItem, gstrExcluirItem
    '     HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    'End If
   'End If
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Sub txtstrCodigo_GotFocus()
    If mblbCarregaPrimeiraVez Then
        proximoCodigoEvento
        mblbCarregaPrimeiraVez = False
    End If
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCodigo
End Sub

Private Function strQuery() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao, strCodigo, intExercicio,"
    strSQL = strSQL & gstrCASEWHEN("intTipoEvento", _
        "0, 'Orçamento', 1, 'Arrecadação', 2, 'Empenho', 3, 'Pagto.Empenho', 4, 'Pagto.Restos a Pagar', 5, 'Pagto.Extra', 6, 'Alterações Orçamentárias', 7, 'Liquidação', 8, 'Transferências', 9, 'Modalidade', 10, 'Adiantamentos', 11, 'Anulação De Receita', 12, 'Cancelamento de Restos à Pagar', 13, 'Fiança Bancária'") & " AS intTipoEvento "
    strSQL = strSQL & " FROM " & gstrEvento & " ORDER BY " & gstrCONVERT(CDT_INT, "strCodigo")
    strQuery = strSQL
End Function

Function strQueryPlanoConta() As String
    Dim strSQL          As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " "
    strSQL = strSQL & "WHERE ABS(blnAnalitica) = 1"
    strQueryPlanoConta = strSQL
End Function

Private Sub BuscaCodigoEvento()
    
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
        
    If Len(Trim(txtstrCodigo)) > 0 Then
      strSQL = ""
      strSQL = strSQL & "SELECT PKID, strCodigo, strDescricao, intTipoEvento, intExercicio "
      strSQL = strSQL & "FROM " & gstrEvento & " "
      strSQL = strSQL & "WHERE strCodigo = '" & Trim(txtstrCodigo) & "'"
      Set gobjBanco = New clsBanco
      If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
          With adoResultado
              If .EOF = False Then
                  txtPKId = (!Pkid)
                  txtstrCodigo = Trim(!strCodigo)
                  txtstrDescricao = Trim(!strDescricao)
                  txtintExercicio = Space$(0) & Trim(!intExercicio)
                  cbo_strTipodeEvento.ListIndex = Val(!intTipoEvento)
                  
                  LeTabelaContaDoEvento gstrEventoContaContabilCredito, lvw_ContaCredito, Val(txtPKId)
                  LeTabelaContaDoEvento gstrEventoContaContabilDebito, lvw_ContaDebito, Val(txtPKId)
                  mblnAlterando = True
              End If
              .Close
          End With
      End If

   End If
End Sub

Private Sub HabilitaDesabilitaCodigo(ByVal habilita As Boolean)
    TrocaCorObjeto txtstrCodigo, Not habilita
    TrocaCorObjeto cbo_strTipodeEvento, Not habilita
    If habilita Then
        cbo_strTipodeEvento.ListIndex = -1
    End If
End Sub

Private Sub proximoCodigoEvento()
    If Not mblnAlterando Then
       gstrProximoCodigo txtstrCodigo, gstrEvento, "strCodigo", gintCodSeguranca
       MarcaCampo txtstrCodigo
    End If
End Sub

Private Function EventoJaUsado() As Boolean
    Dim strSQL      As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = ""
    
    strSQL = strSQL & " SELECT intEvento FROM " & gstrProgramaDeTrabalho
    strSQL = strSQL & " WHERE intEvento = " & txtPKId.Text
    
    strSQL = strSQL & " UNION ALL"
    strSQL = strSQL & " SELECT intEvento FROM " & gstrPrevisaoDaReceita
    strSQL = strSQL & " WHERE intEvento = " & txtPKId.Text
    
    strSQL = strSQL & " UNION ALL"
    strSQL = strSQL & " SELECT intEvento FROM " & gstrSuplementacaoReducaoReceita
    strSQL = strSQL & " WHERE intEvento = " & txtPKId.Text
    
    strSQL = strSQL & " UNION ALL"
    strSQL = strSQL & " SELECT intEvento FROM " & gstrSuplementacaoReducaoDespesa
    strSQL = strSQL & " WHERE intEvento = " & txtPKId.Text
    
    strSQL = strSQL & " UNION ALL"
    strSQL = strSQL & " SELECT intEvento FROM " & gstrDotacaoSuplementadaReduzida
    strSQL = strSQL & " WHERE intEvento = " & txtPKId.Text
    
    strSQL = strSQL & " UNION ALL"
    strSQL = strSQL & " SELECT intEvento FROM " & gstrEmpenho
    strSQL = strSQL & " WHERE intEvento = " & txtPKId.Text
    
    strSQL = strSQL & " UNION ALL"
    strSQL = strSQL & " SELECT intEvento FROM " & gstrSubempenho
    strSQL = strSQL & " WHERE intEvento = " & txtPKId.Text
    
    strSQL = strSQL & " UNION ALL"
    strSQL = strSQL & " SELECT intEvento FROM " & gstrProcessoPagamento
    strSQL = strSQL & " WHERE intEvento = " & txtPKId.Text
    
    strSQL = strSQL & " UNION ALL"
    strSQL = strSQL & " SELECT intEvento FROM " & gstrArrecadacaoReceita
    strSQL = strSQL & " WHERE intEvento = " & txtPKId.Text
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                EventoJaUsado = True
            End If
            .Close
        End With
    End If

    
End Function
Private Function VerificaItemNaLista(intTpEvento As Integer, objLista As ListView, objCombo As ComboBox) As Boolean
   Dim intContador As Integer
   
   If intTpEvento = 0 Or intTpEvento = 6 Then
      
      If objLista.Name = "lvw_ContaDebito" And Mid(objCombo.Text, 1, 1) = gstrDigitoDespesa Then
         ExibeMensagem "Não é possível inserir na guia Debitar um código com o primeiro dígito igual a " & gstrDigitoDespesa & "."
         VerificaItemNaLista = True
         Exit Function
      ElseIf objLista.Name = "lvw_ContaCredito" And Mid(objCombo.Text, 1, 1) = gstrDigitoReceita Then
         ExibeMensagem "Não é possível inserir na guia Creditar um código com o primeiro dígito igual a " & gstrDigitoReceita & "."
         VerificaItemNaLista = True
         Exit Function
      End If
      
      For intContador = 1 To objLista.ListItems.Count
         
         If objLista.Name = "lvw_ContaCredito" And Mid(objCombo.Text, 1, 1) = gstrDigitoDespesa And Mid(objLista.ListItems(intContador).Text, 1, 1) = gstrDigitoDespesa Then
            ExibeMensagem "Não é permitido inserir mais de uma conta contabil com o início do código igual a " & gstrDigitoDespesa & "."
            VerificaItemNaLista = True
            Exit Function
         ElseIf objLista.Name = "lvw_ContaDebito" And Mid(objCombo.Text, 1, 1) = gstrDigitoReceita And Mid(objLista.ListItems(intContador).Text, 1, 1) = gstrDigitoReceita Then
            ExibeMensagem "Não é permitido inserir mais de uma conta contabil com o início do código igual a " & gstrDigitoReceita & "."
            VerificaItemNaLista = True
            Exit Function
         End If
         
      Next
         
      If objLista.Name = "lvw_ContaDebito" Then
         For intContador = 1 To lvw_ContaCredito.ListItems.Count
            If (Mid(objCombo.Text, 1, 1) = gstrDigitoDespesa Or Mid(objCombo.Text, 1, 1) = gstrDigitoReceita) And (Mid(lvw_ContaCredito.ListItems(intContador).Text, 1, 1) = gstrDigitoDespesa Or Mid(lvw_ContaCredito.ListItems(intContador).Text, 1, 1) = gstrDigitoReceita) Then
               ExibeMensagem "Não é permitido inserir uma conta de debito com inicio igual a " & gstrDigitoDespesa & " ou " & gstrDigitoReceita & " pois já existe uma conta iniciada com " & gstrDigitoDespesa & " ou " & gstrDigitoReceita & " na guia credito. "
               VerificaItemNaLista = True
               Exit Function
            End If
         Next
      ElseIf objLista.Name = "lvw_ContaCredito" Then
         For intContador = 1 To lvw_ContaDebito.ListItems.Count
         
            If (Mid(objCombo.Text, 1, 1) = gstrDigitoDespesa Or Mid(objCombo.Text, 1, 1) = gstrDigitoReceita) And (Mid(lvw_ContaDebito.ListItems(intContador).Text, 1, 1) = gstrDigitoDespesa Or Mid(lvw_ContaDebito.ListItems(intContador).Text, 1, 1) = gstrDigitoReceita) Then
               ExibeMensagem "Não é permitido inserir uma conta de credito com inicio igual a " & gstrDigitoDespesa & " ou " & gstrDigitoReceita & " pois já existe uma conta iniciada com " & gstrDigitoDespesa & " ou " & gstrDigitoReceita & " na guia debito. "
               VerificaItemNaLista = True
               Exit Function
            End If
         Next
      End If
   ElseIf intTpEvento = 1 Then
   
      If objLista.Name = "lvw_ContaDebito" And (Mid(objCombo.Text, 1, 1) = gstrDigitoReceita Or Mid(objCombo.Text, 1, 1) = gstrDigitoDespesa) Then
         ExibeMensagem "Não é possível inserir na guia Debitar um código com o primeiro dígito igual a " & gstrDigitoDespesa & " ou " & gstrDigitoReceita & "."
         VerificaItemNaLista = True
         Exit Function
      ElseIf objLista.Name = "lvw_ContaCredito" And (Mid(objCombo.Text, 1, 1) = gstrDigitoDespesa) Then
         ExibeMensagem "Não é possível inserir na guia Creditar um código com o primeiro dígito igual a " & gstrDigitoDespesa & "."
         VerificaItemNaLista = True
         Exit Function
      End If
      
      For intContador = 1 To objLista.ListItems.Count
         
         If Mid(objCombo.Text, 1, 1) = gstrDigitoReceita And Mid(objLista.ListItems(intContador).Text, 1, 1) = gstrDigitoReceita Then
            ExibeMensagem "Não é permitido inserir mais de uma conta contabil com o início do código igual a " & gstrDigitoReceita & "."
            VerificaItemNaLista = True
            Exit Function
         End If
         
      Next
   ElseIf intTpEvento = 2 Then
   
      If objLista.Name = "lvw_ContaCredito" And (Mid(objCombo.Text, 1, 1) = gstrDigitoReceita Or Mid(objCombo.Text, 1, 1) = gstrDigitoReceita) Then
         ExibeMensagem "Não é possível inserir na guia Creditar um código com o primeiro dígito igual a " & gstrDigitoDespesa & " ou " & gstrDigitoReceita & "."
         VerificaItemNaLista = True
         Exit Function
      ElseIf objLista.Name = "lvw_ContaDebito" And (Mid(objCombo.Text, 1, 1) = gstrDigitoReceita) Then
         ExibeMensagem "Não é possível inserir na guia Debitar um código com o primeiro dígito igual a " & gstrDigitoReceita & "."
         VerificaItemNaLista = True
         Exit Function
      End If
      
      For intContador = 1 To objLista.ListItems.Count
         
         If Mid(objCombo.Text, 1, 1) = gstrDigitoDespesa And Mid(objLista.ListItems(intContador).Text, 1, 1) = gstrDigitoDespesa Then
            ExibeMensagem "Não é permitido inserir mais de uma conta contabil com o início do código igual a " & gstrDigitoDespesa & "."
            VerificaItemNaLista = True
            Exit Function
         End If
         
      Next
   ElseIf intTpEvento = 8 Then
      If Not VerificaContaBancaria Then
         ExibeMensagem "Não é permitido inserir esta conta contabil para o tipo de evento TRANSFERÊNCIAS."
         VerificaItemNaLista = True
         Exit Function
      End If
   End If
End Function
Private Function VerificaContaBancaria() As Boolean
   Dim strSQL As String
   Dim adoResultado As New ADODB.Recordset
   
   Select Case tab_3dPastaConta.Tab
      Case 0
         strSQL = "SELECT * FROM " & gstrPlanoConta & " PC "
         strSQL = strSQL & "WHERE PC.PKID = " & gstrItemData(cbointContaDebito) & " AND "
         strSQL = strSQL & "PC.blnAnalitica = 1 AND PC.bytDisponibilidadeDeCaixa = 1"
      Case 1
         strSQL = "SELECT * FROM " & gstrPlanoConta & " PC "
         strSQL = strSQL & "WHERE PC.PKID = " & gstrItemData(cbointContaCredito) & " AND "
         strSQL = strSQL & "PC.blnAnalitica = 1 AND PC.bytDisponibilidadeDeCaixa = 1"
   End Select

   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         VerificaContaBancaria = True
      End If
   End If


End Function

Private Function blnVerificaMovi() As Boolean
    Dim strSQL As String
    Dim adoTemp As ADODB.Recordset
    Dim adoResultado As ADODB.Recordset
    
    blnVerificaMovi = False
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "EVC.intContaContabil "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrEvento & " EV, "
    strSQL = strSQL & gstrEventoContaContabilCredito & "  EVC "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "EV.Pkid = EVC.Intevento AND "
    strSQL = strSQL & "EV.pkid = " & txtPKId
    strSQL = strSQL & " Order By EVC.intContaContabil"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoTemp) Then
        If adoTemp.RecordCount >= 1 Then
            adoTemp.MoveFirst
            Do While Not adoTemp.EOF
                strSQL = ""
                strSQL = strSQL & "Select "
                strSQL = strSQL & "Count(*) IntNumero "
                strSQL = strSQL & "From "
                strSQL = strSQL & gstrLancamentoContabil & " LC, "
                strSQL = strSQL & gstrPlanoConta & " PC, "
                strSQL = strSQL & gstrProcessoPagamento & " PG "
                strSQL = strSQL & "Where "
                strSQL = strSQL & "LC.INTPROCESSO = PG.PKID AND "
                strSQL = strSQL & "LC.Intconta = PC.pkid AND "
                strSQL = strSQL & "Not PG.INTLANCAMENTOCONTABIL is null AND "
                strSQL = strSQL & "PC.Pkid = " & adoTemp!intContaContabil
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    If adoResultado!INTNUMERO >= 1 Then
                        blnVerificaMovi = True
                        Exit Function
                    End If
                End If
                adoTemp.MoveNext
            Loop
        End If
    End If


End Function


