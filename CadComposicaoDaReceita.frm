VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadComposicaoDaReceita 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Composição das Receitas"
   ClientHeight    =   7050
   ClientLeft      =   3390
   ClientTop       =   2205
   ClientWidth     =   6720
   HelpContextID   =   22
   Icon            =   "CadComposicaoDaReceita.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5190
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   9155
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Composição das Receitas"
      TabPicture(0)   =   "CadComposicaoDaReceita.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintTipo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrSigla"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintCodigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintUtilizacao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_Valorminimo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrSigla"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtintCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstrDescricao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkbytDividaAtiva"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fra_Itens"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cbointUtilizacao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cbointTipo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPKId"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtdblParcelaminima"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Fundamento Legal"
      TabPicture(1)   =   "CadComposicaoDaReceita.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_strFundamentoTexto"
      Tab(1).Control(1)=   "txt_strComposicao"
      Tab(1).Control(2)=   "dbc_intExercicio"
      Tab(1).Control(3)=   "lvw_Itens"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "lbl_intExercicio"
      Tab(1).Control(6)=   "lbl_strComposicao"
      Tab(1).ControlCount=   7
      Begin VB.TextBox txt_strFundamentoTexto 
         Height          =   1185
         Left            =   -74790
         MaxLength       =   340
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1980
         Width           =   6255
      End
      Begin VB.TextBox txt_strComposicao 
         Height          =   285
         Left            =   -73050
         MaxLength       =   70
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   720
         Width           =   4530
      End
      Begin VB.TextBox txtdblParcelaminima 
         Height          =   285
         Left            =   5025
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1320
         Width           =   1470
      End
      Begin VB.TextBox txtPKId 
         Height          =   285
         Left            =   2220
         TabIndex        =   16
         Top             =   330
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ComboBox cbointTipo 
         Height          =   315
         ItemData        =   "CadComposicaoDaReceita.frx":107A
         Left            =   960
         List            =   "CadComposicaoDaReceita.frx":1084
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1950
         Width           =   2280
      End
      Begin VB.ComboBox cbointUtilizacao 
         Height          =   315
         ItemData        =   "CadComposicaoDaReceita.frx":10AC
         Left            =   4095
         List            =   "CadComposicaoDaReceita.frx":10AE
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1950
         Width           =   2385
      End
      Begin VB.Frame fra_Itens 
         Caption         =   "Receitas"
         Height          =   2820
         Left            =   60
         TabIndex        =   13
         Top             =   2295
         Width           =   6345
         Begin TrueOleDBGrid70.TDBDropDown tdd_Receita 
            Height          =   1575
            Left            =   150
            TabIndex        =   15
            Top             =   690
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   2778
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
            Columns(2).DataField=   "strsigla"
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
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7541"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7461"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
            _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(55)  =   "Named:id=39:EvenRow"
            _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(57)  =   "Named:id=40:OddRow"
            _StyleDefs(58)  =   ":id=40,.parent=33"
            _StyleDefs(59)  =   "Named:id=41:RecordSelector"
            _StyleDefs(60)  =   ":id=41,.parent=34"
            _StyleDefs(61)  =   "Named:id=42:FilterBar"
            _StyleDefs(62)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid grd_Receita 
            Height          =   2400
            Left            =   150
            TabIndex        =   14
            Top             =   270
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   4233
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Pkid"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descrição"
            Columns(1).DataField=   "strY"
            Columns(1).DropDown=   "tdd_Receita"
            Columns(1).DropDown.vt=   8
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Sigla"
            Columns(2).DataField=   "strsigla"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7541"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7461"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(1).AutoDropDown=1"
            Splits(0)._ColumnProps(13)=   "Column(1).AutoCompletion=1"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
            _StyleDefs(16)  =   ":id=8,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
            _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=39:EvenRow"
            _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(60)  =   "Named:id=40:OddRow"
            _StyleDefs(61)  =   ":id=40,.parent=33"
            _StyleDefs(62)  =   "Named:id=41:RecordSelector"
            _StyleDefs(63)  =   ":id=41,.parent=34"
            _StyleDefs(64)  =   "Named:id=42:FilterBar"
            _StyleDefs(65)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.CheckBox chkbytDividaAtiva 
         Caption         =   "Dívida Ativa"
         Height          =   195
         Left            =   990
         TabIndex        =   8
         Top             =   1695
         Width           =   1260
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   960
         MaxLength       =   70
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1000
         Width           =   5520
      End
      Begin VB.TextBox txtintCodigo 
         Height          =   285
         Left            =   960
         MaxLength       =   8
         TabIndex        =   2
         Top             =   690
         Width           =   1290
      End
      Begin VB.TextBox txtstrSigla 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1310
         Width           =   1290
      End
      Begin MSDataListLib.DataCombo dbc_intExercicio 
         Height          =   315
         Left            =   -73980
         TabIndex        =   24
         Top             =   1290
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSComctlLib.ListView lvw_Itens 
         Height          =   1665
         Left            =   -74790
         TabIndex        =   25
         Top             =   3300
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   2937
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Exercício"
            Object.Width           =   10954
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Fundamento Legal"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fundamento Legal Texto"
         Height          =   195
         Left            =   -74790
         TabIndex        =   23
         Top             =   1710
         Width           =   1770
      End
      Begin VB.Label lbl_intExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   -74790
         TabIndex        =   21
         Top             =   1305
         Width           =   675
      End
      Begin VB.Label lbl_strComposicao 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   -74850
         TabIndex        =   19
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label lbl_Valorminimo 
         AutoSize        =   -1  'True
         Caption         =   "Valor mínimo da Parcela:"
         Height          =   195
         Left            =   3180
         TabIndex        =   17
         Top             =   1365
         Width           =   1770
      End
      Begin VB.Label lblintUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   3330
         TabIndex        =   12
         Top             =   1995
         Width           =   690
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   375
         TabIndex        =   1
         Top             =   750
         Width           =   495
      End
      Begin VB.Label lblstrSigla 
         AutoSize        =   -1  'True
         Caption         =   "Sigla"
         Height          =   195
         Left            =   525
         TabIndex        =   5
         Top             =   1360
         Width           =   345
      End
      Begin VB.Label lblintTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   555
         TabIndex        =   11
         Top             =   1995
         Width           =   315
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   1055
         Width           =   720
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Composicao 
      Height          =   1770
      Left            =   30
      TabIndex        =   22
      Top             =   5250
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   3122
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   "PKId"
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
      Columns(3).Caption=   "Sigla"
      Columns(3).DataField=   "strSigla"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1799"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1720"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=6138"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6059"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadComposicaoDaReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando      As Boolean
    Dim mblnAlterandoLista As Boolean
    Dim mcboAux            As ComboBox
    Dim adoRec             As adodb.Recordset
    Dim adoTdb             As adodb.Recordset
    Dim X                  As XArrayDB
    Dim Y                  As New XArrayDB
    Dim Z                  As New XArrayDB
    Dim intMaxPKId         As Integer
    Dim mobjGeral          As Object
    Dim mblnPrimeiraVez    As Boolean
    Dim mobjAux            As Object
    Dim strDuplicataCodigo As String
    
    Dim mobjLista          As Object
                
    Dim bytOrdenacao1      As Byte
    Dim blnOrdenacaoAsc1   As Boolean
            
    Dim bytOrdenacao2      As Byte
    Dim blnOrdenacaoAsc2   As Boolean
    
    Dim strCodigoAtual     As String
    Dim strDescricaoAtual  As String
    Dim strSiglaAtual      As String
    Dim strCodigo          As String
    
   
Function PreencheTBGrid()
    Dim strSql  As String
    LimpaGrid
    strSql = strQueryGrid
    Set gobjBanco = New clsBanco
    gobjBanco.CriaADO strSql, 5, adoRec
    MontaArray
    LeDaTabelaParaObj "", tdd_Receita, "Select PKId, strDescricao, strsigla From " & gstrReceita & " Order by strDescricao"
    tdd_Receita.ReBind
    tdd_Receita.Refresh
End Function

Private Sub cbointTipo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointUtilizacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbytDividaAtiva_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 445
    VirificaGradeListView Me
    
    If tab_3dPasta.Tab = 1 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
    
    If mblnAlterando Then
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
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub

Private Sub Form_Load()
    Dim strSql  As String
    
    
    TrocaCorObjeto txt_strComposicao, True
        
    bytOrdenacao1 = 2: blnOrdenacaoAsc1 = True
    bytOrdenacao2 = 1: blnOrdenacaoAsc2 = True
   
    If MDIMenu.Tag = "PROTOCOLO" Then
        'strSql = "SELECT PKId, intCodigo, strDescricao, strSigla FROM " & gstrComposicaoDaReceita & " WHERE intUtilizacao > 2 "
    Else
        cbointUtilizacao.AddItem "Imobiliárias "
        cbointUtilizacao.ItemData(cbointUtilizacao.NewIndex) = "1"
        cbointUtilizacao.AddItem "Econômicas"
        cbointUtilizacao.ItemData(cbointUtilizacao.NewIndex) = "2"
    End If
    cbointUtilizacao.AddItem "Dívida Ativa"
    cbointUtilizacao.ItemData(cbointUtilizacao.NewIndex) = "3"
    cbointUtilizacao.AddItem "Acordo"
    cbointUtilizacao.ItemData(cbointUtilizacao.NewIndex) = "4"
    cbointUtilizacao.AddItem "Preço Público"
    cbointUtilizacao.ItemData(cbointUtilizacao.NewIndex) = "5"
    cbointUtilizacao.AddItem "ISS Construção"
    cbointUtilizacao.ItemData(cbointUtilizacao.NewIndex) = "6"
    cbointUtilizacao.AddItem "Imobiliário - Taxas"
    cbointUtilizacao.ItemData(cbointUtilizacao.NewIndex) = "7"
    cbointUtilizacao.AddItem "Iss Movimento - GissOnLine"
    cbointUtilizacao.ItemData(cbointUtilizacao.NewIndex) = "9"
    cbointUtilizacao.AddItem "Outros"
    cbointUtilizacao.ItemData(cbointUtilizacao.NewIndex) = "8"
    
    VerificaObjParaAplicar mobjAux
    PreencheTBGrid
    Set X = New XArrayDB
    X.ReDim 0, 0, 0, 1
    X.Clear
    Set grd_Receita.Array = X
    grd_Receita.ReBind
    grd_Receita.Refresh
End Sub

Private Function strQuery() As String
Dim strSql As String
   
   strSql = ""
   strSql = strSql & "SELECT PKId, intCodigo, strDescricao, strSigla "
   strSql = strSql & "FROM " & gstrComposicaoDaReceita & " "
   
   If MDIMenu.Tag = "PROTOCOLO" Then
      strSql = strSql & "WHERE intUtilizacao > 2 "
   End If
   
 ' TIMTIM - 10/02/2003 - Pendência nº 4
 ' strSql = strSql & "ORDER BY strDescricao"
   
   Select Case bytOrdenacao1
      
      Case Is = 1
         strSql = strSql & "ORDER BY intCodigo" & IIf(blnOrdenacaoAsc1, " ASC", " DESC")
         
      Case Is = 2
         strSql = strSql & "ORDER BY strDescricao" & IIf(blnOrdenacaoAsc1, " ASC", " DESC")
      
      Case Is = 3
         strSql = strSql & "ORDER BY strSigla" & IIf(blnOrdenacaoAsc1, " ASC", " DESC")
         
   End Select
   
   strQuery = strSql
   
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnPrimeiraVez = False
End Sub

Function strQueryGrid() As String

'******************************************************************************************
' Data: 27/03/2003
' Alteração: - Retirado o comando CONVERT da cláusula ORDER BY uma vez que este não era
'            necessário.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String

   strSql = ""
   strSql = strSql & "SELECT VL.intReceita, RC.strDescricao, RC.strSigla FROM "
   strSql = strSql & gstrValorCompoRec & " VL,"
   strSql = strSql & gstrReceita & " RC "
   strSql = strSql & "WHERE RC.PKId = VL.intReceita "
   
   If txtPKId <> "" Then
      strSql = strSql & "AND VL.intComposicaoDaReceita = " & tdb_Composicao.Columns("PKID").Value
   End If
       
   Select Case bytOrdenacao2
      
      Case Is = 0
'         strSql = strSql & " ORDER BY CONVERT(int, RTRIM(VL.intReceita))" & IIf(blnOrdenacaoAsc2, " ASC", " DESC")
         strSql = strSql & " ORDER BY VL.intReceita " & IIf(blnOrdenacaoAsc2, " ASC", " DESC")
         
      Case Is = 1
         strSql = strSql & "ORDER BY RC.strDescricao" & IIf(blnOrdenacaoAsc2, " ASC", " DESC")
      
   End Select
   
   strQueryGrid = strSql
   
End Function

Function PegaMaxPKId()
    Dim strSql          As String
    Dim adoResultado    As adodb.Recordset
    strSql = ""
    strSql = strSql & "SELECT MAX(PKId) as PKId "
    strSql = strSql & "FROM " & gstrComposicaoDaReceita
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
         intMaxPKId = adoResultado!Pkid
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
End Sub



Private Sub lvw_Itens_Click()
If lvw_Itens.ListItems.Count > 0 Then
       dbc_intExercicio.Text = lvw_Itens.SelectedItem.Text
       txt_strFundamentoTexto = lvw_Itens.SelectedItem.SubItems(1)
       mblnAlterandoLista = True
    End If
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3dPasta.Tab = 1 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
End Sub

Private Sub tdb_Composicao_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Composicao) = 1 Then
        tdb_Composicao_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Composicao_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Composicao_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Composicao
End Sub
Private Sub tdb_Composicao_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Composicao, ColIndex
End Sub

Private Sub tdb_Composicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Composicao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Composicao
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                txtPKId.Text = .Columns("PKID").Value
                mblnAlterando = True
                
                HabilitaItens
                
                LeDaTabelaParaObj gstrComposicaoDaReceita, Me
                PreencheTBGrid
                gCorLinhaSelecionada tdb_Composicao
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                strCodigoAtual = txtintCodigo.Text
                strDescricaoAtual = txtstrdescricao.Text
                strSiglaAtual = txtstrSigla.Text
                LimpaItens
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim intAux          As Integer
    Dim pkIDAux         As String
    Dim blnAlterandoAux As Boolean
    
    If UCase(strModoOperacao) = gstrAplicar Then
        ToolBarGeral strModoOperacao, gstrComposicaoDaReceita, mblnAlterando, tdb_Composicao, Me, mobjAux, strQuery, strQueryAplicar
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = gstrLocalizar Or UCase(strModoOperacao) = gstrPreencherLista Then
        ToolBarGeral strModoOperacao, gstrComposicaoDaReceita, mblnAlterando, tdb_Composicao, Me, mobjAux, strQuery
    End If
    
    If mblnAlterando Then
        intAux = Val(txtPKId.Text)
    Else
        intAux = 0
    End If
    
    If UCase(strModoOperacao) = "NOVO" Then
        LimpaItens
        If tab_3dPasta.Tab = 1 Then
            Exit Sub
        End If
    End If
    
    If UCase(strModoOperacao) = "SALVAR" Then
        If Not blnDadosOk Then Exit Sub
    End If
    
    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
        mblnPrimeiraVez = False
    End If
    
    If UCase(strModoOperacao) = UCase(gstrIncluirItem) Then
        If tab_3dPasta.Tab = 1 Then
            IncluirItemNoGrid
        End If
    End If

    If UCase(strModoOperacao) = UCase(gstrExcluirItem) Then
        If tab_3dPasta.Tab = 1 Then
            ExcluirItemNoGrid
        End If
    End If
    
    If UCase(strModoOperacao) = gstrDeletar Then
        DeletaValores intAux, strModoOperacao
    End If
    
    Set gobjBanco = New clsBanco
    
    blnAlterandoAux = mblnAlterando
    pkIDAux = txtPKId
    
    If ToolBarGeral(strModoOperacao, gstrComposicaoDaReceita, mblnAlterando, tdb_Composicao, Me, mobjAux, strQuery, , rptComposicaoDaReceita, strQueryRelatorio) Then
        
        If UCase(strModoOperacao) = UCase(gstrSalvar) Then
            If blnAlterandoAux Then
                If Not gobjBanco.Execute(StrSalvaItem(pkIDAux)) Then
                    ExibeMensagem "Ocorreram erros ao gravar os Fundamentos Legais."
                    Exit Sub
                End If
            End If
        End If
    
        If intAux > 0 Then
            If GravaValores((intAux), UCase(strModoOperacao)) Then
            End If
        Else
            PegaMaxPKId
            If GravaValores((intMaxPKId), UCase(strModoOperacao)) Or DeletaValores((intMaxPKId), UCase(strModoOperacao)) Then
                intMaxPKId = 0
            End If
        End If
        If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
            LeDaTabelaParaObj gstrComposicaoDaReceita, tdb_Composicao, strQuery
            strCodigoAtual = ""
            strDescricaoAtual = ""
            strSiglaAtual = ""
            strCodigo = ""
        End If
        HabilitaItens
    End If
    
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    If UCase(strModoOperacao) = "NOVO" Then
        PreencheTBGrid
        Set X = New XArrayDB
        X.ReDim 0, 0, 0, 1
        X.Clear
        Set grd_Receita.Array = X
        grd_Receita.ReBind
        grd_Receita.Refresh
        txtintCodigo.SetFocus
        strCodigoAtual = ""
        strDescricaoAtual = ""
        strSiglaAtual = ""
        strCodigo = ""
        HabilitaItens
    End If
    

    
End Sub

Private Sub txtdblParcelaminima_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtdblParcelaminima
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub txtstrDescricao_Change()
    txt_strComposicao = txtstrdescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
        CaracterValido KeyAscii, "A", txtstrdescricao
End Sub

Private Sub txtstrSigla_GotFocus()
    MarcaCampo txtstrSigla
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrdescricao
End Sub

Private Sub txtintCodigo_GotFocus()
    gstrProximoCodigo txtintCodigo, gstrComposicaoDaReceita, "intCodigo", gintCodSeguranca
    MarcaCampo txtintCodigo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown
    Select Case KeyCode
    Case vbKeyEscape
        If Not IsNull(tdd_Receita.SelectedItem) Then
            grd_Receita.SelStart = Len(grd_Receita.Text)
        End If
        SendKeys "{RIGHT}"
        Exit Sub
    End Select
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
End Sub

Private Sub grd_Receita_KeyPress(KeyAscii As Integer)
    Select Case grd_Receita.Col
    Case 0
    Case 1
    Case 2
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            SendKeys "%{DOWN}"
        End If
    End Select
        CaracterValido KeyAscii
End Sub
                
' TIMTIM - 11/02/2003 - Pendência nº 4
Private Sub grd_Receita_HeadClick(ByVal ColIndex As Integer)
      
   blnOrdenacaoAsc2 = IIf(bytOrdenacao2 = ColIndex, Not blnOrdenacaoAsc2, True)
   
   bytOrdenacao2 = ColIndex: PreencheTBGrid
   
End Sub

Private Sub tdd_Receita_DropDownClose()
Dim intRow As Integer
    
    On Error GoTo Err_Handle
    
    If Not IsNull(tdd_Receita.SelectedItem) Then
        
        With grd_Receita
            intRow = .Row '+ 1
'            If .Row = (X.Count(1)) Then
                'If X.UpperBound(1) > 1 Then
                    'If .Row <> (X.Count(1)) Then
                    '    X.ReDim 0, X.UpperBound(1) + 1, 0, 2
                    'End If
                'Else
                    X.ReDim 0, X.UpperBound(1), 0, 2
                'End If
                Set grd_Receita.Array = X
                grd_Receita.ReBind
                grd_Receita.Refresh
                DoEvents
                '.MoveFirst
'            End If
            .Row = intRow
            .Col = 0
            .SetFocus
        End With
        
        grd_Receita.Columns(0) = tdd_Receita.Columns(0)
        grd_Receita.Columns(1) = tdd_Receita.Columns(1)
        grd_Receita.Columns(2) = tdd_Receita.Columns(2)

'    Else
'        grd_Receita.Columns(0) = ""
'        grd_Receita.Columns(1) = ""
'        grd_Receita.Columns(2) = ""
    End If

    Exit Sub
    
Err_Handle:

End Sub

Private Sub MontaArray()
    Dim varAux As Variant
    Set X = New XArrayDB
    X.Clear
    With adoRec
        If Not .EOF And mblnAlterando Then
            X.ReDim 0, .RecordCount - 1, 0, 2
            Do While Not .EOF
                varAux = .Fields(0)
                X(.AbsolutePosition - 1, 0) = varAux
                varAux = .Fields(1)
                X(.AbsolutePosition - 1, 1) = varAux
                varAux = .Fields(2)
                X(.AbsolutePosition - 1, 2) = varAux
                .MoveNext
            Loop
        Else
            X.ReDim 0, 0, 0, 1
            X(0, 0) = ""
            X(0, 1) = ""
        End If
    End With
    
   ' TIMTIM - 11/02/2003 - Pendência nº 4
       
   Select Case bytOrdenacao2
      
      Case Is = 0
         X.QuickSort X.LowerBound(1), X.UpperBound(1), 0, IIf(blnOrdenacaoAsc2, XORDER_ASCEND, XORDER_DESCEND), XTYPE_LONG
         
      Case Is = 1
         X.QuickSort X.LowerBound(1), X.UpperBound(1), 1, IIf(blnOrdenacaoAsc2, XORDER_ASCEND, XORDER_DESCEND), XTYPE_STRING
      
   End Select
      
    Set grd_Receita.Array = X
    grd_Receita.ReBind
    grd_Receita.Refresh
    
End Sub

Function DeletaValores(intCodComposicao As Integer, strOperacao As String)
    If strOperacao = "DELETAR" Then
        Dim strSql As String
        If MsgBox("Confirma exclusão das receitas desta composição?", vbQuestion + vbYesNo) = vbYes Then
            strSql = ""
            strSql = strSql & "DELETE FROM " & gstrValorCompoRec & " "
            strSql = strSql & "WHERE  intComposicaoDaReceita = " & intCodComposicao
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSql
            LimpaGrid
        End If
    End If
End Function

Function GravaValores(intCodComposicao As Integer, strOperacao As String)
    
'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim strMsg As String
    Dim i      As Integer
    If strOperacao = "SALVAR" Then
    
        'grd_Receita.MoveFirst
        'If X(0, 0) = "" Or IsEmpty(X(0, 0)) Or IsNull(X(0, 0)) Then
        If grd_Receita.ApproxCount <= 1 And grd_Receita.Columns("Pkid").Value = "" Then
            ExibeMensagem "Não foram incluídas receitas para esta composição!"
            intMaxPKId = 0
            Exit Function
        End If
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrValorCompoRec & " "
        strSql = strSql & "WHERE  intComposicaoDaReceita = " & intCodComposicao
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSql
        
        grd_Receita.MoveFirst
            
'         For i = 0 To X.Count(1) - 1
         For i = 0 To grd_Receita.ApproxCount - 1
             
             If Len(Trim(grd_Receita.Columns(1).Value)) <> 0 Then
                 strSql = ""
                 strSql = strSql & "INSERT INTO " & gstrValorCompoRec & " "
                 strSql = strSql & "(intComposicaoDaReceita, intReceita, "
                 strSql = strSql & " dtmDtAtualizacao, lngCodUsr "
                 strSql = strSql & ") Values ("
                 strSql = strSql & intCodComposicao & ", "
                 strSql = strSql & Val(grd_Receita.Columns("Pkid").Value) & ", "
'                 strSql = strSql & Val(X(i, 0)) & ", "
                 strSql = strSql & strGETDATE & ", "
                 strSql = strSql & glngCodUsr
                 strSql = strSql & ")"
        
                 If Not gobjBanco.Execute(strSql, False) Then
                     gobjBanco.ExecutaRollbackTrans
                 End If
             End If
             
             grd_Receita.MoveNext
             
        Next i
        
        gobjBanco.ExecutaCommitTrans
        LimpaGrid
        
    End If
    
End Function

Private Sub LimpaGrid()
    Set X = New XArrayDB
    Set Y = New XArrayDB
    Set Z = New XArrayDB
    X.ReDim 0, 0, 0, 1
    X.Clear
    Y.Clear
    Z.Clear
    Set grd_Receita.Array = X
    grd_Receita.ReBind
    grd_Receita.Refresh
    Set tdd_Receita.Array = Z
    tdd_Receita.ReBind
    tdd_Receita.Refresh
End Sub

Private Function blnDadosOk() As Boolean
    Dim strSql As String
    Dim adoResultado As adodb.Recordset
    
    blnDadosOk = False
    
    If Val(txtintCodigo.Text) = 0 Then
        ExibeMensagem "O código deve ser informado."
        txtintCodigo.SetFocus
        Exit Function
    ElseIf Trim(txtstrdescricao.Text) = "" Then
        ExibeMensagem "A Descrição deve ser informada."
        txtstrdescricao.SetFocus
        Exit Function
    ElseIf Trim(txtstrSigla.Text) = "" Then
        ExibeMensagem "A Sigla deve ser informada."
        txtstrSigla.SetFocus
        Exit Function
    ElseIf cbointTipo.ListIndex = -1 Then
        ExibeMensagem "O Tipo deve ser informado."
        cbointTipo.SetFocus
        Exit Function
    End If
        
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.Text)) Then

ProximoCodigo:

        If gblnExisteCodigo(1, gstrComposicaoDaReceita, "intCodigo", "'" & txtintCodigo.Text & "'") Then
            strCodigo = (gstrProximoCodigo(txtintCodigo, gstrComposicaoDaReceita, "intCodigo", gintCodSeguranca, , , , True))
            If MsgBox("O código informado já se encontra cadastrado. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                txtintCodigo.SetFocus
                Exit Function
            Else
                txtintCodigo.Text = strCodigo
                GoTo ProximoCodigo
            End If
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrdescricao.Text) <> UCase$(strDescricaoAtual)) Then
            
        If gblnExisteCodigo(1, gstrComposicaoDaReceita, "strDescricao", "'" & txtstrdescricao.Text & "'") Then
            ExibeMensagem "A descrição informada já se encontra cadastrada."
            txtstrdescricao.SetFocus
            Exit Function
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrSigla.Text) <> UCase$(strSiglaAtual)) Then
            
        If gblnExisteCodigo(1, gstrComposicaoDaReceita, "strSigla", "'" & txtstrSigla.Text & "'") Then
            ExibeMensagem "A Sigla informada já se encontra cadastrada."
            txtstrSigla.SetFocus
            Exit Function
        End If
    End If

    strSql = ""
    strSql = strSql & "SELECT R.Bytinscreveda"
    strSql = strSql & "  FROM TBLRECEITA R, TBLRECEITA RR"
    strSql = strSql & " WHERE R.INTDIVIDAATIVA " & strOUTJSQLServer & "= RR.PKID" & strOUTJOracle
    strSql = strSql & "   AND R.Strdescricao = '" & grd_Receita.Columns(1) & "' "
    strSql = strSql & " ORDER BY R.STRDESCRICAO ASC"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
       If Not adoResultado.EOF And Not adoResultado.BOF Then
           
           If chkbytDividaAtiva.Value = vbChecked Then
               If adoResultado!bytinscreveDa = 1 Then
                   blnDadosOk = True
               Else
                   ExibeMensagem "Verificar inscrição em Divida Ativa - Composição e Receitas"
                   Exit Function
               End If
           Else
               If adoResultado!bytinscreveDa = 0 Then
                   blnDadosOk = True
               Else
                   ExibeMensagem "Verificar inscrição em Divida Ativa - Composição e Receitas"
                   Exit Function
               End If
           End If
           
       End If
    End If
    
    blnDadosOk = True
    
End Function

Private Sub txtintCodigo_LostFocus()
 '   If blnTemDuplicata = True Then
 '       LeDaTabelaParaObj gstrComposicaoDaReceita, Me
 '       mblnAlterando = True
 '   End If
End Sub

'Function blnTemDuplicata() As Boolean
'    Dim strSql             As String
'    Dim AdoResultado       As ADODB.Recordset
'    blnTemDuplicata = False
'    strDuplicataCodigo = ""
'    If Val(txtintCodigo) = 0 Then Exit Function
'    strSql = ""
'    strSql = strSql & "SELECT * FROM "
'    strSql = strSql & gstrComposicaoDaReceita
'    strSql = strSql & " WHERE intCodigo = " & txtintCodigo
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, AdoResultado) Then
'        With AdoResultado
'            If Not .EOF Then
'                strDuplicataCodigo = strSql
'                blnTemDuplicata = True
'                txtPkid.Text = !PkId
'                Exit Function
'            End If
'        End With
'    End If
'End Function

Private Sub txtstrSigla_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Function strQueryRelatorio() As String
Dim strSql As String
   
   strSql = ""
   strSql = strSql & "SELECT CR.intCodigo, CR.strDescricao, "
   strSql = strSql & "CR.strSigla, CR.intTipo, "
   strSql = strSql & "CR.bytDividaAtiva, CR.intUtilizacao, "
   strSql = strSql & "VL.intReceita, RC.strDescricao strReceita "
   strSql = strSql & "FROM "
   strSql = strSql & gstrComposicaoDaReceita & " CR, "
   strSql = strSql & gstrValorCompoRec & " VL, "
   strSql = strSql & gstrReceita & " RC "
   strSql = strSql & "WHERE VL.intComposicaoDaReceita = CR.PKId "
   strSql = strSql & "AND RC.PKId = VL.intReceita"
   
   Select Case bytOrdenacao1
      
      Case Is = 1
         strSql = strSql & " ORDER BY CR.intCodigo" & IIf(blnOrdenacaoAsc1, " ASC", " DESC")
         
      Case Is = 2
         strSql = strSql & " ORDER BY CR.strDescricao" & IIf(blnOrdenacaoAsc1, " ASC", " DESC")
      
      Case Is = 3
         strSql = strSql & " ORDER BY CR.strSigla" & IIf(blnOrdenacaoAsc1, " ASC", " DESC")
         
   End Select
   
   strQueryRelatorio = strSql
   
End Function

Private Function strQueryAplicar() As String
    Dim strSql As String
   
    strSql = ""
    strSql = "SELECT Pkid,"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita
    strSql = strSql & " WHERE pkid = " & tdb_Composicao.Columns("PKID").Value
    strQueryAplicar = strSql
   
End Function

Private Function IncluirItemNoGrid()
    Dim intInd          As Integer
    If blnDadosItens = False Then Exit Function
    With lvw_Itens
        If mblnAlterandoLista Then
            For intInd = 1 To .ListItems.Count
                If .SelectedItem.Index <> intInd Then
                    If Trim(dbc_intExercicio.Text) = .ListItems(intInd).Text Then
                        ExibeMensagem "Não é possível incluir itens com exercícios iguais."
                        Exit Function
                    End If
                End If
            Next
            .SelectedItem.Text = dbc_intExercicio.Text
            .SelectedItem.SubItems(1) = txt_strFundamentoTexto
        Else
            For intInd = 1 To .ListItems.Count
                If Trim(dbc_intExercicio.Text) = .ListItems(intInd).Text Then
                    ExibeMensagem "Não é possível incluir itens com exercícios iguais."
                    Exit Function
                End If
            Next

            Set mobjLista = .ListItems.Add(, , dbc_intExercicio.Text)
            mobjLista.SubItems(1) = txt_strFundamentoTexto
            
        End If
    End With
    dbc_intExercicio.Text = ""
    txt_strFundamentoTexto = ""
    LimpaItens
    
End Function

Private Sub HabilitaItens()
Dim strSql As String
Dim adoResultado As adodb.Recordset
    
    lvw_Itens.ListItems.Clear
    
    tab_3dPasta.TabEnabled(1) = mblnAlterando
    
    LimpaItens
    
    If mblnAlterando Then
        strSql = ""
        strSql = strSql & "Select FL.PkID, FL.intExercicio, FL.strDescricao "
        strSql = strSql & "From "
        strSql = strSql & gstrFundamentoLegal & " FL "
        strSql = strSql & "Where "
        strSql = strSql & "FL.intComposicaoDaReceita = " & txtPKId
        strSql = strSql & " Order By FL.intExercicio"
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                If .EOF = False Then
                    Do While Not .EOF
                        Set mobjLista = lvw_Itens.ListItems.Add(, , gstrENulo(!intExercicio))
                        mobjLista.SubItems(1) = gstrENulo(!strDescricao)
                        
                        .MoveNext
                    Loop
                End If
            End With
        End If
        
        Set dbc_intExercicio.RowSource = Nothing
        dbc_intExercicio.Refresh
    
        dbc_intExercicio.Tag = strQueryExercicio & ";intExercicio"
        PreencherListaDeOpcoes dbc_intExercicio
    Else
        tab_3dPasta.Tab = 0
    End If
        
    
End Sub

Private Function strQueryExercicio() As String
Dim strSql As String

    strSql = "SELECT Pkid, intExercicio "
    strSql = strSql & " FROM "
    strSql = strSql & gstrParametroAtualizacao
    strSql = strSql & " WHERE"
    strSql = strSql & " intComposicaoReceita = " & txtPKId
    strSql = strSql & " ORDER BY intExercicio"

    strQueryExercicio = strSql

End Function

Private Function StrSalvaItem(mtxtPkId As String) As String
    Dim strSql  As String
    Dim intInd  As Integer
    
    strSql = ""
    If lvw_Itens.ListItems.Count > 0 Then
        strSql = IIf(bytDBType = Oracle, "Begin", "")
    End If
    

    strSql = strSql & " Delete from " & gstrFundamentoLegal & " Where intComposicaoDaReceita = " & mtxtPkId
    If lvw_Itens.ListItems.Count > 0 Then
       strSql = strSql & IIf(bytDBType = Oracle, ";", "")
    End If

    If lvw_Itens.ListItems.Count > 0 Then
        With lvw_Itens
            For intInd = 1 To .ListItems.Count
                strSql = strSql & " INSERT INTO "
                strSql = strSql & gstrFundamentoLegal & " ("
                strSql = strSql & "intComposicaoDaReceita, "
                strSql = strSql & "intexercicio, "
                strSql = strSql & "strDescricao, "
                strSql = strSql & "dtmDtAtualizacao, "
                strSql = strSql & "lngCodUsr) "
                strSql = strSql & "Values("
                strSql = strSql & mtxtPkId & ", "
                strSql = strSql & .ListItems(intInd).Text & ", "
                strSql = strSql & "'" & gstrConvVrParaSql(.ListItems(intInd).SubItems(1)) & "', "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSql = strSql & glngCodUsr & " "
                strSql = strSql & ")" & IIf(bytDBType = Oracle, ";", "")
            Next
        End With
    End If
    If lvw_Itens.ListItems.Count > 0 Then
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    End If
    StrSalvaItem = strSql
End Function

Private Function blnDadosItens() As Boolean
    
    blnDadosItens = False
    
        If Not dbc_intExercicio.MatchedWithList Then
            ExibeMensagem "Informe um exercício válido"
            dbc_intExercicio.SetFocus
            Exit Function
        End If
    
    blnDadosItens = True
    
End Function

Private Function ExcluirItemNoGrid()
    With lvw_Itens
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
End Function

Private Sub LimpaItens()
    txt_strFundamentoTexto = ""
    dbc_intExercicio.Text = ""
    mblnAlterandoLista = False
End Sub


