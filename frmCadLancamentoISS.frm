VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadLancamentoISS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento ISS Construção"
   ClientHeight    =   6810
   ClientLeft      =   210
   ClientTop       =   2310
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11640
   Begin VB.TextBox txtstrNumeroAviso 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   8040
      MaxLength       =   10
      TabIndex        =   3
      Top             =   90
      Width           =   1125
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4635
      Left            =   60
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   555
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   8176
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lançamento ISS Construção"
      TabPicture(0)   =   "frmCadLancamentoISS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Cabecalho"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Prédios / Demonstrativo"
      TabPicture(1)   =   "frmCadLancamentoISS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Parcelas"
      Tab(1).Control(1)=   "fra_Predios"
      Tab(1).ControlCount=   2
      Begin VB.Frame fra_Predios 
         Caption         =   "Prédios"
         Height          =   2010
         Left            =   -74880
         TabIndex        =   54
         Top             =   360
         Width           =   11265
         Begin TrueOleDBGrid70.TDBGrid tdb_Predios 
            Height          =   1635
            Left            =   120
            TabIndex        =   55
            Top             =   255
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   2884
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
            Columns(1).Caption=   "Construção"
            Columns(1).DataField=   "dtmDataConstrucao"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tipo Construção"
            Columns(2).DataField=   "strTipoConstrucao"
            Columns(2).NumberFormat=   "Standard"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Acabamento"
            Columns(3).DataField=   "strTipoAcabamento"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Área"
            Columns(4).DataField=   "dblarealancada"
            Columns(4).NumberFormat=   "FormatText Event"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Vlr. m2 Serviço"
            Columns(5).DataField=   "dblvalorm2"
            Columns(5).NumberFormat=   "FormatText Event"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Vlr. Serviço"
            Columns(6).DataField=   "dblvalorservico"
            Columns(6).NumberFormat=   "FormatText Event"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Alíquota"
            Columns(7).DataField=   "dblaliquotaiss"
            Columns(7).NumberFormat=   "Standard"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Iss Devido"
            Columns(8).DataField=   "dblvalorlancto"
            Columns(8).NumberFormat=   "FormatText Event"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Iss Abatido"
            Columns(9).DataField=   "dblvalorabatido"
            Columns(9).NumberFormat=   "FormatText Event"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Iss A Pagar"
            Columns(10).DataField=   "dblSaldo"
            Columns(10).NumberFormat=   "FormatText Event"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Demolição"
            Columns(11).DataField=   "dblPorcDemolicao"
            Columns(11).NumberFormat=   "FormatText Event"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   12
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=12"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=1058"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=1640"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerStyle=0"
            Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=1588"
            Splits(0)._ColumnProps(13)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=1"
            Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(16)=   "Column(2).Width=2831"
            Splits(0)._ColumnProps(17)=   "Column(2).DividerStyle=0"
            Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2778"
            Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=0"
            Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(23)=   "Column(3).Width=1720"
            Splits(0)._ColumnProps(24)=   "Column(3).DividerStyle=0"
            Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=1667"
            Splits(0)._ColumnProps(27)=   "Column(3).AllowSizing=0"
            Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=0"
            Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(30)=   "Column(4).Width=1799"
            Splits(0)._ColumnProps(31)=   "Column(4).DividerStyle=0"
            Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=1746"
            Splits(0)._ColumnProps(34)=   "Column(4).AllowSizing=0"
            Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(37)=   "Column(5).Width=2170"
            Splits(0)._ColumnProps(38)=   "Column(5).DividerStyle=0"
            Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=2117"
            Splits(0)._ColumnProps(41)=   "Column(5).AllowSizing=0"
            Splits(0)._ColumnProps(42)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(44)=   "Column(6).Width=1958"
            Splits(0)._ColumnProps(45)=   "Column(6).DividerStyle=0"
            Splits(0)._ColumnProps(46)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(6)._WidthInPix=1905"
            Splits(0)._ColumnProps(48)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(50)=   "Column(7).Width=1191"
            Splits(0)._ColumnProps(51)=   "Column(7).DividerStyle=0"
            Splits(0)._ColumnProps(52)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(53)=   "Column(7)._WidthInPix=1138"
            Splits(0)._ColumnProps(54)=   "Column(7)._ColStyle=2"
            Splits(0)._ColumnProps(55)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(56)=   "Column(8).Width=1931"
            Splits(0)._ColumnProps(57)=   "Column(8).DividerStyle=0"
            Splits(0)._ColumnProps(58)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(59)=   "Column(8)._WidthInPix=1879"
            Splits(0)._ColumnProps(60)=   "Column(8)._ColStyle=2"
            Splits(0)._ColumnProps(61)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(62)=   "Column(9).Width=1958"
            Splits(0)._ColumnProps(63)=   "Column(9).DividerStyle=0"
            Splits(0)._ColumnProps(64)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(65)=   "Column(9)._WidthInPix=1905"
            Splits(0)._ColumnProps(66)=   "Column(9)._ColStyle=2"
            Splits(0)._ColumnProps(67)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(68)=   "Column(10).Width=2064"
            Splits(0)._ColumnProps(69)=   "Column(10).DividerStyle=0"
            Splits(0)._ColumnProps(70)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(71)=   "Column(10)._WidthInPix=2011"
            Splits(0)._ColumnProps(72)=   "Column(10)._ColStyle=2"
            Splits(0)._ColumnProps(73)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(74)=   "Column(11).Width=1455"
            Splits(0)._ColumnProps(75)=   "Column(11).DividerStyle=0"
            Splits(0)._ColumnProps(76)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(77)=   "Column(11)._WidthInPix=1402"
            Splits(0)._ColumnProps(78)=   "Column(11).AllowSizing=0"
            Splits(0)._ColumnProps(79)=   "Column(11)._ColStyle=2"
            Splits(0)._ColumnProps(80)=   "Column(11).Order=12"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            BorderStyle     =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   -2147483633
            RowDividerColor =   12648447
            RowSubDividerColor=   13160660
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000004&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000004&"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1,.transparentBmp=-1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=59,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=47,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=48"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=49"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=59"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=47,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=48"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=49"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=59"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=16,.parent=47,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=48"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=49"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=59"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=74,.parent=47,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=48"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=49"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=59"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=82,.parent=47,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=79,.parent=48"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=80,.parent=49"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=81,.parent=59"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=86,.parent=47,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=48"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=49"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=59"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=24,.parent=47,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=48"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=49"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=59"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=47,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=48"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=49"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=59"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=32,.parent=47,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=48"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=49"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=59"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=46,.parent=47,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=48"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=49"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=59"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=54,.parent=47,.alignment=1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=48"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=49"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=59"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=47,.alignment=1"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=48"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=49"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=59"
            _StyleDefs(84)  =   "Named:id=33:Normal"
            _StyleDefs(85)  =   ":id=33,.parent=0"
            _StyleDefs(86)  =   "Named:id=34:Heading"
            _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(88)  =   ":id=34,.wraptext=-1"
            _StyleDefs(89)  =   "Named:id=35:Footing"
            _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(91)  =   "Named:id=36:Selected"
            _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=37:Caption"
            _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(95)  =   "Named:id=38:HighlightRow"
            _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=39:EvenRow"
            _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(99)  =   "Named:id=40:OddRow"
            _StyleDefs(100) =   ":id=40,.parent=33"
            _StyleDefs(101) =   "Named:id=41:RecordSelector"
            _StyleDefs(102) =   ":id=41,.parent=34"
            _StyleDefs(103) =   "Named:id=42:FilterBar"
            _StyleDefs(104) =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_Parcelas 
         Caption         =   "Parcelas"
         Height          =   2070
         Left            =   -74385
         TabIndex        =   56
         Top             =   2415
         Width           =   10290
         Begin TrueOleDBGrid70.TDBGrid tdb_Parcelas 
            Height          =   1560
            Left            =   90
            TabIndex        =   57
            Top             =   300
            Width           =   9420
            _ExtentX        =   16616
            _ExtentY        =   2752
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Parcela"
            Columns(0).DataField=   "intParcela"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Acordo"
            Columns(1).DataField=   "strAcordo"
            Columns(1).NumberFormat=   "FormatText Event"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).DataField=   "strMoeda"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Valor"
            Columns(3).DataField=   "dblValor"
            Columns(3).NumberFormat=   "Standard"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Vencimento"
            Columns(4).DataField=   "dtmDtVencimento"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "D.A."
            Columns(5).DataField=   "intLancamentoAlfaDAtiva"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Baixa"
            Columns(6).DataField=   "dtmDtPagamento"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Descrição da Baixa"
            Columns(7).DataField=   "STRDESCRICAO"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Observação"
            Columns(8).DataField=   "Strobservacao"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1138"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
            Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=1085"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2672"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerStyle=0"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2619"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1058"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerStyle=0"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1005"
            Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1879"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerStyle=0"
            Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=1826"
            Splits(0)._ColumnProps(24)=   "Column(3).AllowSizing=0"
            Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(27)=   "Column(4).Width=1746"
            Splits(0)._ColumnProps(28)=   "Column(4).DividerStyle=0"
            Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=1693"
            Splits(0)._ColumnProps(31)=   "Column(4).AllowSizing=0"
            Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(34)=   "Column(5).Width=688"
            Splits(0)._ColumnProps(35)=   "Column(5).DividerStyle=0"
            Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=635"
            Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(39)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(40)=   "Column(6).Width=1535"
            Splits(0)._ColumnProps(41)=   "Column(6).DividerStyle=0"
            Splits(0)._ColumnProps(42)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(6)._WidthInPix=1482"
            Splits(0)._ColumnProps(44)=   "Column(6).AllowSizing=0"
            Splits(0)._ColumnProps(45)=   "Column(6)._ColStyle=1"
            Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(47)=   "Column(7).Width=3228"
            Splits(0)._ColumnProps(48)=   "Column(7).DividerStyle=0"
            Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=3175"
            Splits(0)._ColumnProps(51)=   "Column(7).AllowSizing=0"
            Splits(0)._ColumnProps(52)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(53)=   "Column(8).Width=5054"
            Splits(0)._ColumnProps(54)=   "Column(8).DividerStyle=0"
            Splits(0)._ColumnProps(55)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(56)=   "Column(8)._WidthInPix=5001"
            Splits(0)._ColumnProps(57)=   "Column(8).AllowSizing=0"
            Splits(0)._ColumnProps(58)=   "Column(8).Order=9"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            BorderStyle     =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   -2147483633
            RowDividerColor =   12648447
            RowSubDividerColor=   13160660
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000004&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000004&"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1,.transparentBmp=-1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=59,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=47,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=48"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=49"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=59"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=47"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=48"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=49"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=59"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=20,.parent=47,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=17,.parent=48"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=18,.parent=49"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=19,.parent=59"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=16,.parent=47,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=13,.parent=48"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=14,.parent=49"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=15,.parent=59"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=47,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=48"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=49"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=59"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=24,.parent=47,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=48"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=49"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=59"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=47,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=48"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=49"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=59"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=82,.parent=47"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=48"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=49"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=59"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=86,.parent=47"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=83,.parent=48"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=84,.parent=49"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=85,.parent=59"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   ":id=34,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   "Named:id=36:Selected"
            _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=37:Caption"
            _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(83)  =   "Named:id=38:HighlightRow"
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_Cabecalho 
         Height          =   4110
         Left            =   135
         TabIndex        =   23
         Top             =   360
         Width           =   11265
         Begin VB.TextBox txtdblvlIndexador 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9930
            MaxLength       =   20
            TabIndex        =   59
            Top             =   630
            Width           =   1245
         End
         Begin VB.TextBox txtstrIndexador 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6585
            TabIndex        =   58
            Top             =   630
            Width           =   1665
         End
         Begin VB.TextBox txtstrObservacoes 
            Height          =   1275
            Left            =   150
            MaxLength       =   300
            TabIndex        =   18
            Top             =   2700
            Width           =   10980
         End
         Begin VB.TextBox txtdtmData 
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
            Left            =   9915
            MaxLength       =   10
            TabIndex        =   9
            Top             =   195
            Width           =   1245
         End
         Begin VB.TextBox txtbitDigitoProcesso 
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
            Left            =   7950
            MaxLength       =   2
            TabIndex        =   8
            Top             =   195
            Width           =   285
         End
         Begin VB.TextBox txtintExercicioProcesso 
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
            Left            =   7440
            MaxLength       =   4
            TabIndex        =   7
            Top             =   195
            Width           =   465
         End
         Begin VB.TextBox txtstrCodigoProcesso 
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
            Left            =   6570
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   195
            Width           =   825
         End
         Begin MSDataListLib.DataCombo dbcintComposicao 
            Height          =   315
            Left            =   1065
            TabIndex        =   5
            Top             =   195
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcstrNomeProprietario 
            Height          =   315
            Left            =   1065
            TabIndex        =   10
            Top             =   615
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin TabDlg.SSTab tab_3dEnderecos 
            Height          =   1350
            Left            =   135
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1050
            Width           =   11010
            _ExtentX        =   19420
            _ExtentY        =   2381
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Endereço"
            TabPicture(0)   =   "frmCadLancamentoISS.frx":0038
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fra_EndImobiliario"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Endereço de Notificação"
            TabPicture(1)   =   "frmCadLancamentoISS.frx":0054
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame1"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame1 
               Height          =   930
               Left            =   -74910
               TabIndex        =   37
               Top             =   315
               Width           =   10830
               Begin VB.TextBox txtstrMunicipioC 
                  Height          =   300
                  Left            =   5205
                  MaxLength       =   50
                  TabIndex        =   47
                  Top             =   540
                  Width           =   2235
               End
               Begin VB.TextBox txtstrUFC 
                  Height          =   300
                  Left            =   8880
                  MaxLength       =   2
                  TabIndex        =   49
                  Top             =   540
                  Width           =   375
               End
               Begin VB.TextBox txtstrNumeroC 
                  Height          =   300
                  Left            =   6615
                  MaxLength       =   10
                  TabIndex        =   41
                  Top             =   180
                  Width           =   825
               End
               Begin VB.TextBox txtintCepC 
                  Height          =   300
                  Left            =   9660
                  MaxLength       =   9
                  TabIndex        =   51
                  Top             =   525
                  Width           =   1005
               End
               Begin VB.TextBox txtstrComplementoC 
                  Height          =   300
                  Left            =   9075
                  MaxLength       =   10
                  TabIndex        =   43
                  Top             =   180
                  Width           =   1590
               End
               Begin VB.TextBox txtstrLogradouroC 
                  Height          =   300
                  Left            =   1080
                  MaxLength       =   100
                  TabIndex        =   39
                  Top             =   180
                  Width           =   4065
               End
               Begin VB.TextBox txtstrBairroC 
                  Height          =   300
                  Left            =   675
                  MaxLength       =   50
                  TabIndex        =   45
                  Top             =   540
                  Width           =   2670
               End
               Begin VB.Label lbl_MunicipioC 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Município"
                  Height          =   195
                  Left            =   4455
                  TabIndex        =   46
                  Top             =   615
                  Width           =   705
               End
               Begin VB.Label lbl_UFC 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "UF"
                  Height          =   195
                  Left            =   8610
                  TabIndex        =   48
                  Top             =   630
                  Width           =   210
               End
               Begin VB.Label lbl_CepC 
                  AutoSize        =   -1  'True
                  Caption         =   "CEP"
                  Height          =   195
                  Left            =   9315
                  TabIndex        =   50
                  Top             =   615
                  Width           =   315
               End
               Begin VB.Label lbl_ComplementoC 
                  AutoSize        =   -1  'True
                  Caption         =   "Compl."
                  Height          =   195
                  Left            =   8550
                  TabIndex        =   42
                  Top             =   270
                  Width           =   480
               End
               Begin VB.Label lbl_NumeroC 
                  AutoSize        =   -1  'True
                  Caption         =   "N°"
                  Height          =   195
                  Left            =   6390
                  TabIndex        =   40
                  Top             =   270
                  Width           =   180
               End
               Begin VB.Label lbl_BairroC 
                  AutoSize        =   -1  'True
                  Caption         =   "Bairro"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   44
                  Top             =   630
                  Width           =   405
               End
               Begin VB.Label lbl_LogradouroC 
                  AutoSize        =   -1  'True
                  Caption         =   "Logradouro"
                  Height          =   195
                  Left            =   195
                  TabIndex        =   38
                  Top             =   270
                  Width           =   810
               End
            End
            Begin VB.Frame fra_EndImobiliario 
               Height          =   930
               Left            =   90
               TabIndex        =   29
               Top             =   315
               Width           =   10830
               Begin VB.TextBox txtstrBairro 
                  Height          =   300
                  Left            =   675
                  MaxLength       =   50
                  TabIndex        =   14
                  Top             =   540
                  Width           =   2670
               End
               Begin VB.TextBox txtstrLogradouro 
                  Height          =   300
                  Left            =   1080
                  MaxLength       =   100
                  TabIndex        =   11
                  Top             =   180
                  Width           =   4065
               End
               Begin VB.TextBox txtstrComplemento 
                  Height          =   300
                  Left            =   9075
                  MaxLength       =   10
                  TabIndex        =   13
                  Top             =   180
                  Width           =   1590
               End
               Begin VB.TextBox txtintCep 
                  Height          =   300
                  Left            =   9660
                  MaxLength       =   9
                  TabIndex        =   17
                  Top             =   525
                  Width           =   1005
               End
               Begin VB.TextBox txtstrNumero 
                  Height          =   300
                  Left            =   6615
                  MaxLength       =   10
                  TabIndex        =   12
                  Top             =   180
                  Width           =   825
               End
               Begin VB.TextBox txtstrMunicipio 
                  Height          =   300
                  Left            =   5205
                  MaxLength       =   50
                  TabIndex        =   15
                  Top             =   540
                  Width           =   2235
               End
               Begin VB.TextBox txtstrUf 
                  Height          =   300
                  Left            =   8880
                  MaxLength       =   2
                  TabIndex        =   16
                  Top             =   540
                  Width           =   375
               End
               Begin VB.Label lblintLogradouro 
                  AutoSize        =   -1  'True
                  Caption         =   "Logradouro"
                  Height          =   195
                  Left            =   195
                  TabIndex        =   30
                  Top             =   270
                  Width           =   810
               End
               Begin VB.Label lblintBairro 
                  AutoSize        =   -1  'True
                  Caption         =   "Bairro"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   33
                  Top             =   630
                  Width           =   405
               End
               Begin VB.Label lblintNumero 
                  AutoSize        =   -1  'True
                  Caption         =   "N°"
                  Height          =   195
                  Left            =   6390
                  TabIndex        =   31
                  Top             =   270
                  Width           =   180
               End
               Begin VB.Label lblstrComplemento 
                  AutoSize        =   -1  'True
                  Caption         =   "Compl."
                  Height          =   195
                  Left            =   8550
                  TabIndex        =   32
                  Top             =   270
                  Width           =   480
               End
               Begin VB.Label lblintCep 
                  AutoSize        =   -1  'True
                  Caption         =   "CEP"
                  Height          =   195
                  Left            =   9315
                  TabIndex        =   36
                  Top             =   615
                  Width           =   315
               End
               Begin VB.Label lblstrMunicipio 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Município"
                  Height          =   195
                  Left            =   4455
                  TabIndex        =   34
                  Top             =   615
                  Width           =   705
               End
               Begin VB.Label lblstrUf 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "UF"
                  Height          =   195
                  Left            =   8610
                  TabIndex        =   35
                  Top             =   630
                  Width           =   210
               End
            End
         End
         Begin VB.Label lbl_dblVlIndexador 
            AutoSize        =   -1  'True
            Caption         =   "Valor Indexador"
            Height          =   195
            Left            =   8775
            TabIndex        =   61
            Top             =   705
            Width           =   1110
         End
         Begin VB.Label lbl_strIndexador 
            AutoSize        =   -1  'True
            Caption         =   "Indexador"
            Height          =   195
            Left            =   5820
            TabIndex        =   60
            Top             =   720
            Width           =   705
         End
         Begin VB.Label lblRequerente 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte"
            Height          =   195
            Left            =   150
            TabIndex        =   27
            Top             =   705
            Width           =   840
         End
         Begin VB.Label lbl_strInscricaoAnterior 
            AutoSize        =   -1  'True
            Caption         =   "Observações"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   52
            Top             =   2475
            Width           =   945
         End
         Begin VB.Label lbldtmData 
            AutoSize        =   -1  'True
            Caption         =   "Lançamento"
            Height          =   195
            Left            =   8970
            TabIndex        =   26
            Top             =   285
            Width           =   885
         End
         Begin VB.Label lblstrProcesso 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Processo"
            Height          =   195
            Left            =   5865
            TabIndex        =   25
            Top             =   285
            Width           =   660
         End
         Begin VB.Label lblintComposicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   135
            TabIndex        =   24
            Top             =   285
            Width           =   870
         End
      End
   End
   Begin MSDataListLib.DataCombo dbcintExercicio 
      Height          =   315
      Left            =   6030
      TabIndex        =   2
      Top             =   90
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcstrEmissao 
      Height          =   315
      Left            =   10350
      TabIndex        =   4
      Top             =   90
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1455
      Left            =   75
      TabIndex        =   53
      Top             =   5280
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   2566
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
      Columns(1).Caption=   "intComposicaoDaReceita"
      Columns(1).DataField=   "intComposicaoDaReceita"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Composição Da Receita"
      Columns(2).DataField=   "strComposicaoDaReceita"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Inscrição Cadastral"
      Columns(3).DataField=   "strInscricao"
      Columns(3).NumberFormat=   "FormatText Event"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Exercício"
      Columns(4).DataField=   "intExercicio"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Área Predial"
      Columns(5).DataField=   "dblarealancada"
      Columns(5).NumberFormat=   "FormatText Event"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Aviso"
      Columns(6).DataField=   "strNumeroAviso"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Lançamento"
      Columns(7).DataField=   "dtmLancamento"
      Columns(7).NumberFormat=   "FormatText Event"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Total"
      Columns(8).DataField=   "dblValorTotal"
      Columns(8).NumberFormat=   "FormatText Event"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "PkidIss"
      Columns(9).DataField=   "PkidIss"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=6800"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=6720"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=3572"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3493"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1402"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1323"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2196"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2117"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1508"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1429"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=2117"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2037"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(49)=   "Column(8).Width=2117"
      Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=2037"
      Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(55)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(59)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(60)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bgcolor=&H80000009&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(21)  =   ":id=8,.fgcolor=&H8000000E&"
      _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.namedParent=38"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
      _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=74,.parent=13,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(80)  =   "Named:id=33:Normal"
      _StyleDefs(81)  =   ":id=33,.parent=0"
      _StyleDefs(82)  =   "Named:id=34:Heading"
      _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   ":id=34,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=36:Selected"
      _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=37:Caption"
      _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(91)  =   "Named:id=38:HighlightRow"
      _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=39:EvenRow"
      _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(95)  =   "Named:id=40:OddRow"
      _StyleDefs(96)  =   ":id=40,.parent=33"
      _StyleDefs(97)  =   "Named:id=41:RecordSelector"
      _StyleDefs(98)  =   ":id=41,.parent=34"
      _StyleDefs(99)  =   "Named:id=42:FilterBar"
      _StyleDefs(100) =   ":id=42,.parent=33"
   End
   Begin MSMask.MaskEdBox mskstrInscricao 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   24
      PromptChar      =   " "
   End
   Begin VB.Label lblstrNumeroAviso 
      AutoSize        =   -1  'True
      Caption         =   "Aviso"
      Height          =   195
      Left            =   7590
      TabIndex        =   20
      Top             =   180
      Width           =   390
   End
   Begin VB.Label lblstrInscricao 
      AutoSize        =   -1  'True
      Caption         =   "Inscrição Cadastral"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   1350
   End
   Begin VB.Label lblintExercicio 
      AutoSize        =   -1  'True
      Caption         =   "Exercício"
      Height          =   195
      Left            =   5310
      TabIndex        =   19
      Top             =   180
      Width           =   675
   End
   Begin VB.Label lblstrEmissao 
      AutoSize        =   -1  'True
      Caption         =   "Emissão"
      Height          =   195
      Left            =   9690
      TabIndex        =   21
      Top             =   180
      Width           =   585
   End
End
Attribute VB_Name = "frmCadLancamentoISS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnPrimeiraVez      As Boolean

Private Sub dbcintComposicao_Click(Area As Integer)
    DropDownDataCombo dbcintComposicao, Me, Area
End Sub

Private Sub dbcintComposicao_GotFocus()
    MarcaCampo dbcintComposicao
End Sub

Private Sub dbcintComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintComposicao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintComposicao
End Sub

Private Sub dbcintExercicio_Click(Area As Integer)
    DropDownDataCombo dbcintExercicio, Me, Area
End Sub
Private Sub dbcintExercicio_GotFocus()
    MarcaCampo dbcintExercicio
End Sub
Private Sub dbcintExercicio_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintExercicio, Me, , KeyCode, Shift
End Sub
Private Sub dbcintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintExercicio
End Sub

Private Sub dbcstrEmissao_Click(Area As Integer)
    DropDownDataCombo dbcstrEmissao, Me, Area
End Sub
Private Sub dbcstrEmissao_GotFocus()
    MarcaCampo dbcstrEmissao
End Sub
Private Sub dbcstrEmissao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcstrEmissao, Me, , KeyCode, Shift
End Sub

Private Sub dbcstrEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcstrEmissao
End Sub

Private Sub dbcstrNomeProprietario_Click(Area As Integer)
    DropDownDataCombo dbcstrNomeProprietario, Me, Area
End Sub

Private Sub dbcstrNomeProprietario_GotFocus()
    MarcaCampo dbcstrNomeProprietario
End Sub

Private Sub dbcstrNomeProprietario_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcstrNomeProprietario, Me, , KeyCode, Shift
End Sub

Private Sub dbcstrNomeProprietario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcstrNomeProprietario
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1242
    VerificaMascaraInscricao
End Sub

Private Sub Form_Load()

    dbcintComposicao.Tag = strQueryComposicao & ";strDescricao"
    dbcintExercicio.Tag = strQueryExercicio & ";intExercicio"
    dbcstrEmissao.Tag = strQueryEmissao & ";strEmissao"
    dbcstrNomeProprietario.Tag = strQueryRequerente & ";strNome"
    
    TrocaCorObjeto txtstrIndexador, True
    TrocaCorObjeto txtdblvlIndexador, True
    TrocaCorObjeto txtstrLogradouro, True
    TrocaCorObjeto txtstrLogradouroC, True
    TrocaCorObjeto txtstrNumero, True
    TrocaCorObjeto txtstrNumeroC, True
    TrocaCorObjeto txtstrComplemento, True
    TrocaCorObjeto txtstrComplementoC, True
    TrocaCorObjeto txtstrBairro, True
    TrocaCorObjeto txtstrBairroC, True
    TrocaCorObjeto txtstrMunicipio, True
    TrocaCorObjeto txtstrMunicipioC, True
    TrocaCorObjeto txtstrUf, True
    TrocaCorObjeto txtstrUFC, True
    TrocaCorObjeto txtintCep, True
    TrocaCorObjeto txtintCepC, True
    TrocaCorObjeto txtstrObservacoes, True
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    blnPrimeiraVez = False
End Sub

Private Sub mskstrInscricao_GotFocus()
    MarcaCampo mskstrInscricao
End Sub

Sub VerificaMascaraInscricao()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
Dim strMascara   As String
    
    strMascara = ""
    strSQL = ""
    strSQL = strSQL & "Select * From " & gstrCampoDeInscricao & " "
    strSQL = strSQL & "Where intTipoDeInscricao = " & TYP_ISS_CONSTRUCAO
    strSQL = strSQL & "Order By intSequencia"
    
    Set gobjBanco = New clsBanco
        
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    
    mskstrInscricao.Mask = strMascara

End Sub

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao
End Sub

Private Sub tdb_Lista_Click()
    blnPrimeiraVez = True
End Sub

Private Sub tdb_Lista_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 7 Then
        Value = gstrDataFormatada(Value)
    ElseIf ColIndex = 8 Or ColIndex = 5 Then
        Value = gstrConvVrDoSql(Value, 2)
    ElseIf ColIndex = 4 Then
        Value = gstrFormataInscricao(CStr(Value), TYP_ISS_CONSTRUCAO)
    End If
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If blnPrimeiraVez Then
        If Not tdb_Lista.EOF Then
            If tdb_Lista.Columns("Pkid").Value > 0 Then
                PreencheLancamentoISS
                PreenchePredios
            End If
        End If
    End If
End Sub

Private Sub tdb_Parcelas_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 2 Then
        If Len(Value) Then
            Value = gstrFormataInscricao(CStr(Value), TYP_ACORDO)
        End If
    End If
End Sub

Private Sub tdb_Predios_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Select Case ColIndex
        Case 4, 5, 6, 8, 9, 10
            Value = gstrConvVrDoSql(Value, 2)
        Case 7, 11
            Value = gstrConvVrDoSql(Value, 4)
    End Select
End Sub

Private Sub txtbitDigitoProcesso_GotFocus()
    MarcaCampo txtbitDigitoProcesso
End Sub

Private Sub txtbitDigitoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigitoProcesso
End Sub

Private Sub txtdtmData_GotFocus()
    MarcaCampo txtdtmData
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmData
End Sub

Private Sub txtdtmData_LostFocus()
    txtdtmData = gstrDataFormatada(txtdtmData)
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub

Private Sub txtintCepC_GotFocus()
    MarcaCampo txtintCepC
End Sub

Private Sub txtintCepC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepC
End Sub

Private Sub txtintExercicioProcesso_GotFocus()
    MarcaCampo txtintExercicioProcesso
End Sub

Private Sub txtintExercicioProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicioProcesso
End Sub

Private Sub txtstrCodigoProcesso_GotFocus()
    MarcaCampo txtstrCodigoProcesso
End Sub

Private Sub txtstrCodigoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigoProcesso
End Sub

Private Sub txtstrNumero_GotFocus()
    MarcaCampo txtstrNumero
End Sub

Private Sub txtstrNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumero
End Sub

Private Sub txtstrNumeroC_Change()
    MarcaCampo txtstrNumeroC
End Sub

Private Sub txtstrNumeroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumeroC
End Sub

Private Sub txtstrBairro_GotFocus()
    MarcaCampo txtstrBairro
End Sub

Private Sub txtstrBairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrBairro
End Sub

Private Sub txtstrBairroC_GotFocus()
    MarcaCampo txtstrBairroC
End Sub

Private Sub txtstrBairroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrBairroC
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
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

Private Sub txtstrLogradouro_GotFocus()
    MarcaCampo txtstrLogradouro
End Sub

Private Sub txtstrLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLogradouro
End Sub

Private Sub txtstrLogradouroC_GotFocus()
    MarcaCampo txtstrLogradouroC
End Sub

Private Sub txtstrLogradouroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLogradouroC
End Sub

Private Sub txtstrUFC_GotFocus()
    MarcaCampo txtstrUFC
End Sub

Private Sub txtstrUFC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrUFC
End Sub

Private Sub txtstrNumeroAviso_GotFocus()
    MarcaCampo txtstrNumeroAviso
End Sub

Private Sub txtstrNumeroAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNumeroAviso
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        Case Is = UCase(gstrLocalizar)
            blnPrimeiraVez = True
            LeDaTabelaParaObj "", tdb_Lista, strQueryLocalizar
        Case Is = UCase(gstrNovo)
            LimpaObjeto Me
            Set tdb_Predios.DataSource = Nothing
            Set tdb_Parcelas.DataSource = Nothing
            tab_3dPasta.Tab = 0
            If mskstrInscricao.Enabled Then mskstrInscricao.SetFocus
    End Select
End Sub

Private Function strQueryLocalizar() As String
Dim strSQL As String
    
    strSQL = "SELECT LA.Pkid Pkid,"
    strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ISS_CONSTRUCAO)) & " strInscricao, "
    strSQL = strSQL & " LA.intComposicaoDaReceita,"
    strSQL = strSQL & " LA.strComposicaoDaReceita,"
    strSQL = strSQL & " LA.intExercicio,"
    strSQL = strSQL & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso,"
    strSQL = strSQL & " LI.PKid PkidISS,"
    strSQL = strSQL & " LI.dtmLancamento,"
    strSQL = strSQL & " Sum(LIP.dblarealancada) as dblarealancada,"
    strSQL = strSQL & " ( (CASE WHEN SUM(LIP.dblValorLancto) IS NULL THEN 0 ELSE SUM(LIP.dblValorLancto) END) - "
    strSQL = strSQL & "(CASE WHEN SUM(LIP.dblValorAbatido)IS NULL THEN 0 ELSE SUM(LIP.dblValorAbatido) END)) dblValorTotal"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLanctoIssConstrucao & " LI, "
    strSQL = strSQL & gstrLanctoIssConstrucaoPredios & " LIP "
    strSQL = strSQL & " WHERE LA.Pkid = LI.intLancamentoAlfa AND LIP.intLanctoIssConstrucao" & strOUTJOracle & "=" & strOUTJSQLServer & "LI.Pkid"
    
    
    If Len(mskstrInscricao.Text) > 0 Then strSQL = strSQL & " AND strInscricao LIKE " & "'" & UCase(String(gintLenInscricao - gintRetornaTamanhoMascara(TYP_IMOBILIARIA), "0") & mskstrInscricao) & "%'"
    If Len(dbcintExercicio.Text) > 0 Then strSQL = strSQL & " AND intExercicio LIKE " & "'" & UCase(dbcintExercicio.Text) & "%'"
    If Len(dbcstrEmissao.Text) > 0 Then strSQL = strSQL & " AND strEmissao LIKE " & "'" & UCase(dbcstrEmissao.Text) & "%'"
    If Len(txtstrNumeroAviso.Text) > 0 Then strSQL = strSQL & " AND strNumeroAviso LIKE " & "'" & UCase(String(gintLenNumAviso - Len(txtstrNumeroAviso), "0") & txtstrNumeroAviso.Text) & "'"
    If Len(dbcintComposicao.Text) > 0 Then strSQL = strSQL & " AND UPPER(strComposicaoDaReceita) LIKE " & "'" & UCase(dbcintComposicao.Text) & "%'"
    If Len(txtstrCodigoProcesso.Text) > 0 Then strSQL = strSQL & " AND LI.strCodigoProcesso = " & "'" & UCase(txtstrCodigoProcesso.Text) & "'"
    If Len(txtbitDigitoProcesso.Text) > 0 Then strSQL = strSQL & " AND LI.bitDigitoProcesso = " & UCase(txtbitDigitoProcesso.Text)
    If Len(txtintExercicioProcesso.Text) > 0 Then strSQL = strSQL & " AND LI.intExercicioProcesso = " & UCase(txtintExercicioProcesso.Text)
    If Len(dbcstrNomeProprietario.Text) > 0 Then strSQL = strSQL & " AND UPPER(strNomeProprietario) LIKE " & "'" & UCase(dbcstrNomeProprietario.Text) & "%'"
    If Len(txtdtmData.Text) > 0 Then strSQL = strSQL & " AND LI.dtmLancamento = '" & txtdtmData.Text & "'"
    
    strSQL = strSQL & " GROUP BY LA.Pkid, LA.strInscricao, LA.intComposicaoDaReceita, LA.strComposicaoDaReceita, LA.intExercicio, LA.strNumeroAviso, LI.PKid, LI.dtmLancamento"
    strSQL = strSQL & " ORDER BY LA.intComposicaoDaReceita,LA.strInscricao,LA.intExercicio"
    
    strQueryLocalizar = strSQL

End Function

Private Function strQueryExercicio() As String
Dim strSQL As String
    
    strSQL = "SELECT DISTINCT 1, intExercicio"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParametroIPTU
    strSQL = strSQL & " ORDER BY intExercicio"
    
    strQueryExercicio = strSQL
    
End Function

Private Function strQueryComposicao() As String
Dim strSQL As String
    
    strSQL = "SELECT Pkid,"
    strSQL = strSQL & " strDescricao Descricao"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita
    strSQL = strSQL & " WHERE bytDividaAtiva = 1 AND intUtilizacao = " & TYP_ISS_CONSTRUCAO
    strSQL = strSQL & " ORDER BY strDescricao"
    
    strQueryComposicao = strSQL

End Function

Private Function strQueryEmissao() As String
Dim strSQL As String
    
    strSQL = "SELECT DISTINCT 1, strEmissao "
    'String(gintLenEmissao - Len(strValor), "0") & strValor
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParametroIPTU
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_numeric, "strEmissao")
    
    strQueryEmissao = strSQL
    
End Function

Private Function strQueryRequerente() As String
Dim strSQL As String

    strSQL = "SELECT Pkid,"
    strSQL = strSQL & " strNome"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrContribuinte
    strSQL = strSQL & " ORDER BY"
    strSQL = strSQL & " strNome"
    
    strQueryRequerente = strSQL
    
End Function


Private Sub PreencheLancamentoISS()
Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
    
    strSQL = "SELECT LA.strComposicaoDaReceita,"
    strSQL = strSQL & " LA.intExercicio,"
    strSQL = strSQL & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso,"
    strSQL = strSQL & " LA.strEmissao,"
    strSQL = strSQL & " LA.strNomeProprietario,"
    strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ISS_CONSTRUCAO)) & " strInscricao, "
    strSQL = strSQL & " LA.strIndexador,"
    strSQL = strSQL & " LA.dblvlIndexador,"
    strSQL = strSQL & " LA.strLogradouro,"
    strSQL = strSQL & " LA.strNumero,"
    strSQL = strSQL & " LA.strComplemento,"
    strSQL = strSQL & " LA.strBairro,"
    strSQL = strSQL & " LA.strMunicipio,"
    strSQL = strSQL & " LA.strUf,"
    strSQL = strSQL & " LA.intCep,"
    strSQL = strSQL & " LA.strLogradouroC,"
    strSQL = strSQL & " LA.strNumeroC,"
    strSQL = strSQL & " LA.strComplementoC,"
    strSQL = strSQL & " LA.strBairroC,"
    strSQL = strSQL & " LA.strMunicipioC,"
    strSQL = strSQL & " LA.strUfC,"
    strSQL = strSQL & " LA.intCepC,"
    strSQL = strSQL & " LI.strCodigoProcesso,"
    strSQL = strSQL & " LI.intExercicioProcesso,"
    strSQL = strSQL & " LI.bitDigitoProcesso,"
    strSQL = strSQL & " LI.strObservacoes,"
    strSQL = strSQL & " LI.dtmLancamento"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, " & gstrLanctoIssConstrucao & " LI "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " LA.Pkid = " & tdb_Lista.Columns("Pkid").Value & " AND LI.intLancamentoAlfa = LA.Pkid"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dbcintComposicao.Text = Space$(0) & adoResultado!strComposicaoDaReceita
            dbcintExercicio.Text = Space$(0) & adoResultado!intExercicio
            txtstrNumeroAviso.Text = Space$(0) & adoResultado!strNumeroAviso
            dbcstrEmissao.Text = Space$(0) & adoResultado!strEmissao
            mskstrInscricao.Text = Space$(0) & adoResultado!strInscricao
            txtstrIndexador.Text = Space$(0) & adoResultado!Strindexador
            txtdblvlIndexador.Text = Space$(0) & adoResultado!dblvlIndexador
            txtstrCodigoProcesso.Text = Space$(0) & adoResultado!strCodigoProcesso
            txtintExercicioProcesso.Text = Space$(0) & adoResultado!intExercicioProcesso
            txtbitDigitoProcesso.Text = Space$(0) & adoResultado!bitDigitoProcesso
            txtdtmData.Text = Space$(0) & adoResultado!dtmLancamento
            dbcstrNomeProprietario.Text = Space$(0) & adoResultado!strnomeproprietario
            txtstrLogradouro.Text = Space$(0) & adoResultado!strLogradouro
            txtstrNumero.Text = Space$(0) & adoResultado!strNumero
            txtstrComplemento.Text = Space$(0) & adoResultado!STRCOMPLEMENTO
            txtstrBairro.Text = Space$(0) & adoResultado!strBairro
            txtstrMunicipio.Text = Space$(0) & adoResultado!STRMUNICIPIO
            txtstrUf.Text = Space$(0) & adoResultado!STRUF
            txtintCep.Text = Space$(0) & gstrCEPFormatado(adoResultado!INTCEP)
            txtstrLogradouroC.Text = Space$(0) & adoResultado!strlogradouroc
            txtstrNumeroC.Text = Space$(0) & adoResultado!strNumeroC
            txtstrComplementoC.Text = Space$(0) & adoResultado!strComplementoC
            txtstrBairroC.Text = Space$(0) & adoResultado!strBairroC
            txtstrMunicipioC.Text = Space$(0) & adoResultado!strMunicipioC
            txtstrUFC.Text = Space$(0) & adoResultado!strUFC
            txtintCepC.Text = Space$(0) & gstrCEPFormatado(adoResultado!intcepc)
            txtstrObservacoes.Text = Space$(0) & adoResultado!strObservacoes
        End If
    End If

End Sub

Private Sub PreenchePredios()
Dim strSQL As String

    strSQL = "SELECT dtmDataConstrucao, "
    strSQL = strSQL & " strTipoConstrucao, "
    strSQL = strSQL & " strTipoAcabamento, "
    strSQL = strSQL & " dblPorcDemolicao, "
    strSQL = strSQL & " dblarealancada, "
    strSQL = strSQL & " dblvalorm2, "
    strSQL = strSQL & "dblvalorservico, "
    strSQL = strSQL & "dblaliquotaiss, "
    strSQL = strSQL & "dblvalorlancto, "
    strSQL = strSQL & "dblvalorabatido, "
    strSQL = strSQL & "(" & gstrCONVERT(CDT_INT, gstrISNULL("dblvalorlancto", "0")) & " - " & gstrCONVERT(CDT_INT, gstrISNULL("dblvalorabatido", "0")) & ") dblSaldo "
    strSQL = strSQL & "FROM " & gstrLanctoIssConstrucaoPredios
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " intlanctoIssConstrucao = " & Val(tdb_Lista.Columns("PkidISS").Value)
    
    LeDaTabelaParaObj "", tdb_Predios, strSQL
    
    CarregaParcelas
    
End Sub

Private Sub CarregaParcelas()
Dim strSQL As String
    
    If bytDBType = Oracle Then
        strSQL = "SELECT LV.intParcela,"
        strSQL = strSQL & " LV.dblValor,"
        strSQL = strSQL & " LV.dtmDtVencimento,"
        strSQL = strSQL & "CASE WHEN LV.intLancamentoAlfaDAtiva IS NULL THEN '' ELSE 'X' END intLancamentoAlfaDAtiva ,"
        strSQL = strSQL & " LP.dtmDtPagamento,"
        strSQL = strSQL & " CB.STRDESCRICAO, "
        strSQL = strSQL & " LP.Strobservacao, "
        strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ACORDO)) & " strAcordo "
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrCodigoDeBaixa & " CB, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP,"
        strSQL = strSQL & gstrLancamentoAlfa & " LA "
        strSQL = strSQL & " WHERE LV.Pkid " & strOUTJSQLServer & "=" & " LP.intLancamentoValor " & strOUTJOracle
        strSQL = strSQL & " AND CB.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " LP.Intcodigobaixa  AND"
        strSQL = strSQL & " LV.Intlancamentoalfaacordo " & strOUTJSQLServer & "= LA.pkid " & strOUTJOracle & " AND "
        strSQL = strSQL & " LV.intLancamentoAlfa = " & Val(tdb_Lista.Columns("Pkid").Value)
        strSQL = strSQL & " ORDER BY LV.intParcela"
    Else
        strSQL = "SELECT "
        strSQL = strSQL & " LV.intParcela, "
        strSQL = strSQL & " LV.dblValor, "
        strSQL = strSQL & " LV.dtmDtVencimento, "
        strSQL = strSQL & " CASE WHEN LV.intLancamentoAlfaDAtiva IS NULL THEN '' ELSE 'X' END intLancamentoAlfaDAtiva , "
        strSQL = strSQL & " LP.dtmDtPagamento, "
        strSQL = strSQL & " CB.STRDESCRICAO, "
        strSQL = strSQL & " LP.strObservacao, "
        strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ACORDO)) & " strAcordo "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrLancamentoValor & " LV LEFT OUTER JOIN "
        strSQL = strSQL & gstrLancamentoPagamento & " LP ON LV.PKId = LP.intLancamentoValor LEFT OUTER JOIN "
        strSQL = strSQL & gstrCodigoDeBaixa & " CB ON LP.intcodigobaixa = CB.PKID LEFT OUTER JOIN "
        strSQL = strSQL & gstrLancamentoAlfa & " LA ON LV.intLancamentoAlfaAcordo = LA.PKId "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & " LV.intLancamentoAlfa = " & Val(tdb_Lista.Columns("Pkid").Value)
        strSQL = strSQL & "ORDER BY "
        strSQL = strSQL & " LV.intParcela "
    End If
    
    LeDaTabelaParaObj "", tdb_Parcelas, strSQL
    
End Sub

