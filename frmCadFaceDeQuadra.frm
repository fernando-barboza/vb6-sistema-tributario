VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadFaceDeQuadra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faces de Quadra"
   ClientHeight    =   6840
   ClientLeft      =   1860
   ClientTop       =   2115
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7020
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4560
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   8043
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Faces de Quadra"
      TabPicture(0)   =   "frmCadFaceDeQuadra.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrSetor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrQuadra"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_strBairro"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintLogradouro"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPKId"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtstrSetor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstrQuadra"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtstrSequenciaDeFace"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_Logradouro"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fra_HistoricoFaceDeQuadra"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_strBairro"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Equipamentos"
      TabPicture(1)   =   "frmCadFaceDeQuadra.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_dblSomaDosFatores"
      Tab(1).Control(1)=   "cmd_Remover"
      Tab(1).Control(2)=   "cmd_Adicionar"
      Tab(1).Control(3)=   "txt_intAno"
      Tab(1).Control(4)=   "lvw_MelhoramentosCadastrados"
      Tab(1).Control(5)=   "lvw_Melhoramentos"
      Tab(1).Control(6)=   "lbl_dblSomaDosFatores"
      Tab(1).Control(7)=   "lbl_Ano"
      Tab(1).Control(8)=   "lbl_Melhoramentos"
      Tab(1).Control(9)=   "lbl_MelhoramentosExistentes"
      Tab(1).ControlCount=   10
      Begin VB.TextBox txt_dblSomaDosFatores 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -69060
         TabIndex        =   26
         Top             =   3840
         Width           =   795
      End
      Begin VB.TextBox txt_strBairro 
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
         Left            =   1605
         MaxLength       =   100
         TabIndex        =   9
         Top             =   1410
         Width           =   4005
      End
      Begin VB.Frame fra_HistoricoFaceDeQuadra 
         Caption         =   "Histórico da Face de Quadra"
         Height          =   2115
         Left            =   105
         TabIndex        =   22
         Top             =   2310
         Width           =   6705
         Begin TrueOleDBGrid70.TDBDropDown tdd_ValorMetroTerreno 
            Height          =   1305
            Left            =   1230
            TabIndex        =   24
            Top             =   540
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   2302
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
            Columns(1).Caption=   "Exercício"
            Columns(1).DataField=   "intExercicio"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Código"
            Columns(2).DataField=   "intCodigo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Valor Metro Terreno"
            Columns(3).DataField=   "dblValor"
            Columns(3).NumberFormat=   "FormatText Event"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1376"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1296"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1296"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1217"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2196"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2117"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
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
            ListField       =   "dblValor"
            DataField       =   ""
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
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
         Begin TrueOleDBGrid70.TDBGrid tdb_HistoricoFaceDeQuadra 
            Height          =   1695
            Left            =   405
            TabIndex        =   23
            Top             =   270
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   2990
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Exercício"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Valor Metro Terreno"
            Columns(2).DataField=   ""
            Columns(2).NumberFormat=   "FormatText Event"
            Columns(2).DropDown=   "tdd_ValorMetroTerreno"
            Columns(2).DropDown.vt=   8
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "intValorMetroTerreno"
            Columns(3).DataField=   "intValorMetroTerreno"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Valor Metro Terreno com Equipamentos"
            Columns(4).DataField=   ""
            Columns(4).NumberFormat=   "FormatText Event"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=1376"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1296"
            Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=1"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(21)=   "Column(2).AutoCompletion=1"
            Splits(0)._ColumnProps(22)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(28)=   "Column(4).Width=5106"
            Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=5027"
            Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
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
      Begin VB.CommandButton cmd_Remover 
         Height          =   465
         Left            =   -72105
         MouseIcon       =   "frmCadFaceDeQuadra.frx":0038
         MousePointer    =   99  'Custom
         Picture         =   "frmCadFaceDeQuadra.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Remover"
         Top             =   1395
         Width           =   465
      End
      Begin VB.CommandButton cmd_Adicionar 
         Height          =   465
         Left            =   -72105
         MouseIcon       =   "frmCadFaceDeQuadra.frx":0784
         MousePointer    =   99  'Custom
         Picture         =   "frmCadFaceDeQuadra.frx":0A8E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Adicionar"
         Top             =   855
         Width           =   465
      End
      Begin VB.TextBox txt_intAno 
         Height          =   285
         Left            =   -72825
         MaxLength       =   4
         TabIndex        =   13
         Top             =   3840
         Width           =   645
      End
      Begin VB.CommandButton cmd_Logradouro 
         Height          =   300
         Left            =   5280
         Picture         =   "frmCadFaceDeQuadra.frx":0ED0
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "584"
         ToolTipText     =   "Ativa Cadastro de Logradouro"
         Top             =   1005
         Width           =   330
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
         Left            =   1605
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1800
         Width           =   930
      End
      Begin VB.TextBox txtstrQuadra 
         Alignment       =   1  'Right Justify
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
         Left            =   3645
         MaxLength       =   4
         TabIndex        =   4
         Top             =   645
         Width           =   930
      End
      Begin VB.TextBox txtstrSetor 
         Alignment       =   1  'Right Justify
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
         Left            =   1605
         MaxLength       =   2
         TabIndex        =   2
         Top             =   645
         Width           =   930
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2775
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   12
         Top             =   30
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dbcintLogradouro 
         Height          =   315
         Left            =   1605
         TabIndex        =   6
         Top             =   1005
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComctlLib.ListView lvw_MelhoramentosCadastrados 
         Height          =   2910
         Left            =   -74910
         TabIndex        =   16
         Top             =   825
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   5133
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
      Begin MSComctlLib.ListView lvw_Melhoramentos 
         Height          =   2910
         Left            =   -71535
         TabIndex        =   17
         Top             =   840
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   5133
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
      Begin VB.Label lbl_dblSomaDosFatores 
         Caption         =   "Soma dos Fatores"
         Height          =   255
         Left            =   -70440
         TabIndex        =   25
         Top             =   3930
         Width           =   1485
      End
      Begin VB.Label lbl_strBairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   1110
         TabIndex        =   8
         Top             =   1485
         Width           =   405
      End
      Begin VB.Label lbl_Ano 
         AutoSize        =   -1  'True
         Caption         =   "Ano do melhoramento"
         Height          =   195
         Left            =   -74460
         TabIndex        =   20
         Top             =   3915
         Width           =   1545
      End
      Begin VB.Label lbl_Melhoramentos 
         AutoSize        =   -1  'True
         Caption         =   "Melhoramentos cadastrados"
         Height          =   195
         Left            =   -74910
         TabIndex        =   19
         Top             =   585
         Width           =   1995
      End
      Begin VB.Label lbl_MelhoramentosExistentes 
         AutoSize        =   -1  'True
         Caption         =   "Melhoramentos existentes na seção:"
         Height          =   195
         Left            =   -71565
         TabIndex        =   18
         Top             =   585
         Width           =   2580
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Logradouro"
         Height          =   195
         Left            =   705
         TabIndex        =   5
         Top             =   1060
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sequência de face"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1860
         Width           =   1350
      End
      Begin VB.Label lblstrQuadra 
         AutoSize        =   -1  'True
         Caption         =   "Quadra"
         Height          =   195
         Left            =   3015
         TabIndex        =   3
         Top             =   675
         Width           =   525
      End
      Begin VB.Label lblstrSetor 
         AutoSize        =   -1  'True
         Caption         =   "Setor"
         Height          =   195
         Left            =   1080
         TabIndex        =   1
         Top             =   705
         Width           =   375
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   2130
      Left            =   60
      TabIndex        =   21
      Top             =   4665
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3757
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "PKId"
      Columns(0).DataField=   "PKID"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Setor"
      Columns(1).DataField=   "strSetor"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Quadra"
      Columns(2).DataField=   "strQuadra"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Logradouro"
      Columns(3).DataField=   "Logradouro"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Bairro"
      Columns(4).DataField=   "Bairro"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Sequência de Face"
      Columns(5).DataField=   "strSequenciaDeFace"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "intLogradouro"
      Columns(6).DataField=   "intLogradouro"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Logradouro"
      Columns(7).DataField=   "bytCancelado"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=423"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=344"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1429"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1349"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1402"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1323"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=3810"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3731"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2355"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2275"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=2619"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2540"
      Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(36)=   "Column(6).Width=159"
      Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=79"
      Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(45)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
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
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadFaceDeQuadra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando     As Boolean
Dim mobjAux           As Object
Dim mblnSelecionou    As Boolean
Dim mblnClickOk       As Boolean
Dim objList           As Object
Dim bytOrdenacao      As Byte
Dim blnOrdenacaoAsc   As Boolean
Dim mblnPrimeiraVez   As Boolean
    
Dim lngLograAtual     As Long
Dim strQuadraAtual    As String
Dim strSetorAtual     As String
Dim strSeqAtual       As String
        
Private Function strQuery() As String
Dim strSQL  As String
    
    strSQL = ""
    
    strSQL = strSQL & "SELECT FQ.PKId, FQ.strSetor, FQ.strQuadra, FQ.strSequenciaDeFace, "
    strSQL = strSQL & "LO.strDescricao as Logradouro, FQ.intLogradouro, BA.strDescricao Bairro, "
    strSQL = strSQL & gstrISNULL("LO.dtmdtExclusao", "'Não Cancelado'", "'Cancelado'") & " bytCancelado "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrFaceDeQuadra & " FQ, "
    strSQL = strSQL & gstrLogradouro & " LO, "
    strSQL = strSQL & gstrBairro & " BA "
    strSQL = strSQL & "WHERE LO.PKId " & strOUTJOracle & "=" & " FQ.intLogradouro"
    strSQL = strSQL & " AND BA.Pkid " & strOUTJOracle & "=" & " LO.intBairro"
    'strSql = strSql & " AND LO.Dtmdtexclusao is null "
    
   
    Select Case bytOrdenacao
        Case Is = 1
            strSQL = strSQL & " ORDER BY strSetor" & IIf(blnOrdenacaoAsc, " ASC", " DESC") & ", strQuadra"
        Case Is = 2
            strSQL = strSQL & " ORDER BY strQuadra" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strSQL = strSQL & " ORDER BY Logradouro" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 4
            strSQL = strSQL & " ORDER BY Bairro" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 5
            strSQL = strSQL & " ORDER BY strSequenciaDeFace" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 7
            strSQL = strSQL & " ORDER BY bytCancelado" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
   
    strQuery = strSQL
End Function

Private Sub cmd_Adicionar_Click()
    AdicionarMelhoramento
    GravaMelhoramentos Val(txtPKId)
End Sub

Private Sub cmd_Logradouro_Click()
    CarregaForm frmCadLogradouro, dbcintLogradouro
End Sub

Private Sub cmd_Remover_Click()
    RemoverMelhoramento
    GravaMelhoramentos Val(txtPKId)
End Sub

Private Sub dbcintLogradouro_Change()
    txt_strBairro.Text = Mid(dbcintLogradouro.Text, InStr(1, dbcintLogradouro.Text, "->") + 3, Len(dbcintLogradouro.Text) - InStr(1, dbcintLogradouro.Text, "->"))
End Sub

Private Sub dbcintLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintLogradouro
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1026
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

Private Sub Form_Load()
   bytOrdenacao = 1: blnOrdenacaoAsc = True
   TrocaCorObjeto txt_strBairro, True
   VerificaObjParaAplicar mobjAux
   dbcintLogradouro.Tag = strQueryLogradouro & ";L.strDescricao"
   MontaColumnHeaders
   VerificaListaAutomatica gstrMelhoramentoPublico, lvw_MelhoramentosCadastrados, "PKId, strNomeDoMelhoramento, dblFator"
   LimpaGrdHistorico
   mblnAlterando = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub lvw_Melhoramentos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaColunaClicada lvw_Melhoramentos, ColumnHeader
End Sub

Private Sub lvw_Melhoramentos_GotFocus()
    tab_3dPasta.Tab = 1
End Sub

Private Sub lvw_Melhoramentos_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_Melhoramentos
End Sub

Private Sub lvw_MelhoramentosCadastrados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaColunaClicada lvw_MelhoramentosCadastrados, ColumnHeader
End Sub

Private Sub lvw_MelhoramentosCadastrados_DblClick()
    cmd_Adicionar_Click
    GravaMelhoramentos Val(txtPKId)
End Sub

Private Sub lvw_MelhoramentosCadastrados_GotFocus()
    tab_3dPasta.Tab = 1
End Sub

Private Sub lvw_MelhoramentosCadastrados_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_MelhoramentosCadastrados
End Sub

Private Sub tdb_HistoricoFaceDeQuadra_ButtonClick(ByVal ColIndex As Integer)
  If ColIndex = 2 Then
     MontaArrayValorMetroTerreno
  End If
End Sub

Private Sub tdb_HistoricoFaceDeQuadra_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 2 Then Value = gstrConvVrDoSql(Value, 5)
End Sub

Private Sub tdb_HistoricoFaceDeQuadra_KeyPress(KeyAscii As Integer)
    Select Case tdb_HistoricoFaceDeQuadra.Col
        Case Is = 1
            CaracterValido KeyAscii, "N", tdb_HistoricoFaceDeQuadra
        Case Is = 2
            If KeyAscii = 9 Then
                MontaArrayValorMetroTerreno
            End If
            KeyAscii = 0
        Case Is = 4
            KeyAscii = 0
    End Select
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
    bytOrdenacao = ColIndex
    blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    Select Case tdb_Lista.Col
        Case 1
            CaracterValido KeyAscii, "N", tdb_Lista
        Case 2
            CaracterValido KeyAscii, "N", tdb_Lista
        Case 3
            CaracterValido KeyAscii, "A", tdb_Lista
        Case 4
            CaracterValido KeyAscii, "N", tdb_Lista
    End Select
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId.Text = .Columns("PKID").Value
            If mblnPrimeiraVez Then
                LeDaTabelaParaObj gstrFaceDeQuadra, Me
                MontaArrayHistorico
                gCorLinhaSelecionada tdb_Lista
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
                CarregaMelhoramentos txtPKId.Text
                lngLograAtual = tdb_Lista.Columns("intLogradouro").Value
                strQuadraAtual = tdb_Lista.Columns("Quadra").Value
                strSetorAtual = tdb_Lista.Columns("Setor").Value
                strSeqAtual = tdb_Lista.Columns("Sequência de Face").Value
                Do While Not tdb_HistoricoFaceDeQuadra.EOF
                   tdb_HistoricoFaceDeQuadra.Columns(4).Text = gstrConvVrDoSql(tdb_HistoricoFaceDeQuadra.Columns(2).Text * txt_dblSomaDosFatores)
                   tdb_HistoricoFaceDeQuadra.MoveNext
                Loop
                txtstrQuadra_LostFocus
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
        
    If UCase(strModoOperacao) = gstrSalvar Then
        If Not blnDadosOK Then Exit Sub
        If VerificaExer = True Then
           MsgBox "O Exercício informado já existe.", vbInformation + vbOKOnly
           tdb_HistoricoFaceDeQuadra.Refresh
           Exit Sub
        End If
        
        If ToolBarGeral(strModoOperacao, gstrFaceDeQuadra, mblnAlterando, tdb_Lista, _
                        Me, mobjAux, strQuery, , rptCadFaceDeQuadra, strQueryRelatorio, False) Then
           If mblnAlterando Then
               GravaHistorico Val(txtPKId)
           Else
               GravaHistorico lngPegaPkid
               GravaMelhoramentos lngPegaPkid
           End If
           mblnAlterando = False
           LimpaObjeto Me
           NovaFaceDeQuadra
           LimpaGrdHistorico
        End If
    ElseIf UCase(strModoOperacao) = UCase(gstrNovo) Then
        mblnAlterando = False
        LimpaObjeto Me
        NovaFaceDeQuadra
        LimpaGrdHistorico
        dbcintLogradouro.ListField = ""
    ElseIf UCase(strModoOperacao) = UCase(gstrDeletar) Then
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        DeletaHistorico Val(txtPKId)
        If ToolBarGeral(strModoOperacao, gstrFaceDeQuadra, mblnAlterando, tdb_Lista, _
                 Me, mobjAux, strQuery, , rptCadFaceDeQuadra, strQueryRelatorio) Then
            gobjBanco.ExecutaCommitTrans
        Else
            gobjBanco.ExecutaRollbackTrans
        End If
        mblnAlterando = False
        LimpaObjeto Me
        NovaFaceDeQuadra
        LimpaGrdHistorico
    Else
        ToolBarGeral strModoOperacao, gstrFaceDeQuadra, mblnAlterando, tdb_Lista, _
                 Me, mobjAux, strQuery, , rptCadFaceDeQuadra, strQueryRelatorio
    End If
    tdb_HistoricoFaceDeQuadra.Refresh
    
End Sub

Private Sub tdd_ValorMetroTerreno_DropDownClose()
    If tdd_ValorMetroTerreno.Row <> -1 Then
       tdb_HistoricoFaceDeQuadra.Columns("intValorMetroTerreno") = tdd_ValorMetroTerreno.Columns("Pkid")
       tdb_HistoricoFaceDeQuadra.Columns(1).Value = tdd_ValorMetroTerreno.Columns("intExercicio").Value
    End If
End Sub

Private Sub tdd_ValorMetroTerreno_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 3 Then Value = gstrConvVrDoSql(Value, 5)
End Sub

Private Sub txt_intAno_GotFocus()
    MarcaCampo txt_intAno
End Sub

Private Sub txt_intAno_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intAno
End Sub

Private Sub txtPKId_GotFocus()
    MarcaCampo txtPKId
End Sub

Private Sub txtPKId_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Function strQueryRelatorio() As String

    Dim strSQL          As String
   
    strSQL = ""
    
    strSQL = strSQL & "SELECT FQ.PKId, HFQ.intExercicio, FQ.strSetor, FQ.strQuadra, FQ.strSequenciaDeFace, "
    strSQL = strSQL & "LO.strDescricao as Logradouro, HFQ.intValorMetroTerreno, "
    strSQL = strSQL & gstrISNULL("LO.dtmdtExclusao", "'Não Cancelado'", "'Cancelado'") & " bytCancelado "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrFaceDeQuadra & " FQ, "
    strSQL = strSQL & gstrHistoricoFaceDeQuadra & " HFQ, "
    strSQL = strSQL & gstrLogradouro & " LO "
    strSQL = strSQL & "WHERE LO.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & " FQ.intLogradouro"
    strSQL = strSQL & " AND HFQ.intFaceDeQuadra " & strOUTJOracle & "=" & strOUTJSQLServer & " FQ.Pkid"
      
    Select Case bytOrdenacao
        Case Is = 1
            strSQL = strSQL & " ORDER BY strSetor" & IIf(blnOrdenacaoAsc, " ASC", " DESC") & ", strQuadra"
        Case Is = 2
            strSQL = strSQL & " ORDER BY strQuadra" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strSQL = strSQL & " ORDER BY Logradouro" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 4
            strSQL = strSQL & " ORDER BY strSequenciaDeFace" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 7
            strSQL = strSQL & " ORDER BY bytCancelado" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
   
    strQueryRelatorio = strSQL
   
End Function

Private Function blnDadosOK()
    
    blnDadosOK = False
    
    If Trim(txtstrSetor.Text) = "" Then
        ExibeMensagem "Preencha corretamente o campo setor!"
        txtstrSetor.SetFocus
        Exit Function
    ElseIf Trim(txtstrQuadra.Text) = "" Then
        ExibeMensagem "Preencha corretamente o campo quadra!"
        txtstrQuadra.SetFocus
        Exit Function
    ElseIf Not dbcintLogradouro.MatchedWithList Then
        ExibeMensagem "Preencha corretamente o campo logradouro!"
        dbcintLogradouro.SetFocus
        Exit Function
    ElseIf Trim(txtstrSequenciaDeFace.Text) = "" Then
        ExibeMensagem "Preencha corretamente o campo Sequência de Face!"
        txtstrSequenciaDeFace.SetFocus
        Exit Function
    End If
    
    If Not VerificaGridHistorico Then Exit Function
    
    If Not mblnAlterando Or (mblnAlterando And lngLograAtual <> Val(dbcintLogradouro.BoundText) Or _
                    UCase$(RTrim(strQuadraAtual)) <> UCase$(txtstrQuadra.Text) Or _
                    UCase$(RTrim(strSetorAtual)) <> UCase$(txtstrSetor.Text) Or UCase$(RTrim(strSeqAtual)) <> UCase$(txtstrSequenciaDeFace)) Then
        If lngPegaPkid <> 0 Then
             ExibeMensagem "Esta Sequência de Face já se encontra cadastrada."
             tdb_HistoricoFaceDeQuadra.ReBind
             tdb_HistoricoFaceDeQuadra.Refresh
             DoEvents
             Exit Function
        End If
    End If
    
    blnDadosOK = True
    
End Function

Private Sub txtstrQuadra_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrQuadra
End Sub

Private Sub txtstrQuadra_LostFocus()
'    txtstrQuadra = Format$(txtstrQuadra, "0000")
End Sub
   
Private Sub txtstrSequenciaDeFace_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSequenciaDeFace
End Sub

Private Sub txtstrSequenciaDeFace_LostFocus()
'    txtstrSequenciaDeFace = Format$(txtstrSequenciaDeFace, "00")
End Sub

Private Sub txtstrSetor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSetor
End Sub

Sub MontaColumnHeaders()
    
    With lvw_MelhoramentosCadastrados
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Descrição", 2000 '2690
        .ColumnHeaders.Add 2, , "Fator", 690
    End With
    
     With lvw_Melhoramentos
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Descrição", 1800 '2400
        .ColumnHeaders.Add 2, , "Ano", 600 '800
        .ColumnHeaders.Add 3, , "Fator", 800
    End With
End Sub
Sub AdicionarMelhoramento()
    If lvw_MelhoramentosCadastrados.ListItems.Count = 0 Then Exit Sub
    If lvw_MelhoramentosCadastrados.SelectedItem.Selected = False Then Exit Sub
    
    For giContador = 1 To lvw_Melhoramentos.ListItems.Count
        If lvw_Melhoramentos.ListItems(giContador).Tag = lvw_MelhoramentosCadastrados.SelectedItem.Tag Then
            ExibeMensagem "Melhoramento já relacionado com a seção."
            Exit Sub
        End If
    Next
    If Val(txt_intAno) = 0 Then
        ExibeMensagem "O ano do melhoramento tem que ser digitado."
        txt_intAno.SetFocus
        Exit Sub
    End If
    Set objList = lvw_Melhoramentos.ListItems.Add(, , lvw_MelhoramentosCadastrados.SelectedItem.Text)
    objList.SubItems(1) = txt_intAno
    objList.SubItems(2) = lvw_MelhoramentosCadastrados.SelectedItem.SubItems(1)
    objList.Tag = lvw_MelhoramentosCadastrados.SelectedItem.Tag
    txt_dblSomaDosFatores = gstrConvVrDoSql(dblSomaDosFatores)
End Sub

Sub RemoverMelhoramento()
    If lvw_Melhoramentos.ListItems.Count = 0 Then Exit Sub
    If lvw_Melhoramentos.SelectedItem.Selected = False Then Exit Sub
    
    lvw_Melhoramentos.ListItems.Remove lvw_Melhoramentos.SelectedItem.Index
    lvw_Melhoramentos.Sorted = True
    txt_dblSomaDosFatores = gstrConvVrDoSql(dblSomaDosFatores)
End Sub

Sub GravaMelhoramentos(lngPKId As Long)
    
    Dim strSQL  As String
    Dim intI    As Integer
    
    If lngPKId = 0 Then Exit Sub
    
    DeletaMelhoramentos lngPKId
    
    With lvw_Melhoramentos
        For intI = 1 To .ListItems.Count
            strSQL = ""
            strSQL = strSQL & "Insert Into " & gstrMelhoramentoDaSecaoDeLogradouro & " "
            strSQL = strSQL & "(intFaceDeQuadra, intMelhoramento, intAno, "
            strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr "
            strSQL = strSQL & ") Values ("
            strSQL = strSQL & lngPKId & ", "
            strSQL = strSQL & .ListItems(intI).Tag & ", "
            strSQL = strSQL & .ListItems(intI).SubItems(1) & ", "
            strSQL = strSQL & strGETDATE & ", "
            strSQL = strSQL & glngCodUsr
            strSQL = strSQL & ")"
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSQL
        Next
    End With
End Sub

Sub DeletaMelhoramentos(lngPKId As Long)
    Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "Delete From " & gstrMelhoramentoDaSecaoDeLogradouro & " "
    strSQL = strSQL & "Where intFaceDeQuadra = " & lngPKId
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSQL
End Sub

Sub NovaFaceDeQuadra()
    lvw_Melhoramentos.ListItems.Clear
    txt_intAno = ""
    tab_3dPasta.Tab = 0
    txt_dblSomaDosFatores.Text = ""
End Sub

Sub CarregaMelhoramentos(lngPKId As Long)
    Dim strSQL            As String
    Dim adoResultado      As ADODB.Recordset
    
    lvw_Melhoramentos.ListItems.Clear
    txt_intAno = ""
    
    strSQL = ""
    strSQL = strSQL & "Select M.PKId, M.strNomeDoMelhoramento Descricao, MS.intAno, M.dblFator "
    strSQL = strSQL & "From " & gstrMelhoramentoDaSecaoDeLogradouro & " MS,  "
    strSQL = strSQL & gstrMelhoramentoPublico & " M "
    strSQL = strSQL & "Where MS.intMelhoramento = M.PKId And MS.intFaceDeQuadra = " & lngPKId
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set objList = lvw_Melhoramentos.ListItems.Add(, , !Descricao)
                objList.SubItems(1) = !intAno
                If IsNull(!dblFator) = False Then
                    objList.SubItems(2) = gstrConvVrDoSql(!dblFator)
                End If
                objList.Tag = !Pkid
                .MoveNext
            Loop
            txt_dblSomaDosFatores = gstrConvVrDoSql(dblSomaDosFatores)
        End With
    End If
End Sub

Private Sub MontaArrayHistorico()
    Dim strSQL          As String
    Dim x               As XArrayDB
    Dim adoResultado    As ADODB.Recordset
    
    Set x = New XArrayDB
    
    strSQL = "SELECT HFQ.Pkid, HFQ.intExercicio, VT.dblVAlor,"
    strSQL = strSQL & " HFQ.intValorMetroTerreno, '' dblValorMetroPorEqui "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrHistoricoFaceDeQuadra & " HFQ, "
    strSQL = strSQL & gstrValorMetroTerreno & " VT"
    strSQL = strSQL & " WHERE HFQ.intFaceDeQuadra ='" & txtPKId & "' AND"
    strSQL = strSQL & " VT.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " HFQ.intValorMetroTerreno"
    strSQL = strSQL & " ORDER BY HFQ.intExercicio DESC"
    
    Set gobjBanco = New clsBanco
    If Not gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        Exit Sub
    End If
            
    If Not adoResultado.EOF Then
        x.ReDim 0, adoResultado.RecordCount - 1, 0, 4
        Dim varAux As Variant
        Do While Not adoResultado.EOF
            varAux = adoResultado!Pkid
            x(adoResultado.AbsolutePosition - 1, 0) = varAux
            
            varAux = adoResultado!intExercicio
            x(adoResultado.AbsolutePosition - 1, 1) = varAux
                    
            varAux = gstrConvVrDoSql(adoResultado!dblValor, 5)
            x(adoResultado.AbsolutePosition - 1, 2) = varAux
            
            varAux = adoResultado!intValorMetroTerreno
            x(adoResultado.AbsolutePosition - 1, 3) = varAux
            
            varAux = gstrConvVrDoSql(adoResultado!dblValorMetroPorEqui, 5)
            x(adoResultado.AbsolutePosition - 1, 4) = varAux
        
            adoResultado.MoveNext
       Loop
    
    Else
    
    x.ReDim 0, 0, 0, 3
    x(0, 0) = ""
    x(0, 1) = ""
    x(0, 2) = ""
    x(0, 3) = ""
    x(0, 4) = ""
    
    End If
            
    Set tdb_HistoricoFaceDeQuadra.Array = x
    tdb_HistoricoFaceDeQuadra.ReBind
    tdb_HistoricoFaceDeQuadra.Refresh

End Sub

Private Sub MontaArrayValorMetroTerreno()
    Dim strSQL          As String
    
    strSQL = "SELECT Pkid, Intexercicio, intCodigo, dblValor"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrValorMetroTerreno
    If Len(tdb_HistoricoFaceDeQuadra.Columns(1).Value) = 4 Then
        strSQL = strSQL & " Where intExercicio = " & tdb_HistoricoFaceDeQuadra.Columns(1).Value & " "
    End If
    strSQL = strSQL & " ORDER BY Intexercicio, intCodigo"
    
    LeDaTabelaParaObj "", tdd_ValorMetroTerreno, strSQL

End Sub


Private Function VerificaGridHistorico() As Boolean
    Dim A       As XArrayDB
    Dim intCont As Integer
    
    VerificaGridHistorico = False
    
    tdb_HistoricoFaceDeQuadra.Update
    tdb_HistoricoFaceDeQuadra.MoveFirst
        
    If tdb_HistoricoFaceDeQuadra.ApproxCount > 0 Then
        
        Set A = tdb_HistoricoFaceDeQuadra.Array
        If A.Count(1) = 0 Then
            ExibeMensagem "É necessário pelo menos um Exercicio e Valor."
            Exit Function
        Else
            For intCont = 0 To A.Count(1) - 1
                If Val(A.Value(intCont, 1)) = 0 Or Val(A.Value(intCont, 2)) = 0 Then
                    ExibeMensagem "Dados inválidos para o Histórico."
                    Exit Function
                End If
            Next
        End If
    End If
    
    VerificaGridHistorico = True
End Function

Private Sub GravaHistorico(lngPKId As Long)
    Dim A       As XArrayDB
    Dim intCont As Integer
    Dim strSQL  As String
    
            
    tdb_HistoricoFaceDeQuadra.Update
    tdb_HistoricoFaceDeQuadra.MoveFirst
        
    Set A = tdb_HistoricoFaceDeQuadra.Array
    
    strSQL = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    For intCont = 0 To A.Count(1) - 1
    
        If Val(A.Value(intCont, 0)) = 0 Then
            strSQL = strSQL & "INSERT INTO " & gstrHistoricoFaceDeQuadra
            strSQL = strSQL & " (intFaceDeQuadra, intExercicio, intValorMetroTerreno, dtmDtAtualizacao, lngCodUsr)"
            strSQL = strSQL & " VALUES("
            strSQL = strSQL & lngPKId & ", "
            strSQL = strSQL & A.Value(intCont, 1) & ", "
            strSQL = strSQL & A.Value(intCont, 3) & ", "
            strSQL = strSQL & strGETDATE & ", "
            strSQL = strSQL & glngCodUsr & ")"
            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), ";", "")
        Else
            strSQL = strSQL & "UPDATE " & gstrHistoricoFaceDeQuadra
            strSQL = strSQL & " SET intFaceDeQuadra = " & txtPKId & ", "
            strSQL = strSQL & " intExercicio = " & A.Value(intCont, 1) & ", "
            strSQL = strSQL & " intValorMetroTerreno = " & A.Value(intCont, 3) & ", "
            strSQL = strSQL & " dtmDtAtualizacao = " & strGETDATE & ", "
            strSQL = strSQL & " lngCodUsr = " & glngCodUsr
            strSQL = strSQL & " WHERE Pkid = " & A.Value(intCont, 0)
            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), ";", "")
        End If
    Next
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.Execute (strSQL)
End Sub

Private Sub LimpaGrdHistorico()
    Dim x As XArrayDB
    
    Set x = New XArrayDB
    
    x.ReDim 0, 0, 0, 3
    x(0, 0) = ""
    x(0, 1) = ""
    x(0, 2) = ""
    x(0, 3) = ""
    
    Set tdb_HistoricoFaceDeQuadra.Array = x
    tdb_HistoricoFaceDeQuadra.ReBind
    tdb_HistoricoFaceDeQuadra.Refresh
End Sub

Private Function lngPegaPkid() As Long
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = "SELECT Pkid"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrFaceDeQuadra
    strSQL = strSQL & " WHERE strSetor = '" & txtstrSetor.Text & "'"
    strSQL = strSQL & " AND strQuadra = '" & txtstrQuadra.Text & "'"
    strSQL = strSQL & " AND strSequenciaDeFace = '" & txtstrSequenciaDeFace & "'"
    strSQL = strSQL & " AND intLogradouro = '" & dbcintLogradouro.BoundText & "'"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngPegaPkid = adoResultado!Pkid
        End If
    End If

End Function

Private Sub DeletaHistorico(lngPKId As Long)
    Dim strSQL As String
    
    strSQL = "DELETE FROM " & gstrHistoricoFaceDeQuadra
    strSQL = strSQL & " WHERE intFaceDequadra = " & lngPKId
    
    gobjBanco.Execute strSQL

End Sub

Private Function VerificaExer() As Boolean
    Dim intFor, intAux, intExercicio As Integer
    Dim A       As XArrayDB
    
    Set A = tdb_HistoricoFaceDeQuadra.Array
    If A.Count(1) > 0 Then
       For intFor = 0 To A.Count(1) - 2
           intExercicio = Val(A.Value(intFor, 1))
           For intAux = intFor + 1 To A.Count(1) - 1
               If intExercicio = Val(A.Value(intAux, 1)) Then
                  VerificaExer = True
                  Exit Function
               End If
           Next
       Next
    End If
End Function

Private Sub txtstrSetor_LostFocus()
'    txtstrSetor = Format$(txtstrSetor, "00")
End Sub

Private Function strQueryLogradouro() As String
Dim strSQL  As String
     
    strSQL = ""
    
    strSQL = strSQL & "SELECT L.PKId, "
    strSQL = strSQL & " RTRIM(LTRIM(L.strDescricao)) " & strCONCAT & gstrISNULL("TL.strSigla", "''", "', '") & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & _
             strCONCAT & gstrISNULL("U.strDescricao", "' '", "', '") & strCONCAT & gstrISNULL("U.strDescricao", "''") & ")) " & strCONCAT & "' -> '" & strCONCAT & gstrISNULL("BA.strDescricao", "''") & " AS Logradouro "
    strSQL = strSQL & "FROM " & gstrLogradouro & " L, "
    strSQL = strSQL & gstrTituloLogradouro & " U, "
    strSQL = strSQL & gstrTipoLogradouro & " TL, "
    strSQL = strSQL & gstrBairro & " BA "
    strSQL = strSQL & " WHERE L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle
    'strSql = strSql & " AND L.Dtmdtexclusao is null "
    strSQL = strSQL & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
    strSQL = strSQL & " AND L.intBairro " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle
    strSQL = strSQL & " ORDER BY L.strDescricao "
    
    strQueryLogradouro = strSQL
        
End Function

Private Function dblSomaDosFatores() As Double
    
    For giContador = 1 To lvw_Melhoramentos.ListItems.Count
        dblSomaDosFatores = dblSomaDosFatores + gstrConvVrDoSql(lvw_Melhoramentos.ListItems(giContador).SubItems(2))
    Next
    
End Function


