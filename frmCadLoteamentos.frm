VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadLoteamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loteamentos"
   ClientHeight    =   7110
   ClientLeft      =   2205
   ClientTop       =   2175
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8820
   Begin TrueOleDBGrid70.TDBGrid tdb_lista 
      Height          =   2175
      Left            =   105
      TabIndex        =   43
      Top             =   4830
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   3836
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
      Columns(1).Caption=   "Código"
      Columns(1).DataField=   "intCodigo"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nome do Loteamento"
      Columns(2).DataField=   "strNome"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Bairro"
      Columns(3).DataField=   "strBairro"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   16
      Columns(4)._MaxComboItems=   5
      Columns(4).ValueItems(0)._DefaultItem=   0
      Columns(4).ValueItems(0).Value=   "A"
      Columns(4).ValueItems(0).Value.vt=   8
      Columns(4).ValueItems(0).DisplayValue=   "Aprovado"
      Columns(4).ValueItems(0).DisplayValue.vt=   8
      Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(1)._DefaultItem=   0
      Columns(4).ValueItems(1).Value=   "I"
      Columns(4).ValueItems(1).Value.vt=   8
      Columns(4).ValueItems(1).DisplayValue=   "Indeferido"
      Columns(4).ValueItems(1).DisplayValue.vt=   8
      Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems.Count=   2
      Columns(4).Caption=   "Situação"
      Columns(4).DataField=   "strSituAprovacao"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Dt Aprovação"
      Columns(5).DataField=   "dtmDtAprovacao"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).PartialRightColumn=   0   'False
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowSizing=   -1  'True
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowRowSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1217"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1138"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=4313"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4233"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=4577"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=4498"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=2170"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2090"
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
      EditDropDown    =   0   'False
      HeadLines       =   1
      FootLines       =   1
      TabAction       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
      MultiSelect     =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   1620,284
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=9,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.bgcolor=&H80000016&"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=25,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=44,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=27,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=30,.parent=6,.fgcolor=&H8000000E&"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=29,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=31,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=32,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=43,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=45,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=46,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=25"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=26"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=27"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=29"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=25,.alignment=1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=26"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=27"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=29"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=90,.parent=25"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=26"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=27"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=29"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=16,.parent=25"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=13,.parent=26"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=14,.parent=27"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=15,.parent=29"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=94,.parent=25"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=26"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=27"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=29"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=98,.parent=25"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=95,.parent=26"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=96,.parent=27"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=97,.parent=29"
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
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin VB.TextBox txtPKId 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4680
      Left            =   90
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   90
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   8255
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Loteamento"
      TabPicture(0)   =   "frmCadLoteamentos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Codigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_tipo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_CodigoPadrao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_Nome"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_MetrosTestada"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_AreaPadrao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_AreaLote"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_AreaRuas"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_AreaVerde"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_AreaInst"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl_SituApr"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl_DtApr"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl_DecrApr"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl_CREA"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl_RegCartorio"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl_NumMatricula"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl_DtMatricula"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl_NumProcesso"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl_AnoProcesso"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl_AnoDecrApr"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dbcintBairro"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtintCodigo"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtstrTipo"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtstrCodPadrao"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtstrNome"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtdblQtTestada"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtdblAreaPadrao"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtdblAreaLote"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtdblAreaRuas"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtdblAreaVerde"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtdblAreaInst"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cbo_strSituAprovacao"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtdtmDtAprovacao"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtintDecretoAprovacao"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtstrCREA"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtintNumCartorio"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtintNumMatricula"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtdtmDtMatricula"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtintNumProcesso"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtintAnoProcesso"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtintAnoAprDecreto"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).ControlCount=   42
      Begin VB.TextBox txtintAnoAprDecreto 
         Height          =   285
         Left            =   7410
         MaxLength       =   4
         TabIndex        =   14
         Top             =   2790
         Width           =   1065
      End
      Begin VB.TextBox txtintAnoProcesso 
         Height          =   285
         Left            =   7410
         MaxLength       =   4
         TabIndex        =   20
         Top             =   3780
         Width           =   1065
      End
      Begin VB.TextBox txtintNumProcesso 
         Height          =   285
         Left            =   1890
         TabIndex        =   19
         Top             =   3810
         Width           =   1485
      End
      Begin VB.TextBox txtdtmDtMatricula 
         Height          =   285
         Left            =   7410
         TabIndex        =   18
         Top             =   3450
         Width           =   1065
      End
      Begin VB.TextBox txtintNumMatricula 
         Height          =   285
         Left            =   1890
         TabIndex        =   17
         Top             =   3480
         Width           =   1485
      End
      Begin VB.TextBox txtintNumCartorio 
         Height          =   285
         Left            =   7410
         TabIndex        =   16
         Top             =   3120
         Width           =   1065
      End
      Begin VB.TextBox txtstrCREA 
         Height          =   285
         Left            =   1890
         TabIndex        =   15
         Top             =   3150
         Width           =   1485
      End
      Begin VB.TextBox txtintDecretoAprovacao 
         Height          =   285
         Left            =   1890
         TabIndex        =   13
         Top             =   2820
         Width           =   1485
      End
      Begin VB.TextBox txtdtmDtAprovacao 
         Height          =   285
         Left            =   7410
         TabIndex        =   12
         Top             =   2460
         Width           =   1065
      End
      Begin VB.ComboBox cbo_strSituAprovacao 
         Height          =   315
         ItemData        =   "frmCadLoteamentos.frx":001C
         Left            =   1890
         List            =   "frmCadLoteamentos.frx":0026
         TabIndex        =   11
         Top             =   2460
         Width           =   1485
      End
      Begin VB.TextBox txtdblAreaInst 
         Height          =   285
         Left            =   7410
         TabIndex        =   10
         Top             =   2130
         Width           =   1065
      End
      Begin VB.TextBox txtdblAreaVerde 
         Height          =   285
         Left            =   1890
         TabIndex        =   9
         Top             =   2130
         Width           =   1485
      End
      Begin VB.TextBox txtdblAreaRuas 
         Height          =   285
         Left            =   7410
         TabIndex        =   8
         Top             =   1800
         Width           =   1065
      End
      Begin VB.TextBox txtdblAreaLote 
         Height          =   285
         Left            =   1890
         TabIndex        =   7
         Top             =   1800
         Width           =   1485
      End
      Begin VB.TextBox txtdblAreaPadrao 
         Height          =   285
         Left            =   7410
         TabIndex        =   6
         Top             =   1470
         Width           =   1065
      End
      Begin VB.TextBox txtdblQtTestada 
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
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1470
         Width           =   1485
      End
      Begin VB.TextBox txtstrNome 
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
         Left            =   1890
         MaxLength       =   35
         TabIndex        =   3
         Top             =   810
         Width           =   6585
      End
      Begin VB.TextBox txtstrCodPadrao 
         Height          =   285
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   1
         Top             =   495
         Width           =   465
      End
      Begin VB.TextBox txtstrTipo 
         Height          =   285
         Left            =   7875
         MaxLength       =   4
         TabIndex        =   2
         Top             =   495
         Width           =   585
      End
      Begin VB.TextBox txtintCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1905
         TabIndex        =   0
         Top             =   480
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo dbcintBairro 
         Height          =   315
         Left            =   1890
         TabIndex        =   4
         Top             =   1125
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   1350
         TabIndex        =   44
         Top             =   1215
         Width           =   405
      End
      Begin VB.Label lbl_AnoDecrApr 
         AutoSize        =   -1  'True
         Caption         =   "Ano Aprovação Decr"
         Height          =   195
         Left            =   5820
         TabIndex        =   42
         Top             =   2850
         Width           =   1500
      End
      Begin VB.Label lbl_AnoProcesso 
         AutoSize        =   -1  'True
         Caption         =   "Ano do Processo"
         Height          =   195
         Left            =   6120
         TabIndex        =   41
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lbl_NumProcesso 
         AutoSize        =   -1  'True
         Caption         =   "Número do Processo"
         Height          =   195
         Left            =   300
         TabIndex        =   40
         Top             =   3870
         Width           =   1485
      End
      Begin VB.Label lbl_DtMatricula 
         AutoSize        =   -1  'True
         Caption         =   "Data Matrícula"
         Height          =   195
         Left            =   6270
         TabIndex        =   39
         Top             =   3510
         Width           =   1065
      End
      Begin VB.Label lbl_NumMatricula 
         AutoSize        =   -1  'True
         Caption         =   "Número Matrícula"
         Height          =   195
         Left            =   510
         TabIndex        =   38
         Top             =   3540
         Width           =   1275
      End
      Begin VB.Label lbl_RegCartorio 
         AutoSize        =   -1  'True
         Caption         =   "Registro no Cartório"
         Height          =   195
         Left            =   5940
         TabIndex        =   37
         Top             =   3180
         Width           =   1395
      End
      Begin VB.Label lbl_CREA 
         AutoSize        =   -1  'True
         Caption         =   "Registro no CREA"
         Height          =   195
         Left            =   510
         TabIndex        =   36
         Top             =   3210
         Width           =   1290
      End
      Begin VB.Label lbl_DecrApr 
         AutoSize        =   -1  'True
         Caption         =   "Decreto de Aprovação"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   2880
         Width           =   1620
      End
      Begin VB.Label lbl_DtApr 
         AutoSize        =   -1  'True
         Caption         =   "Data Aprovação"
         Height          =   195
         Left            =   6165
         TabIndex        =   34
         Top             =   2520
         Width           =   1170
      End
      Begin VB.Label lbl_SituApr 
         AutoSize        =   -1  'True
         Caption         =   "Situação da Aprovação"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2550
         Width           =   1680
      End
      Begin VB.Label lbl_AreaInst 
         AutoSize        =   -1  'True
         Caption         =   "Área Institucional"
         Height          =   195
         Left            =   6120
         TabIndex        =   32
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label lbl_AreaVerde 
         AutoSize        =   -1  'True
         Caption         =   "Área Verde"
         Height          =   195
         Left            =   990
         TabIndex        =   31
         Top             =   2190
         Width           =   795
      End
      Begin VB.Label lbl_AreaRuas 
         AutoSize        =   -1  'True
         Caption         =   "Área Ruas"
         Height          =   195
         Left            =   6585
         TabIndex        =   30
         Top             =   1860
         Width           =   750
      End
      Begin VB.Label lbl_AreaLote 
         AutoSize        =   -1  'True
         Caption         =   "Área Lote"
         Height          =   195
         Left            =   1095
         TabIndex        =   29
         Top             =   1860
         Width           =   690
      End
      Begin VB.Label lbl_AreaPadrao 
         AutoSize        =   -1  'True
         Caption         =   "Área Padrão"
         Height          =   195
         Left            =   6450
         TabIndex        =   28
         Top             =   1530
         Width           =   885
      End
      Begin VB.Label lbl_MetrosTestada 
         AutoSize        =   -1  'True
         Caption         =   "Qt Metros de Testada"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label lbl_Nome 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   1350
         TabIndex        =   26
         Top             =   900
         Width           =   420
      End
      Begin VB.Label lbl_CodigoPadrao 
         AutoSize        =   -1  'True
         Caption         =   "Região Administrativa"
         Height          =   195
         Left            =   4245
         TabIndex        =   25
         Top             =   540
         Width           =   1530
      End
      Begin VB.Label lbl_tipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   7455
         TabIndex        =   24
         Top             =   540
         Width           =   315
      End
      Begin VB.Label lbl_Codigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1305
         TabIndex        =   22
         Top             =   540
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCadLoteamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando     As Boolean
Dim mobjAux           As Object
Dim mblnPrimeiraVez   As Boolean
Dim strCodigoAtual     As String
Dim strDescricaoAtual  As String


Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case UCase(strModoOperacao)
    Case "NOVO"
        LimpaObjeto Me, mblnAlterando
        LimpaCombo
        mblnPrimeiraVez = False
        gstrProximoCodigo txtintCodigo, gstrLoteamento, "intCodigo", gintCodSeguranca
    
    Case "SALVAR"
        If blnDadosOk Then
            If ToolBarGeral(strModoOperacao, gstrLoteamento, mblnAlterando, tdb_lista, Me, mobjAux, strQuery, , , , False) Then
               'If mblnPrimeiraVez = True Then
               'If mblnAlterando = True Then
                  'If IsNull(tdb_lista.Columns(0).Value) Then
                     'txtPKId = gstrENulo(tdb_lista.Columns(0).Value)
                  'Else
                     'txtPKId = tdb_lista.Columns(0).Value
                     GravaSituacaoAprovacao
                  'End If
               'End If
               tdb_lista.Rebind
               tdb_lista.Refresh
               LimpaObjeto Me, mblnAlterando
               LimpaCombo
               mblnPrimeiraVez = False
               LeDaTabelaParaObj gstrLoteamento, tdb_lista, strQuery
            End If
        End If
    
    Case "DELETAR"
            If ExcluiRegistro Then
                LimpaCombo
                LimpaObjeto Me
                tdb_lista.Rebind
                tdb_lista.Refresh
                mblnAlterando = False
                mblnPrimeiraVez = False
                LeDaTabelaParaObj gstrLoteamento, tdb_lista, strQuery
            End If
    
    Case "APLICAR"
        ToolBarGeral strModoOperacao, gstrLoteamento, mblnAlterando, tdb_lista, Me, mobjAux, strQuery
    
    Case Else
        ToolBarGeral strModoOperacao, gstrLoteamento, mblnAlterando, tdb_lista, Me, mobjAux, strQuery, , rptLoteamento, strQueryRelatorio
        
    End Select
    
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
        
    
End Sub


Private Sub Form_Activate()
    
    gintCodSeguranca = 587
    VirificaGradeListView Me
    
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
End Sub

Private Sub Form_Load()
   mblnAlterando = False
   VerificaObjParaAplicar mobjAux
   dbcintBairro.Tag = "SELECT PKId, strDescricao FROM " & gstrBairro & " ORDER BY strDescricao;strDescricao"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub


Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_lista
End Sub
Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_lista, ColIndex
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
With tdb_lista
    If (Not .EOF And Not .BOF) Then
        If mblnPrimeiraVez Then
            mblnAlterando = True
            txtPKId.text = .Columns("PKID").Value
            LeDaTabelaParaObj gstrLoteamento, Me
            gCorLinhaSelecionada tdb_lista
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            strCodigoAtual = tdb_lista.Columns("intcodigo").Value
            strDescricaoAtual = tdb_lista.Columns("strNome").Value
            SelecionaSituAprovacao
        End If
    End If
End With
End Sub

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    
    If Trim(txtintCodigo.text) = "" Then
        ExibeMensagem "O Código deve ser informado."
        txtintCodigo.SetFocus
        Exit Function
    ElseIf Trim(txtstrNome.text) = "" Then
        ExibeMensagem "O Nome deve ser informado."
        txtstrNome.SetFocus
        Exit Function
    ElseIf Trim(dbcintBairro.text) = Empty Then
        ExibeMensagem "O Bairro deve ser informado."
        dbcintBairro.SetFocus
        Exit Function
    ElseIf Trim(cbo_strSituAprovacao.text) <> "" Then
        If cbo_strSituAprovacao.ListIndex = -1 Then
            ExibeMensagem "Selecione um Tipo de Situação da Aprovação válido."
            cbo_strSituAprovacao.SetFocus
            Exit Function
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtintCodigo.text)) Then
        If gblnExisteCodigo(1, gstrLoteamento, "intCodigo", txtintCodigo.text) Then
            ExibeMensagem "Já existe registro com esse código."
            Exit Function
        End If
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrNome.text) <> UCase$(strDescricaoAtual)) Then
        If gblnExisteCodigo(1, gstrLoteamento, "strNome", "'" & txtstrNome & "'") Then
            ExibeMensagem "Já existe registro com essa descrição."
            Exit Function
        End If
    End If

    
    blnDadosOk = True
    
End Function

Private Sub txtdblAreaInst_GotFocus()
    MarcaCampo txtdblAreaInst
End Sub

Private Sub txtdblAreaInst_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblAreaInst
End Sub

Private Sub txtdblAreaLote_GotFocus()
    MarcaCampo txtdblAreaLote
End Sub

Private Sub txtdblAreaLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblAreaLote
End Sub

Private Sub txtdblAreaPadrao_GotFocus()
    MarcaCampo txtdblAreaPadrao
End Sub

Private Sub txtdblAreaPadrao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblAreaPadrao
End Sub

Private Sub txtdblAreaRuas_GotFocus()
    MarcaCampo txtdblAreaRuas
End Sub

Private Sub txtdblAreaRuas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblAreaRuas
End Sub

Private Sub txtdblAreaVerde_GotFocus()
    MarcaCampo txtdblAreaVerde
End Sub

Private Sub txtdblAreaVerde_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblAreaVerde
End Sub

Private Sub txtdblQtTestada_GotFocus()
    MarcaCampo txtdblQtTestada
End Sub

Private Sub txtdblQtTestada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblQtTestada
End Sub

Private Sub txtdtmDtAprovacao_GotFocus()
    MarcaCampo txtdtmDtAprovacao
End Sub

Private Sub txtdtmDtAprovacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtAprovacao
End Sub

Private Sub txtdtmDtAprovacao_LostFocus()
    txtdtmDtAprovacao = gstrDataFormatada(txtdtmDtAprovacao, False)
End Sub

Private Sub txtdtmDtMatricula_GotFocus()
    MarcaCampo txtdtmDtMatricula
End Sub

Private Sub txtdtmDtMatricula_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtMatricula
End Sub

Private Sub txtdtmDtMatricula_LostFocus()
    txtdtmDtMatricula = gstrDataFormatada(txtdtmDtMatricula, False)
End Sub

Private Sub txtintAnoAprDecreto_GotFocus()
    MarcaCampo txtintAnoAprDecreto
End Sub

Private Sub txtintAnoAprDecreto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintAnoAprDecreto
End Sub

Private Sub txtintAnoAprDecreto_LostFocus()
    txtintAnoAprDecreto = Right(gstrDataFormatada("01/01/" & txtintAnoAprDecreto, False), 4)
End Sub

Private Sub txtintAnoProcesso_GotFocus()
    MarcaCampo txtintAnoProcesso
End Sub

Private Sub txtintAnoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintAnoProcesso
End Sub

Private Sub txtintAnoProcesso_LostFocus()
    txtintAnoProcesso = Right(gstrDataFormatada("01/01/" & txtintAnoProcesso, False), 4)
End Sub

Private Sub txtintCodigo_GotFocus()
    MarcaCampo txtintCodigo
    gstrProximoCodigo txtintCodigo, gstrLoteamento, "intCodigo", gintCodSeguranca
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub txtintDecretoAprovacao_GotFocus()
    MarcaCampo txtintDecretoAprovacao
End Sub

Private Sub txtintDecretoAprovacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintDecretoAprovacao
End Sub

Private Sub txtintNumCartorio_GotFocus()
    MarcaCampo txtintNumCartorio
End Sub

Private Sub txtintNumCartorio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumCartorio
End Sub

Private Sub txtintNumMatricula_GotFocus()
    MarcaCampo txtintNumMatricula
End Sub

Private Sub txtintNumMatricula_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumMatricula
End Sub

Private Sub txtintNumProcesso_GotFocus()
    MarcaCampo txtintNumProcesso
End Sub

Private Sub txtintNumProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumProcesso
End Sub

Private Sub txtstrTipo_GotFocus()
    MarcaCampo txtstrTipo
End Sub

Private Sub txtstrTipo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrTipo
End Sub

Private Sub txtstrCodPadrao_GotFocus()
    MarcaCampo txtstrCodPadrao
End Sub

Private Sub txtstrCodPadrao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrCodPadrao
End Sub

Private Sub txtstrCREA_GotFocus()
    MarcaCampo txtstrCREA
End Sub

Private Sub txtstrNome_GotFocus()
    MarcaCampo txtstrNome
End Sub

Private Sub txtstrNome_LostFocus()
    txtstrNome = Trim(txtstrNome)
End Sub

Private Sub GravaSituacaoAprovacao()
    Dim strSql As String
    Dim adoRec As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "UPDATE " & gstrLoteamento & " "
    strSql = strSql & "SET strSituAprovacao = "
    
    If cbo_strSituAprovacao.ListIndex = 0 Then
        strSql = strSql & "'A' "
    ElseIf cbo_strSituAprovacao.ListIndex = 1 Then
        strSql = strSql & "'I' "
    Else
        strSql = strSql & "NULL "
    End If
    
    strSql = strSql & "WHERE intCodigo = " & Trim(txtintCodigo.text)

    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql, False

End Sub

Private Sub SelecionaSituAprovacao()
    If tdb_lista.Columns("strSituAprovacao").Value = "A" Then
        cbo_strSituAprovacao.ListIndex = 0
    ElseIf tdb_lista.Columns("strSituAprovacao").Value = "I" Then
        cbo_strSituAprovacao.ListIndex = 1
    Else
        cbo_strSituAprovacao.ListIndex = -1
    End If
End Sub

Private Sub LimpaCombo()
    cbo_strSituAprovacao.ListIndex = -1
End Sub

Private Function ExcluiRegistro() As Boolean
    Dim strSql As String
    Dim adoRec As ADODB.Recordset
    Dim strDescricao As String
    
    strDescricao = txtstrNome
    
    If gblnExclusaoGravacaoOk("E", strDescricao) Then
        
        strSql = ""
        strSql = strSql & " DELETE FROM " & gstrLoteamento
        strSql = strSql & " WHERE PKId = " & txtPKId
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 30, adoRec) Then
            ExcluiRegistro = True
            Exit Function
        End If
    
    End If
    
    ExcluiRegistro = False
        
End Function

Private Function strQuery() As String

Dim strSql As String

    strSql = ""
    strSql = strSql & "SELECT LO.PKId pkID, LO.intCodigo intCodigo, LO.strNome strNome, "
    strSql = strSql & "LO.strSituAprovacao strSituAprovacao, "
    strSql = strSql & "LO.dtmDtAprovacao dtmdtAprovacao, BA.strDescricao strBairro "
    strSql = strSql & "FROM " & gstrLoteamento & " LO, "
    strSql = strSql & gstrBairro & " BA "
    strSql = strSql & "WHERE BA.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LO.intBairro "
    strSql = strSql & " ORDER BY strNome "

    strQuery = strSql
    
End Function

Private Function strQueryRelatorio() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, intCodigo, strNome, strTipo"
    strSql = strSql & " FROM " & gstrLoteamento
    strQueryRelatorio = strSql
End Function
