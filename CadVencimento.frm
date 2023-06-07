VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadVencimentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Vencimento de Parcelas"
   ClientHeight    =   5175
   ClientLeft      =   3075
   ClientTop       =   2625
   ClientWidth     =   6450
   HelpContextID   =   48
   Icon            =   "CadVencimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6450
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Vencimento"
      TabPicture(0)   =   "CadVencimento.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrNomeDoVencimento"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintTributo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtstrNomeDoVencimento"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tdb_Vencimento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintTributo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Parcelas"
      TabPicture(1)   =   "CadVencimento.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_PKIdParcela"
      Tab(1).Control(1)=   "txt_intNumero"
      Tab(1).Control(2)=   "txt_dtmDataDaParcela"
      Tab(1).Control(3)=   "txt_intExercicio"
      Tab(1).Control(4)=   "tdb_Parcelas"
      Tab(1).Control(5)=   "lbl_intNumero"
      Tab(1).Control(6)=   "lbl_dtmDataDaParcela"
      Tab(1).Control(7)=   "lbl_intExercicio"
      Tab(1).ControlCount=   8
      Begin VB.TextBox txt_PKIdParcela 
         Height          =   285
         Left            =   -69525
         TabIndex        =   14
         Top             =   1095
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txt_intNumero 
         Height          =   285
         Left            =   -72870
         MaxLength       =   4
         OLEDragMode     =   1  'Automatic
         TabIndex        =   3
         Top             =   645
         Width           =   540
      End
      Begin VB.TextBox txt_dtmDataDaParcela 
         Height          =   300
         Left            =   -70635
         MaxLength       =   10
         TabIndex        =   4
         Top             =   615
         Width           =   1080
      End
      Begin VB.TextBox txt_intExercicio 
         Height          =   285
         Left            =   -72870
         MaxLength       =   4
         OLEDragMode     =   1  'Automatic
         TabIndex        =   5
         Top             =   1005
         Width           =   765
      End
      Begin MSDataListLib.DataCombo dbcintTributo 
         Height          =   315
         Left            =   1005
         TabIndex        =   1
         Top             =   915
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Vencimento 
         Height          =   3435
         Left            =   135
         TabIndex        =   2
         Top             =   1305
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   6059
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
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "strNomeDoVencimento"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=7250"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=7170"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
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
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtstrNomeDoVencimento 
         Height          =   285
         Left            =   1005
         MaxLength       =   50
         TabIndex        =   0
         Top             =   510
         Width           =   5025
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Parcelas 
         Height          =   3060
         Left            =   -73920
         TabIndex        =   6
         Top             =   1455
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   5398
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
         Columns(1).Caption=   "Nº da parcela"
         Columns(1).DataField=   "intNumero"
         Columns(1).DropDown=   "tdd_Materiais"
         Columns(1).DropDown.vt=   8
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data de Vencimento"
         Columns(2).DataField=   "dtmDataDaParcela"
         Columns(2).DataWidth=   12
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Exercício"
         Columns(3).DataField=   "intExercicio"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(7)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=2487"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2408"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=256"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=3387"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=3307"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=258"
         Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(21)=   "Column(3).Width=1693"
         Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=1614"
         Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=258"
         Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(14)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42,.alignment=3"
         _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14,.alignment=0"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(44)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14,.alignment=0"
         _StyleDefs(45)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Named:id=33:Normal"
         _StyleDefs(48)  =   ":id=33,.parent=0"
         _StyleDefs(49)  =   "Named:id=34:Heading"
         _StyleDefs(50)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(51)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
         _StyleDefs(52)  =   "Named:id=35:Footing"
         _StyleDefs(53)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   "Named:id=36:Selected"
         _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(56)  =   "Named:id=37:Caption"
         _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_intNumero 
         AutoSize        =   -1  'True
         Caption         =   "Nº da parcela"
         Height          =   195
         Left            =   -73920
         TabIndex        =   13
         Top             =   690
         Width           =   975
      End
      Begin VB.Label lbl_dtmDataDaParcela 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vencimento"
         Height          =   195
         Left            =   -72180
         TabIndex        =   12
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label lbl_intExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   -73620
         TabIndex        =   11
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label lblintTributo 
         AutoSize        =   -1  'True
         Caption         =   "Tributo"
         Height          =   195
         Left            =   435
         TabIndex        =   10
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label lblstrNomeDoVencimento 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   600
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadVencimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnAlterando                    As Boolean
Dim blnAlterandoParcela             As Boolean
Dim blnPrimeiraVezParcela           As Boolean
Dim mobjAux                         As Object
Dim oList                           As Object
Dim bytOrdenacao                    As Byte
Dim mblnSelecionou                  As Boolean
Dim blnPrimeiraVez                  As Boolean
Dim blnOrdenacaoAsc                 As Boolean
Dim XParcelas                       As New XArrayDB

Private Sub dbcintTributo_Click(Area As Integer)
    DropDownDataCombo dbcintTributo, Me, Area
End Sub

Private Sub dbcintTributo_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintTributo_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTributo, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTributo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTributo
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 600
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
    tab_3dPasta.TabEnabled(1) = False
    dbcintTributo.Tag = strQueryDataComboComposicaoDaReceita & ";strDescricao"
    bytOrdenacao = 1: blnOrdenacaoAsc = True
    'VerificaListaAutomatica gstrVencimentos, tdb_Vencimento, strQuery
    VerificaParametroCombox mobjAux
End Sub

Private Function strQueryDataComboComposicaoDaReceita()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrComposicaoDaReceita & " "
    strSql = strSql & "ORDER BY strDescricao"
    strQueryDataComboComposicaoDaReceita = strSql
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    blnPrimeiraVez = False
End Sub

Private Sub tdb_Parcelas_Click()
    blnPrimeiraVezParcela = True
End Sub

Private Sub tdb_Parcelas_FilterChange()
    blnPrimeiraVezParcela = False
    gblnFilraCampos tdb_Parcelas
End Sub

Private Sub tdb_Parcelas_KeyPress(KeyAscii As Integer)
    Select Case tdb_Parcelas.Col
        Case 1, 3
            CaracterValido KeyAscii, "N", tdb_Parcelas
        Case 2
            CaracterValido KeyAscii, "D", tdb_Parcelas
        Case Else
            CaracterValido KeyAscii, "A", tdb_Parcelas
    End Select
End Sub

Private Sub tdb_Parcelas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If blnPrimeiraVezParcela Then
        With tdb_Parcelas
            If Not .EOF And Not .BOF Then
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                blnAlterandoParcela = True
                TrocaCorObjeto txt_intNumero, True
                TrocaCorObjeto txt_intExercicio, True
                txt_PKIdParcela.Text = .Columns("PKID").Value
                txt_intNumero.Text = .Columns("intNumero").Value
                txt_dtmDataDaParcela.Text = .Columns("dtmDataDaParcela").Value
                txt_intExercicio.Text = .Columns("intExercicio").Value
            End If
        End With
    End If
End Sub

Private Sub tdb_Vencimento_Click()
    blnPrimeiraVez = True
End Sub

Private Sub tdb_Vencimento_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_Vencimento
End Sub

Private Sub tdb_Vencimento_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub tdb_Vencimento_HeadClick(ByVal ColIndex As Integer)

    blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
   
    bytOrdenacao = ColIndex: MantemForm gstrRefresh
   
End Sub

Private Sub tdb_Vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CaracterValido KeyAscii, "N", tdb_Vencimento
    Else
        Select Case tdb_Vencimento.Col
            Case 0 'PKId
                CaracterValido KeyAscii, "N", tdb_Vencimento
        End Select
    End If
End Sub

Private Function strQueryParcela() As String
    Dim strSql  As String
    
    strSql = "Select * "
    strSql = strSql & "From " & gstrVencimentosDasParcelas & " "
    strSql = strSql & "Where intVencimento = " & txtPKId.Text & " "
    strSql = strSql & "Order By intExercicio, intNumero"
    strQueryParcela = strSql
    
End Function

Private Sub tdb_Vencimento_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strSql As String
    
    With tdb_Vencimento
        If Not .EOF And Not .BOF Then
            If blnPrimeiraVez Then
                tab_3dPasta.TabEnabled(1) = True
                blnAlterando = True
                txtPKId = .Columns("PKId").Value
                LeDaTabelaParaObj gstrVencimentos, Me
                gCorLinhaSelecionada tdb_Vencimento
                If txtPKId.Text <> "" Then
                    
                    LeDaTabelaParaObj gstrVencimentosDasParcelas, tdb_Parcelas, strQueryParcela
                End If
'=============
                gCorLinhaSelecionada tdb_Vencimento
                
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else

                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
'=============
                mblnSelecionou = True
                blnAlterando = True
            End If
        End If
    End With
End Sub

Private Sub txt_dtmDataDaParcela_GotFocus()
    MarcaCampo txt_dtmDataDaParcela
    tab_3dPasta.Tab = 1
End Sub

Private Sub txt_dtmDataDaParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataDaParcela
End Sub

Private Sub txt_dtmDataDaParcela_LostFocus()
    txt_dtmDataDaParcela.Text = gstrDataFormatada(txt_dtmDataDaParcela.Text)
    txt_intExercicio = Right(txt_dtmDataDaParcela.Text, 4)
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
    tab_3dPasta.Tab = 1
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_intNumero_GotFocus()
    MarcaCampo txt_intNumero
    tab_3dPasta.Tab = 1
End Sub

Private Sub txt_intNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intNumero
End Sub

Private Sub txtstrNomeDoVencimento_GotFocus()
    MarcaCampo txtstrNomeDoVencimento
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrNomeDoVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNomeDoVencimento
End Sub

Function strQuery() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNomeDoVencimento "
    strSql = strSql & "FROM " & gstrVencimentos & " "
    
    'ORDENA NO CABEÇALHO DO GRID
    Select Case bytOrdenacao
   
      Case Is = 0
            strSql = strSql & " ORDER BY PKId" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
         
      Case Is = 1
         strSql = strSql & " ORDER BY strNomeDoVencimento" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    strQuery = strSql
    
End Function

Private Function GravaParcela() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    If ParcelaOK Then
        If gblnExclusaoGravacaoOk(IIf(blnAlterandoParcela, "A", "I"), "de parcela") Then
            If blnAlterandoParcela Then
                strSql = "UPDATE " & gstrVencimentosDasParcelas & _
                " SET intExercicio = " & txt_intExercicio.Text & _
                ", intNumero = " & txt_intNumero.Text & _
                ", dtmDataDaParcela = " & gstrConvDtParaSql(txt_dtmDataDaParcela.Text)
'                ", dtmDtAtualizacao = GETDATE() "
                strSql = strSql & ", dtmDtAtualizacao = " & strGETDATE & _
                ", lngCodUsr = " & glngCodUsr & _
                " WHERE PKID = " & txt_PKIdParcela.Text
            Else
                strSql = "INSERT INTO " & gstrVencimentosDasParcelas & _
                " (intVencimento, intExercicio, intNumero, dtmDataDaParcela, " & _
                "dtmDtAtualizacao, lngCodUsr) " & _
                "VALUES (" & _
                txtPKId.Text & _
                ", " & txt_intExercicio.Text & _
                ", " & txt_intNumero.Text & _
                ", " & gstrConvDtParaSql(txt_dtmDataDaParcela.Text)
'                ", GETDATE()"
                strSql = strSql & ", " & strGETDATE & _
                ", " & glngCodUsr & ")"
            End If
            
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            If gobjBanco.Execute(strSql) Then
                gobjBanco.ExecutaCommitTrans
                LeDaTabelaParaObj gstrVencimentosDasParcelas, tdb_Parcelas, strQueryParcela
                
                Set adoResultado = tdb_Parcelas.DataSource
                If blnAlterandoParcela Then
                    adoResultado.Find ("PKId = '" & txt_PKIdParcela.Text & "'")
                Else
                    adoResultado.Find ("PKId = '" & glngPegaUltimaChave(gstrVencimentosDasParcelas, "PKID") & "'")
                End If
                Set adoResultado = Nothing
                GravaParcela = True
            Else
                gobjBanco.ExecutaRollbackTrans
            End If
        End If
    End If
    
End Function

Private Function Ultima() As Boolean
    Dim blnUltima As Boolean
    Dim adoResultado As ADODB.Recordset
    Dim intPosicao As Integer
    
    Set adoResultado = tdb_Parcelas.DataSource
    blnUltima = True
    With adoResultado
        intPosicao = .AbsolutePosition
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If Val(txt_intExercicio.Text) = Val(!intExercicio) And (Val(txt_intNumero.Text) + 1) = Val(!intNumero) Then
                    blnUltima = False
                    .Move intPosicao - 1, 1
                    ExibeMensagem "Exclua as últimas parcelas do Exercício " & txt_intExercicio.Text & " para manter a sequência das parcelas."
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
        
    End With
    Set adoResultado = Nothing
    Ultima = blnUltima

End Function

Private Function DeletaParcela() As Boolean
    Dim strSql As String
    
    If blnAlterandoParcela Then
        If Ultima Then
            If gblnExclusaoGravacaoOk("E", "da parcela") Then
                strSql = "DELETE FROM " & gstrVencimentosDasParcelas & _
                " WHERE PKId = " & txt_PKIdParcela.Text
                
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaBeginTrans
                If gobjBanco.Execute(strSql) Then
                    gobjBanco.ExecutaCommitTrans
                    LeDaTabelaParaObj gstrVencimentosDasParcelas, tdb_Parcelas, strQueryParcela
                    DeletaParcela = True
                Else
                    gobjBanco.ExecutaRollbackTrans
                End If
            End If
        End If
    End If
    
End Function

Private Function blnDeletaVencimento() As Boolean

'******************************************************************************************
' Data: 06/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
    
    If Trim(txtPKId) = "" Then
        Exit Function
    End If
    If MsgBox("Confirma a exclusão deste vencimento e todas as suas parcelas?", vbQuestion + vbYesNo) = vbYes Then
        strSql = ""
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
        
        strSql = strSql & "DELETE FROM " & gstrVencimentosDasParcelas & " "
        strSql = strSql & "WHERE intVencimento = " & txtPKId.Text
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        strSql = strSql & " DELETE FROM " & gstrVencimentos & " "
        strSql = strSql & "WHERE PKId = " & txtPKId.Text
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        If gobjBanco.Execute(strSql) Then
            gobjBanco.ExecutaCommitTrans
            blnDeletaVencimento = True
        Else
            gobjBanco.ExecutaRollbackTrans
        End If
                
    End If
End Function

Private Sub NovoVencimento()
    blnPrimeiraVez = False
    blnAlterando = False
    NovaParcela
    Set tdb_Parcelas.DataSource = Nothing
    tab_3dPasta.TabEnabled(1) = False
End Sub

Private Sub NovaParcela()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    blnPrimeiraVezParcela = False
    blnAlterandoParcela = False
    txt_intNumero.Text = ""
    txt_dtmDataDaParcela.Text = ""
    txt_intExercicio.Text = ""
    TrocaCorObjeto txt_intNumero, False
    TrocaCorObjeto txt_intExercicio, False
End Sub

Function ParcelaOK() As Boolean
    Dim adoResultado As ADODB.Recordset
    Dim blnTemSequencia As Boolean
    Dim blnInicioSequencia As Boolean
    Dim intPosicao As Integer
    
    If Trim(txt_intNumero.Text) = "" Then
        ExibeMensagem "Digite o número da parcela."
        txt_intNumero.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_dtmDataDaParcela.Text) Then
        ExibeMensagem "Data de vencimento inválida."
        txt_dtmDataDaParcela.SetFocus
        Exit Function
    ElseIf gblnExisteValorNaTabela(gstrDiasNaoUteis, "dtmData", gstrConvDtParaSql(txt_dtmDataDaParcela.Text)) Then
        ExibeMensagem "Data de vencimento da parcela é um dia não útil."
        txt_dtmDataDaParcela.SetFocus
        Exit Function
    ElseIf Trim(txt_intExercicio.Text) = "" Then
        ExibeMensagem "Digite o Exercício da parcela."
        txt_intExercicio.SetFocus
        Exit Function
    End If
    If Not blnAlterandoParcela Then
        blnInicioSequencia = True
        blnTemSequencia = False
        
        Set adoResultado = tdb_Parcelas.DataSource
        With adoResultado
            If .RecordCount > 0 Then
                intPosicao = .AbsolutePosition
                .MoveFirst
                While Not .EOF
                    If Val(txt_intNumero.Text) = Val(!intNumero) And Val(txt_intExercicio.Text) = Val(!intExercicio) Then
                        ExibeMensagem "Número de parcela já cadastrado para o Exercício " & !intExercicio & "."
                        txt_intNumero.SetFocus
                        .Move intPosicao, 0
                        Set adoResultado = Nothing
                        Exit Function
                    End If
                    If Val(txt_intExercicio.Text) = Val(!intExercicio) Then
                        blnInicioSequencia = False
                        If (Val(txt_intNumero.Text) - 1) = Val(!intNumero) Then
                            blnTemSequencia = True
                        End If
                    End If
                    .MoveNext
                Wend
            End If
' Identado MRA .Move intPosicao - 1, 1
        End With
        
        Set adoResultado = Nothing
        If (Not blnInicioSequencia) And (Not blnTemSequencia) Then
            ExibeMensagem "Número de parcela fora da sequência das parcelas para o Exercício " & txt_intExercicio.Text & "."
            txt_intNumero.SetFocus
            Exit Function
        End If
    End If
    ParcelaOK = True
End Function

Function strQuerryRelatorio() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT VC.PKId as Codigo, VC.strNomeDoVencimento as NomeDoVencimento, VP.intNumero as NumeroDeParcelas, VP.dtmDataDaParcela as DataDaParcela "
    strSql = strSql & " FROM " & gstrVencimentos & " VC,"
    strSql = strSql & gstrVencimentosDasParcelas & " VP "
    If blnAlterando = True Then
'        strSql = strSql & " WHERE VP.intVencimento =* VC.PKId and VC.PKId = " & tdb_Vencimento.Columns("PKId").Value
        strSql = strSql & " WHERE VP.intVencimento " & strOUTJOracle & "=" & strOUTJSQLServer & " VC.PKId and VC.PKId = " & tdb_Vencimento.Columns("PKId").Value
        Else
'        strSql = strSql & " WHERE VP.intVencimento =* VC.PKId"
        strSql = strSql & " WHERE VP.intVencimento =" & strOUTJSQLServer & " VC.PKId" & strOUTJOracle
    End If
    strSql = strSql & " ORDER BY VP.PKId"
    strQuerryRelatorio = strSql
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark As Variant
    Dim strSql      As String
    Dim strOperacao As String
    
    Select Case UCase(strModoOperacao)
        Case gstrNovo
            If tab_3dPasta.Tab = 0 Then
                LimpaObjeto Me, blnAlterando
                NovoVencimento
            Else
                NovaParcela
                txt_intNumero.SetFocus
            End If
        Case gstrSalvar
            If tab_3dPasta.Tab = 0 Then
                If blnAlterando Then
                    strOperacao = "A"
                Else: strOperacao = "I"
                End If
                If SalvarGeral(gstrVencimentos, strOperacao, Me, tdb_Vencimento, strQuery) Then
                    NovoVencimento
                End If
            Else
                If GravaParcela Then
                    NovaParcela
                End If
            End If
        
        Case gstrDeletar
            If tab_3dPasta.Tab = 0 Then
                If blnDeletaVencimento Then
                    LimpaObjeto Me
                    NovoVencimento
                    LeDaTabelaParaObj gstrVencimentos, tdb_Vencimento, strQuery
                End If
            Else
                 If DeletaParcela Then
                    NovaParcela
                End If
            End If
        Case gstrAplicar
            
        Case gstrGrade
            
        Case gstrImprimir
            ImprimeRelatorio rptVencimentos, strQuerryRelatorio
            
        Case gstrRefresh
            LeDaTabelaParaObj gstrVencimentos, tdb_Vencimento, strQuery
            
        Case gstrLocalizar, gstrPreencherLista
            ToolBarGeral strModoOperacao, gstrVencimentos, False, tdb_Vencimento, Me, mobjAux, strQuery
            
        Case gstrFechar
            Unload Me
            
    End Select
        
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
End Sub

