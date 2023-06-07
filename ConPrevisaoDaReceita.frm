VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmConPrevisaoDaReceita 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Previsão da Receita"
   ClientHeight    =   6300
   ClientLeft      =   240
   ClientTop       =   1245
   ClientWidth     =   9570
   HelpContextID   =   27
   Icon            =   "ConPrevisaoDaReceita.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6165
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   10874
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Previsão da Receita"
      TabPicture(0)   =   "ConPrevisaoDaReceita.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbldblValor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintFonteRecurso"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintCodigoorcamentario"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintCodigoReduzido"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_Total"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrLegislacao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tdb_Lista"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtdblValor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPKId"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_Total"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtintCodigoReduzido"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtintExercicio"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtintCodigoOrcamentario"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtstrCodigoOrcamentario"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtstrFonteRecurso"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtstrLegislacao"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.TextBox txtstrLegislacao 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1650
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   15
         Top             =   1410
         Width           =   7665
      End
      Begin VB.TextBox txtstrFonteRecurso 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1650
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         Top             =   1080
         Width           =   7665
      End
      Begin VB.TextBox txtstrCodigoOrcamentario 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3300
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   13
         Top             =   750
         Width           =   6015
      End
      Begin VB.TextBox txtintCodigoOrcamentario 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1650
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   12
         Top             =   750
         Width           =   1665
      End
      Begin VB.TextBox txtintExercicio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   11
         Top             =   30
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtintCodigoReduzido 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1650
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   10
         Top             =   420
         Width           =   1665
      End
      Begin VB.TextBox txt_Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   4020
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         Top             =   1740
         Width           =   1485
      End
      Begin VB.TextBox txtPKId 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8640
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1650
         MaxLength       =   15
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   1740
         Width           =   1485
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   3945
         Left            =   150
         TabIndex        =   0
         Top             =   2100
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   6959
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
         Columns(1).Caption=   "Cod. Reduzido"
         Columns(1).DataField=   "intCodigoReduzido"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Código"
         Columns(2).DataField=   "strCodigoOrcamentario"
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descrição"
         Columns(3).DataField=   "strDescricao"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor"
         Columns(4).DataField=   "dblValor"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Total"
         Columns(5).DataField=   "dblTotal"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "bytAprovado"
         Columns(6).DataField=   "bytAprovado"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2037"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1958"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2461"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2381"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=8652"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=8573"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=2461"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2381"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(34)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(41)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(65)  =   "Named:id=33:Normal"
         _StyleDefs(66)  =   ":id=33,.parent=0"
         _StyleDefs(67)  =   "Named:id=34:Heading"
         _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   ":id=34,.wraptext=-1"
         _StyleDefs(70)  =   "Named:id=35:Footing"
         _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   "Named:id=36:Selected"
         _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   "Named:id=37:Caption"
         _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(76)  =   "Named:id=38:HighlightRow"
         _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   "Named:id=39:EvenRow"
         _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(80)  =   "Named:id=40:OddRow"
         _StyleDefs(81)  =   ":id=40,.parent=33"
         _StyleDefs(82)  =   "Named:id=41:RecordSelector"
         _StyleDefs(83)  =   ":id=41,.parent=34"
         _StyleDefs(84)  =   "Named:id=42:FilterBar"
         _StyleDefs(85)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblstrLegislacao 
         AutoSize        =   -1  'True
         Caption         =   "Legislacao"
         Height          =   195
         Left            =   840
         TabIndex        =   16
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_Total 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label lblintCodigoReduzido 
         AutoSize        =   -1  'True
         Caption         =   "Código Reduzido"
         Height          =   195
         Left            =   390
         TabIndex        =   7
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label lblintCodigoorcamentario 
         AutoSize        =   -1  'True
         Caption         =   "Código Orçamentário"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblintFonteRecurso 
         AutoSize        =   -1  'True
         Caption         =   "Fonte de Recurso"
         Height          =   195
         Left            =   330
         TabIndex        =   5
         Top             =   1110
         Width           =   1275
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   1245
         TabIndex        =   4
         Top             =   1770
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmConPrevisaoDaReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnClickOk             As Boolean
    Dim mobjAux                 As Object
    Dim mblnSelecionou          As Boolean
    Dim mstrQueryAplicar        As String
    Dim mblocalizar             As Boolean

Private Function strQueryAplicar() As String
    Dim strSql As String
    If Trim(mstrQueryAplicar) = "" Then
        strSql = ""
        strSql = strSql & "SELECT CO.PKId, CO.strDescricao FROM "
        strSql = strSql & gstrCodigoOrcamentario & " CO, "
        strSql = strSql & gstrPrevisaoDaReceita & " PR "
        strSql = strSql & "WHERE CO.PKId = PR.intCodigoOrcamentario "
        strSql = strSql & "ORDER BY CO.strDescricao"
        strQueryAplicar = strSql
    Else
        strQueryAplicar = mstrQueryAplicar
    End If
End Function

Private Function strQueryCO() As String

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PR.PKId, PR.intCodigoReduzido, CO.strCodigoOrcamentario, "
    strSql = strSql & "CO.strDescricao, PR.dblValor, PR.bytSituacao, "
'    strSql = strSql & "(SELECT ISNULL(SUM(dblValor), 0) FROM "
    strSql = strSql & "(SELECT " & gstrISNULL("SUM(dblValor)", "0") & " FROM "
    strSql = strSql & gstrPrevisaoDaReceita & " PR WHERE PR.bytSituacao = 1) AS dblTotal "
    strSql = strSql & "FROM "
    strSql = strSql & gstrPrevisaoDaReceita & " PR, "
    strSql = strSql & gstrCodigoOrcamentario & " CO "
    strSql = strSql & "WHERE PR.intCodigoOrcamentario = CO.PKId "
    strSql = strSql & "AND PR.bytSituacao = 1 "
    strSql = strSql & "AND PR.intExercicio = " & gintExercicio
    strSql = strSql & "ORDER BY CO.strCodigoOrcamentario"
    strQueryCO = strSql
End Function

Private Sub Form_Activate()
    gintCodSeguranca = 761
    VirificaGradeListView Me
    txtintExercicio = gintExercicio
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrNovo, gstrSalvar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
    VerificaListaAutomatica "", tdb_Lista, strQueryCO
    VerificaObjParaAplicar mobjAux, mstrQueryAplicar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mblocalizar = False
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
    mblocalizar = True
End Sub

Private Sub tdb_Lista_DataSourceChanged()
    With tdb_Lista
        txt_Total = .Columns("dblTotal")
    End With
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
    mblocalizar = True
End Sub

Private Sub tdb_Lista_FilterChange()
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Value = gvntFormatacaoEspecifica(Value, 2)
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
    mblocalizar = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Lista
    mblocalizar = True
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
    mblocalizar = True
End Sub

Private Sub LePrevisao()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    strSql = ""
    strSql = strSql & "SELECT CO.PKId, CO.strCodigoOrcamentario, "
    strSql = strSql & "CO.strDescricao, FR.strDescricao AS strFonte, "
    strSql = strSql & "PV.intCodigoReduzido, PV.dblValor, PV.strLegislacao "
    strSql = strSql & "FROM " & gstrFonteRecurso & " FR, "
    strSql = strSql & gstrCodigoOrcamentario & " CO, "
    strSql = strSql & gstrPrevisaoDaReceita & " PV "
    strSql = strSql & "WHERE PV.intCodigoOrcamentario = CO.PKId "
'    strSql = strSql & "AND PV.intFonteRecurso *= FR.PKId "
    strSql = strSql & "AND PV.intFonteRecurso " & strOUTJOracle & strOUTJSQLServer & "= FR.PKId "
    strSql = strSql & "AND PV.PKId = " & Val(tdb_Lista.Columns("PKID").Value)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txtintCodigoReduzido = gstrENulo(!intCodigoReduzido)
                txtintCodigoOrcamentario = gvntFormatacaoEspecifica(!strCodigoOrcamentario, 2)
                txtstrCodigoOrcamentario = !strDescricao
                txtstrFonteRecurso = gstrENulo(!strFonte)
                txtstrLegislacao = gstrENulo(!strLegislacao)
                txtdblValor = gstrConvVrDoSql(!dblValor)
                txtPKId = !Pkid
            Else
                txtintCodigoReduzido = ""
                txtintCodigoOrcamentario = ""
                txtstrCodigoOrcamentario = ""
                txtstrFonteRecurso = ""
                txtstrLegislacao = ""
                txtdblValor = ""
            End If
        End With
    End If
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            gCorLinhaSelecionada tdb_Lista
            LePrevisao
            'If mobjAux Is Nothing Then
            '    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            'Else
            '    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            'End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    If strModoOperacao = gstrLocalizar Then
        If mblocalizar = True Then
            Exit Sub
        End If
    End If
    ToolBarGeral strModoOperacao, gstrPrevisaoDaReceita, False, _
                 tdb_Lista, Me, mobjAux, strQueryCO, strQueryAplicar, _
                 rptPrevisaoDaReceita, strQueryRelatorio
                 
    If strModoOperacao = gstrNovo Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrNovo, gstrSalvar
    End If
    
End Sub

Private Sub txtPKId_Change()
    txtPKId = txtPKId
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txt_Total
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblValor = gstrConvVrDoSql(txtdblValor)
End Sub

Public Function strQueryRelatorio()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT PR.intCodigoreduzido AS CodigoReduzido, "
    strSql = strSql & "CO.strCodigoOrcamentario, "
    strSql = strSql & "CO.strDescricao AS DesCodigoOrcamentario, "
    strSql = strSql & "FR.strDescricao AS FonteRecurso, PR.dblValor "
    strSql = strSql & "FROM "
    strSql = strSql & gstrPrevisaoDaReceita & " PR, "
    strSql = strSql & gstrCodigoOrcamentario & " CO, "
    strSql = strSql & gstrFonteRecurso & " FR "
'    strSql = strSql & "WHERE PR.intCodigoOrcamentario *= CO.PKId "
    strSql = strSql & "WHERE PR.intCodigoOrcamentario " & strOUTJSQLServer & "= CO.PKId " & strOUTJOracle
'    strSql = strSql & "AND PR.intFonteRecurso *= FR.PKId "
    strSql = strSql & "AND PR.intFonteRecurso " & strOUTJSQLServer & "= FR.PKId " & strOUTJOracle
    strSql = strSql & "ORDER BY CO.strCodigoOrcamentario"
    strQueryRelatorio = strSql
End Function
