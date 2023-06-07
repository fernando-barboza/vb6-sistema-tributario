VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadDespesaExtraOrcamentaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despesa Extra-Orçamentária"
   ClientHeight    =   5430
   ClientLeft      =   2370
   ClientTop       =   2910
   ClientWidth     =   8835
   HelpContextID   =   4
   Icon            =   "CadDespesaExtraOrcamentaria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8835
   Begin VB.TextBox txtPKId 
      Height          =   285
      Left            =   6720
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   825
   End
   Begin TabDlg.SSTab tab_3DDados 
      Height          =   5175
      Left            =   90
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Despesa Extra-Orçamentária"
      TabPicture(0)   =   "CadDespesaExtraOrcamentaria.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbldtmData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbldblValor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDesconto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintContaExtra"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintContribuinte"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrNumero"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbcintContribuinte"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tdb_Lista"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_HistoricoSubEmpenho"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtdtmData"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmd_Credor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmd_Descricao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtintNumero"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fra_Processo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtdblDesconto"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cbo_DescricaoExtra"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtdblValor"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_intNContribuinte"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cbointContaContabil"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      Begin VB.ComboBox cbointContaContabil 
         Height          =   315
         ItemData        =   "CadDespesaExtraOrcamentaria.frx":105E
         Left            =   930
         List            =   "CadDespesaExtraOrcamentaria.frx":1060
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1260
         Width           =   1455
      End
      Begin VB.TextBox txt_intNContribuinte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   930
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   840
         Width           =   1440
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2430
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         TabIndex        =   1
         Top             =   450
         Width           =   1185
      End
      Begin VB.ComboBox cbo_DescricaoExtra 
         Height          =   315
         Left            =   2430
         TabIndex        =   7
         Top             =   1260
         Width           =   5745
      End
      Begin VB.TextBox txtdblDesconto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4725
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         TabIndex        =   2
         Top             =   450
         Width           =   1185
      End
      Begin VB.Frame fra_Processo 
         Caption         =   " Processo "
         Height          =   765
         Left            =   90
         TabIndex        =   22
         Top             =   1665
         Width           =   1725
         Begin VB.TextBox txt_strcodigo 
            CausesValidation=   0   'False
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
            Left            =   60
            MaxLength       =   15
            TabIndex        =   8
            Top             =   360
            Width           =   825
         End
         Begin VB.TextBox txt_intExercicio 
            CausesValidation=   0   'False
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
            Left            =   900
            MaxLength       =   4
            TabIndex        =   9
            Top             =   360
            Width           =   465
         End
         Begin VB.TextBox txt_bitDigito 
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
            Left            =   1380
            MaxLength       =   2
            TabIndex        =   10
            Top             =   360
            Width           =   285
         End
      End
      Begin VB.TextBox txtintNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   930
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         TabIndex        =   0
         Top             =   450
         Width           =   825
      End
      Begin VB.CommandButton cmd_Descricao 
         Height          =   300
         Left            =   8190
         Picture         =   "CadDespesaExtraOrcamentaria.frx":1062
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Clique para cadastar conta"
         Top             =   1260
         Width           =   330
      End
      Begin VB.CommandButton cmd_Credor 
         Height          =   300
         Left            =   8190
         Picture         =   "CadDespesaExtraOrcamentaria.frx":13EC
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastrar o contribuint"
         Top             =   840
         Width           =   330
      End
      Begin VB.TextBox txtdtmData 
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
         Left            =   6675
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   435
         Width           =   1005
      End
      Begin VB.Frame fra_HistoricoSubEmpenho 
         Caption         =   " Histórico "
         Height          =   1185
         Left            =   1890
         TabIndex        =   17
         Top             =   1665
         Width           =   6660
         Begin VB.TextBox txt_strCodigoHistorico 
            Height          =   315
            Left            =   105
            TabIndex        =   12
            Top             =   795
            Width           =   795
         End
         Begin VB.CommandButton cmd_Historico 
            Height          =   300
            Left            =   6210
            Picture         =   "CadDespesaExtraOrcamentaria.frx":1776
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Clique aqui para cadastrar histórico"
            Top             =   795
            Width           =   330
         End
         Begin VB.TextBox txtstrHistorico 
            Height          =   589
            Left            =   105
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   180
            Width           =   6450
         End
         Begin MSDataListLib.DataCombo dbc_Historico 
            Height          =   315
            Left            =   945
            TabIndex        =   13
            Top             =   795
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2115
         Left            =   90
         TabIndex        =   14
         Top             =   2925
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   3731
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
         Columns(1).Caption=   "Número"
         Columns(1).DataField=   "intNumero"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Conta"
         Columns(2).DataField=   "strContaContabil"
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descrição"
         Columns(3).DataField=   "strDescricao"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Situação"
         Columns(4).DataField=   "strSituacao"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "O.P."
         Columns(5).DataField=   "OP"
         Columns(5).NumberFormat=   "General Number"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Valor"
         Columns(6).DataField=   "dblValor"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1826"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1746"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2461"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2381"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=7276"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=7197"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=1693"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1614"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(35)=   "Column(6).Width=2461"
         Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2381"
         Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=188,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14,.alignment=2"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14,.alignment=2"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintContribuinte 
         Height          =   315
         Left            =   2430
         TabIndex        =   5
         Top             =   840
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblstrNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   300
         TabIndex        =   26
         Top             =   495
         Width           =   555
      End
      Begin VB.Label lblintContribuinte 
         AutoSize        =   -1  'True
         Caption         =   "Credor"
         Height          =   195
         Left            =   390
         TabIndex        =   25
         Top             =   885
         Width           =   465
      End
      Begin VB.Label lblintContaExtra 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         Height          =   195
         Left            =   450
         TabIndex        =   24
         Top             =   1350
         Width           =   420
      End
      Begin VB.Label lblDesconto 
         AutoSize        =   -1  'True
         Caption         =   "Desconto"
         Height          =   195
         Left            =   3945
         TabIndex        =   23
         Top             =   495
         Width           =   690
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   2010
         TabIndex        =   20
         Top             =   495
         Width           =   360
      End
      Begin VB.Label lbldtmData 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   6240
         TabIndex        =   18
         Top             =   495
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmCadDespesaExtraOrcamentaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mobjAux             As Object
    Dim mstrQueryAplicar    As String
    Dim mblnAlterando       As Boolean
    Dim mblnSelecionou      As Boolean
    Dim mblnClickOk         As Boolean
    Dim itemAnterior        As String

Private Function blnDataDespesaOK()
Dim dtmDtEncerramento As Date

    If gblnDataValida(txtdtmData) Then
    
        dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
        
        If dtmDtEncerramento = Empty Then
           If txtdtmData.Enabled Then txtdtmData.SetFocus
           Exit Function
        Else
           If CDate(txtdtmData) <= dtmDtEncerramento Then
              ExibeMensagem "A data deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
              If txtdtmData.Enabled Then txtdtmData.SetFocus
              Exit Function
           ElseIf Year(txtdtmData) <> gintExercicio Then
              ExibeMensagem "O ano não pode ser diferente do exercício."
              If txtdtmData.Enabled Then txtdtmData.SetFocus
              Exit Function
           End If
        End If
    Else
        ExibeMensagem "A data da despesa tem que ser informada corretamente."
        If txtdtmData.Enabled Then txtdtmData.SetFocus
        Exit Function
    End If
    
    If blnValidarProcesso Then
       If Len(Trim(txt_strcodigo)) > 0 Or Len(Trim(txt_bitDigito)) > 0 Or Len(Trim(txt_intExercicio)) > 0 Then
          If VerificaEmpenhoProcesso = "NULL" Then
             ExibeMensagem "Processo não localizado."
             If txt_strcodigo.Enabled Then txt_strcodigo.SetFocus
             Exit Function
          End If
       End If
    End If
    
    If Val(gstrConvVrParaSql(txtdblValor)) = 0 Then
        ExibeMensagem "O campo valor não aceita nulo."
        txtdblValor.SetFocus
        Exit Function
    ElseIf Val(txtintNumero) = 0 Then
        ExibeMensagem "O número não pode ser nulo."
        If txtintNumero.Enabled Then txtintNumero.SetFocus
        Exit Function
    'ElseIf gblnExisteValorNaTabela(gstrDespesaExtraOrcamentaria, "intNumero", txtintNumero) And Not mblnAlterando Then
    ElseIf gblnExisteCodigo(2, gstrDespesaExtraOrcamentaria, "intNumero", txtintNumero, "intExercicio", Val(gintExercicio)) And Not mblnAlterando Then
        ExibeMensagem "Este número já existe na tabela e não pode ser repetido."
        localizaDespesabyNumero
        If txtintNumero.Enabled Then txtintNumero.SetFocus
        Exit Function
    ElseIf dbcintContribuinte.MatchedWithList = False Then
        ExibeMensagem "O contribuinte tem que ser informado corretamente."
        dbcintContribuinte.SetFocus
        Exit Function
    ElseIf cbointContaContabil.ListIndex = -1 Then
        ExibeMensagem "A conta tem que ser informada corretamente."
        cbointContaContabil.SetFocus
        Exit Function
    End If
    If Val(txtdblDesconto) <> 0 Then
        If txtdblDesconto < 0 Then
            ExibeMensagem "O Desconto não pode ser negativo."
            If txtdblDesconto.Enabled Then txtdblDesconto.SetFocus
            Exit Function
        End If
        If CDbl(txtdblDesconto) > CDbl(txtdblValor) Then
            ExibeMensagem "O Desconto não pode ser maior que o Valor."
            If txtdblDesconto.Enabled Then txtdblDesconto.SetFocus
            Exit Function
        End If
    End If
    
    blnDataDespesaOK = True
    
End Function
Private Function VerificaEmpenhoProcesso() As String
   Dim strSql As String
   Dim adoResultado As New ADODB.Recordset
   
   strSql = "SELECT PP.PKID FROM " & gstrProtocolizacaoProcesso & " PP "
   strSql = strSql & "WHERE RTRIM(LTRIM(PP.strCodigo)) = '" & Trim(txt_strcodigo) & "' AND "
   strSql = strSql & IIf(Len(Trim(txt_bitDigito)) > 0, "RTRIM(LTRIM(PP.bitDigito)) =  " & Trim(txt_bitDigito), "bitDigito IS NULL") & " AND "
   strSql = strSql & IIf(Len(Trim(txt_intExercicio)) > 0, "RTRIM(LTRIM(PP.intExercicio)) =  " & Trim(txt_intExercicio), "intExercicio IS NULL")
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         VerificaEmpenhoProcesso = adoResultado!Pkid
       Else
         VerificaEmpenhoProcesso = "NULL"
      End If
   End If
   
End Function
Private Sub GravaDespesaExtra()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql  As String
        
    If blnDataDespesaOK Then
        If gblnExclusaoGravacaoOk(IIf(mblnAlterando, "A", "I"), " Despesa extra-orçamentária") Then
            strSql = ""
            If mblnAlterando Then
                strSql = strSql & "UPDATE " & gstrDespesaExtraOrcamentaria & " SET "
                strSql = strSql & "intContribuinte = " & dbcintContribuinte.BoundText & ", "
                strSql = strSql & "intContaContabil = " & gstrItemData(cbointContaContabil) & ", "
                strSql = strSql & "dblValor = " & gstrConvVrParaSql(txtdblValor) & ", "
                strSql = strSql & "dblDesconto = " & gstrConvVrParaSql(txtdblDesconto) & ", "
                strSql = strSql & "dtmData = " & gstrConvDtParaSql(txtdtmData) & ", "
                strSql = strSql & "strHistorico = '" & Trim(txtstrHistorico) & "', "
                 strSql = strSql & "intProtocolizacaoProcesso = " & (VerificaEmpenhoProcesso) & " "
                strSql = strSql & "WHERE PKId = " & Val(txtPKId)
            Else
                strSql = strSql & "INSERT INTO " & gstrDespesaExtraOrcamentaria & " ("
                strSql = strSql & "intNumero, intContribuinte, intContaContabil, "
                strSql = strSql & "dblValor, dblDesconto, bytSituacao, dtmData, strHistorico, "
                strSql = strSql & "intExercicio, dtmDtAtualizacao, lngCodUsr,intProtocolizacaoProcesso) "
'                strSql = strSql & "SELECT ISNULL(MAX(intNumero), 0) + 1, "
                'StrSql = StrSql & "SELECT " & gstrISNULL("MAX(intNumero)", "0") & " + 1, "
                strSql = strSql & " VALUES (" & txtintNumero & ", "
                strSql = strSql & dbcintContribuinte.BoundText & ", "
                strSql = strSql & gstrItemData(cbointContaContabil) & ", "
                strSql = strSql & gstrConvVrParaSql(txtdblValor) & ", "
                strSql = strSql & gstrConvVrParaSql(txtdblDesconto) & ", "
                strSql = strSql & " 0, "
                strSql = strSql & gstrConvDtParaSql(txtdtmData) & ", "
                strSql = strSql & "'" & Trim(txtstrHistorico) & "', "
                strSql = strSql & gintExercicio & ", "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ","
                strSql = strSql & glngCodUsr & ","
                strSql = strSql & "" & (VerificaEmpenhoProcesso) & ")"
                'StrSql = StrSql & "FROM " & gstrDespesaExtraOrcamentaria
            End If
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSql) Then
                LimpaObjeto Me, mblnAlterando
                txt_bitDigito.Text = ""
                txt_strcodigo.Text = ""
                txt_intExercicio.Text = ""
                'VerificaListaAutomatica "", tdb_Lista, strQuery
                MantemForm gstrNovo
                MantemForm gstrLocalizar
            End If
        End If
    End If
End Sub

Private Sub cbo_DescricaoExtra_DropDown()
    If cbo_DescricaoExtra.ListIndex = -1 Then
        cbointContaContabil.Text = ""
    End If
    ComboPlanoDeConta
End Sub

Private Sub cbointContaContabil_DropDown()
    If cbointContaContabil.ListIndex = -1 Then
        cbo_DescricaoExtra.Text = ""
    End If
    ComboPlanoDeConta
End Sub

Private Sub cbointContaContabil_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", cbointContaContabil
End Sub

Private Sub cbo_DescricaoExtra_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cbo_DescricaoExtra
End Sub

Private Sub cmd_Credor_Click()
    CarregaForm frmCadContribuinte, dbcintContribuinte
    frmCadContribuinte.Caption = "Cadastro de Credores"
    frmCadContribuinte.Tag = "Credor"
End Sub

Private Sub cmd_Descricao_Click()
    CarregaForm frmCadPlanoConta, cbo_DescricaoExtra, strQueryAplicar
End Sub

Private Sub cmd_Historico_Click()
    CarregaForm frmCadHistorico, dbc_Historico
End Sub

Private Sub dbc_Historico_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_Historico, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinte_Click(Area As Integer)
    DropDownDataCombo dbcintContribuinte, Me, Area
    
    If Area = 2 Then
        If dbcintContribuinte.MatchedWithList Then
            txt_intNContribuinte = LeCDCCredor(dbcintContribuinte.BoundText)
            If itemAnterior = dbcintContribuinte.BoundText Then Exit Sub
            itemAnterior = dbcintContribuinte.BoundText
       End If
    End If
End Sub

Private Sub dbcintContribuinte_GotFocus()
    If dbcintContribuinte.BoundText <> txt_intNContribuinte And dbcintContribuinte.BoundText <> "" Then
         dbcintContribuinte_Click 0
    End If
End Sub

Private Sub dbcintContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinte_KeyPress(KeyAscii As Integer)
  CaracterValido KeyAscii
End Sub

Private Sub dbc_Historico_Click(Area As Integer)
    Dim adoResultado As ADODB.Recordset
    Dim strSql       As String
    
    DropDownDataCombo dbc_Historico, Me, Area
    txtstrHistorico = dbc_Historico.Text
    txt_strCodigoHistorico = dbc_Historico.BoundText
    
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & " strCodigo "
    strSql = strSql & "FROM "
    strSql = strSql & gstrHistorico
    strSql = strSql & " WHERE "
    strSql = strSql & " pkid = '" & dbc_Historico.BoundText & "' "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Me.txt_strCodigoHistorico.Text = gstrENulo(!strCodigo)
                Me.txtstrHistorico.Text = Me.dbc_Historico.Text
            End If
        End With
    End If
    
End Sub

Private Sub dbc_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 287
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    dbcintContribuinte.Tag = strQueryContribuinte & ";strNome"
    dbc_Historico.Tag = "SELECT PKID, strDescricao FROM " & gstrHistorico & " ORDER BY strDescricao;strDescricao"
    VerificaListaAutomatica "", tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux, mstrQueryAplicar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Function strQuery() As String

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT DE.PKId, DE.intNumero, PC.strContaContabil, "
    strSql = strSql & "PC.strDescricao, DE.bytSituacao, DE.dblValor,  "
'    strSql = strSql & "CASE DE.bytSituacao "
'    strSql = strSql & "WHEN 0 THEN 'Programada' "
'    strSql = strSql & "WHEN 2 THEN 'Paga' END AS strSituacao "
    strSql = strSql & gstrCASEWHEN("DE.bytSituacao", _
        "0, 'Programada', 2, 'Paga'") & " AS strSituacao "
    
    strSql = strSql & "FROM "
    strSql = strSql & gstrDespesaExtraOrcamentaria & " DE, "
    strSql = strSql & gstrPlanoConta & " PC "
    strSql = strSql & "WHERE PC.PKId = DE.intContaContabil "
    strSql = strSql & "ORDER BY DE.intNumero"
    'StrSql = StrSql & "ORDER BY PC.strContaContabil"
    strQuery = strSql
End Function

Private Function strQueryAplicar() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrPlanoConta & " "
    strSql = strSql & "WHERE ABS(blnExtraOrcamentaria) = 1 "
    strSql = strSql & "AND ABS(blnAnalitica) = 1 "
    strSql = strSql & "ORDER BY strDescricao"
    strQueryAplicar = strSql
End Function

Private Sub cbointContaContabil_Click()
    cbo_DescricaoExtra.ListIndex = gintIndiceCBO(cbo_DescricaoExtra, _
                                                 gstrItemData(cbointContaContabil))
End Sub

Private Sub cbo_DescricaoExtra_Click()
    cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, _
                                                  gstrItemData(cbo_DescricaoExtra))
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 2 Then
        Value = gvntFormatacaoEspecifica(Value, 1)
    End If
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
     gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Lista
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId = .Columns("PKID").Value
            If cbointContaContabil.ListCount = 0 Then
                ComboPlanoDeConta
            End If
            
            LeDaTabelaParaObj gstrDespesaExtraOrcamentaria, Me
            RetornaProcesso
'            LeDaTabelaParaObj "", dbcintContribuinte, strQueryContribuinte
            txt_strCodigoHistorico.Text = ""
            dbc_Historico.Text = ""

            
            txt_intNContribuinte.Text = RetornaCredor(IIf(dbcintContribuinte.BoundText = "", 0, dbcintContribuinte.BoundText), False)
            
            If Trim$(IIf(txt_intNContribuinte = "", 0, txt_intNContribuinte)) = 0 Then
                txt_intNContribuinte = Space$(0)
                dbcintContribuinte.Text = Space$(0)
            End If
            
            txt_intNContribuinte_LostFocus
            
            gCorLinhaSelecionada tdb_Lista
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            mblnSelecionou = True
            mblnAlterando = True
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strCodigo As String
    If UCase(strModoOperacao) = gstrSalvar Then
        If mblnAlterando = False Then
            'If gblnExisteCodigo(1, gstrDespesaExtraOrcamentaria, "intnumero", "'" & txtintNumero.Text & "'") Then
            If gblnExisteCodigo(2, gstrDespesaExtraOrcamentaria, "intNumero", txtintNumero, "intExercicio", Val(gintExercicio)) Then
                strCodigo = (gstrProximoCodigo(txtintNumero, gstrDespesaExtraOrcamentaria, "intNumero", gintCodSeguranca, "intExercicio", Val(gintExercicio), , True, , , "intExercicio", Val(gintExercicio)))
                If MsgBox("O código informado já se encontra cadastrado. Deseja cadastrar usando o novo código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                    txtintNumero.SetFocus
                    Exit Sub
                Else
                    txtintNumero.Text = strCodigo
                    GravaDespesaExtra
                End If
            Else
                GravaDespesaExtra
            End If
        Else
            GravaDespesaExtra
        End If
            
    Else
        If UCase(strModoOperacao) = gstrDeletar Then
            If UCase(tdb_Lista.Columns(4)) = "PAGA" Then
                MsgBox "Não é permitido excuir uma despesa que já foi paga.", vbExclamation, "Atenção"
                Exit Sub
            End If
        End If
        
        If UCase(strModoOperacao) = UCase(gstrPreencherLista) Then

'            If Me.ActiveControl.Name = dbc_Historico.Name Then
'                   LeDaTabelaParaObj gstrHistorico, dbc_Historico
'            End If

            If Me.ActiveControl.Name = cbointContaContabil.Name Or Me.ActiveControl.Name = cbo_DescricaoExtra.Name Then
                   cbointContaContabil.Text = ""
                   cbo_DescricaoExtra.Text = ""
                   ComboPlanoDeConta
                   Exit Sub
            End If
        End If
        
        If UCase(strModoOperacao) = gstrImprimir Then
           
           ImprimeDocumentoExtra
        
        End If
        
        If UCase(strModoOperacao) = gstrLocalizar Then
            LeDaTabelaParaObj "", tdb_Lista, strQueryLocalizar
            Exit Sub
        End If
        
        ToolBarGeral strModoOperacao, gstrDespesaExtraOrcamentaria, mblnAlterando, _
                     tdb_Lista, Me, mobjAux, strQuery, mstrQueryAplicar
        
        If UCase(strModoOperacao) = gstrDeletar And Trim(Len(txtPKId)) = 0 Then
            txt_bitDigito.Text = ""
            txt_strcodigo.Text = ""
            txt_intExercicio.Text = ""
            txt_intNContribuinte = ""
        End If
                     
        If UCase(strModoOperacao) = gstrNovo Then
            txt_intNContribuinte = ""
            txt_bitDigito.Text = ""
            txt_strcodigo.Text = ""
            txt_intExercicio.Text = ""
            txt_strCodigoHistorico.Text = ""
            dbc_Historico.Text = ""
            proximoCodigoDespesa
            If txtintNumero.Enabled Then txtintNumero.SetFocus
        End If
        
    End If
End Sub


Private Sub txt_bitDigito_KeyPress(KeyAscii As Integer)
     CaracterValido KeyAscii, "N", txt_bitDigito
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
     CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_intNContribuinte_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "N", txt_intNContribuinte
End Sub

Private Sub txt_intNContribuinte_LostFocus()
Dim strPKId As String
Dim strSql As String

    strPKId = LeCDCCredor(, txt_intNContribuinte)
   
   If strPKId = "" Then
        dbcintContribuinte.BoundText = ""
        Exit Sub
   End If

   
    If Len(Trim(txt_intNContribuinte)) > 0 Then
        If dbcintContribuinte.Enabled Then dbcintContribuinte.SetFocus
            strSql = "SELECT CO.PKID,"
           strSql = strSql & " CO.STRNOME"
           strSql = strSql & " FROM "
           strSql = strSql & gstrContribuinte & " CO, "
           strSql = strSql & gstrItens & " IT, "
           strSql = strSql & gstrModuloContribuinte & " MC"
           strSql = strSql & " WHERE CO.PKID = " & strPKId & "AND"
           strSql = strSql & " IT.PKId = MC.intItem AND"
           strSql = strSql & " MC.intContribuinte = CO.Pkid AND"
           strSql = strSql & " IT.Pkid =" & gintModulo & " AND CO.BLNINATIVO = 0"
           LeDaTabelaParaObj gstrContribuinte, dbcintContribuinte, strSql
           dbcintContribuinte.BoundText = strPKId
    End If

End Sub

Private Sub txt_strCodigoHistorico_GotFocus()
    MarcaCampo txt_strCodigoHistorico
End Sub

Private Sub txt_strCodigoHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strCodigoHistorico
End Sub

Private Sub txt_strCodigoHistorico_LostFocus()
    Dim adoResultado As ADODB.Recordset
    Dim strSql       As String
    
    strSql = ""
    strSql = strSql & "SELECT strDescricao "
    strSql = strSql & "  FROM " & gstrHistorico
    strSql = strSql & " WHERE STRCODIGO = '" & Me.txt_strCodigoHistorico.Text & "'"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Me.dbc_Historico = gstrENulo(!strDescricao)
                Me.txtstrHistorico.Text = gstrENulo(!strDescricao)
            Else
                Me.dbc_Historico = ""
                Me.txtstrHistorico.Text = ""
            End If
        End With
    End If
End Sub

Private Sub txtdblDesconto_GotFocus()
    MarcaCampo txtdblDesconto
End Sub

Private Sub txtdblDesconto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblDesconto
End Sub

Private Sub txtdblDesconto_LostFocus()
    txtdblDesconto = gstrConvVrDoSql(txtdblDesconto, 2)
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblValor = gstrConvVrDoSql(txtdblValor)
End Sub

Private Sub txtdtmData_GotFocus()
    MarcaCampo txtdtmData
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmData
End Sub

Private Sub txtdtmData_LostFocus()

    txtdtmData = gstrDataFormatada(txtdtmData)
    
    'ORC677
    If IsDate(txtdtmData) Then
        If Year(CDate(txtdtmData)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data tem que estar no exercício de " & gintExercicio & "."
            If txtdtmData.Enabled Then txtdtmData.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txtintNumero_GotFocus()
    If Trim(txtintNumero) = "" Then
        proximoCodigoDespesa
    End If
    MarcaCampo txtintNumero
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumero
End Sub

Private Sub txtintNumero_LostFocus()
    localizaDespesabyNumero
End Sub

Private Sub localizaDespesabyNumero()
    If Len(Trim$(txtintNumero)) > 0 Then
        'If gblnExisteValorNaTabela(gstrDespesaExtraOrcamentaria, "intNumero", txtintNumero) Then
        If gblnExisteCodigo(2, gstrDespesaExtraOrcamentaria, "intNumero", txtintNumero, "intExercicio", Val(gintExercicio)) Then
            With frmCadDespesaExtraOrcamentaria
                If .tdb_Lista.EOF Then Exit Sub
                .tdb_Lista.MoveFirst
                Do While Not tdb_Lista.EOF
                    If .tdb_Lista.Columns("intNumero") = txtintNumero Then
                        gCorLinhaSelecionada .tdb_Lista
                        .LeByLinhaSelecionada
                        Exit Do
                    Else
                        .tdb_Lista.MoveNext
                    End If
                Loop
            End With
        End If
    End If
End Sub


Private Sub txtstrHistorico_KeyPress(KeyAscii As Integer)
 CaracterValido KeyAscii, "A", txtstrHistorico
End Sub

Private Function strQuerryRelatorio() As String

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT DE.PKId, DE.intContribuinte, CO.strNome AS CREDOR, DE.dtmData, DE.intNumero, PC.strContaContabil, "
    strSql = strSql & "PC.strDescricao, DE.bytSituacao, DE.dblValor,  "
'    strSql = strSql & "CASE DE.bytSituacao "
'    strSql = strSql & "WHEN 0 THEN 'Programada' "
'    strSql = strSql & "WHEN 2 THEN 'Paga' END AS strSituacao "
    strSql = strSql & gstrCASEWHEN("DE.bytSituacao", _
        "0, 'Programada', 2, 'Paga'") & " AS strSituacao "
    
    strSql = strSql & "FROM "
    strSql = strSql & gstrDespesaExtraOrcamentaria & " DE, "
    strSql = strSql & gstrPlanoConta & " PC, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & "WHERE PC.PKId = DE.intContaContabil "
    strSql = strSql & " AND DE.intContribuinte = CO.PKId "
    strSql = strSql & " AND DE.intExercicio = " & gintExercicio
    strSql = strSql & " ORDER BY CO.strNome, PC.strContaContabil"
strQuerryRelatorio = strSql
End Function

Public Sub LeByLinhaSelecionada()
    mblnClickOk = True
    tdb_Lista_RowColChange 0, 0
End Sub

Private Sub proximoCodigoDespesa()
       'gstrProximoCodigo txtintNumero, gstrDespesaExtraOrcamentaria, "intNumero", gintCodSeguranca, gstrDATEPART(strYEAR, "dtmData"), Val(gintExercicio), , , , , gstrDATEPART(strYEAR, "dtmData"), Val(gintExercicio)
       gstrProximoCodigo txtintNumero, gstrDespesaExtraOrcamentaria, "intNumero", gintCodSeguranca, "intExercicio", Val(gintExercicio), , , , , "intExercicio", Val(gintExercicio)
       'gstrProximoCodigo(txtintNumero, gstrReservaDotacao          , "intNumero", gintCodSeguranca, "intExercicioReserva", Val(gintExercicio), , True, , , "intExercicioReserva", Val(gintExercicio))
       
End Sub

Private Function strQueryContribuinte() As String
Dim strSql As String
    
'    strSQL = ""
'    strSQL = strSQL & "SELECT C.PKID, C.strNome "
'    strSQL = strSQL & "FROM "
'    strSQL = strSQL & gstrContribuinte & " C, "
'    strSQL = strSQL & gstrModuloContribuinte & " M "
'    strSQL = strSQL & "WHERE M.intItem = " & gintModulo
'    strSQL = strSQL & " AND C.PKId = M.intContribuinte"
'    strSQL = strSQL & " GROUP BY C.PKID, C.STRNOME"
'    strSQL = strSQL & " ORDER BY C.STRNOME"

    strSql = "SELECT PKID, strNome FROM " & gstrContribuinte & " ORDER BY strNome"
    strQueryContribuinte = strSql

End Function

Private Function RetornaCredor(lngCodCredor As Long, Optional blnID As Boolean = True) As String
Dim strSql       As String
Dim adoResultado As ADODB.Recordset

    strSql = "SELECT " & IIf(blnID, "C.PKId", "C.CDC") & " Codigo FROM " & gstrContribuinte & " C WHERE " & IIf(blnID, "C.CDC", "C.PKId") & " = " & lngCodCredor
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 30, adoResultado) Then
        If Not adoResultado.EOF Then
            RetornaCredor = IIf(IsNull(adoResultado!Codigo) = True, 0, adoResultado!Codigo)
        End If
        adoResultado.Close: Set adoResultado = Nothing
    End If
    
End Function
Private Function RetornaProcesso() As String
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    strSql = ""
    strSql = strSql & "Select PP.intexercicio,PP.strcodigo,PP.bitdigito "
    strSql = strSql & "from " & gstrDespesaExtraOrcamentaria & " DE," & gstrProtocolizacaoProcesso & " PP where PP.pkid = DE.intProtocolizacaoProcesso "
    strSql = strSql & "and DE.pkid = " & txtPKId & ""
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txt_strcodigo = !strCodigo
                txt_intExercicio = !intExercicio
                txt_bitDigito = !bitDigito
            End With
        Else
                txt_strcodigo.Text = ""
                txt_intExercicio.Text = ""
                txt_bitDigito.Text = ""
        End If
        adoResultado.Close: Set adoResultado = Nothing
    End If
    
End Function

Private Sub ComboPlanoDeConta()
Dim strSql       As String
Dim adoResultado As ADODB.Recordset
    
    strSql = strSql & "SELECT PKId, strContaContabil, strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrPlanoConta
    strSql = strSql & " WHERE blnExtraOrcamentaria = 1 AND blnAnalitica = 1 "
    
    If cbo_DescricaoExtra.Text <> "" Then
            strSql = strSql & " AND UPPER(strDescricao) LIKE " & UCase(IIf(Mid(cbo_DescricaoExtra.Text, 1, 1) = "%", "'%", "'") & cbo_DescricaoExtra.Text & "%'")
    End If
    
    If cbointContaContabil.Text <> "" Then
            strSql = strSql & " AND UPPER(strContaContabil) LIKE " & UCase(IIf(Mid(cbointContaContabil.Text, 1, 1) = "%", "'%", "'") & Replace(cbointContaContabil.Text, ".", "") & "%'")
    End If
                       
    strSql = strSql & " ORDER BY strContaContabil"
    
    cbointContaContabil.Clear
    cbo_DescricaoExtra.Clear
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                cbointContaContabil.AddItem gvntFormatacaoEspecifica(!strContaContabil, 1)
                cbointContaContabil.ItemData(cbointContaContabil.NewIndex) = !Pkid
                cbo_DescricaoExtra.AddItem !strDescricao
                cbo_DescricaoExtra.ItemData(cbo_DescricaoExtra.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If

End Sub

Private Function strQueryLocalizar() As String
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT DE.PKId, DE.intNumero, PC.strContaContabil, "
    strSql = strSql & "PC.strDescricao, DE.bytSituacao, DE.dblValor, OP.intNumero OP,  "
    
    strSql = strSql & gstrCASEWHEN("DE.bytSituacao", "0, 'Programada', 2, 'Paga'") & " AS strSituacao "
    
    strSql = strSql & "FROM "
    strSql = strSql & gstrDespesaExtraOrcamentaria & " DE "
    strSql = strSql & " INNER JOIN " & gstrPlanoConta & " PC ON (PC.PKId = DE.intContaContabil) "
    strSql = strSql & " LEFT JOIN " & gstrOrdemPagamentoDespesaExtra & " OPDE ON (DE.PKID = OPDE.intDespesaExtraOrcamentaria)"
    strSql = strSql & " LEFT JOIN " & gstrOrdemPagamento & " OP ON (OPDE.intOrdemPagamento = OP.PKID) "
    strSql = strSql & " WHERE  DE.intExercicio = " & gintExercicio
    
    If Len(Trim$(txtintNumero)) > 0 Then
        strSql = strSql & " AND DE.intNumero = " & txtintNumero
    End If
    
    If Len(Trim$(txtdblValor)) > 0 Then
        strSql = strSql & " AND DE.dblValor = " & gstrConvVrParaSql(txtdblValor)
    End If
    
    If Len(Trim$(txtdtmData)) > 0 Then
        strSql = strSql & " AND DE.dtmData = " & gstrConvDtParaSql(txtdtmData)
    End If
    
    If Len(Trim$(txt_intNContribuinte)) > 0 Then
        strSql = strSql & " AND DE.intContribuinte = " & RetornaCredor(txt_intNContribuinte)
    End If
    
    If dbcintContribuinte.MatchedWithList Then
        strSql = strSql & " AND DE.intContribuinte = " & RetornaCredor(dbcintContribuinte.BoundText)
    End If
    
    If cbointContaContabil.ListIndex > -1 Then
        strSql = strSql & " AND DE.intContaContabil = " & cbointContaContabil.ItemData(cbointContaContabil.ListIndex)
    ElseIf cbo_DescricaoExtra.ListIndex > -1 Then
        strSql = strSql & " AND DE.intContaContabil = " & cbo_DescricaoExtra.ItemData(cbo_DescricaoExtra.ListIndex)
    End If
    
    If Len(Trim$(txtstrHistorico)) > 0 Then
        strSql = strSql & " AND UPPER(DE.strHistorico) LIKE '" & UCase$(txtstrHistorico) & "%'"
    End If
    
    strSql = strSql & "ORDER BY DE.intNumero"

    strQueryLocalizar = strSql

End Function
Public Sub SelecionaDespesaExtra()
       mblnClickOk = True
       DoEvents
       tdb_Lista_RowColChange 0, 0
End Sub

Private Sub ImprimeDocumentoExtra()
         
   Dim strSql As String

   strSql = "SELECT "
   strSql = strSql & "DEX.PKID, "
   strSql = strSql & "2 bytTipo, "
   strSql = strSql & "DEX.intNumero intOrdem, "
   strSql = strSql & "DEX.strHistorico typHistorico, "
   strSql = strSql & "DEX.dtmData, "
   strSql = strSql & "NULL dtmDataVencimento, "
   strSql = strSql & "DEX.intExercicio IntExercicioOP,"
   strSql = strSql & "PP.strCodigo, "
   strSql = strSql & "PP.intExercicio intExercicioProcesso, "
   strSql = strSql & "PP.bitDigito, "
   strSql = strSql & "CT.strNome, "
   strSql = strSql & "CT.CDC intContribuinte,"
   strSql = strSql & "CT.strCNPJCPF, "
   strSql = strSql & "CT.strLogradouroC strEndereco, "
   strSql = strSql & "CT.intNumero, "
   strSql = strSql & "CT.strComplemento strComplemento, "
   strSql = strSql & "CT.bytNaturezaJuridica,"
   strSql = strSql & "MP.strDescricao strMunicipio, "
   strSql = strSql & "UF.strSigla strUF, "
   strSql = strSql & "BR.strDescricao strBairro, "
   strSql = strSql & "CP.intCEP, "
   strSql = strSql & gstrConvVrParaSql(CDbl(txtdblValor) - CDbl(IIf(Trim(txtdblDesconto) = "", "0,00", txtdblDesconto))) & " dblLiquidoTotal,"
   strSql = strSql & gstrConvVrParaSql(CDbl(txtdblValor)) & " dblLiquidadoTotal , "
   strSql = strSql & gstrConvVrParaSql(IIf(Trim(txtdblDesconto) = "", "0,00", txtdblDesconto)) & "dblDesconto, "
   strSql = strSql & " DEX.PKID PKIDDespExtra "
   strSql = strSql & "FROM "
   strSql = strSql & gstrProtocolizacaoProcesso & " PP, "
   strSql = strSql & gstrDespesaExtraOrcamentaria & " DEX, "
   strSql = strSql & gstrContribuinte & " CT, "
   strSql = strSql & gstrCidade & " MP, "
   strSql = strSql & gstrUF & " UF, "
   strSql = strSql & gstrBairro & " BR, "
   strSql = strSql & gstrCeps & " CP "
   strSql = strSql & "WHERE "
   strSql = strSql & "PP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " DEX.intProtocolizacaoProcesso AND "
   strSql = strSql & "DEX.intContribuinte = CT.PKID AND "
   strSql = strSql & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
   strSql = strSql & "CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep AND "
   strSql = strSql & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
   strSql = strSql & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
   strSql = strSql & "DEX.PKID = " & txtPKId & " "

   ImprimeRelatorio rptDespesaExtra, strSql, "Nota de Despesa Extra - Orcamentaria"

End Sub
