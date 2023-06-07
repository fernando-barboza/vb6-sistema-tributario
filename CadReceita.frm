VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadReceita 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receitas"
   ClientHeight    =   6600
   ClientLeft      =   1965
   ClientTop       =   3030
   ClientWidth     =   7365
   HelpContextID   =   23
   Icon            =   "CadReceita.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6540
      Left            =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   15
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   11536
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Receita "
      TabPicture(0)   =   "CadReceita.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_Receita"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Receita"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Exercícios"
      TabPicture(1)   =   "CadReceita.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblExercicio"
      Tab(1).Control(1)=   "lblintClassificacaoDaReceita"
      Tab(1).Control(2)=   "tdb_ReceitasExercicio"
      Tab(1).Control(3)=   "dbc_intClassificacaoDaReceita"
      Tab(1).Control(4)=   "fra_PreçoPublico"
      Tab(1).Control(5)=   "fra_Lancamento"
      Tab(1).Control(6)=   "txtPKId"
      Tab(1).Control(7)=   "txt_PkidExercicios"
      Tab(1).Control(8)=   "txt_intExercicio"
      Tab(1).ControlCount=   9
      Begin VB.Frame fra_Receita 
         Height          =   2985
         Left            =   90
         TabIndex        =   36
         Top             =   510
         Width           =   7125
         Begin VB.CheckBox chkbytinscreveDa 
            Caption         =   "Inscreve em Dívida Ativa"
            Height          =   285
            Left            =   4830
            TabIndex        =   3
            Top             =   1500
            Width           =   2205
         End
         Begin VB.TextBox txtstrSigla 
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
            MaxLength       =   10
            TabIndex        =   1
            Top             =   975
            Width           =   1035
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
            MaxLength       =   100
            TabIndex        =   0
            Top             =   495
            Width           =   5385
         End
         Begin VB.ComboBox cbobytTipo 
            Height          =   315
            ItemData        =   "CadReceita.frx":107A
            Left            =   1305
            List            =   "CadReceita.frx":107C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1440
            Width           =   3435
         End
         Begin VB.OptionButton optbytTiporeceita 
            Caption         =   "Orçamentária"
            Height          =   315
            Index           =   0
            Left            =   1305
            TabIndex        =   5
            Top             =   2430
            Width           =   1245
         End
         Begin VB.OptionButton optbytTiporeceita 
            Caption         =   "Extra - Orçamentária"
            Height          =   315
            Index           =   1
            Left            =   2985
            TabIndex        =   6
            Top             =   2430
            Width           =   1755
         End
         Begin MSDataListLib.DataCombo dbcintDividaAtiva 
            Height          =   315
            Left            =   1305
            TabIndex        =   4
            Top             =   1920
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label lblstrSigla 
            AutoSize        =   -1  'True
            Caption         =   "Sigla"
            Height          =   195
            Left            =   885
            TabIndex        =   40
            Top             =   1035
            Width           =   345
         End
         Begin VB.Label lblstrDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   510
            TabIndex        =   39
            Top             =   585
            Width           =   720
         End
         Begin VB.Label lblbytTipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   915
            TabIndex        =   38
            Top             =   1560
            Width           =   315
         End
         Begin VB.Label lblCodigoDividaAtiva 
            AutoSize        =   -1  'True
            Caption         =   "Dívida Ativa"
            Height          =   195
            Left            =   345
            TabIndex        =   37
            Top             =   1995
            Width           =   885
         End
      End
      Begin VB.TextBox txt_intExercicio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   -73980
         MaxLength       =   9
         TabIndex        =   8
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox txt_PkidExercicios 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72090
         TabIndex        =   33
         Top             =   30
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtPKId 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72795
         TabIndex        =   32
         Top             =   45
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame fra_Lancamento 
         Caption         =   "Lançamento"
         Height          =   2475
         Left            =   -74880
         TabIndex        =   24
         Top             =   1785
         Width           =   7095
         Begin VB.TextBox txt_dblValor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1575
            MaxLength       =   9
            TabIndex        =   15
            Top             =   2040
            Width           =   1395
         End
         Begin VB.CheckBox chk_blnParcelar 
            Caption         =   "Não Parcelar"
            Height          =   195
            Left            =   5070
            TabIndex        =   16
            Top             =   2130
            Width           =   1275
         End
         Begin VB.CheckBox chk_blnECalculada 
            Caption         =   "É Calculada"
            Height          =   195
            Left            =   630
            TabIndex        =   11
            Top             =   375
            Width           =   1215
         End
         Begin VB.Frame fra_E_Calculada 
            Enabled         =   0   'False
            Height          =   1545
            Left            =   390
            TabIndex        =   25
            Top             =   360
            Width           =   6330
            Begin VB.CheckBox chk_blnUsaFaixaDeValor 
               Caption         =   "Usa Faixa de Valor"
               Height          =   195
               Left            =   465
               TabIndex        =   12
               Top             =   315
               Width           =   1695
            End
            Begin VB.CommandButton cmd_FormulaDeCalculo 
               Height          =   300
               Left            =   5400
               Picture         =   "CadReceita.frx":107E
               Style           =   1  'Graphical
               TabIndex        =   26
               TabStop         =   0   'False
               Tag             =   "590"
               ToolTipText     =   "Clique para cadastrar fórmula de cálculo"
               Top             =   1110
               Width           =   360
            End
            Begin MSDataListLib.DataCombo dbc_intFormulaDeCalculo 
               Height          =   315
               Left            =   1665
               TabIndex        =   14
               Top             =   1110
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Frame fra_UsaFaixaDeValor 
               Height          =   720
               Left            =   225
               TabIndex        =   27
               Top             =   300
               Width           =   5805
               Begin VB.CommandButton cmd_Faixa 
                  Height          =   300
                  Left            =   5175
                  Picture         =   "CadReceita.frx":119C
                  Style           =   1  'Graphical
                  TabIndex        =   28
                  TabStop         =   0   'False
                  Tag             =   "590"
                  ToolTipText     =   "Clique para cadastrar faixa"
                  Top             =   270
                  Width           =   360
               End
               Begin MSDataListLib.DataCombo dbc_intFaixaDeValor 
                  Height          =   315
                  Left            =   660
                  TabIndex        =   13
                  Top             =   270
                  Width           =   4515
                  _ExtentX        =   7964
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
               End
               Begin VB.Label lblintFaixaDeValor 
                  AutoSize        =   -1  'True
                  Caption         =   "Faixa"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   29
                  Top             =   360
                  Width           =   375
               End
            End
            Begin VB.Label lblintFormulaDeCalculo 
               AutoSize        =   -1  'True
               Caption         =   "Fórmula de Cálculo"
               Height          =   195
               Left            =   240
               TabIndex        =   30
               Top             =   1215
               Width           =   1350
            End
         End
         Begin VB.Label lbldblValor 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   1110
            TabIndex        =   31
            Top             =   2130
            Width           =   360
         End
      End
      Begin VB.Frame fra_PreçoPublico 
         Caption         =   "Preço Público"
         Height          =   825
         Left            =   -74790
         TabIndex        =   20
         Top             =   885
         Width           =   4635
         Begin VB.TextBox txt_dblPrecoPublico 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   675
            MaxLength       =   9
            TabIndex        =   9
            Top             =   375
            Width           =   1035
         End
         Begin VB.CommandButton cmd_FormaAtualizacao 
            Height          =   300
            Left            =   4125
            Picture         =   "CadReceita.frx":12BA
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Clique para cadastrar Indexadores"
            Top             =   375
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbc_intFormaAtualizacao 
            Height          =   315
            Left            =   2655
            TabIndex        =   10
            Top             =   375
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblPreçoGlobal 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   210
            TabIndex        =   23
            Top             =   420
            Width           =   360
         End
         Begin VB.Label lblMoeda 
            AutoSize        =   -1  'True
            Caption         =   "Indexador"
            Height          =   195
            Left            =   1875
            TabIndex        =   22
            Top             =   420
            Width           =   705
         End
      End
      Begin MSDataListLib.DataCombo dbc_intClassificacaoDaReceita 
         Height          =   315
         Left            =   -73005
         TabIndex        =   17
         Top             =   4350
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Receita 
         Height          =   2625
         Left            =   90
         TabIndex        =   18
         Top             =   3750
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   4630
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
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "strDescricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Sigla"
         Columns(2).DataField=   "strSigla"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Dívida Ativa"
         Columns(3).DataField=   "Divida"
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
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=6562"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6482"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1667"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1588"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=7223"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=7144"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=164,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid tdb_ReceitasExercicio 
         Height          =   1680
         Left            =   -74895
         TabIndex        =   19
         Top             =   4725
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   2963
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
         Columns(1).Caption=   "Exercício"
         Columns(1).DataField=   "Exercicio"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Classificação"
         Columns(2).DataField=   "Classificacao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Preço Público"
         Columns(3).DataField=   "PrecoPublico"
         Columns(3).NumberFormat=   "Standard"
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
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6482"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6403"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=3836"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=3757"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=164,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
      Begin VB.Label lblintClassificacaoDaReceita 
         AutoSize        =   -1  'True
         Caption         =   "Classificação da Receita"
         Height          =   195
         Left            =   -74895
         TabIndex        =   35
         Top             =   4455
         Width           =   1755
      End
      Begin VB.Label lblExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   -74745
         TabIndex        =   34
         Top             =   555
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmCadReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando                   As Boolean
Dim mobjAux                         As Object
Dim mblnClickOk                     As Boolean
Dim mblnSelecionou                  As Boolean
Dim mblnPrimeiraVez                 As Boolean
Dim strDescricaoAtual               As String
Dim strSiglaAtual                   As String
Dim strCodigo                       As String
Dim bytOrdenacao                    As Byte
Dim blnOrdenacaoAsc                 As Boolean
Dim bytOrdenacaoExer                As Byte
Dim blnOrdenacaoAscExer             As Boolean

Private Sub cbobytTipo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", cbobytTipo
End Sub

Private Sub chk_blnECalculada_Click()
    Select Case chk_blnECalculada.Value
        Case 0
            fra_E_Calculada.Enabled = False
            chk_blnUsaFaixaDeValor.Value = 0
            chk_blnUsaFaixaDeValor.Enabled = False
            fra_UsaFaixaDeValor.Enabled = False
            lblintFaixaDeValor.Enabled = False
            lblintFormulaDeCalculo.Enabled = False
            TrocaCorObjeto dbc_intFormulaDeCalculo, True
            TrocaCorObjeto dbc_intFaixaDeValor, True, True
            dbc_intFormulaDeCalculo.BoundText = ""
            TrocaCorObjeto txt_dblvalor, False, True
            TrocaCorObjeto cmd_Faixa, True
            TrocaCorObjeto cmd_FormulaDeCalculo, True
        Case 1
            chk_blnUsaFaixaDeValor.Enabled = True
            fra_E_Calculada.Enabled = True
            fra_UsaFaixaDeValor.Enabled = True
            lblintFaixaDeValor.Enabled = True
            lblintFormulaDeCalculo.Enabled = True
            TrocaCorObjeto dbc_intFormulaDeCalculo, False
            'TrocaCorObjeto dbc_intFaixaDeValor, False
            TrocaCorObjeto txt_dblvalor, True, True
            'TrocaCorObjeto cmd_Faixa, False
            TrocaCorObjeto cmd_FormulaDeCalculo, False
    End Select
End Sub


Private Sub chk_blnUsaFaixaDeValor_Click()
    Select Case chk_blnUsaFaixaDeValor.Value
        Case 0
            dbc_intFaixaDeValor.BoundText = ""
            TrocaCorObjeto dbc_intFaixaDeValor, True
        Case 1
            TrocaCorObjeto dbc_intFaixaDeValor, False
    End Select
End Sub

Private Sub chk_blnUsaFaixaDeValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_blnUsaFaixaDeValor
End Sub

Private Sub cmd_Faixa_Click()
    CarregaForm frmCadFaixaDeValores, dbc_intFaixaDeValor
End Sub

Private Sub cmd_FormaAtualizacao_Click()
    CarregaForm frmIndexadorEconomico, dbc_intFormaAtualizacao
End Sub

Private Sub cmd_FormulaDeCalculo_Click()
    CarregaForm frmCadFormulaDeCalculos, dbc_intFormulaDeCalculo, "SELECT PKId, strNome FROM " & gstrFormulaDeCalculo & " ORDER BY strNome"
End Sub

Private Sub dbc_intClassificacaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intClassificacaoDaReceita, Me, Area
End Sub

Private Sub dbc_intClassificacaoDaReceita_GotFocus()
    MarcaCampo dbc_intClassificacaoDaReceita
End Sub

Private Sub dbc_intClassificacaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intClassificacaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intClassificacaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intClassificacaoDaReceita
End Sub


Private Sub dbc_intFaixaDeValor_Click(Area As Integer)
    DropDownDataCombo dbc_intFaixaDeValor, Me, Area
End Sub

Private Sub dbc_intFaixaDeValor_GotFocus()
    MarcaCampo dbc_intFaixaDeValor
End Sub

Private Sub dbc_intFaixaDeValor_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intFaixaDeValor, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intFaixaDeValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intFaixaDeValor
End Sub

Private Sub dbc_intFormaAtualizacao_Click(Area As Integer)
    DropDownDataCombo dbc_intFormaAtualizacao, Me, Area
End Sub

Private Sub dbc_intFormaAtualizacao_GotFocus()
    MarcaCampo dbc_intFormaAtualizacao
End Sub

Private Sub dbc_intFormaAtualizacao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intFormaAtualizacao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intFormaAtualizacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intFormaAtualizacao
End Sub

Private Sub dbc_intFormulaDeCalculo_Click(Area As Integer)
    DropDownDataCombo dbc_intFormulaDeCalculo, Me, Area
End Sub

Private Sub dbc_intFormulaDeCalculo_GotFocus()
    MarcaCampo dbc_intFormulaDeCalculo
End Sub

Private Sub dbc_intFormulaDeCalculo_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intFormulaDeCalculo, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intFormulaDeCalculo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intFormulaDeCalculo
End Sub

Private Sub dbcintDividaAtiva_Click(Area As Integer)
    DropDownDataCombo dbcintDividaAtiva, Me, Area
End Sub

Private Sub dbcintDividaAtiva_GotFocus()
    MarcaCampo dbcintDividaAtiva
End Sub

Private Sub dbcintDividaAtiva_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintDividaAtiva, Me, , KeyCode, Shift
End Sub

Private Sub dbcintDividaAtiva_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintDividaAtiva
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 444
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
    dbc_intFaixaDeValor.Tag = strQueryDataComboFaixadeValor & ";strNomeDaFaixa"
'dbc_intClassificacaoDaReceita.Tag = strQueryPrevisaoDaReceita & ";CO.strCodigoOrcamentario"
    dbc_intFormulaDeCalculo.Tag = "SELECT PKId, strNome FROM " & gstrFormulaDeCalculo & " ORDER BY strNome " & ";strNome"
    TrocaCorObjeto dbc_intFaixaDeValor, True
    TrocaCorObjeto cmd_Faixa, True
    CarregaCboTipo
    chk_blnUsaFaixaDeValor.Enabled = False
    fra_UsaFaixaDeValor.Enabled = False
    lblintFaixaDeValor.Enabled = False
    lblintFormulaDeCalculo.Enabled = False
    TrocaCorObjeto dbc_intFormulaDeCalculo, True
    dbc_intFaixaDeValor.Enabled = False
    cmd_Faixa.Enabled = False
    cmd_FormulaDeCalculo.Enabled = False
    VerificaObjParaAplicar mobjAux
    tab_3dPasta.TabEnabled(1) = False
    dbc_intFormaAtualizacao.Tag = strQueryFormaAtualizacao & ";strAbreviatura"
    dbcintDividaAtiva.Tag = strQueryCodigoDividaAtiva & ";strDescricao"
    bytOrdenacao = 1: blnOrdenacaoAsc = True
    bytOrdenacaoExer = 1: blnOrdenacaoAsc = True
    optbytTiporeceita(0).Value = True
    
End Sub

Private Function strQueryDataComboFaixadeValor()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNomeDaFaixa "
    strSql = strSql & "FROM " & gstrFaixaDeValor & " "
    strSql = strSql & "ORDER BY strNomeDaFaixa"
    strQueryDataComboFaixadeValor = strSql
End Function


Private Sub CarregaCboTipo()
    
    cbobytTipo.AddItem "Convênio"
    cbobytTipo.ItemData(cbobytTipo.NewIndex) = "1"
    
    cbobytTipo.AddItem "Imposto"
    cbobytTipo.ItemData(cbobytTipo.NewIndex) = "2"
    
    cbobytTipo.AddItem "Taxa"
    cbobytTipo.ItemData(cbobytTipo.NewIndex) = "3"
    
    cbobytTipo.AddItem "Tarifa"
    cbobytTipo.ItemData(cbobytTipo.NewIndex) = "4"
    
    cbobytTipo.AddItem "Repasse Governamental"
    cbobytTipo.ItemData(cbobytTipo.NewIndex) = "5"
    
    cbobytTipo.AddItem "Repasse não Governamental"
    cbobytTipo.ItemData(cbobytTipo.NewIndex) = "6"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
End Sub

Private Sub optbytTiporeceita_Click(Index As Integer)
    If Index = 0 Then
        dbc_intClassificacaoDaReceita.Tag = strQueryPrevisaoDaReceita & ";CO.strCodigoOrcamentario"
    Else
        dbc_intClassificacaoDaReceita.Tag = strQueryPrevisaoDaReceita2 & ";strDescricao"
    End If
End Sub

Private Sub tdb_Receita_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Receita) = 1 Then
        tdb_Receita_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Receita_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Receita_FilterChange()
    gblnFilraCampos tdb_Receita
End Sub

Private Sub tdb_Receita_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Receita, ColIndex
End Sub

Private Sub tdb_Receita_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown
        mblnClickOk = True
    Case Else
        mblnClickOk = False
    End Select
End Sub

Private Sub tdb_Receita_KeyPress(KeyAscii As Integer)
    Select Case tdb_Receita.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_Receita
        Case Else
            CaracterValido KeyAscii, "A", tdb_Receita
    End Select
End Sub

Private Sub tdb_Receita_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Receita_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim intCont As Byte
    With tdb_Receita
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                If mblnClickOk Then
                    LimpaTabExercicio
                    mblnClickOk = False
                    mblnAlterando = True
                    txtPKId.Text = .Columns("PKID").Value
                    LeDaTabelaParaObj gstrReceita, Me
                    gCorLinhaSelecionada tdb_Receita
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                    If mobjAux Is Nothing Then
                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                    End If
                    mblnSelecionou = True
                    strDescricaoAtual = txtstrDescricao.Text
                    strSiglaAtual = txtstrSigla.Text
                    tab_3dPasta.TabEnabled(1) = True
                    PreencheGrdExercicios Val(txtPKId)
                    'For intCont = 0 To 1
                    '    TrocaCorObjeto optbytTiporeceita(intCont), True
                    'Next
                End If
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
Dim strSql As String
Dim intCont As Byte
    
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrSalvar)
            Dim lngPkid As Long
            
            mblnPrimeiraVez = False
            If Not blnDadosOk Then Exit Sub
            If tab_3dPasta.Tab = 0 Then

                If Not mblnAlterando Then
                    lngPkid = 0
                Else
                    lngPkid = Val(txtPKId)
                End If

                ToolBarGeral strModoOperacao, gstrReceita, mblnAlterando, tdb_Receita, Me, mobjAux, "", , rptReceita, strQueryRelatorio
                
                'Ja vamos posicionar no registro cadastrados, ou alterado
                If lngPkid = 0 Then
                    lngPkid = glngRetornaPkidTabelaPai("seqtblReceita", gstrReceita)
                End If

                LeDaTabelaParaObj gstrReceita, tdb_Receita, strQuery(lngPkid)
                mblnClickOk = True
                tdb_Receita_Click
                
                LimpaTabExercicio
                txtstrDescricao.SetFocus
                
            Else
                If gblnExclusaoGravacaoOk(strModoOperacao, "Confirma " & IIf(Val(txt_PkidExercicios.Text) = 0, "Inclusão", "Alteração")) Then
                    SalvaReceitasExercicios IIf(Val(txt_PkidExercicios.Text) = 0, False, True)
                End If
                PreencheGrdExercicios Val(txtPKId)
            End If
        
        Case Is = UCase(gstrDeletar)
            If tab_3dPasta.Tab = 0 Then
                mblnPrimeiraVez = False
                mblnAlterando = False
                ToolBarGeral strModoOperacao, gstrReceita, mblnAlterando, tdb_Receita, Me, mobjAux, strQuery
                LimpaTabExercicio
                txtstrDescricao.SetFocus
            Else
                If gblnExclusaoGravacaoOk(strModoOperacao, "Confirma Exclusão") Then
                    Set gobjBanco = New clsBanco
                    strSql = "DELETE FROM " & gstrReceitasExercicio
                    strSql = strSql & " WHERE"
                    strSql = strSql & " Pkid = " & Val(txt_PkidExercicios.Text)
                    gobjBanco.Execute strSql
                    PreencheGrdExercicios Val(txtPKId)
                End If
            End If
        Case Is = UCase(gstrNovo)
            If tab_3dPasta.Tab = 0 Then
                mblnAlterando = False
                mblnSelecionou = False
                mblnPrimeiraVez = False
                mblnClickOk = False
                LimpaObjeto Me
                tab_3dPasta.TabEnabled(1) = False
                LimpaTabExercicio
                'For intCont = 0 To 1
                '    TrocaCorObjeto optbytTiporeceita(intCont), False
                'Next
                txtstrDescricao.SetFocus
            Else
                LimpaTabExercicio
            End If
        Case Else
            If UCase(strModoOperacao) = gstrPreencherLista And UCase(Me.ActiveControl.Name) = "DBC_INTCLASSIFICACAODARECEITA" Then
                If Len(Trim(txt_intExercicio)) = 4 Then
                    If optbytTiporeceita(0).Value Then
                        dbc_intClassificacaoDaReceita.Tag = strQueryPrevisaoDaReceita & ";CO.strCodigoOrcamentario"
                    Else
                        dbc_intClassificacaoDaReceita.Tag = strQueryPrevisaoDaReceita2 & ";strDescricao"
                    End If
                    ToolBarGeral strModoOperacao, gstrReceita, mblnAlterando, tdb_Receita, Me, mobjAux, strQuery, , rptReceita, strQueryRelatorio
                    Exit Sub
                Else
                    ExibeMensagem "É necessário o preenchimento do exercício"
                    If txt_intExercicio.Enabled Then txt_intExercicio.SetFocus
                    Exit Sub
                End If
            End If
            ToolBarGeral strModoOperacao, gstrReceita, mblnAlterando, tdb_Receita, Me, mobjAux, strQuery, , rptReceita, strQueryRelatorio
    End Select
    
End Sub

Private Sub tdb_ReceitasExercicio_FilterChange()
    gblnFilraCampos tdb_ReceitasExercicio
End Sub

Private Sub tdb_ReceitasExercicio_HeadClick(ByVal ColIndex As Integer)
    blnOrdenacaoAscExer = IIf(bytOrdenacaoExer = ColIndex, Not blnOrdenacaoAscExer, True)
    bytOrdenacaoExer = ColIndex: PreencheGrdExercicios Val(txtPKId)
End Sub

Private Sub tdb_ReceitasExercicio_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tdb_ReceitasExercicio.EOF Then
        txt_PkidExercicios = tdb_ReceitasExercicio.Columns("Pkid").Value
        PreencheTabExercicios Val(tdb_ReceitasExercicio.Columns("Pkid").Value)
        If gstrConvVrDoSql(txt_dblPrecoPublico, , , True) <= 0 Then
            TrocaCorObjeto dbc_intFormaAtualizacao, True
            cmd_FormaAtualizacao.Enabled = False
        Else
            TrocaCorObjeto dbc_intFormaAtualizacao, False
            cmd_FormaAtualizacao.Enabled = True
        End If
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    End If
End Sub
Private Sub txt_dblPrecoPublico_GotFocus()
    MarcaCampo txt_dblPrecoPublico
End Sub

Private Sub txt_dblPrecoPublico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblPrecoPublico
End Sub

Private Sub txt_dblPrecoPublico_LostFocus()
    txt_dblPrecoPublico = gstrConvVrDoSql(txt_dblPrecoPublico, 2)
    If gstrConvVrDoSql(txt_dblPrecoPublico, , , True) <= 0 Then
        TrocaCorObjeto dbc_intFormaAtualizacao, True
        cmd_FormaAtualizacao.Enabled = False
        If chk_blnECalculada.Enabled = True Then chk_blnECalculada.SetFocus
    Else
        TrocaCorObjeto dbc_intFormaAtualizacao, False
        cmd_FormaAtualizacao.Enabled = True
    End If
End Sub

Private Sub txt_intExercicio_GotFocus()
    tab_3dPasta.Tab = 1
    If txt_intExercicio.Text = "" Then txt_intExercicio = gintExercicio
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txtstrSigla_GotFocus()
    MarcaCampo txtstrSigla
End Sub

Private Sub txtstrSigla_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSigla
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Function strQueryRelatorio() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT strDescricao, strSigla FROM " & gstrReceita
    If mblnAlterando = True Then
        strSql = strSql & " WHERE PKId = " & Val(txtPKId)
    End If
    strSql = strSql & " ORDER BY strDescricao"
    strQueryRelatorio = strSql
End Function

Function strQueryPrevisaoDaReceita() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT PR.PKId, RTRIM(CO.strCodigoOrcamentario) " & strCONCAT & "' - '" & strCONCAT & " CO.strDescricao AS Previsao "
    strSql = strSql & "FROM " & gstrPrevisaoDaReceita & " PR , "
    strSql = strSql & gstrCodigoOrcamentario & " CO "
    strSql = strSql & "WHERE PR.intCodigoOrcamentario = CO.PKId AND "
    strSql = strSql & "PR.Intexercicio = " & Trim(txt_intExercicio) & " "
    strSql = strSql & "ORDER BY CO.strCodigoOrcamentario"
    
    strQueryPrevisaoDaReceita = strSql
    
End Function

Function strQueryPrevisaoDaReceita2() As String
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "Pkid, "
    strSql = strSql & "strDescricao "
    strSql = strSql & "From "
    strSql = strSql & gstrPlanoConta & " "
    strSql = strSql & "Where "
    strSql = strSql & "blnanalitica = 1 AND "
    strSql = strSql & "blnpatrimonial = 1 AND "
    strSql = strSql & "Blnextraorcamentaria = 1 "
    strSql = strSql & "Order By "
    strSql = strSql & "strDescricao "
    
    strQueryPrevisaoDaReceita2 = strSql
    
End Function



Private Sub txt_dblValor_GotFocus()
    MarcaCampo txt_dblvalor
End Sub

Private Sub txt_DBLVALOR_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblvalor
End Sub

Private Sub txt_DBLVALOR_LostFocus()
    txt_dblvalor = gvntConvVrDoSql(txt_dblvalor)
End Sub


Private Function strQuery(Optional lngPkid As Long) As String
Dim strSql  As String
    
    strSql = "SELECT R.PKId,"
    strSql = strSql & " R.strdescricao, "
    strSql = strSql & " R.Byttiporeceita, "
    strSql = strSql & " R.strsigla, "
    strSql = strSql & " RR.Strdescricao as Divida"
    strSql = strSql & " FROM "
    strSql = strSql & gstrReceita & " R, "
    strSql = strSql & gstrReceita & " RR "
    strSql = strSql & "WHERE "
    strSql = strSql & "R.intDividaAtiva" & strOUTJSQLServer & "= RR.Pkid" & strOUTJOracle
    
    If lngPkid > 0 Then
        strSql = strSql & " AND R.Pkid = " & lngPkid
    End If
    
    Select Case bytOrdenacao
        Case Is = 1
            strSql = strSql & " ORDER BY R.strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strSql = strSql & " ORDER BY R.strSigla" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
            
    End Select
    
    strQuery = strSql
    
End Function

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    
    Select Case tab_3dPasta.Tab
        Case Is = 0
            If txtstrDescricao.Text = "" Then
                ExibeMensagem "A Descrição deve ser informada."
                txtstrDescricao.SetFocus
                Exit Function
            ElseIf txtstrSigla.Text = "" Then
                ExibeMensagem "A Sigla deve ser informada."
                txtstrSigla.SetFocus
                Exit Function
            ElseIf cbobytTipo.ListIndex = -1 Then
                ExibeMensagem "O Tipo deve ser informado."
                cbobytTipo.SetFocus
                Exit Function
            End If
            If dbcintDividaAtiva.Text <> "" Then
                If Not dbcintDividaAtiva.MatchedWithList Then
                    ExibeMensagem "Selecione um Código de Dívida Ativa válido."
                    If dbcintDividaAtiva.Enabled Then dbcintDividaAtiva.SetFocus
                    Exit Function
                End If
            End If
                
            If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrDescricao.Text) <> UCase$(strDescricaoAtual)) Then
                    
                If gblnExisteCodigo(1, gstrReceita, "strDescricao", "'" & txtstrDescricao.Text & "'") Then
                    ExibeMensagem "A descrição informada já se encontra cadastrada."
                    txtstrDescricao.SetFocus
                    Exit Function
                End If
            End If
            
            If Not mblnAlterando Or (mblnAlterando And UCase$(txtstrSigla.Text) <> UCase$(strSiglaAtual)) Then
                    
                If gblnExisteCodigo(1, gstrReceita, "strSigla", "'" & txtstrSigla.Text & "'") Then
                    ExibeMensagem "A Sigla informada já se encontra cadastrada."
                    txtstrSigla.SetFocus
                    Exit Function
                End If
            End If
        Case Is = 1
            If Val(txtPKId) = 0 Then
                ExibeMensagem "Selecione uma Receita válida."
                tab_3dPasta.Tab = 0
            End If
            
            If Val(txt_intExercicio.Text) = 0 Then
                ExibeMensagem "É necessário informar o Exercício."
                If txt_intExercicio.Enabled Then txt_intExercicio.SetFocus
                Exit Function
            End If
            
            If dbc_intFormaAtualizacao.Text <> "" Then
                If Not dbc_intFormaAtualizacao.MatchedWithList Then
                    ExibeMensagem "Selecione uma Moeda Válida."
                    If dbc_intFormaAtualizacao.Enabled Then dbc_intFormaAtualizacao.SetFocus
                    Exit Function
                End If
            End If
            
            If chk_blnECalculada.Value = 1 Then
                If chk_blnUsaFaixaDeValor.Value = 1 Then
                    If Not dbc_intFaixaDeValor.MatchedWithList Then
                        ExibeMensagem "Selecione uma Faixa de Valor válida."
                        If dbc_intFaixaDeValor.Enabled Then dbc_intFaixaDeValor.SetFocus
                        Exit Function
                    End If
                End If
                If Not dbc_intFormulaDeCalculo.MatchedWithList Then
                    ExibeMensagem "Selecione uma Fórmula de Cálculo válida."
                    If dbc_intFormulaDeCalculo.Enabled Then dbc_intFormulaDeCalculo.SetFocus
                    Exit Function
                End If
            End If
                            
            If chk_blnECalculada.Value = 0 Then
                If txt_dblvalor.Text = "" Then
                    ExibeMensagem "Informe algum valor."
                    If txt_dblvalor.Enabled Then txt_dblvalor.SetFocus
                    Exit Function
                End If
            End If
            If Not dbc_intClassificacaoDaReceita.MatchedWithList Then
                ExibeMensagem "O campo Classificação da Receita deve ser preenchido."
                If dbc_intClassificacaoDaReceita.Enabled Then dbc_intClassificacaoDaReceita.SetFocus
                Exit Function
            End If
                
    End Select
    
    blnDadosOk = True
    
End Function

Private Sub SalvaReceitasExercicios(blnAlterando As Boolean)
Dim strSql As String

    Set gobjBanco = New clsBanco

    If blnAlterando Then
        
        strSql = "UPDATE " & gstrReceitasExercicio & " SET "
        strSql = strSql & " intExercicio = " & Val(txt_intExercicio.Text) & ", "
        strSql = strSql & " dblPrecoPublico = " & gstrConvVrParaSql(txt_dblPrecoPublico.Text) & ", "
        strSql = strSql & " INTINDEXADORECONOMICO = " & gstrENulo(dbc_intFormaAtualizacao.BoundText, , True) & ", "
        strSql = strSql & " blnECalculada = " & chk_blnECalculada.Value & ", "
        strSql = strSql & " blnParcelar = " & chk_blnParcelar.Value & ", "
        strSql = strSql & " dblValor = " & gstrConvVrParaSql(txt_dblvalor.Text) & ", "
        strSql = strSql & " blnUsaFaixaDeValor = " & chk_blnUsaFaixaDeValor.Value & ", "
        strSql = strSql & " intFaixaDeValor = " & gstrENulo(dbc_intFaixaDeValor.BoundText, , True) & ", "
        strSql = strSql & " intFormulaDeCalculo = " & gstrENulo(dbc_intFormulaDeCalculo.BoundText, , True) & ", "
        If optbytTiporeceita(0).Value = True Then
            strSql = strSql & " intClassificacaoDaReceita = " & gstrENulo(dbc_intClassificacaoDaReceita.BoundText, , True) & ", "
        Else
            strSql = strSql & " intplanoconta = " & gstrENulo(dbc_intClassificacaoDaReceita.BoundText, , True) & ", "
        End If
        strSql = strSql & " lngCodUsr = " & glngCodUsr & ", "
        strSql = strSql & " dtmDtAtualizacao = " & strGETDATE
        strSql = strSql & " WHERE Pkid = " & Val(txt_PkidExercicios.Text)
    Else
        strSql = "INSERT INTO " & gstrReceitasExercicio & "("
        strSql = strSql & "intReceita,"
        strSql = strSql & " intExercicio,"
        strSql = strSql & " dblPrecoPublico,"
        strSql = strSql & " INTINDEXADORECONOMICO,"
        strSql = strSql & " blnECalculada,"
        strSql = strSql & " blnParcelar,"
        strSql = strSql & " dblValor,"
        strSql = strSql & " blnUsaFaixaDeVAlor,"
        strSql = strSql & " intFaixaDeValor,"
        strSql = strSql & " intFormulaDeCalculo,"
        If optbytTiporeceita(0).Value = True Then
            strSql = strSql & " intClassificacaoDaReceita,"
        Else
            strSql = strSql & " intplanoconta,"
        End If
        strSql = strSql & " lngCodUsr,"
        strSql = strSql & " dtmDtAtualizacao)"
        strSql = strSql & " VALUES("
        strSql = strSql & Val(txtPKId.Text) & ", "
        strSql = strSql & Val(txt_intExercicio.Text) & ", "
        strSql = strSql & gstrConvVrParaSql(txt_dblPrecoPublico.Text) & ", "
        strSql = strSql & gstrENulo(dbc_intFormaAtualizacao.BoundText, , True) & ", "
        strSql = strSql & chk_blnECalculada.Value & ", "
        strSql = strSql & chk_blnParcelar.Value & ", "
        strSql = strSql & gstrConvVrParaSql(txt_dblvalor.Text) & ", "
        strSql = strSql & chk_blnUsaFaixaDeValor.Value & ", "
        strSql = strSql & gstrENulo(dbc_intFaixaDeValor.BoundText, , True) & ", "
        strSql = strSql & gstrENulo(dbc_intFormulaDeCalculo.BoundText, , True) & ", "
        strSql = strSql & gstrENulo(dbc_intClassificacaoDaReceita.BoundText, , True) & ", "
        strSql = strSql & glngCodUsr & ", "
        strSql = strSql & strGETDATE & ")"
    End If
    
    gobjBanco.Execute strSql

End Sub

Private Sub LimpaTabExercicio()

    txt_PkidExercicios.Text = ""
    txt_intExercicio.Text = ""
    chk_blnECalculada.Value = 0
    chk_blnUsaFaixaDeValor.Value = 0
    dbc_intFormaAtualizacao.ListField = ""
    dbc_intFormaAtualizacao.Text = ""
    TrocaCorObjeto dbc_intFaixaDeValor, True, True
    TrocaCorObjeto cmd_Faixa, True
    TrocaCorObjeto dbc_intFormulaDeCalculo, True, True
    TrocaCorObjeto cmd_FormulaDeCalculo, True
    TrocaCorObjeto txt_dblvalor, False, True
    txt_dblPrecoPublico = ""
    txt_dblvalor = ""
    chk_blnParcelar.Value = 0
    TrocaCorObjeto dbc_intClassificacaoDaReceita, False, True
    dbc_intClassificacaoDaReceita.ListField = ""
    dbc_intClassificacaoDaReceita.Text = ""
    
    
End Sub

Private Sub PreencheGrdExercicios(lngPkidReceitas As Long)
Dim strSql As String

    strSql = "SELECT RE.Pkid,"
    strSql = strSql & " RE.intExercicio Exercicio, "
    
    If optbytTiporeceita(0).Value = True Then
        strSql = strSql & " RTRIM(CO.strCodigoOrcamentario) " & strCONCAT & "' - '" & strCONCAT
        strSql = strSql & " CO.strDescricao AS Classificacao,"
    Else
        strSql = strSql & "PC.strDescricao AS Classificacao,"
    End If
    
    strSql = strSql & " RE.dblPrecoPublico PrecoPublico"
    strSql = strSql & " FROM " & gstrReceitasExercicio & " RE, "
    
    If optbytTiporeceita(0).Value = True Then
        strSql = strSql & gstrCodigoOrcamentario & " CO, "
        strSql = strSql & gstrPrevisaoDaReceita & " PR "
        strSql = strSql & " WHERE "
        strSql = strSql & "PR.intCodigoOrcamentario " & strOUTJSQLServer & "= CO.Pkid " & strOUTJOracle & " And "
        strSql = strSql & " PR.Pkid " & strOUTJOracle & "= RE.intClassificacaoDaReceita AND"
    Else
        strSql = strSql & gstrPlanoConta & " PC "
        strSql = strSql & " WHERE"
        strSql = strSql & " PC.Pkid = RE.intplanoConta AND"
    End If
    
    strSql = strSql & " RE.intReceita = " & Val(lngPkidReceitas)
    
    Select Case bytOrdenacaoExer
        Case Is = 1
            strSql = strSql & " ORDER BY Exercicio" & IIf(blnOrdenacaoAscExer, " ASC", " DESC")
        Case Is = 2
            strSql = strSql & " ORDER BY Classificacao" & IIf(blnOrdenacaoAscExer, " ASC", " DESC")
        Case Is = 3
            strSql = strSql & " ORDER BY PrecoPublico" & IIf(blnOrdenacaoAscExer, " ASC", " DESC")
    End Select
    
    LeDaTabelaParaObj "", tdb_ReceitasExercicio, strSql

End Sub

Private Sub PreencheTabExercicios(lngPkidExercicio As Long)
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset

    strSql = "SELECT RE.intExercicio,"
    strSql = strSql & " RE.blnECalculada,"
    strSql = strSql & " RE.blnParcelar,"
    strSql = strSql & " RE.dblValor,"
    strSql = strSql & " RE.blnUsaFaixaDeValor,"
    strSql = strSql & " RE.intFaixaDeValor,"
    strSql = strSql & " RE.intFormulaDeCalculo,"
    strSql = strSql & " RE.intClassificacaoDaReceita,"
    strSql = strSql & " RE.dblPrecoPublico,"
    strSql = strSql & " RE.INTINDEXADORECONOMICO,"
    strSql = strSql & " RE.intPlanoConta"
    strSql = strSql & " FROM "
    strSql = strSql & gstrReceitasExercicio & " RE "
'    strSql = strSql & gstrFaixaDeValor & " FV, "
'    strSql = strSql & gstrFormulaDeCalculo & " FC "
    strSql = strSql & " WHERE RE.Pkid = " & lngPkidExercicio

    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_intExercicio = gstrENulo(adoResultado!intExercicio)
            chk_blnECalculada.Value = gstrENulo(adoResultado!blnECalculada)
            chk_blnUsaFaixaDeValor.Value = gstrENulo(adoResultado!blnUsaFaixaDeValor)
            PreencherListaDeOpcoes dbc_intFaixaDeValor, gstrENulo(adoResultado!intFaixaDeValor)
            PreencherListaDeOpcoes dbc_intFormulaDeCalculo, gstrENulo(adoResultado!intFormulaDeCalculo)
            txt_dblvalor = gstrConvVrDoSql(gstrENulo(adoResultado!dblValor), 2)
            chk_blnParcelar = gstrENulo(adoResultado!blnParcelar)
            If optbytTiporeceita(0).Value = True Then
                dbc_intClassificacaoDaReceita.Tag = strQueryPrevisaoDaReceita & ";CO.strCodigoOrcamentario"
                PreencherListaDeOpcoes dbc_intClassificacaoDaReceita, gstrENulo(adoResultado!intClassificacaoDaReceita)
            Else
                dbc_intClassificacaoDaReceita.Tag = strQueryPrevisaoDaReceita2 & ";strDescricao"
                PreencherListaDeOpcoes dbc_intClassificacaoDaReceita, gstrENulo(adoResultado!intPlanoConta)
            End If
            txt_dblPrecoPublico = gstrConvVrDoSql(gstrENulo(adoResultado!dblPrecoPublico), 2)
            PreencherListaDeOpcoes dbc_intFormaAtualizacao, gstrENulo(adoResultado!INTINDEXADORECONOMICO)
        End If
    End If
        

End Sub

Private Function strQueryFormaAtualizacao() As String
Dim strSql As String

    strSql = "SELECT Pkid,"
    strSql = strSql & " strAbreviatura"
    strSql = strSql & " FROM "
    strSql = strSql & gstrIndexadorEconomico
    strSql = strSql & " ORDER BY strAbreviatura"
    
    strQueryFormaAtualizacao = strSql

End Function

Private Function strQueryCodigoDividaAtiva() As String
Dim strSql As String

    strSql = "SELECT Pkid,"
    strSql = strSql & " strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrReceita
    strSql = strSql & " ORDER BY strDescricao"
    
    strQueryCodigoDividaAtiva = strSql

End Function
