VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadDebito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamentos em Conta Corrente"
   ClientHeight    =   7095
   ClientLeft      =   1050
   ClientTop       =   2310
   ClientWidth     =   9390
   HelpContextID   =   636
   Icon            =   "CadDebito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_PKId 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3180
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   7065
      Left            =   60
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   30
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   12462
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lançamentos em Conta Corrente"
      TabPicture(0)   =   "CadDebito.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Cadastro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Pesquisa"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tab_Receita"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame Frame1 
         Height          =   1935
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   8985
         Begin VB.CommandButton cmd_intComposicaoReceita 
            Height          =   315
            Left            =   8145
            Picture         =   "CadDebito.frx":105E
            Style           =   1  'Graphical
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "Ativa o Cadastro de Composição da Receita"
            Top             =   240
            Width           =   330
         End
         Begin VB.TextBox txt_strSequencia 
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
            Left            =   6045
            MaxLength       =   5
            TabIndex        =   9
            Top             =   1245
            Width           =   555
         End
         Begin VB.CommandButton cmd_intOcorrencia 
            Height          =   315
            Left            =   8145
            Picture         =   "CadDebito.frx":117C
            Style           =   1  'Graphical
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Ativa o Cadastro de Ocorrências"
            Top             =   900
            Width           =   330
         End
         Begin VB.CommandButton cmd_intReceitas 
            Height          =   315
            Left            =   8145
            Picture         =   "CadDebito.frx":129A
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Ativa o Cadastro de Receitas"
            Top             =   570
            Width           =   330
         End
         Begin VB.TextBox txt_dtmDataLancamento 
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
            Left            =   4095
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1545
            Width           =   975
         End
         Begin VB.TextBox txt_dtmDataVencimento 
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
            MaxLength       =   15
            TabIndex        =   10
            Top             =   1545
            Width           =   975
         End
         Begin VB.TextBox txt_intNumeroParcela 
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
            Left            =   4095
            MaxLength       =   5
            TabIndex        =   8
            Top             =   1245
            Width           =   555
         End
         Begin VB.TextBox txt_intExercicio 
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
            MaxLength       =   4
            TabIndex        =   7
            Top             =   1245
            Width           =   555
         End
         Begin VB.TextBox txt_dblValorParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6045
            TabIndex        =   12
            Top             =   1545
            Width           =   1350
         End
         Begin MSDataListLib.DataCombo dbc_intComposicaoReceita 
            Height          =   315
            Left            =   1605
            TabIndex        =   4
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intOcorrencia 
            Height          =   315
            Left            =   1605
            TabIndex        =   6
            Top             =   900
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intReceitas 
            Height          =   315
            Left            =   1605
            TabIndex        =   5
            Top             =   570
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_strSequencia 
            AutoSize        =   -1  'True
            Caption         =   "Sequencia"
            Height          =   195
            Left            =   5235
            TabIndex        =   39
            Top             =   1335
            Width           =   765
         End
         Begin VB.Label lbl_dtmDataLancamento 
            AutoSize        =   -1  'True
            Caption         =   "Data Lançamento"
            Height          =   195
            Left            =   2730
            TabIndex        =   38
            Top             =   1635
            Width           =   1275
         End
         Begin VB.Label lbl_dtmDataVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento"
            Height          =   195
            Left            =   660
            TabIndex        =   37
            Top             =   1635
            Width           =   840
         End
         Begin VB.Label lbl_intNumeroParcela 
            AutoSize        =   -1  'True
            Caption         =   "Número Parcela"
            Height          =   195
            Left            =   2865
            TabIndex        =   36
            Top             =   1335
            Width           =   1140
         End
         Begin VB.Label lbl_intComposicaoReceita 
            AutoSize        =   -1  'True
            Caption         =   "Origem da Receita"
            Height          =   195
            Left            =   180
            TabIndex        =   35
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label lbl_dblValorParcela 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   5640
            TabIndex        =   34
            Top             =   1635
            Width           =   360
         End
         Begin VB.Label lbl_intOcorrencia 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrência"
            Height          =   195
            Left            =   720
            TabIndex        =   33
            Top             =   1020
            Width           =   780
         End
         Begin VB.Label lbl_intExercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   825
            TabIndex        =   32
            Top             =   1335
            Width           =   675
         End
         Begin VB.Label lbl_intReceitas 
            AutoSize        =   -1  'True
            Caption         =   "Receitas"
            Height          =   195
            Left            =   870
            TabIndex        =   31
            Top             =   720
            Width           =   630
         End
      End
      Begin TabDlg.SSTab tab_Receita 
         Height          =   3240
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   5715
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Dívidas"
         TabPicture(0)   =   "CadDebito.frx":13B8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "tdb_Divida"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Composição da Receita"
         TabPicture(1)   =   "CadDebito.frx":13D4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "tdb_Composicao"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Receitas"
         TabPicture(2)   =   "CadDebito.frx":13F0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "tdb_Parcela"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Mensagens"
         TabPicture(3)   =   "CadDebito.frx":140C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txt_Mensagem2"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "txt_Mensagem1"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).ControlCount=   2
         Begin VB.TextBox txt_Mensagem1 
            Height          =   1335
            Left            =   -74880
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   480
            Width           =   8355
         End
         Begin VB.TextBox txt_Mensagem2 
            Height          =   1335
            Left            =   -74880
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   1860
            Width           =   8355
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Parcela 
            Height          =   2715
            Left            =   -74880
            TabIndex        =   15
            Top             =   480
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   4789
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
            Columns(2).Caption=   "Exercício"
            Columns(2).DataField=   "intExercicio"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Parcela"
            Columns(3).DataField=   "intNumeroParcela"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Sequência"
            Columns(4).DataField=   "strSequencia"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Vencimento"
            Columns(5).DataField=   "dtmDataVencimento"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Valor"
            Columns(6).DataField=   "dblValorParcela"
            Columns(6).NumberFormat=   "Standard"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "bitUtilizacaoDebito"
            Columns(7).DataField=   "bitUtilizacaoDebito"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   " intOcorrencia"
            Columns(8).DataField=   "intOcorrencia"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=6138"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6059"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1376"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1296"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1138"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1058"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=1561"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1482"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(28)=   "Column(5).Width=1958"
            Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=1879"
            Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(33)=   "Column(6).Width=2858"
            Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=2778"
            Splits(0)._ColumnProps(36)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(39)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(42)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(43)=   "Column(7).Visible=0"
            Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(45)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(46)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(48)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(49)=   "Column(8).Visible=0"
            Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
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
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            InsertMode      =   0   'False
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=39"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
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
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Composicao 
            Height          =   2715
            Left            =   -74880
            TabIndex        =   14
            Top             =   480
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   4789
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
            Columns(1).Caption=   "Parcela"
            Columns(1).DataField=   "intNumeroParcela"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Seq."
            Columns(2).DataField=   "strSequencia"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Vencimento"
            Columns(3).DataField=   "dtmDataVencimento"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Sit."
            Columns(4).DataField=   "Situacao"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Pagamento"
            Columns(5).DataField=   "dtmDataPagamento"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Valor Orig."
            Columns(6).DataField=   "dblValorParcela"
            Columns(6).NumberFormat=   "Standard"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Juros"
            Columns(7).DataField=   "dblJuros"
            Columns(7).NumberFormat=   "Standard"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Multa"
            Columns(8).DataField=   "dblMulta"
            Columns(8).NumberFormat=   "Standard"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Valor Total"
            Columns(9).DataField=   "dblTotalPago"
            Columns(9).NumberFormat=   "Standard"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "dtmLancamento"
            Columns(10).DataField=   "dtmLancamento"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "strInscricaoCadastral"
            Columns(11).DataField=   "strInscricaoCadastral"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "strDescricao"
            Columns(12).DataField=   "strDescricao"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   13
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=13"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1138"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1058"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=794"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=714"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1958"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1879"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(24)=   "Column(4).Width=661"
            Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=582"
            Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(30)=   "Column(5).Width=1958"
            Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=1879"
            Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(35)=   "Column(6).Width=2117"
            Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2037"
            Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(41)=   "Column(7).Width=2117"
            Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2037"
            Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=2"
            Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(47)=   "Column(8).Width=2117"
            Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2037"
            Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._ColStyle=2"
            Splits(0)._ColumnProps(52)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(53)=   "Column(9).Width=2117"
            Splits(0)._ColumnProps(54)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(55)=   "Column(9)._WidthInPix=2037"
            Splits(0)._ColumnProps(56)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._ColStyle=2"
            Splits(0)._ColumnProps(58)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(59)=   "Column(10).Width=1852"
            Splits(0)._ColumnProps(60)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(61)=   "Column(10)._WidthInPix=1773"
            Splits(0)._ColumnProps(62)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(63)=   "Column(10).AllowSizing=0"
            Splits(0)._ColumnProps(64)=   "Column(10).Visible=0"
            Splits(0)._ColumnProps(65)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(66)=   "Column(11).Width=847"
            Splits(0)._ColumnProps(67)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(68)=   "Column(11)._WidthInPix=767"
            Splits(0)._ColumnProps(69)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(70)=   "Column(11).AllowSizing=0"
            Splits(0)._ColumnProps(71)=   "Column(11).Visible=0"
            Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(73)=   "Column(12).Width=2725"
            Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=2646"
            Splits(0)._ColumnProps(76)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(77)=   "Column(12).AllowSizing=0"
            Splits(0)._ColumnProps(78)=   "Column(12).Visible=0"
            Splits(0)._ColumnProps(79)=   "Column(12).Order=13"
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
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            InsertMode      =   0   'False
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=39"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=82,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=79,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=80,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=81,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=86,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=83,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=84,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=85,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=28,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=25,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=26,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=27,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=70,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=90,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=87,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=88,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=89,.parent=17"
            _StyleDefs(88)  =   "Named:id=33:Normal"
            _StyleDefs(89)  =   ":id=33,.parent=0"
            _StyleDefs(90)  =   "Named:id=34:Heading"
            _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(92)  =   ":id=34,.wraptext=-1"
            _StyleDefs(93)  =   "Named:id=35:Footing"
            _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(95)  =   "Named:id=36:Selected"
            _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=37:Caption"
            _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(99)  =   "Named:id=38:HighlightRow"
            _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(101) =   "Named:id=39:EvenRow"
            _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(103) =   "Named:id=40:OddRow"
            _StyleDefs(104) =   ":id=40,.parent=33"
            _StyleDefs(105) =   "Named:id=41:RecordSelector"
            _StyleDefs(106) =   ":id=41,.parent=34"
            _StyleDefs(107) =   "Named:id=42:FilterBar"
            _StyleDefs(108) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Divida 
            Height          =   2715
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   4789
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
            Columns(1).Caption=   "Utilização"
            Columns(1).DataField=   "bitUtilizacao"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Código"
            Columns(2).DataField=   "intCodigoOrigem"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Inscrição"
            Columns(3).DataField=   "strInscricao"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Receita"
            Columns(4).DataField=   "strDescricao"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Exercício"
            Columns(5).DataField=   "intExercicio"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Valor"
            Columns(6).DataField=   "dblValor"
            Columns(6).NumberFormat=   "Standard"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "PKIdContribuinte"
            Columns(7).DataField=   "PKIdContribuinte"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1826"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1746"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=2487"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2408"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=4498"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=4419"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(28)=   "Column(5).Width=1693"
            Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=1614"
            Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=2328"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2249"
            Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(38)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(40)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(43)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(44)=   "Column(7).AllowSizing=0"
            Splits(0)._ColumnProps(45)=   "Column(7).Visible=0"
            Splits(0)._ColumnProps(46)=   "Column(7).AllowFocus=0"
            Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
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
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            InsertMode      =   0   'False
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=39"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
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
      Begin VB.Frame fra_Pesquisa 
         Height          =   1335
         Left            =   120
         TabIndex        =   19
         Top             =   330
         Width           =   8985
         Begin VB.CommandButton cmd_intContribuinte 
            Height          =   315
            Left            =   8145
            Picture         =   "CadDebito.frx":1428
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Ativa o Cadastro de Contribuintes"
            Top             =   900
            Width           =   330
         End
         Begin VB.TextBox txt_strCodigo 
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
            MaxLength       =   15
            TabIndex        =   1
            Top             =   570
            Width           =   975
         End
         Begin VB.TextBox txt_strInscricaoCadastral 
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
            Left            =   4125
            MaxLength       =   50
            TabIndex        =   2
            Top             =   570
            Width           =   1935
         End
         Begin VB.CommandButton cmd_intUtilizacaoDebito 
            Height          =   315
            Left            =   8145
            Picture         =   "CadDebito.frx":1546
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Ativa o Cadastro de Utilização"
            Top             =   240
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.ComboBox cbo_intUtilizacaoDebito 
            Height          =   315
            Left            =   1605
            TabIndex        =   0
            Top             =   240
            Width           =   6495
         End
         Begin MSDataListLib.DataCombo dbc_intContribuinte 
            Height          =   315
            Left            =   1605
            TabIndex        =   3
            Top             =   885
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_strCodigo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   1005
            TabIndex        =   28
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lbl_strInscricao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   2670
            TabIndex        =   27
            Top             =   660
            Width           =   1350
         End
         Begin VB.Label lbl_intContribuinte 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código/Contribuinte"
            Height          =   195
            Left            =   90
            TabIndex        =   26
            Top             =   1005
            Width           =   1410
         End
         Begin VB.Label lbl_intUtilizacaoDebito 
            AutoSize        =   -1  'True
            Caption         =   "Utilização"
            Height          =   195
            Left            =   810
            TabIndex        =   25
            Top             =   360
            Width           =   690
         End
      End
      Begin VB.Label lbl_Cadastro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1830
         TabIndex        =   23
         Top             =   1140
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmCadDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando          As Boolean
Dim mblnPrimeiraVez        As Boolean
Dim mstrModoOperacao       As String

Private Sub cbo_intUtilizacaoDebito_Change()
Dim intIndice As Integer

   With cbo_intUtilizacaoDebito
      If .ListIndex >= 0 Then
          intIndice = .ItemData(.ListIndex)
           
          If intIndice = 2 Or intIndice = 3 Then 'Imobiliária ou Econômica
           
          End If
      End If
   End With

End Sub

Private Sub cmd_intComposicaoReceita_Click()
   ChamaFormCadastro frmCadComposicaoDaReceita, dbc_intComposicaoReceita
End Sub

Private Sub cmd_intContribuinte_Click()
ChamaFormCadastro frmCadContribuinte, dbc_intContribuinte
End Sub

Private Sub cmd_intOcorrencia_Click()
ChamaFormCadastro frmCadOcorrencia, dbc_intOcorrencia
End Sub

Private Sub cmd_intReceitas_Click()
ChamaFormCadastro frmCadReceita, dbc_intReceitas
End Sub

Private Sub cmd_intUtilizacaoDebito_Click()
'ChamaFormCadastro
End Sub

Private Sub dbc_intComposicaoReceita_Click(Area As Integer)
Dim strSQL As String

   DropDownDataCombo dbc_intComposicaoReceita, Me, Area
   
   If Area = 1 Or Area = 2 Then
       With dbc_intComposicaoReceita
           If .MatchedWithList Then
               strSQL = strQueryReceitas(dbc_intComposicaoReceita.BoundText)
               LeDaTabelaParaObj "", dbc_intReceitas, strSQL
           End If
       End With
   End If
End Sub

Private Sub dbc_intComposicaoReceita_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intComposicaoReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intContribuinte_Click(Area As Integer)
   DropDownDataCombo dbc_intContribuinte, Me, Area
   If Area = 1 Or Area = 2 Then
      With dbc_intContribuinte
          If .MatchedWithList Then
              'LimpaFormulario
              
              mblnPrimeiraVez = True
              
              Set tdb_Parcela.DataSource = Nothing
              tdb_Parcela.ReBind
              tdb_Parcela.Refresh
   
   '            CarregaComposicacaoReceita .BoundText
          End If
      End With
   End If
End Sub

Private Sub dbc_intContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intOcorrencia_Click(Area As Integer)
   DropDownDataCombo dbc_intOcorrencia, Me, Area
End Sub

Private Sub dbc_intOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intReceitas_Click(Area As Integer)
   DropDownDataCombo dbc_intReceitas, Me, Area
End Sub

Private Sub dbc_intReceitas_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intReceitas, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 636
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    MDIMenu.Tag = "DEBITO"
    dbc_intContribuinte.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
Dim strSQL As String

mblnAlterando = False
mblnPrimeiraVez = False
HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
                
                
strSQL = ""
strSQL = strSQL & " SELECT PKId, strNome FROM " & gstrContribuinte
If Trim(dbc_intContribuinte.Text) <> "" Then
    strSQL = strSQL & " WHERE strNome LIKE '" & dbc_intContribuinte.Text & "%'"
End If
strSQL = strSQL & " ORDER BY strNome "

dbc_intContribuinte.Tag = strSQL & ";strNome"

cbo_intUtilizacaoDebito.AddItem "Imobiliárias "
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "1"
cbo_intUtilizacaoDebito.AddItem "Apêndice"
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "2"
cbo_intUtilizacaoDebito.AddItem "Fiscalização"
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "3"
cbo_intUtilizacaoDebito.AddItem "Outras Receitas"
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "4"

LeDaTabelaParaObj gstrOcorrencia, dbc_intOcorrencia, strQueryOcorrencia
LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoReceita
Screen.MousePointer = vbDefault
End Sub

Private Function strQueryOcorrencia() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo CAST do SQL Server pela função gstrCONVERT.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL As String
    strSQL = ""
'    strSQL = strSQL & "SELECT O.PKID, RTRIM(CAST(O.intCodigo AS CHAR)) + ' - ' + O.strDescricao AS Ocorrencia "
    strSQL = strSQL & "SELECT O.PKID, RTRIM(" & gstrCONVERT(CDT_VARCHAR, "O.intCodigo") & ") " & strCONCAT & " ' - ' " & strCONCAT & " O.strDescricao AS Ocorrencia "
    strSQL = strSQL & "FROM " & gstrOcorrencia & " O "
    strSQL = strSQL & "WHERE O.intUtilizacaoDaOcorrencia = 1 "
    strSQL = strSQL & "ORDER BY O.intCodigo"
    
    strQueryOcorrencia = strSQL
End Function

Private Function blnDadosOk() As Boolean
Dim strCampo As String
Dim strValor As String

blnDadosOk = False

'Contribuinte
strCampo = lbl_intContribuinte.Caption
If dbc_intContribuinte.Text = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If dbc_intContribuinte.Enabled Then
        dbc_intContribuinte.SetFocus
    End If
    Exit Function
Else
    If Not dbc_intContribuinte.MatchedWithList Then
        strValor = dbc_intContribuinte.Text
        ExibeMensagem strValor & " não é um valor válido para o campo " & strCampo
        If dbc_intContribuinte.Enabled Then
            dbc_intContribuinte.SetFocus
        End If
        Exit Function
    End If
End If
'Origem da Receita
strCampo = lbl_intComposicaoReceita.Caption
If dbc_intComposicaoReceita.Text = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If dbc_intComposicaoReceita.Enabled Then
        dbc_intComposicaoReceita.SetFocus
    End If
    Exit Function
Else
    If Not dbc_intComposicaoReceita.MatchedWithList Then
        strValor = dbc_intComposicaoReceita.Text
        ExibeMensagem strValor & " não é um valor válido para o campo " & strCampo
        If dbc_intComposicaoReceita.Enabled Then
            dbc_intComposicaoReceita.SetFocus
        End If
        Exit Function
    End If
End If
'Receita
strCampo = lbl_intReceitas.Caption
If dbc_intReceitas.Text = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If dbc_intReceitas.Enabled Then
        dbc_intReceitas.SetFocus
    End If
    Exit Function
Else
    If Not dbc_intReceitas.MatchedWithList Then
        strValor = dbc_intReceitas.Text
        ExibeMensagem strValor & " não é um valor válido para o campo " & strCampo
        If dbc_intReceitas.Enabled Then
            dbc_intReceitas.SetFocus
        End If
        Exit Function
    End If
End If
'Utilização do Débito
strCampo = lbl_intUtilizacaoDebito.Caption
If cbo_intUtilizacaoDebito.Text = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If cbo_intUtilizacaoDebito.Enabled Then
        cbo_intUtilizacaoDebito.SetFocus
    End If
    Exit Function
Else
    If cbo_intUtilizacaoDebito.ListIndex < 0 Then
        strValor = cbo_intUtilizacaoDebito.Text
        ExibeMensagem strValor & " não é um valor válido para o campo " & strCampo
        If cbo_intUtilizacaoDebito.Enabled Then
            cbo_intUtilizacaoDebito.SetFocus
        End If
        Exit Function
    End If
End If
'Ocorrência
strCampo = lbl_intOcorrencia.Caption
If dbc_intOcorrencia.Text = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If dbc_intOcorrencia.Enabled Then
        dbc_intOcorrencia.SetFocus
    End If
    Exit Function
Else
    If Not dbc_intOcorrencia.MatchedWithList Then
        strValor = dbc_intOcorrencia.Text
        ExibeMensagem strValor & " não é um valor válido para o campo " & strCampo
        If dbc_intOcorrencia.Enabled Then
            dbc_intOcorrencia.SetFocus
        End If
        Exit Function
    End If
End If
'Exercício
strCampo = lbl_intExercicio.Caption
If Trim(txt_intExercicio.Text) = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If txt_intExercicio.Enabled Then
        txt_intExercicio.SetFocus
    End If
    Exit Function
Else
    If Not gblnDataValida("01/01/" & txt_intExercicio.Text) Then
        strValor = txt_intExercicio.Text
        ExibeMensagem strValor & " não é um valor válido para o campo " & strCampo
        If txt_intExercicio.Enabled Then
            txt_intExercicio.SetFocus
        End If
        Exit Function
    End If
End If
'Número de parcelas
strCampo = lbl_intNumeroParcela.Caption
If Trim(txt_intNumeroParcela.Text) = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If txt_intNumeroParcela.Enabled Then
        txt_intNumeroParcela.SetFocus
    End If
    Exit Function
End If
'Sequencia
strCampo = lbl_strSequencia.Caption
If Trim(txt_strSequencia.Text) = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If txt_strSequencia.Enabled Then
        txt_strSequencia.SetFocus
    End If
    Exit Function
End If
'Data de Vencimento
strCampo = lbl_dtmDataVencimento.Caption
If Trim(txt_dtmDataVencimento.Text) = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If txt_dtmDataVencimento.Enabled Then
        txt_dtmDataVencimento.SetFocus
    End If
    Exit Function
Else
    If Not gblnDataValida(txt_dtmDataVencimento) Then
        strValor = txt_dtmDataVencimento.Text
        ExibeMensagem strValor & " não é um valor válido para o campo " & strCampo
        If txt_dtmDataVencimento.Enabled Then
            txt_dtmDataVencimento.SetFocus
        End If
        Exit Function
    End If
End If
'Data de Lançamento
strCampo = lbl_dtmDataLancamento.Caption
If Trim(txt_dtmDataLancamento.Text) = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If txt_dtmDataLancamento.Enabled Then
        txt_dtmDataLancamento.SetFocus
    End If
    Exit Function
Else
    If Not gblnDataValida(txt_dtmDataLancamento) Then
        strValor = txt_dtmDataLancamento.Text
        ExibeMensagem strValor & " não é um valor válido para o campo " & strCampo
        If txt_dtmDataLancamento.Enabled Then
            txt_dtmDataLancamento.SetFocus
        End If
        Exit Function
    End If
End If
'Valor
strCampo = lbl_dblValorParcela.Caption
If Trim(txt_dblValorParcela.Text) = "" Then
    ExibeMensagem "O campo " & Trim(strCampo) & " não pode ser nulo!"
    If txt_dblValorParcela.Enabled Then
        txt_dblValorParcela.SetFocus
    End If
    Exit Function
End If

blnDadosOk = True
End Function

Private Function VerificaLancamento() As Long
Dim strSQL As String
Dim lngPKId As Long
Dim adoRec As ADODB.Recordset

lngPKId = 0

strSQL = ""
strSQL = strSQL & " SELECT PKId FROM " & gstrLancamentoCalculo
strSQL = strSQL & " WHERE intExercicio = " & txt_intExercicio.Text
strSQL = strSQL & " AND intContribuinte = " & dbc_intContribuinte.BoundText
strSQL = strSQL & " AND intComposicaoReceita = " & dbc_intComposicaoReceita.BoundText
strSQL = strSQL & " AND strSequencia = '" & txt_strSequencia.Text & "'"

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
    With adoRec
        If Not (.BOF And .EOF) Then
            lngPKId = !Pkid
        End If
    End With
Else
    lngPKId = 0
End If

VerificaLancamento = lngPKId
End Function

Private Sub IncluiLancamento()

'******************************************************************************************
' Data: 07/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************

'Dim strSQL As String
'Dim strMsg As String
'Dim i As Integer
'Dim intNumeroDeParcelas As Integer
'Dim lngPKIdLancamentoCalculo As Long
'Dim lngPKIdParcelaReceita As Long
'Dim adoRec As ADODB.Recordset
'
'strMsg = ""
'strMsg = strMsg & "Deseja incluir a Receita " & dbc_intComposicaoReceita.Text
'
'If gblnExclusaoGravacaoOk("I", strMsg, True) Then
'
'    If blnDadosOk Then
'
'        intNumeroDeParcelas = txt_intNumeroParcela.Text
'
'        lngPKIdLancamentoCalculo = VerificaLancamento()
'
'        Set gobjBanco = New clsBanco
'        gobjBanco.ExecutaBeginTrans
'
'        If lngPKIdLancamentoCalculo = 0 Then
'
'            strSQL = ""
'            strSQL = strSQL & " INSERT INTO " & gstrLancamentoCalculo
'            strSQL = strSQL & " ( intExercicio, intContribuinte, intComposicaoReceita, strInscricaoCadastral, "
'            strSQL = strSQL & " dtmLancamento, dtmVencimento, intNumeroDeParcelas, intIntervaloEntreParcelas, "
'            strSQL = strSQL & " bitUtilizacaoDebito, intOcorrencia, bytOrigem, strSequencia, dblAliquota, "
'            strSQL = strSQL & " dtmDtAtualizacao, lngCodUsr ) VALUES ( "
'
'            'Exercício
'            strSQL = strSQL & txt_intExercicio.Text
'            'Contribuinte
'            strSQL = strSQL & ", '" & dbc_intContribuinte.BoundText & "'"
'            'Origem da Receita
'            strSQL = strSQL & ", '" & dbc_intComposicaoReceita.BoundText & "'"
'            'Inscrição Cadastral
'            strSQL = strSQL & ", '" & dbc_intContribuinte.BoundText & "'"
'            'Data de Lançamento
'            strSQL = strSQL & ", " & gstrConvDtParaSql(txt_dtmDataLancamento.Text)
'            'Data de Vencimento
'            strSQL = strSQL & ", " & gstrConvDtParaSql(txt_dtmDataVencimento.Text)
'            'Número de parcelas
'            strSQL = strSQL & ", '" & txt_intNumeroParcela.Text & "'"
'            'Intervalo entre parcelas
'            strSQL = strSQL & ", '0'"
'            'Utilização
'            strSQL = strSQL & ", '" & cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.ListIndex) & "'"
'            'Ocorrência
'            strSQL = strSQL & ", '" & dbc_intOcorrencia.BoundText & "'"
'            'Origem
'            strSQL = strSQL & ", '1'"
'            'Sequencia
'            strSQL = strSQL & ", '" & txt_strSequencia.Text & "'"
'            'Alíquota
'            strSQL = strSQL & ", '0'"
'            'Data Atualização
''            strSQL = strSQL & ", GETDATE()"
'            strSQL = strSQL & ", " & strGETDATE
'            'Usuário
'            strSQL = strSQL & ", '" & glngCodUsr & "'"
'            strSQL = strSQL & " )"
'
'            If Not gobjBanco.Execute(strSQL) Then
'                gobjBanco.ExecutaRollbackTrans
'                Exit Sub
'            End If
'
'            lngPKIdLancamentoCalculo = glngPegaUltimaChave(gstrLancamentoCalculo, "PKId")
'
'            strSQL = ""
'            strSQL = strSQL & " INSERT INTO " & gstrParcelaReceita
'            strSQL = strSQL & " ( intLancamentoCalculo, intComposicaoDaReceita, intNumeroParcela, dtmDataVencimento, "
'            strSQL = strSQL & " dblValorParcela, dblValorDesconto, bytDividaAjuizada, bytSimulado, bytPrescrita, "
'            strSQL = strSQL & " bytCancelada, bytAtiva, bytSuspensaoDeExigencia, dtmDtAtualizacao, lngCodUsr ) "
'            strSQL = strSQL & " VALUES ( "
'
'            'Lancamento Calculo
'            strSQL = strSQL & lngPKIdLancamentoCalculo
'            'Composição da receita
'            strSQL = strSQL & ", '" & dbc_intComposicaoReceita.BoundText & "'"
'            'Número da parcela
'            strSQL = strSQL & ", '" & txt_intNumeroParcela.Text & "'"
'            'Data de vencimento
'            strSQL = strSQL & ", " & gstrConvDtParaSql(txt_dtmDataVencimento.Text)
'            'Valor
'            strSQL = strSQL & ", " & gstrConvVrParaSql(txt_dblValorParcela.Text)
'            'Desconto
'            strSQL = strSQL & ", '0'"
'            'Dívida ajuizada
'            strSQL = strSQL & ", '" & chk_blnDividaAjuizada.Value & "'"
'            'Simulada
'            strSQL = strSQL & ", '" & chk_blnSimulado.Value & "'"
'            'Prescrita
'            strSQL = strSQL & ", '" & chk_blnPrescrita.Value & "'"
'            'Cancelada
'            strSQL = strSQL & ", '" & chk_blnCancelada.Value & "'"
'            'Dívida Ativa
'            strSQL = strSQL & ", '" & chk_blnDividaAtiva.Value & "'"
'            'Suspensão de exigência (Não lança suspensão de exigência neste formulário)
'            strSQL = strSQL & ", '0'"
'            'Atualização
''            strSQL = strSQL & ", GETDATE()"
'            strSQL = strSQL & ", " & strGETDATE
'            'Usuário
'            strSQL = strSQL & ", '" & glngCodUsr & "'"
'            strSQL = strSQL & " )"
'
'            If Not gobjBanco.Execute(strSQL) Then
'                gobjBanco.ExecutaRollbackTrans
'                Exit Sub
'            End If
'        Else
'            strSQL = ""
'            strSQL = strSQL & " SELECT PKId FROM " & gstrParcelaReceita
'            strSQL = strSQL & " WHERE intLancamentoCalculo = " & lngPKIdLancamentoCalculo
'
'            If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
'                If Not adoRec.BOF And adoRec.EOF Then
'                    lngPKIdParcelaReceita = adoRec!PKId
'                End If
'            End If
'
'            strSQL = ""
'            strSQL = strSQL & " UPDATE " & gstrParcelaReceita
'            strSQL = strSQL & " SET dblValorParcela = dblValorParcela + " & gstrConvVrParaSql(txt_dblValorParcela.Text)
'
'            If Not gobjBanco.Execute(strSQL) Then
'                gobjBanco.ExecutaRollbackTrans
'                Exit Sub
'            End If
'
'        End If
'
'        strSQL = ""
'        strSQL = strSQL & " INSERT INTO " & gstrParcelaTaxa
'        strSQL = strSQL & " ( intLancamentoCalculo, intReceita, intNumeroParcela, dtmDataVencimento, "
'        strSQL = strSQL & " dblValorParcela, dtmDtAtualizacao, lngCodUsr ) VALUES ( "
'        'Lançamento Cálculo
'        strSQL = strSQL & lngPKIdLancamentoCalculo
'        'Receita
'        strSQL = strSQL & ", '" & dbc_intReceitas.BoundText & "'"
'        'Número da parcela
'        strSQL = strSQL & ", '" & txt_intNumeroParcela.Text & "'"
'        'Vencimento
'        strSQL = strSQL & ", " & gstrConvDtParaSql(txt_dtmDataVencimento.Text)
'        'Valor
'        strSQL = strSQL & ", " & gstrConvVrParaSql(txt_dblValorParcela.Text)
'        'Atualização
''        strSQL = strSQL & ", GETDATE()"
'        strSQL = strSQL & ", " & strGETDATE
'        'Usuário
'        strSQL = strSQL & ", '" & glngCodUsr & "'"
'        strSQL = strSQL & " )"
'
'        If Not gobjBanco.Execute(strSQL) Then
'            gobjBanco.ExecutaRollbackTrans
'            Exit Sub
'        End If
'
'        gobjBanco.ExecutaCommitTrans
'
'        CarregaComposicacaoReceita dbc_intContribuinte.BoundText
'        mblnPrimeiraVez = True
'        tdb_Composicao_RowColChange 0, 0
'        mblnPrimeiraVez = False
'    End If 'Dados OK
'
'End If 'Confirmação
End Sub

Private Sub CarregaDivida(lngPKIdContribuinte As Long)

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo YEAR() do SQL Server pela função gstrDATEPART
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 07/05/2003
' Alteração: - Foram substituídos os nomes das colunas na cláusula ORDER BY pelos apelidos
'            utilizados na cláusula SELECT.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSQL As String

If cbo_intUtilizacaoDebito.ListIndex = -1 Then
    strSQL = ""
    strSQL = strSQL & " SELECT 'Imobiliário' AS bitUtilizacao, LC.PKId, IM.PKId AS intCodigoOrigem, IM.strInscricaoAnterior AS strInscricao, "
'    strSql = strSql & " CR.strDescricao , YEAR(LC.dtmVencimento) AS intExercicio, SUM(dblValorParcela) AS dblValor, "
    strSQL = strSQL & " CR.strDescricao , " & gstrDATEPART(strYEAR, "LC.dtmVencimento") & " AS intExercicio, SUM(dblValorParcela) AS dblValor, "
    strSQL = strSQL & " CO.PKId as PKIdContribuinte "
    strSQL = strSQL & " FROM " & gstrContribuinte & " CO, "
    strSQL = strSQL & gstrImobiliario & " IM, "
    strSQL = strSQL & gstrLancamentoCalculo & " LC, "
    strSQL = strSQL & gstrParcelaReceita & " PR, "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR "
    strSQL = strSQL & " WHERE CO.PKId = IM.intContribuinte "
    strSQL = strSQL & " AND CO.PKId = LC.intContribuinte "
    strSQL = strSQL & " AND IM.strInscricaoAnterior = LC.strInscricaoCadastral"
    strSQL = strSQL & " AND LC.PKId = PR.intLancamentoCalculo"
    strSQL = strSQL & " AND CR.PKId = PR.intComposicaoDaReceita"
    
    If lngPKIdContribuinte > 0 Then
        If dbc_intContribuinte.MatchedWithList Then
            strSQL = strSQL & " AND CO.PKId = '" & lngPKIdContribuinte & "'"
        Else
            strSQL = strSQL & " AND CO.strCodigoAnterior = '" & lngPKIdContribuinte & "'"
        End If
    ElseIf Trim(dbc_intContribuinte.Text) <> "" Then
        strSQL = strSQL & " AND CO.strNome LIKE '" & dbc_intContribuinte.Text & "%'"
    End If
    
    If Trim(txt_strCodigo.Text) <> "" Then
        strSQL = strSQL & " AND IM.PKId = " & Val(txt_strCodigo.Text)
    End If
    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
        strSQL = strSQL & " AND IM.strInscricaoAnterior = '" & txt_strInscricaoCadastral.Text & "'"
    End If
    strSQL = strSQL & " AND PR.intNumeroParcela <> 0 "
'    strSql = strSql & " GROUP BY LC.PKId, IM.PKId, YEAR(LC.dtmVencimento), PR.intLancamentoCalculo, IM.strInscricaoAnterior, "
    strSQL = strSQL & " GROUP BY LC.PKId, IM.PKId, " & gstrDATEPART(strYEAR, "LC.dtmVencimento") & ", PR.intLancamentoCalculo, IM.strInscricaoAnterior, "
    strSQL = strSQL & " CO.PKId, CR.strDescricao "
    
    strSQL = strSQL & " UNION "

    strSQL = strSQL & " SELECT 'Apêndice' AS bitUtilizacao, LC.PKId, EC.PKId AS intCodigoOrigem, EC.strInscricaoCadastral AS strInscricao, "
'    strSql = strSql & " CR.strDescricao , YEAR(LC.dtmVencimento) AS intExercicio, SUM(dblValorParcela) AS dblValor, "
    strSQL = strSQL & " CR.strDescricao , " & gstrDATEPART(strYEAR, "LC.dtmVencimento") & " AS intExercicio, SUM(dblValorParcela) AS dblValor, "
    strSQL = strSQL & " CO.PKId as PKIdContribuinte "
    strSQL = strSQL & " FROM " & gstrContribuinte & " CO, "
    strSQL = strSQL & gstrEconomico & " EC, "
    strSQL = strSQL & gstrLancamentoCalculo & " LC, "
    strSQL = strSQL & gstrParcelaReceita & " PR, "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR "
    strSQL = strSQL & " WHERE CO.PKId = EC.intContribuinte "
    strSQL = strSQL & " AND CO.PKId = LC.intContribuinte "
    strSQL = strSQL & " AND EC.strInscricaoCadastral = LC.strInscricaoCadastral"
    strSQL = strSQL & " AND LC.PKId = PR.intLancamentoCalculo"
    strSQL = strSQL & " AND CR.PKId = PR.intComposicaoDaReceita"
    
    If lngPKIdContribuinte > 0 Then
        If dbc_intContribuinte.MatchedWithList Then
            strSQL = strSQL & " AND CO.PKId = '" & lngPKIdContribuinte & "'"
        Else
            strSQL = strSQL & " AND CO.strCodigoAnterior = '" & lngPKIdContribuinte & "'"
        End If
    ElseIf Trim(dbc_intContribuinte.Text) <> "" Then
        strSQL = strSQL & " AND CO.strNome LIKE '" & dbc_intContribuinte.Text & "%'"
    End If
    
    If Trim(txt_strCodigo.Text) <> "" Then
        strSQL = strSQL & " AND EC.PKId = " & Val(txt_strCodigo.Text)
    End If
    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
        strSQL = strSQL & " AND EC.strInscricaoCadastral = '" & txt_strInscricaoCadastral.Text & "'"
    End If
    strSQL = strSQL & " AND PR.intNumeroParcela <> 0 "
'    strSql = strSql & " GROUP BY LC.PKId, EC.PKId, YEAR(LC.dtmVencimento), PR.intLancamentoCalculo, EC.strInscricaoCadastral, "
    strSQL = strSQL & " GROUP BY LC.PKId, EC.PKId, " & gstrDATEPART(strYEAR, "LC.dtmVencimento") & ", PR.intLancamentoCalculo, EC.strInscricaoCadastral, "
    strSQL = strSQL & " CO.PKId, CR.strDescricao "
    
    LeDaTabelaParaObj "", tdb_Divida, strSQL
    
    If Not IsNull(tdb_Divida.Columns("PKIdContribuinte").Value) Then
        LeDaTabelaParaObj gstrContribuinte, dbc_intContribuinte, " SELECT PKId, strNome FROM " & gstrContribuinte & " WHERE PKId = " & tdb_Divida.Columns("PKIdContribuinte").Value
        dbc_intContribuinte.BoundText = tdb_Divida.Columns("PKIdContribuinte").Value
    Else
        Set dbc_intContribuinte.RowSource = Nothing
        dbc_intContribuinte.Text = ""
    End If

ElseIf cbo_intUtilizacaoDebito.ListIndex = 0 Then 'Imobiliário
    strSQL = ""
    strSQL = strSQL & " SELECT 'Imobiliário' AS bitUtilizacao, LC.PKId, IM.PKId AS intCodigoOrigem, IM.strInscricaoAnterior AS strInscricao, "
'    strSql = strSql & " CR.strDescricao , YEAR(LC.dtmVencimento) AS intExercicio, SUM(dblValorParcela) AS dblValor, "
    strSQL = strSQL & " CR.strDescricao , " & gstrDATEPART(strYEAR, "LC.dtmVencimento") & " AS intExercicio, SUM(dblValorParcela) AS dblValor, "
    strSQL = strSQL & " CO.PKId as PKIdContribuinte "
    strSQL = strSQL & " FROM " & gstrContribuinte & " CO, "
    strSQL = strSQL & gstrImobiliario & " IM, "
    strSQL = strSQL & gstrLancamentoCalculo & " LC, "
    strSQL = strSQL & gstrParcelaReceita & " PR, "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR "
    strSQL = strSQL & " WHERE CO.PKId = IM.intContribuinte "
    strSQL = strSQL & " AND CO.PKId = LC.intContribuinte "
    strSQL = strSQL & " AND IM.strInscricaoAnterior = LC.strInscricaoCadastral"
    strSQL = strSQL & " AND LC.PKId = PR.intLancamentoCalculo"
    strSQL = strSQL & " AND CR.PKId = PR.intComposicaoDaReceita"
    
    If lngPKIdContribuinte > 0 Then
        If dbc_intContribuinte.MatchedWithList Then
            strSQL = strSQL & " AND CO.PKId = '" & lngPKIdContribuinte & "'"
        Else
            strSQL = strSQL & " AND CO.strCodigoAnterior = '" & lngPKIdContribuinte & "'"
        End If
    ElseIf Trim(dbc_intContribuinte.Text) <> "" Then
        strSQL = strSQL & " AND CO.strNome LIKE '" & dbc_intContribuinte.Text & "%'"
    End If
    
    If Trim(txt_strCodigo.Text) <> "" Then
        strSQL = strSQL & " AND IM.PKId = " & Val(txt_strCodigo.Text)
    End If
    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
        strSQL = strSQL & " AND IM.strInscricaoAnterior = '" & txt_strInscricaoCadastral.Text & "'"
    End If
    strSQL = strSQL & " AND PR.intNumeroParcela <> 0 "
'    strSql = strSql & " GROUP BY LC.PKId, IM.PKId, YEAR(LC.dtmVencimento), PR.intLancamentoCalculo, IM.strInscricaoAnterior, "
    strSQL = strSQL & " GROUP BY LC.PKId, IM.PKId, " & gstrDATEPART(strYEAR, "LC.dtmVencimento") & ", PR.intLancamentoCalculo, IM.strInscricaoAnterior, "
    strSQL = strSQL & " CO.PKId, CR.strDescricao "
'    strSQL = strSQL & " ORDER BY IM.PKId, LC.intExercicio "
    strSQL = strSQL & " ORDER BY IM.PKId, intExercicio "
    
    LeDaTabelaParaObj "", tdb_Divida, strSQL
    If Not IsNull(tdb_Divida.Columns("PKIdContribuinte").Value) Then
        LeDaTabelaParaObj gstrContribuinte, dbc_intContribuinte, " SELECT PKId, strNome FROM " & gstrContribuinte & " WHERE PKId = " & tdb_Divida.Columns("PKIdContribuinte").Value
        dbc_intContribuinte.BoundText = tdb_Divida.Columns("PKIdContribuinte").Value
    Else
        Set dbc_intContribuinte.RowSource = Nothing
        dbc_intContribuinte.Text = ""
    End If
ElseIf cbo_intUtilizacaoDebito.ListIndex = 1 Then 'Econômica
    strSQL = ""
    strSQL = strSQL & " SELECT 'Apêndice' AS bitUtilizacao, LC.PKId, EC.PKId AS intCodigoOrigem, EC.strInscricaoCadastral AS strInscricao, "
'    strSql = strSql & " CR.strDescricao , YEAR(LC.dtmVencimento) AS intExercicio, SUM(dblValorParcela) AS dblValor, "
    strSQL = strSQL & " CR.strDescricao , " & gstrDATEPART(strYEAR, "LC.dtmVencimento") & " AS intExercicio, SUM(dblValorParcela) AS dblValor, "
    strSQL = strSQL & " CO.PKId as PKIdContribuinte "
    strSQL = strSQL & " FROM " & gstrContribuinte & " CO, "
    strSQL = strSQL & gstrEconomico & " EC, "
    strSQL = strSQL & gstrLancamentoCalculo & " LC, "
    strSQL = strSQL & gstrParcelaReceita & " PR, "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR "
    strSQL = strSQL & " WHERE CO.PKId = EC.intContribuinte "
    strSQL = strSQL & " AND CO.PKId = LC.intContribuinte "
    strSQL = strSQL & " AND EC.strInscricaoCadastral = LC.strInscricaoCadastral"
    strSQL = strSQL & " AND LC.PKId = PR.intLancamentoCalculo"
    strSQL = strSQL & " AND CR.PKId = PR.intComposicaoDaReceita"
    
    If lngPKIdContribuinte > 0 Then
        If dbc_intContribuinte.MatchedWithList Then
            strSQL = strSQL & " AND CO.PKId = '" & lngPKIdContribuinte & "'"
        Else
            strSQL = strSQL & " AND CO.strCodigoAnterior = '" & lngPKIdContribuinte & "'"
        End If
    ElseIf Trim(dbc_intContribuinte.Text) <> "" Then
        strSQL = strSQL & " AND CO.strNome LIKE '" & dbc_intContribuinte.Text & "%'"
    End If
    
    If Trim(txt_strCodigo.Text) <> "" Then
        strSQL = strSQL & " AND EC.PKId = " & Val(txt_strCodigo.Text)
    End If
    If Trim(txt_strInscricaoCadastral.Text) <> "" Then
        strSQL = strSQL & " AND EC.strInscricaoCadastral = '" & txt_strInscricaoCadastral.Text & "'"
    End If
    strSQL = strSQL & " AND PR.intNumeroParcela <> 0 "
'    strSql = strSql & " GROUP BY LC.PKId, EC.PKId, YEAR(LC.dtmVencimento), PR.intLancamentoCalculo, EC.strInscricaoCadastral, "
    strSQL = strSQL & " GROUP BY LC.PKId, EC.PKId, " & gstrDATEPART(strYEAR, "LC.dtmVencimento") & ", PR.intLancamentoCalculo, EC.strInscricaoCadastral, "
    strSQL = strSQL & " CO.PKId, CR.strDescricao "
'    strSQL = strSQL & " ORDER BY EC.PKId, LC.intExercicio "
    strSQL = strSQL & " ORDER BY EC.PKId, intExercicio "
    
    LeDaTabelaParaObj "", tdb_Divida, strSQL
    If Not IsNull(tdb_Divida.Columns("PKIdContribuinte").Value) Then
        LeDaTabelaParaObj gstrContribuinte, dbc_intContribuinte, " SELECT PKId, strNome FROM " & gstrContribuinte & " WHERE PKId = " & tdb_Divida.Columns("PKIdContribuinte").Value
        dbc_intContribuinte.BoundText = tdb_Divida.Columns("PKIdContribuinte").Value
    Else
        Set dbc_intContribuinte.RowSource = Nothing
        dbc_intContribuinte.Text = ""
    End If
End If
tab_Receita.Tab = 0
End Sub

Public Sub MantemForm(strModoOperacao As String)

'******************************************************************************************
' Data: 08/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSQL As String
Dim strInscricaoCadastral As String
Dim blnExisteLancamento As Boolean

mstrModoOperacao = strModoOperacao
Screen.MousePointer = vbHourglass
            
Select Case strModoOperacao
    Case gstrNovo
        LimpaFormulario
        mblnAlterando = False
    Case gstrSalvar
        IncluiLancamento
        mblnAlterando = False
    Case gstrLocalizar
        'If dbc_intContribuinte.MatchedWithList Then

            CarregaDivida Val(dbc_intContribuinte.BoundText)
            mstrModoOperacao = ""
        'End If
    Case gstrPreencherLista
        
        LimpaFormulario
        
        If Me.ActiveControl.Name = "dbc_intContribuinte" Then
            PreencherListaDeOpcoes dbc_intContribuinte
        End If

    Case gstrImprimir
        If Not tdb_Composicao.BOF And Not tdb_Composicao.EOF Then
            MDIMenu.Tag = "frmCadDebito"
            strInscricaoCadastral = tdb_Composicao.Columns("strInscricaoCadastral").Value
            
            Set gobjBanco = New clsBanco
'            gobjBanco.Execute "sp_calculoMultaJuros " & tdb_Divida.Columns("PKId").Value
            gobjBanco.Execute gstrStoredProcedure("sp_calculoMultaJuros", tdb_Divida.Columns("PKId").Value)
            
            strSQL = gstrQueryRelatorioGuiaDeArrecadacao(blnExisteLancamento, strInscricaoCadastral, strInscricaoCadastral, txt_intExercicio.Text, dbc_intComposicaoReceita.BoundText, False, , txt_intNumeroParcela.Text, txt_intNumeroParcela.Text)
            
            If blnExisteLancamento Then
                Set gfrmFormularioQueEstaImprimindoGuia = Me
                
                rptGuiaDeArrecadacaoMunicipal.strImposto = dbc_intComposicaoReceita.Text
                ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, strSQL
            End If
        End If
End Select

Screen.MousePointer = vbDefault
End Sub

Private Sub LimpaFormulario()

'Limpa o formulário para não manter os dados do contribuinte anterior
dbc_intComposicaoReceita.Text = ""
dbc_intReceitas.Text = ""
'cbo_intUtilizacaoDebito.ListIndex = -1
dbc_intOcorrencia.Text = ""
txt_intExercicio.Text = ""
txt_intNumeroParcela.Text = ""
txt_strSequencia.Text = ""
txt_dtmDataVencimento.Text = ""
txt_dtmDataLancamento.Text = ""
txt_dblValorParcela.Text = ""
txt_strCodigo.Text = ""
'txt_strCodigo.SetFocus
txt_strInscricaoCadastral.Text = ""
dbc_intContribuinte.Text = ""

Set tdb_Composicao.DataSource = Nothing
Set tdb_Divida.DataSource = Nothing
Set tdb_Parcela.DataSource = Nothing

tab_Receita.Tab = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
mblnPrimeiraVez = False
End Sub

Private Sub tdb_Composicao_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_Composicao_DblClick()
tab_Receita.Tab = 2
End Sub

Private Sub tdb_Composicao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim strSQL As String
Screen.MousePointer = vbHourglass
'Verifica se selecionou contribuinte
If Not dbc_intContribuinte.MatchedWithList Then
    Screen.MousePointer = vbDefault
    Exit Sub

End If

'If mblnPrimeiraVez Then

    With tdb_Composicao
        gCorLinhaSelecionada tdb_Composicao
        
        mblnAlterando = True
        
        dbc_intReceitas.Text = ""
        
        If Not IsNull(.Columns("dtmDataVencimento").Value) Then
            txt_dtmDataVencimento.Text = gstrDataFormatada(.Columns("dtmDataVencimento").Value)
        Else
            txt_dtmDataVencimento.Text = ""
        End If
        If Not IsNull(tdb_Composicao.Columns("dtmLancamento").Value) Then
            txt_dtmDataLancamento.Text = gstrDataFormatada(tdb_Composicao.Columns("dtmLancamento").Value)
        Else
            txt_dtmDataLancamento.Text = ""
        End If
        If Not IsNull(.Columns("dblValorParcela").Value) Then
            txt_dblValorParcela.Text = gstrConvVrDoSql(.Columns("dblTotalPago").Value)
        Else
            txt_dblValorParcela.Text = ""
        End If
        
        If Not IsNull(.Columns("intNumeroParcela").Value) Then
            CarregaReceitas .Columns("PKId").Value, .Columns("intNumeroParcela").Value
        End If
        
        'LimpaFormulario
   
        'Preenche formulário
        dbc_intComposicaoReceita.Text = IIf(IsNull(.Columns("strDescricao").Value), "", .Columns("strDescricao").Value)
        txt_intNumeroParcela.Text = IIf(IsNull(.Columns("intNumeroParcela").Value), "", .Columns("intNumeroParcela").Value)
        
        tdb_Parcela_RowColChange 0, 0
    End With

    'Monta o Data Combo de receitas
    If dbc_intComposicaoReceita.BoundText <> "" Then
'        strSql = strQueryReceitas(dbc_intComposicaoReceita.BoundText)
'        LeDaTabelaParaObj "", dbc_intReceitas, strSql
    End If
    'tdb_Parcela_RowColChange 0, 0
    
    mblnPrimeiraVez = False
'End If

Screen.MousePointer = vbDefault
End Sub

Private Function strQueryReceitas(lngPKIdComposicaoDaReceita As Long) As String
Dim strSQL As String

strSQL = ""
strSQL = strSQL & " SELECT A.PKId, A.strDescricao FROM "
strSQL = strSQL & gstrReceita & " A,"
strSQL = strSQL & gstrValorCompoRec & " B"
strSQL = strSQL & " WHERE A.PKId = B.intReceita "
strSQL = strSQL & " AND B.intComposicaoDaReceita = " & lngPKIdComposicaoDaReceita

strQueryReceitas = strSQL
End Function

Private Sub tdb_Divida_DblClick()
tab_Receita.Tab = 1
End Sub

Private Sub tdb_Divida_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Screen.MousePointer = vbHourglass
With tdb_Divida
    If Not .BOF And Not .EOF Then
        gCorLinhaSelecionada tdb_Divida
        txt_strCodigo.Text = .Columns("intCodigoOrigem").Value
        txt_strInscricaoCadastral.Text = .Columns("strInscricao").Value
        LeDaTabelaParaObj gstrContribuinte, dbc_intContribuinte, "SELECT PKId, strNome FROM " & gstrContribuinte & " WHERE PKId = " & .Columns("PKIdContribuinte").Value
        dbc_intContribuinte.BoundText = .Columns("PKIdContribuinte").Value
        CarregaComposicacaoReceita .Columns("PKId").Value, .Columns("strInscricao").Value
        tdb_Composicao_RowColChange 0, 0
    End If
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub tdb_Parcela_Click()
mblnPrimeiraVez = True
End Sub

Private Sub tdb_Parcela_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Screen.MousePointer = vbHourglass
'If mblnPrimeiraVez Then
    With tdb_Parcela
        gCorLinhaSelecionada tdb_Parcela
    
        dbc_intReceitas.Text = IIf(IsNull(.Columns("strDescricao").Value), "", .Columns("strDescricao").Value)
        'cbo_intUtilizacaoDebito.ListIndex = IIf(IsNull(.Columns("bitUtilizacaoDebito").Value), -1, .Columns("bitUtilizacaoDebito").Value - 1)
        dbc_intOcorrencia.BoundText = IIf(IsNull(.Columns("intOcorrencia").Value), "", .Columns("intOcorrencia").Value)
        txt_intExercicio.Text = IIf(IsNull(.Columns("intExercicio").Value), "", .Columns("intExercicio").Value)
        txt_intNumeroParcela.Text = IIf(IsNull(.Columns("intNumeroParcela").Value), "", .Columns("intNumeroParcela").Value)
        txt_strSequencia.Text = IIf(IsNull(.Columns("strSequencia").Value), "", .Columns("strSequencia").Value)
        
        mblnPrimeiraVez = False
    End With
'End If
Screen.MousePointer = vbDefault
End Sub

Private Sub CarregaReceitas(lngPKIdLancamentoCalculo As Long, intNumeroParcela As Integer)
Dim strSQL As String

strSQL = ""
strSQL = strSQL & "SELECT PT.PKId, LC.intExercicio, PT.intNumeroParcela, LC.strSequencia, PT.dtmDataVencimento, "
strSQL = strSQL & "PT.dblValorParcela, LC.bitUtilizacaoDebito, LC.intOcorrencia, CR.strDescricao "
strSQL = strSQL & "FROM " & gstrReceita & " CR, "
strSQL = strSQL & gstrParcelaTaxa & " PT, "
strSQL = strSQL & gstrLancamentoCalculo & " LC "
strSQL = strSQL & "WHERE CR.PKID = PT.intReceita "
strSQL = strSQL & " AND LC.PKId = PT.intLancamentoCalculo "
strSQL = strSQL & " AND LC.intContribuinte = " & dbc_intContribuinte.BoundText
strSQL = strSQL & " AND PT.intLancamentoCalculo = '" & lngPKIdLancamentoCalculo & "'"
strSQL = strSQL & " AND PT.intNumeroParcela = '" & intNumeroParcela & "'"

LeDaTabelaParaObj gstrParcelaReceita, tdb_Parcela, strSQL
End Sub

Private Sub CarregaComposicacaoReceita(lngPKIdLancamento As Long, strInscricao As String)

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo YEAR() do SQL Server pela função gstrDATEPART
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 08/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSQL As String
Dim strSqlAux As String
Dim adoRec As ADODB.Recordset
Static lngLancamento As Long

If lngLancamento <> lngPKIdLancamento Then
    Set gobjBanco = New clsBanco
'    gobjBanco.Execute "sp_CalculoMultaJuros " & lngPKIdLancamento
    gobjBanco.Execute gstrStoredProcedure("sp_CalculoMultaJuros", CStr(lngPKIdLancamento))
    lngLancamento = lngPKIdLancamento
End If

strSQL = ""
'strSql = strSql & " SELECT LC.PKId, CR.strDescricao, YEAR(LC.dtmVencimento)  AS intExercicio, PR.intNumeroParcela, LC.strSequencia, "
strSQL = strSQL & " SELECT LC.PKId, CR.strDescricao, " & gstrDATEPART(strYEAR, "LC.dtmVencimento") & "  AS intExercicio, PR.intNumeroParcela, LC.strSequencia, "
strSQL = strSQL & " PR.dtmDataVencimento, PR.dtmDataPagamento, dblJuros, dblMulta, PR.dblValorParcela, "
'strSql = strSql & " (dblJuros + dblMulta + PR.dblValorParcela) AS dblTotalPago, LC.dtmLancamento, LC.strInscricaoCadastral, ISNULL(strSituacao,'A') AS Situacao"
strSQL = strSQL & " (dblJuros + dblMulta + PR.dblValorParcela) AS dblTotalPago, LC.dtmLancamento, LC.strInscricaoCadastral, " & gstrISNULL("strSituacao", "'A'") & " AS Situacao"
strSQL = strSQL & " FROM "
strSQL = strSQL & gstrComposicaoDaReceita & " CR, "
strSQL = strSQL & gstrParcelaReceita & " PR, "
strSQL = strSQL & gstrLancamentoCalculo & " LC "
strSQL = strSQL & " WHERE "
strSQL = strSQL & " CR.PKId = PR.intComposicaoDaReceita "
strSQL = strSQL & " AND LC.PKId = PR.intLancamentoCalculo "
strSQL = strSQL & " AND LC.PKId = " & lngPKIdLancamento

'strSql = strSql & " UNION "
'
'strSql = strSql & " SELECT PP.PKId, CR.strDescricao, PP.intExercicio, PP.intNumeroDaParcela, PP.strSequencia, "
'strSql = strSql & " PP.dtmDataVencimento, PP.dtmDataPagamento, PP.dblJuros, PP.dblMulta, PP.dblValorParcela, "
'strSql = strSql & " PP.dblTotalPago, PP.dtmDataLancamento, PP.strInscricaoCadastral, 'P' AS Situacao"
'strSql = strSql & " FROM "
'strSql = strSql & gstrComposicaoDaReceita & " CR, "
'strSql = strSql & gstrPagamentoParcela & " PP "
'strSql = strSql & " WHERE "
'strSql = strSql & " CR.PKId = PP.intComposicaoDaReceita"
'strSql = strSql & " AND PP.strInscricaoCadastral = '" & Val(strInscricao) & "'"
'strSql = strSql & " ORDER BY CR.strDescricao , LC.intExercicio, LC.strSequencia, PR.intNumeroParcela, "
'strSql = strSql & " PR.dtmDataVencimento"

'strSql = ""
'strSql = strSql & " SELECT LC.PKId, CR.strDescricao, LC.intExercicio, PR.intNumeroParcela, LC.strSequencia, "
'strSql = strSql & " PR.dtmDataVencimento, PR.dblValorParcela, LC.dtmLancamento, LC.strInscricaoCadastral, 'FALSE' AS blnImprimir "
'strSql = strSql & " FROM " & gstrComposicaoDaReceita & " CR, "
'strSql = strSql & gstrParcelaReceita & " PR, "
'strSql = strSql & gstrLancamentoCalculo & " LC "
'strSql = strSql & " WHERE CR.PKID = PR.intComposicaoDaReceita "
'strSql = strSql & " AND LC.PKId = PR.intLancamentoCalculo "
'strSql = strSql & " AND LC.PKId = " & lngPKIdLancamento
'strSql = strSql & " ORDER BY CR.strDescricao, LC.intExercicio, LC.strSequencia, PR.intNumeroParcela, "
'strSql = strSql & " PR.dtmDataVencimento "

LeDaTabelaParaObj "", tdb_Composicao, strSQL
End Sub

Private Sub txt_dblValorParcela_GotFocus()
MarcaCampo txt_dblValorParcela
End Sub

Private Sub txt_dblValorParcela_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "V", txt_dblValorParcela
End Sub

Private Sub txt_dblValorParcela_LostFocus()
txt_dblValorParcela = gstrConvVrDoSql(txt_dblValorParcela, 2)
End Sub

Private Sub txt_dtmDataLancamento_GotFocus()
MarcaCampo txt_dtmDataLancamento
End Sub

Private Sub txt_dtmDataLancamento_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "D", txt_dtmDataLancamento
End Sub

Private Sub txt_dtmDataLancamento_LostFocus()
txt_dtmDataLancamento = gstrDataFormatada(txt_dtmDataLancamento)
End Sub

Private Sub txt_dtmDataVencimento_GotFocus()
MarcaCampo txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "D", txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_LostFocus()
txt_dtmDataVencimento.Text = gstrDataFormatada(txt_dtmDataVencimento.Text)
txt_dtmDataLancamento.Text = txt_dtmDataVencimento.Text
End Sub

Private Sub txt_intExercicio_GotFocus()
MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_intNumeroParcela_GotFocus()
MarcaCampo txt_intNumeroParcela
End Sub

Private Sub txt_intNumeroParcela_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txt_intNumeroParcela
End Sub

