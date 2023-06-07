VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadReceitasDiversas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receitas Diversas"
   ClientHeight    =   7230
   ClientLeft      =   735
   ClientTop       =   1035
   ClientWidth     =   8220
   Icon            =   "CadReceitasDiversas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8220
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   7185
      Left            =   30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   30
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12674
      _Version        =   393216
      Style           =   1
      TabHeight       =   529
      TabCaption(0)   =   "Receitas Diversas"
      TabPicture(0)   =   "CadReceitasDiversas.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintContribuinte"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_intOcorrencia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dbc_intOcorrencia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tdb_Geral"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintContribuinte"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_Producao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_Data"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtintCodigoGeral"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Contribuinte"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtPKId"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Taxas"
      TabPicture(1)   =   "CadReceitasDiversas.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Atividades"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Emissão de Guias de Arrecadação"
      TabPicture(2)   =   "CadReceitasDiversas.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_EmissaoDeGuias"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtPKId 
         Height          =   285
         Left            =   4800
         TabIndex        =   46
         Top             =   30
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5955
         Left            =   -74790
         TabIndex        =   30
         Top             =   480
         Width           =   7605
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1695
            Left            =   330
            TabIndex        =   39
            Top             =   2760
            Width           =   6945
            Begin VB.TextBox txt_Mensagem2 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   34
               Top             =   0
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   35
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label lbl_Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   40
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1695
            Left            =   330
            TabIndex        =   37
            Top             =   690
            Width           =   6945
            Begin VB.TextBox txt_Mensagem1 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin VB.CheckBox chk_EmBranco1 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   31
               Top             =   0
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem1 
               Height          =   315
               Left            =   1080
               TabIndex        =   32
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label lbl_Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   390
               Width           =   780
            End
         End
      End
      Begin VB.Frame fra_Atividades 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   7845
         Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
            Height          =   5595
            Left            =   120
            TabIndex        =   29
            Top             =   270
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   9869
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   "PKId"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   68
            Columns(1)._MaxComboItems=   20
            Columns(1).ValueItems(0)._DefaultItem=   0
            Columns(1).ValueItems(0).Value=   ""
            Columns(1).ValueItems(0).Value.vt=   8
            Columns(1).ValueItems(0).DisplayValue=   ""
            Columns(1).ValueItems(0).DisplayValue.vt=   8
            Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(1).ValueItems.Count=   1
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descrição"
            Columns(2).DataField=   "strDescricao"
            Columns(2).DropDown=   "tdd_Atividades"
            Columns(2).DropDown.vt=   8
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   1
            Splits(0).MarqueeStyle=   5
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=450"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=370"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=10848"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=10769"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(2).AutoDropDown=1"
            Splits(0)._ColumnProps(20)=   "Column(2).DropDownList=1"
            Splits(0)._ColumnProps(21)=   "Column(2).AutoCompletion=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DataMode        =   4
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   0
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
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.CommandButton cmd_Contribuinte 
         Height          =   315
         Left            =   7620
         Picture         =   "CadReceitasDiversas.frx":1096
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de Contribuintes"
         Top             =   495
         Width           =   360
      End
      Begin VB.TextBox txtintCodigoGeral 
         Height          =   315
         Left            =   1500
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin VB.Frame fra_Data 
         Height          =   2085
         Left            =   150
         TabIndex        =   19
         Top             =   3510
         Width           =   7845
         Begin VB.TextBox txtintNumero 
            Height          =   315
            Left            =   330
            TabIndex        =   45
            Top             =   1560
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   4185
            MaxLength       =   4
            TabIndex        =   4
            Top             =   210
            Width           =   525
         End
         Begin VB.TextBox txt_intParcelaInicial 
            Height          =   285
            Left            =   1755
            MaxLength       =   15
            TabIndex        =   6
            Top             =   555
            Width           =   480
         End
         Begin VB.TextBox txt_intParcelaFinal 
            Height          =   285
            Left            =   2745
            MaxLength       =   15
            TabIndex        =   7
            Top             =   555
            Width           =   480
         End
         Begin VB.TextBox txtintIntervalo 
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
            Left            =   6720
            MaxLength       =   3
            TabIndex        =   8
            Top             =   570
            Width           =   975
         End
         Begin VB.TextBox txtdtmVencimento 
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
            Left            =   6720
            MaxLength       =   10
            TabIndex        =   5
            Top             =   210
            Width           =   975
         End
         Begin VB.TextBox txtdtmLancamento 
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
            Left            =   1770
            MaxLength       =   10
            TabIndex        =   3
            Top             =   210
            Width           =   975
         End
         Begin VB.TextBox txt_Mensagem 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   675
            Left            =   1770
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1290
            Width           =   5895
         End
         Begin VB.CommandButton cmd_Mensagem 
            Height          =   315
            Left            =   7320
            Picture         =   "CadReceitasDiversas.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Tag             =   "605"
            ToolTipText     =   "Ativa Cadastro de Mensagens"
            Top             =   900
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintMensagem 
            Height          =   315
            Left            =   1770
            TabIndex        =   9
            Top             =   930
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   3360
            TabIndex        =   44
            Top             =   255
            Width           =   675
         End
         Begin VB.Label lbl_ParcelaIncial 
            AutoSize        =   -1  'True
            Caption         =   "Parcela"
            Height          =   195
            Left            =   1095
            TabIndex        =   43
            Top             =   600
            Width           =   540
         End
         Begin VB.Label lbl_ParcelaFinal 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   2370
            TabIndex        =   42
            Top             =   600
            Width           =   225
         End
         Begin VB.Label lblintIntervalo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Intervalo de dias entre parcelas"
            Height          =   195
            Left            =   4380
            TabIndex        =   24
            Top             =   630
            Width           =   2220
         End
         Begin VB.Label lbldtmLancamento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data de Lançamento"
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   255
            Width           =   1500
         End
         Begin VB.Label lbldtmVencimento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data de vencimento"
            Height          =   195
            Left            =   5160
            TabIndex        =   22
            Top             =   255
            Width           =   1440
         End
         Begin VB.Label lblintMensagem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mensagem"
            Height          =   195
            Left            =   870
            TabIndex        =   21
            Top             =   1020
            Width           =   780
         End
      End
      Begin VB.Frame fra_Producao 
         Height          =   2235
         Left            =   150
         TabIndex        =   13
         Top             =   1260
         Width           =   7845
         Begin VB.TextBox txt_ValorTotalEstimado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6585
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   4290
            Width           =   1290
         End
         Begin TrueOleDBGrid70.TDBDropDown tdd_Valores 
            Height          =   1515
            Left            =   4230
            TabIndex        =   15
            Top             =   720
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   2672
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   "strDescricao"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Faixa Inicial"
            Columns(1).DataField=   "dblFaixaInicial"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Faixa Final"
            Columns(2).DataField=   "dblFaixaFinal"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2461"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2381"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=2858"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2778"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   4
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   ""
            DataField       =   ""
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
            ValueTranslate  =   0   'False
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
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
            _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=39:EvenRow"
            _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(57)  =   "Named:id=40:OddRow"
            _StyleDefs(58)  =   ":id=40,.parent=33"
            _StyleDefs(59)  =   "Named:id=41:RecordSelector"
            _StyleDefs(60)  =   ":id=41,.parent=34"
            _StyleDefs(61)  =   "Named:id=42:FilterBar"
            _StyleDefs(62)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBDropDown tdd_Faixa 
            Height          =   1515
            Left            =   2850
            TabIndex        =   16
            Top             =   720
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   2672
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
            Columns(1).Caption=   "Faixa de Valor"
            Columns(1).DataField=   "strNomeDaFaixa"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   4
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   "strNomeDaFaixa"
            DataField       =   "PKId"
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
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
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
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
         Begin TrueOleDBGrid70.TDBDropDown tdd_Receita 
            Height          =   1485
            Left            =   390
            TabIndex        =   17
            Top             =   720
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   2619
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
            Columns(1).Caption=   "Composição da Receita"
            Columns(1).DataField=   "strDescricao"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   4
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   "strDescricao"
            DataField       =   "PKId"
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   -2147483644
            ValueTranslate  =   -1  'True
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Composicao 
            Height          =   1815
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   3201
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   66
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Composição da Receita"
            Columns(1).DataField=   ""
            Columns(1).DataWidth=   40
            Columns(1).DropDown=   "tdd_Receita"
            Columns(1).DropDown.vt=   8
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Faixa de Valor"
            Columns(2).DataField=   ""
            Columns(2).DataWidth=   30
            Columns(2).DropDown=   "tdd_Faixa"
            Columns(2).DropDown.vt=   8
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Faixa Inicial"
            Columns(3).DataField=   ""
            Columns(3).DataWidth=   10
            Columns(3).NumberFormat=   "Standard"
            Columns(3).EditMaskUpdate=   -1  'True
            Columns(3).EditMaskRight=   -1  'True
            Columns(3).DropDown=   "tdd_Valores"
            Columns(3).DropDown.vt=   8
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Faixa Final"
            Columns(4).DataField=   ""
            Columns(4).DataWidth=   10
            Columns(4).NumberFormat=   "Standard"
            Columns(4).EditMaskUpdate=   -1  'True
            Columns(4).EditMaskRight=   -1  'True
            Columns(4).DropDown=   "tdd_Valores"
            Columns(4).DropDown.vt=   8
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Quantidade"
            Columns(5).DataField=   ""
            Columns(5).DataWidth=   12
            Columns(5).DropDown=   "tdd_UnidadeMedida"
            Columns(5).DropDown.vt=   8
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerColor=   11059392
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=476"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=397"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=4710"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=4630"
            Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=260"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(1).AutoDropDown=1"
            Splits(0)._ColumnProps(16)=   "Column(1).AutoCompletion=1"
            Splits(0)._ColumnProps(17)=   "Column(2).Width=2408"
            Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2328"
            Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(22)=   "Column(3).Width=1931"
            Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1852"
            Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=260"
            Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(28)=   "Column(4).Width=1879"
            Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=1799"
            Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=260"
            Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(34)=   "Column(5).Width=1773"
            Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=1693"
            Splits(0)._ColumnProps(37)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=260"
            Splits(0)._ColumnProps(39)=   "Column(5).Order=6"
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
            DeadAreaBackColor=   -2147483644
            RowDividerColor =   11059392
            RowSubDividerColor=   11059392
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
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=17,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=26,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=18,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=19,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=20,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=22,.parent=6"
            _StyleDefs(26)  =   "Splits(0).EditorStyle:id=21,.parent=7"
            _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=23,.parent=8"
            _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=24,.parent=9"
            _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=25,.parent=10"
            _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=27,.parent=11"
            _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=28,.parent=12"
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=16,.parent=17"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=18"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=19"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=21"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=17"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=18,.alignment=0"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=19"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=21"
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=18"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=19"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=21"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=18,.alignment=0"
            _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=19"
            _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=21"
            _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=17"
            _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=18,.alignment=0"
            _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=19"
            _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=21"
            _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=58,.parent=17"
            _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=18,.alignment=0"
            _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=19"
            _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=21"
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
         Begin VB.Label lbl_ValorTotal 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total Estimado"
            Height          =   195
            Left            =   5040
            TabIndex        =   18
            Top             =   4380
            Width           =   1455
         End
      End
      Begin MSDataListLib.DataCombo dbcintContribuinte 
         Height          =   315
         Left            =   2400
         TabIndex        =   0
         Top             =   480
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Geral 
         Height          =   1305
         Left            =   180
         TabIndex        =   11
         Top             =   5670
         Width           =   7815
         _ExtentX        =   13785
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "intCodigoGeral"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nome"
         Columns(2).DataField=   "strNome"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2619"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=2646"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2566"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=10583"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=10504"
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
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.locked=0"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbc_intOcorrencia 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   870
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lbl_intOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   570
         TabIndex        =   41
         Top             =   990
         Width           =   780
      End
      Begin VB.Label lblintContribuinte 
         AutoSize        =   -1  'True
         Caption         =   "Contribuinte"
         Height          =   195
         Left            =   510
         TabIndex        =   27
         Top             =   600
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmCadReceitasDiversas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando               As Boolean
Dim mobjAux                     As Object
Dim mblnSelecionou              As Boolean
Dim mblnPrimeiraVez             As Boolean
Dim adoRec                      As ADODB.Recordset
Dim adoTdb                      As ADODB.Recordset

Dim X                           As XArrayDB
Dim Y                           As XArrayDB
Dim Z                           As XArrayDB
Dim A                           As XArrayDB
Dim xarReceita                  As XArrayDB

Private Sub cmd_Mensagem_Click()
   ChamaFormCadastro frmCadMensagem, dbcintMensagem
End Sub

Private Sub dbc_intMensagem1_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intMensagem1, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intMensagem2_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intMensagem2, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intOcorrencia_Click(Area As Integer)
    DropDownDataCombo dbc_intOcorrencia, Me, Area
End Sub

Private Sub dbc_intOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintContribuinte, Me, , KeyCode, Shift
   If KeyCode = vbKeyReturn Then
      If dbcintContribuinte.MatchedWithList Then
          txtintCodigoGeral = Val(dbcintContribuinte.BoundText)
      End If
   End If
End Sub

Private Sub dbcintMensagem_Click(Area As Integer)
    DropDownDataCombo dbcintMensagem, Me, Area
    If Area = 2 Then
       If dbcintMensagem.BoundText <> "" Then
          strQuerryMensagem
       End If
    End If
End Sub

Private Function strQuerry() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT RD.PKId, RD.intCodigoGeral, CO.strNome "
    strSql = strSql & " FROM "
    strSql = strSql & gstrReceitaDiversa & " RD,"
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE RD.intContribuinte = CO.PKId "
    strSql = strSql & " ORDER BY RD.intCodigoGeral "
strQuerry = strSql
End Function
Private Function strQuerryMensagem()
    Dim strSql As String
    Dim adoResultado       As ADODB.Recordset
    strSql = "SELECT PKId, strMensagem " & _
             " FROM " & gstrMensagem & _
             " WHERE PKId = " & dbcintMensagem.BoundText
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                txt_Mensagem = !strMensagem
                .MoveNext
            End If
        End With
    End If
End Function

Private Function strQueryInscricao() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strInscricaoCadastral "
    strSql = strSql & " FROM "
    strSql = strSql & gstrEconomico
    strSql = strSql & " WHERE "
    strSql = strSql & " dtmDataBaixa IS NULL " 'Verifica se existe data de baixa
    strSql = strSql & " ORDER BY "
'    strSql = strSql & " CONVERT(NUMERIC,strInscricaoCadastral) "
    strSql = strSql & gstrCONVERT(CDT_NUMERIC, "strInscricaoCadastral")
strQueryInscricao = strSql
End Function

Private Sub dbcintMensagem_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintMensagem, Me, , KeyCode, Shift
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    gintCodSeguranca = 674
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Activate()
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
    mblnAlterando = False
    dbcintContribuinte.Tag = "SELECT PKId, strNome FROM " & gstrContribuinte & " ORDER BY strNome;strNome"
    LeDaTabelaParaObj gstrMensagem, dbcintMensagem, "PKId, strDescricao"
    LeDaTabelaParaObj gstrReceitaDiversa, tdb_Geral, strQuerry
    PreencheGRD2
    VerificaObjParaAplicar mobjAux
    txtintCodigoGeral.Enabled = False
    TrocaCorObjeto txtintCodigoGeral, True
    
    LeDaTabelaParaObj gstrOcorrencia, dbc_intOcorrencia, strQuerryOcorrencia
    
'''GUIA
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem1, strQueryMensagem
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem2, strQueryMensagem
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    Select Case tab_3DPasta.Tab
        Case 0
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
        Case 1
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
        Case 2
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrNovo
    End Select
End Sub

Private Sub tdb_Atividades_AfterColUpdate(ByVal ColIndex As Integer)
tdb_Atividades.Update
End Sub

Private Sub tdb_Atividades_AfterUpdate()
tdb_Atividades.Update
End Sub

Private Sub tdb_Composicao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    MontaAtividade Val(tdb_Composicao.Columns(1).Value)
End Sub

Private Sub tdb_Geral_Click()
    mblnPrimeiraVez = True
    With tdb_Geral
        If Not .EOF And Not .BOF Then
'            If .Bookmark = 1 Then
'                tdb_Geral_RowColChange 0, 0
'            End If
        End If
    End With
End Sub

Sub tdb_Geral_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Geral_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Geral
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKId.Text = .Columns("PKId").Value

                LeDaTabelaParaObj gstrReceitaDiversa, Me
                PreencheGRD2
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
                dbcintMensagem_Click 2

                gCorLinhaSelecionada tdb_Geral

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
'                mblnPrimeiraVez = False
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
            End If

        End If
    End With
End Sub

'==Para Efetuar Cálculo==
Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    Dim i As Integer
    If (txt_intParcelaInicial.Text = "") Or (txt_intParcelaFinal.Text = "") Then
        ExibeMensagem "O intervalo de parcelas deve ser definido "
        tab_3DPasta.Tab = 0
        If txt_intParcelaInicial.Text = "" Then
            txt_intParcelaInicial.SetFocus
        Else
            txt_intParcelaFinal.SetFocus
        End If
        Exit Function
    End If
    If Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal) Then
        ExibeMensagem "O número da parcela final deve ser maior que o número da parcela inicial"
        tab_3DPasta.Tab = 0
        txt_intParcelaFinal.SetFocus
        Exit Function
    End If
    If dbc_intOcorrencia.MatchedWithList = False Then
        ExibeMensagem "O campo Ocorrência não pode ser nulo."
        tab_3DPasta.Tab = 0
        dbc_intOcorrencia.SetFocus
        Exit Function
    End If
    If txtintIntervalo.Text = "" Then
        ExibeMensagem "O campo Intervalo não pode ser nulo."
        tab_3DPasta.Tab = 0
        txtintIntervalo.SetFocus
        Exit Function
    End If
    
    If txtdtmVencimento.Text = "" Then
        ExibeMensagem "O campo data de vencimento não pode ser nulo."
        tab_3DPasta.Tab = 0
        txtdtmVencimento.SetFocus
        Exit Function
    Else
        If gblnDataValida(txtdtmVencimento.Text) = False Then
            ExibeMensagem "A data de vencimento não é válida."
            tab_3DPasta.Tab = 0
            txtdtmVencimento.SetFocus
            Exit Function
        End If
    End If
    
    If txtdtmLancamento.Text = "" Then
        ExibeMensagem "O campo data de lançamento não pode ser nulo."
        tab_3DPasta.Tab = 0
        txtdtmLancamento.SetFocus
        Exit Function
    Else
        If gblnDataValida(txtdtmLancamento.Text) = False Then
            ExibeMensagem "A data de lançamento não é válida."
            tab_3DPasta.Tab = 0
            txtdtmLancamento.SetFocus
            Exit Function
        End If
    End If
    
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 Then
            blnDadosOk = True
            Exit Function
        End If
    Next
    ExibeMensagem "Selecione uma receita para efetuar o cálculo."
    tab_3DPasta.Tab = 1
End Function

Private Function strPKId() As String
    Dim strSql As String
    Dim i As Integer
    strSql = ""
    For i = 0 To xarReceita.Count(1) - 1

        If xarReceita(i, 2) = -1 Then
            If strSql <> "" Then
               strSql = strSql & ","
            End If
        strSql = strSql & xarReceita(i, 0)
        End If
    Next
   
    strPKId = strSql
End Function

Private Function strQuerryOcorrencia() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrOcorrencia
    strSql = strSql & " WHERE "
    strSql = strSql & " intUtilizacaoDaOcorrencia = 1 "
    strSql = strSql & " ORDER BY strDescricao "
strQuerryOcorrencia = strSql
End Function

Private Function strQuerryRelatorioGuiaDeArrecadacao(ByRef blnExiteLancamento As Boolean) As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    strSql = ""
    
    strSql = strSql & " SELECT DISTINCT I.PKId, "
    strSql = strSql & " I.intExercicio, J.intNumeroParcela, I.strInscricaoCadastral, I.intComposicaoReceita CodReceita,"
    strSql = strSql & " J.dtmDataVencimento, G.strDescricao Municipio, H.strDescricao Bairro,C.strNome Contribuinte,"
    strSql = strSql & " C.intNumero, C.strComplemento, C.intCep,"
'    strSql = strSql & " ltrim(rtrim(isnull(E.strSigla,''))) + ' ' + ltrim(rtrim(isnull(F.strSigla,''))) + ' ' + ltrim(rtrim(D.strDescricao)) AS Logradouro "
    strSql = strSql & " ltrim(rtrim(" & gstrISNULL("E.strSigla", "''") & ")) " & strCONCAT & " ' ' " & strCONCAT & " ltrim(rtrim(" & gstrISNULL("F.strSigla", "''") & ")) " & strCONCAT & " ' ' " & strCONCAT & " ltrim(rtrim(D.strDescricao)) AS Logradouro "

    strSql = strSql & " FROM "
    strSql = strSql & gstrContribuinte & " C,"
    strSql = strSql & gstrLogradouro & " D,"
    strSql = strSql & gstrTipoLogradouro & " E,"
    strSql = strSql & gstrTituloLogradouro & " F,"
    strSql = strSql & gstrCidade & " G,"
    strSql = strSql & gstrBairro & " H, "
    strSql = strSql & gstrLancamentoCalculo & " I, "
    strSql = strSql & gstrParcelaTaxa & " J "

    strSql = strSql & " WHERE "
    strSql = strSql & " J.intLancamentoCalculo = I.PKId "
    strSql = strSql & " AND I.bitUtilizacaoDebito = 4 "
    strSql = strSql & " AND I.intContribuinte      = C.PKId "
'    strSql = strSql & " AND C.intMunicipio         *= G.PKId "
    strSql = strSql & " AND C.intMunicipio         " & strOUTJSQLServer & "= G.PKId " & strOUTJOracle
'    strSql = strSql & " AND C.intBairro            *= H.PKId "
    strSql = strSql & " AND C.intBairro            " & strOUTJSQLServer & "= H.PKId " & strOUTJOracle
    strSql = strSql & " AND C.intLogradouro        = D.PKId "
'    strSql = strSql & " AND D.intTipoLogradouro    *= E.PKId "
    strSql = strSql & " AND D.intTipoLogradouro    " & strOUTJSQLServer & "= E.PKId " & strOUTJOracle
'    strSql = strSql & " AND D.intTituloLogradouro  *= F.PKId "
    strSql = strSql & " AND D.intTituloLogradouro  " & strOUTJSQLServer & "= F.PKId " & strOUTJOracle
    strSql = strSql & " AND I.intContribuinte = " & txtintCodigoGeral.Text
    strSql = strSql & " AND I.intExercicio = " & Val(txt_intExercicio)
    strSql = strSql & " AND J.intNumeroParcela BETWEEN " & Val(txt_intParcelaInicial) & " AND " & Val(txt_intParcelaFinal)

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            blnExiteLancamento = True
        Else
            ExibeMensagem "Não Existe(m) Lançamento(s) para o Contribuinte Selecionado"
            blnExiteLancamento = False
        End If
    End If

strQuerryRelatorioGuiaDeArrecadacao = strSql
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim varBookMark As Variant
Dim strSql As String
Dim AlterandoAux As Boolean
Dim txt_PKId As Integer
Dim i        As Integer
Dim blnExiteLancamento As Boolean

    If tab_3DPasta.Tab = 2 Then
        If UCase(strModoOperacao) = UCase(gstrImprimir) Then
            If blnDadosGuiaOK Then
                strSql = strQuerryRelatorioGuiaDeArrecadacao(blnExiteLancamento)
                If blnExiteLancamento Then
                    Set gfrmFormularioQueEstaImprimindoGuia = Me
    '                rptGuiaDeArrecadacaoMunicipal.strImposto = dbc_intComposicaoDaReceita.Text
                    ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, strSql
                End If
            End If
        End If
        If UCase(strModoOperacao) = UCase(gstrNovo) Then
            LimpaObjetos
        End If
        If UCase(strModoOperacao) = UCase(gstrFechar) Then
            Unload Me
        End If
        
    Else
        If UCase(strModoOperacao) = gstrPreencherLista Then
            PreencherListaDeOpcoes Me.ActiveControl
            Exit Sub
        End If
        AlterandoAux = False
        txt_PKId = 0
        If mblnAlterando Then
            AlterandoAux = True
            txt_PKId = Val(txtPKId)
        End If
        
        If Not tdb_Geral.EOF Then
            varBookMark = tdb_Geral.Bookmark
        Else
            mblnAlterando = False
        End If
        
        strSql = strQuerry
        If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
            mblnPrimeiraVez = False
        End If
    
    '''
        txtintNumero.Text = Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1
        If strModoOperacao = gstrCalcularReajuste Then
            CalculoLancamento
            Exit Sub
        ElseIf ToolBarGeral(strModoOperacao, gstrReceitaDiversa, mblnAlterando, tdb_Geral, Me, mobjAux, strSql, strSql) Then
            If AlterandoAux = True Then
                If UCase(strModoOperacao) = UCase(gstrSalvar) Then
                    If GravaValores2(Val(txt_PKId), strModoOperacao) Then
                
                    End If
                End If
                If UCase(strModoOperacao) = UCase(gstrDeletar) Then
                    If DeletaValores2(Val(txt_PKId), strModoOperacao) Then
                    End If
                End If
            Else
                If UCase(strModoOperacao) = UCase(gstrSalvar) Then
                    If GravaValores2(PegaMaxPKId, strModoOperacao) Then
                        
                    End If
                End If
                If UCase(strModoOperacao) = UCase(gstrDeletar) Then
                    If DeletaValores2(PegaMaxPKId, strModoOperacao) Then
                    End If
                End If
            End If
        End If
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
        If UCase(strModoOperacao) = UCase(gstrNovo) Then
            LimpaGrid2
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
        End If
        If UCase(strModoOperacao) <> UCase(gstrFechar) And (Not tdb_Geral.EOF And Not tdb_Geral.BOF) Then
            If Not IsEmpty(varBookMark) Then
                If UCase(strModoOperacao) = UCase(gstrDeletar) Then
                    tdb_Geral.MoveFirst
                Else
                    tdb_Geral.Bookmark = varBookMark
                End If
            End If
        End If
        PreencheGRD2
    End If
    
End Sub


'====================

Private Sub CalculoLancamento()

'******************************************************************************************
' Data: 08/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 08/05/2003
' Alteração: - Substituição da chamada à função CriaADO por uma chamada à função
'            ExecuteStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql                   As String
'    Dim adoResultado             As ADODB.Recordset
    Dim adoParameters            As ADODB.Parameters
    Dim intNumeroParcelas        As Integer
    Screen.MousePointer = vbHourglass
    If blnDadosOk Then
'        strSQL = "sp_CalculoParaUsuario '" & dbcintContribuinte.BoundText & "','" & strPKId & "',4," & _
'                 tdd_Receita.Columns(0).Value
        strSql = gstrStoredProcedure("sp_CalculoParaUsuario", "'" & dbcintContribuinte.BoundText & "','" & strPKId & "',4," & _
                 tdd_Receita.Columns(0).Value & ", 0, ' :V_dblValor'")
        Set gobjBanco = New clsBanco
'        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If gobjBanco.ExecuteStoredProcedure(strSql, 10, , adoParameters) Then
'            If Not (adoResultado.BOF And adoResultado.EOF) Then
            If Not (adoParameters Is Nothing) Then
                'Mostra Valores para Usuário
'                strSQL = "Confirma o cálculo de " & gstrConvVrDoSql(adoResultado!dblValorAparcelar) & _
'                        Chr(10) & " + " & gstrConvVrDoSql(adoResultado!dblValorNaoParcelado) & " em " & _
'                        (Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1) & " parcela(s) ?"
                strSql = "Confirma o cálculo de " & gstrConvVrDoSql(adoParameters("dblValorAparcelar")) & _
                        Chr(10) & " + " & gstrConvVrDoSql(adoParameters("dblValorNaoParcelado")) & " em " & _
                        (Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1) & " parcela(s) ?"
                Screen.MousePointer = vbNormal
                If MsgBox(strSql, vbYesNo, "Tributário") = vbYes Then
                    Screen.MousePointer = vbHourglass
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaBeginTrans
                    If Not gBlnVerificaLancamentos(txt_intExercicio, tdd_Receita.Columns(0).Value, _
                                                    tdd_Receita.Columns(1).Value, (Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1), _
                                                    gstrConvDtParaSql(txtdtmLancamento), 0, dbcintContribuinte.BoundText) Then
                        gobjBanco.ExecutaRollbackTrans
                        Screen.MousePointer = vbNormal
                        Exit Sub
                    End If
'                    strSql = strLancamentos(adoResultado!dblValorAparcelar, adoResultado!dblValorNaoParcelado)
                    strSql = strLancamentos(adoParameters("dblValorAparcelar"), adoParameters("dblValorNaoParcelado"))
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaBeginTrans
                    'Executa o Lançamento
                    If gobjBanco.Execute(strSql, False) Then
                        gobjBanco.ExecutaCommitTrans
                        Screen.MousePointer = vbNormal
                        ExibeMensagem "Cálculo efetuado com sucesso!"
                    Else
                        gobjBanco.ExecutaRollbackTrans
                    End If
                End If
            End If
        End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Function strContribuintes() As String
    Dim strSql          As String
    strSql = " SELECT " & dbcintContribuinte.BoundText & "," & dbcintContribuinte.BoundText
    strContribuintes = strSql
End Function

Private Function strLancamentos(dblValorAparcelar As String, dblValorNaoParcelado As String) As String

'******************************************************************************************
' Data: 09/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
    
    Dim strParameters As String
    
'    strSql = " sp_CalculoLancamentoReceitas 1, " & gstrConvVrParaSql(dblValorAparcelar) & " , "
    strParameters = "1, " & gstrConvVrParaSql(dblValorAparcelar) & " , " & _
            gstrConvVrParaSql(dblValorNaoParcelado) & ",'" & strPKId & "','" & strContribuintes & _
            "'," & Val(txt_intExercicio) & "," & tdd_Receita.Columns(0).Value
    
    If dbcintMensagem.MatchedWithList Then
'        strSql = strSql & ", " & dbcintMensagem.BoundText & ","
        strParameters = strParameters & ", " & dbcintMensagem.BoundText & ","
    Else
'        strSql = strSql & ", NULL,"
        strParameters = strParameters & ", NULL,"
    End If
    
'    strSql = strSql & gstrConvDtParaSql(txtdtmLancamento.Text)
    strParameters = strParameters & gstrConvDtParaSql(txtdtmLancamento.Text) & _
            "," & gstrConvDtParaSql(txtdtmVencimento.Text) & "," & _
            txt_intParcelaInicial.Text & "," & txt_intParcelaFinal.Text & "," & _
            Val(txtintIntervalo.Text) & ",4," & Val(dbc_intOcorrencia.BoundText) & _
            ",4," & glngCodUsr
    
    strSql = gstrStoredProcedure("sp_CalculoLancamentoReceitas", strParameters)
    
    strLancamentos = strSql
End Function
'====================

Function PreencheGRD2()
Dim strSql As String
Dim varAux  As Variant

LimpaGrid2
'' Receita

        strSql = ""
        strSql = strSql & " SELECT PKId, strDescricao "
        strSql = strSql & " FROM " & gstrComposicaoDaReceita
        strSql = strSql & " WHERE intUtilizacao = 5 "
        strSql = strSql & " ORDER BY strDescricao "

        Set gobjBanco = New clsBanco
        gobjBanco.CriaADO strSql, 5, adoTdb
    If Not adoTdb.EOF Then
        Y.ReDim 0, adoTdb.RecordCount - 1, 0, 1
        Do While Not adoTdb.EOF
            varAux = adoTdb!Pkid
            Y(adoTdb.AbsolutePosition - 1, 0) = varAux
            varAux = adoTdb!strDescricao
            Y(adoTdb.AbsolutePosition - 1, 1) = varAux

            adoTdb.MoveNext
        Loop
    End If
        Set tdd_Receita.Array = Y
        tdd_Receita.Rebind
        tdd_Receita.Refresh

' DROP Faixa
        
        strSql = ""
        strSql = strSql & "SELECT PKId, strNomeDaFaixa "
        strSql = strSql & " FROM " & gstrFaixaDeValor
        strSql = strSql & " ORDER BY strNomeDaFaixa "

        Set gobjBanco = New clsBanco
        gobjBanco.CriaADO strSql, 5, adoTdb
    If Not adoTdb.EOF Then
        A.ReDim 0, adoTdb.RecordCount - 1, 0, 1
        Do While Not adoTdb.EOF

            varAux = adoTdb!Pkid
            A(adoTdb.AbsolutePosition - 1, 0) = varAux
            varAux = adoTdb!strNomeDaFaixa
            A(adoTdb.AbsolutePosition - 1, 1) = varAux
            
            adoTdb.MoveNext
        Loop
    End If
        Set tdd_Faixa.Array = A
        tdd_Faixa.Rebind
        tdd_Faixa.Refresh

'GRID Composicao
        strSql = ""
        strSql = strSql & "SELECT PKId, intComposicaoDaReceita, intFaixaDeValor, dblFaixaInicial, "
        strSql = strSql & " dblFaixaFinal, intQuantidade "
        strSql = strSql & " FROM " & gstrReceitaDiversaValor
        If mblnAlterando = True Then
            strSql = strSql & " WHERE intReceitaDiversa = " & Val(txtPKId)
        End If
        Set gobjBanco = New clsBanco
        gobjBanco.CriaADO strSql, 5, adoRec
        MontaArray2
        

End Function

Private Sub tdb_Composicao_KeyPress(KeyAscii As Integer)
    Select Case tdb_Composicao.Col
         Case 2
            CaracterValido KeyAscii, "N", tdb_Composicao.Columns(2)
        Case 3
            CaracterValido KeyAscii, "N", tdb_Composicao.Columns(3)
        Case 4
            CaracterValido KeyAscii, "V", tdb_Composicao.Columns(4)
    End Select
End Sub


Private Sub MontaArray2()
    Dim varAux As Variant

    Set X = New XArrayDB
    X.Clear
    With adoRec
        If Not .EOF And mblnAlterando Then
            X.ReDim 0, .RecordCount - 1, 0, 5
            Do While Not .EOF
                varAux = .Fields(0)
                X(.AbsolutePosition - 1, 0) = varAux
                varAux = .Fields(1)
                X(.AbsolutePosition - 1, 1) = varAux
                varAux = .Fields(2)
                X(.AbsolutePosition - 1, 2) = varAux
                varAux = .Fields(3)
                X(.AbsolutePosition - 1, 3) = varAux
                varAux = .Fields(4)
                X(.AbsolutePosition - 1, 4) = varAux
                varAux = .Fields(5)
                X(.AbsolutePosition - 1, 5) = varAux
                .MoveNext
            Loop
        Else
            X.ReDim 0, 0, 0, 5
            X(0, 0) = ""
            X(0, 1) = ""
            X(0, 2) = ""
            X(0, 3) = ""
            X(0, 4) = ""
            X(0, 5) = ""
        End If
    End With

    Set tdb_Composicao.Array = X
    tdb_Composicao.Rebind
    tdb_Composicao.Refresh
End Sub

Private Sub LimpaGrid2()
    Set X = New XArrayDB
    Set Y = New XArrayDB
    Set Z = New XArrayDB
    Set A = New XArrayDB

    X.Clear
    Y.Clear
    Z.Clear
    A.Clear

    Set tdb_Composicao.Array = X
    tdb_Composicao.Rebind
    tdb_Composicao.Refresh

    Set tdd_Receita.Array = Y
    tdd_Receita.Rebind
    tdd_Receita.Refresh

    Set tdd_Faixa.Array = Z
    tdd_Faixa.Rebind
    tdd_Faixa.Refresh

    Set tdd_Valores.Array = A
    tdd_Valores.Rebind
    tdd_Valores.Refresh
    
    txt_Mensagem = ""
    
End Sub

Private Sub tdb_Geral_FilterChange()
    gblnFilraCampos tdb_Geral
End Sub

Private Sub MontaAtividade(intComposicaoReceita As Integer)
Dim strSql As String
Dim adoRec As ADODB.Recordset
Dim varAux As String

On Error GoTo Err_Handle

Set xarReceita = New XArrayDB
xarReceita.Clear

xarReceita.ReDim 0, 0, 0, 2

strSql = ""
strSql = strSql & " SELECT A.PKId, A.strDescricao FROM "
strSql = strSql & gstrReceita & " A,"
strSql = strSql & gstrValorCompoRec & " B"
strSql = strSql & " WHERE A.PKId = B.intReceita "
strSql = strSql & " AND B.intComposicaoDaReceita = " & intComposicaoReceita

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    With adoRec
        If Not .EOF Then
            xarReceita.ReDim 0, .RecordCount - 1, 0, 2
            Do While Not .EOF
                varAux = !Pkid
                xarReceita(.AbsolutePosition - 1, 0) = varAux
                
                varAux = False
                xarReceita(.AbsolutePosition - 1, 1) = varAux
            
                varAux = !strDescricao
                xarReceita(.AbsolutePosition - 1, 2) = varAux
                
                .MoveNext
            Loop
        End If
    End With
End If

Set tdb_Atividades.Array = xarReceita
tdb_Atividades.Rebind
tdb_Atividades.Refresh

Exit Sub
Err_Handle:

End Sub

Private Sub tdd_Receita_DropDownClose()
    MontaAtividade tdd_Receita.Columns(0).Value
End Sub

Private Function DeletaValores2(intCodImobiliario As Integer, strOperacao As String) As Boolean
    Dim strSql As String
    If strOperacao = "DELETAR" Then
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrReceitaDiversaValor & " "
        strSql = strSql & "WHERE  intReceitaDiversa = " & intCodImobiliario
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSql

        LimpaGrid2
    
    End If
End Function

Private Function GravaValores2(intCodImobiliario As Integer, strOperacao As String) As Boolean
    Dim strSql As String
    Dim strMsg As String
    Dim i      As Integer

On Error GoTo err_GravaValores2
    If strOperacao = "SALVAR" Then
    
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
    
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrReceitaDiversaValor & " "
        strSql = strSql & "WHERE  intReceitaDiversa = " & intCodImobiliario
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSql
    
        tdb_Composicao.MoveFirst
         
        For i = 0 To X.Count(1) - 1
            
            strSql = ""
            strSql = strSql & "INSERT INTO " & gstrReceitaDiversaValor & " "
            strSql = strSql & "(intReceitaDiversa, intComposicaoDaReceita, intFaixaDeValor, "
            strSql = strSql & " dblFaixaInicial, dblFaixaFinal, intQuantidade "
            strSql = strSql & ") Values ("
            strSql = strSql & intCodImobiliario & ", "
            strSql = strSql & X(i, 1) & ", "
            strSql = strSql & X(i, 2) & ", "
            strSql = strSql & gstrConvVrParaSql(X(i, 3)) & ", "
            strSql = strSql & gstrConvVrParaSql(X(i, 4)) & ", "
            strSql = strSql & Val(X(i, 5)) & " "
            strSql = strSql & ")"
    
            If Not gobjBanco.Execute(strSql, False) Then
                gobjBanco.ExecutaRollbackTrans
            End If
        Next i
    End If
    gobjBanco.ExecutaCommitTrans
    LimpaGrid2

Exit Function
err_GravaValores2:
    gobjBanco.ExecutaRollbackTrans
End Function


Private Sub tdd_Faixa_DropDownClose()
Dim PPKKid As Integer
Dim strSql As String
PPKKid = 0
    Dim intRow As Integer
    On Error GoTo Err_Handle
    If Not IsNull(tdd_Faixa.SelectedItem) Or Not IsEmpty(tdd_Faixa.SelectedItem) Then
        tdb_Composicao.Columns(2) = tdd_Faixa.Columns(1)
        PPKKid = Val(tdd_Faixa.Columns(0))
    Else
        tdb_Composicao.Columns(2) = ""
    End If
        
        If PPKKid <> 0 Then
            strSql = ""
            strSql = strSql & "SELECT PKId, dblFaixaInicial, dblFaixaFinal "
            strSql = strSql & "FROM " & gstrValorDaFaixa
            strSql = strSql & " WHERE intFaixaDeValores = " & PPKKid
            strSql = strSql & " AND intUtilizacao BETWEEN 12 AND 13 "
            strSql = strSql & " ORDER BY PKId "
    
            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strSql, 5, adoTdb
            If Not adoTdb.EOF Then
                Z.ReDim 0, adoTdb.RecordCount - 1, 0, 2
                Dim varAux2   As Variant
                Dim varAux21  As Variant
                Dim varAux211 As Variant
                Do While Not adoTdb.EOF
        
                    varAux2 = adoTdb!Pkid
                    varAux21 = adoTdb!dblFaixaInicial
                    varAux211 = adoTdb!dblFaixaFinal
                    Z(adoTdb.AbsolutePosition - 1, 0) = varAux2
                    Z(adoTdb.AbsolutePosition - 1, 1) = varAux21
                    Z(adoTdb.AbsolutePosition - 1, 2) = varAux211
        
                    adoTdb.MoveNext
                Loop
            End If
            Set tdd_Valores.Array = Z
            tdd_Valores.Rebind
            tdd_Valores.Refresh
        End If
    Exit Sub
Err_Handle:
End Sub

Private Sub tdd_Valores_DropDownClose()
    Dim intRow As Integer
    On Error GoTo Err_Handle
    If Not IsNull(tdd_Valores.SelectedItem) Or Not IsEmpty(tdd_Valores.SelectedItem) Then
        tdb_Composicao.Columns(3) = tdd_Valores.Columns(1)
        tdb_Composicao.Columns(4) = tdd_Valores.Columns(2)
    Else
        tdb_Composicao.Columns(3) = ""
        tdb_Composicao.Columns(4) = ""
    End If

    Exit Sub
Err_Handle:
End Sub

Private Sub dbcintContribuinte_Click(Area As Integer)
    DropDownDataCombo dbcintContribuinte, Me, Area
    If Area = 2 Then
        txtintCodigoGeral = Val(dbcintContribuinte.BoundText)
    End If
End Sub

Private Sub txtdtmLancamento_GotFocus()
    MarcaCampo txtdtmLancamento
End Sub

Private Sub txtdtmLancamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmLancamento
End Sub

Private Sub txtdtmLancamento_LostFocus()
    txtdtmLancamento = gstrDataFormatada(txtdtmLancamento)
End Sub

Private Sub txtdtmVencimento_GotFocus()
    MarcaCampo txtdtmVencimento
End Sub

Private Sub txtdtmVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmVencimento
End Sub

Private Sub txtdtmVencimento_LostFocus()
    txtdtmVencimento = gstrDataFormatada(txtdtmVencimento)
End Sub

Private Sub txtintIntervalo_GotFocus()
    MarcaCampo txtintIntervalo
End Sub

Private Sub txtintIntervalo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintIntervalo
End Sub

Function PegaMaxPKId() As Integer
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    strSql = "SELECT MAX(PKId) as PKId " & _
             " FROM " & gstrReceitaDiversa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
         PegaMaxPKId = adoResultado!Pkid
    End If
End Function





'''>>>>>>>>>>>>>>>>>>GUIA DE ARRECADAÇÃO




Private Sub chk_EmBranco1_Click()
    If chk_EmBranco1.Value = 1 Then
        dbc_intMensagem1.BoundText = ""
        dbc_intMensagem1.Enabled = False
        TrocaCorObjeto dbc_intMensagem1, True
        txt_Mensagem1.Text = ""
        txt_Mensagem1.Enabled = False
        TrocaCorObjeto txt_Mensagem1, True
    Else
        dbc_intMensagem1.Enabled = True
        TrocaCorObjeto dbc_intMensagem1, False
        txt_Mensagem1.Enabled = True
        TrocaCorObjeto txt_Mensagem1, False
    End If
End Sub

Private Sub chk_EmBranco2_Click()
    If chk_EmBranco2.Value = 1 Then
        dbc_intMensagem2.BoundText = ""
        dbc_intMensagem2.Enabled = False
        TrocaCorObjeto dbc_intMensagem2, True
        txt_Mensagem2.Text = ""
        txt_Mensagem2.Enabled = False
        TrocaCorObjeto txt_Mensagem2, True
    Else
        dbc_intMensagem2.Enabled = True
        TrocaCorObjeto dbc_intMensagem2, False
        txt_Mensagem2.Enabled = True
        TrocaCorObjeto txt_Mensagem2, False
    End If
End Sub

Private Sub dbc_intMensagem1_Click(Area As Integer)
    DropDownDataCombo dbc_intMensagem1, Me, Area
    If Area = 2 Then
        LeDoComboParaTXT1
    End If
End Sub

Private Sub dbc_intMensagem2_Click(Area As Integer)
    DropDownDataCombo dbc_intMensagem2, Me, Area
    If Area = 2 Then
        LeDoComboParaTXT2
    End If
End Sub

Private Function LeDoComboParaTXT1()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    strSql = " SELECT strMensagem " & _
             " FROM " & gstrMensagem & _
             " WHERE PKId = " & Val(dbc_intMensagem1.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem1.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem1.Text = ""
        End If
    End If
End Function

Private Function LeDoComboParaTXT2()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    strSql = " SELECT strMensagem " & _
             " FROM " & gstrMensagem & _
             " WHERE PKId = " & Val(dbc_intMensagem2.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem2.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem2.Text = ""
        End If
    End If
End Function

Private Function strQueryMensagem() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSql As String

    strSql = ""
'    strSQL = strSQL & "SELECT PKId, ltrim(rtrim(PKId)) + ' - ' + ltrim(rtrim(strDescricao)) as Descricao "
    strSql = strSql & "SELECT PKId, ltrim(rtrim(PKId)) " & strCONCAT & " ' - ' " & strCONCAT & " ltrim(rtrim(strDescricao)) as Descricao "
    strSql = strSql & " FROM " & gstrMensagem
    strSql = strSql & " ORDER BY PKId "

strQueryMensagem = strSql
End Function

Private Function blnDadosGuiaOK() As Boolean
    If (txt_intParcelaInicial.Text = "") Or (txt_intParcelaFinal.Text = "") Then
        ExibeMensagem "O intervalo de parcelas deve ser definido "
        tab_3DPasta.Tab = 0
        If txt_intParcelaInicial.Text = "" Then
            txt_intParcelaInicial.SetFocus
        Else
            txt_intParcelaFinal.SetFocus
        End If
        Exit Function
    End If
    If Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal) Then
        ExibeMensagem "O número da parcela final deve ser maior que o número da parcela inicial"
        tab_3DPasta.Tab = 0
        txt_intParcelaFinal.SetFocus
        Exit Function
    End If
    
    If dbcintContribuinte.Text = "" Then
        ExibeMensagem "Selecione um Contribuinte "
        tab_3DPasta.Tab = 0
        dbcintContribuinte.SetFocus
        Exit Function
    End If
    If txt_intExercicio.Text = "" Then
        ExibeMensagem "O Exercício deve ser Digitado."
        tab_3DPasta.Tab = 0
        txt_intExercicio.SetFocus
        Exit Function
    End If
    
    If chk_EmBranco1.Value = 0 Then
        If txt_Mensagem1.Text = "" Then
            ExibeMensagem "A mensagem 1 tem que ser selecionada."
            Exit Function
        End If
    End If
    
    If chk_EmBranco2.Value = 0 Then
        If txt_Mensagem2.Text = "" Then
            ExibeMensagem "A mensagem 2 tem que ser selecionada."
            Exit Function
        End If
    End If
    
blnDadosGuiaOK = True
End Function

Private Sub LimpaObjetos()
    dbc_intMensagem1.BoundText = ""
    dbc_intMensagem2.BoundText = ""
    txt_Mensagem1 = ""
    txt_Mensagem2 = ""
End Sub

Private Sub dbc_intMensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem1
End Sub

Private Sub dbc_intMensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem2
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub
Private Sub txt_intParcelaInicial_GotFocus()
    tab_3DPasta.Tab = 0
    MarcaCampo txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaFinal_GotFocus()
    tab_3DPasta.Tab = 0
    MarcaCampo txt_intParcelaFinal
End Sub

Private Sub txt_intParcelaFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaFinal
End Sub
