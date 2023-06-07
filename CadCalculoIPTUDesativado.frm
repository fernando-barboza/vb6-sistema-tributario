VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadCalculoIPTUDesativado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo do IPTU"
   ClientHeight    =   6300
   ClientLeft      =   1920
   ClientTop       =   2445
   ClientWidth     =   8250
   Icon            =   "CadCalculoIPTUDesativado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8250
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6135
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   90
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "IPTU"
      TabPicture(0)   =   "CadCalculoIPTUDesativado.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Receitas a serem calculadas"
      TabPicture(1)   =   "CadCalculoIPTUDesativado.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdb_Atividades"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Emissão de Guias de Arrecadação"
      TabPicture(2)   =   "CadCalculoIPTUDesativado.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_EmissaoDeGuias"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5325
         Left            =   -74820
         TabIndex        =   21
         Top             =   600
         Width           =   7665
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1695
            Left            =   420
            TabIndex        =   27
            Top             =   675
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco1 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   29
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem1 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem1 
               Height          =   315
               Left            =   1080
               TabIndex        =   30
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbl_Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1695
            Left            =   420
            TabIndex        =   22
            Top             =   3015
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   24
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem2 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   25
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbl_Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   390
               Width           =   780
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados"
         Height          =   5265
         Left            =   240
         TabIndex        =   10
         Top             =   570
         Width           =   7605
         Begin VB.TextBox txt_intParcelaFinal 
            Height          =   285
            Left            =   3180
            MaxLength       =   15
            TabIndex        =   8
            Top             =   2355
            Width           =   480
         End
         Begin VB.TextBox txt_intParcelaInicial 
            Height          =   285
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   7
            Top             =   2355
            Width           =   480
         End
         Begin VB.CheckBox chk_SelecionarTodos 
            Caption         =   "Selecionar Todas as Inscrições"
            Height          =   225
            Left            =   2160
            TabIndex        =   2
            Top             =   1020
            Width           =   4245
         End
         Begin VB.TextBox txt_dblDesconto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3945
            MaxLength       =   2
            TabIndex        =   4
            Top             =   1290
            Width           =   855
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   3
            Top             =   1290
            Width           =   675
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoFinal 
            Height          =   315
            Left            =   2160
            TabIndex        =   1
            Tag             =   $"CadCalculoIPTUDesativado.frx":1096
            Top             =   660
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
            Height          =   315
            Left            =   2160
            TabIndex        =   5
            Top             =   1620
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoInicial 
            Height          =   315
            Left            =   2160
            TabIndex        =   0
            Tag             =   $"CadCalculoIPTUDesativado.frx":1198
            Top             =   300
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_VencimentoParcelas 
            Height          =   2355
            Left            =   180
            TabIndex        =   9
            Top             =   2715
            Width           =   7245
            _ExtentX        =   12779
            _ExtentY        =   4154
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   "PKID"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nº Parcela"
            Columns(1).DataField=   "intNumero"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Data Vencimento"
            Columns(2).DataField=   "dtmDataDaParcela"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=3043"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2963"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=3545"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3466"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
            Left            =   2160
            TabIndex        =   6
            Top             =   1980
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label lbl_ParcelaFinal 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   2775
            TabIndex        =   19
            Top             =   2400
            Width           =   225
         End
         Begin VB.Label lbl_ParcelaIncial 
            AutoSize        =   -1  'True
            Caption         =   "Parcela"
            Height          =   195
            Left            =   1500
            TabIndex        =   18
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label lbl_intOcorrencia 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrência"
            Height          =   195
            Left            =   1260
            TabIndex        =   17
            Top             =   2040
            Width           =   780
         End
         Begin VB.Label lbl_strComposicaoReceita 
            AutoSize        =   -1  'True
            Caption         =   "Composição da Receita"
            Height          =   195
            Left            =   345
            TabIndex        =   16
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lbl_dblDesconto 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
            Height          =   195
            Left            =   3150
            TabIndex        =   15
            Top             =   1335
            Width           =   690
         End
         Begin VB.Label lbl_p2 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   4875
            TabIndex        =   14
            Top             =   1335
            Width           =   120
         End
         Begin VB.Label lbl_Exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   1365
            TabIndex        =   13
            Top             =   1335
            Width           =   675
         End
         Begin VB.Label lbl_strInscricaoCadastralInicial 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1800
         End
         Begin VB.Label lbl_strInscricaoCadastralFinal 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição  Cadastral Final"
            Height          =   195
            Left            =   270
            TabIndex        =   11
            Top             =   720
            Width           =   1770
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
         Height          =   4935
         Left            =   -74700
         TabIndex        =   32
         Top             =   810
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   8705
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
         Splits(0)._ColumnProps(14)=   "Column(2).Width=12065"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=11986"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "frmCadCalculoIPTUDesativado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando             As Boolean
Dim mobjAux                   As Object
Dim mblnSelecionou            As Boolean
Dim xarReceita                As XArrayDB
Dim mblnPrimeiraVez           As Boolean

Private Sub chk_EmBranco1_GotFocus()
    tab_3dPasta.Tab = 2
End Sub

Private Sub chk_EmBranco1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_EmBranco1
End Sub

Private Sub chk_EmBranco2_GotFocus()
    tab_3dPasta.Tab = 2
End Sub

Private Sub chk_EmBranco2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_EmBranco2
End Sub

Private Sub chk_SelecionarTodos_Click()
    If chk_SelecionarTodos.Value = 1 Then
        dbc_strInscricaoInicial.BoundText = ""
        dbc_strInscricaoFinal.BoundText = ""
        dbc_strInscricaoInicial.Enabled = False
        TrocaCorObjeto dbc_strInscricaoInicial, True
        dbc_strInscricaoFinal.Enabled = False
        TrocaCorObjeto dbc_strInscricaoFinal, True
    Else
        dbc_strInscricaoInicial.Enabled = True
        TrocaCorObjeto dbc_strInscricaoInicial, False
        dbc_strInscricaoFinal.Enabled = True
        TrocaCorObjeto dbc_strInscricaoFinal, False
    End If
End Sub

Private Sub chk_SelecionarTodos_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub chk_SelecionarTodos_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chk_SelecionarTodos
End Sub

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
   DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
    If Area = 2 And dbc_intComposicaoDaReceita.MatchedWithList Then
        MontaAtividade
        If txt_intExercicio.Text <> "" Then
            LeDaTabelaParaObj gstrVencimentosDasParcelas, tdb_VencimentoParcelas, gQueryTDB_VencimentoParcelasReceita(dbc_intComposicaoDaReceita.BoundText, Val(txt_intExercicio.Text))
            tdb_VencimentoParcelas.MoveFirst
            txt_intParcelaInicial.Text = gstrENulo(tdb_VencimentoParcelas.Columns("intNumero").Value)
            tdb_VencimentoParcelas.MoveLast
            txt_intParcelaFinal.Text = gstrENulo(tdb_VencimentoParcelas.Columns("intNumero").Value)
            tdb_VencimentoParcelas.MoveFirst
        End If
    End If
End Sub

Private Sub dbc_intComposicaoDaReceita_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicaoDaReceita
End Sub

Private Sub dbc_intMensagem1_GotFocus()
    tab_3dPasta.Tab = 2
End Sub

Private Sub dbc_intMensagem1_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intMensagem1, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intMensagem2_GotFocus()
    tab_3dPasta.Tab = 2
End Sub

Private Sub dbc_intMensagem2_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intMensagem2, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intOcorrencia_Click(Area As Integer)
   DropDownDataCombo dbc_intOcorrencia, Me, Area
End Sub

Private Sub dbc_intOcorrencia_GotFocus()
   tab_3dPasta.Tab = 0
End Sub

Private Sub dbc_intOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intOcorrencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intOcorrencia
End Sub

Private Sub dbc_strInscricaoFinal_Click(Area As Integer)
   DropDownDataCombo dbc_strInscricaoFinal, Me, Area
End Sub

Private Sub dbc_strInscricaoFinal_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbc_strInscricaoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoInicial_Click(Area As Integer)
   DropDownDataCombo dbc_strInscricaoInicial, Me, Area
   If dbc_strInscricaoInicial.Text <> "" Then dbc_strInscricaoFinal.Text = dbc_strInscricaoInicial.Text
End Sub

Private Sub dbc_strInscricaoInicial_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbc_strInscricaoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoInicial, Me, , KeyCode, Shift
End Sub


Private Sub Form_Activate()
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
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
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub Form_Load()
   mblnAlterando = False
   dbc_strInscricaoInicial.Tag = strQuery & ";IM.strInscricaoAnterior"
   dbc_strInscricaoFinal.Tag = strQuery & ";IM.strInscricaoAnterior"
   dbc_intOcorrencia.Tag = strQueryOcorrencia & ";strDescricao"
   dbc_intComposicaoDaReceita.Tag = QueryComposicao & ";strDescricao"
   dbc_intMensagem1.Tag = strQueryMensagem & ";strDescricao"
   dbc_intMensagem2.Tag = strQueryMensagem & ";strDescricao"
   VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

'******************************************************************************************
' Data: 07/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql              As String
    Dim adoResultado        As ADODB.Recordset
    Dim blnExiteLancamento  As Boolean
    Dim strInscricaoInicial As String
    Dim strInscricaoFinal   As String

    Dim strSPParameters     As String

    On Error Resume Next

    If strModoOperacao = gstrPreencherLista Or strModoOperacao = gstrLocalizar Then
        ToolBarGeral strModoOperacao, gstrVencimentosDasParcelas, False, tdb_VencimentoParcelas, Me, mobjAux, gQueryTDB_VencimentoParcelasReceita(16, Val(txt_intExercicio.Text))
        Exit Sub
    End If
    
    If chk_SelecionarTodos.Value = 0 Then
        strInscricaoInicial = dbc_strInscricaoInicial.BoundText
        strInscricaoFinal = dbc_strInscricaoFinal.BoundText
    End If
    
    If strModoOperacao = gstrImprimir Then
        If blnDadosGuiaOK Then
            strSql = gstrQueryRelatorioGuiaDeArrecadacao(blnExiteLancamento, strInscricaoInicial, strInscricaoFinal, txt_intExercicio.Text, dbc_intComposicaoDaReceita.BoundText, IIf(chk_SelecionarTodos.Value = 1, True, False), , txt_intParcelaInicial.Text, txt_intParcelaFinal.Text)
            If blnExiteLancamento Then
                Set gfrmFormularioQueEstaImprimindoGuia = Me
                rptGuiaDeArrecadacaoMunicipal.strImposto = dbc_intComposicaoDaReceita.Text
                ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, strSql
            End If
        End If
    End If
    
    If strModoOperacao = gstrCalcularReajuste Then
        If blnValidaDados Then
            If MsgBox("Deseja Efetuar o Cálculo de " & dbc_intComposicaoDaReceita.Text & " ? ", vbYesNo, "Tributário") = vbNo Then
                GoTo Fim
            End If
            Set gobjBanco = New clsBanco
            Screen.MousePointer = vbHourglass
            gobjBanco.ExecutaBeginTrans
            If Not gBlnVerificaLancamentos(txt_intExercicio.Text, _
                                           dbc_intComposicaoDaReceita.BoundText, _
                                           dbc_intComposicaoDaReceita.Text, _
                                           (Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1), _
                                            "IPTU", _
                                           chk_SelecionarTodos.Value, _
                                           strInscricaoInicial, strInscricaoFinal) Then
                gobjBanco.ExecutaRollbackTrans
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
           'Concatena os Procedimento e seus Devidos Parâmetros
           
'            strSql = Chr(34) & strPKId & Chr(34) & "," & Chr(34) & strInscricaoInicial & Chr(34) & ","
            strSPParameters = Chr(34) & strPKId & Chr(34) & "," & Chr(34) & strInscricaoInicial & Chr(34) & "," & _
                      Chr(34) & strInscricaoFinal & Chr(34) & "," & dbc_intComposicaoDaReceita.BoundText & "," & _
                      Val(txt_intParcelaInicial.Text) & "," & Val(txt_intParcelaFinal.Text) & "," & _
                      Val(txt_intExercicio.Text) & "," & dbc_intOcorrencia.BoundText & "," & gstrConvVrParaSql(txt_dblDesconto) & "," & _
                      glngCodUsr

            If (bytDBType = EDatabases.Oracle) Then
                strSql = "DECLARE "
                strSql = strSql & "dblValor NUMBER := 0; "
                strSql = strSql & "BEGIN "
            
            End If

'            strSql = " sp_CalculoFormulaExecutada -24,NULL, '" & strSql & "' "
            strSql = strSql & "sp_CalculoFormulaExecutada " & IIf((bytDBType = EDatabases.Oracle), "(", "") & _
                "-24," & IIf((bytDBType = EDatabases.Oracle), "dblValor", "NULL") & _
                ", '" & strSPParameters & "'"
    
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "); ", "")
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
            Set gobjBanco = New clsBanco
'            If gobjBanco.CriaADO(strSql, 59, adoResultado) Then
            If gobjBanco.Execute(strSql) Then
                gobjBanco.ExecutaCommitTrans
                Screen.MousePointer = vbNormal
                ExibeMensagem "IPTU Gerado com Sucesso!"
            Else
                gobjBanco.ExecutaRollbackTrans
Fim:
                Screen.MousePointer = vbNormal
            End If
        End If
    End If
    
    If strModoOperacao = gstrNovo Then
        LimpaControlesDoFormulario
    End If
    
    If strModoOperacao = gstrFechar Then
        Unload Me
    End If
    If strModoOperacao = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If
End Sub

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

Private Sub LimpaControlesDoFormulario()
    dbc_strInscricaoInicial.BoundText = ""
    dbc_strInscricaoFinal.BoundText = ""
    chk_SelecionarTodos.Value = 0
    txt_intExercicio.Text = ""
    txt_dblDesconto.Text = ""
    dbc_intComposicaoDaReceita.BoundText = ""
    Set tdb_VencimentoParcelas.DataSource = Nothing
    dbc_intOcorrencia.BoundText = ""
    txt_intParcelaInicial.Text = ""
    txt_intParcelaFinal.Text = ""
    
    chk_EmBranco1.Value = 0
    chk_EmBranco2.Value = 0
    dbc_intMensagem1.BoundText = ""
    dbc_intMensagem2.BoundText = ""
    txt_Mensagem1.Text = ""
    txt_Mensagem2.Text = ""
    tab_3dPasta.Tab = 0
    dbc_strInscricaoInicial.SetFocus
End Sub

Private Function blnValidaDados() As Boolean
    Dim i As Integer
    If chk_SelecionarTodos.Value = 0 Then
        If Trim(dbc_strInscricaoInicial.Text) = "" Then
            ExibeMensagem "A Inscrição Inicial deve ser selecionada."
            tab_3dPasta.Tab = 0
            dbc_strInscricaoInicial.SetFocus
            Exit Function
        End If
        If Trim(dbc_strInscricaoFinal.Text) = "" Then
            ExibeMensagem "A Inscrição Final deve ser selecionada."
            tab_3dPasta.Tab = 0
            dbc_strInscricaoFinal.SetFocus
            Exit Function
        End If
    End If
    If Trim(dbc_intComposicaoDaReceita.Text) = "" Then
        ExibeMensagem "A Composição da Receita deve ser selecionada."
        tab_3dPasta.Tab = 0
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If
    If Trim(txt_intExercicio.Text) = "" Then
        ExibeMensagem "O exercício deve ser digitado."
        tab_3dPasta.Tab = 0
        txt_intExercicio.SetFocus
        Exit Function
    End If
    If txt_dblDesconto.Text = "" Then
        txt_dblDesconto.Text = "0"
    End If
    If Trim(dbc_intOcorrencia.Text) = "" Then
        ExibeMensagem "A Ocorrencia deve ser Selecionada."
        tab_3dPasta.Tab = 0
        txt_intExercicio.SetFocus
        Exit Function
    End If
    If (txt_intParcelaInicial.Text = "") Or (txt_intParcelaFinal.Text = "") Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "O intervalo de parcelas deve ser definido "
        If txt_intParcelaInicial.Text = "" Then
            txt_intParcelaInicial.SetFocus
        Else
            txt_intParcelaFinal.SetFocus
        End If
        Exit Function
    End If
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 Then
            blnValidaDados = True
            Exit Function
        End If
    Next
    ExibeMensagem "Selecione uma receita para efetuar o cálculo!"
    tab_3dPasta.Tab = 1
End Function

Private Function strQueryOcorrencia() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrOcorrencia
    strSql = strSql & " WHERE "
    strSql = strSql & " intUtilizacaoDaOcorrencia = 1 "
    strSql = strSql & " ORDER BY strDescricao "
    strQueryOcorrencia = strSql
End Function

Private Function strQuery() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    strSql = " SELECT IM.strInscricaoAnterior, "
'             " LTRIM(RTRIM(IM.strInscricaoAnterior)) + ' - ' +  "
    strSql = strSql & " LTRIM(RTRIM(IM.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & _
             " LTRIM(RTRIM(CC.strNome)) AS Descricao FROM " & gstrImobiliario & _
             " IM," & gstrContribuinte & " CC " & _
             " WHERE CC.PKId = IM.intContribuinte " & _
             " ORDER BY strInscricaoAnterior"
    strQuery = strSql
End Function

Private Function QueryComposicao() As String
    Dim strSql As String
    strSql = " SELECT PKId, strDescricao FROM " & _
              gstrComposicaoDaReceita & _
             " WHERE intUtilizacao = 1 "
    QueryComposicao = strSql
End Function

Private Sub MontaAtividade()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
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
    strSql = strSql & " AND B.intComposicaoDaReceita = " & dbc_intComposicaoDaReceita.BoundText
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                xarReceita.ReDim 0, .RecordCount - 1, 0, 2
                Do While Not .EOF
                    varAux = !Pkid
                    xarReceita(.AbsolutePosition - 1, 0) = varAux
                    
                    varAux = False
                    xarReceita(.AbsolutePosition - 1, 2) = varAux
                
                    varAux = !strDescricao
                    xarReceita(.AbsolutePosition - 1, 3) = varAux
                    
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

Private Sub tdb_VencimentoParcelas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_dblDesconto
End Sub

Private Sub txt_dblDesconto_GotFocus()
    MarcaCampo txt_dblDesconto
    tab_3dPasta.Tab = 0
End Sub

Private Sub txt_dblDesconto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_dblDesconto
End Sub


Private Sub txt_intExercicio_LostFocus()
    If txt_intExercicio <> "" And dbc_intComposicaoDaReceita.BoundText <> "" Then
        LeDaTabelaParaObj gstrVencimentosDasParcelas, tdb_VencimentoParcelas, gQueryTDB_VencimentoParcelasReceita(16, Val(txt_intExercicio.Text))
    End If
End Sub




'''>>>>>>>>>>>>>>>>>> DE ARRECADAÇÃO



Private Function strQueryInscricao() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT EC.PKId, EC.strInscricaoCadastral "
    strSql = strSql & " FROM "
    strSql = strSql & gstrEconomico & " EC,"
    strSql = strSql & gstrTributoEmpresa & " EM"
    strSql = strSql & " WHERE EC.PKId = EM.intEconomico AND EC.dtmDataBaixa IS NULL "
'    strSql = strSql & " ORDER BY CONVERT(NUMERIC, EC.strInscricaoCadastral) "
    strSql = strSql & " ORDER BY " & gstrCONVERT(cdt_numeric, "EC.strInscricaoCadastral")
strQueryInscricao = strSql
End Function

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
   If Area = 2 And dbc_intMensagem1.MatchedWithList Then
       LeDoComboParaTXT1
   End If
End Sub

Private Sub dbc_intMensagem2_Click(Area As Integer)
   DropDownDataCombo dbc_intMensagem2, Me, Area
   If Area = 2 And dbc_intMensagem2.MatchedWithList Then
       LeDoComboParaTXT2
   End If
End Sub

Private Function LeDoComboParaTXT1()
Dim strSql       As String
Dim adoResultado As ADODB.Recordset
   strSql = ""
   strSql = strSql & " SELECT strMensagem "
   strSql = strSql & " FROM " & gstrMensagem
   strSql = strSql & " WHERE PKId = " & Val(dbc_intMensagem1.BoundText)
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
Dim strSql       As String
Dim adoResultado As ADODB.Recordset
    strSql = ""
    strSql = strSql & " SELECT strMensagem "
    strSql = strSql & " FROM " & gstrMensagem
    strSql = strSql & " WHERE PKId = " & Val(dbc_intMensagem2.BoundText)
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
    
    If chk_SelecionarTodos.Value = 0 Then
        If dbc_strInscricaoInicial.BoundText = "" Then
            tab_3dPasta.Tab = 0
            ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para gerar a Guia de Arrecadação."
            dbc_strInscricaoInicial.SetFocus
            Exit Function
        End If
        
        If dbc_strInscricaoFinal.BoundText = "" Then
            tab_3dPasta.Tab = 0
            ExibeMensagem "Selecione uma Inscrição Cadastral Final para gerar a  Guia de Arrecadação."
            dbc_strInscricaoFinal.SetFocus
            Exit Function
        End If
    End If
    If Trim(txt_intExercicio.Text) = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "O Exercício deve ser Digitado."
        txt_intExercicio.SetFocus
        Exit Function
    End If
    If dbc_intComposicaoDaReceita.BoundText = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Composição da Receita deve ser selecionada."
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If
    If Trim(txt_intParcelaInicial.Text) = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Parcela Inicial deve ser digitada."
        txt_intParcelaInicial.SetFocus
        Exit Function
    ElseIf Trim(txt_intParcelaFinal.Text) = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Parcela Final deve ser digitada."
        txt_intParcelaFinal.SetFocus
        Exit Function
    ElseIf Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal.Text) Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Parcela Inicial deve ser anterior a Parcela Final."
        txt_intParcelaInicial.SetFocus
        Exit Function
    Else
        tdb_VencimentoParcelas.MoveFirst
        If Val(txt_intParcelaInicial.Text) < Val(tdb_VencimentoParcelas.Columns("intNumero").Value) Then
            tab_3dPasta.Tab = 0
            ExibeMensagem "A Parcela Incial não ser pode menor que o número da primeira parcela."
            txt_intParcelaInicial.SetFocus
            Exit Function
        Else
            tdb_VencimentoParcelas.MoveLast
            If Val(txt_intParcelaFinal.Text) > Val(tdb_VencimentoParcelas.Columns("intNumero").Value) Then
                tab_3dPasta.Tab = 0
                ExibeMensagem "A Parcela Final não ser pode maior que o número de última parcela."
                tdb_VencimentoParcelas.MoveFirst
                txt_intParcelaFinal.SetFocus
                Exit Function
            End If
        End If
    End If
        
    If chk_EmBranco1.Value = 0 Then
        If Trim(txt_Mensagem1.Text) = "" Then
            tab_3dPasta.Tab = 2
            ExibeMensagem "A mensagem 1 tem que ser selecionada."
            dbc_intMensagem1.SetFocus
            Exit Function
        End If
    End If
    If chk_EmBranco2.Value = 0 Then
        If Trim(txt_Mensagem2.Text) = "" Then
            tab_3dPasta.Tab = 2
            ExibeMensagem "A mensagem 2 tem que ser selecionada."
            dbc_intMensagem2.SetFocus
            Exit Function
        End If
    End If
    blnDadosGuiaOK = True
End Function

Private Sub dbc_intMensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem1
End Sub

Private Sub dbc_intMensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem2
End Sub

Private Sub txt_intExercicio_GotFocus()
    tab_3dPasta.Tab = 0
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub dbc_strInscricaoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoInicial
End Sub

Private Sub dbc_strInscricaoFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoFinal
End Sub

Private Sub txt_intParcelaInicial_GotFocus()
    tab_3dPasta.Tab = 0
    MarcaCampo txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaFinal_GotFocus()
    tab_3dPasta.Tab = 0
    MarcaCampo txt_intParcelaFinal
End Sub

Private Sub txt_intParcelaFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaFinal
End Sub

Private Sub txt_Mensagem1_GotFocus()
    tab_3dPasta.Tab = 2
    MarcaCampo txt_Mensagem1
End Sub

Private Sub txt_Mensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Mensagem1
End Sub

Private Sub txt_Mensagem2_GotFocus()
    tab_3dPasta.Tab = 2
    MarcaCampo txt_Mensagem2
End Sub

Private Sub txt_Mensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Mensagem2
End Sub
Private Sub tdb_Atividades_AfterColUpdate(ByVal ColIndex As Integer)
    tdb_Atividades.Update
End Sub

Private Sub tdb_Atividades_AfterUpdate()
    tdb_Atividades.Update
End Sub
