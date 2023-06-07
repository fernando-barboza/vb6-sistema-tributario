VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCalculoContribuicaoMelhoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo da Contribuição de Melhoria"
   ClientHeight    =   6060
   ClientLeft      =   2250
   ClientTop       =   2400
   ClientWidth     =   8100
   Icon            =   "CalculoContribuicaoMelhoria.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5865
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   10345
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Contribuição de Melhoria"
      TabPicture(0)   =   "CalculoContribuicaoMelhoria.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_intOcorrencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_intComposicaoDaReceita"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintTabelaDeEdital"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintSecaoDeLogradouro"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintIntervalo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbldtmLancamento"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbldtmVencimento"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_strInscricaoCadastral"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_ParcelaFinal"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_ParcelaIncial"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl_intExercicio"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dbc_strInscricaoCadastral"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "tdb_Atividades"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dbcintSecaoDeLogradouro"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dbcintTabelaDeEdital"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dbc_intOcorrencia"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dbc_intComposicaoDaReceita"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtintIntervalo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtdtmVencimento"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtdtmLancamento"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_intParcelaFinal"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_intParcelaInicial"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txt_intExercicio"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Emissão de Guias de Arrecadação"
      TabPicture(1)   =   "CalculoContribuicaoMelhoria.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_EmissaoDeGuias"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox txt_intExercicio 
         Height          =   285
         Left            =   4155
         MaxLength       =   4
         TabIndex        =   33
         Top             =   2400
         Width           =   525
      End
      Begin VB.TextBox txt_intParcelaInicial 
         Height          =   285
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   30
         Top             =   2760
         Width           =   480
      End
      Begin VB.TextBox txt_intParcelaFinal 
         Height          =   285
         Left            =   2970
         MaxLength       =   15
         TabIndex        =   29
         Top             =   2760
         Width           =   480
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
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2400
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
         Left            =   6570
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2400
         Width           =   975
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
         Left            =   6570
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2760
         Width           =   975
      End
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5145
         Left            =   -74850
         TabIndex        =   17
         Top             =   450
         Width           =   7575
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1695
            Left            =   390
            TabIndex        =   21
            Top             =   780
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco1 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   8
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem1 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem1 
               Height          =   315
               Left            =   1080
               TabIndex        =   9
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
               TabIndex        =   23
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1695
            Left            =   390
            TabIndex        =   18
            Top             =   2790
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   10
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem2 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   11
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
               TabIndex        =   20
               Top             =   390
               Width           =   780
            End
         End
      End
      Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
         Height          =   315
         Left            =   2010
         TabIndex        =   1
         Top             =   960
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intOcorrencia 
         Height          =   315
         Left            =   2010
         TabIndex        =   0
         Top             =   600
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintTabelaDeEdital 
         Height          =   315
         Left            =   2010
         TabIndex        =   2
         Top             =   1320
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintSecaoDeLogradouro 
         Height          =   315
         Left            =   2010
         TabIndex        =   3
         Top             =   1680
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
         Height          =   2505
         Left            =   300
         TabIndex        =   16
         Top             =   3180
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   4419
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
         Splits(0)._ColumnProps(14)=   "Column(2).Width=11642"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=11562"
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
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastral 
         Height          =   315
         Left            =   2010
         TabIndex        =   27
         Top             =   2040
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin VB.Label lbl_intExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   3360
         TabIndex        =   34
         Top             =   2475
         Width           =   675
      End
      Begin VB.Label lbl_ParcelaIncial 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         Height          =   195
         Left            =   1350
         TabIndex        =   32
         Top             =   2790
         Width           =   540
      End
      Begin VB.Label lbl_ParcelaFinal 
         AutoSize        =   -1  'True
         Caption         =   "até"
         Height          =   195
         Left            =   2580
         TabIndex        =   31
         Top             =   2790
         Width           =   225
      End
      Begin VB.Label lbl_strInscricaoCadastral 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   480
         TabIndex        =   28
         Top             =   2130
         Width           =   1350
      End
      Begin VB.Label lbldtmVencimento 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data de vencimento"
         Height          =   195
         Left            =   5010
         TabIndex        =   26
         Top             =   2490
         Width           =   1440
      End
      Begin VB.Label lbldtmLancamento 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data de Lançamento"
         Height          =   195
         Left            =   375
         TabIndex        =   25
         Top             =   2490
         Width           =   1500
      End
      Begin VB.Label lblintIntervalo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Intervalo de dias entre as Parcelas"
         Height          =   195
         Left            =   4005
         TabIndex        =   24
         Top             =   2820
         Width           =   2445
      End
      Begin VB.Label lblintSecaoDeLogradouro 
         AutoSize        =   -1  'True
         Caption         =   "Seção de Logradouro"
         Height          =   195
         Left            =   330
         TabIndex        =   15
         Top             =   1770
         Width           =   1545
      End
      Begin VB.Label lblintTabelaDeEdital 
         AutoSize        =   -1  'True
         Caption         =   "Edital"
         Height          =   195
         Left            =   1485
         TabIndex        =   14
         Top             =   1410
         Width           =   390
      End
      Begin VB.Label lbl_intComposicaoDaReceita 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1050
         Width           =   1695
      End
      Begin VB.Label lbl_intOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmCalculoContribuicaoMelhoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xarReceita                  As XArrayDB
Dim adoResultado                As ADODB.Recordset

Private Sub MontaAtividade(intComposicaoReceita As Integer)
Dim strSQL As String
Dim adoRec As ADODB.Recordset
Dim varAux As String

On Error GoTo Err_Handle

Set xarReceita = New XArrayDB
xarReceita.Clear

xarReceita.ReDim 0, 0, 0, 2

strSQL = ""
strSQL = strSQL & " SELECT A.PKId, A.strDescricao FROM "
strSQL = strSQL & gstrReceita & " A,"
strSQL = strSQL & gstrValorCompoRec & " B"
strSQL = strSQL & " WHERE A.PKId = B.intReceita "
strSQL = strSQL & " AND B.intComposicaoDaReceita = " & intComposicaoReceita

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
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

Private Function strQueryContribuicaoMelhoria() As String
Dim strSQL As String

strSQL = ""
strSQL = strSQL & "SELECT CO.PKId, CO.strNome, IM.strInscricaoAnterior, TE.dblCustoDaParcela FROM "
strSQL = strSQL & gstrContribuinte & " CO, "
strSQL = strSQL & gstrImobiliario & " IM, "
strSQL = strSQL & gstrSecaoLogradouro & " SL, "
strSQL = strSQL & gstrTabelaDeEdital & " TE"
strSQL = strSQL & " WHERE "
strSQL = strSQL & " CO.PKId = IM.intContribuinte"
strSQL = strSQL & " AND SL.PKId = IM.intSecoes"
strSQL = strSQL & " AND SL.PKId = TE.intSecaoDeLogradouro"
strSQL = strSQL & " AND TE.PKId = " & dbcintTabelaDeEdital.BoundText
strSQL = strSQL & " ORDER BY IM.strInscricaoAnterior "
strQueryContribuicaoMelhoria = strSQL
End Function

Private Function strQuerySecao() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
'            Foi mantida a forma antiga para o SQL Server pois não era possível o
'            deslocamento completo devido à incompatibilidade entres os bancos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "SELECT SL.PKId , "
'    strSql = strSql & "ISNULL(SL.strInscricaoCadastral, '') + ' - ' + RTRIM(LTRIM(ISNULL(TL.strSigla, '') + ' ' + ISNULL(U.strDescricao,'') + ' ' + L.strDescricao)) AS Logradouro "
    strSQL = strSQL & gstrISNULL("SL.strInscricaoCadastral", "''") & strCONCAT & " ' - ' " & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & strCONCAT & " ' ' " & strCONCAT & gstrISNULL("U.strDescricao", "''") & strCONCAT & " ' ' " & strCONCAT & " L.strDescricao)) AS Logradouro "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrTabelaDeEdital & " TE, "
    strSQL = strSQL & gstrEditalSecaoLogradouro & " ESL, "
        
    If (bytDBType = EDatabases.SQLServer) Then
        strSQL = strSQL & " ((" & gstrSecaoLogradouro & " SL "
        strSQL = strSQL & "LEFT JOIN " & gstrLogradouro & " L ON SL.intLogradouro = L.PKId) "
        strSQL = strSQL & "LEFT JOIN " & gstrTituloLogradouro & " U ON L.intTituloLogradouro = U.PKId) "
        strSQL = strSQL & "LEFT JOIN " & gstrTipoLogradouro & " TL ON L.intTipoLogradouro = TL.PKId "
    
    ElseIf (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & gstrSecaoLogradouro & " SL, "
        strSQL = strSQL & gstrLogradouro & " L, "
        strSQL = strSQL & gstrTituloLogradouro & " U, "
        strSQL = strSQL & gstrTipoLogradouro & " TL "
    
    End If
    
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & " SL.PKID = ESL.intSecaoDeLogradouro "
    strSQL = strSQL & " AND TE.PKID = ESL.intTabelaDeEdital "
    
    If (bytDBType = EDatabases.Oracle) Then
        strSQL = strSQL & " AND SL.intLogradouro " & strOUTJOracle & "= L.PKId "
        strSQL = strSQL & " AND L.intTituloLogradouro = U.PKId " & strOUTJOracle
        strSQL = strSQL & " AND L.intTipoLogradouro = TL.PKId " & strOUTJOracle
    
    End If
    
    strSQL = strSQL & " AND TE.PKId = " & dbcintTabelaDeEdital.BoundText
    strSQL = strSQL & " ORDER BY Logradouro"
strQuerySecao = strSQL
End Function

Private Function strQuerryOcorrencia() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrOcorrencia
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " intUtilizacaoDaOcorrencia = 1 "
    strSQL = strSQL & " ORDER BY strDescricao "
strQuerryOcorrencia = strSQL
End Function

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
    If Area = 2 And dbc_intComposicaoDaReceita.MatchedWithList Then
        MontaAtividade dbc_intComposicaoDaReceita.BoundText
    End If
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, , KeyCode, Shift
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

Private Sub dbc_strInscricaoCadastral_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, Area
End Sub

Private Sub dbc_strInscricaoCadastral_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, , KeyCode, Shift
End Sub

Private Sub dbcintSecaoDeLogradouro_Click(Area As Integer)
    DropDownDataCombo dbcintSecaoDeLogradouro, Me, Area
    If Area = 2 And dbcintSecaoDeLogradouro.MatchedWithList Then
        LeDaTabelaParaObj gstrImobiliario, dbc_strInscricaoCadastral, strQueryInscricao
    End If
End Sub

Private Sub dbcintSecaoDeLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintSecaoDeLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTabelaDeEdital_Click(Area As Integer)
    DropDownDataCombo dbcintTabelaDeEdital, Me, Area
    If Area = 2 Then
        If dbcintTabelaDeEdital.MatchedWithList Then
            VerificaListaAutomatica gstrSecaoLogradouro, dbcintSecaoDeLogradouro, strQuerySecao
        End If
    End If
End Sub

Private Sub dbcintTabelaDeEdital_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTabelaDeEdital, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 670
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub Form_Load()
LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoDaReceita, " SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita & " WHERE intUtilizacao = 1 ORDER BY strDescricao "
LeDaTabelaParaObj gstrOcorrencia, dbc_intOcorrencia, strQuerryOcorrencia
LeDaTabelaParaObj "", dbcintTabelaDeEdital, "SELECT PKId, strNomeDoEdital FROM " & gstrTabelaDeEdital & " ORDER BY strNomeDoEdital"
End Sub

Public Sub MantemForm(strModoOperacao As String)
    Dim strSQL As String
    Dim blnExiteLancamento As Boolean
    Dim strInscricao       As String
    
    If tab_3dPasta.Tab = 1 Then
        If UCase(strModoOperacao) = UCase(gstrImprimir) Then
            If blnDadosGuiaOK = True Then
                strInscricao = Trim(Left(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, _
                                "-", vbTextCompare) - 1))
                strSQL = gstrQueryRelatorioGuiaDeArrecadacao(blnExiteLancamento, strInscricao, strInscricao, _
                                    txt_intExercicio.Text, dbc_intComposicaoDaReceita.BoundText, , , _
                                    Val(txt_intParcelaInicial.Text), Val(txt_intParcelaFinal.Text))
                If blnExiteLancamento Then
                    Set gfrmFormularioQueEstaImprimindoGuia = Me
                    rptGuiaDeArrecadacaoMunicipal.strImposto = dbc_intComposicaoDaReceita.Text
                    ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, strSQL
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
        Select Case strModoOperacao
            Case gstrNovo
                LimpaCampos
            Case gstrCalcularReajuste
                CalculoLancamento
            Case gstrPreencherLista
                If dbcintSecaoDeLogradouro.MatchedWithList = False Then Exit Sub
                dbc_strInscricaoCadastral.Tag = strQueryInscricao & ";I.strInscricaoAnterior"
                PreencherListaDeOpcoes dbc_strInscricaoCadastral
                dbc_strInscricaoCadastral.Tag = ""
                Exit Sub
        End Select
    End If

End Sub

Private Function blnDadosOk() As Boolean
    Dim i As Integer
    If dbc_intOcorrencia.Text = "" Then
        ExibeMensagem "Selecione uma " & lbl_intOcorrencia.Caption & "."
        blnDadosOk = False
        dbc_intOcorrencia.SetFocus
        Exit Function
    End If
    If dbc_intComposicaoDaReceita.Text = "" Then
        ExibeMensagem "Selecione uma " & lbl_intComposicaoDaReceita.Caption & "."
        blnDadosOk = False
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If
    If dbcintTabelaDeEdital.Text = "" Then
        ExibeMensagem "Selecione um " & lblintTabelaDeEdital.Caption & "."
        blnDadosOk = False
        dbcintTabelaDeEdital.SetFocus
        Exit Function
    End If
    If dbcintSecaoDeLogradouro.Text = "" Then
        ExibeMensagem "Selecione uma " & lblintSecaoDeLogradouro.Caption & "."
        blnDadosOk = False
        dbcintSecaoDeLogradouro.SetFocus
        Exit Function
    End If
    If txtdtmLancamento.Text = "" Then
        ExibeMensagem "O campo " & lbldtmLancamento.Caption & ", deve ser informado ."
        blnDadosOk = False
        txtdtmLancamento.SetFocus
        Exit Function
    End If
    If txt_intExercicio.Text = "" Then
        ExibeMensagem "O campo " & lbl_intExercicio.Caption & ", deve ser informado ."
        blnDadosOk = False
        txt_intExercicio.SetFocus
        Exit Function
    End If
    If txtdtmVencimento.Text = "" Then
        ExibeMensagem "O campo " & lbldtmVencimento.Caption & ", deve ser informado ."
        blnDadosOk = False
        txtdtmVencimento.SetFocus
        Exit Function
    End If
    If txtintIntervalo.Text = "" Then
        ExibeMensagem "O campo " & lblintIntervalo.Caption & ", deve ser informado ."
        blnDadosOk = False
        txtintIntervalo.SetFocus
        Exit Function
    End If
    If (txt_intParcelaInicial.Text = "") Or (txt_intParcelaFinal.Text = "") Then
        ExibeMensagem "O intervalo de parcelas deve ser definido "
        If txt_intParcelaInicial.Text = "" Then
            txt_intParcelaInicial.SetFocus
        Else
            txt_intParcelaFinal.SetFocus
        End If
        Exit Function
    End If
    If Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal) Then
        ExibeMensagem "O número da parcela final deve ser maior que o número da parcela inicial"
        txt_intParcelaFinal.SetFocus
        Exit Function
    End If
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 Then
            blnDadosOk = True
            Exit Function
        End If
    Next
    ExibeMensagem "Selecione uma receita para efetuar o cálculo!"
    
End Function


Private Sub CalculoLancamento()

'******************************************************************************************
' Data: 09/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/05/2003
' Alteração: - Substituição da chamada à função CriaADO por uma chamada à função
'            ExecuteStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL               As String
'    Dim adoResultado         As ADODB.Recordset
    Dim adoParameters        As ADODB.Parameters
    Dim strInscricao         As String
    Dim dblValorAparcelar    As String
    Dim dblValorNaoParcelado As String
    
    Screen.MousePointer = vbHourglass
    If blnDadosOk Then
        If dbc_strInscricaoCadastral.Text <> "" Then
            strInscricao = Trim(Left(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, _
                                    "-", vbTextCompare) - 1))
'            strSQL = "sp_CalculoParaUsuario '" & strInscricao & "','" & strPKId & "',1," & _
'                 dbc_intComposicaoDaReceita.BoundText & ",0,0,0,'" & dbcintTabelaDeEdital.BoundText & _
'                "," & dbcintSecaoDeLogradouro.BoundText & "," & NumeroDeReceitas & "," & IIf(strInscricao = "", 0, 1) & ", @dblValor OUTPUT'"
            strSQL = gstrStoredProcedure("sp_CalculoParaUsuario", "'" & strInscricao & "','" & strPKId & "',1," & _
                 dbc_intComposicaoDaReceita.BoundText & ",0,0,0,'" & dbcintTabelaDeEdital.BoundText & _
                "," & dbcintSecaoDeLogradouro.BoundText & "," & NumeroDeReceitas & "," & IIf(strInscricao = "", 0, 1) & _
                ", " & IIf((bytDBType = EDatabases.Oracle), ":dblValor", "@dblValor OUTPUT") & "'")
            Set gobjBanco = New clsBanco
'            If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
            If gobjBanco.ExecuteStoredProcedure(strSQL, 10, , adoParameters) Then
'                If Not (adoResultado.BOF And adoResultado.EOF) Then
                If Not (adoParameters Is Nothing) Then
                    'Mostra Valores para Usuário
'                    strSQL = "Confirma o cálculo de " & gstrConvVrDoSql(adoResultado!dblValorAparcelar) & _
'                            Chr(10) & " + " & gstrConvVrDoSql(adoResultado!dblValorNaoParcelado) & " em " & _
'                            (Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1) & " parcela(s) ?"
                    strSQL = "Confirma o cálculo de " & gstrConvVrDoSql(adoParameters("dblValorAparcelar")) & _
                            Chr(10) & " + " & gstrConvVrDoSql(adoParameters("dblValorNaoParcelado")) & " em " & _
                            (Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1) & " parcela(s) ?"
                    Screen.MousePointer = vbNormal
                    If MsgBox(strSQL, vbYesNo, "Tributário") = vbNo Then
                        Exit Sub
                    End If
'                    dblValorAparcelar = CStr(adoResultado!dblValorAparcelar)
                    dblValorAparcelar = CStr(adoParameters("dblValorAparcelar"))
'                    dblValorNaoParcelado = CStr(adoResultado!dblValorNaoParcelado)
                    dblValorNaoParcelado = CStr(adoParameters("dblValorNaoParcelado"))
                End If
            End If
        Else
            strInscricao = ""
            Screen.MousePointer = vbNormal
            If MsgBox("Deseja Efetuar o Cálculo da Contribuição de Melhoria ?", vbYesNo, "Tributário") = vbNo Then
                Exit Sub
            End If
        End If
        Screen.MousePointer = vbHourglass
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        If Not gBlnVerificaLancamentos(txt_intExercicio, dbc_intComposicaoDaReceita.BoundText, _
                                        dbc_intComposicaoDaReceita.Text, (Val(txt_intParcelaFinal) - _
                                        Val(txt_intParcelaInicial) + 1), gstrConvDtParaSql(txtdtmLancamento), _
                                        IIf(strInscricao = "", 1, 0), strInscricao) Then
            gobjBanco.ExecutaRollbackTrans
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        If strInscricao = "" Then
            dblValorAparcelar = "-1"
            dblValorNaoParcelado = "-1"
        End If
        strSQL = strLancamentos(dblValorAparcelar, dblValorNaoParcelado, strInscricao)
        If gobjBanco.Execute(strSQL, False) Then
            gobjBanco.ExecutaCommitTrans
            Screen.MousePointer = vbNormal
            ExibeMensagem "Cálculo efetuado com sucesso!"
        Else
            gobjBanco.ExecutaRollbackTrans
        End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Function strLancamentos(dblValorAparcelar As String, dblValorNaoParcelado As String, strInscricao As String) As String

'******************************************************************************************
' Data: 09/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL As String
'    strSQL = " sp_CalculoLancamentoReceitas 1, " & gstrConvVrParaSql(dblValorAparcelar) & " , " & gstrConvVrParaSql(dblValorNaoParcelado) & ",'" & _
'            strPKId & "','" & strContribuintes(strInscricao) & "'," & Val(txt_intExercicio.Text) & _
'            "," & dbc_intComposicaoDaReceita.BoundText & ", NULL, " & _
'            gstrConvDtParaSql(txtdtmLancamento.Text) & _
'            "," & gstrConvDtParaSql(txtdtmVencimento) & "," & Val(txt_intParcelaInicial.Text) & "," & Val(txt_intParcelaFinal.Text) & "," & _
'            Val(txtintIntervalo.Text) & ",3," & Val(dbc_intOcorrencia.BoundText) & ",1," & glngCodUsr & ",0,0,-1,'" & dbcintTabelaDeEdital.BoundText & _
'            "," & dbcintSecaoDeLogradouro.BoundText & "," & NumeroDeReceitas & "," & IIf(strInscricao = "", 0, 1) & ", @dblValor OUTPUT'"
    strSQL = gstrStoredProcedure("sp_CalculoLancamentoReceitas", "1, " & gstrConvVrParaSql(dblValorAparcelar) & " , " & gstrConvVrParaSql(dblValorNaoParcelado) & ",'" & _
            strPKId & "','" & strContribuintes(strInscricao) & "'," & Val(txt_intExercicio.Text) & _
            "," & dbc_intComposicaoDaReceita.BoundText & ", NULL, " & _
            gstrConvDtParaSql(txtdtmLancamento.Text) & _
            "," & gstrConvDtParaSql(txtdtmVencimento) & "," & Val(txt_intParcelaInicial.Text) & "," & Val(txt_intParcelaFinal.Text) & "," & _
            Val(txtintIntervalo.Text) & ",3," & Val(dbc_intOcorrencia.BoundText) & ",1," & glngCodUsr & ",0,0,-1,'" & dbcintTabelaDeEdital.BoundText & _
            "," & dbcintSecaoDeLogradouro.BoundText & "," & NumeroDeReceitas & "," & IIf(strInscricao = "", 0, 1) & _
            ", " & IIf((bytDBType = EDatabases.Oracle), ":dblValor", "@dblValor OUTPUT") & "'")
    
    strLancamentos = strSQL
End Function

Private Function strContribuintes(strInscricao As String) As String
    Dim strSQL As String
    
    strSQL = "SELECT intContribuinte , strInscricaoAnterior " & _
            " FROM " & gstrImobiliario & _
            " WHERE intSecoes = " & dbcintSecaoDeLogradouro.BoundText
    If strInscricao <> "" Then
        strSQL = strSQL & " AND strInscricaoAnterior = """ & strInscricao & """"
    End If
    strContribuintes = strSQL
End Function

Private Function NumeroDeReceitas() As Integer
    Dim iCont As Integer
    Dim i As Integer
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 And (xarReceita(i, 0) = 25 Or xarReceita(i, 0) = 95) Then
            iCont = iCont + 1
        End If
    Next
   
    NumeroDeReceitas = iCont
End Function

'Private Function RetornaCustoEdital(dblCustoDaParcela As Double) As Double
'Dim StrSql As String
'Dim adoRec As ADODB.Recordset
'
'StrSql = ""
'StrSql = StrSql & " SELECT ISNULL(SUM(CONVERT(NUMERIC(12,8),strMedidaDaTestada)),0) AS dblValor "
'StrSql = StrSql & " FROM " & gstrTestadaImobiliario & " A, "
'StrSql = StrSql & gstrImobiliario & " B, "
'StrSql = StrSql & gstrContribuinte & " C  "
'StrSql = StrSql & " WHERE "
'StrSql = StrSql & " B.PKId = A.intImobiliario"
'StrSql = StrSql & " AND C.PKId = B.intContribuinte "
'StrSql = StrSql & " AND intSecoes = " & dbcintSecaoDeLogradouro.BoundText
'
'Set gobjBanco = New clsBanco
'
'If gobjBanco.CriaADO(StrSql, 10, adoRec) Then
'    If adoRec!dblValor > 0 Then
'        RetornaCustoEdital = dblCustoDaParcela / adoRec!dblValor
'    Else
'        RetornaCustoEdital = 0
'    End If
'Else
'    RetornaCustoEdital = 0
'End If
'End Function

'Private Function RetornaValorTestada(lngPKId As Long) As Double
'Dim StrSql As String
'Dim adoRec As ADODB.Recordset
'
'StrSql = ""
'StrSql = StrSql & " SELECT ISNULL(SUM(CONVERT(NUMERIC(12,8),strMedidaDaTestada)),0) AS dblValor "
'StrSql = StrSql & " FROM " & gstrTestadaImobiliario & " A, "
'StrSql = StrSql & gstrImobiliario & " B, "
'StrSql = StrSql & gstrContribuinte & " C  "
'StrSql = StrSql & " WHERE "
'StrSql = StrSql & " B.PKId = A.intImobiliario"
'StrSql = StrSql & " AND C.PKId = B.intContribuinte "
'StrSql = StrSql & " AND intSecoes = " & dbcintSecaoDeLogradouro.BoundText
'StrSql = StrSql & " AND C.PKId = " & lngPKId
'
'Set gobjBanco = New clsBanco
'
'If gobjBanco.CriaADO(StrSql, 10, adoRec) Then
'    If adoRec!dblValor > 0 Then
'        RetornaValorTestada = adoRec!dblValor
'    Else
'        RetornaValorTestada = 0
'    End If
'Else
'    RetornaValorTestada = 0
'End If
'End Function

'Private Sub EfetuaCalculoContribuicaoMelhoria()
'Dim StrSql As String
'Dim adoRec As ADODB.Recordset
'Dim adoContribuinte As ADODB.Recordset
'Dim i As Integer
'Dim dblValor As Double
'Dim dblValorParcelado As Double
'Dim dblValorCalculado As Double
'Dim dblIndexador As Double
'Dim dblValorResto As Double
'Dim strMsg As String
'Dim lngSequencia As Long
'Dim datDataVencimento As Date
'Dim dblValorCusto As Double
'Dim dblCustoPorContribuinte As Double
'
'On Error GoTo err_EfetuaCalculoContribuicaoMelhoria
'
'If blnDadosOK Then
'    Screen.MousePointer = vbDefault
'    Set gobjBanco = New clsBanco
'
'    StrSql = "sp_EfetuaCalculo '" & strPKId & "'," & dbc_intComposicaoDaReceita.BoundText & ",11,0,NULL,0,0,0," & glngCodUsr
'
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(StrSql, 5, adoRec) Then
'        With adoRec
'            If Not (.BOF And .EOF) Then
'                dblValorCalculado = (!dblValorCalculado)
'                dblIndexador = (!dblIndexador)
'            End If
'        End With
'
'        'Pesquisa a sequência da composição da receita
'        StrSql = strQueryContribuicaoMelhoria
'        If Not gobjBanco.CriaADO(StrSql, 10, adoContribuinte) Then
'            Exit Sub
'        End If
'
'        If adoContribuinte.EOF And adoContribuinte.BOF Then
'            MsgBox "Não existem contribuintes para o Edital."
'            Exit Sub
'        End If
'
'        strMsg = ""
'        strMsg = strMsg & "Confirma o cálculo de " & gstrConvVrDoSql(dblValorCalculado + adoContribuinte!dblCustoDaParcela) & Chr(10)
'        strMsg = strMsg & "em " & txtintNumero.Text & " parcela(s) ?"
'
'        If Not gblnExclusaoGravacaoOk("", strMsg, True) Then
'            Exit Sub
'        End If
'
'        Screen.MousePointer = 11
'        dblValorCusto = RetornaCustoEdital(adoContribuinte!dblCustoDaParcela)
'
'        Do While Not adoContribuinte.EOF
'
'            dblCustoPorContribuinte = RetornaValorTestada(adoContribuinte!PKId)
'
'            dblCustoPorContribuinte = (dblCustoPorContribuinte * dblValorCusto)
'
'            StrSql = ""
'            StrSql = StrSql & " SELECT ISNULL(MAX(strSequencia),0) + 1 AS Maximo FROM " & gstrLancamentoCalculo
'            StrSql = StrSql & " WHERE intComposicaoReceita = " & dbc_intComposicaoDaReceita.BoundText
'            StrSql = StrSql & " AND intContribuinte = " & adoContribuinte!PKId
'            StrSql = StrSql & " AND intExercicio = " & gintExercicio
'
'            Set gobjBanco = New clsBanco
'            If gobjBanco.CriaADO(StrSql, 5, adoRec) Then
'                lngSequencia = adoRec!Maximo
'            Else
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            End If
'
'            Set gobjBanco = New clsBanco
'            gobjBanco.ExecutaBeginTrans
'
'            StrSql = ""
'            StrSql = StrSql & " INSERT INTO " & gstrLancamentoCalculo
'            StrSql = StrSql & " (intExercicio, intContribuinte, intComposicaoReceita, intMensagem, strInscricaoCadastral, "
'            StrSql = StrSql & " dtmLancamento, dtmVencimento, intNumeroDeParcelas, intIntervaloEntreParcelas, "
'            StrSql = StrSql & " bitUtilizacaoDebito, intOcorrencia, bytOrigem, strSequencia, dtmDtAtualizacao, lngCodUsr ) VALUES ( "
'            StrSql = StrSql & gintExercicio
'            StrSql = StrSql & ", " & adoContribuinte!PKId
'            StrSql = StrSql & ", " & dbc_intComposicaoDaReceita.BoundText
'
'            'If dbcintMensagem.MatchedWithList Then
'            '    strSql = strSql & ", " & dbcintMensagem.BoundText
'            'Else
'                StrSql = StrSql & ", NULL"
'            'End If
'
'            'Inscrição cadastral (Código do contribuinte para receitas diversas)
'            StrSql = StrSql & ", " & adoContribuinte!strInscricaoAnterior
'            StrSql = StrSql & ", " & gstrConvDtParaSql(txtdtmLancamento.Text)
'            StrSql = StrSql & ", " & gstrConvDtParaSql(txtdtmVencimento.Text)
'            StrSql = StrSql & ", " & Val(txtintNumero.Text)
'            StrSql = StrSql & ", " & Val(txtintIntervalo.Text)
'            StrSql = StrSql & ", 4" 'Utilização do débito = 4 - Receitas diversas
'            StrSql = StrSql & ", " & Val(dbc_intOcorrencia.BoundText) 'Ocorrência
'            StrSql = StrSql & " , 4" 'Origem
'            StrSql = StrSql & ", " & CStr(lngSequencia)
'            StrSql = StrSql & ", GETDATE()"
'            StrSql = StrSql & ", " & glngCodUsr
'            StrSql = StrSql & " ) "
'
'            'Gravar as Parcelas Taxas
'            StrSql = StrSql & " EXECUTE "
'            StrSql = StrSql & " sp_EfetuaCalculo '" & strPKId & "'," & dbc_intComposicaoDaReceita.BoundText & ",21,"
'            StrSql = StrSql & txtintNumero & "," & gstrConvDtParaSql(txtdtmVencimento) & "," & txtintIntervalo
'            StrSql = StrSql & ",0,0," & glngCodUsr
'            'Fim Gravar
'
'            dblValorParcelado = (dblValorCalculado + dblCustoPorContribuinte) / Val(txtintNumero.Text)
'            dblValor = 0
'
'            datDataVencimento = txtdtmVencimento.Text
'
'            For i = 1 To Val(txtintNumero.Text)
'                If i = Val(txtintNumero.Text) Then
'                    dblValorParcelado = (dblValorParcelado * Val(txtintNumero.Text)) - dblValorResto
'                Else
'                    dblValorResto = gstrConvVrDoSql(dblValorResto + dblValorParcelado)
'                    dblValor = gstrConvVrDoSql(dblValorParcelado)
'                End If
'
'                StrSql = StrSql & " INSERT INTO " & gstrParcelaReceita
'                StrSql = StrSql & " (intLancamentoCalculo, intComposicaoDaReceita, intNumeroParcela, dtmDataVencimento, "
'                StrSql = StrSql & " dblValorParcela, bytDividaAjuizada, bytSimulado, bytPrescrita, "
'                StrSql = StrSql & " bytCancelada, bytAtiva, bytSuspensaoDeExigencia, dtmDtAtualizacao, lngCodUsr) "
'                StrSql = StrSql & " (SELECT MAX(PKId) "
'                StrSql = StrSql & ", " & dbc_intComposicaoDaReceita.BoundText
'                StrSql = StrSql & ", " & i
'
'                StrSql = StrSql & ", " & gstrConvDtParaSql(datDataVencimento)
'                datDataVencimento = datDataVencimento + Val(txtintIntervalo.Text)
'
'                If i < Val(txtintNumero.Text) Then
'                    StrSql = StrSql & ", " & gstrConvVrParaSql(gstrConvVrDoSql(dblValor))
'                Else
'                    StrSql = StrSql & ", " & gstrConvVrParaSql(gstrConvVrDoSql(dblValorParcelado))
'                End If
'                StrSql = StrSql & ", 0" 'Dívida Ajuizada
'                StrSql = StrSql & ", 0" 'Simulado
'                StrSql = StrSql & ", 0" 'Prescrita
'                StrSql = StrSql & ", 0" 'Cancelada
'                StrSql = StrSql & ", 0" 'Divida Ativa
'                StrSql = StrSql & ", 0, GETDATE()"
'                StrSql = StrSql & ", " & glngCodUsr
'                StrSql = StrSql & " FROM " & gstrLancamentoCalculo & ") "
'            Next i
'
'            Set gobjBanco = New clsBanco
'            If gobjBanco.Execute(StrSql, False) Then
'                gobjBanco.ExecutaCommitTrans
'            Else
'                gobjBanco.ExecutaRollbackTrans
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            End If
'            adoContribuinte.MoveNext
'        Loop
'    Else
'        Exit Sub
'    End If
'End If
'ExibeMensagem "Cálculo efetuado com sucesso!"
'Screen.MousePointer = vbDefault
'
'err_EfetuaCalculoContribuicaoMelhoria:

'End Sub

Private Function strPKId() As String
    Dim strSQL As String
    Dim i As Integer
    strSQL = ""
    For i = 0 To xarReceita.Count(1) - 1

        If xarReceita(i, 2) = -1 Then
            If strSQL <> "" Then
               strSQL = strSQL & ","
            End If
        strSQL = strSQL & xarReceita(i, 0)
        End If
    Next
   
    strPKId = strSQL
End Function

Private Function strQueryInscricao() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSQL As String
'    strSQL = "SELECT I.PKId AS PKId, (I.strInscricaoAnterior + ' - ' + CO.strNome) "
    strSQL = "SELECT I.PKId AS PKId, (I.strInscricaoAnterior " & strCONCAT & " ' - ' " & strCONCAT & " CO.strNome) " & _
            "AS Inscricao  FROM " & gstrImobiliario & " I, " & gstrContribuinte & " CO " & _
            " WHERE I.intSecoes = " & dbcintSecaoDeLogradouro.BoundText & _
            " AND CO.PKId = I.intContribuinte "
strQueryInscricao = strSQL
End Function

Private Sub LimpaCampos()
dbc_intOcorrencia.Text = ""
dbc_intComposicaoDaReceita.Text = ""
dbcintTabelaDeEdital.Text = ""
dbcintSecaoDeLogradouro.Text = ""
Set dbcintSecaoDeLogradouro.RowSource = Nothing
dbcintSecaoDeLogradouro.Refresh

Set xarReceita = New XArrayDB
xarReceita.Clear
xarReceita.ReDim 0, 0, 0, 2

Set tdb_Atividades.Array = xarReceita
tdb_Atividades.Rebind
tdb_Atividades.Refresh

dbc_intOcorrencia.SetFocus
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    Select Case tab_3dPasta.Tab
        Case 0
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
        Case 1
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrNovo
    End Select
End Sub

Private Sub txtdtmLancamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmLancamento
End Sub

Private Sub txtdtmLancamento_LostFocus()
    txtdtmLancamento = gstrDataFormatada(txtdtmLancamento)
End Sub

Private Sub txtdtmVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmVencimento
End Sub

Private Sub txtdtmVencimento_LostFocus()
    txtdtmVencimento.Text = gstrDataFormatada(txtdtmVencimento.Text)
End Sub

Private Sub txtintIntervalo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintIntervalo
End Sub


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
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT strMensagem "
    strSQL = strSQL & " FROM " & gstrMensagem
    strSQL = strSQL & " WHERE PKId = " & Val(dbc_intMensagem1.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem1.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem1.Text = ""
        End If
    End If
End Function

Private Function LeDoComboParaTXT2()
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT strMensagem "
    strSQL = strSQL & " FROM " & gstrMensagem
    strSQL = strSQL & " WHERE PKId = " & Val(dbc_intMensagem2.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
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

Dim strSQL As String

    strSQL = ""
'    strSQL = strSQL & "SELECT PKId, ltrim(rtrim(PKId)) + ' - ' + ltrim(rtrim(strDescricao)) as Descricao "
    strSQL = strSQL & "SELECT PKId, ltrim(rtrim(PKId)) " & strCONCAT & " ' - ' " & strCONCAT & " ltrim(rtrim(strDescricao)) as Descricao "
    strSQL = strSQL & " FROM " & gstrMensagem
    strSQL = strSQL & " ORDER BY PKId "

strQueryMensagem = strSQL
End Function

Private Function blnDadosGuiaOK() As Boolean
    
    If (txt_intParcelaInicial.Text = "") Or (txt_intParcelaFinal.Text = "") Then
        ExibeMensagem "O intervalo de parcelas deve ser definido "
        If txt_intParcelaInicial.Text = "" Then
            txt_intParcelaInicial.SetFocus
        Else
            txt_intParcelaFinal.SetFocus
        End If
        Exit Function
    End If
    If Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal) Then
        ExibeMensagem "O número da parcela final deve ser maior que o número da parcela inicial"
        txt_intParcelaFinal.SetFocus
        Exit Function
    End If
   
    If dbc_intComposicaoDaReceita.BoundText = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Composição da Receita deve ser selecionada."
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If
    
    If txt_intExercicio.Text = "" Then
        ExibeMensagem "O Exercício deve ser Digitado."
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

