VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadLancamentoIPTU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento IPTU"
   ClientHeight    =   8595
   ClientLeft      =   660
   ClientTop       =   1935
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11325
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6675
      Left            =   30
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   30
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   11774
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "Lançamento IPTU"
      TabPicture(0)   =   "frmCadLancamentoIPTU.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Contribuinte"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtPkid"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtPkidIPTU"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Valor Venal e Tributos"
      TabPicture(1)   =   "frmCadLancamentoIPTU.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_ValorVenalETributos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Prédios/Características"
      TabPicture(2)   =   "frmCadLancamentoIPTU.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Parcelas"
      TabPicture(3)   =   "frmCadLancamentoIPTU.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra_Parcelas(1)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fra_Parcelas 
         Height          =   6195
         Index           =   1
         Left            =   -74940
         TabIndex        =   108
         Top             =   360
         Width           =   11115
         Begin VB.Frame fra_Cabecalho 
            Height          =   615
            Index           =   3
            Left            =   1040
            TabIndex        =   115
            Top             =   120
            Width           =   9090
            Begin VB.TextBox txtstrEmissao4 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   6015
               MaxLength       =   4
               TabIndex        =   117
               Top             =   225
               Width           =   570
            End
            Begin VB.TextBox txtstrNumDoAviso4 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   7140
               MaxLength       =   6
               TabIndex        =   116
               Top             =   225
               Width           =   975
            End
            Begin MSMask.MaskEdBox mskstrInscricao4 
               Height          =   300
               Left            =   1590
               TabIndex        =   118
               Top             =   240
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   24
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo dbcintExercicio4 
               Height          =   315
               Left            =   4260
               TabIndex        =   119
               Top             =   225
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbl_Exercicio 
               AutoSize        =   -1  'True
               Caption         =   "Exercício"
               Height          =   195
               Index           =   3
               Left            =   3540
               TabIndex        =   123
               Top             =   315
               Width           =   675
            End
            Begin VB.Label lbl_strInscricaoAnterior 
               AutoSize        =   -1  'True
               Caption         =   "Inscrição Cadastral"
               Height          =   195
               Index           =   3
               Left            =   90
               TabIndex        =   122
               Top             =   300
               Width           =   1350
            End
            Begin VB.Label lbl_Emissao 
               AutoSize        =   -1  'True
               Caption         =   "Emissão"
               Height          =   195
               Index           =   3
               Left            =   5385
               TabIndex        =   121
               Top             =   315
               Width           =   585
            End
            Begin VB.Label lbl_Aviso 
               AutoSize        =   -1  'True
               Caption         =   "Aviso"
               Height          =   195
               Index           =   3
               Left            =   6690
               TabIndex        =   120
               Top             =   315
               Width           =   390
            End
         End
         Begin VB.Frame fra_Parcelas 
            Caption         =   "Parcelas"
            Height          =   5115
            Index           =   0
            Left            =   90
            TabIndex        =   109
            Top             =   690
            Width           =   10950
            Begin TrueOleDBGrid70.TDBGrid tdb_Parcelas 
               Height          =   4635
               Left            =   150
               TabIndex        =   110
               Top             =   300
               Width           =   10755
               _ExtentX        =   18971
               _ExtentY        =   8176
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Parcela"
               Columns(0).DataField=   "intParcela"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Acordo"
               Columns(1).DataField=   "strAcordo"
               Columns(1).NumberFormat=   "FormatText Event"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).DataField=   "strMoeda"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "Valor"
               Columns(3).DataField=   "dblValor"
               Columns(3).NumberFormat=   "Standard"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "Vencimento"
               Columns(4).DataField=   "dtmDtVencimento"
               Columns(4).NumberFormat=   "FormatText Event"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "D.A."
               Columns(5).DataField=   "intLancamentoAlfaDAtiva"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "Baixa"
               Columns(6).DataField=   "dtmDtPagamento"
               Columns(6).NumberFormat=   "FormatText Event"
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).Caption=   "Descrição da Baixa"
               Columns(7).DataField=   "STRDESCRICAO"
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(8)._VlistStyle=   0
               Columns(8)._MaxComboItems=   5
               Columns(8).Caption=   "Observação"
               Columns(8).DataField=   "Strobservacao"
               Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   9
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=9"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
               Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=1058"
               Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
               Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(8)=   "Column(1).Width=2672"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerStyle=0"
               Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2619"
               Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
               Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=2"
               Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(15)=   "Column(2).Width=609"
               Splits(0)._ColumnProps(16)=   "Column(2).DividerStyle=0"
               Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=556"
               Splits(0)._ColumnProps(19)=   "Column(2).AllowSizing=0"
               Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(21)=   "Column(3).Width=1879"
               Splits(0)._ColumnProps(22)=   "Column(3).DividerStyle=0"
               Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1826"
               Splits(0)._ColumnProps(25)=   "Column(3).AllowSizing=0"
               Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=2"
               Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(28)=   "Column(4).Width=1746"
               Splits(0)._ColumnProps(29)=   "Column(4).DividerStyle=0"
               Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1693"
               Splits(0)._ColumnProps(32)=   "Column(4).AllowSizing=0"
               Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=1"
               Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(35)=   "Column(5).Width=688"
               Splits(0)._ColumnProps(36)=   "Column(5).DividerStyle=0"
               Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=635"
               Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=1"
               Splits(0)._ColumnProps(40)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(41)=   "Column(6).Width=2514"
               Splits(0)._ColumnProps(42)=   "Column(6).DividerStyle=0"
               Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=2461"
               Splits(0)._ColumnProps(45)=   "Column(6).AllowSizing=0"
               Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=1"
               Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(48)=   "Column(7).Width=3228"
               Splits(0)._ColumnProps(49)=   "Column(7).DividerStyle=0"
               Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=3175"
               Splits(0)._ColumnProps(52)=   "Column(7).AllowSizing=0"
               Splits(0)._ColumnProps(53)=   "Column(7)._ColStyle=0"
               Splits(0)._ColumnProps(54)=   "Column(7).Order=8"
               Splits(0)._ColumnProps(55)=   "Column(8).Width=4921"
               Splits(0)._ColumnProps(56)=   "Column(8).DividerStyle=0"
               Splits(0)._ColumnProps(57)=   "Column(8).DividerColor=0"
               Splits(0)._ColumnProps(58)=   "Column(8)._WidthInPix=4868"
               Splits(0)._ColumnProps(59)=   "Column(8).AllowSizing=0"
               Splits(0)._ColumnProps(60)=   "Column(8)._ColStyle=0"
               Splits(0)._ColumnProps(61)=   "Column(8).Order=9"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   0
               BorderStyle     =   0
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   0
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   -2147483633
               RowDividerColor =   12648447
               RowSubDividerColor=   13160660
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000004&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&,.bold=0"
               _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000004&"
               _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1,.transparentBmp=-1"
               _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
               _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
               _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
               _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
               _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
               _StyleDefs(30)  =   "Splits(0).EditorStyle:id=59,.parent=7"
               _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
               _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
               _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
               _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
               _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=47,.alignment=2"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=48"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=49"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=59"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=24,.parent=47,.alignment=1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=48,.wraptext=-1"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=49"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=59"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=47,.wraptext=0,.transparentBmp=-1"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=48,.wraptext=-1"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=49"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=59"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=16,.parent=47,.alignment=1"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=13,.parent=48"
               _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=14,.parent=49"
               _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=15,.parent=59"
               _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=47,.alignment=2"
               _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=48"
               _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=49"
               _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=59"
               _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=20,.parent=47,.alignment=2"
               _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=17,.parent=48"
               _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=18,.parent=49"
               _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=19,.parent=59"
               _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=47,.alignment=2"
               _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=48"
               _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=49"
               _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=59"
               _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=82,.parent=47,.alignment=0"
               _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=48"
               _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=49"
               _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=59"
               _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=86,.parent=47,.alignment=0"
               _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=83,.parent=48"
               _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=84,.parent=49"
               _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=85,.parent=59"
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
               _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(85)  =   "Named:id=39:EvenRow"
               _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(87)  =   "Named:id=40:OddRow"
               _StyleDefs(88)  =   ":id=40,.parent=33"
               _StyleDefs(89)  =   "Named:id=41:RecordSelector"
               _StyleDefs(90)  =   ":id=41,.parent=34"
               _StyleDefs(91)  =   "Named:id=42:FilterBar"
               _StyleDefs(92)  =   ":id=42,.parent=33"
            End
         End
      End
      Begin VB.TextBox txtPkidIPTU 
         Height          =   315
         Left            =   7890
         TabIndex        =   75
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPkid 
         Height          =   315
         Left            =   6270
         TabIndex        =   74
         Top             =   -15
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Frame Frame2 
         Height          =   6105
         Left            =   -74940
         TabIndex        =   83
         Top             =   360
         Width           =   11025
         Begin VB.Frame fra_CaracBoletim 
            Caption         =   "Características/Boletim"
            Height          =   2070
            Left            =   1040
            TabIndex        =   103
            Top             =   2475
            Width           =   9090
            Begin TrueOleDBGrid70.TDBGrid tdb_CaracBoletim 
               Height          =   1395
               Left            =   90
               TabIndex        =   104
               Top             =   375
               Width           =   8820
               _ExtentX        =   15558
               _ExtentY        =   2461
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Característica"
               Columns(0).DataField=   "strNomeCaracGeral"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Detalhe"
               Columns(1).DataField=   "strNomeDoDetalhe"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Pontos"
               Columns(2).DataField=   "dblValorDetalhe"
               Columns(2).NumberFormat=   "Standard"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=7567"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
               Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=7514"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=5689"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerStyle=0"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5636"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=1746"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerStyle=0"
               Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1693"
               Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=2"
               Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   0
               BorderStyle     =   0
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   0
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   -2147483633
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000000&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(11)  =   ":id=2,.strikethrough=0,.charset=0"
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
               _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.transparentBmp=-1"
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
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.alignment=1"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
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
         Begin VB.Frame fra_Prédios 
            Caption         =   "Prédios"
            Height          =   1560
            Left            =   1040
            TabIndex        =   101
            Top             =   735
            Width           =   9090
            Begin TrueOleDBGrid70.TDBGrid tdb_PrediosCaracteristicas 
               Height          =   1185
               Left            =   120
               TabIndex        =   102
               Top             =   225
               Width           =   8955
               _ExtentX        =   15796
               _ExtentY        =   2090
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   2
               Columns(0)._MaxComboItems=   5
               Columns(0).DataField=   "Option"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   18
               Columns(1)._MaxComboItems=   5
               Columns(1).DataField=   "Area"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Área"
               Columns(2).DataField=   "dblMedidaDaArea"
               Columns(2).NumberFormat=   "Standard"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "Pontos"
               Columns(3).DataField=   "dblPontos"
               Columns(3).NumberFormat=   "Standard"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "Valor Metro"
               Columns(4).DataField=   "dblValorMetro"
               Columns(4).NumberFormat=   "Standard"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "Fator"
               Columns(5).DataField=   "dblFatorObsolescencia"
               Columns(5).NumberFormat=   "Standard"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "Valor Venal"
               Columns(6).DataField=   "dblValorVenalPredio"
               Columns(6).NumberFormat=   "Standard"
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).Caption=   "Pkid"
               Columns(7).DataField=   "Pkid"
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   8
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=8"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=370"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
               Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=318"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=3651"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerStyle=0"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3598"
               Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=2"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(12)=   "Column(2).Width=1640"
               Splits(0)._ColumnProps(13)=   "Column(2).DividerStyle=0"
               Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1588"
               Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=2"
               Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(18)=   "Column(3).Width=979"
               Splits(0)._ColumnProps(19)=   "Column(3).DividerStyle=0"
               Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=926"
               Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=2"
               Splits(0)._ColumnProps(23)=   "Column(3).Visible=0"
               Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(25)=   "Column(4).Width=3122"
               Splits(0)._ColumnProps(26)=   "Column(4).DividerStyle=0"
               Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=3069"
               Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=2"
               Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(31)=   "Column(5).Width=3228"
               Splits(0)._ColumnProps(32)=   "Column(5).DividerStyle=0"
               Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=3175"
               Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=0"
               Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(37)=   "Column(6).Width=2858"
               Splits(0)._ColumnProps(38)=   "Column(6).DividerStyle=0"
               Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2805"
               Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=2"
               Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(43)=   "Column(7).Width=2699"
               Splits(0)._ColumnProps(44)=   "Column(7).DividerStyle=0"
               Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2646"
               Splits(0)._ColumnProps(47)=   "Column(7).Visible=0"
               Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   0
               BorderStyle     =   0
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   0
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   -2147483633
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000F&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
               _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.transparentBmp=-1"
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
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
               _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
               _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
               _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
               _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
               _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
               _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=0"
               _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
               _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
               _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
               _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
               _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
               _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
               _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
               _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
               _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
               _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
               _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
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
               _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
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
         Begin VB.Frame fra_Cabecalho 
            Height          =   615
            Index           =   2
            Left            =   1040
            TabIndex        =   66
            Top             =   120
            Width           =   9090
            Begin VB.TextBox txtstrNumDoAviso3 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   7140
               MaxLength       =   6
               TabIndex        =   32
               Top             =   225
               Width           =   975
            End
            Begin VB.TextBox txtstrEmissao3 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   6015
               MaxLength       =   4
               TabIndex        =   31
               Top             =   225
               Width           =   570
            End
            Begin MSMask.MaskEdBox mskstrInscricao3 
               Height          =   300
               Left            =   1590
               TabIndex        =   29
               Top             =   240
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   24
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo dbcintExercicio3 
               Height          =   315
               Left            =   4260
               TabIndex        =   30
               Top             =   225
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbl_Aviso 
               AutoSize        =   -1  'True
               Caption         =   "Aviso"
               Height          =   195
               Index           =   2
               Left            =   6690
               TabIndex        =   70
               Top             =   315
               Width           =   390
            End
            Begin VB.Label lbl_Emissao 
               AutoSize        =   -1  'True
               Caption         =   "Emissão"
               Height          =   195
               Index           =   2
               Left            =   5385
               TabIndex        =   69
               Top             =   315
               Width           =   585
            End
            Begin VB.Label lbl_strInscricaoAnterior 
               AutoSize        =   -1  'True
               Caption         =   "Inscrição Cadastral"
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   68
               Top             =   300
               Width           =   1350
            End
            Begin VB.Label lbl_Exercicio 
               AutoSize        =   -1  'True
               Caption         =   "Exercício"
               Height          =   195
               Index           =   2
               Left            =   3540
               TabIndex        =   67
               Top             =   315
               Width           =   675
            End
         End
      End
      Begin VB.Frame fra_ValorVenalETributos 
         Height          =   6225
         Left            =   -74940
         TabIndex        =   60
         Top             =   360
         Width           =   11115
         Begin VB.Frame fra_Demonstrativo 
            Caption         =   "Demonstrativo"
            Height          =   3420
            Left            =   75
            TabIndex        =   82
            Top             =   735
            Width           =   10950
            Begin VB.TextBox txt_dblTotaisValorVenal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               CausesValidation=   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   6600
               Locked          =   -1  'True
               TabIndex        =   125
               Top             =   3030
               Width           =   1770
            End
            Begin VB.TextBox txt_dblTotaisImposto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   8880
               Locked          =   -1  'True
               TabIndex        =   124
               Top             =   3030
               Width           =   1635
            End
            Begin VB.TextBox txtdblTotalImposto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   9195
               Locked          =   -1  'True
               TabIndex        =   113
               Top             =   2685
               Width           =   1320
            End
            Begin VB.TextBox txtdblTotalValorVenal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   7050
               Locked          =   -1  'True
               TabIndex        =   112
               Top             =   2685
               Width           =   1320
            End
            Begin VB.TextBox txtdblTotalArea 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   1755
               Locked          =   -1  'True
               TabIndex        =   111
               Top             =   2685
               Width           =   1320
            End
            Begin VB.TextBox txtdblImpostoTerreno 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   9105
               Locked          =   -1  'True
               TabIndex        =   106
               Top             =   540
               Width           =   1410
            End
            Begin VB.TextBox txtdblValorVenalTerreno 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               CausesValidation=   0   'False
               Height          =   300
               Left            =   6600
               Locked          =   -1  'True
               TabIndex        =   84
               Top             =   540
               Width           =   1770
            End
            Begin VB.TextBox txtdblAreaTerreno 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   92
               Top             =   540
               Width           =   1500
            End
            Begin VB.TextBox txtdblAliquotaTerreno 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   8415
               Locked          =   -1  'True
               TabIndex        =   91
               Top             =   540
               Width           =   660
            End
            Begin VB.TextBox txtdblAreaExcedente 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   90
               Top             =   1110
               Width           =   1500
            End
            Begin VB.TextBox txtdblValorTerrenoExcedente 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   6600
               Locked          =   -1  'True
               TabIndex        =   89
               Top             =   1110
               Width           =   1770
            End
            Begin VB.TextBox txtdblAliquotaExcedente 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   8415
               Locked          =   -1  'True
               TabIndex        =   88
               Top             =   1110
               Width           =   660
            End
            Begin VB.TextBox txtdblImpostoExcedente 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   9105
               Locked          =   -1  'True
               TabIndex        =   87
               Top             =   1110
               Width           =   1410
            End
            Begin VB.TextBox txtdblValorMetro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   3060
               Locked          =   -1  'True
               TabIndex        =   85
               Top             =   540
               Width           =   990
            End
            Begin TrueOleDBGrid70.TDBGrid tdb_Predios 
               Height          =   960
               Left            =   90
               TabIndex        =   86
               Top             =   1650
               Width           =   10755
               _ExtentX        =   18971
               _ExtentY        =   1693
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "strNomeDaArea"
               Columns(0).DataField=   "strNomeDaArea"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "dblMedidaDaArea"
               Columns(1).DataField=   "dblMedidaDaArea"
               Columns(1).NumberFormat=   "Standard"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "dblValorMetro"
               Columns(2).DataField=   "dblValorMetro"
               Columns(2).NumberFormat=   "Standard"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "strNomeFator"
               Columns(3).DataField=   "strNomeFator"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "dblFatorObsolescencia"
               Columns(4).DataField=   "dblFatorObsolescencia"
               Columns(4).NumberFormat=   "Standard"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "dblValorVenalPredio"
               Columns(5).DataField=   "dblValorVenalPredio"
               Columns(5).NumberFormat=   "Standard"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "dblAliquota"
               Columns(6).DataField=   "dblAliquota"
               Columns(6).NumberFormat=   "Standard"
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).Caption=   "dblImposto"
               Columns(7).DataField=   "dblImposto"
               Columns(7).NumberFormat=   "Standard"
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   8
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=8"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2884"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
               Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2831"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2408"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerStyle=0"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2355"
               Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=2"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(13)=   "Column(2).Width=1746"
               Splits(0)._ColumnProps(14)=   "Column(2).DividerStyle=0"
               Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1693"
               Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=2"
               Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(19)=   "Column(3).Width=2275"
               Splits(0)._ColumnProps(20)=   "Column(3).DividerStyle=0"
               Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2223"
               Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(24)=   "Column(4).Width=1614"
               Splits(0)._ColumnProps(25)=   "Column(4).DividerStyle=0"
               Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1561"
               Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=2"
               Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(30)=   "Column(5).Width=3757"
               Splits(0)._ColumnProps(31)=   "Column(5).DividerStyle=0"
               Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=3704"
               Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=2"
               Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(36)=   "Column(6).Width=1191"
               Splits(0)._ColumnProps(37)=   "Column(6).DividerStyle=0"
               Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1138"
               Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=2"
               Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(42)=   "Column(7).Width=2540"
               Splits(0)._ColumnProps(43)=   "Column(7).DividerStyle=0"
               Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2487"
               Splits(0)._ColumnProps(46)=   "Column(7)._ColStyle=2"
               Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   2
               BorderStyle     =   0
               ColumnHeaders   =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   0
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   -2147483633
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000F&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
               _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
               _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
               _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
               _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
               _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
               _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
               _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
               _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
               _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
               _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
               _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
               _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
               _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
               _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
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
               _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(81)  =   "Named:id=39:EvenRow"
               _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(83)  =   "Named:id=40:OddRow"
               _StyleDefs(84)  =   ":id=40,.parent=33"
               _StyleDefs(85)  =   "Named:id=41:RecordSelector"
               _StyleDefs(86)  =   ":id=41,.parent=34"
               _StyleDefs(87)  =   "Named:id=42:FilterBar"
               _StyleDefs(88)  =   ":id=42,.parent=33"
            End
            Begin TrueOleDBGrid70.TDBGrid tdb_Fatores 
               Height          =   960
               Left            =   4110
               TabIndex        =   93
               Top             =   540
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   1693
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "strDescricao"
               Columns(0).DataField=   "strDescricao"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "dblFator"
               Columns(1).DataField=   "dblFator"
               Columns(1).NumberFormat=   "Standard"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   2
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2249"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
               Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=2196"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=1588"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerStyle=0"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1535"
               Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=2"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               BorderStyle     =   0
               ColumnHeaders   =   0   'False
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   0
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   -2147483633
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H8000000F&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
               _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.transparentBmp=-1"
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
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
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
            Begin VB.Label lbl_Imposto 
               AutoSize        =   -1  'True
               Caption         =   "Imposto"
               Height          =   195
               Left            =   9930
               TabIndex        =   107
               Top             =   135
               Width           =   555
            End
            Begin VB.Label lbl_Area 
               AutoSize        =   -1  'True
               Caption         =   "Área"
               Height          =   195
               Left            =   2700
               TabIndex        =   100
               Top             =   135
               Width           =   330
            End
            Begin VB.Label lbl_ValorMetro 
               AutoSize        =   -1  'True
               Caption         =   "Valor Metro"
               Height          =   195
               Left            =   3240
               TabIndex        =   99
               Top             =   135
               Width           =   810
            End
            Begin VB.Label lbl_Fatores 
               AutoSize        =   -1  'True
               Caption         =   "Fatores"
               Height          =   195
               Left            =   4080
               TabIndex        =   98
               Top             =   135
               Width           =   525
            End
            Begin VB.Label lbl_ValorTotal 
               AutoSize        =   -1  'True
               Caption         =   "Valor Venal"
               Height          =   195
               Left            =   7560
               TabIndex        =   97
               Top             =   135
               Width           =   810
            End
            Begin VB.Label lbl_Aliquota 
               AutoSize        =   -1  'True
               Caption         =   "Alíquota"
               Height          =   195
               Left            =   8475
               TabIndex        =   96
               Top             =   135
               Width           =   600
            End
            Begin VB.Label lbl_Terreno 
               AutoSize        =   -1  'True
               Caption         =   "Terreno"
               Height          =   195
               Left            =   60
               TabIndex        =   95
               Top             =   555
               Width           =   555
            End
            Begin VB.Label lbl_Excedente 
               AutoSize        =   -1  'True
               Caption         =   "Excedente"
               Height          =   195
               Left            =   60
               TabIndex        =   94
               Top             =   1080
               Width           =   765
            End
         End
         Begin VB.Frame fra_Tributos 
            Caption         =   "Tributos"
            Height          =   2010
            Left            =   3045
            TabIndex        =   76
            Top             =   4155
            Width           =   5235
            Begin VB.TextBox txtdblCompensacao 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   3795
               Locked          =   -1  'True
               TabIndex        =   79
               Top             =   1470
               Width           =   1080
            End
            Begin VB.TextBox txtTotalTributo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   3795
               Locked          =   -1  'True
               TabIndex        =   81
               Top             =   1680
               Width           =   1080
            End
            Begin TrueOleDBGrid70.TDBGrid tdb_Tributos 
               Height          =   1110
               Left            =   150
               TabIndex        =   77
               Top             =   300
               Width           =   4995
               _ExtentX        =   8811
               _ExtentY        =   1958
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Tributo"
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Valor"
               Columns(1).DataField=   ""
               Columns(1).NumberFormat=   "Standard"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   2
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=5874"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerStyle=0"
               Splits(0)._ColumnProps(3)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(4)=   "Column(0)._WidthInPix=5821"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=2487"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerStyle=0"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2434"
               Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8194"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   0
               BorderStyle     =   0
               DataMode        =   4
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               RowDividerStyle =   0
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   -2147483633
               RowDividerColor =   13160660
               RowSubDividerColor=   13160660
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000004&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
               _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.transparentBmp=-1"
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
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=1,.locked=-1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
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
            Begin VB.Label lbl_ValorCompensacao 
               AutoSize        =   -1  'True
               Caption         =   "Compensação"
               Height          =   195
               Left            =   180
               TabIndex        =   78
               Top             =   1485
               Width           =   1020
            End
            Begin VB.Label lbl_ValorTributoTotal 
               AutoSize        =   -1  'True
               Caption         =   "Total"
               Height          =   195
               Left            =   180
               TabIndex        =   80
               Top             =   1695
               Width           =   360
            End
         End
         Begin VB.Frame fra_Cabecalho 
            Height          =   615
            Index           =   1
            Left            =   1040
            TabIndex        =   61
            Top             =   120
            Width           =   9090
            Begin VB.TextBox txtstrEmissao2 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   6015
               MaxLength       =   4
               TabIndex        =   27
               Top             =   225
               Width           =   570
            End
            Begin VB.TextBox txtstrNumDoAviso2 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   7140
               MaxLength       =   6
               TabIndex        =   28
               Top             =   225
               Width           =   975
            End
            Begin MSMask.MaskEdBox mskstrInscricao2 
               Height          =   300
               Left            =   1590
               TabIndex        =   25
               Top             =   240
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   24
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo dbcintExercicio2 
               Height          =   315
               Left            =   4260
               TabIndex        =   26
               Top             =   225
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbl_Exercicio 
               AutoSize        =   -1  'True
               Caption         =   "Exercício"
               Height          =   195
               Index           =   1
               Left            =   3540
               TabIndex        =   65
               Top             =   315
               Width           =   675
            End
            Begin VB.Label lbl_strInscricaoAnterior 
               AutoSize        =   -1  'True
               Caption         =   "Inscrição Cadastral"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   64
               Top             =   300
               Width           =   1350
            End
            Begin VB.Label lbl_Emissao 
               AutoSize        =   -1  'True
               Caption         =   "Emissão"
               Height          =   195
               Index           =   1
               Left            =   5385
               TabIndex        =   63
               Top             =   315
               Width           =   585
            End
            Begin VB.Label lbl_Aviso 
               AutoSize        =   -1  'True
               Caption         =   "Aviso"
               Height          =   195
               Index           =   1
               Left            =   6690
               TabIndex        =   62
               Top             =   315
               Width           =   390
            End
         End
      End
      Begin VB.Frame fra_Contribuinte 
         Height          =   6180
         Left            =   1005
         TabIndex        =   34
         Top             =   360
         Width           =   9285
         Begin VB.Frame Fra_Isencoes 
            Caption         =   "Isenção Imunidade"
            Height          =   615
            Left            =   90
            TabIndex        =   126
            Top             =   1200
            Width           =   9000
            Begin VB.TextBox txtstrtipoisencaoimunidade 
               Height          =   315
               Left            =   5310
               Locked          =   -1  'True
               MaxLength       =   35
               TabIndex        =   130
               Top             =   210
               Width           =   3540
            End
            Begin VB.TextBox txtstrdefinicaoisencao 
               Height          =   315
               Left            =   1020
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   128
               Top             =   210
               Width           =   2490
            End
            Begin VB.Label lbl_strtipoisencaoimunidade 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Isenção Imunidade"
               Height          =   195
               Left            =   3570
               TabIndex        =   129
               Top             =   270
               Width           =   1710
            End
            Begin VB.Label lbl_strdefinicaoisencao 
               AutoSize        =   -1  'True
               Caption         =   "Definição"
               Height          =   195
               Left            =   240
               TabIndex        =   127
               Top             =   270
               Width           =   675
            End
         End
         Begin VB.TextBox txtstrSequenciaDeFace 
            Height          =   300
            Left            =   4590
            TabIndex        =   7
            Top             =   1965
            Width           =   600
         End
         Begin VB.TextBox txtstrPromissario 
            Height          =   300
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   10
            Top             =   2655
            Width           =   6885
         End
         Begin VB.TextBox txtstrNomeProprietario 
            Height          =   300
            Left            =   1125
            MaxLength       =   100
            TabIndex        =   9
            Top             =   2280
            Width           =   6885
         End
         Begin VB.TextBox txtstrLoteamento 
            Height          =   300
            Left            =   6240
            MaxLength       =   35
            TabIndex        =   8
            Top             =   1935
            Width           =   2940
         End
         Begin VB.TextBox txtstrQuadra 
            Height          =   300
            Left            =   1125
            MaxLength       =   20
            TabIndex        =   5
            Top             =   1920
            Width           =   1290
         End
         Begin VB.TextBox txtintLote 
            Height          =   300
            Left            =   2925
            MaxLength       =   4
            TabIndex        =   6
            Top             =   1935
            Width           =   660
         End
         Begin VB.Frame fra_Cabecalho 
            Height          =   1080
            Index           =   0
            Left            =   90
            TabIndex        =   35
            Top             =   120
            Width           =   9000
            Begin VB.TextBox txtDtmCancelamento 
               Height          =   315
               Left            =   7050
               TabIndex        =   105
               Top             =   675
               Width           =   1230
            End
            Begin VB.TextBox txtstrComposicaoDaReceita 
               Height          =   315
               Left            =   1575
               TabIndex        =   4
               Top             =   675
               Width           =   3150
            End
            Begin VB.TextBox txtstrNumeroAviso 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   7140
               MaxLength       =   10
               TabIndex        =   3
               Top             =   225
               Width           =   1125
            End
            Begin VB.TextBox txtstrEmissao 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   6015
               MaxLength       =   4
               TabIndex        =   2
               Top             =   225
               Width           =   570
            End
            Begin MSMask.MaskEdBox mskstrInscricao 
               Height          =   300
               Left            =   1590
               TabIndex        =   0
               Top             =   240
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   24
               PromptChar      =   " "
            End
            Begin MSDataListLib.DataCombo dbcintExercicio 
               Height          =   315
               Left            =   4260
               TabIndex        =   1
               Top             =   225
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbldtmCancelamento 
               AutoSize        =   -1  'True
               Caption         =   "Data de Cancelamento"
               Height          =   195
               Left            =   5340
               TabIndex        =   72
               Top             =   750
               Width           =   1635
            End
            Begin VB.Label lblintComposicao 
               AutoSize        =   -1  'True
               Caption         =   "Composição"
               Height          =   195
               Left            =   585
               TabIndex        =   71
               Top             =   750
               Width           =   870
            End
            Begin VB.Label lbl_Aviso 
               AutoSize        =   -1  'True
               Caption         =   "Aviso"
               Height          =   195
               Index           =   0
               Left            =   6690
               TabIndex        =   39
               Top             =   315
               Width           =   390
            End
            Begin VB.Label lbl_Emissao 
               AutoSize        =   -1  'True
               Caption         =   "Emissão"
               Height          =   195
               Index           =   0
               Left            =   5385
               TabIndex        =   38
               Top             =   315
               Width           =   585
            End
            Begin VB.Label lbl_strInscricaoAnterior 
               AutoSize        =   -1  'True
               Caption         =   "Inscrição Cadastral"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   37
               Top             =   300
               Width           =   1350
            End
            Begin VB.Label lbl_Exercicio 
               AutoSize        =   -1  'True
               Caption         =   "Exercício"
               Height          =   195
               Index           =   0
               Left            =   3540
               TabIndex        =   36
               Top             =   315
               Width           =   675
            End
         End
         Begin MSComctlLib.ListView lvwEnvolvidos 
            Height          =   1305
            Left            =   1110
            TabIndex        =   11
            Top             =   3000
            Width           =   8070
            _ExtentX        =   14235
            _ExtentY        =   2302
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "Envolvido"
               Text            =   "Contribuinte"
               Object.Width           =   9878
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Key             =   "bitProprietario"
               Text            =   "Vinculo"
               Object.Width           =   2293
            EndProperty
         End
         Begin TabDlg.SSTab tab_3dEnderecos 
            Height          =   1680
            Left            =   105
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   4395
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   2963
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Endereço do Imobiliário"
            TabPicture(0)   =   "frmCadLancamentoIPTU.frx":0070
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fra_EndImobiliario"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Endereço de Notificação"
            TabPicture(1)   =   "frmCadLancamentoIPTU.frx":008C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame1"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame1 
               Height          =   1110
               Left            =   -74850
               TabIndex        =   52
               Top             =   315
               Width           =   8655
               Begin VB.TextBox txtstrMunicipioC 
                  Height          =   300
                  Left            =   4185
                  MaxLength       =   50
                  TabIndex        =   21
                  Top             =   735
                  Width           =   2235
               End
               Begin VB.TextBox txtstrUFC 
                  Height          =   300
                  Left            =   6765
                  MaxLength       =   2
                  TabIndex        =   22
                  Top             =   735
                  Width           =   375
               End
               Begin VB.TextBox txtstrNumeroC 
                  Height          =   300
                  Left            =   5475
                  MaxLength       =   10
                  TabIndex        =   18
                  Top             =   330
                  Width           =   825
               End
               Begin VB.TextBox txtintCepC 
                  Height          =   300
                  Left            =   7560
                  MaxLength       =   9
                  TabIndex        =   23
                  Top             =   720
                  Width           =   1005
               End
               Begin VB.TextBox txtstrComplementoC 
                  Height          =   300
                  Left            =   6960
                  MaxLength       =   10
                  TabIndex        =   19
                  Top             =   315
                  Width           =   1590
               End
               Begin VB.TextBox txtstrLogradouroC 
                  Height          =   300
                  Left            =   1080
                  MaxLength       =   100
                  TabIndex        =   17
                  Top             =   300
                  Width           =   4065
               End
               Begin VB.TextBox txtstrBairroC 
                  Height          =   300
                  Left            =   510
                  MaxLength       =   50
                  TabIndex        =   20
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.Label lbl_MunicipioC 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Município"
                  Height          =   195
                  Left            =   3435
                  TabIndex        =   59
                  Top             =   810
                  Width           =   705
               End
               Begin VB.Label lbl_UFC 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "UF"
                  Height          =   195
                  Left            =   6495
                  TabIndex        =   58
                  Top             =   810
                  Width           =   210
               End
               Begin VB.Label lbl_CepC 
                  AutoSize        =   -1  'True
                  Caption         =   "CEP"
                  Height          =   195
                  Left            =   7170
                  TabIndex        =   57
                  Top             =   780
                  Width           =   315
               End
               Begin VB.Label lbl_ComplementoC 
                  AutoSize        =   -1  'True
                  Caption         =   "Compl."
                  Height          =   195
                  Left            =   6435
                  TabIndex        =   56
                  Top             =   390
                  Width           =   480
               End
               Begin VB.Label lbl_NumeroC 
                  AutoSize        =   -1  'True
                  Caption         =   "N°"
                  Height          =   195
                  Left            =   5250
                  TabIndex        =   55
                  Top             =   390
                  Width           =   180
               End
               Begin VB.Label lbl_BairroC 
                  AutoSize        =   -1  'True
                  Caption         =   "Bairro"
                  Height          =   195
                  Left            =   75
                  TabIndex        =   54
                  Top             =   765
                  Width           =   405
               End
               Begin VB.Label lbl_LogradouroC 
                  AutoSize        =   -1  'True
                  Caption         =   "Logradouro"
                  Height          =   195
                  Left            =   195
                  TabIndex        =   53
                  Top             =   390
                  Width           =   810
               End
            End
            Begin VB.Frame fra_EndImobiliario 
               Height          =   1170
               Left            =   150
               TabIndex        =   46
               Top             =   315
               Width           =   8655
               Begin VB.TextBox txtstrBairro 
                  Height          =   300
                  Left            =   1080
                  MaxLength       =   50
                  TabIndex        =   15
                  Top             =   720
                  Width           =   4050
               End
               Begin VB.TextBox txtstrLogradouro 
                  Height          =   300
                  Left            =   1080
                  MaxLength       =   100
                  TabIndex        =   12
                  Top             =   300
                  Width           =   4065
               End
               Begin VB.TextBox txtstrComplemento 
                  Height          =   300
                  Left            =   6960
                  MaxLength       =   10
                  TabIndex        =   14
                  Top             =   315
                  Width           =   1590
               End
               Begin VB.TextBox txtintCep 
                  Height          =   300
                  Left            =   5685
                  MaxLength       =   9
                  TabIndex        =   16
                  Top             =   720
                  Width           =   1005
               End
               Begin VB.TextBox txtstrNumero 
                  Height          =   300
                  Left            =   5475
                  MaxLength       =   10
                  TabIndex        =   13
                  Top             =   330
                  Width           =   825
               End
               Begin VB.Label lblintLogradouro 
                  AutoSize        =   -1  'True
                  Caption         =   "Logradouro"
                  Height          =   195
                  Left            =   195
                  TabIndex        =   51
                  Top             =   390
                  Width           =   810
               End
               Begin VB.Label lblintBairro 
                  AutoSize        =   -1  'True
                  Caption         =   "Bairro"
                  Height          =   195
                  Left            =   585
                  TabIndex        =   50
                  Top             =   765
                  Width           =   405
               End
               Begin VB.Label lblintNumero 
                  AutoSize        =   -1  'True
                  Caption         =   "N°"
                  Height          =   195
                  Left            =   5250
                  TabIndex        =   49
                  Top             =   390
                  Width           =   180
               End
               Begin VB.Label lblstrComplemento 
                  AutoSize        =   -1  'True
                  Caption         =   "Compl."
                  Height          =   195
                  Left            =   6435
                  TabIndex        =   48
                  Top             =   390
                  Width           =   480
               End
               Begin VB.Label lblintCep 
                  AutoSize        =   -1  'True
                  Caption         =   "CEP"
                  Height          =   195
                  Left            =   5295
                  TabIndex        =   47
                  Top             =   780
                  Width           =   315
               End
            End
         End
         Begin VB.Label lbl_Sequencia 
            AutoSize        =   -1  'True
            Caption         =   "Sequência"
            Height          =   195
            Left            =   3750
            TabIndex        =   73
            Top             =   2040
            Width           =   765
         End
         Begin VB.Label lbl_Envolvidos 
            AutoSize        =   -1  'True
            Caption         =   "Envolvidos"
            Height          =   195
            Left            =   225
            TabIndex        =   45
            Top             =   3075
            Width           =   780
         End
         Begin VB.Label lbl_Promissario 
            AutoSize        =   -1  'True
            Caption         =   "Promissário"
            Height          =   195
            Left            =   225
            TabIndex        =   44
            Top             =   2730
            Width           =   795
         End
         Begin VB.Label lbl_Proprietario 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   210
            TabIndex        =   43
            Top             =   2310
            Width           =   795
         End
         Begin VB.Label lblstrQuadra 
            AutoSize        =   -1  'True
            Caption         =   "Quadra"
            Height          =   195
            Left            =   480
            TabIndex        =   42
            Top             =   2010
            Width           =   525
         End
         Begin VB.Label lblstrLote 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            Height          =   195
            Left            =   2535
            TabIndex        =   41
            Top             =   2025
            Width           =   315
         End
         Begin VB.Label lblstrLoteamento 
            AutoSize        =   -1  'True
            Caption         =   "Loteamento"
            Height          =   195
            Left            =   5340
            TabIndex        =   40
            Top             =   2025
            Width           =   840
         End
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1740
      Left            =   30
      TabIndex        =   114
      Top             =   6780
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   3069
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
      Columns(1).Caption=   "Inscricao"
      Columns(1).DataField=   "strInscricao"
      Columns(1).NumberFormat=   "FormatText Event"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Composição"
      Columns(2).DataField=   "Composicao"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Exercício"
      Columns(3).DataField=   "Exercicio"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Número do Aviso"
      Columns(4).DataField=   "NumeroAviso"
      Columns(4).NumberFormat=   "FormatText Event"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Proprietário"
      Columns(5).DataField=   "Proprietario"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Promissário"
      Columns(6).DataField=   "Promissario"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Quadra"
      Columns(7).DataField=   "Quadra"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Lote"
      Columns(8).DataField=   "Lote"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "PkidIPTU"
      Columns(9).DataField=   "PkidIPTU"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2672"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2593"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=5689"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=5609"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1455"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1376"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2328"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2249"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=7382"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=7303"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=8281"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=8202"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=2646"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2566"
      Splits(0)._ColumnProps(45)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=1164"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=1085"
      Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(52)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(55)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(56)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bgcolor=&H80000009&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(18)  =   ":id=6,.fgcolor=&H8000000E&"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(21)  =   ":id=8,.fgcolor=&H8000000E&"
      _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.namedParent=38"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(32)  =   ":id=18,.fgcolor=&H8000000E&"
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(35)  =   ":id=19,.fgcolor=&H8000000E&"
      _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=74,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(80)  =   "Named:id=33:Normal"
      _StyleDefs(81)  =   ":id=33,.parent=0"
      _StyleDefs(82)  =   "Named:id=34:Heading"
      _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   ":id=34,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=36:Selected"
      _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=37:Caption"
      _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(91)  =   "Named:id=38:HighlightRow"
      _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=39:EvenRow"
      _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(95)  =   "Named:id=40:OddRow"
      _StyleDefs(96)  =   ":id=40,.parent=33"
      _StyleDefs(97)  =   "Named:id=41:RecordSelector"
      _StyleDefs(98)  =   ":id=41,.parent=34"
      _StyleDefs(99)  =   "Named:id=42:FilterBar"
      _StyleDefs(100) =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadLancamentoIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnPrimeiraVez      As Boolean

Private Sub dbcintExercicio_Click(Area As Integer)
    DropDownDataCombo dbcintExercicio, Me, Area
End Sub
Private Sub dbcintExercicio_GotFocus()
    MarcaCampo dbcintExercicio
End Sub
Private Sub dbcintExercicio_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintExercicio, Me, , KeyCode, Shift
End Sub
Private Sub dbcintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintExercicio
End Sub
Private Sub Form_Activate()
    gintCodSeguranca = 1076
    VerificaMascaraInscricao
End Sub

Private Sub Form_Load()
    With tdb_PrediosCaracteristicas.Columns("Pontos").ValueItems
        .Presentation = dbgRadioButton
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    blnPrimeiraVez = False
End Sub

Private Sub lbl_strInscricaoAnterior_DblClick(Index As Integer)
    frm_Arquiv_Banco.Show
End Sub

Private Sub mskstrInscricao_GotFocus()
    MarcaCampo mskstrInscricao
End Sub

Sub VerificaMascaraInscricao()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
    
    strMascara = ""
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & TYP_IMOBILIARIA
    strSql = strSql & "Order By intSequencia"
    
    Set gobjBanco = New clsBanco
        
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    
    mskstrInscricao.Mask = strMascara
    mskstrInscricao2.Mask = strMascara
    mskstrInscricao3.Mask = strMascara
    mskstrInscricao4.Mask = strMascara
End Sub
Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao
End Sub
Private Sub mskstrInscricao_LostFocus()
    If txtPkid = "" Then
        LeDaTabelaParaObj "", dbcintExercicio, strQueryExercicio
    End If
End Sub

Private Sub mskstrInscricao2_GotFocus()
    MarcaCampo mskstrInscricao2
End Sub

Private Sub mskstrInscricao2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao2
End Sub

Private Sub mskstrInscricao3_GotFocus()
    MarcaCampo mskstrInscricao3
End Sub

Private Sub mskstrInscricao3_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao3
End Sub

Private Sub mskstrInscricao4_GotFocus()
    MarcaCampo mskstrInscricao4
End Sub

Private Sub mskstrInscricao4_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskstrInscricao4
End Sub

Private Sub tdb_Lista_Click()
    blnPrimeiraVez = True
End Sub

Private Sub tdb_Lista_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
Dim strValor As String

    If ColIndex = 1 Then
        If Len(Value) > 0 Then
           strValor = Value
           Value = gstrFormataInscricao(strValor, TYP_IMOBILIARIA)
        End If
    End If

End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If blnPrimeiraVez Then
        If Not tdb_Lista.EOF And Len(tdb_Lista.Columns("Pkid").Value) > 0 Then
            txtPkid.Text = tdb_Lista.Columns("Pkid").Value
            txtPkidIPTU.Text = tdb_Lista.Columns("PkidIPTU").Value
            PreencheCabecalho
            PreencheTabLancamentoIptu
            PreencheTabValorVenalETributos (Val(txtPkid.Text))
            PreencheTabCaracteristicas (Val(txtPkidIPTU.Text))
            PreecheTotaisValorVenalEImpostos
        End If
    End If
End Sub
Private Sub PreecheTotaisValorVenalEImpostos()
txt_dblTotaisValorVenal = Val(gstrConvVrParaSql(txtdblValorVenalTerreno)) + Val(gstrConvVrParaSql(txtdblValorTerrenoExcedente)) + Val(gstrConvVrParaSql(txtdblTotalValorVenal))
txt_dblTotaisImposto = Val(gstrConvVrParaSql(txtdblImpostoTerreno)) + Val(gstrConvVrParaSql(txtdblImpostoExcedente)) + Val(gstrConvVrParaSql(txtdblTotalImposto))
txt_dblTotaisValorVenal = gstrConvVrDoSql(txt_dblTotaisValorVenal)
txt_dblTotaisImposto = gstrConvVrDoSql(txt_dblTotaisImposto)
End Sub


Private Sub tdb_Parcelas_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
Dim strValor As String

    If ColIndex = 1 Then
        If Len(Value) > 0 Then
           strValor = Value
           Value = gstrFormataInscricao(strValor, TYP_ACORDO)
        End If
    ElseIf ColIndex = 6 Or ColIndex = 4 Then
        Value = gstrDataFormatada(Value)
    End If
    
End Sub

Private Sub tdb_PrediosCaracteristicas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tdb_PrediosCaracteristicas.EOF Then
        PreencheCaracteristicas txtPkidIPTU, tdb_PrediosCaracteristicas.Columns("Pkid").Value
    End If
End Sub

Private Sub txtdtmCancelamento_GotFocus()
    MarcaCampo txtDtmCancelamento
End Sub

Private Sub txtdtmCancelamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDtmCancelamento
End Sub

Private Sub txtdtmCancelamento_LostFocus()
    txtDtmCancelamento.Text = gstrDataFormatada(txtDtmCancelamento.Text)
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
End Sub
Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub
Private Sub txtintCepC_GotFocus()
    MarcaCampo txtintCepC
End Sub
Private Sub txtintCepC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepC
End Sub
Private Sub txtintLote_GotFocus()
    MarcaCampo txtintLote
End Sub
Private Sub txtintLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintLote
End Sub
Private Sub txtstrNumero_GotFocus()
    MarcaCampo txtstrNumero
End Sub
Private Sub txtstrNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumero
End Sub
Private Sub txtstrNumeroC_Change()
    MarcaCampo txtstrNumeroC
End Sub
Private Sub txtstrNumeroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumeroC
End Sub
Private Sub txtstrBairro_GotFocus()
    MarcaCampo txtstrBairro
End Sub
Private Sub txtstrBairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrBairro
End Sub
Private Sub txtstrBairroC_GotFocus()
    MarcaCampo txtstrBairroC
End Sub
Private Sub txtstrBairroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrBairroC
End Sub
Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
End Sub
Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub
Private Sub txtstrComplementoC_GotFocus()
    MarcaCampo txtstrComplementoC
End Sub
Private Sub txtstrComplementoC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplementoC
End Sub
Private Sub txtstrComposicaoDaReceita_GotFocus()
    MarcaCampo txtstrComposicaoDaReceita
End Sub
Private Sub txtstrComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComposicaoDaReceita
End Sub
Private Sub txtstrEmissao_GotFocus()
    MarcaCampo txtstrEmissao
End Sub
Private Sub txtstrEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrEmissao
End Sub
Private Sub txtstrLogradouro_GotFocus()
    MarcaCampo txtstrLogradouro
End Sub
Private Sub txtstrLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLogradouro
End Sub
Private Sub txtstrLogradouroC_GotFocus()
    MarcaCampo txtstrLogradouroC
End Sub
Private Sub txtstrLogradouroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLogradouroC
End Sub
Private Sub txtstrLoteamento_GotFocus()
    MarcaCampo txtstrLoteamento
End Sub
Private Sub txtstrLoteamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLoteamento
End Sub
Private Sub txtstrNomeProprietario_GotFocus()
    MarcaCampo txtstrNomeProprietario
End Sub
Private Sub txtstrNomeProprietario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNomeProprietario
End Sub

Private Sub txtstrNumeroAviso_GotFocus()
    MarcaCampo txtstrNumeroAviso
End Sub
Private Sub txtstrNumeroAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNumeroAviso
End Sub

Private Sub txtstrPromissario_GotFocus()
    MarcaCampo txtstrPromissario
End Sub
Private Sub txtstrPromissario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, txtstrPromissario
End Sub
Private Sub txtstrQuadra_GotFocus()
    MarcaCampo txtstrQuadra
End Sub
Private Sub txtstrQuadra_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrQuadra
End Sub
Private Sub txtstrUFC_GotFocus()
    MarcaCampo txtstrUFC
End Sub
Private Sub txtstrUFC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrUFC
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrLocalizar)
            blnPrimeiraVez = True
            LeDaTabelaParaObj "", tdb_Lista, strQueryLocalizar(False)
        Case Is = UCase(gstrRefresh)
            LeDaTabelaParaObj "", tdb_Lista, strQueryLocalizar(True)
        Case Is = UCase(gstrNovo)
            LimpaObjeto Me
            LimpaGrids
            tab_3DPasta.Tab = 0
        Case Is = UCase(gstrImprimir)
            ImprimeRelatorio rptLancamentoIPTU, strQueryRelatorio
    End Select
End Sub

Private Function strQueryRelatorio() As String
    Dim strSql As String
    
    strSql = "SELECT "
        'strSQL = strSQL & " LA.strInscricao Inscricao,"
        strSql = strSql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
        strSql = strSql & gstrCONVERT(CDT_numeric, "strNumeroAviso") & " NumeroAviso,"
        strSql = strSql & " LA.strNumeroAviso NumeroAviso,"
        strSql = strSql & " LA.strNomeProprietario Proprietario,"
        strSql = strSql & " LA.strPromissario Promissario,"
        strSql = strSql & " LI.strQuadra Quadra,"
        strSql = strSql & " LI.strLote Lote"
    
    strSql = strSql & " FROM "
        strSql = strSql & gstrLancamentoAlfa & " LA, "
        strSql = strSql & gstrLancamentoIPTU & " LI"
    
    strSql = strSql & " WHERE LA.Pkid = LI.intLancamentoAlfa"
        strSql = strSql & " AND UPPER(strInscricao) LIKE " & "'" & UCase(mskstrInscricao.Text) & "%'"
        strSql = strSql & " AND UPPER(intExercicio) LIKE " & "'" & UCase(dbcintExercicio.Text) & "%'"
        strSql = strSql & " AND UPPER(strEmissao) LIKE " & "'" & UCase(txtstrEmissao.Text) & "%'"
        strSql = strSql & " AND UPPER(strNumeroAviso) LIKE " & "'" & UCase(txtstrNumeroAviso.Text) & "%'"
        strSql = strSql & " AND UPPER(strComposicaoDaReceita) LIKE " & "'" & UCase(txtstrComposicaoDaReceita.Text) & "%'"
    strSql = strSql & " ORDER BY LA.strInscricao"
    
    strQueryRelatorio = strSql
End Function

Private Function strQueryLocalizar(blnRefresh As Boolean) As String

    Dim strSql As String
    
    strSql = "SELECT LA.Pkid Pkid,"
    strSql = strSql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
    strSql = strSql & " LA.strComposicaoDaReceita Composicao,"
    strSql = strSql & " LA.intExercicio Exercicio,"
    strSql = strSql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " NumeroAviso,"
    strSql = strSql & " LA.strNomeProprietario Proprietario,"
    strSql = strSql & " LA.strPromissario Promissario,"
    strSql = strSql & " LI.strQuadra Quadra,"
    strSql = strSql & " LI.strLote Lote,"
    strSql = strSql & " LI.PKid PkidIPTU"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoIPTU & " LI"
    strSql = strSql & " WHERE LA.Pkid = LI.intLancamentoAlfa"
    
    If Not blnRefresh Then
        If Len(mskstrInscricao.Text) > 0 Then strSql = strSql & " AND (strInscricao LIKE " & "'" & UCase(String(gintLenInscricao - gintRetornaTamanhoMascara(TYP_IMOBILIARIA), "0") & mskstrInscricao) & "%' OR strInscricao LIKE " & "'" & UCase(String(gintLenInscricao - Len(mskstrInscricao.Text), "0") & mskstrInscricao) & "%')"
        If Len(dbcintExercicio.Text) > 0 Then strSql = strSql & " AND intExercicio LIKE " & "'" & UCase(dbcintExercicio.Text) & "%'"
        If Len(txtstrEmissao.Text) > 0 Then strSql = strSql & " AND strEmissao LIKE " & "'" & UCase(String(gintLenEmissao - Len(txtstrEmissao), "0") & txtstrEmissao) & "'"
        If Len(txtstrNumeroAviso.Text) > 0 Then strSql = strSql & " AND strNumeroAviso LIKE " & "'" & UCase(String(gintLenNumAviso - Len(txtstrNumeroAviso), "0") & txtstrNumeroAviso.Text) & "%'"
        If Len(txtstrComposicaoDaReceita.Text) > 0 Then strSql = strSql & " AND UPPER(strComposicaoDaReceita) LIKE " & "'" & UCase(txtstrComposicaoDaReceita.Text) & "%'"
        If Len(txtstrQuadra.Text) > 0 Then strSql = strSql & " AND LI.strQuadra = " & "'" & txtstrQuadra.Text & "'"
        If Len(txtintLote.Text) > 0 Then strSql = strSql & " AND LI.strLote = " & "'" & txtintLote.Text & "'"
        If Len(txtstrSequenciaDeFace.Text) > 0 Then strSql = strSql & " AND LI.strSequenciaDeFace = " & txtstrSequenciaDeFace.Text
        If Len(txtstrLoteamento.Text) > 0 Then strSql = strSql & " AND UPPER(LI.strLoteamento) LIKE " & "'" & UCase(txtstrLoteamento.Text) & "%'"
        If Len(txtstrNomeProprietario.Text) > 0 Then strSql = strSql & " AND UPPER(LA.strNomeProprietario) LIKE " & "'" & UCase(txtstrNomeProprietario.Text) & "%'"
        If Len(txtstrPromissario.Text) > 0 Then strSql = strSql & " AND UPPER(LA.strPromissario) LIKE " & "'" & UCase(txtstrPromissario.Text) & "%'"
        If Len(txtstrLogradouro.Text) > 0 Then strSql = strSql & " AND UPPER(LA.strLogradouro) LIKE " & "'" & UCase(txtstrLogradouro.Text) & "%'"
        If Len(txtstrBairro.Text) > 0 Then strSql = strSql & " AND UPPER(LA.strBairro) LIKE " & "'" & UCase(txtstrBairro.Text) & "%'"
        If Len(txtintCep.Text) > 0 Then strSql = strSql & " AND LA.intCEP = " & Replace(txtintCep.Text, "-", "")
    End If
    
    strSql = strSql & " ORDER BY LTRIM(RTRIM(LA.strInscricao)), LTRIM(RTRIM(LA.strComposicaoDaReceita)), Exercicio DESC, NumeroAviso DESC "

    strQueryLocalizar = strSql

End Function

Private Function strQueryExercicio() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid, intExercicio"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa
    strSql = strSql & " WHERE"
    strSql = strSql & " strInscricao = '" & mskstrInscricao.Text & "'"
    strSql = strSql & " ORDER BY intExercicio"
    
    strQueryExercicio = strSql
    
End Function

Private Sub PreencheCabecalho()
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "SELECT " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " Inscricao, "
    strSql = strSql & " intExercicio Exercicio,"
    strSql = strSql & " strEmissao Emissao,"
    strSql = strSql & gstrCONVERT(CDT_numeric, "strNumeroAviso") & " NumeroAviso,"
    strSql = strSql & " strComposicaoDaReceita ComposicaoDaReceita,"
    strSql = strSql & " dtmDtCancelamento Cancelamento"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa
    strSql = strSql & " WHERE"
    strSql = strSql & " Pkid = '" & txtPkid.Text & "'"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            mskstrInscricao.Text = Space$(0) & adoResultado!Inscricao
            mskstrInscricao2.Text = Space$(0) & adoResultado!Inscricao
            mskstrInscricao3.Text = Space$(0) & adoResultado!Inscricao
            mskstrInscricao4.Text = Space$(0) & adoResultado!Inscricao
            dbcintExercicio.Text = Space$(0) & adoResultado!EXERCICIO
            dbcintExercicio2.Text = Space$(0) & adoResultado!EXERCICIO
            dbcintExercicio3.Text = Space$(0) & adoResultado!EXERCICIO
            dbcintExercicio4.Text = Space$(0) & adoResultado!EXERCICIO
            txtstrEmissao.Text = Space$(0) & adoResultado!Emissao
            txtstrEmissao2.Text = Space$(0) & adoResultado!Emissao
            txtstrEmissao3.Text = Space$(0) & adoResultado!Emissao
            txtstrEmissao4.Text = Space$(0) & adoResultado!Emissao
            txtstrNumeroAviso.Text = Space$(0) & adoResultado!NumeroAviso
            txtstrNumDoAviso2.Text = Space$(0) & adoResultado!NumeroAviso
            txtstrNumDoAviso3.Text = Space$(0) & adoResultado!NumeroAviso
            txtstrNumDoAviso4.Text = Space$(0) & adoResultado!NumeroAviso
            txtstrComposicaoDaReceita.Text = Space$(0) & adoResultado!ComposicaoDaReceita
            txtDtmCancelamento.Text = Space$(0) & gstrDataFormatada(adoResultado!cancelamento)
        End If
    End If
End Sub

Private Sub PreencheTabLancamentoIptu()
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "SELECT LI.Pkid PkidIPTU,"
    strSql = strSql & " LI.strQuadra Quadra,"
    strSql = strSql & " LI.strLote Lote,"
    strSql = strSql & " LI.strSequenciaDeFace SequenciaDeFace,"
    strSql = strSql & " LI.strLoteamento Loteamento,"
    strSql = strSql & " LA.strdefinicaoisencao, LA.strtipoisencaoimunidade,"
    strSql = strSql & " LA.strNomeProprietario Proprietario,"
    strSql = strSql & " LA.strPromissario Promissario,"
    strSql = strSql & " LA.strLogradouro Logradouro,"
    strSql = strSql & " LA.strNumero Numero,"
    strSql = strSql & " LA.strComplemento Complemento,"
    strSql = strSql & " LA.strBairro Bairro,"
    strSql = strSql & " LA.intCEP CEP,"
    strSql = strSql & " LA.strLogradouroC LogradouroN,"
    strSql = strSql & " LA.strNumeroC NumeroN,"
    strSql = strSql & " LA.strComplementoC ComplementoN,"
    strSql = strSql & " LA.strBairroC BairroN,"
    strSql = strSql & " LA.strMunicipioC MunicipioN,"
    strSql = strSql & " LA.strUFC UFN,"
    strSql = strSql & " LA.intCEPC CEPN"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoIPTU & " LI, "
    strSql = strSql & gstrLancamentoAlfa & " LA"
    strSql = strSql & " WHERE "
    strSql = strSql & " LA.Pkid = '" & txtPkid.Text & "'"
    strSql = strSql & " AND LA.Pkid " & strOUTJSQLServer & "=" & " LI.intLancamentoAlfa " & strOUTJOracle
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 30, adoResultado) Then
        If Not adoResultado.EOF Then
            PreencheEnvolvidos (adoResultado!PkidIPTU)
            txtstrdefinicaoisencao = gstrENulo(adoResultado!strdefinicaoisencao)
            txtstrtipoisencaoimunidade = gstrENulo(adoResultado!strtipoisencaoimunidade)
            txtstrQuadra.Text = gstrENulo(adoResultado!Quadra)
            txtintLote.Text = gstrENulo(adoResultado!Lote)
            txtstrSequenciaDeFace.Text = gstrENulo(adoResultado!SequenciaDeFace)
            txtstrLoteamento.Text = gstrENulo(adoResultado!Loteamento)
            txtstrNomeProprietario.Text = gstrENulo(adoResultado!Proprietario)
            txtstrPromissario.Text = gstrENulo(adoResultado!Promissario)
            txtstrLogradouro.Text = gstrENulo(adoResultado!Logradouro)
            txtstrNumero.Text = gstrENulo(adoResultado!Numero)
            txtstrComplemento.Text = gstrENulo(adoResultado!Complemento)
            txtstrBairro.Text = gstrENulo(adoResultado!Bairro)
            txtintCep.Text = gstrCEPFormatado(gstrENulo(adoResultado!CEP))
            txtstrLogradouroC.Text = gstrENulo(adoResultado!LogradouroN)
            txtstrNumeroC.Text = gstrENulo(adoResultado!NumeroN)
            txtstrComplementoC.Text = gstrENulo(adoResultado!ComplementoN)
            txtstrBairroC.Text = gstrENulo(adoResultado!BairroN)
            txtstrMunicipioC.Text = gstrENulo(adoResultado!MunicipioN)
            txtstrUFC.Text = gstrENulo(adoResultado!UFN)
            txtintCepC.Text = gstrCEPFormatado(gstrENulo(adoResultado!CEPN))
        End If
    End If
    
End Sub

Private Sub PreencheEnvolvidos(lngPkid As Long)
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT bitProprietario,"
    strSql = strSql & " strNomeEnvolvido Envolvido"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoEnvolvidos
    strSql = strSql & " WHERE"
    strSql = strSql & " intLancamentoIPTU ='" & lngPkid & "'"
    
    lvwEnvolvidos.ListItems.Clear
        
    Set gobjBanco = New clsBanco
       
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                lvwEnvolvidos.ListItems.Add 1, , adoResultado!Envolvido
                lvwEnvolvidos.ListItems.Item(1).ListSubItems.Add , , IIf(adoResultado!BITPROPRIETARIO = 1, "Proprietário", "Promissário")
                .MoveNext
            Loop
        End With
    End If
    
    If lvwEnvolvidos.ListItems.Count <> 0 Then
        lvwEnvolvidos.SelectedItem.Selected = False
    End If

End Sub

Private Sub PreencheTabValorVenalETributos(lngPkidAlfaLancamento As Long)
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT dblAreaTerreno,"
    strSql = strSql & " dblValorMetro,"
    strSql = strSql & " dblValorVenalTerreno,"
    strSql = strSql & " dblAliquotaTerreno,"
    strSql = strSql & " dblImpostoTerreno,"
    strSql = strSql & " dblAreaExcedente,"
    strSql = strSql & " dblValorTerrenoExcedente,"
    strSql = strSql & " dblAliquotaExcedente,"
    strSql = strSql & " dblImpostoExcedente,"
    strSql = strSql & " LA.dblValorCompensacao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoIPTU & " LI, "
    strSql = strSql & gstrLancamentoAlfa & " LA "
    strSql = strSql & " WHERE "
    strSql = strSql & " LI.intLancamentoAlfa = '" & Val(lngPkidAlfaLancamento) & "' AND "
    strSql = strSql & " LA.Pkid = LI.intLancamentoAlfa"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txtdblAreaTerreno.Text = gstrConvVrDoSql(adoResultado!Dblareaterreno, 2)
            txtdblValorMetro.Text = gstrConvVrDoSql(adoResultado!dblValorMetro, 2)
            txtdblValorVenalTerreno.Text = gstrConvVrDoSql(adoResultado!dblValorVenalTerreno, 2)
            txtdblAliquotaTerreno.Text = gstrConvVrDoSql(adoResultado!dblAliquotaTerreno, 2)
            txtdblImpostoTerreno.Text = gstrConvVrDoSql(adoResultado!dblimpostoterreno, 2)
            txtdblAreaExcedente.Text = gstrConvVrDoSql(adoResultado!dblAreaExcedente, 2)
            txtdblValorTerrenoExcedente.Text = gstrConvVrDoSql(adoResultado!dblValorTerrenoExcedente, 2)
            txtdblAliquotaExcedente.Text = gstrConvVrDoSql(adoResultado!dblAliquotaExcedente, 2)
            txtdblImpostoExcedente.Text = gstrConvVrDoSql(adoResultado!dblImpostoExcedente, 2)
            txtdblCompensacao.Text = IIf(IsNull(adoResultado!dblValorCompensacao), "0,00", gstrConvVrDoSql(adoResultado!dblValorCompensacao, 2))
        End If
    End If
    
    CarregaFatores txtPkidIPTU
    CarregaPredios txtPkidIPTU
    CarregaTributos
    CarregaValorTotal
    CarregaParcelas txtPkid
    
End Sub

Private Sub CarregaFatores(lngPkid As Long)
    Dim strSql As String
    
    strSql = "SELECT strDescricao,"
    strSql = strSql & " dblFator"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoFatores
    strSql = strSql & " WHERE"
    strSql = strSql & " intLancamentoIPTU = '" & Val(lngPkid) & "'"
    
    LeDaTabelaParaObj "", tdb_Fatores, strSql
    
End Sub

Private Sub CarregaPredios(lngPkid As Long)
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT strNomeDaArea,"
    strSql = strSql & " dblMedidaDaArea,"
    strSql = strSql & " dblValorMetro,"
    strSql = strSql & " strNomeFator,"
    strSql = strSql & " dblFatorObsolescencia,"
    strSql = strSql & " dblValorVenalPredio,"
    strSql = strSql & " dblAliquota,"
    strSql = strSql & " dblImposto"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoPredioIPTU
    strSql = strSql & " WHERE"
    strSql = strSql & " intLancamentoIPTU ='" & Val(lngPkid) & "'"
    
    LeDaTabelaParaObj "", tdb_Predios, strSql
                      
    strSql = "SELECT SUM(dblMedidaDaArea) dblTotalArea,"
    strSql = strSql & " SUM(dblValorVenalPredio) dblTotalValorVenal,"
    strSql = strSql & " SUM(dblImposto) dblTotalImposto"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoPredioIPTU
    strSql = strSql & " WHERE"
    strSql = strSql & " intLancamentoIPTU ='" & Val(lngPkid) & "'"
                      
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txtdblTotalArea.Text = gstrConvVrDoSql(adoResultado("dblTotalArea"), 2)
            txtdblTotalValorVenal.Text = gstrConvVrDoSql(adoResultado("dblTotalValorVenal"), 2)
            txtdblTotalImposto.Text = gstrConvVrDoSql(adoResultado("dblTotalImposto"), 2)
        End If
    End If
    
End Sub

Private Sub CarregaTributos()
    Dim x As XArrayDB
    Dim adoResultado As ADODB.Recordset
    Dim dblTotalReceita As Currency
    
    Set x = New XArrayDB
    
    x.ReDim 0, 2, 0, 1
    
    x.Value(0, 0) = "Imposto Territorial Urbano"
    x.Value(1, 0) = "Imposto Excedente"
    x.Value(2, 0) = "Imposto Predial Urbano"
    
    
    x.Value(0, 1) = txtdblImpostoTerreno
    x.Value(1, 1) = txtdblImpostoExcedente
    x.Value(2, 1) = dblImpostoPredial
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strQueryReceita, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            x.ReDim 0, adoResultado.RecordCount + 2, 0, 1
            Do While Not adoResultado.EOF
                x.Value(2 + adoResultado.AbsolutePosition, 0) = gstrENulo(adoResultado("strReceita").Value)
                x.Value(2 + adoResultado.AbsolutePosition, 1) = gstrConvVrDoSql(gstrENulo(adoResultado("dblValorReceita").Value))
                dblTotalReceita = dblTotalReceita + gstrConvVrDoSql(gstrENulo(adoResultado("dblValorReceita").Value))
                adoResultado.MoveNext
            Loop
        End If
    End If
    
    Set tdb_Tributos.Array = x
    tdb_Tributos.ReBind
    tdb_Tributos.Refresh
    
End Sub

Private Function dblImpostoPredial() As Double
    tdb_Predios.Update
    tdb_Predios.MoveFirst
    
    Do While Not tdb_Predios.EOF
        dblImpostoPredial = dblImpostoPredial + gstrConvVrDoSql(tdb_Predios.Columns("dblImposto").Value, 2)
        tdb_Predios.MoveNext
    Loop
    tdb_Predios.MoveFirst
End Function

Private Sub CarregaParcelas(lngPkid As Long)
    Dim strSql As String
    
    strSql = "SELECT LV.intParcela, "
    strSql = strSql & "LV.dblValor, "
    strSql = strSql & "CASE WHEN LV.intLancamentoAlfaDAtiva IS NULL THEN '' ELSE 'X' END intLancamentoAlfaDAtiva ,"
    'strSql = strSql & strSUBSTRING & "(LA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & " " & gstrRIGHT("LA.strInscricao", 4) & " strAcordo,"
    strSql = strSql & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ACORDO)) & " strAcordo, "
    strSql = strSql & "M.Strabreviatura as strMoeda, "
    strSql = strSql & "LV.dtmDtVencimento, "
    strSql = strSql & "LP.dtmDtPagamento, "
    If bytDBType = Oracle Then
        strSql = strSql & "CB.STRDESCRICAO, "
    Else
        strSql = strSql & "(Select STRDESCRICAO From tblCodigoBaixa where pkid = LP.Intcodigobaixa) strDescricao, "
    End If
    strSql = strSql & "LP.Strobservacao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoValor & " LV, "
    
    If bytDBType = Oracle Then strSql = strSql & gstrCodigoDeBaixa & " CB, "
    
    strSql = strSql & gstrMoedas & " M, "
    strSql = strSql & gstrLancamentoPagamento & " LP, "
    strSql = strSql & gstrLancamentoAlfa & " LA "
    strSql = strSql & "WHERE "
    strSql = strSql & "LV.Pkid " & strOUTJSQLServer & "=" & " LP.intLancamentoValor " & strOUTJOracle & " AND "
    
    If bytDBType = Oracle Then strSql = strSql & "CB.Pkid " & strOUTJOracle & "= LP.Intcodigobaixa AND "
    
    strSql = strSql & "M.Pkid = LV.Intmoeda AND "
    strSql = strSql & "LV.intLancamentoAlfa = " & Val(lngPkid) & " AND "
    strSql = strSql & "LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= LA.pkID " & strOUTJOracle & " "
    
    strSql = strSql & " ORDER BY LV.intParcela"
    
    LeDaTabelaParaObj "", tdb_Parcelas, strSql
    
End Sub

Private Sub CarregaValorTotal()
    tdb_Tributos.Update
    tdb_Tributos.MoveFirst
    txtTotalTributo.Text = ""
    Do While Not tdb_Tributos.EOF
        txtTotalTributo = CDbl(gstrConvVrDoSql(gstrENulo(txtTotalTributo.Text), 2, , True)) + CDbl(gstrConvVrDoSql(gstrENulo(tdb_Tributos.Columns("Valor").Value), 2, , True))
        tdb_Tributos.MoveNext
    Loop
    txtTotalTributo = CDbl(gstrConvVrDoSql(gstrENulo(txtTotalTributo.Text), 2, , True)) - CDbl(gstrConvVrDoSql(gstrENulo(txtdblCompensacao.Text), 2, , True))
    txtTotalTributo = gstrConvVrDoSql(txtTotalTributo, 2, , True)
End Sub

Private Sub PreencheTabCaracteristicas(lngPkid As Long)
    Dim strSql As String
    
    strSql = "SELECT Pkid,"
    strSql = strSql & " strNomeDaArea Area,"
    strSql = strSql & " dblMedidaDaArea,"
    strSql = strSql & " dblPontos,"
    strSql = strSql & " dblValorMetro,"
    strSql = strSql & " dblFatorObsolescencia,"
    strSql = strSql & " dblValorVenalPredio"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoPredioIPTU
    strSql = strSql & " WHERE"
    strSql = strSql & " intLancamentoIPTU = '" & Val(lngPkid) & "'"
    strSql = strSql & " ORDER BY intNEdificacao"
    
    LeDaTabelaParaObj "", tdb_PrediosCaracteristicas, strSql
    
End Sub

Private Sub PreencheCaracteristicas(lngPkidIPTU As Long, lngPkidPredio As Long)
    Dim strSql As String
    
    strSql = "SELECT strNomeCaracGeral,"
    strSql = strSql & " strNomeDoDetalhe,"
    strSql = strSql & " dblValorDetalhe"
    strSql = strSql & " FROM "
    strSql = strSql & gstrCaracBoletimIPTU
    strSql = strSql & " WHERE "
    strSql = strSql & " intLancamentoIPTU = '" & Val(lngPkidIPTU) & "'"
    strSql = strSql & " AND intLancamentoPredioIPTU = '" & Val(lngPkidPredio) & "'"
    strSql = strSql & " AND intCodigoUtilizacaoDaTabelaDeV = 3" '3 - Característica do prédio
    strSql = strSql & " ORDER BY strNomeDoDetalhe"
    
    LeDaTabelaParaObj "", tdb_CaracBoletim, strSql
    
End Sub

Private Sub LimpaGrids()
    Dim x As XArrayDB
    
    Set tdb_Predios.DataSource = Nothing
    Set tdb_Parcelas.DataSource = Nothing
    Set tdb_PrediosCaracteristicas.DataSource = Nothing
    Set tdb_CaracBoletim.DataSource = Nothing
    Set tdb_Fatores.DataSource = Nothing
    
    Set x = New XArrayDB
    
    x.Clear
    x.ReDim 0, 0, 0, 2
    
    Set tdb_Tributos.Array = x
    tdb_Tributos.ReBind
    tdb_Tributos.Refresh
    
    lvwEnvolvidos.ListItems.Clear
End Sub

Private Function strQueryReceita() As String
    Dim strSql As String
    
    strSql = strSql & "Select "
    strSql = strSql & "R.strSigla strReceita, "
    strSql = strSql & "Sum(" & gstrISNULL("LR.DblValor", "0") & ") as dblValorReceita "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoIPTU & " LI, "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrLancamentoReceita & " LR, "
    strSql = strSql & gstrReceita & " R "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LI.Intlancamentoalfa And "
    strSql = strSql & "LA.Pkid = LV.Intlancamentoalfa And "
    strSql = strSql & "LV.Pkid = LR.Intlancamentovalor And "
    strSql = strSql & "R.Pkid = LR.Intreceita And "
    strSql = strSql & "LV.bitParcelaValida = 1 And "
    strSql = strSql & "R.byttipo = 3 AND "
    strSql = strSql & "LA.Pkid = " & txtPkid.Text & " "
    strSql = strSql & "Group By "
    strSql = strSql & "R.pkid, "
    strSql = strSql & "R.strSigla, "
    strSql = strSql & "LR.dblvalor "
    strSql = strSql & "Order by "
    strSql = strSql & "strReceita "
    
    strQueryReceita = strSql
    
End Function
