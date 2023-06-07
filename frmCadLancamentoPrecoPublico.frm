VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadLancamentoPrecoPublico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preço Público - Guias"
   ClientHeight    =   7965
   ClientLeft      =   1320
   ClientTop       =   2415
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8640
   Begin VB.CommandButton cmd_intComposicao 
      Height          =   300
      Left            =   8130
      Picture         =   "frmCadLancamentoPrecoPublico.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "590"
      ToolTipText     =   "Ativa Cadastro de Composições"
      Top             =   487
      Width           =   360
   End
   Begin VB.CommandButton cmd_intIdentificacao 
      Height          =   315
      Left            =   8190
      Picture         =   "frmCadLancamentoPrecoPublico.frx":011E
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Ativa Cadastro Único"
      Top             =   2190
      Width           =   360
   End
   Begin VB.TextBox txtbitDigito 
      Alignment       =   1  'Right Justify
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
      Left            =   2310
      MaxLength       =   2
      TabIndex        =   35
      Top             =   4335
      Width           =   285
   End
   Begin VB.TextBox txtintExercicio 
      Alignment       =   1  'Right Justify
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
      Left            =   1830
      MaxLength       =   4
      TabIndex        =   34
      Top             =   4335
      Width           =   465
   End
   Begin VB.TextBox txtstrCodigo 
      Alignment       =   1  'Right Justify
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
      Left            =   990
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   4335
      Width           =   825
   End
   Begin VB.Frame fra_Endereco 
      Caption         =   " Endereço "
      Height          =   1320
      Left            =   90
      TabIndex        =   17
      Top             =   2940
      Width           =   8475
      Begin VB.TextBox txtstrUF 
         Height          =   285
         Left            =   6105
         MaxLength       =   6
         TabIndex        =   29
         Top             =   915
         Width           =   345
      End
      Begin VB.TextBox txtstrMunicipio 
         Height          =   285
         Left            =   1050
         MaxLength       =   20
         TabIndex        =   27
         Top             =   915
         Width           =   4065
      End
      Begin VB.TextBox txtstrBairro 
         Height          =   285
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   25
         Top             =   555
         Width           =   4065
      End
      Begin VB.TextBox txtstrLogradouro 
         Height          =   285
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   19
         Top             =   195
         Width           =   4065
      End
      Begin VB.TextBox txtintCep 
         Height          =   285
         Left            =   7215
         MaxLength       =   9
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   915
         Width           =   1155
      End
      Begin VB.TextBox txtstrComplemento 
         Height          =   285
         Left            =   7575
         MaxLength       =   20
         TabIndex        =   23
         Top             =   195
         Width           =   795
      End
      Begin VB.TextBox txtintNumero 
         Height          =   285
         Left            =   5580
         MaxLength       =   6
         TabIndex        =   21
         Top             =   195
         Width           =   885
      End
      Begin VB.Label lblstrLogradouro 
         AutoSize        =   -1  'True
         Caption         =   "Logradouro"
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblstrBairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   600
         TabIndex        =   24
         Top             =   615
         Width           =   405
      End
      Begin VB.Label lblstrMunicipio 
         AutoSize        =   -1  'True
         Caption         =   "Município"
         Height          =   195
         Left            =   300
         TabIndex        =   26
         Top             =   960
         Width           =   705
      End
      Begin VB.Label lblintCep 
         AutoSize        =   -1  'True
         Caption         =   "Cep"
         Height          =   195
         Left            =   6855
         TabIndex        =   30
         Top             =   960
         Width           =   285
      End
      Begin VB.Label lblstrUF 
         AutoSize        =   -1  'True
         Caption         =   "UF"
         Height          =   195
         Left            =   5835
         TabIndex        =   28
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblintNumero 
         AutoSize        =   -1  'True
         Caption         =   "Nº"
         Height          =   195
         Left            =   5310
         TabIndex        =   20
         Top             =   270
         Width           =   180
      End
      Begin VB.Label lblstrComplemento 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Compl."
         Height          =   195
         Left            =   7035
         TabIndex        =   22
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.CommandButton cmd_intContribuinte 
      Height          =   315
      Left            =   8190
      Picture         =   "frmCadLancamentoPrecoPublico.frx":04A8
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Ativa Cadastro de Contribuintes"
      Top             =   2565
      Width           =   360
   End
   Begin VB.Frame fra_Receitas 
      Caption         =   " Receitas "
      Height          =   3180
      Left            =   75
      TabIndex        =   36
      Top             =   4725
      Width           =   8505
      Begin VB.CommandButton cmd_intReceita 
         Height          =   315
         Left            =   2985
         Picture         =   "frmCadLancamentoPrecoPublico.frx":0832
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de Receitas"
         Top             =   225
         Width           =   360
      End
      Begin VB.TextBox txtstrIndexador 
         Height          =   285
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   44
         Top             =   225
         Width           =   930
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5115
         MaxLength       =   18
         TabIndex        =   42
         Top             =   225
         Width           =   1530
      End
      Begin VB.TextBox txtintQtdeReceita 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3780
         MaxLength       =   7
         TabIndex        =   40
         Top             =   225
         Width           =   870
      End
      Begin TrueOleDBGrid70.TDBGrid tdbReceitas 
         Height          =   2145
         Left            =   120
         TabIndex        =   45
         Top             =   645
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   3784
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "intReceita"
         Columns(0).DataField=   "intReceita"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Receita"
         Columns(1).DataField=   "strReceita"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Qtde Index."
         Columns(2).DataField=   "dblQtdeIndexador"
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Abrev. Index."
         Columns(3).DataField=   "strIndexador"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor Index."
         Columns(4).DataField=   "dblValorIndexador"
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Valor R$"
         Columns(5).DataField=   "dblValor"
         Columns(5).NumberFormat=   "FormatText Event"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Qtde Rec."
         Columns(6).DataField=   "intQtdeReceita"
         Columns(6).NumberFormat=   "FormatText Event"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Valor Total R$"
         Columns(7).DataField=   "dblValorTotal"
         Columns(7).NumberFormat=   "FormatText Event"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Posição"
         Columns(8).DataField=   "intPosicao"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "strSigla"
         Columns(9).DataField=   "strSigla"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2937"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2858"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=1588"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=1508"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=8194"
         Splits(0)._ColumnProps(20)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=1799"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1720"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(27)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(29)=   "Column(4).Width=2170"
         Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=2090"
         Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=8194"
         Splits(0)._ColumnProps(34)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(36)=   "Column(5).Width=2011"
         Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=1931"
         Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(42)=   "Column(6).Width=1429"
         Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=1349"
         Splits(0)._ColumnProps(45)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(48)=   "Column(7).Width=2011"
         Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=1931"
         Splits(0)._ColumnProps(51)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=8194"
         Splits(0)._ColumnProps(53)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(54)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(55)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(56)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(58)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(60)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(61)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(62)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(63)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(64)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(65)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(66)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(67)=   "Column(9).Order=10"
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
         CellTips        =   2
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
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.wraptext=-1,.locked=0,.bold=0"
         _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000002&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000014&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.bgcolor=&HC0C0C0&,.locked=-1"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17,.bgcolor=&H80000016&"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HC0C0C0&"
         _StyleDefs(46)  =   ":id=46,.locked=-1"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.bgcolor=&HC0C0C0&,.locked=-1"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1,.bgcolor=&HC0C0C0&"
         _StyleDefs(55)  =   ":id=54,.locked=-1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1,.bgcolor=&HC0C0C0&"
         _StyleDefs(68)  =   ":id=66,.locked=-1"
         _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(80)  =   "Named:id=33:Normal"
         _StyleDefs(81)  =   ":id=33,.parent=0,.transparentBmp=0"
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
         _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(93)  =   "Named:id=39:EvenRow"
         _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(95)  =   "Named:id=40:OddRow"
         _StyleDefs(96)  =   ":id=40,.parent=33"
         _StyleDefs(97)  =   "Named:id=41:RecordSelector"
         _StyleDefs(98)  =   ":id=41,.parent=34"
         _StyleDefs(99)  =   "Named:id=42:FilterBar"
         _StyleDefs(100) =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintReceita 
         Height          =   315
         Left            =   135
         TabIndex        =   37
         Top             =   225
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbldblTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5820
         TabIndex        =   46
         Tag             =   "1"
         Top             =   2805
         Width           =   2460
      End
      Begin VB.Line lin_2 
         BorderColor     =   &H00000000&
         X1              =   135
         X2              =   8340
         Y1              =   3075
         Y2              =   3075
      End
      Begin VB.Line lin_3 
         BorderColor     =   &H00FFFFFF&
         X1              =   8355
         X2              =   8355
         Y1              =   2790
         Y2              =   3075
      End
      Begin VB.Line lin_1 
         X1              =   135
         X2              =   135
         Y1              =   2790
         Y2              =   3075
      End
      Begin VB.Label lblstrIndexador 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   195
         Left            =   6690
         TabIndex        =   43
         Top             =   285
         Width           =   705
      End
      Begin VB.Label lbldblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   4725
         TabIndex        =   41
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblintQtde 
         AutoSize        =   -1  'True
         Caption         =   "Qtde"
         Height          =   195
         Left            =   3405
         TabIndex        =   39
         Top             =   285
         Width           =   345
      End
   End
   Begin VB.ComboBox cbointTipoIdentificacao 
      Height          =   315
      ItemData        =   "frmCadLancamentoPrecoPublico.frx":0BBC
      Left            =   990
      List            =   "frmCadLancamentoPrecoPublico.frx":0BC9
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2190
      Width           =   1725
   End
   Begin VB.TextBox txtdtmDataVecto 
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
      Left            =   990
      MaxLength       =   19
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1845
      Width           =   1035
   End
   Begin VB.TextBox txtstrHistoricoPadrao 
      Height          =   885
      Left            =   1005
      MaxLength       =   750
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   900
      Width           =   7500
   End
   Begin VB.CommandButton cmd_intAssunto 
      Height          =   300
      Left            =   8130
      Picture         =   "frmCadLancamentoPrecoPublico.frx":0BE8
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "590"
      ToolTipText     =   "Ativa Cadastro de Assuntos"
      Top             =   105
      Width           =   360
   End
   Begin MSDataListLib.DataCombo dbcintAssunto 
      Height          =   315
      Left            =   1005
      TabIndex        =   1
      Top             =   105
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcintIdentificacao 
      Height          =   315
      Left            =   2790
      TabIndex        =   12
      Top             =   2190
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcintContribuinte 
      Height          =   315
      Left            =   990
      TabIndex        =   15
      Top             =   2565
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dbcintComposicao 
      Height          =   315
      Left            =   1005
      TabIndex        =   4
      Top             =   480
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lblintComposicao 
      AutoSize        =   -1  'True
      Caption         =   "Composição"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   540
      Width           =   870
   End
   Begin VB.Label lblProcesso 
      AutoSize        =   -1  'True
      Caption         =   "Processo"
      Height          =   195
      Left            =   285
      TabIndex        =   32
      Top             =   4380
      Width           =   660
   End
   Begin VB.Label lblintContribuinte 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Contribuinte"
      Height          =   195
      Left            =   105
      TabIndex        =   14
      Top             =   2625
      Width           =   840
   End
   Begin VB.Label lblintIdentificacao 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Identificação"
      Height          =   195
      Left            =   30
      TabIndex        =   10
      Top             =   2250
      Width           =   915
   End
   Begin VB.Label lbldtmDataVecto 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1905
      Width           =   840
   End
   Begin VB.Label lblstrHistoricoPadrao 
      AutoSize        =   -1  'True
      Caption         =   "Histórico"
      Height          =   195
      Left            =   345
      TabIndex        =   6
      Top             =   900
      Width           =   615
   End
   Begin VB.Label lblintAssunto 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Assunto"
      Height          =   195
      Left            =   390
      TabIndex        =   0
      Top             =   195
      Width           =   570
   End
End
Attribute VB_Name = "frmCadLancamentoPrecoPublico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando     As Boolean
Dim mobjAux           As Object
Dim mblnSelecionou    As Boolean
Dim mblnClickOk       As Boolean
Dim bytDividaAtiva    As Byte

Dim xadbReceitas      As XArrayDB


Private Sub cmd_intComposicao_Click()
    CarregaForm frmCadComposicaoDaReceita, dbcintComposicao
End Sub

Private Sub dbcintComposicao_Change()
    dbcintReceita.BoundText = Space$(0)
    Set dbcintReceita.RowSource = Nothing
    dbcintReceita.Tag = strQueryReceitas & ";strDescricao"
End Sub

Private Sub dbcintComposicao_Click(Area As Integer)
    DropDownDataCombo dbcintComposicao, Me, Area
End Sub

Private Sub dbcintComposicao_GotFocus()
    MarcaCampo dbcintComposicao
End Sub

Private Sub dbcintComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintComposicao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintComposicao
End Sub

Private Sub cbointTipoIdentificacao_Click()

    TrocaCorObjeto dbcintIdentificacao, cbointTipoIdentificacao.ListIndex = 0
    
    If cbointTipoIdentificacao.ListIndex > -1 Then
        CarregaIdentificacao cbointTipoIdentificacao.ItemData(cbointTipoIdentificacao.ListIndex)
    End If
    
End Sub

Private Sub cmd_intAssunto_Click()
    CarregaForm frmCadCatalogoAssunto, dbcintAssunto
End Sub

Private Sub cmd_intContribuinte_Click()
    ChamaFormCadastro frmCadContribuinte, dbcintContribuinte
End Sub

Private Sub cmd_intIdentificacao_Click()
    If cbointTipoIdentificacao.ListIndex = 1 Then
        ChamaFormCadastro frmCadImobiliario, dbcintIdentificacao
    ElseIf cbointTipoIdentificacao.ListIndex = 2 Then
        ChamaFormCadastro frmCadEconomico, dbcintIdentificacao
    End If
End Sub

Private Sub cmd_intReceita_Click()
    ChamaFormCadastro frmCadReceita, dbcintReceita
End Sub

Private Sub dbcintAssunto_Change()
    If dbcintAssunto.MatchedWithList Then
        CarregaDadosAssunto
        PreencheGrid
        If dbcintAssunto.Enabled = True Then dbcintAssunto.SetFocus
    Else
        LimpaGrid
    End If
End Sub

Private Sub dbcintAssunto_Click(Area As Integer)
    If Area = 0 Then
       DropDownDataCombo dbcintAssunto, Me, Area
    End If
End Sub

Private Sub dbcintAssunto_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintAssunto, Me, , KeyCode, Shift
End Sub

Private Sub dbcintAssunto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintAssunto
End Sub

Private Sub dbcintContribuinte_Change()
    If dbcintContribuinte.MatchedWithList And Not dbcintIdentificacao.MatchedWithList Then
        CarregaDadosContribuinte dbcintContribuinte.BoundText
    End If
End Sub

Private Sub dbcintContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinte_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContribuinte
End Sub

Private Sub dbcintIdentificacao_Change()
    If dbcintIdentificacao.MatchedWithList Then
        CarregaDadosIdentificacao dbcintIdentificacao.BoundText
    End If
End Sub

Private Sub dbcintIdentificacao_GotFocus()
    MarcaCampo dbcintIdentificacao
End Sub

Private Sub dbcintReceita_Change()
    If dbcintReceita.MatchedWithList Then
        CarregaDadosReceita dbcintReceita.BoundText
    End If
    dbcintReceita.ToolTipText = Trim(dbcintReceita.Text)
End Sub

Private Sub dbcintReceita_GotFocus()
    MarcaCampo dbcintReceita
End Sub

Private Sub tdbReceitas_AfterColEdit(ByVal ColIndex As Integer)
Dim intFor As Integer
    
    With tdbReceitas
        If Not .EOF And Not .BOF Then
            'If Len(Trim(.Columns("intReceita").Value)) > 0 Then
                If ColIndex = 5 Or ColIndex = 6 Then
                    For intFor = 0 To xadbReceitas.UpperBound(1)
                        If xadbReceitas(intFor, 8) = .Columns("intPosicao").Value Then
                            
                            If Len(.Columns(ColIndex).Value) = 0 Then .Columns(ColIndex).Value = 0
                            
                            xadbReceitas(intFor, ColIndex) = .Columns(ColIndex).Value
                            
                            'Vamos atualizar o valor Total do Grid
                            lbldblTotal.Caption = gstrConvVrDoSql(lbldblTotal - gstrConvVrDoSql(tdbReceitas.Columns("dblValorTotal").Value, , , True), 2)
                            
                            tdbReceitas.Columns("dblValorTotal").Value = tdbReceitas.Columns("dblValor").Value * IIf(tdbReceitas.Columns("intQtdeReceita").Value = "", 0, tdbReceitas.Columns("intQtdeReceita").Value)
                            xadbReceitas(intFor, 7) = tdbReceitas.Columns("dblValorTotal").Value
                            
                            lbldblTotal.Caption = gstrConvVrDoSql(lbldblTotal + tdbReceitas.Columns("dblValorTotal").Value, 2)
                            
                            Exit For
                        
                        End If
                    Next
                End If
            'End If
        End If
    End With

End Sub

Private Sub tdbReceitas_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    
    With tdbReceitas
        If .EOF Or .BOF Or Len(Trim(.Columns("intReceita").Value)) = 0 Then
            Cancel = 1
        ElseIf ColIndex = 5 And tdbReceitas.Columns("dblValor").Value <> 0 Then
            Cancel = 1
        End If
        
    End With

End Sub

Private Sub tdbReceitas_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    
    Select Case ColIndex
        Case Is = 2, 4
            Value = gstrConvVrDoSql(Value, 6)
        Case Is = 5, 7
            Value = gstrConvVrDoSql(Value, 2)
    End Select

End Sub

Private Sub tdbReceitas_KeyPress(KeyAscii As Integer)
        
    If tdbReceitas.Col = 5 Then
        CaracterValido KeyAscii, "V", tdbReceitas
    ElseIf tdbReceitas.Col = 6 Then
        CaracterValido KeyAscii, "N", tdbReceitas
        If Len(tdbReceitas.Columns(6).Value) >= 4 Then KeyAscii = 0
    End If

End Sub

Private Sub txtbitDigito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigito
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And txtdblValor.SelLength <> Len(txtdblValor) Then
        If Len(Mid(txtdblValor, InStr(1, txtdblValor, ",") + 1, Len(txtdblValor))) = 9 Then
            KeyAscii = 0
        End If
    End If
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblValor = gstrConvVrDoSql(txtdblValor, 6)
End Sub

Private Sub txtintExercicio_GotFocus()
    If txtintExercicio = "" Then
        txtintExercicio.Text = Year(gstrDataDoSistema())
    End If
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintQtdeReceita_GotFocus()
    MarcaCampo txtintQtdeReceita
End Sub

Private Sub txtintQtdeReceita_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And txtintQtdeReceita.SelLength <> Len(txtintQtdeReceita) Then
        If InStr(1, txtintQtdeReceita, ",") > 0 Then
            If Len(Mid(txtintQtdeReceita, InStr(1, txtintQtdeReceita, ",") + 1, Len(txtintQtdeReceita))) = 2 Then
                KeyAscii = 0
            End If
        End If
    End If
    CaracterValido KeyAscii, "V", txtintQtdeReceita
End Sub

Private Sub txtintQtdeReceita_LostFocus()
    txtintQtdeReceita = gstrConvVrDoSql(txtintQtdeReceita, 2)
End Sub

Private Sub txtstrCodigo_GotFocus()
     MarcaCampo txtstrCodigo
End Sub

Private Sub txtbitDigito_GotFocus()
    MarcaCampo txtbitDigito
End Sub

Private Sub txtdtmDataVecto_LostFocus()
    txtdtmDataVecto.Text = gstrDataFormatada(txtdtmDataVecto.Text)
End Sub

Private Sub txtdtmDataVecto_GotFocus()
    MarcaCampo txtdtmDataVecto
    If txtdtmDataVecto = "" Then txtdtmDataVecto = gstrDataDoSistema
End Sub

Private Sub txtdtmDataVecto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataVecto
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1204
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
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrImprimir
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub Form_Load()
    
    TrocaCorObjeto dbcintIdentificacao, True
    TrocaCorObjeto txtstrLogradouro, True
    TrocaCorObjeto txtintNumero, True
    TrocaCorObjeto txtstrComplemento, True
    TrocaCorObjeto txtstrBairro, True
    TrocaCorObjeto txtstrMunicipio, True
    TrocaCorObjeto txtstrUF, True
    TrocaCorObjeto txtintCep, True
    TrocaCorObjeto txtdblValor, True
    TrocaCorObjeto txtstrIndexador, True
    
    dbcintAssunto.Tag = strQueryAssuntos & ";strDescricao"
    dbcintComposicao.Tag = strQueryComposicao & ";strDescricao"
    dbcintContribuinte.Tag = strQueryContribuintes & ";strNome"
    dbcintReceita.Tag = strQueryReceitas & ";strDescricao"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case UCase(strModoOperacao)
        
        Case Is = gstrPreencherLista
            PreencherListaDeOpcoes Me.ActiveControl
    
        Case Is = UCase(gstrNovo)
            LimpaObjeto Me
            LimpaGrid
            TrocaCorObjeto txtdblValor, True
            
        Case Is = UCase(gstrIncluirItem)
            IncluiReceitaNoGrid
        Case Is = UCase(gstrExcluirItem)
            ExcluiReceitaNoGrid
        
        Case Is = UCase(gstrCalcularReajuste)
            If blnDadosOk Then
                If GeraPrecoPublico Then
                    
                    If UCase(App.ProductName) = "PROTOCOLO" Then
                        Unload Me
                    Else
                        ExibeMensagem "Operação concluída."
                    End If
                    
                End If
            End If
            
    End Select
    
End Sub

Private Sub CarregaDadosAssunto()
Dim adoRec As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO("SELECT strHistoricoPadrao FROM " & gstrCatalogoAssunto & " WHERE PkID = " & dbcintAssunto.BoundText, 10, adoRec) Then
    
        With adoRec
            If Not (.BOF And .EOF) Then
                txtstrHistoricoPadrao = gstrENulo(!strHistoricoPadrao)
            Else
                txtstrHistoricoPadrao = Space$(0)
            End If
        End With
        
    End If
    
    Set gobjBanco = Nothing

End Sub

Private Sub CarregaDadosContribuinte(lngContribuinte As Long)
Dim adoRec As ADODB.Recordset
Dim strSql As String

    Set gobjBanco = New clsBanco
        
    strSql = "SELECT CO.STRLOGRADOUROC LogradouroC, CO.intnumeroC, CO.strcomplementoC, CO.intcepC, CO.strBairroC, UF.strSigla, MU.strDescricao MunicipioC " & _
             "FROM " & gstrContribuinte & " CO, " & gstrUF & " UF, " & gstrCidade & " MU " & _
             "WHERE UF.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " CO.intUfC AND " & _
             "MU.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " CO.intMunicipioC AND " & _
             "CO.pkid = " & lngContribuinte

    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
            If Not (.BOF And .EOF) Then
                txtstrLogradouro.Text = Space$(0) & adoRec("LogradouroC")
                txtintNumero.Text = Space$(0) & adoRec("intNumeroC")
                txtstrComplemento.Text = Space$(0) & adoRec("strComplementoC")
                txtstrBairro.Text = Space$(0) & adoRec("strBairroC")
                txtstrMunicipio.Text = Space$(0) & adoRec("MunicipioC")
                txtstrUF.Text = Space$(0) & adoRec("strSigla")
                txtintCep.Text = Space$(0) & adoRec("intCepC")
            Else
                txtstrLogradouro.Text = Space$(0)
                txtintNumero.Text = Space$(0)
                txtstrComplemento.Text = Space$(0)
                txtstrBairro.Text = Space$(0)
                txtstrMunicipio.Text = Space$(0)
                txtstrUF.Text = Space$(0)
                txtintCep.Text = Space$(0)
            End If
        End With
        
    End If
    
    Set gobjBanco = Nothing

End Sub

Private Sub CarregaDadosIdentificacao(lngIdentificacao As Long)
Dim adoRec As ADODB.Recordset
Dim strSql As String

    Set gobjBanco = New clsBanco
        
    If bytDBType = Oracle Then
    strSql = "SELECT TL.STRSIGLA " & strCONCAT & "' '" & strCONCAT & " TTL.STRSIGLA " & strCONCAT & "' '" & strCONCAT & " LO.STRDESCRICAO Logradouro, IM.intnumero, IM.strcomplemento, IM.intcep, BA.STRDESCRICAO Bairro, IM.intContribuinte " & _
             "FROM " & IIf(cbointTipoIdentificacao.ListIndex = 1, gstrImobiliario, gstrEconomico) & " IM, " & gstrLogradouro & " LO, " & gstrTipoLogradouro & " TL, " & gstrTituloLogradouro & " TTL, " & gstrBairro & " BA " & _
             "WHERE LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " IM.Intlogradouro AND " & _
             "TL.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.INTTIPOLOGRADOURO AND " & _
             "TTL.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.INTTITULOLOGRADOURO AND " & _
             "BA.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.intBairro AND IM.pkid = " & lngIdentificacao
    Else
        'Criado select específico para SQL pois não aceita relacionamento no "Where"
        strSql = strSql & "SELECT "
        strSql = strSql & gstrISNULL("TL.strSigla", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("TTL.strSigla", "''") & strCONCAT & "' '" & strCONCAT & gstrISNULL("LO.strDescricao", "''") & " AS Logradouro, "
        strSql = strSql & " IM.intNumero, "
        strSql = strSql & " IM.strComplemento, "
        strSql = strSql & " IM.intCEP, "
        strSql = strSql & " BA.strDescricao AS Bairro, "
        strSql = strSql & " IM.intContribuinte "
        strSql = strSql & "FROM "
        strSql = strSql & IIf(cbointTipoIdentificacao.ListIndex = 1, gstrImobiliario, gstrEconomico) & " IM INNER JOIN "
        strSql = strSql & gstrBairro & " BA ON IM.intBairro = BA.PKId LEFT OUTER JOIN "
        strSql = strSql & gstrLogradouro & " LO ON IM.intLogradouro = LO.PKId LEFT OUTER JOIN "
        strSql = strSql & gstrTipoLogradouro & " TL ON LO.intTipoLogradouro = TL.PKId LEFT OUTER JOIN "
        strSql = strSql & gstrTituloLogradouro & " TTL ON LO.intTituloLogradouro = TTL.PKId "
        strSql = strSql & "WHERE IM.pkid = " & lngIdentificacao
    End If

    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
            If Not (.BOF And .EOF) Then
                AtribuiValorDoSql dbcintContribuinte, adoRec("intContribuinte")
                txtstrLogradouro.Text = Space$(0) & adoRec("Logradouro")
                txtintNumero.Text = Space$(0) & adoRec("intNumero")
                txtstrComplemento.Text = Space$(0) & adoRec("strComplemento")
                txtstrBairro.Text = Space$(0) & adoRec("Bairro")
                txtstrMunicipio.Text = gstrCidadeEmpresa
                txtstrUF.Text = gstrUFEmpresa
                txtintCep.Text = Space$(0) & adoRec("intCep")
            Else
                dbcintContribuinte.BoundText = Space$(0)
                txtstrLogradouro.Text = Space$(0)
                txtintNumero.Text = Space$(0)
                txtstrComplemento.Text = Space$(0)
                txtstrBairro.Text = Space$(0)
                txtstrMunicipio.Text = Space$(0)
                txtstrUF.Text = Space$(0)
                txtintCep.Text = Space$(0)
            End If
        End With
        
    End If
    
    Set gobjBanco = Nothing

End Sub

Private Sub CarregaDadosReceita(lngReceita As Long)
Dim adoRec As ADODB.Recordset
Dim strSql As String

    Set gobjBanco = New clsBanco
        
    strSql = "SELECT RE.dblPrecoPublico, IE.strAbreviatura " & _
             "FROM " & gstrReceitasExercicio & " RE, " & gstrIndexadorEconomico & " IE " & _
             "WHERE RE.intReceita = " & lngReceita & " AND " & _
             "RE.intExercicio = " & Year(gstrDataDoSistema) & " AND " & _
             "IE.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " RE.intIndexadorEconomico"

    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
            If Not (.BOF And .EOF) Then
                txtdblValor.Text = Space$(0) & gstrConvVrDoSql(adoRec("dblPrecoPublico").Value, 6, , True)
                txtstrIndexador.Text = Space$(0) & adoRec("strAbreviatura").Value
                txtintQtdeReceita.Text = "1,00"
                TrocaCorObjeto txtdblValor, Not gstrConvVrDoSql(adoRec("dblPrecoPublico").Value, , , True) = 0
            Else
                txtdblValor.Text = Space$(0)
                txtstrIndexador.Text = Space$(0)
                txtintQtdeReceita.Text = Space$(0)
            End If
        End With
        
    End If
    
    Set gobjBanco = Nothing

End Sub

Sub PreencheGrid()
Dim c       As Integer
    
    c = tdbReceitas.Col
    tdbReceitas.HoldFields
    
    MontaArray
    
    tdbReceitas.Col = c
    tdbReceitas.EditActive = True
    tdbReceitas.CurrentCellModified = True
    
End Sub

Private Sub MontaArray()
Dim adoResultado      As ADODB.Recordset
Dim intFor            As Integer
Dim strSql            As String
Dim varAux            As Variant

Dim dblValorTotal     As Double

On Error GoTo Problema_Na_Rotina
    
    'Vamos obter os valores das parcelas da inscricao selecionada
    Set gobjBanco = New clsBanco
    
    Set xadbReceitas = New XArrayDB
    xadbReceitas.Clear
    
    strSql = "SELECT RE.strsigla, AR.intReceita, RX.dblPrecoPublico,  RE.strDescricao strReceita, IE.strAbreviatura, "
             
    If bytDBType = Oracle Then
    
        strSql = strSql & "(SELECT FA.DBLVALOR FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = RX.INTINDEXADORECONOMICO AND FA.DTMDATA = '" & Date & "') dblValorDataAtual, " & _
                          "(SELECT FA.DBLVALOR FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = RX.INTINDEXADORECONOMICO AND FA.DTMDATA = '01/" & Month(gstrDataDoSistema) & "/" & Year(gstrDataDoSistema) & "') dblValorMes "
    Else
        strSql = strSql & "(SELECT FA.DBLVALOR FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = RX.INTINDEXADORECONOMICO AND FA.DTMDATA = '" & Format$(Date, "MM/DD/YYYY") & "') dblValorDataAtual, " & _
                          "(SELECT FA.DBLVALOR FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = RX.INTINDEXADORECONOMICO AND FA.DTMDATA = '" & Month(gstrDataDoSistema) & "/01/" & Year(gstrDataDoSistema) & "') dblValorMes "
    
    End If
    
    strSql = strSql & "FROM " & gstrCatalogoAssuntoReceita & " AR, " & gstrReceita & " RE, " & gstrReceitasExercicio & " RX, " & gstrIndexadorEconomico & " IE " & _
                      "WHERE AR.intCatalogoAssunto = " & dbcintAssunto.BoundText & " AND " & _
                      "RE.pkid = AR.intReceita AND " & _
                      "RX.INTRECEITA = RE.PKID AND " & _
                      "RX.Intexercicio = " & Year(gstrDataDoSistema) & " AND " & _
                      "IE.Pkid = RX.INTINDEXADORECONOMICO"
      
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
                
            If Not adoResultado.EOF Then
            
                For intFor = 0 To adoResultado.RecordCount - 1
            
                    xadbReceitas.ReDim 0, intFor, 0, 9
                
                    varAux = Space$(0) & adoResultado("intReceita").Value
                    xadbReceitas(intFor, 0) = varAux
                    varAux = Space$(0) & adoResultado("strReceita").Value
                    xadbReceitas(intFor, 1) = varAux
                    
                    'Caso exista indexador a qtde do indexador sera o valor da receita
                    If Len(adoResultado("strAbreviatura")) > 0 Then
                        varAux = gstrConvVrDoSql(adoResultado("dblPrecoPublico").Value)
                    Else
                        varAux = Space$(0)
                    End If
                    xadbReceitas(intFor, 2) = varAux
                    
                    varAux = Space$(0) & adoResultado("strAbreviatura").Value
                    xadbReceitas(intFor, 3) = varAux
                    
                    'Caso exista indexador vamos colocar o valor do indexador
                    If Len(adoResultado("strAbreviatura")) > 0 Then
                        If Not IsNull(adoResultado("dblValorDataAtual").Value) Then
                            varAux = Space$(0) & gstrConvVrDoSql(adoResultado("dblValorDataAtual").Value, 6)
                        ElseIf Not IsNull(adoResultado("dblValorMes").Value) Then
                            varAux = Space$(0) & gstrConvVrDoSql(adoResultado("dblValorMes").Value, 6)
                        Else
                            ExibeMensagem "Não foi encontrado valor para o Indexador Econômico (" & adoResultado("strAbreviatura") & ")"
                            LimpaGrid
                            dbcintAssunto.BoundText = Space$(0)
                            Exit Sub
                        End If
                    Else
                        varAux = Space$(0)
                    End If
                    xadbReceitas(intFor, 4) = varAux
                    
                    If Val(xadbReceitas(intFor, 4)) = 0 Then
                        varAux = Space$(0) & gstrConvVrDoSql(adoResultado("dblPrecoPublico").Value)
                    Else
                        varAux = Space$(0) & gstrConvVrDoSql(xadbReceitas(intFor, 2) * xadbReceitas(intFor, 4))
                    End If
                    xadbReceitas(intFor, 5) = varAux
                    
                    varAux = Space$(0) & "1,00"
                    xadbReceitas(intFor, 6) = varAux
                    varAux = Space$(0) & gstrConvVrDoSql(xadbReceitas(intFor, 5))
                    xadbReceitas(intFor, 7) = varAux
                    varAux = Space$(0) & intFor
                    xadbReceitas(intFor, 8) = varAux
                    
                    varAux = Space$(0) & adoResultado("strSigla")
                    xadbReceitas(intFor, 9) = varAux

                    dblValorTotal = dblValorTotal + xadbReceitas(intFor, 7)
                    
                    adoResultado.MoveNext
                    
                Next
                
            Else
                xadbReceitas.ReDim 0, 0, 0, 9
                xadbReceitas(0, 0) = ""
                xadbReceitas(0, 1) = ""
                xadbReceitas(0, 2) = ""
                xadbReceitas(0, 3) = ""
                xadbReceitas(0, 4) = ""
                xadbReceitas(0, 5) = ""
                xadbReceitas(0, 6) = ""
                xadbReceitas(0, 7) = ""
                xadbReceitas(0, 8) = ""
                xadbReceitas(0, 9) = ""
                
                dblValorTotal = 0
                
            End If
    
            Set tdbReceitas.Array = xadbReceitas
            tdbReceitas.ReBind
            tdbReceitas.Refresh
            
            'Vamos atualizar o valor Total do Grid
            lbldblTotal.Caption = gstrConvVrDoSql(dblValorTotal, 2)

        End With
        
    End If
    
    Exit Sub
    
Problema_Na_Rotina:
    ExibeDetalheErro Err.Description
    Exit Sub
    
End Sub

Private Sub CarregaIdentificacao(bytTipoIdentificacao As Byte)
    
    If bytTipoIdentificacao = TYP_IMOBILIARIA Then
        dbcintIdentificacao.Tag = "SELECT Pkid, " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao FROM " & gstrImobiliario & " Where Dtmdtcancelamento is null ORDER BY strInscricao;strInscricao"
    ElseIf bytTipoIdentificacao = TYP_ECONOMICA Then
        dbcintIdentificacao.Tag = "SELECT Pkid, " & gstrRIGHT("strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricaoCadastral FROM " & gstrEconomico & " ORDER BY strInscricaoCadastral;strInscricaoCadastral"
    Else
        dbcintIdentificacao.Tag = ""
    End If
    
    dbcintIdentificacao.BoundText = ""
    Set dbcintIdentificacao.RowSource = Nothing
    
End Sub

Private Sub LimpaGrid()
    
    Set xadbReceitas = New XArrayDB
    xadbReceitas.Clear
    xadbReceitas.ReDim -1, -1, 0, 9
            
    Set tdbReceitas.Array = xadbReceitas
    tdbReceitas.ReBind
    tdbReceitas.Refresh

End Sub

Private Sub IncluiReceitaNoGrid()
    Dim varAux            As Variant
    Dim dblValorIndexador As Double
    Dim adoRec            As ADODB.Recordset
    Dim adoRecAux         As ADODB.Recordset
    Dim strSql            As String
    Dim intPosicao        As Integer
    Dim strsigla          As String
    Dim dblValorEmReais   As Double
    
    If Not dbcintReceita.MatchedWithList Then
        ExibeMensagem "É preciso selecionar alguma Receita."
        Exit Sub
    End If
            
    If Not dbcintAssunto.MatchedWithList Then
        ExibeMensagem "É preciso selecionar algum Assunto."
        Exit Sub
    End If
    
    If txtdblValor.Enabled = True Then
        If CDbl(gstrConvVrDoSql(txtdblValor.Text, 6, , True)) <= 0 Then
            ExibeMensagem "É preciso informar um valor."
            txtdblValor.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtintQtdeReceita) <= 0 Then
        ExibeMensagem "É preciso informar a quantidade de receitas."
        txtdblValor.SetFocus
        Exit Sub
    End If
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO("Select strSigla, BYTINSCREVEDA From " & gstrReceita & " Where pkid = " & dbcintReceita.BoundText, 5, adoRecAux) Then
        If Not adoRecAux.EOF Then
            If blnVerificaReceitas Then
                If bytDividaAtiva <> Val(gstrENulo(adoRecAux!BYTINSCREVEDA)) Then
                    ExibeMensagem "Já foi incluída(s) receitas que " & IIf(bytDividaAtiva = 0, "não inscreve(m) em dívida ativa.", "inscreve(m) em dívida ativa.") & Chr(13) & "Incluir receitas de mesma situação."
                    Exit Sub
                End If
            End If
            strsigla = gstrENulo(adoRecAux!strsigla)
            bytDividaAtiva = Val(gstrENulo(adoRecAux!BYTINSCREVEDA))
        End If
    End If

    
    'Vamos obter os valor do indexador
    If bytDBType = Oracle Then
        strSql = "SELECT RE.dblPrecoPublico, IE.strAbreviatura, " & _
                 "(SELECT FA.DBLVALOR FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = RE.INTINDEXADORECONOMICO AND FA.DTMDATA = " & gstrConvDtParaSql(gstrDataDoSistema) & ") dblValorDataAtual, " & _
                 "(SELECT FA.DBLVALOR FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = RE.INTINDEXADORECONOMICO AND FA.DTMDATA = " & gstrConvDtParaSql("01/" & Month(gstrDataDoSistema) & "/" & Year(gstrDataDoSistema)) & ") dblValorMes " & _
                 "FROM " & gstrReceitasExercicio & " RE, " & gstrIndexadorEconomico & " IE " & _
                 "WHERE RE.intReceita = " & dbcintReceita.BoundText & " AND " & _
                 "RE.intExercicio = " & Year(gstrDataDoSistema) & " AND " & _
                 "IE.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " RE.intIndexadorEconomico"
    Else
        'Criado select específico para SQL pois não aceita relacionamento no "Where"
        strSql = "SELECT RE.dblPrecoPublico, IE.strAbreviatura, "
        strSql = strSql & "(SELECT FA.DBLVALOR FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = RE.INTINDEXADORECONOMICO AND FA.DTMDATA = " & gstrConvDtParaSql(gstrDataDoSistema) & ") dblValorDataAtual, "
        strSql = strSql & "(SELECT FA.DBLVALOR FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = RE.INTINDEXADORECONOMICO AND FA.DTMDATA = " & gstrConvDtParaSql("01/" & Month(gstrDataDoSistema) & "/" & Year(gstrDataDoSistema)) & ") dblValorMes "
        strSql = strSql & "FROM "
        strSql = strSql & gstrReceitasExercicio & "  RE LEFT OUTER JOIN "
        strSql = strSql & gstrIndexadorEconomico & " IE ON RE.INTINDEXADORECONOMICO = IE.PKId "
        strSql = strSql & "WHERE "
        strSql = strSql & " RE.intReceita = " & dbcintReceita.BoundText & " AND RE.intExercicio = " & Year(gstrDataDoSistema) & " "
    End If

    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
            If Not (.BOF And .EOF) Then
                'Caso exista indexador vamos colocar o valor do indexador
                If Len(adoRec("strAbreviatura")) > 0 Then
                    If Not IsNull(adoRec("dblValorDataAtual")) Then
                        dblValorIndexador = Space$(0) & gstrConvVrDoSql(adoRec("dblValorDataAtual").Value, 6)
                    ElseIf Not IsNull(adoRec("dblValorMes")) Then
                        dblValorIndexador = Space$(0) & gstrConvVrDoSql(adoRec("dblValorMes").Value, 6)
                    Else
                        ExibeMensagem "Não foi encontrado valor para o Indexador Econômico."
                        Exit Sub
                    End If
                End If
            Else
                ExibeMensagem "Não foi possível encontrar o Indexador Econômico."
                Exit Sub
            End If
        End With
        
    End If
    
    If xadbReceitas.UpperBound(1) > -1 Then
        'caso ja exista uma linha em branco nao vamos criar outra
        If Len(Trim(xadbReceitas(xadbReceitas.UpperBound(1), 8))) = 0 Then
            intPosicao = 0
        Else
            intPosicao = Val(xadbReceitas(xadbReceitas.UpperBound(1), 8)) + 1
            xadbReceitas.AppendRows 1
        End If
    Else
        intPosicao = 0
        xadbReceitas.AppendRows 1
    End If
    
    varAux = Space$(0) & dbcintReceita.BoundText
    xadbReceitas(xadbReceitas.UpperBound(1), 0) = varAux                'IntReceita
    
    
    varAux = strsigla
    xadbReceitas(xadbReceitas.UpperBound(1), 9) = varAux                'strSigla
    
    Set gobjBanco = Nothing
    
    varAux = Space$(0) & dbcintReceita.Text
    xadbReceitas(xadbReceitas.UpperBound(1), 1) = varAux                'Receita
    
    'Caso exista indexador a qtde do indexador sera o valor da receita
    
    If Trim(gstrENulo(adoRec!Strabreviatura)) <> "" Then
        If txtdblValor.Enabled = False Then
            varAux = gstrConvVrDoSql(adoRec("dblPrecoPublico").Value, 6)
        Else
            varAux = Space$(0) & gstrConvVrDoSql(txtdblValor.Text, 6, 6)
        End If
    Else
        varAux = Space$(0) & "0,000000"
    End If
    xadbReceitas(xadbReceitas.UpperBound(1), 2) = varAux                'Qtde de Indexdor
    
    varAux = Space$(0) & txtstrIndexador.Text
    xadbReceitas(xadbReceitas.UpperBound(1), 3) = varAux                'Abrev. Indexador
    varAux = Space$(0) & dblValorIndexador
    xadbReceitas(xadbReceitas.UpperBound(1), 4) = varAux                'Valor Indexador
    
    If Trim(txtstrIndexador.Text) <> "" Then 'Tem Indexador
        If txtdblValor.Enabled = False Then
            varAux = Space$(0) & gstrConvVrDoSql(dblValorIndexador * xadbReceitas(xadbReceitas.UpperBound(1), 2), 6)
        Else
            varAux = Space$(0) & gstrConvVrDoSql(dblValorIndexador * gstrConvVrDoSql(txtdblValor.Text, 6, 6), 6)
        End If
    Else
        If txtdblValor.Enabled = False Then 'Não tem indexador
            varAux = Space$(0) & gstrConvVrDoSql(txtdblValor.Text, 6, 6)
        Else
            varAux = Space$(0) & gstrConvVrDoSql(txtdblValor.Text, 6, 6)
        End If
    End If
    'Vamos armazenar o valor sem truncar para utilizar no calculo da receita
    dblValorEmReais = varAux
    varAux = Mid(gstrConvVrDoSql(varAux, 6), 1, InStr(varAux, ",") - 1) & Mid(gstrConvVrDoSql(varAux, 6), InStr(varAux, ","), 3)
    xadbReceitas(xadbReceitas.UpperBound(1), 5) = varAux                'Valor
    
    varAux = Space$(0) & txtintQtdeReceita.Text
    xadbReceitas(xadbReceitas.UpperBound(1), 6) = varAux                'Qtde Rec
    
    varAux = Space$(0) & gstrConvVrDoSql(dblValorEmReais * xadbReceitas(xadbReceitas.UpperBound(1), 6), 6)
    varAux = Mid(gstrConvVrDoSql(varAux, 6), 1, InStr(varAux, ",") - 1) & Mid(gstrConvVrDoSql(varAux, 6), InStr(varAux, ","), 3)
    xadbReceitas(xadbReceitas.UpperBound(1), 7) = varAux                'Valor Total
    
    varAux = intPosicao
    xadbReceitas(xadbReceitas.UpperBound(1), 8) = varAux                'Posição
    
    'Vamos atualizar o valor Total do Grid
    lbldblTotal.Caption = gstrConvVrDoSql(CCur(lbldblTotal) + xadbReceitas(xadbReceitas.UpperBound(1), 7), 2)
    
    Set tdbReceitas.Array = xadbReceitas
    tdbReceitas.ReBind
    tdbReceitas.Refresh
    
    If gstrConvVrDoSql(lbldblTotal.Caption) > 999999999.99 Then
        ExibeMensagem "A soma dos valores é superior ao máximo permitido."
        tdbReceitas.MoveLast
        ExcluiReceitaNoGrid
        Exit Sub
    End If
    
    dbcintReceita.BoundText = Space$(0)
    txtintQtdeReceita.Text = Space$(0)
    txtdblValor.Text = Space$(0)
    txtstrIndexador.Text = Space$(0)
    
End Sub

Private Sub ExcluiReceitaNoGrid()
Dim varAux As Variant
Dim intFor As Integer

    If tdbReceitas.EOF Then
        ExibeMensagem "É preciso selecionar alguma Receita da lista."
        Exit Sub
    End If
            
    For intFor = 0 To xadbReceitas.UpperBound(1)
        
        If xadbReceitas(intFor, 8) = tdbReceitas.Columns("intPosicao") Then
            
            'Vamos atualizar o valor Total do Grid
            lbldblTotal.Caption = gstrConvVrDoSql(lbldblTotal - xadbReceitas(intFor, 7), 2)
            
            xadbReceitas.DeleteRows intFor
            
            Exit For
            
        End If
        
    Next
    
    Set tdbReceitas.Array = xadbReceitas
    tdbReceitas.ReBind
    tdbReceitas.Refresh
    
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If Not dbcintAssunto.MatchedWithList Then
        MsgBox "O Assunto deve ser selecionado.", vbOKOnly, "Mensagem ao Usuário"
        dbcintAssunto.SetFocus
        Exit Function
    End If
    
    If Not dbcintComposicao.MatchedWithList Then
        MsgBox "A composição deve ser selecionada.", vbOKOnly, "Mensagem ao Usuário"
        dbcintComposicao.SetFocus
        Exit Function
    End If
    
    If txtdtmDataVecto = Space$(0) Then
       MsgBox "A Data de Vencimento deve ser preenchida.", vbOKOnly, "Mensagem ao Usuário"
       txtdtmDataVecto.SetFocus
       Exit Function
    End If
    
    If gblnDataValida(txtdtmDataVecto.Text) = False Then
       MsgBox "Data de Vencimento inválida.", vbOKOnly, "Mensagem ao Usuário"
       txtdtmDataVecto.SetFocus
       Exit Function
    End If
    
    'If CDate(txtdtmDataVecto) < CDate(gstrDataDoSistema) Then
    '   ExibeMensagem "Data de Vencimento não pode ser menor que a data do dia."
    '   txtdtmDataVecto.SetFocus
    '   Exit Function
    'End If
    
    If cbointTipoIdentificacao.ListIndex <= 0 Then
        If Not dbcintContribuinte.MatchedWithList Then
            ExibeMensagem "O contribuinte deve ser selecionado"
            If dbcintContribuinte.Enabled = True Then dbcintContribuinte.SetFocus
            Exit Function
        End If
    End If
    
    
    If cbointTipoIdentificacao.ListIndex > 0 Then
        If Not dbcintIdentificacao.MatchedWithList Then
            MsgBox "A Identificação deve ser selecionada.", vbOKOnly, "Mensagem ao Usuário"
            dbcintIdentificacao.SetFocus
            Exit Function
        End If
    End If
    
    If Len(Trim(txtstrCodigo.Text)) > 0 Or Len(Trim(txtintExercicio.Text)) > 0 Or Len(Trim(txtbitDigito.Text)) > 0 Then
        If Not VerificaEmpenhoProcesso(Trim(txtstrCodigo), Val(txtbitDigito), Val(txtintExercicio)) Then
            MsgBox "O Processo não é válido.", vbOKOnly, "Mensagem ao Usuário"
            txtstrCodigo.SetFocus
            Exit Function
        End If
    End If
    
    If blnVerificaReceitas = False Then
        ExibeMensagem "Não há nenhuma receita para gerar o preço público."
        Exit Function
    End If
    
    blnDadosOk = True
       
End Function

Private Function strQueryAssuntos()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrCatalogoAssunto & " "
    strSql = strSql & " WHERE dtmDtCancelamento IS NULL"
    strSql = strSql & " ORDER BY strDescricao"
    strQueryAssuntos = strSql
End Function

Private Function strQueryComposicao() As String
Dim strSql As String
    
    strSql = "SELECT Pkid,"
    strSql = strSql & " strDescricao Descricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita
    strSql = strSql & " WHERE intUtilizacao = " & TYP_PRECO_PUBLICO
    strSql = strSql & " ORDER BY strDescricao"
    
    strQueryComposicao = strSql

End Function

Private Function strQueryContribuintes()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNome "
    strSql = strSql & "FROM " & gstrContribuinte & " "
    strSql = strSql & "ORDER BY strNome"
    strQueryContribuintes = strSql
End Function

Private Function strQueryReceitas()
Dim strSql As String
    strSql = ""
'    strSql = strSql & "SELECT R.PKId, Ltrim(Rtrim(R.strDescricao)) as strDescricao "
'    strSql = strSql & "FROM " & gstrReceita & " R, " & gstrReceitasExercicio & " RE "
'    strSql = strSql & "WHERE RE.intReceita = R.Pkid AND RE.intExercicio = " & Year(gstrDataDoSistema) & " AND NOT RE.dblPrecoPublico IS NULL "
    
    strSql = strSql & "select distinct "
       strSql = strSql & "tr.PKId, "
       strSql = strSql & "tr.strDescricao "
    strSql = strSql & "from "
       strSql = strSql & gstrValorCompoRec & " tv, "
       strSql = strSql & gstrComposicaoDaReceita & " tc, "
       strSql = strSql & gstrReceita & " tr, "
       strSql = strSql & gstrReceitasExercicio & " te "
    strSql = strSql & "Where "
       strSql = strSql & "tc.Pkid = tv.intComposicaoDaReceita "
       strSql = strSql & "and tr.PKId = tv.intReceita "
       strSql = strSql & "and te.intReceita = tv.intReceita "
       strSql = strSql & "and te.intExercicio = " & Year(gstrDataDoSistema) & " "
       strSql = strSql & "and not te.dblPrecoPublico is null "

    If dbcintComposicao.MatchedWithList Then
        strSql = strSql & "and tc.PKId = " & dbcintComposicao.BoundText
    Else
        strSql = strSql & "and tc.PKId = 0 "
    End If
    
    strSql = strSql & " ORDER BY tr.strDescricao"
    strQueryReceitas = strSql
End Function

Private Function GeraPrecoPublico() As Boolean
Dim adoRec                As ADODB.Recordset
Dim strSql                As String
Dim intFor                As Integer

Dim strNumeroGuia         As String
Dim strInscricao          As String
Dim strNumeroAviso        As String
Dim intNumeroDaGuia       As Long
Dim lngComposicao         As Long
Dim strComposicao         As String

Dim lngContaBancaria      As Long

Dim vetParcelas(4, 0)     As String

'Variáveis para recebimento de Pkid das Sequences
Dim intLancamentoAlfa     As Long
Dim intLancamentoValor    As Long
Dim intLancamentoPPublico As Long

Dim strsigla              As String

Dim STRMUNICIPIO          As String
Dim STRUF                 As String
Dim lngMoedaAtual         As Long

Dim ADOTemp               As ADODB.Recordset
    
    GeraPrecoPublico = False
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
'    'Vamos obter a composicao do tipo Preco Publico
'    strSql = "Select " & gstrTOPnSQLServer(1) & " Pkid, Strdescricao From " & gstrComposicaoDaReceita & " Where intutilizacao = " & TYP_PRECO_PUBLICO & " And bytdividaativa = " & bytDividaAtiva & " Order By pkid asc"
'    strSql = gstrTOPnOracle(strSql, 1)
'
'    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
'
'        With adoRec
'            If Not (.BOF And .EOF) Then
'                lngComposicao = adoRec("Pkid").Value
'                strComposicao = gstrENulo(adoRec("Strdescricao").Value)
'            Else
'                ExibeMensagem "Não existem Composições de Receita para o tipo Preço Público."
'                gobjBanco.ExecutaRollbackTrans
'                Exit Function
'            End If
'        End With
'
'    End If
    lngComposicao = dbcintComposicao.BoundText
    strComposicao = dbcintComposicao.Text
    
    'Vamos obter a conta bancaria da composicao
    strSql = "Select PA.intContaBancaria From " & gstrParametroAtualizacao & " PA Where PA.intComposicaoReceita  = " & lngComposicao & " And PA.intExercicio = " & Year(gstrDataDoSistema)
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    
        With adoRec
            If Not (.BOF And .EOF) Then
                lngContaBancaria = IIf(IsNull(adoRec("intContaBancaria").Value), 0, adoRec("intContaBancaria").Value)
            End If
        End With
        
    End If

    If gobjBanco.CriaADO("SELECT E.intMoeda, M.strDescricao Cidade, u.strSigla UF from tblempresa E, tblmunicipio M, tblUf u Where M.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " E.intcidade and U.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " E.intUf", 30, ADOTemp) Then
        If Not ADOTemp.EOF Then
            STRMUNICIPIO = ADOTemp("Cidade")
            STRUF = ADOTemp("UF")
            lngMoedaAtual = ADOTemp("intMoeda")
        Else
            STRMUNICIPIO = ""
            STRUF = ""
            lngMoedaAtual = Null
        End If
    End If
    
    'Vamos obter o numero da guia
    strNumeroGuia = glngRetornaProximoNumeroGuia
    If Val(strNumeroGuia) = 0 Then
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    strNumeroAviso = Format$(strNumeroGuia, "000000")
    intNumeroDaGuia = CLng(gstrENulo(strNumeroGuia))
        
    If Trim(cbointTipoIdentificacao.Text) = "" Then
        strInscricao = strNumeroGuia
    ElseIf gstrItemData(cbointTipoIdentificacao) = 0 Then
        strInscricao = strNumeroGuia
    Else
        If dbcintIdentificacao.MatchedWithList = True And dbcintIdentificacao.BoundText > 0 Then
            strInscricao = gstrFormataInscricao(dbcintIdentificacao.Text, CInt(gstrItemData(cbointTipoIdentificacao)))
        Else
            ExibeMensagem "Não foi possível obter o número da inscrição."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    End If
        
    '**************** LANCAMENTO ALFA ************************
        
    'Vamos gravar a tabela LancamentoAlfa
    strSql = "INSERT INTO " & gstrLancamentoAlfa & "(Strinscricao, Strcomposicaodareceita, Strocorrencia, Strnomeproprietario, Strcnpjcpf," & _
                                        "Stridentidade, Strlogradouro, Strnumero, Strcomplemento, Strbairro, Strmunicipio," & _
                                        "Struf, Intcep, Strlogradouroc, Strnumeroc, Strcomplementoc, Strbairroc, Strmunicipioc," & _
                                        "Strufc, Intcepc, Strnumeroaviso, Stremissao, Intexercicio, Intcomposicaodareceita, intUtilizacao, dtmDtAtualizacao, lngCodUsr, bytnaoinscreveda)"
    'Caso seja Imobiliario ou Economico
    If cbointTipoIdentificacao.ListIndex > 0 Then
        'Feita esta verificacao, devido problemas de join
        If bytDBType = EDatabases.Oracle Then
            strSql = strSql & "(SELECT " & IIf(dbcintIdentificacao.MatchedWithList, dbcintIdentificacao.Text, strNumeroGuia & Year(gstrDataDoSistema)) & ", CR.STRDESCRICAO strComposicao, OC.STRDESCRICAO strOcorrencia, CO.Strnome strNomeProprietario," & _
                                    "CO.Strcnpjcpf, CO.Stridentidade,  Ltrim(Rtrim(TPL.STRSIGLA)) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(TTL.STRdescricao)) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(LO.STRDESCRICAO)) strLogradouro," & _
                                    "X.INTNUMERO, X.Strcomplemento, BA.STRDESCRICAO strBairro, '" & STRMUNICIPIO & "' strMunicipio, '" & STRUF & "' strUf, X.Intcep," & _
                                    "Ltrim(Rtrim(TPL.STRSIGLA)) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(TTL.STRdescricao)) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(LO.STRDESCRICAO)) strLogradouroC," & _
                                    "X.INTNUMERO, X.Strcomplemento, BA.STRDESCRICAO strBairroC, '" & STRMUNICIPIO & "' strMunicipioC, '" & STRUF & "' strUfC, X.Intcep," & _
                                    strNumeroAviso & "," & strSUBSTRING & "(X.strEmissao,1,3)," & Year(gstrDataDoSistema) & "," & lngComposicao & ", " & cbointTipoIdentificacao.ListIndex & ", " & strGETDATE & "," & glngCodUsr & "," & IIf(bytDividaAtiva = 0, 1, 0) & _
                                " FROM " & IIf(cbointTipoIdentificacao.ListIndex = 1, gstrImobiliario, gstrEconomico) & " X," & gstrOcorrencia & " OC," & gstrComposicaoDaReceita & " CR," & gstrContribuinte & " CO," & gstrLogradouro & " LO," & _
                                    gstrTipoLogradouro & " TPL," & gstrTituloLogradouro & " TTL," & gstrBairro & " BA " & _
                                " WHERE OC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & IIf(cbointTipoIdentificacao.ListIndex = 1, "X.intOcorrrencia", "X.intOcorrencia") & " AND " & _
                                    "CO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " X.INTCONTRIBUINTE AND " & _
                                    "LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " X.Intlogradouro AND " & _
                                    "TPL.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.INTTIPOLOGRADOURO AND " & _
                                    "TTL.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.INTTITULOLOGRADOURO AND " & _
                                    "BA.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " X.Intbairro AND " & _
                                    "CR.pkid = " & lngComposicao & " AND " & _
                                    "X.PKID = " & dbcintIdentificacao.BoundText & ")"
        Else
            strSql = strSql & "(SELECT " & IIf(dbcintIdentificacao.MatchedWithList, dbcintIdentificacao.Text, strNumeroGuia & Year(gstrDataDoSistema)) & ", CR.STRDESCRICAO strComposicao, OC.STRDESCRICAO strOcorrencia, CO.Strnome strNomeProprietario," & _
                                    "CO.Strcnpjcpf, CO.Stridentidade, Ltrim(Rtrim(" & gstrISNULL("TPL.STRSIGLA", "''") & ")) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(" & gstrISNULL("TTL.STRdescricao", "''") & ")) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(" & gstrISNULL("LO.STRDESCRICAO", "''") & ")) strLogradouro," & _
                                    "X.INTNUMERO, X.Strcomplemento, BA.STRDESCRICAO strBairro, '" & STRMUNICIPIO & "' strMunicipio, '" & STRUF & "' strUf, X.Intcep, " & _
                                    "Ltrim(Rtrim(" & gstrISNULL("TPL.STRSIGLA", "''") & ")) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(" & gstrISNULL("TTL.STRdescricao", "''") & ")) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(" & gstrISNULL("LO.STRDESCRICAO", "''") & ")) strLogradouroC," & _
                                    "X.INTNUMERO, X.Strcomplemento, BA.STRDESCRICAO strBairroC, '" & STRMUNICIPIO & "' strMunicipioC, '" & STRUF & "' strUfC, X.Intcep," & _
                                    strNumeroAviso & "," & strSUBSTRING & "(X.strEmissao,1,3)," & Year(gstrDataDoSistema) & "," & lngComposicao & ", " & cbointTipoIdentificacao.ListIndex & ", " & strGETDATE & "," & glngCodUsr & "," & IIf(bytDividaAtiva = 0, 1, 0) & _
                               " FROM " & IIf(cbointTipoIdentificacao.ListIndex = 1, gstrImobiliario, gstrEconomico) & " X LEFT OUTER JOIN " & _
                                    gstrOcorrencia & " OC ON " & IIf(cbointTipoIdentificacao.ListIndex = 1, "X.intOcorrrencia", "X.intOcorrencia") & " = OC.PKId LEFT OUTER JOIN " & _
                                    gstrContribuinte & " CO ON X.intContribuinte = CO.PKId LEFT OUTER JOIN " & _
                                    gstrLogradouro & " LO ON X.intLogradouro = LO.PKId LEFT OUTER JOIN " & _
                                    gstrTipoLogradouro & " TPL ON LO.intTipoLogradouro = TPL.PKId LEFT OUTER JOIN " & _
                                    gstrTituloLogradouro & " TTL ON LO.intTituloLogradouro = TTL.PKId LEFT OUTER JOIN " & _
                                    gstrBairro & " BA ON X.intBairro = BA.PKId CROSS JOIN " & _
                                    gstrComposicaoDaReceita & " CR " & _
                                " WHERE CR.pkid = " & lngComposicao & " AND " & _
                                    "X.PKID = " & dbcintIdentificacao.BoundText & ")"
        End If
    Else
        'Feita esta verificacao, devido problemas de join
        If bytDBType = EDatabases.Oracle Then
            strSql = strSql & "(SELECT " & IIf(dbcintIdentificacao.MatchedWithList, dbcintIdentificacao.Text, strNumeroGuia & Year(gstrDataDoSistema)) & ", CR.STRDESCRICAO strComposicao, '', CO.Strnome strNomeProprietario," & _
                                    "CO.Strcnpjcpf, CO.Stridentidade,  Ltrim(Rtrim(TPL.STRSIGLA)) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(TTL.STRdescricao)) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(LO.STRDESCRICAO)) strLogradouro," & _
                                    "CO.INTNUMERO, Ltrim(Rtrim(CO.Strcomplemento)) as Strcomplemento, BA.STRDESCRICAO strBairro, '" & STRMUNICIPIO & "' strMunicipio, '" & STRUF & "' strUf, CO.Intcep," & _
                                    "CO.strLogradouroC strLogradouroC," & _
                                    "CO.INTNUMEROC, CO.StrcomplementoC, CO.strBairroC, '" & STRMUNICIPIO & "' strMunicipioC, '" & STRUF & "' strUfC, CO.IntcepC," & _
                                    strNumeroAviso & ",'000'," & Year(gstrDataDoSistema) & "," & lngComposicao & ", CR.intUtilizacao," & strGETDATE & "," & glngCodUsr & "," & IIf(bytDividaAtiva = 0, 1, 0) & _
                                " FROM " & gstrComposicaoDaReceita & " CR," & gstrContribuinte & " CO," & gstrLogradouro & " LO," & _
                                    gstrTipoLogradouro & " TPL," & gstrTituloLogradouro & " TTL," & gstrBairro & " BA " & _
                                " WHERE CO.Pkid  = " & dbcintContribuinte.BoundText & " AND " & _
                                    "LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " CO.Intlogradouro AND " & _
                                    "TPL.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.INTTIPOLOGRADOURO AND " & _
                                    "TTL.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LO.INTTITULOLOGRADOURO AND " & _
                                    "BA.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " CO.Intbairro AND " & _
                                    "CR.pkid = " & lngComposicao & ")"
        Else
            strSql = strSql & "(SELECT " & IIf(dbcintIdentificacao.MatchedWithList, dbcintIdentificacao.Text, strNumeroGuia & Year(gstrDataDoSistema)) & ", CR.STRDESCRICAO strComposicao, '', CO.Strnome strNomeProprietario," & _
                                   "CO.Strcnpjcpf, CO.Stridentidade, Ltrim(Rtrim(" & gstrISNULL("TPL.STRSIGLA", "''") & ")) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(" & gstrISNULL("TTL.STRdescricao", "''") & ")) " & strCONCAT & "' '" & strCONCAT & " Ltrim(Rtrim(" & gstrISNULL("LO.STRDESCRICAO", "''") & ")) strLogradouro," & _
                                   "CO.INTNUMERO, Ltrim(Rtrim(CO.Strcomplemento)) as Strcomplemento, BA.STRDESCRICAO strBairro, '" & STRMUNICIPIO & "' strMunicipio, '" & STRUF & "' strUf, CO.Intcep," & _
                                   "CO.strLogradouroC strLogradouroC," & _
                                   "CO.INTNUMEROC, CO.StrcomplementoC, CO.strBairroC, '" & STRMUNICIPIO & "' strMunicipioC, '" & STRUF & "' strUfC, CO.IntcepC," & _
                                   strNumeroAviso & ",'000'," & Year(gstrDataDoSistema) & "," & lngComposicao & ", CR.intUtilizacao," & strGETDATE & "," & glngCodUsr & "," & IIf(bytDividaAtiva = 0, 1, 0) & _
                            " FROM " & gstrTituloLogradouro & " TTL RIGHT OUTER JOIN " & _
                                   gstrLogradouro & " LO ON TTL.PKId = LO.intTituloLogradouro LEFT OUTER JOIN " & _
                                   gstrTipoLogradouro & " TPL ON LO.intTipoLogradouro = TPL.PKId RIGHT OUTER JOIN " & _
                                   gstrContribuinte & " CO LEFT OUTER JOIN " & _
                                   gstrBairro & " BA ON CO.intBairro = BA.PKId ON LO.PKId = CO.intLogradouro CROSS JOIN " & _
                                   gstrComposicaoDaReceita & " CR " & _
                            " WHERE CO.Pkid  = " & dbcintContribuinte.BoundText & " AND " & _
                                   "CR.pkid = " & lngComposicao & ")"
        End If
    End If
    
    If Not gobjBanco.Execute(strSql, False) Then
        ExibeMensagem "Não foi possível criar o Lançamento Alfa. A operação foi cancelada."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    intLancamentoAlfa = glngRetornaPkidTabelaPai("seqtblLancamentoAlfa", gstrLancamentoAlfa)
    
    '**************** LANCAMENTO VALOR ************************
    
    'Vamos gravar a parcela na tabela tblLancamentoValor
    strSql = "INSERT INTO " & gstrLancamentoValor & " " & _
             "(intLancamentoAlfa, intParcela, dtmDtVencimento, dblValor, intMoeda, bitParcelaValida, dtmDtAtualizacao, lngCodUsr)" & _
             " VALUES " & _
             "(" & intLancamentoAlfa & ", 1," & gstrConvDtParaSql(txtdtmDataVecto.Text) & "," & gstrConvVrParaSql(lbldblTotal.Caption) & "," & lngMoedaAtual & ",1," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                              
    If Not gobjBanco.Execute(strSql, False) Then
        gobjBanco.ExecutaRollbackTrans
        ExibeMensagem "Não foi possível criar o Lançamento Valor. A operação foi cancelada."
        Exit Function
    End If
        
    intLancamentoValor = glngRetornaPkidTabelaPai("seqtblLancamentoValor", gstrLancamentoValor)
    
    'Vamos carregar o vetor de parcelas para a emissao da guia
    vetParcelas(0, 0) = intLancamentoValor
    vetParcelas(1, 0) = lbldblTotal.Caption
    vetParcelas(2, 0) = "0,00"
    vetParcelas(3, 0) = "0,00"
    vetParcelas(4, 0) = "0,00"
        
    '**************** LANCAMENTO PRECO PUBLICO ************************
        
    'Vamos gravar na tabela tblLancamentoPPublico
    strSql = "INSERT INTO " & gstrLancamentoPPublico & " " & _
             "(intLancamentoAlfa, strCodigo, bitDigito, intExercicio, strHistorico, strAssunto, strIdCContaBancaria, dblValor, dtmDtVencimento, dtmDtAtualizacao, intutilizacao, strinscricao,lngCodUsr)" & _
             " VALUES " & _
             "(" & intLancamentoAlfa & ",'" & txtstrCodigo & "' ," & gstrENulo(txtbitDigito, , True) & "," & gstrENulo(txtintExercicio, , True) & ",'" & txtstrHistoricoPadrao.Text & "','" & dbcintAssunto.Text & "','0000'," & gstrConvVrParaSql(lbldblTotal.Caption) & ", " & gstrConvDtParaSql(txtdtmDataVecto.Text) & ", " & gstrConvDtParaSql(gstrDataDoSistema) & "," & IIf(Trim(cbointTipoIdentificacao.Text) = "", 0, gstrItemData(cbointTipoIdentificacao)) & ",'" & IIf(dbcintIdentificacao.MatchedWithList, Trim(dbcintIdentificacao.Text), "") & "'," & glngCodUsr & ")"
                              
    If Not gobjBanco.Execute(strSql, False) Then
        ExibeMensagem "Não foi possível criar o Lançamento Preço Público. A operação foi cancelada."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    intLancamentoPPublico = glngRetornaPkidTabelaPai("seqtblLancamentoPPublico", gstrLancamentoPPublico)
        
    '**************** LANCAMENTO RECEITA E PPUBLICO RECEITA ************************
        
    'Vamos gravar as receitas na tabela tblLancamentoReceita
    For intFor = 0 To xadbReceitas.UpperBound(1)
    
        strSql = "INSERT INTO " & gstrLancamentoReceita & " " & _
                 "(intLancamentoValor, intReceita, dblValor, dtmDtAtualizacao, lngCodUsr)" & _
                 " VALUES " & _
                "(" & intLancamentoValor & "," & xadbReceitas(intFor, 0) & "," & gstrConvVrParaSql(xadbReceitas(intFor, 7)) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
    
        If Not gobjBanco.Execute(strSql, False) Then
            ExibeMensagem "Não foi possível criar o Lançamento Receita. A operação foi cancelada."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    
        strSql = "INSERT INTO " & gstrLancamentoPPublicoReceita & " " & _
                 "(intLancamentoPPublico, intQtdeReceita, dblQtdeIndexador, dblValorIndexador, strIndexador, dblValor, strReceita, dtmDtAtualizacao, lngCodUsr)" & _
                 " VALUES " & _
                "(" & intLancamentoPPublico & "," & gstrConvVrParaSql(xadbReceitas(intFor, 6)) & "," & gstrConvVrParaSql(xadbReceitas(intFor, 2)) & "," & gstrConvVrParaSql(xadbReceitas(intFor, 4)) & ",'" & xadbReceitas(intFor, 3) & "'," & gstrConvVrParaSql(xadbReceitas(intFor, 7)) & ",'" & xadbReceitas(intFor, 1) & "'," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
    
        If Not gobjBanco.Execute(strSql, False) Then
            ExibeMensagem "Não foi possível criar o Lançamento Preço Publico Receita. A operação foi cancelada."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    
    Next
    
    gobjBanco.ExecutaCommitTrans
    
    For intFor = 0 To xadbReceitas.UpperBound(1)
        strsigla = strsigla & xadbReceitas(intFor, 9) & " / "
    Next
    strsigla = Mid(strsigla, 1, Len(strsigla) - 2)
     
    'Vammos imprimir a Guia de Preco Publico
    ImprimePrecoPublico strInscricao, intNumeroDaGuia, dbcintContribuinte.Text, Trim(txtstrLogradouro) & ", " & Trim(txtintNumero) & " - " & Trim(txtstrComplemento), Trim(txtstrBairro), txtstrMunicipio, txtstrUF, gstrValorSemMascara(txtintCep), "", "", strNumeroAviso, IIf(Len(Trim(txtstrCodigo)) > 0, txtstrCodigo & "/" & txtintExercicio & "-" & txtbitDigito, ""), strsigla, Trim(txtstrHistoricoPadrao), lngContaBancaria, gstrConvVrDoSql(lbldblTotal.Caption), 0, 0, 0, gstrConvVrDoSql(lbldblTotal.Caption), txtdtmDataVecto.Text, vetParcelas, lngContaBancaria = 0
    
    GeraPrecoPublico = True
    
End Function

Private Function blnVerificaReceitas() As Boolean
    Dim intFor As Integer
    
    blnVerificaReceitas = False

    For intFor = 0 To xadbReceitas.UpperBound(1)
        If Val(xadbReceitas(intFor, 0)) > 0 Then
            blnVerificaReceitas = True
            Exit Function
        End If
    Next
    
End Function
