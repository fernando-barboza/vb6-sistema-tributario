VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmAtualizacaoDebitos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualização de Débitos"
   ClientHeight    =   8355
   ClientLeft      =   465
   ClientTop       =   2625
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11925
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   8250
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   14552
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Receita "
      TabPicture(0)   =   "frmAtualizacaoDebitos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTributo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrInscricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintContribuinte"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "shp_Opcionais"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_Opcionais"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "shp_Totais"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_dblTotOriginal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_dblTotPrincipal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_dblTotMulta"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_dblTotJuros"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl_dblTotCorrecao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl_dblTotTotal"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl_dblTotHonorarios"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl_dblTotGeral"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbldblCredito"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblintUtilizacao"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dbcintUtilizacao"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "tdbValoresParcelasOpcionais"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "tdbValoresParcelas"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "tdbValoresAcumulado"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "dbcintComposicaoDaReceita"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdComposicao"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkTodasReceitas"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtstrInscricao"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtstrContribuinte"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtdblCredito"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkTodasUtilizacoes"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      Begin VB.CheckBox chkTodasUtilizacoes 
         Caption         =   "Selecionar todas as utilizações"
         Height          =   225
         Left            =   7680
         TabIndex        =   25
         Top             =   1440
         Width           =   2475
      End
      Begin VB.TextBox txtdblCredito 
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
         Left            =   1950
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1020
         Width           =   2280
      End
      Begin VB.TextBox txtstrContribuinte 
         Height          =   285
         Left            =   7695
         TabIndex        =   8
         Top             =   750
         Width           =   4005
      End
      Begin VB.TextBox txtstrInscricao 
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
         Left            =   8190
         MaxLength       =   20
         TabIndex        =   6
         Top             =   420
         Width           =   3510
      End
      Begin VB.CheckBox chkTodasReceitas 
         Caption         =   "Selecionar todas as receitas"
         Height          =   225
         Left            =   1950
         TabIndex        =   4
         Top             =   750
         Width           =   2475
      End
      Begin VB.CommandButton cmdComposicao 
         Height          =   300
         Left            =   5670
         Picture         =   "frmAtualizacaoDebitos.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Composição da Receita"
         Top             =   405
         Width           =   360
      End
      Begin MSDataListLib.DataCombo dbcintComposicaoDaReceita 
         Height          =   315
         HelpContextID   =   1
         Left            =   1950
         TabIndex        =   2
         Top             =   405
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdbValoresAcumulado 
         Height          =   1440
         Left            =   120
         TabIndex        =   9
         Top             =   1740
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   2540
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pkid"
         Columns(0).DataField=   "PkidLA"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   68
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Inscr. cadastral"
         Columns(2).DataField=   "strInsricao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "intComposicaoDaReceita"
         Columns(3).DataField=   "intComposicaoDaReceita"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Composição da receita"
         Columns(4).DataField=   "strComposicaoDaReceita"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Exercício"
         Columns(5).DataField=   "intExercicio"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Valor original"
         Columns(6).DataField=   "dblValorOriginal"
         Columns(6).NumberFormat=   "FormatText Event"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Principal"
         Columns(7).DataField=   "dblPrincipal"
         Columns(7).NumberFormat=   "FormatText Event"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Multa"
         Columns(8).DataField=   "dblMulta"
         Columns(8).NumberFormat=   "FormatText Event"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Juros"
         Columns(9).DataField=   "dblJuros"
         Columns(9).NumberFormat=   "FormatText Event"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Correção"
         Columns(10).DataField=   "dblCorrecao"
         Columns(10).NumberFormat=   "FormatText Event"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Total"
         Columns(11).DataField=   "dblTotal"
         Columns(11).NumberFormat=   "FormatText Event"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Honorários"
         Columns(12).DataField=   "dblHonorarios"
         Columns(12).NumberFormat=   "FormatText Event"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Dirigências"
         Columns(13).DataField=   "dblDirigencias"
         Columns(13).NumberFormat=   "FormatText Event"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Custas"
         Columns(14).DataField=   "dblCustas"
         Columns(14).NumberFormat=   "FormatText Event"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Total Geral"
         Columns(15).DataField=   "dblTotalGeral"
         Columns(15).NumberFormat=   "FormatText Event"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   16
         Columns(16)._MaxComboItems=   5
         Columns(16).ValueItems(0)._DefaultItem=   0
         Columns(16).ValueItems(0).Value=   "true"
         Columns(16).ValueItems(0).Value.vt=   8
         Columns(16).ValueItems(0).DisplayValue=   "Sim"
         Columns(16).ValueItems(0).DisplayValue.vt=   8
         Columns(16).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(16).ValueItems(1)._DefaultItem=   0
         Columns(16).ValueItems(1).Value=   "false"
         Columns(16).ValueItems(1).Value.vt=   8
         Columns(16).ValueItems(1).DisplayValue=   "Não"
         Columns(16).ValueItems(1).DisplayValue.vt=   8
         Columns(16).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(16).ValueItems.Count=   2
         Columns(16).Caption=   "Acordo"
         Columns(16).DataField=   "intLancamentoAlfaAcordo"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Executivo"
         Columns(17).DataField=   "intLancamentoAlfaExecutivo"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "QtdeParcelasInicial"
         Columns(18).DataField=   ""
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "Vencto1Parcela"
         Columns(19).DataField=   ""
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "ValorParcelasInicial"
         Columns(20).DataField=   "dblTotal"
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   0
         Columns(21)._MaxComboItems=   5
         Columns(21).Caption=   "AcordosRelacionados"
         Columns(21).DataField=   ""
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(22)._VlistStyle=   0
         Columns(22)._MaxComboItems=   5
         Columns(22).Caption=   "nº Aviso"
         Columns(22).DataField=   "strNumeroAviso"
         Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(23)._VlistStyle=   0
         Columns(23)._MaxComboItems=   5
         Columns(23).Caption=   "Conta"
         Columns(23).DataField=   "intContaBancaria"
         Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   24
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=24"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=476"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=397"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2170"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2090"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=3625"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3545"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(27)=   "Column(4).Width=3043"
         Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=2963"
         Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=1323"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1244"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=1905"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1826"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(7).Width=1720"
         Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1640"
         Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(49)=   "Column(8).Width=1640"
         Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1561"
         Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=2"
         Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(55)=   "Column(9).Width=1402"
         Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1323"
         Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=2"
         Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(61)=   "Column(10).Width=1508"
         Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=1429"
         Splits(0)._ColumnProps(64)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(10)._ColStyle=2"
         Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(67)=   "Column(11).Width=1693"
         Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=1614"
         Splits(0)._ColumnProps(70)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(11)._ColStyle=2"
         Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(73)=   "Column(12).Width=1588"
         Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=1508"
         Splits(0)._ColumnProps(76)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(77)=   "Column(12)._ColStyle=2"
         Splits(0)._ColumnProps(78)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(79)=   "Column(13).Width=1693"
         Splits(0)._ColumnProps(80)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(13)._WidthInPix=1614"
         Splits(0)._ColumnProps(82)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(83)=   "Column(13).AllowSizing=0"
         Splits(0)._ColumnProps(84)=   "Column(13)._ColStyle=2"
         Splits(0)._ColumnProps(85)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(86)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(87)=   "Column(14).Width=1535"
         Splits(0)._ColumnProps(88)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(89)=   "Column(14)._WidthInPix=1455"
         Splits(0)._ColumnProps(90)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(91)=   "Column(14).AllowSizing=0"
         Splits(0)._ColumnProps(92)=   "Column(14)._ColStyle=2"
         Splits(0)._ColumnProps(93)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(94)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(95)=   "Column(15).Width=1588"
         Splits(0)._ColumnProps(96)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(97)=   "Column(15)._WidthInPix=1508"
         Splits(0)._ColumnProps(98)=   "Column(15)._EditAlways=0"
         Splits(0)._ColumnProps(99)=   "Column(15)._ColStyle=2"
         Splits(0)._ColumnProps(100)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(101)=   "Column(16).Width=1799"
         Splits(0)._ColumnProps(102)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(103)=   "Column(16)._WidthInPix=1720"
         Splits(0)._ColumnProps(104)=   "Column(16)._EditAlways=0"
         Splits(0)._ColumnProps(105)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(106)=   "Column(17).Width=1429"
         Splits(0)._ColumnProps(107)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(108)=   "Column(17)._WidthInPix=1349"
         Splits(0)._ColumnProps(109)=   "Column(17)._EditAlways=0"
         Splits(0)._ColumnProps(110)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(111)=   "Column(18).Width=2725"
         Splits(0)._ColumnProps(112)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(113)=   "Column(18)._WidthInPix=2646"
         Splits(0)._ColumnProps(114)=   "Column(18)._EditAlways=0"
         Splits(0)._ColumnProps(115)=   "Column(18).AllowSizing=0"
         Splits(0)._ColumnProps(116)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(117)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(118)=   "Column(19).Width=2725"
         Splits(0)._ColumnProps(119)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(120)=   "Column(19)._WidthInPix=2646"
         Splits(0)._ColumnProps(121)=   "Column(19)._EditAlways=0"
         Splits(0)._ColumnProps(122)=   "Column(19).AllowSizing=0"
         Splits(0)._ColumnProps(123)=   "Column(19).Visible=0"
         Splits(0)._ColumnProps(124)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(125)=   "Column(20).Width=2725"
         Splits(0)._ColumnProps(126)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(127)=   "Column(20)._WidthInPix=2646"
         Splits(0)._ColumnProps(128)=   "Column(20)._EditAlways=0"
         Splits(0)._ColumnProps(129)=   "Column(20).AllowSizing=0"
         Splits(0)._ColumnProps(130)=   "Column(20).Visible=0"
         Splits(0)._ColumnProps(131)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(132)=   "Column(21).Width=2725"
         Splits(0)._ColumnProps(133)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(134)=   "Column(21)._WidthInPix=2646"
         Splits(0)._ColumnProps(135)=   "Column(21)._EditAlways=0"
         Splits(0)._ColumnProps(136)=   "Column(21).AllowSizing=0"
         Splits(0)._ColumnProps(137)=   "Column(21).Visible=0"
         Splits(0)._ColumnProps(138)=   "Column(21).Order=22"
         Splits(0)._ColumnProps(139)=   "Column(22).Width=1535"
         Splits(0)._ColumnProps(140)=   "Column(22).DividerColor=0"
         Splits(0)._ColumnProps(141)=   "Column(22)._WidthInPix=1455"
         Splits(0)._ColumnProps(142)=   "Column(22)._EditAlways=0"
         Splits(0)._ColumnProps(143)=   "Column(22)._ColStyle=2"
         Splits(0)._ColumnProps(144)=   "Column(22).Order=23"
         Splits(0)._ColumnProps(145)=   "Column(23).Width=2725"
         Splits(0)._ColumnProps(146)=   "Column(23).DividerColor=0"
         Splits(0)._ColumnProps(147)=   "Column(23)._WidthInPix=2646"
         Splits(0)._ColumnProps(148)=   "Column(23)._EditAlways=0"
         Splits(0)._ColumnProps(149)=   "Column(23).Visible=0"
         Splits(0)._ColumnProps(150)=   "Column(23).Order=24"
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
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
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=90,.parent=13,.alignment=1"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=1"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=118,.parent=13,.alignment=1"
         _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=115,.parent=14"
         _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=116,.parent=15"
         _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=117,.parent=17"
         _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=122,.parent=13,.alignment=1"
         _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=119,.parent=14"
         _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=120,.parent=15"
         _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=121,.parent=17"
         _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=126,.parent=13,.alignment=1"
         _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=123,.parent=14"
         _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=124,.parent=15"
         _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=125,.parent=17"
         _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=130,.parent=13,.alignment=1"
         _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=127,.parent=14"
         _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=128,.parent=15"
         _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=129,.parent=17"
         _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=82,.parent=13"
         _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=79,.parent=14"
         _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=80,.parent=15"
         _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=81,.parent=17"
         _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=86,.parent=13"
         _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=14"
         _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=15"
         _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=17"
         _StyleDefs(109) =   "Splits(0).Columns(18).Style:id=94,.parent=13"
         _StyleDefs(110) =   "Splits(0).Columns(18).HeadingStyle:id=91,.parent=14"
         _StyleDefs(111) =   "Splits(0).Columns(18).FooterStyle:id=92,.parent=15"
         _StyleDefs(112) =   "Splits(0).Columns(18).EditorStyle:id=93,.parent=17"
         _StyleDefs(113) =   "Splits(0).Columns(19).Style:id=102,.parent=13"
         _StyleDefs(114) =   "Splits(0).Columns(19).HeadingStyle:id=99,.parent=14"
         _StyleDefs(115) =   "Splits(0).Columns(19).FooterStyle:id=100,.parent=15"
         _StyleDefs(116) =   "Splits(0).Columns(19).EditorStyle:id=101,.parent=17"
         _StyleDefs(117) =   "Splits(0).Columns(20).Style:id=106,.parent=13"
         _StyleDefs(118) =   "Splits(0).Columns(20).HeadingStyle:id=103,.parent=14"
         _StyleDefs(119) =   "Splits(0).Columns(20).FooterStyle:id=104,.parent=15"
         _StyleDefs(120) =   "Splits(0).Columns(20).EditorStyle:id=105,.parent=17"
         _StyleDefs(121) =   "Splits(0).Columns(21).Style:id=110,.parent=13"
         _StyleDefs(122) =   "Splits(0).Columns(21).HeadingStyle:id=107,.parent=14"
         _StyleDefs(123) =   "Splits(0).Columns(21).FooterStyle:id=108,.parent=15"
         _StyleDefs(124) =   "Splits(0).Columns(21).EditorStyle:id=109,.parent=17"
         _StyleDefs(125) =   "Splits(0).Columns(22).Style:id=98,.parent=13,.alignment=1"
         _StyleDefs(126) =   "Splits(0).Columns(22).HeadingStyle:id=95,.parent=14"
         _StyleDefs(127) =   "Splits(0).Columns(22).FooterStyle:id=96,.parent=15"
         _StyleDefs(128) =   "Splits(0).Columns(22).EditorStyle:id=97,.parent=17"
         _StyleDefs(129) =   "Splits(0).Columns(23).Style:id=114,.parent=13"
         _StyleDefs(130) =   "Splits(0).Columns(23).HeadingStyle:id=111,.parent=14"
         _StyleDefs(131) =   "Splits(0).Columns(23).FooterStyle:id=112,.parent=15"
         _StyleDefs(132) =   "Splits(0).Columns(23).EditorStyle:id=113,.parent=17"
         _StyleDefs(133) =   "Named:id=33:Normal"
         _StyleDefs(134) =   ":id=33,.parent=0"
         _StyleDefs(135) =   "Named:id=34:Heading"
         _StyleDefs(136) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(137) =   ":id=34,.wraptext=-1"
         _StyleDefs(138) =   "Named:id=35:Footing"
         _StyleDefs(139) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(140) =   "Named:id=36:Selected"
         _StyleDefs(141) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(142) =   "Named:id=37:Caption"
         _StyleDefs(143) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(144) =   "Named:id=38:HighlightRow"
         _StyleDefs(145) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(146) =   "Named:id=39:EvenRow"
         _StyleDefs(147) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(148) =   "Named:id=40:OddRow"
         _StyleDefs(149) =   ":id=40,.parent=33"
         _StyleDefs(150) =   "Named:id=41:RecordSelector"
         _StyleDefs(151) =   ":id=41,.parent=34"
         _StyleDefs(152) =   "Named:id=42:FilterBar"
         _StyleDefs(153) =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid tdbValoresParcelas 
         Height          =   2955
         Left            =   120
         TabIndex        =   10
         Top             =   3555
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5212
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pkid"
         Columns(0).DataField=   "PkidLV"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   68
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "intLancamentoAlfa"
         Columns(2).DataField=   "intLancamentoAlfa"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Parcela"
         Columns(3).DataField=   "intParcela"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor original"
         Columns(4).DataField=   "ValorOrig"
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Principal"
         Columns(5).DataField=   "dblValorPrincipal"
         Columns(5).NumberFormat=   "FormatText Event"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Multa"
         Columns(6).DataField=   "dblValorMulta"
         Columns(6).NumberFormat=   "FormatText Event"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Juros"
         Columns(7).DataField=   "dblValorJuros"
         Columns(7).NumberFormat=   "FormatText Event"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Correção"
         Columns(8).DataField=   "dblValorCorrecao"
         Columns(8).NumberFormat=   "FormatText Event"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Total"
         Columns(9).DataField=   "dblValorTotal"
         Columns(9).NumberFormat=   "FormatText Event"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Acordo"
         Columns(10).DataField=   "intLancamentoAlfaAcordo"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   16
         Columns(11)._MaxComboItems=   5
         Columns(11).ValueItems(0)._DefaultItem=   0
         Columns(11).ValueItems(0).Value=   "true"
         Columns(11).ValueItems(0).Value.vt=   8
         Columns(11).ValueItems(0).DisplayValue=   "Sim"
         Columns(11).ValueItems(0).DisplayValue.vt=   8
         Columns(11).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(11).ValueItems(1)._DefaultItem=   0
         Columns(11).ValueItems(1).Value=   "false"
         Columns(11).ValueItems(1).Value.vt=   8
         Columns(11).ValueItems(1).DisplayValue=   "Não"
         Columns(11).ValueItems(1).DisplayValue.vt=   8
         Columns(11).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(11).ValueItems.Count=   2
         Columns(11).Caption=   "Executivo"
         Columns(11).DataField=   "intLancamentoAlfaExecutivo"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Vencimento"
         Columns(12).DataField=   "dtmdtvencimento"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "D.A."
         Columns(13).DataField=   "intLancamentoAlfaDAtiva"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   14
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=14"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=476"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=397"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2170"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2090"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=1138"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1058"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8194"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=2037"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1958"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=2"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1693"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(38)=   "Column(6).Width=1720"
         Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1640"
         Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(44)=   "Column(7).Width=1693"
         Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1614"
         Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(50)=   "Column(8).Width=1746"
         Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1667"
         Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=2"
         Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(56)=   "Column(9).Width=2011"
         Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1931"
         Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=2"
         Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(62)=   "Column(10).Width=2302"
         Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=2223"
         Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=2"
         Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(68)=   "Column(11).Width=1905"
         Splits(0)._ColumnProps(69)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(70)=   "Column(11)._WidthInPix=1826"
         Splits(0)._ColumnProps(71)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(72)=   "Column(11)._ColStyle=2"
         Splits(0)._ColumnProps(73)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(74)=   "Column(12).Width=1640"
         Splits(0)._ColumnProps(75)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(76)=   "Column(12)._WidthInPix=1561"
         Splits(0)._ColumnProps(77)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(78)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(79)=   "Column(13).Width=979"
         Splits(0)._ColumnProps(80)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(13)._WidthInPix=900"
         Splits(0)._ColumnProps(82)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(83)=   "Column(13).Order=14"
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
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
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=90,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=1"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=82,.parent=13"
         _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
         _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
         _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
         _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=86,.parent=13"
         _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
         _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
         _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
         _StyleDefs(93)  =   "Named:id=33:Normal"
         _StyleDefs(94)  =   ":id=33,.parent=0"
         _StyleDefs(95)  =   "Named:id=34:Heading"
         _StyleDefs(96)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(97)  =   ":id=34,.wraptext=-1"
         _StyleDefs(98)  =   "Named:id=35:Footing"
         _StyleDefs(99)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(100) =   "Named:id=36:Selected"
         _StyleDefs(101) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(102) =   "Named:id=37:Caption"
         _StyleDefs(103) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(104) =   "Named:id=38:HighlightRow"
         _StyleDefs(105) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(106) =   "Named:id=39:EvenRow"
         _StyleDefs(107) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(108) =   "Named:id=40:OddRow"
         _StyleDefs(109) =   ":id=40,.parent=33"
         _StyleDefs(110) =   "Named:id=41:RecordSelector"
         _StyleDefs(111) =   ":id=41,.parent=34"
         _StyleDefs(112) =   "Named:id=42:FilterBar"
         _StyleDefs(113) =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid tdbValoresParcelasOpcionais 
         Height          =   1290
         Left            =   195
         TabIndex        =   11
         Top             =   6780
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   2275
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pkid"
         Columns(0).DataField=   "PkidLV"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   68
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "intLancamentoAlfa"
         Columns(2).DataField=   "intLancamentoAlfa"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Parcela"
         Columns(3).DataField=   "intParcela"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Valor original"
         Columns(4).DataField=   "ValorOrig"
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Principal"
         Columns(5).DataField=   "dblValorPrincipal"
         Columns(5).NumberFormat=   "FormatText Event"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Multa"
         Columns(6).DataField=   "dblValorMulta"
         Columns(6).NumberFormat=   "FormatText Event"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Juros"
         Columns(7).DataField=   "dblValorJuros"
         Columns(7).NumberFormat=   "FormatText Event"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Correção"
         Columns(8).DataField=   "dblValorCorrecao"
         Columns(8).NumberFormat=   "FormatText Event"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Total"
         Columns(9).DataField=   "dblValorTotal"
         Columns(9).NumberFormat=   "FormatText Event"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Acordo"
         Columns(10).DataField=   "intLancamentoAlfaAcordo"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   16
         Columns(11)._MaxComboItems=   5
         Columns(11).ValueItems(0)._DefaultItem=   0
         Columns(11).ValueItems(0).Value=   "true"
         Columns(11).ValueItems(0).Value.vt=   8
         Columns(11).ValueItems(0).DisplayValue=   "Sim"
         Columns(11).ValueItems(0).DisplayValue.vt=   8
         Columns(11).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(11).ValueItems(1)._DefaultItem=   0
         Columns(11).ValueItems(1).Value=   "false"
         Columns(11).ValueItems(1).Value.vt=   8
         Columns(11).ValueItems(1).DisplayValue=   "Não"
         Columns(11).ValueItems(1).DisplayValue.vt=   8
         Columns(11).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(11).ValueItems.Count=   2
         Columns(11).Caption=   "Executivo"
         Columns(11).DataField=   "intLancamentoAlfaExecutivo"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Vencimento"
         Columns(12).DataField=   "dtmdtvencimento"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "D.A."
         Columns(13).DataField=   "intLancamentoAlfaDAtiva"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   14
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=14"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=476"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=397"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2170"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2090"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=1138"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1058"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8194"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=2037"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1958"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=2"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1693"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(38)=   "Column(6).Width=1720"
         Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1640"
         Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(44)=   "Column(7).Width=1693"
         Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1614"
         Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(50)=   "Column(8).Width=1746"
         Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1667"
         Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=2"
         Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(56)=   "Column(9).Width=2011"
         Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1931"
         Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=2"
         Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(62)=   "Column(10).Width=2302"
         Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=2223"
         Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=2"
         Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(68)=   "Column(11).Width=1773"
         Splits(0)._ColumnProps(69)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(70)=   "Column(11)._WidthInPix=1693"
         Splits(0)._ColumnProps(71)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(73)=   "Column(12).Width=1614"
         Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=1535"
         Splits(0)._ColumnProps(76)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(77)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(78)=   "Column(13).Width=979"
         Splits(0)._ColumnProps(79)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(80)=   "Column(13)._WidthInPix=900"
         Splits(0)._ColumnProps(81)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(82)=   "Column(13).Order=14"
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
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
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=90,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=78,.parent=13"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=82,.parent=13"
         _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
         _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
         _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
         _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=86,.parent=13"
         _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
         _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
         _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
         _StyleDefs(93)  =   "Named:id=33:Normal"
         _StyleDefs(94)  =   ":id=33,.parent=0"
         _StyleDefs(95)  =   "Named:id=34:Heading"
         _StyleDefs(96)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(97)  =   ":id=34,.wraptext=-1"
         _StyleDefs(98)  =   "Named:id=35:Footing"
         _StyleDefs(99)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(100) =   "Named:id=36:Selected"
         _StyleDefs(101) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(102) =   "Named:id=37:Caption"
         _StyleDefs(103) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(104) =   "Named:id=38:HighlightRow"
         _StyleDefs(105) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(106) =   "Named:id=39:EvenRow"
         _StyleDefs(107) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(108) =   "Named:id=40:OddRow"
         _StyleDefs(109) =   ":id=40,.parent=33"
         _StyleDefs(110) =   "Named:id=41:RecordSelector"
         _StyleDefs(111) =   ":id=41,.parent=34"
         _StyleDefs(112) =   "Named:id=42:FilterBar"
         _StyleDefs(113) =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintUtilizacao 
         Height          =   315
         HelpContextID   =   1
         Left            =   7695
         TabIndex        =   23
         Top             =   1080
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblintUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   6930
         TabIndex        =   24
         Top             =   1155
         Width           =   690
      End
      Begin VB.Label lbldblCredito 
         AutoSize        =   -1  'True
         Caption         =   "Crédito"
         Height          =   195
         Left            =   1350
         TabIndex        =   21
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_dblTotGeral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   10590
         TabIndex        =   20
         Tag             =   "1"
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label lbl_dblTotHonorarios 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   9720
         TabIndex        =   19
         Tag             =   "1"
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label lbl_dblTotTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8790
         TabIndex        =   18
         Tag             =   "1"
         Top             =   3240
         Width           =   960
      End
      Begin VB.Label lbl_dblTotCorrecao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7950
         TabIndex        =   17
         Tag             =   "1"
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lbl_dblTotJuros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7125
         TabIndex        =   16
         Tag             =   "1"
         Top             =   3240
         Width           =   840
      End
      Begin VB.Label lbl_dblTotMulta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6195
         TabIndex        =   15
         Tag             =   "1"
         Top             =   3240
         Width           =   945
      End
      Begin VB.Label lbl_dblTotPrincipal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5175
         TabIndex        =   14
         Tag             =   "1"
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Label lbl_dblTotOriginal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4155
         TabIndex        =   13
         Tag             =   "1"
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Shape shp_Totais 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Left            =   120
         Top             =   3240
         Width           =   11580
      End
      Begin VB.Label lbl_Opcionais 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opcionais"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   5985
         Width           =   705
      End
      Begin VB.Shape shp_Opcionais 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   1560
         Left            =   135
         Top             =   6600
         Width           =   11580
      End
      Begin VB.Label lblintContribuinte 
         AutoSize        =   -1  'True
         Caption         =   "Contribuinte"
         Height          =   195
         Left            =   6780
         TabIndex        =   7
         Top             =   810
         Width           =   840
      End
      Begin VB.Label lblstrInscricao 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   6780
         TabIndex        =   5
         Top             =   465
         Width           =   1350
      End
      Begin VB.Label lblTributo 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAtualizacaoDebitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dtmDataAtualizacao      As Date
Dim xadbLista               As XArrayDB
Dim xadbParcelas            As XArrayDB
Dim xadbParcelasOpcionais   As XArrayDB
Dim vetParcelas()           As String
Dim vetParcelasOpcionais()  As String
Dim blnReajuste             As Boolean
Dim blnNegativa             As Boolean

Dim strLogradouro           As String
Dim STRENDERECO             As String
Dim strNumero               As String
Dim STRCOMPLEMENTO          As String
Dim STRBAIRRO               As String
Dim STRMUNICIPIO            As String
Dim STRUF                   As String
Dim strCep                  As String
Dim strEnderecoC            As String
Dim strNumeroC              As String
Dim strComplementoC         As String
Dim strBairroC              As String
Dim strMunicipioC           As String
Dim strUFC                  As String
Dim strCepC                 As String
Dim lngContribuinte         As Long
Dim dblValorAcordo          As Double
Dim VetParcelasAcordo()     As String

Dim strQuadra               As String
Dim strLote                 As String

Dim strNumeroAviso          As String
Dim strInscricaoAuxiliar    As String
Dim strIPTU                 As String
Dim intUtilizacao           As Integer

Public strProprietario      As String
Public strNumeroProcesso    As String
Public dtmVencimento        As Date

Public blnAcordoAtivo       As Boolean

'Variáveis usadas para impressão da certidão de Divida Ativa
Dim lngDividaAtiva          As Long
Dim xadbDividaAtiva         As XArrayDB
Dim blnDAtiva               As Boolean

'Variaveis que indicam a habilitacao das guias
Dim blnGuiaPositiva         As Boolean

Dim blnAcordoEmDividaAtiva  As Boolean
Dim blnAcordoCreditaParcela As Boolean
Dim blnPrimeiraVezFoco      As Boolean

Dim vetTotais()             As String

Private Sub chkTodasReceitas_Click()
    If chkTodasReceitas.Value = 1 Then
        chkTodasUtilizacoes.Value = 0
        chkTodasUtilizacoes.Enabled = False
        
        dbcintComposicaoDaReceita.BoundText = ""
        dbcintComposicaoDaReceita.Enabled = False
        TrocaCorObjeto dbcintComposicaoDaReceita, True
        
        dbcintUtilizacao.BoundText = ""
        dbcintUtilizacao.Enabled = False
        TrocaCorObjeto dbcintUtilizacao, True
    Else
        chkTodasUtilizacoes.Enabled = True
        
        dbcintComposicaoDaReceita.Enabled = True
        TrocaCorObjeto dbcintComposicaoDaReceita, False
        
        dbcintUtilizacao.Enabled = True
        TrocaCorObjeto dbcintUtilizacao, False
    End If
End Sub

Private Sub chkTodasReceitas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkTodasReceitas
End Sub

Private Sub chkTodasUtilizacoes_Click()
    If chkTodasUtilizacoes.Value = 1 Then
        chkTodasReceitas.Value = 0
        chkTodasReceitas.Enabled = False
        
        dbcintComposicaoDaReceita.BoundText = ""
        dbcintComposicaoDaReceita.Enabled = False
        TrocaCorObjeto dbcintComposicaoDaReceita, True
        
        dbcintUtilizacao.BoundText = ""
        dbcintUtilizacao.Enabled = False
        TrocaCorObjeto dbcintUtilizacao, True
    Else
        chkTodasReceitas.Enabled = True
        
        dbcintComposicaoDaReceita.Enabled = True
        TrocaCorObjeto dbcintComposicaoDaReceita, False
        
        dbcintUtilizacao.Enabled = True
        TrocaCorObjeto dbcintUtilizacao, False
    End If
End Sub

Private Sub cmdComposicao_Click()
    CarregaForm frmCadComposicaoDaReceita, dbcintComposicaoDaReceita
End Sub

Private Sub dbcintComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbcintComposicaoDaReceita, Me, Area
    If dbcintComposicaoDaReceita.MatchedWithList Then
        chkTodasUtilizacoes.Value = 0
        chkTodasUtilizacoes.Enabled = False
        
        dbcintUtilizacao.BoundText = ""
        dbcintUtilizacao.Enabled = False
        TrocaCorObjeto dbcintUtilizacao, True
    End If
End Sub

Private Sub dbcintComposicaoDaReceita_GotFocus()
    MarcaCampo dbcintComposicaoDaReceita
End Sub

Private Sub dbcintComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbcintComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintComposicaoDaReceita
End Sub

Private Sub dbcintUtilizacao_Click(Area As Integer)
    DropDownDataCombo dbcintUtilizacao, Me, Area
    If dbcintUtilizacao.MatchedWithList Then
        chkTodasReceitas.Value = 0
        chkTodasReceitas.Enabled = False
        
        dbcintComposicaoDaReceita.BoundText = ""
        dbcintComposicaoDaReceita.Enabled = False
        TrocaCorObjeto dbcintComposicaoDaReceita, True
    End If
End Sub

Private Sub dbcintUtilizacao_GotFocus()
    MarcaCampo dbcintUtilizacao
End Sub

Private Sub dbcintUtilizacao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintUtilizacao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUtilizacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintUtilizacao
End Sub

Private Sub Form_Activate()
    If blnAcordoAtivo Then
       frmCadAcordos.SetFocus
       Exit Sub
    End If
    gintCodSeguranca = 1161
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar, gstrDeletar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
    HabilitaDesabilitaBotao1 blnGuiaPositiva, gstrBtnArquivo, gstrGuiaCertidaoPositiva
    If blnReajuste Then HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimirGuia, gstrParcelamentoDebitoAtualizado
    If blnDAtiva Then HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrGuiaCertidaoDividaAtiva
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrGuiaRelacaoDeDebitos
    If blnPrimeiraVezFoco = False Then
       txtstrInscricao.SetFocus
       blnPrimeiraVezFoco = True
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste, gstrImprimirGuia, gstrParcelamentoDebitoAtualizado, gstrGuiaCertidaoPositiva, gstrGuiaCertidaoDividaAtiva
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrGuiaRelacaoDeDebitos
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub Form_Load()
    TrocaCorObjeto txtstrContribuinte, True
    TrocaCorObjeto txtdblCredito, True
    dtmDataAtualizacao = gstrDataDoSistema
    dbcintComposicaoDaReceita.Tag = strQueryComposicao & ";strDescricao"
    dbcintUtilizacao.Tag = strQueryUtilizacao & ";strDescricao"
    blnNegativa = False
    blnPrimeiraVezFoco = False
    chkTodasReceitas.Value = 1
    
    VerificaAcordoEmDividaAtiva
    
    tdbValoresAcumulado_ColResize 0, 0
    
   'Define o Style quando a data de pagamento esta preenchida
   'Dim JaPagas As New TrueOleDBGrid70.Style
   'JaPagas.BackColor = vbRed
   'tdbValoresParcelas.Columns("intParcela").AddRegexCellStyle dbgNormalCell, JaPagas, "^1"
   'tdbValoresParcelas.Columns("intParcela").AddRegexCellStyle dbgNormalCell + dbgCurrentCell, JaPagas, "^0"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimirGuia, gstrParcelamentoDebitoAtualizado, gstrGuiaCertidaoDividaAtiva
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrGuiaRelacaoDeDebitos
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrGuiaCertidaoPositiva
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrGuiaCertidaoPositivaEfeitoNegativa
    blnReajuste = False
    blnGuiaPositiva = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim intCont    As Integer
Dim intPosicao As Integer
Dim DBLVALOR   As Long
Dim blnCheck   As Boolean
    
    Select Case UCase(strModoOperacao)
        Case Is = gstrGuiaCertidaoDividaAtiva
            If blnMontaArrayDividaAtiva Then
                ImprimeRelatorioPorArray rptCertidaoDividaAtiva, , "Certidão de Dívida Ativa", , xadbDividaAtiva, True
                blnDAtiva = True
            End If
         Case Is = UCase(gstrNovo)
            LimpaObjeto Me
            LimpaGrids
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimirGuia, gstrParcelamentoDebitoAtualizado
            blnReajuste = False
            blnGuiaPositiva = False
            chkTodasReceitas.Enabled = True
            chkTodasReceitas.Value = 1
            txtstrInscricao.SetFocus
        
        Case Is = UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        
        Case Is = UCase(gstrCalcularReajuste), UCase(gstrLocalizar)
            If blnDadosOK Then
                LimpaGrids
                AtualizaValores
                If blnEmDividaAtiva Then
                    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrGuiaCertidaoDividaAtiva
                End If
                blnReajuste = True
            End If
        
        Case gstrImprimirGuia
            If tdbValoresAcumulado.ApproxCount > 0 Then
                If Not tdbValoresAcumulado.EOF And Not tdbValoresAcumulado.BOF And Len(Trim(tdbValoresAcumulado.Columns("Pkid").Value)) > 0 Then
                    For intCont = 0 To xadbLista.Count(1) - 1
                        If xadbLista(intCont, 1) = -1 Then
                            DBLVALOR = DBLVALOR + xadbLista(intCont, 11)
                            blnCheck = True
                        End If
                    Next
                    If DBLVALOR = 0 And blnCheck = False Then
                        ExibeMensagem "Não existe nenhum registro checado para impressão da guia."
                    ElseIf DBLVALOR = 0 And blnCheck = True Then
                        ExibeMensagem "O total da guia deve ser maior que '0'."
                    Else
                        frmImprimirGuia.txt_strRequerente.Text = txtstrContribuinte.Text
                        CarregaForm frmImprimirGuia
                    End If
                End If
            Else
                ExibeMensagem "É necessário o cálculo para impressão da guia."
            End If
        
        Case gstrGuiaCertidaoNegativa
            If blnNegativa Then
                'If chkTodasReceitas.Value Then
                '    strIPTU = strVerificaIPTU
                'Else
                '    strIPTU = dbcintComposicaoDaReceita.Text
                'End If
                ImprimiNegativa
            End If
        
        Case gstrGuiaCertidaoPositiva
            If Not xadbLista Is Nothing Then
               If xadbLista(0, 0) > 0 Then
                  CarregaForm frmImprimirGuiaPositiva
                  frmImprimirGuiaPositiva.bytTipoPositiva = 0
               Else
                  ExibeMensagem "É necessário o cálculo para impressão."
               End If
            Else
               ExibeMensagem "É necessário o cálculo para impressão."
            End If
        
        Case gstrGuiaCertidaoPositivaEfeitoNegativa
            If Not xadbLista Is Nothing Then
               If xadbLista(0, 0) > 0 Then
                  CarregaForm frmImprimirGuiaPositiva
                  frmImprimirGuiaPositiva.bytTipoPositiva = 1
               Else
                  ExibeMensagem "É necessário o cálculo para impressão."
               End If
            Else
               ExibeMensagem "É necessário o cálculo para impressão."
            End If
        
        Case gstrGuiaRelacaoDeDebitos
            If Not xadbLista Is Nothing Then
                If xadbLista(0, 0) > 0 Then
                    'If MsgBox("Relação de Débitos Detalhada", vbYesNo, "Modelo de relação") = vbYes Then
                        ImprimeRelatorioPorArray rptRelacaoDebitos, vetParcelas, "Informação de Débitos Detalhada"
                    'Else
                    '    ImprimeRelatorioPorArray rptRelacaoDebitosPorExercicio, vetParcelas, "Informação de Débitos Por Exercício"
                    'End If
                Else
                    ExibeMensagem "É necessário o cálculo para impressão."
                End If
            Else
                ExibeMensagem "É necessário o cálculo para impressão."
            End If
            
        Case gstrParcelamentoDebitoAtualizado
            
            Dim lngLinha As Long
            
            Load frmCadAcordos
            frmCadAcordos.MantemForm gstrNovo
            frmCadAcordos.blnParcelamentoDebito = True
            frmCadAcordos.txtdtmData.Text = gstrDataDoSistema(False, False, True)
            frmCadAcordos.txtstrLogradouro.Text = STRENDERECO
            frmCadAcordos.txtstrNumero.Text = strNumero
            frmCadAcordos.txtstrComplemento.Text = STRCOMPLEMENTO
            frmCadAcordos.txtstrBairro.Text = STRBAIRRO
            frmCadAcordos.txtstrMunicipio.Text = STRMUNICIPIO
            frmCadAcordos.txtstrUF.Text = STRUF
            frmCadAcordos.txtintCEP.Text = strCep
            frmCadAcordos.txtstrLogradouroC.Text = strEnderecoC
            frmCadAcordos.txtstrNumeroC.Text = strNumeroC
            frmCadAcordos.txtstrComplementoC.Text = strComplementoC
            frmCadAcordos.txtstrBairroC.Text = strBairroC
            frmCadAcordos.txtstrMunicipioC.Text = strMunicipioC
            frmCadAcordos.txtstrUFC.Text = strUFC
            frmCadAcordos.txtintCEPC.Text = strCepC
            PreencherListaDeOpcoes frmCadAcordos.dbcstrNomeProprietario, lngContribuinte
            
            If Len(vetParcelas(0, 0)) > 0 Then
            
                dblValorAcordo = 0
                ReDim VetParcelasAcordo(15, 0)
                
                For intCont = 0 To UBound(vetParcelas, 2)
                
                    'Vamos verificar se estas parcelas estao com a inscricao checada
                    lngLinha = xadbLista.Find(0, 0, vetParcelas(1, intCont))
                    
                    If vetParcelas(11, intCont) = True And xadbLista(lngLinha, 1) = True And CDbl(gstrConvVrDoSql(gstrENulo(xadbLista(lngLinha, 11)), , , True)) > 0 Then
                        'Vamos verificar se parcelas em Divida Ativa entram no Acordo, desde que nao sejam acordos
                        If (blnAcordoEmDividaAtiva And vetParcelas(19, intCont) = "Não") And Not vetParcelas(18, intCont) = TYP_ACORDO Then GoTo ProximaParcela
                        'Vamos verificar se parcelas ja se encontram em acordos
                        If Len(vetParcelas(9, intCont)) > 1 Then GoTo ProximaParcela
                        
                        dblValorAcordo = dblValorAcordo + vetParcelas(8, intCont)
                        
                        ReDim Preserve VetParcelasAcordo(15, intPosicao)

                        VetParcelasAcordo(0, intPosicao) = vetParcelas(0, intCont)
                        VetParcelasAcordo(1, intPosicao) = vetParcelas(3, intCont)
                        VetParcelasAcordo(2, intPosicao) = vetParcelas(4, intCont)
                        VetParcelasAcordo(3, intPosicao) = vetParcelas(5, intCont)
                        VetParcelasAcordo(4, intPosicao) = vetParcelas(6, intCont)
                        VetParcelasAcordo(5, intPosicao) = vetParcelas(7, intCont)
                        VetParcelasAcordo(6, intPosicao) = vetParcelas(8, intCont)
                        VetParcelasAcordo(7, intPosicao) = vetParcelas(13, intCont)
                        VetParcelasAcordo(8, intPosicao) = vetParcelas(14, intCont)
                        VetParcelasAcordo(9, intPosicao) = vetParcelas(12, intCont)
                        VetParcelasAcordo(10, intPosicao) = vetParcelas(15, intCont)
                        VetParcelasAcordo(11, intPosicao) = vetParcelas(2, intCont)
                        VetParcelasAcordo(12, intPosicao) = vetParcelas(16, intCont)
                        VetParcelasAcordo(13, intPosicao) = vetParcelas(17, intCont) 'Inscrição Pura
                        VetParcelasAcordo(14, intPosicao) = vetParcelas(18, intCont) 'Utilização
                        VetParcelasAcordo(15, intPosicao) = vetParcelas(10, intCont) 'Executivo
                        
                        intPosicao = intPosicao + 1
                        
                    End If
ProximaParcela:
                Next
            End If
            If dblValorAcordo > 0 Then
                
                'Total parcelas + Honorarios
                frmCadAcordos.txtdblValor.Text = Format$(dblValorAcordo + vetTotais(0, 6), "#,##0.00")
                
                frmCadAcordos.dblValorAcordoOriginal = Format$(dblValorAcordo + vetTotais(0, 6), "#,##0.00")
                frmCadAcordos.dblValorHonorarios = Format$(vetTotais(0, 6), "#,##0.00")
                frmCadAcordos.InicializaArrayParcelas VetParcelasAcordo
                frmCadAcordos.PreencheValorIndexador
                
                frmCadAcordos.blnVBModal = True
                Set frmCadAcordos.tdb_Acordos.DataSource = Nothing
                
                CarregaForm frmCadAcordos
                blnAcordoAtivo = True
            Else
                ExibeMensagem "Não foi possivel gerar parcelamento pois o total é igual à (0,00) Zero."
            End If
            
    End Select
                 
End Sub

Private Function blnDadosOK() As Boolean

    blnDadosOK = False
    
    On Error GoTo err_blnDadosOK
    
    If chkTodasReceitas.Value = 0 And chkTodasUtilizacoes.Value = 0 Then
        If dbcintComposicaoDaReceita.MatchedWithList = False And dbcintUtilizacao.MatchedWithList = False Then
           ExibeMensagem "A composição da receita ou a utilização tem que ser selecionada."
           dbcintComposicaoDaReceita.SetFocus
           Exit Function
        End If
    End If
        
    If Len(Trim(txtstrInscricao.Text)) = 0 Then
        ExibeMensagem "A inscrição cadastral deve ser informada."
        txtstrInscricao.SetFocus
        Exit Function
    End If
    
    blnDadosOK = True

err_blnDadosOK:

End Function

Private Function strQueryComposicao() As String
Dim strSQL As String
    
    strSQL = "SELECT Pkid,"
    strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita
    strSQL = strSQL & " WHERE bytDividaAtiva = 1 "
    strSQL = strSQL & " ORDER BY strDescricao"
    
    strQueryComposicao = strSQL

End Function

Private Sub tab_3dPasta_DblClick()
'frm_Arquiv_Banco.Show
End Sub

Private Sub tdbValoresAcumulado_ColEdit(ByVal ColIndex As Integer)
Dim intForAcumulado As Integer
    With tdbValoresAcumulado
        If Not .EOF And Not .BOF And Len(Trim(.Columns("Pkid").Value)) > 0 Then
            If ColIndex = 1 Then
                For intForAcumulado = 0 To .ApproxCount - 1
                    If xadbLista(intForAcumulado, 0) = .Columns("Pkid").Value Then
                        xadbLista(intForAcumulado, 1) = .Columns(1).Value
                        
                        vetTotais(0, 0) = IIf(Len(vetTotais(0, 0)) > 0, vetTotais(0, 0), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 6)), CCur(xadbLista(intForAcumulado, 6)) * -1)
                        vetTotais(0, 1) = IIf(Len(vetTotais(0, 1)) > 0, vetTotais(0, 1), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 7)), CCur(xadbLista(intForAcumulado, 7)) * -1)
                        vetTotais(0, 2) = IIf(Len(vetTotais(0, 2)) > 0, vetTotais(0, 2), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 8)), CCur(xadbLista(intForAcumulado, 8)) * -1)
                        vetTotais(0, 3) = IIf(Len(vetTotais(0, 3)) > 0, vetTotais(0, 3), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 9)), CCur(xadbLista(intForAcumulado, 9)) * -1)
                        vetTotais(0, 4) = IIf(Len(vetTotais(0, 4)) > 0, vetTotais(0, 4), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 10)), CCur(xadbLista(intForAcumulado, 10)) * -1)
                        vetTotais(0, 5) = IIf(Len(vetTotais(0, 5)) > 0, vetTotais(0, 5), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 11)), CCur(xadbLista(intForAcumulado, 11)) * -1)
                        vetTotais(0, 6) = IIf(Len(vetTotais(0, 6)) > 0, vetTotais(0, 6), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 12)), CCur(xadbLista(intForAcumulado, 12)) * -1)
                        vetTotais(0, 7) = IIf(Len(vetTotais(0, 7)) > 0, vetTotais(0, 7), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 13)), CCur(xadbLista(intForAcumulado, 13)) * -1)
                        vetTotais(0, 8) = IIf(Len(vetTotais(0, 8)) > 0, vetTotais(0, 8), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 14)), CCur(xadbLista(intForAcumulado, 14)) * -1)
                        vetTotais(0, 9) = IIf(Len(vetTotais(0, 9)) > 0, vetTotais(0, 9), 0) + IIf(xadbLista(intForAcumulado, 1) = True, CCur(xadbLista(intForAcumulado, 15)), CCur(xadbLista(intForAcumulado, 15)) * -1)
                        
                        lbl_dblTotOriginal = gstrConvVrDoSql(vetTotais(0, 0), 2)
                        lbl_dblTotPrincipal = gstrConvVrDoSql(vetTotais(0, 1), 2)
                        lbl_dblTotMulta = gstrConvVrDoSql(vetTotais(0, 2), 2)
                        lbl_dblTotJuros = gstrConvVrDoSql(vetTotais(0, 3), 2)
                        lbl_dblTotCorrecao = gstrConvVrDoSql(vetTotais(0, 4), 2)
                        lbl_dblTotTotal = gstrConvVrDoSql(vetTotais(0, 5), 2)
                        lbl_dblTotHonorarios = gstrConvVrDoSql(vetTotais(0, 6), 2)
                        lbl_dblTotGeral = gstrConvVrDoSql(vetTotais(0, 9), 2)
                        
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub tdbValoresAcumulado_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

    lbl_dblTotOriginal.Width = IIf(tdbValoresAcumulado.Columns("dblValorOriginal").Visible, tdbValoresAcumulado.Columns("dblValorOriginal").Width, 0)
    lbl_dblTotPrincipal.Width = IIf(tdbValoresAcumulado.Columns("dblPrincipal").Visible, tdbValoresAcumulado.Columns("dblPrincipal").Width, 0) + 15
    lbl_dblTotMulta.Width = IIf(tdbValoresAcumulado.Columns("dblMulta").Visible, tdbValoresAcumulado.Columns("dblMulta").Width, 0) + 15
    lbl_dblTotJuros.Width = IIf(tdbValoresAcumulado.Columns("dblJuros").Visible, tdbValoresAcumulado.Columns("dblJuros").Width, 0) + 15
    lbl_dblTotCorrecao.Width = IIf(tdbValoresAcumulado.Columns("dblCorrecao").Visible, tdbValoresAcumulado.Columns("dblCorrecao").Width, 0) + 15
    lbl_dblTotTotal.Width = IIf(tdbValoresAcumulado.Columns("dblTotal").Visible, tdbValoresAcumulado.Columns("dblTotal").Width, 0) + 15
    lbl_dblTotHonorarios.Width = IIf(tdbValoresAcumulado.Columns("dblHonorarios").Visible, tdbValoresAcumulado.Columns("dblHonorarios").Width, 0) + 15
    lbl_dblTotGeral.Width = IIf(tdbValoresAcumulado.Columns("dblTotalGeral").Visible, tdbValoresAcumulado.Columns("dblTotalGeral").Width, 0) + 15
    
    If ColIndex <= tdbValoresAcumulado.Columns("dblValorOriginal").ColIndex Then
        lbl_dblTotOriginal.Left = tdbValoresAcumulado.Columns("dblValorOriginal").Left + 155
    End If
    If ColIndex <= tdbValoresAcumulado.Columns("dblPrincipal").ColIndex Then
        lbl_dblTotPrincipal.Left = lbl_dblTotOriginal.Left + lbl_dblTotOriginal.Width - 15
    End If
    If ColIndex <= tdbValoresAcumulado.Columns("dblMulta").ColIndex Then
        lbl_dblTotMulta.Left = lbl_dblTotPrincipal.Left + lbl_dblTotPrincipal.Width - 15
    End If
    If ColIndex <= tdbValoresAcumulado.Columns("dblJuros").ColIndex Then
        lbl_dblTotJuros.Left = lbl_dblTotMulta.Left + lbl_dblTotMulta.Width - 15
    End If
    If ColIndex <= tdbValoresAcumulado.Columns("dblCorrecao").ColIndex Then
        lbl_dblTotCorrecao.Left = lbl_dblTotJuros.Left + lbl_dblTotJuros.Width - 15
    End If
    If ColIndex <= tdbValoresAcumulado.Columns("dblTotal").ColIndex Then
        lbl_dblTotTotal.Left = lbl_dblTotCorrecao.Left + lbl_dblTotCorrecao.Width - 15
    End If
    If ColIndex <= tdbValoresAcumulado.Columns("dblHonorarios").ColIndex Then
        lbl_dblTotHonorarios.Left = lbl_dblTotTotal.Left + lbl_dblTotTotal.Width - 15
    End If
    If ColIndex <= tdbValoresAcumulado.Columns("dblTotalGeral").ColIndex Then
        lbl_dblTotGeral.Left = lbl_dblTotHonorarios.Left + lbl_dblTotHonorarios.Width - 15
    End If
    
End Sub

Private Sub tdbValoresAcumulado_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Select Case ColIndex
        Case 6, 7, 8, 9, 10, 11, 12, 13, 14, 15
            Value = gstrConvVrDoSql(Value, 2)
    End Select
End Sub

Private Sub tdbValoresAcumulado_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdbValoresAcumulado
        
        If Not .EOF And Not .BOF And Len(Trim(.Columns("PkidLA").Value)) > 0 Then
            
            gCorLinhaSelecionada tdbValoresAcumulado

            ExibeParcelas .Columns("PkidLA").Value
            ExibeParcelasOpcionais .Columns("PkidLA").Value
            
        End If
    
    End With
    
End Sub

Private Sub tdbValoresParcelas_ColEdit(ByVal ColIndex As Integer)
Dim intForParcela     As Integer
Dim intForAcumulado   As Integer
Dim blnParcelaChecada As Boolean
Dim intParcelaValida  As Integer
    
    DoEvents
    
    blnParcelaChecada = False
    intParcelaValida = 0
        
    Screen.MousePointer = vbHourglass
    
    'Vamos atualizar o valor do check no array
    With tdbValoresParcelas
        .Enabled = False 'Para nao encavalar processamentos
        .RowDividerColor = dbgLightGrayLine
        If Not .EOF And Not .BOF And Len(Trim(.Columns("PkidLV").Value)) > 0 Then
            If ColIndex = 1 Then
            
                If tdbValoresParcelasOpcionais.ApproxCount > 0 Then
                         
                    'Vamos varrer todas as parcelas opcionais e descheca-las
                    For intForParcela = 0 To UBound(vetParcelasOpcionais, 2)
                        
                        If vetParcelasOpcionais(1, intForParcela) = .Columns("intLancamentoAlfa").Value Then
                        
                            'Caso esteja checada vamos deschecar
                            If vetParcelasOpcionais(11, intForParcela) = True Then
                            
                                xadbParcelasOpcionais(intParcelaValida, 1) = 0
                                vetParcelasOpcionais(11, intForParcela) = 0
                        
                                'Vamos atualizar os valores do grid Acumulado, caso nao seja uma parcela em acordo
                                If Len(Trim(vetParcelasOpcionais(9, intForParcela))) <= 1 Then
                                
                                    For intForAcumulado = 0 To xadbLista.UpperBound(1)
                                        If xadbLista(intForAcumulado, 0) = vetParcelasOpcionais(1, intForParcela) Then
                                            xadbLista(intForAcumulado, 6) = xadbLista(intForAcumulado, 6) + vetParcelasOpcionais(3, intForParcela) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 0) = IIf(Len(vetTotais(0, 0)) > 0, vetTotais(0, 0), 0) + CCur(vetParcelasOpcionais(3, intForParcela)) * -1
                                            xadbLista(intForAcumulado, 7) = xadbLista(intForAcumulado, 7) + vetParcelasOpcionais(4, intForParcela) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 1) = IIf(Len(vetTotais(0, 1)) > 0, vetTotais(0, 1), 0) + CCur(vetParcelasOpcionais(4, intForParcela)) * -1
                                            xadbLista(intForAcumulado, 8) = xadbLista(intForAcumulado, 8) + vetParcelasOpcionais(5, intForParcela) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 2) = IIf(Len(vetTotais(0, 2)) > 0, vetTotais(0, 2), 0) + CCur(vetParcelasOpcionais(5, intForParcela)) * -1
                                            xadbLista(intForAcumulado, 9) = xadbLista(intForAcumulado, 9) + vetParcelasOpcionais(6, intForParcela) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 3) = IIf(Len(vetTotais(0, 3)) > 0, vetTotais(0, 3), 0) + CCur(vetParcelasOpcionais(6, intForParcela)) * -1
                                            xadbLista(intForAcumulado, 10) = xadbLista(intForAcumulado, 10) + vetParcelasOpcionais(7, intForParcela) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 4) = IIf(Len(vetTotais(0, 4)) > 0, vetTotais(0, 4), 0) + CCur(vetParcelasOpcionais(7, intForParcela)) * -1
                                            xadbLista(intForAcumulado, 11) = xadbLista(intForAcumulado, 11) + vetParcelasOpcionais(8, intForParcela) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 5) = IIf(Len(vetTotais(0, 5)) > 0, vetTotais(0, 5), 0) + CCur(vetParcelasOpcionais(8, intForParcela)) * -1
                                            
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 6) = vetTotais(0, 6) - xadbLista(intForAcumulado, 12)
                                            xadbLista(intForAcumulado, 12) = dblCalculaEncargos(BIT_HONORARIOS, xadbLista(intForAcumulado, 11), xadbLista(intForAcumulado, 17))
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 6) = IIf(Len(vetTotais(0, 6)) > 0, vetTotais(0, 6), 0) + xadbLista(intForAcumulado, 12)
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 9) = vetTotais(0, 9) - xadbLista(intForAcumulado, 15)
                                            xadbLista(intForAcumulado, 15) = xadbLista(intForAcumulado, 11) + xadbLista(intForAcumulado, 12)
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 9) = IIf(Len(vetTotais(0, 9)) > 0, vetTotais(0, 9), 0) + xadbLista(intForAcumulado, 15)
                                            
                                        End If
                                    Next
                                
                                End If
                                
                                DoEvents
                            
                            End If
                            
                            intParcelaValida = intParcelaValida + 1
                            
                        End If
                     Next
                            
                    tdbValoresParcelasOpcionais.ReBind
                    tdbValoresParcelasOpcionais.Refresh
                
                End If
                
                'Vamos varrer todas as parcelas
                For intForParcela = 0 To UBound(vetParcelas, 2)
                    
                    If vetParcelas(0, intForParcela) = .Columns("PkidLV").Value Then
                        'vetParcelas(11, intForParcela) = .Columns(1).Value
                        vetParcelas(11, intForParcela) = IIf(vetParcelas(11, intForParcela) = True, False, True)
                        
                        'Vamos atualizar os valores do grid Acumulado, caso nao seja uma parcela em acordo
                        If Len(Trim(vetParcelas(9, intForParcela))) <= 1 Then
                        
                            For intForAcumulado = 0 To xadbLista.UpperBound(1)
                                If xadbLista(intForAcumulado, 0) = vetParcelas(1, intForParcela) Then
                                    
                                    xadbLista(intForAcumulado, 6) = xadbLista(intForAcumulado, 6) + IIf(vetParcelas(11, intForParcela) = True, vetParcelas(3, intForParcela), vetParcelas(3, intForParcela) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 0) = IIf(Len(vetTotais(0, 0)) > 0, vetTotais(0, 0), 0) + IIf(vetParcelas(11, intForParcela) = True, CCur(vetParcelas(3, intForParcela)), CCur(vetParcelas(3, intForParcela)) * -1)
                                    xadbLista(intForAcumulado, 7) = xadbLista(intForAcumulado, 7) + IIf(vetParcelas(11, intForParcela) = True, vetParcelas(4, intForParcela), vetParcelas(4, intForParcela) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 1) = IIf(Len(vetTotais(0, 1)) > 0, vetTotais(0, 1), 0) + IIf(vetParcelas(11, intForParcela) = True, CCur(vetParcelas(4, intForParcela)), CCur(vetParcelas(4, intForParcela)) * -1)
                                    xadbLista(intForAcumulado, 8) = xadbLista(intForAcumulado, 8) + IIf(vetParcelas(11, intForParcela) = True, vetParcelas(5, intForParcela), vetParcelas(5, intForParcela) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 2) = IIf(Len(vetTotais(0, 2)) > 0, vetTotais(0, 2), 0) + IIf(vetParcelas(11, intForParcela) = True, CCur(vetParcelas(5, intForParcela)), CCur(vetParcelas(5, intForParcela)) * -1)
                                    xadbLista(intForAcumulado, 9) = xadbLista(intForAcumulado, 9) + IIf(vetParcelas(11, intForParcela) = True, vetParcelas(6, intForParcela), vetParcelas(6, intForParcela) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 3) = IIf(Len(vetTotais(0, 3)) > 0, vetTotais(0, 3), 0) + IIf(vetParcelas(11, intForParcela) = True, CCur(vetParcelas(6, intForParcela)), CCur(vetParcelas(6, intForParcela)) * -1)
                                    xadbLista(intForAcumulado, 10) = xadbLista(intForAcumulado, 10) + IIf(vetParcelas(11, intForParcela) = True, vetParcelas(7, intForParcela), vetParcelas(7, intForParcela) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 4) = IIf(Len(vetTotais(0, 4)) > 0, vetTotais(0, 4), 0) + IIf(vetParcelas(11, intForParcela) = True, CCur(vetParcelas(7, intForParcela)), CCur(vetParcelas(7, intForParcela)) * -1)
                                    xadbLista(intForAcumulado, 11) = xadbLista(intForAcumulado, 11) + IIf(vetParcelas(11, intForParcela) = True, vetParcelas(8, intForParcela), vetParcelas(8, intForParcela) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 5) = IIf(Len(vetTotais(0, 5)) > 0, vetTotais(0, 5), 0) + IIf(vetParcelas(11, intForParcela) = True, CCur(vetParcelas(8, intForParcela)), CCur(vetParcelas(8, intForParcela)) * -1)
                                    
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 6) = vetTotais(0, 6) - xadbLista(intForAcumulado, 12)
                                    xadbLista(intForAcumulado, 12) = dblCalculaEncargos(BIT_HONORARIOS, xadbLista(intForAcumulado, 11), xadbLista(intForAcumulado, 17))
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 6) = IIf(Len(vetTotais(0, 6)) > 0, vetTotais(0, 6), 0) + xadbLista(intForAcumulado, 12)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 9) = vetTotais(0, 9) - xadbLista(intForAcumulado, 15)
                                    xadbLista(intForAcumulado, 15) = xadbLista(intForAcumulado, 11) + xadbLista(intForAcumulado, 12)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 9) = IIf(Len(vetTotais(0, 9)) > 0, vetTotais(0, 9), 0) + xadbLista(intForAcumulado, 15)
                                    
                                End If
                            Next
                        
                        End If
                        
                        DoEvents
                        
                        tdbValoresAcumulado.ReBind
                        tdbValoresAcumulado.Refresh
                        
                        lbl_dblTotOriginal = gstrConvVrDoSql(vetTotais(0, 0), 2)
                        lbl_dblTotPrincipal = gstrConvVrDoSql(vetTotais(0, 1), 2)
                        lbl_dblTotMulta = gstrConvVrDoSql(vetTotais(0, 2), 2)
                        lbl_dblTotJuros = gstrConvVrDoSql(vetTotais(0, 3), 2)
                        lbl_dblTotCorrecao = gstrConvVrDoSql(vetTotais(0, 4), 2)
                        lbl_dblTotTotal = gstrConvVrDoSql(vetTotais(0, 5), 2)
                        lbl_dblTotHonorarios = gstrConvVrDoSql(vetTotais(0, 6), 2)
                        lbl_dblTotGeral = gstrConvVrDoSql(vetTotais(0, 9), 2)
                        
                    End If
                    
                    If vetParcelas(11, intForParcela) <> "" Then blnParcelaChecada = True
                    
                Next
                
                HabilitaDesabilitaBotao1 blnParcelaChecada, gstrBtnArquivo, gstrParcelamentoDebitoAtualizado
                
            End If
        End If
        .Enabled = True 'Para nao encavalar processamentos
    End With
    
    Screen.MousePointer = vbDefault
    
    DoEvents
    
End Sub

Private Sub tdbValoresParcelas_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Select Case ColIndex
        Case 4, 5, 6, 7, 8, 9
            Value = gstrConvVrDoSql(Value, 2)
    End Select
End Sub

Private Sub tdbValoresParcelas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdbValoresParcelas
        
        If Not .EOF And Not .BOF Then
            gCorLinhaSelecionada tdbValoresParcelas
        End If
    End With
    
End Sub

Private Sub tdbValoresParcelasOpcionais_ColEdit(ByVal ColIndex As Integer)
Dim intForParcelas As Integer
Dim intForAcumulado  As Integer
Dim intParcelaValida As Integer
    
    DoEvents
    
    intParcelaValida = 0
    
    Screen.MousePointer = vbHourglass
    
    'Vamos atualizar o valor do check no array
    With tdbValoresParcelasOpcionais
        If Not .EOF And Not .BOF And Len(Trim(.Columns("PkidLV").Value)) > 0 Then
            If ColIndex = 1 Then
            
                If tdbValoresParcelas.ApproxCount > 0 Then
                    
                    'Vamos varrer todas as parcelas principais e descheca-las
                    For intForParcelas = 0 To UBound(vetParcelas, 2)
                        
                        If vetParcelas(1, intForParcelas) = .Columns("intLancamentoAlfa").Value Then
                            
                            If vetParcelas(11, intForParcelas) = True Then
                            
                                xadbParcelas(intParcelaValida, 1) = 0
                                vetParcelas(11, intForParcelas) = 0
                        
                                'Vamos atualizar os valores do grid Acumulado, caso nao seja uma parcela em acordo
                                If Len(Trim(vetParcelas(9, intForParcelas))) <= 1 Then
                                
                                    For intForAcumulado = 0 To xadbLista.UpperBound(1)
                                        If xadbLista(intForAcumulado, 0) = vetParcelas(1, intForParcelas) Then
                                            xadbLista(intForAcumulado, 6) = xadbLista(intForAcumulado, 6) + vetParcelas(3, intForParcelas) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 0) = IIf(Len(vetTotais(0, 0)) > 0, vetTotais(0, 0), 0) + CCur(vetParcelas(3, intForParcelas)) * -1
                                            xadbLista(intForAcumulado, 7) = xadbLista(intForAcumulado, 7) + vetParcelas(4, intForParcelas) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 1) = IIf(Len(vetTotais(0, 1)) > 0, vetTotais(0, 1), 0) + CCur(vetParcelas(4, intForParcelas)) * -1
                                            xadbLista(intForAcumulado, 8) = xadbLista(intForAcumulado, 8) + vetParcelas(5, intForParcelas) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 2) = IIf(Len(vetTotais(0, 2)) > 0, vetTotais(0, 2), 0) + CCur(vetParcelas(5, intForParcelas)) * -1
                                            xadbLista(intForAcumulado, 9) = xadbLista(intForAcumulado, 9) + vetParcelas(6, intForParcelas) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 3) = IIf(Len(vetTotais(0, 3)) > 0, vetTotais(0, 3), 0) + CCur(vetParcelas(6, intForParcelas)) * -1
                                            xadbLista(intForAcumulado, 10) = xadbLista(intForAcumulado, 10) + vetParcelas(7, intForParcelas) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 4) = IIf(Len(vetTotais(0, 4)) > 0, vetTotais(0, 4), 0) + CCur(vetParcelas(7, intForParcelas)) * -1
                                            xadbLista(intForAcumulado, 11) = xadbLista(intForAcumulado, 11) + vetParcelas(8, intForParcelas) * -1
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 5) = IIf(Len(vetTotais(0, 5)) > 0, vetTotais(0, 5), 0) + CCur(vetParcelas(8, intForParcelas)) * -1
                                            
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 6) = vetTotais(0, 6) - xadbLista(intForAcumulado, 12)
                                            xadbLista(intForAcumulado, 12) = dblCalculaEncargos(BIT_HONORARIOS, xadbLista(intForAcumulado, 11), xadbLista(intForAcumulado, 17))
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 6) = IIf(Len(vetTotais(0, 6)) > 0, vetTotais(0, 6), 0) + xadbLista(intForAcumulado, 12)
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 9) = vetTotais(0, 9) - xadbLista(intForAcumulado, 15)
                                            xadbLista(intForAcumulado, 15) = xadbLista(intForAcumulado, 11) + xadbLista(intForAcumulado, 12)
                                            If xadbLista(intForAcumulado, 1) Then vetTotais(0, 9) = IIf(Len(vetTotais(0, 9)) > 0, vetTotais(0, 9), 0) + xadbLista(intForAcumulado, 15)
                                            
                                            
                                        End If
                                    Next
                                
                                End If
                                
                            End If
                            
                            'Aponta para a parcela atual referente ao lancamento alfa
                            intParcelaValida = intParcelaValida + 1
                            
                            DoEvents
                        
                        End If
                     Next
                            
                    tdbValoresParcelas.ReBind
                    tdbValoresParcelas.Refresh
                    
                End If
                
                'Vamos desabilitar a opcao de parcelamento de debitos
                HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrParcelamentoDebitoAtualizado
                
                'Vamos todas as parcelas opcionais
                For intForParcelas = 0 To UBound(vetParcelasOpcionais, 2)
                    If vetParcelasOpcionais(0, intForParcelas) = .Columns("PkidLV").Value Then
                    
                        'vetParcelasOpcionais(11, intForParcelas) = .Columns(1).Value
                        vetParcelasOpcionais(11, intForParcelas) = IIf(vetParcelasOpcionais(11, intForParcelas) = True, False, True)
                        
                        'Vamos atualizar os valores do grid Acumulado, caso nao seja uma parcela em acordo
                        If Len(Trim(vetParcelasOpcionais(9, intForParcelas))) <= 1 Then
                        
                            For intForAcumulado = 0 To xadbLista.UpperBound(1)
                                If xadbLista(intForAcumulado, 0) = vetParcelasOpcionais(1, intForParcelas) Then
                                    
                                    xadbLista(intForAcumulado, 6) = xadbLista(intForAcumulado, 6) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, vetParcelasOpcionais(3, intForParcelas), vetParcelasOpcionais(3, intForParcelas) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 0) = IIf(Len(vetTotais(0, 0)) > 0, vetTotais(0, 0), 0) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, CCur(vetParcelasOpcionais(3, intForParcelas)), CCur(vetParcelasOpcionais(3, intForParcelas)) * -1)
                                    xadbLista(intForAcumulado, 7) = xadbLista(intForAcumulado, 7) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, vetParcelasOpcionais(4, intForParcelas), vetParcelasOpcionais(4, intForParcelas) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 1) = IIf(Len(vetTotais(0, 1)) > 0, vetTotais(0, 1), 0) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, CCur(vetParcelasOpcionais(4, intForParcelas)), CCur(vetParcelasOpcionais(4, intForParcelas)) * -1)
                                    xadbLista(intForAcumulado, 8) = xadbLista(intForAcumulado, 8) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, vetParcelasOpcionais(5, intForParcelas), vetParcelasOpcionais(5, intForParcelas) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 2) = IIf(Len(vetTotais(0, 2)) > 0, vetTotais(0, 2), 0) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, CCur(vetParcelasOpcionais(5, intForParcelas)), CCur(vetParcelasOpcionais(5, intForParcelas)) * -1)
                                    xadbLista(intForAcumulado, 9) = xadbLista(intForAcumulado, 9) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, vetParcelasOpcionais(6, intForParcelas), vetParcelasOpcionais(6, intForParcelas) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 3) = IIf(Len(vetTotais(0, 3)) > 0, vetTotais(0, 3), 0) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, CCur(vetParcelasOpcionais(6, intForParcelas)), CCur(vetParcelasOpcionais(6, intForParcelas)) * -1)
                                    xadbLista(intForAcumulado, 10) = xadbLista(intForAcumulado, 10) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, vetParcelasOpcionais(7, intForParcelas), vetParcelasOpcionais(7, intForParcelas) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 4) = IIf(Len(vetTotais(0, 4)) > 0, vetTotais(0, 4), 0) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, CCur(vetParcelasOpcionais(7, intForParcelas)), CCur(vetParcelasOpcionais(7, intForParcelas)) * -1)
                                    xadbLista(intForAcumulado, 11) = xadbLista(intForAcumulado, 11) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, vetParcelasOpcionais(8, intForParcelas), vetParcelasOpcionais(8, intForParcelas) * -1)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 5) = IIf(Len(vetTotais(0, 5)) > 0, vetTotais(0, 5), 0) + IIf(vetParcelasOpcionais(11, intForParcelas) = True, CCur(vetParcelasOpcionais(8, intForParcelas)), CCur(vetParcelasOpcionais(8, intForParcelas)) * -1)
                                    
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 6) = vetTotais(0, 6) - xadbLista(intForAcumulado, 12)
                                    xadbLista(intForAcumulado, 12) = dblCalculaEncargos(BIT_HONORARIOS, xadbLista(intForAcumulado, 11), xadbLista(intForAcumulado, 17))
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 6) = IIf(Len(vetTotais(0, 6)) > 0, vetTotais(0, 6), 0) + xadbLista(intForAcumulado, 12)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 9) = vetTotais(0, 9) - xadbLista(intForAcumulado, 15)
                                    xadbLista(intForAcumulado, 15) = xadbLista(intForAcumulado, 11) + xadbLista(intForAcumulado, 12)
                                    If xadbLista(intForAcumulado, 1) Then vetTotais(0, 9) = IIf(Len(vetTotais(0, 9)) > 0, vetTotais(0, 9), 0) + xadbLista(intForAcumulado, 15)
                                    
                                End If
                            Next
                        
                        End If
                        
                        DoEvents
                        
                        tdbValoresAcumulado.ReBind
                        tdbValoresAcumulado.Refresh
                        
                        lbl_dblTotOriginal = gstrConvVrDoSql(vetTotais(0, 0), 2)
                        lbl_dblTotPrincipal = gstrConvVrDoSql(vetTotais(0, 1), 2)
                        lbl_dblTotMulta = gstrConvVrDoSql(vetTotais(0, 2), 2)
                        lbl_dblTotJuros = gstrConvVrDoSql(vetTotais(0, 3), 2)
                        lbl_dblTotCorrecao = gstrConvVrDoSql(vetTotais(0, 4), 2)
                        lbl_dblTotTotal = gstrConvVrDoSql(vetTotais(0, 5), 2)
                        lbl_dblTotHonorarios = gstrConvVrDoSql(vetTotais(0, 6), 2)
                        lbl_dblTotGeral = gstrConvVrDoSql(vetTotais(0, 9), 2)
                        
                    End If
                Next
                    
            End If
        End If
    
    End With
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub tdbValoresParcelasOpcionais_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Select Case ColIndex
        Case 4, 5, 6, 7, 8, 9
            Value = gstrConvVrDoSql(Value, 2)
    End Select
End Sub

Private Sub tdbValoresParcelasOpcionais_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdbValoresParcelasOpcionais
        If Not .EOF And Not .BOF Then
            gCorLinhaSelecionada tdbValoresParcelasOpcionais
        End If
    End With
    
End Sub

Private Sub txtstrInscricao_GotFocus()
    MarcaCampo txtstrInscricao
End Sub

Private Sub txtstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrInscricao
End Sub

Private Sub AtualizaValores()
Dim adoResultado  As New ADODB.Recordset
Dim adoLogradouro As New ADODB.Recordset
Dim strSQL        As String
    
    Set gobjBanco = New clsBanco
    Screen.MousePointer = vbHourglass
    
    'Vamos fazer a busca da inscricao informada
    strSQL = "SELECT LA.pkid, LA.strInscricaoAuxiliar, LA.strLogradouro, LA.strNUMERO, LA.STRCOMPLEMENTO, LA.INTCEP, " & _
             "LA.STRBAIRRO, LA.struf, LA.STRMUNICIPIO, LA.strlogradouroc, LA.strNumeroC, LA.strComplementoC, LA.strBairroC, LA.strMunicipioC, LA.strufc, LA.intcepc, LA.strNomeProprietario, " & gstrCONVERT(cdt_numeric, "LA.strNumeroAviso") & " strNumeroAviso, LA.intUtilizacao FROM " & gstrLancamentoAlfa & " LA WHERE strInscricao = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "' AND " & _
             "(LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
    'strSql = "SELECT LA.strNomeProprietario, LA.strNumeroAviso FROM " & gstrLancamentoAlfa & " LA WHERE strInscricao = '" & txtstrInscricao.Text & "'"
    If dbcintComposicaoDaReceita.MatchedWithList Then
        strSQL = strSQL & " AND LA.intComposicaoDaReceita = " & dbcintComposicaoDaReceita.BoundText
    ElseIf dbcintUtilizacao.MatchedWithList Then
        strSQL = strSQL & " AND LA.intutilizacao = " & dbcintUtilizacao.BoundText
    End If
    strSQL = strSQL & " ORDER BY LA.intUtilizacao, LA.intExercicio DESC, LA.Pkid DESC"
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txtstrContribuinte.Text = adoResultado("strNomeProprietario").Value 'Inserir o Proprietario
            strNumeroAviso = gstrENulo(adoResultado("strNumeroAviso").Value)    'Inserir o Aviso
            lngDividaAtiva = gstrENulo(adoResultado("Pkid").Value)              'Inserir o Pkid para verifiacar divida ativa
            strInscricaoAuxiliar = gstrENulo(adoResultado("strInscricaoAuxiliar").Value)    'Inserir a Inscricao Auxiliar
            intUtilizacao = gstrENulo(adoResultado("intUtilizacao").Value)
            'Vamos achar o endereço do Imóvel
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strQueryLogradouroImovel(adoResultado("intUtilizacao").Value), 5, adoLogradouro) Then
                If Not adoLogradouro.EOF Then
                    strLogradouro = gstrENulo(adoLogradouro!strLogradouro) & " " & gstrENulo(adoLogradouro!INTNUMERO) & " " & gstrENulo(adoLogradouro!STRCOMPLEMENTO) & " CEP: " & gstrCEPFormatado(gstrENulo(adoLogradouro!INTCEP))
                    STRENDERECO = gstrENulo(adoLogradouro!strLogradouro)
                    strNumero = gstrENulo(adoLogradouro!INTNUMERO)
                    STRCOMPLEMENTO = gstrENulo(adoLogradouro!STRCOMPLEMENTO)
                    strCep = Format(gstrENulo(adoLogradouro!INTCEP), "00000-000")
                    STRBAIRRO = gstrENulo(adoLogradouro!STRBAIRRO) & " / " & gstrENulo(adoLogradouro!strEstado)
                    STRMUNICIPIO = gstrENulo(adoLogradouro!STRMUNICIPIO)
                    STRUF = gstrENulo(adoLogradouro!strEstado)
                    strEnderecoC = gstrENulo(adoLogradouro!strLogradouroC)
                    strNumeroC = gstrENulo(adoLogradouro!intNumeroC)
                    strComplementoC = gstrENulo(adoLogradouro!strComplementoC)
                    strBairroC = gstrENulo(adoLogradouro!strBairroC)
                    strMunicipioC = gstrENulo(adoLogradouro!strMunicipioC)
                    strUFC = gstrENulo(adoLogradouro!strestadoc)
                    strCepC = Format(gstrENulo(adoLogradouro!INTCEPC), "00000-000")
                Else
                    strLogradouro = gstrENulo(adoResultado!strLogradouro) & " " & gstrENulo(adoResultado!strNumero) & " " & gstrENulo(adoResultado!STRCOMPLEMENTO) & " CEP: " & gstrCEPFormatado(gstrENulo(adoResultado!INTCEP))
                    STRENDERECO = gstrENulo(adoResultado!strLogradouro)
                    strNumero = gstrENulo(adoResultado!strNumero)
                    STRCOMPLEMENTO = gstrENulo(adoResultado!STRCOMPLEMENTO)
                    strCep = Format(gstrENulo(adoResultado!INTCEP), "00000-000")
                    STRBAIRRO = gstrENulo(adoResultado!STRBAIRRO) & " / " & gstrENulo(adoResultado!STRUF)
                    STRMUNICIPIO = gstrENulo(adoResultado!STRMUNICIPIO)
                    STRUF = gstrENulo(adoResultado!STRUF)
                    strEnderecoC = gstrENulo(adoResultado!strLogradouroC)
                    strNumeroC = gstrENulo(adoResultado!strNumeroC)
                    strComplementoC = gstrENulo(adoResultado!strComplementoC)
                    strBairroC = gstrENulo(adoResultado!strBairroC)
                    strMunicipioC = gstrENulo(adoResultado!strMunicipioC)
                    strUFC = gstrENulo(adoResultado!strUFC)
                    strCepC = Format(gstrENulo(adoResultado!INTCEPC), "00000-000")
                End If
                blnNegativa = True
            End If
            PreencheGrid
        Else
            ExibeMensagem "Não foi(ram) encontrado(s) registro(s) com esta Inscrição em Lançamento Alfa, ou a Composição de Receita não inscreve em Divida Ativa."
            LimpaGrids
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub LimpaGrids()
    
    ReDim vetTotais(0, 9)
    
    Set xadbLista = New XArrayDB
    xadbLista.Clear
    xadbLista.ReDim 0, 0, 0, 23
    
    Set tdbValoresAcumulado.Array = xadbLista
    tdbValoresAcumulado.ReBind
    tdbValoresAcumulado.Refresh
    
    Set xadbParcelas = New XArrayDB
    xadbParcelas.Clear
    xadbParcelas.ReDim 0, 0, 0, 13
    
    Set tdbValoresParcelas.Array = xadbParcelas
    tdbValoresParcelas.ReBind
    tdbValoresParcelas.Refresh

    Set xadbParcelasOpcionais = New XArrayDB
    xadbParcelasOpcionais.Clear
    xadbParcelasOpcionais.ReDim 0, 0, 0, 13
    
    Set tdbValoresParcelasOpcionais.Array = xadbParcelasOpcionais
    tdbValoresParcelasOpcionais.ReBind
    tdbValoresParcelasOpcionais.Refresh

    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrGuiaCertidaoPositivaEfeitoNegativa
    
End Sub

Sub PreencheGrid()
Dim c       As Integer
    
    c = tdbValoresAcumulado.Col
    tdbValoresAcumulado.HoldFields
    
    MontaArray
    
    tdbValoresAcumulado.Col = c
    'tdbValoresAcumulado.EditActive = True
    tdbValoresAcumulado.CurrentCellModified = True
    
End Sub

Private Sub MontaArray()
Dim adoResultado      As ADODB.Recordset
Dim adoParcelas       As ADODB.Recordset
'Dim adoParcelas       As ADODB.Parameters
Dim blnParcelas       As Boolean
Dim intFor            As Integer
Dim intForOpcionais   As Integer
Dim strSQL            As String
Dim varAux            As Variant
Dim intPosition       As Integer

Dim lngUltInscricao    As Long
Dim lngProxInscricao   As Long
Dim dblTotalOriginal   As Double
Dim dblTotalPrincipal  As Double
Dim dblTotalMulta      As Double
Dim dblTotalJuros      As Double
Dim dblTotalCorrecao   As Double
    
Dim blnParcelaChecada      As Boolean
Dim blnParcelaComAcordo    As Boolean

Dim strInscricoes          As String
Dim strAcordosParaConsulta As String

'Variaveis de apoio para preencher colunas do grid utilizadas no relat. certidao positiva
Dim intQtdeParcelasInicial  As Integer
Dim dblValorParcelasInicial As Double
Dim strVencto1Parcela       As String
Dim strAcordosRelacionados  As String
Dim blnParcelaVencida       As Boolean
    
Dim dblValorParcelasVencidas As Double
Dim dblValorParcelasAVencer  As Double

Dim dtmdtData               As Date

    strIPTU = ""
    On Error GoTo Problema_Na_Rotina
        
    blnParcelaChecada = False
    blnParcelaComAcordo = False
    blnParcelaVencida = False
    
    Set xadbLista = New XArrayDB
    xadbLista.Clear
    
    'Vamos obter os valores das parcelas da inscricao selecionada
    Set gobjBanco = New clsBanco
        
    'Vamos obter os Pkids das inscricoes para fazer consulta de acordos
    strSQL = "SELECT  LA.Pkid " & _
             "FROM " & gstrLancamentoAlfa & " LA " & _
             "WHERE LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "' AND " & _
             "(LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
             If dbcintComposicaoDaReceita.MatchedWithList Then
                 strSQL = strSQL & " AND LA.intComposicaoDaReceita = " & dbcintComposicaoDaReceita.BoundText
             ElseIf dbcintUtilizacao.MatchedWithList Then
                 strSQL = strSQL & " AND LA.intUtilizacao = " & dbcintUtilizacao.BoundText
             End If
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Do While Not adoResultado.EOF
                strAcordosParaConsulta = strAcordosParaConsulta & adoResultado("Pkid").Value & ","
                adoResultado.MoveNext
            Loop
            strAcordosParaConsulta = Mid(strAcordosParaConsulta, 1, Len(strAcordosParaConsulta) - 1)
        End If
    End If
    
ConsultarAcordos:

    'Vamos obter os acordos, caso exista, para exibir no grid Pai
    strSQL = "SELECT  LV.intLancamentoAlfaAcordo " & _
             "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA " & _
             "WHERE LV.intLancamentoAlfa = LA.pkid AND " & _
             "LA.Pkid IN (" & strAcordosParaConsulta & ") AND Not LV.intLancamentoAlfaAcordo Is Null " & _
             "GROUP BY LV.intLancamentoAlfaAcordo "
    
    If gobjBanco.CriaADO(strSQL, 15, adoResultado) Then
        If Not adoResultado.EOF Then
            strAcordosParaConsulta = Space$(0)
            Do While Not adoResultado.EOF
                strInscricoes = strInscricoes & adoResultado("intlancamentoalfaacordo").Value & ","
                strAcordosParaConsulta = strAcordosParaConsulta & adoResultado("intlancamentoalfaacordo").Value & ","
                adoResultado.MoveNext
            Loop
            strAcordosParaConsulta = Mid(strAcordosParaConsulta, 1, Len(strAcordosParaConsulta) - 1)
            GoTo ConsultarAcordos
        End If
    End If
    
    strSQL = "SELECT LV.Pkid PkidLV, LV.bitParcelaValida, LA.intExercicio, LV.intLancamentoAlfa, LV.intParcela, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.intLancamentoAlfaAcordo, LV.intLancamentoAlfaDAtiva, "
    strSQL = strSQL & "LA.Pkid PkidLA, LA.strInscricao, " & gstrCONVERT(cdt_numeric, "LA.strNumeroAviso") & " strNumeroAviso, LA.intComposicaoDaReceita, LA.strComposicaoDaReceita, " & strSUBSTRING & "(LAA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & " " & gstrRIGHT("LAA.strInscricao", 4) & " Acordo, LA.intUtilizacao, PA.intContaBancaria, "
    strSQL = strSQL & "(Select STRNUMDISTRIBUIDOR " & strCONCAT & " '/' " & strCONCAT & " STRSERIEDISTRIBUIDOR From tblexecutivo Where pkid = DA.intExecutivo) strExecutivo "
    strSQL = strSQL & "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA, " & gstrAcordo & " AC, " & gstrLancamentoAlfa & " LAA, " & gstrParametroAtualizacao & " PA, " & gstrDativa & " DA "
    strSQL = strSQL & "WHERE LV.intLancamentoAlfa = LA.pkid AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= AC.intLancamentoAlfa " & strOUTJOracle
    strSQL = strSQL & " AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= LAA.Pkid " & strOUTJOracle
    strSQL = strSQL & " AND PA.intComposicaoReceita = LA.intComposicaoDaReceita "
    strSQL = strSQL & " AND PA.intExercicio = LA.intExercicio "
    strSQL = strSQL & " AND LV.Pkid not in(Select Intlancamentovalor From " & gstrLancamentoPagamento & ")"
    strSQL = strSQL & " AND LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "' "
    strSQL = strSQL & " AND (LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
    strSQL = strSQL & " AND DA.intlancamentoAlfa " & strOUTJOracle & " =" & strOUTJSQLServer & " LA.Pkid "
    If dbcintComposicaoDaReceita.MatchedWithList Then
        strSQL = strSQL & " AND LA.intComposicaoDaReceita = " & dbcintComposicaoDaReceita.BoundText
    ElseIf dbcintUtilizacao.MatchedWithList Then
        strSQL = strSQL & " AND LA.intUtilizacao = " & dbcintUtilizacao.BoundText
    End If
             
    'Consulta que retorna os acordos
    If Len(strInscricoes) > 0 Then
        
        strInscricoes = Mid(strInscricoes, 1, Len(strInscricoes) - 1)
        
        strSQL = strSQL & " UNION ALL "
        strSQL = strSQL & "SELECT LV.Pkid PkidLV, LV.bitParcelaValida, LA.intExercicio, LV.intLancamentoAlfa, LV.intParcela, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.intLancamentoAlfaAcordo, LV.intLancamentoAlfaDAtiva, "
        strSQL = strSQL & "LA.Pkid PkidLA, LA.strInscricao, " & gstrCONVERT(cdt_numeric, "LA.strNumeroAviso") & " strNumeroAviso, LA.intComposicaoDaReceita, LA.strComposicaoDaReceita, " & strSUBSTRING & "(LAA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & " " & gstrRIGHT("LAA.strInscricao", 4) & " Acordo, LA.intUtilizacao, PA.intContaBancaria, "
        strSQL = strSQL & "(Select STRNUMDISTRIBUIDOR " & strCONCAT & " '/' " & strCONCAT & " STRSERIEDISTRIBUIDOR From tblexecutivo Where pkid = DA.intExecutivo) strExecutivo "
        strSQL = strSQL & "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA, " & gstrAcordo & " AC, " & gstrLancamentoAlfa & " LAA, " & gstrParametroAtualizacao & " PA, " & gstrDativa & " DA "
        strSQL = strSQL & "WHERE LV.intLancamentoAlfa = LA.pkid AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= AC.intLancamentoAlfa " & strOUTJOracle
        strSQL = strSQL & " AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= LAA.Pkid " & strOUTJOracle
        strSQL = strSQL & " AND PA.intComposicaoReceita = LA.intComposicaoDaReceita "
        strSQL = strSQL & " AND PA.intExercicio = LA.intExercicio "
        strSQL = strSQL & " AND DA.intlancamentoAlfa " & strOUTJOracle & " =" & strOUTJSQLServer & " LA.Pkid "
        strSQL = strSQL & " AND LV.Pkid not in(Select Intlancamentovalor From " & gstrLancamentoPagamento & ")"
        strSQL = strSQL & " AND LA.Pkid IN (" & strInscricoes & ") "
    
    End If
    
    If bytDBType = EDatabases.Oracle Then
       'strSql = strSql & " ORDER BY intLancamentoAlfa, intParcela"
       strSQL = strSQL & " ORDER BY strComposicaoDaReceita, intExercicio, strNumeroAviso, intLancamentoAlfa, intParcela"
    Else
       'strSql = strSql & " ORDER BY LV.intLancamentoAlfa, LV.intParcela"
       strSQL = strSQL & " ORDER BY LA.strComposicaoDaReceita, LA.intExercicio, LA.strNumeroAviso, LV.intLancamentoAlfa, LV.intParcela"
    End If
        
    If gobjBanco.CriaADO(strSQL, 20, adoResultado) Then
        
        If Not adoResultado.EOF Then
        
        With adoResultado
            
            ReDim vetParcelas(19, 0)
            ReDim vetParcelasOpcionais(19, 0)
            
            If Not adoResultado.EOF Then
            
                lngUltInscricao = !PkidLA
                
                'xadbLista.ReDim 0, !TotalAgrup - 1, 0, 18
                
                intForOpcionais = 0
                
                'Vamos calcular os valores de cada parcela
                For intFor = 0 To adoResultado.RecordCount - 1
                    
                    dtmdtData = VerificaDiasNaoUteis(!Dtmdtvencimento)
                    
                    strSQL = gstrStoredProcedure("sp_AtualizaParcela", !intComposicaoDaReceita & ", " & !intExercicio & ", " & !intParcela & ", " & gstrConvDtParaSql(dtmdtData) & ", " & gstrConvDtParaSql(dtmDataAtualizacao) & ", " & gstrConvVrParaSql(!ValorOrig) & ", " & !intMoeda, True)

                    Set gobjBanco = New clsBanco
                    
                    If gobjBanco.CriaADO(strSQL, 80, adoParcelas) Then
                    'If gobjBanco.ExecuteStoredProcedure(strSql, 80, adoParcelas) Then
                    
                        'Vamos alimentar o array considerando o grid a ser exibido
                        If adoResultado("bitParcelaValida").Value = True Or adoResultado("bitParcelaValida").Value = 1 Then
                        
                            ReDim Preserve vetParcelas(19, intFor - intForOpcionais)
                            
                            vetParcelas(0, intFor - intForOpcionais) = Space$(0) & adoResultado("PkidLV").Value
                            vetParcelas(1, intFor - intForOpcionais) = Space$(0) & adoResultado("intLancamentoAlfa").Value
                            vetParcelas(2, intFor - intForOpcionais) = Space$(0) & adoResultado("intParcela").Value
                            vetParcelas(3, intFor - intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoResultado("ValorOrig").Value))
                            vetParcelas(4, intFor - intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                            vetParcelas(5, intFor - intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value))
                            vetParcelas(6, intFor - intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value))
                            vetParcelas(7, intFor - intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                            vetParcelas(8, intFor - intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                            vetParcelas(9, intFor - intForOpcionais) = Space$(0) & adoResultado("Acordo").Value
                            vetParcelas(10, intFor - intForOpcionais) = Not IsNull(adoResultado("strExecutivo").Value) 'Space$(0) & RetornaExecutivo(adoResultado("intLancamentoAlfaDAtiva").Value)
                            vetParcelas(11, intFor - intForOpcionais) = IsNull(adoResultado("intLancamentoAlfaAcordo").Value)
                            vetParcelas(12, intFor - intForOpcionais) = Space$(0) & gstrDataFormatada(adoResultado("dtmDtVencimento").Value)
                            vetParcelas(13, intFor - intForOpcionais) = Space$(0) & gstrFormataInscricao(Right(adoResultado("strInscricao").Value, gintRetornaTamanhoMascara(adoResultado("intUtilizacao").Value)), adoResultado("intUtilizacao").Value)
                            vetParcelas(14, intFor - intForOpcionais) = Space$(0) & adoResultado("strNumeroAviso").Value
                            vetParcelas(15, intFor - intForOpcionais) = Space$(0) & adoResultado("intExercicio").Value
                            vetParcelas(16, intFor - intForOpcionais) = Space$(0) & adoResultado("strComposicaoDaReceita").Value
                            vetParcelas(17, intFor - intForOpcionais) = Space$(0) & adoResultado("strInscricao").Value
                            vetParcelas(18, intFor - intForOpcionais) = Space$(0) & adoResultado("intUtilizacao").Value
                            vetParcelas(19, intFor - intForOpcionais) = Space$(0) & IIf(IsNull(adoResultado("intLancamentoAlfaDAtiva").Value), "Não", "Sim")
                            
                            If CDate(adoResultado("dtmDtVencimento").Value) < CDate(gstrDataDoSistema) Then
                                
                                intQtdeParcelasInicial = intQtdeParcelasInicial + 1
                                
                                dblValorParcelasInicial = dblValorParcelasInicial + (CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)))
                                   
                                If IsNull(adoResultado("intLancamentoAlfaAcordo").Value) Then
                                    dblValorParcelasVencidas = dblValorParcelasVencidas + (CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)))
                                End If
                                
                                If Len(strVencto1Parcela) = 0 Then strVencto1Parcela = gstrDataFormatada(adoResultado("dtmDtVencimento").Value)
                                     
                                blnParcelaVencida = True
                            
                            Else
                                
                                If IsNull(adoResultado("intLancamentoAlfaAcordo").Value) Then
                                    dblValorParcelasAVencer = dblValorParcelasAVencer + (CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)))
                                End If
                                
                            End If
                            
                            'Caso esteja checado vamos somar
                            If vetParcelas(11, intFor - intForOpcionais) = True Then
                                dblTotalOriginal = dblTotalOriginal + CCur(gstrConvVrDoSql(adoResultado("ValorOrig").Value))
                                dblTotalPrincipal = dblTotalPrincipal + CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                                dblTotalMulta = dblTotalMulta + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value))
                                dblTotalJuros = dblTotalJuros + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value))
                                dblTotalCorrecao = dblTotalCorrecao + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                                
                                'So vamos incluir parcelas vencidas
'                                If CDate(adoResultado("dtmDtVencimento").Value) < CDate(gstrDataDoSistema) Then
                                    
'                                    intQtdeParcelasInicial = intQtdeParcelasInicial + 1
                                    
'                                    dblValorParcelasInicial = dblValorParcelasInicial + (CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)))
                                    
'                                    If Len(strVencto1Parcela) = 0 Then strVencto1Parcela = gstrDataFormatada(adoResultado("dtmDtVencimento").Value)
                                      
'                                End If
                                
                                blnParcelaChecada = True
                            Else
                                blnParcelaComAcordo = True
                                'Namo vamos inserir o acordo se ele ja estiver na string
                                If InStr(1, strAcordosRelacionados, ", " & adoResultado("Acordo").Value) = 0 Then
                                    strAcordosRelacionados = strAcordosRelacionados & ", " & adoResultado("Acordo").Value
                                End If
                            End If
                            
                        Else
                            
                            ReDim Preserve vetParcelasOpcionais(19, intForOpcionais)
                            
                            vetParcelasOpcionais(0, intForOpcionais) = Space$(0) & adoResultado("PkidLV").Value
                            vetParcelasOpcionais(1, intForOpcionais) = Space$(0) & adoResultado("intLancamentoAlfa").Value
                            vetParcelasOpcionais(2, intForOpcionais) = Space$(0) & adoResultado("intParcela").Value
                            vetParcelasOpcionais(3, intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoResultado("ValorOrig").Value))
                            vetParcelasOpcionais(4, intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                            vetParcelasOpcionais(5, intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value))
                            vetParcelasOpcionais(6, intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value))
                            vetParcelasOpcionais(7, intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                            vetParcelasOpcionais(8, intForOpcionais) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                            vetParcelasOpcionais(9, intForOpcionais) = Space$(0) & adoResultado("Acordo").Value
                            vetParcelasOpcionais(10, intForOpcionais) = False
                            vetParcelasOpcionais(11, intForOpcionais) = 0
                            vetParcelasOpcionais(12, intForOpcionais) = Space$(0) & adoResultado("dtmDtVencimento").Value
                            vetParcelasOpcionais(13, intForOpcionais) = Space$(0) & gstrFormataInscricao(Right(adoResultado("strInscricao").Value, gintRetornaTamanhoMascara(adoResultado("intUtilizacao").Value)), adoResultado("intUtilizacao").Value)
                            vetParcelasOpcionais(14, intForOpcionais) = Space$(0) & adoResultado("strNumeroAviso").Value
                            vetParcelasOpcionais(15, intForOpcionais) = Space$(0) & adoResultado("intExercicio").Value
                            vetParcelasOpcionais(16, intForOpcionais) = Space$(0) & adoResultado("strComposicaoDaReceita").Value
                            vetParcelasOpcionais(17, intForOpcionais) = Space$(0) & adoResultado("strInscricao").Value
                            vetParcelasOpcionais(18, intForOpcionais) = Space$(0) & adoResultado("intUtilizacao").Value
                            vetParcelasOpcionais(19, intForOpcionais) = Space$(0) & IIf(IsNull(adoResultado("intLancamentoAlfaDAtiva").Value), "Não", "Sim")
                            intForOpcionais = intForOpcionais + 1
                        End If
                        
                    Else
                        LimpaGrids
                        Exit Sub
                    End If
                    
                    'Vamos obter o proximo registro para verificacao do termino do agrupamento
                    adoResultado.MoveNext
                    If Not adoResultado.EOF Then
                        lngProxInscricao = adoResultado("PkidLA").Value
                    Else
                        lngProxInscricao = -1
                    End If
                    
                    adoResultado.MovePrevious
                    
                    'Vamos preencher o grid acumulado com cada agrupamento totalizado
                    If lngUltInscricao <> lngProxInscricao Then
                        
                        xadbLista.ReDim 0, intPosition, 0, 23
                        
                        lngUltInscricao = lngProxInscricao
                    
                        varAux = Space$(0) & !PkidLA
                        xadbLista(intPosition, 0) = varAux
                        xadbLista(intPosition, 1) = -1
                        varAux = Space$(0) & gstrFormataInscricao(Right(adoResultado("strInscricao").Value, gintRetornaTamanhoMascara(adoResultado("intUtilizacao").Value)), adoResultado("intUtilizacao").Value)
                        xadbLista(intPosition, 2) = varAux
                        varAux = Space$(0) & !intComposicaoDaReceita
                        xadbLista(intPosition, 3) = varAux
                        varAux = Space$(0) & !strComposicaoDaReceita
                        strIPTU = strIPTU & IIf(strIPTU = "", "", " / ") & gstrENulo(!strComposicaoDaReceita)
                        xadbLista(intPosition, 4) = varAux
                        varAux = Space$(0) & !intExercicio
                        xadbLista(intPosition, 5) = varAux
                        varAux = dblTotalOriginal
                        xadbLista(intPosition, 6) = varAux
                        varAux = dblTotalPrincipal
                        xadbLista(intPosition, 7) = varAux
                        varAux = dblTotalMulta
                        xadbLista(intPosition, 8) = varAux
                        varAux = dblTotalJuros
                        xadbLista(intPosition, 9) = varAux
                        varAux = dblTotalCorrecao
                        xadbLista(intPosition, 10) = varAux
                        varAux = dblTotalPrincipal + dblTotalMulta + dblTotalJuros + dblTotalCorrecao
                        xadbLista(intPosition, 11) = varAux
                        
                        varAux = dblCalculaEncargos(BIT_HONORARIOS, (dblTotalPrincipal + dblTotalMulta + dblTotalJuros + dblTotalCorrecao), Space$(0) & !strExecutivo)
                        xadbLista(intPosition, 12) = varAux
                        varAux = 0
                        xadbLista(intPosition, 14) = varAux
                        varAux = 0
                        xadbLista(intPosition, 13) = varAux
                        varAux = dblTotalPrincipal + dblTotalMulta + dblTotalJuros + dblTotalCorrecao + xadbLista(intPosition, 12) + xadbLista(intPosition, 13) + xadbLista(intPosition, 14)
                        xadbLista(intPosition, 15) = varAux
                        
                        varAux = Space$(0) & blnParcelaComAcordo
                        xadbLista(intPosition, 16) = varAux
                        varAux = Space$(0) & !strExecutivo
                        xadbLista(intPosition, 17) = varAux
                        varAux = intQtdeParcelasInicial
                        xadbLista(intPosition, 18) = varAux
                        varAux = strVencto1Parcela
                        xadbLista(intPosition, 19) = varAux
                        varAux = dblValorParcelasInicial
                        xadbLista(intPosition, 20) = varAux
                        If Len(strAcordosRelacionados) > 0 Then
                            varAux = Mid(strAcordosRelacionados, 3, Len(strAcordosRelacionados))
                        Else
                            varAux = strAcordosRelacionados
                        End If
                        xadbLista(intPosition, 21) = varAux
                        varAux = !strNumeroAviso
                        xadbLista(intPosition, 22) = varAux
                        varAux = !intContaBancaria
                        xadbLista(intPosition, 23) = varAux
                        
                        vetTotais(0, 0) = IIf(Len(vetTotais(0, 0)) > 0, vetTotais(0, 0), 0) + dblTotalOriginal
                        vetTotais(0, 1) = IIf(Len(vetTotais(0, 1)) > 0, vetTotais(0, 1), 0) + dblTotalPrincipal
                        vetTotais(0, 2) = IIf(Len(vetTotais(0, 2)) > 0, vetTotais(0, 2), 0) + dblTotalMulta
                        vetTotais(0, 3) = IIf(Len(vetTotais(0, 3)) > 0, vetTotais(0, 3), 0) + dblTotalJuros
                        vetTotais(0, 4) = IIf(Len(vetTotais(0, 4)) > 0, vetTotais(0, 4), 0) + dblTotalCorrecao
                        vetTotais(0, 5) = IIf(Len(vetTotais(0, 5)) > 0, vetTotais(0, 5), 0) + dblTotalPrincipal + dblTotalMulta + dblTotalJuros + dblTotalCorrecao
                        vetTotais(0, 6) = IIf(Len(vetTotais(0, 6)) > 0, vetTotais(0, 6), 0) + xadbLista(intPosition, 12)
                        vetTotais(0, 9) = IIf(Len(vetTotais(0, 9)) > 0, vetTotais(0, 9), 0) + xadbLista(intPosition, 15)
                        
                        dblTotalOriginal = 0
                        dblTotalPrincipal = 0
                        dblTotalMulta = 0
                        dblTotalJuros = 0
                        dblTotalCorrecao = 0
                        
                        intPosition = intPosition + 1
                        
                        intQtdeParcelasInicial = 0
                        dblValorParcelasInicial = 0
                        strVencto1Parcela = ""
                        strAcordosRelacionados = ""
                        
                        blnParcelaComAcordo = False
                        
                    End If
                    adoResultado.MoveNext
                Next
                
            Else
                LimpaGrids
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrGuiaCertidaoNegativa
                Exit Sub
            End If
            'Vamos habilitar a certidao positiva quando o valor das parcelas vencidas for > 0 (zero)
            blnGuiaPositiva = dblValorParcelasVencidas > 0
            HabilitaDesabilitaBotao1 blnGuiaPositiva, gstrBtnArquivo, gstrGuiaCertidaoPositiva
        End With
        
        End If
        
        Set tdbValoresAcumulado.Array = xadbLista
        tdbValoresAcumulado.ReBind
        tdbValoresAcumulado.Refresh
        
        lbl_dblTotOriginal = gstrConvVrDoSql(vetTotais(0, 0), 2, , True)
        lbl_dblTotPrincipal = gstrConvVrDoSql(vetTotais(0, 1), 2, , True)
        lbl_dblTotMulta = gstrConvVrDoSql(vetTotais(0, 2), 2, , True)
        lbl_dblTotJuros = gstrConvVrDoSql(vetTotais(0, 3), 2, , True)
        lbl_dblTotCorrecao = gstrConvVrDoSql(vetTotais(0, 4), 2, , True)
        lbl_dblTotTotal = gstrConvVrDoSql(vetTotais(0, 5), 2, , True)
        lbl_dblTotHonorarios = gstrConvVrDoSql(vetTotais(0, 6), 2, , True)
        lbl_dblTotGeral = gstrConvVrDoSql(vetTotais(0, 9), 2, , True)
        
        'Caso parametrizado, vamos realizar a somatoria de Acordos pagos cancelados por inadimplencia
        If blnAcordoCreditaParcela Then
        strSQL = " SELECT SUM(ValorPago) ValorPago FROM (SELECT (" & gstrISNULL("sum(LP.dblValorPrincipal)", "0", "sum(LP.dblValorPrincipal)") & " / PA.dblValor) * PA1.dblValor ValorPago " & _
                 " FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA, " & gstrLancamentoPagamento & " LP, " & gstrAcordo & " AC, " & gstrParametroAtualizacao & " PA, " & gstrParametroAtualizacao & " PA1 " & _
                 " WHERE LV.intLancamentoAlfa = LA.pkid AND " & _
                       " LV.Pkid = LP.intLancamentoValor AND " & _
                       " AC.intLancamentoAlfa = LA.Pkid AND " & _
                       " AC.dtmDtCancelamento is not NUll AND " & _
                       " AC.dtmDtUtilizacao is NUll AND " & _
                       " Year(LP.dtmDtPagamento) = PA.intExercicio AND " & _
                       " PA.intComposicaoReceita = LA.intComposicaoDaReceita AND " & _
                       " PA1.intExercicio = " & Year(gstrDataDoSistema) & " AND " & _
                       " PA1.intComposicaoReceita = LA.intComposicaoDaReceita AND " & _
                       " LA.Pkid IN (" & strInscricoes & ") " & _
                       " GROUP BY year(LP.dtmDtPagamento), PA.dblValor, PA1.dblValor) PG "
            If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                
                If Not adoResultado.EOF Then
                    txtdblCredito = gstrConvVrDoSql(adoResultado("ValorPago").Value, 2, , True)
                Else
                    txtdblCredito = gstrConvVrDoSql(0, 2, , True)
                End If
                
            End If
        Else
            txtdblCredito = ""
        End If
        
    End If
    
    HabilitaDesabilitaBotao1 blnParcelaChecada, gstrBtnArquivo, gstrImprimirGuia, gstrParcelamentoDebitoAtualizado
    
    'Vamos habilitar a certidao negativa quando o valor das parcelas vencidas for 0 (zero)
    HabilitaDesabilitaBotao1 dblValorParcelasVencidas = 0, gstrBtnArquivo, gstrGuiaCertidaoNegativa
    'Vamos habilitar a positiva com efeito negativa quando o valor das parcelas vencidas for 0 (zero)
    'e o valor das parcelas a vencer for > 0 (zero)
    HabilitaDesabilitaBotao1 dblValorParcelasVencidas = 0 And dblValorParcelasAVencer > 0, gstrBtnArquivo, gstrGuiaCertidaoPositivaEfeitoNegativa
    
    Exit Sub
    
Problema_Na_Rotina:
    ExibeDetalheErro Err.Description
    Exit Sub
    
End Sub

Private Sub ExibeParcelas(PkidAlfa As Long)
Dim intFor            As Integer
Dim varAux            As Variant
Dim intPosition       As Integer
    
    Set xadbParcelas = New XArrayDB
    xadbParcelas.Clear
    
    If Len(vetParcelas(0, 0)) > 0 Then
        
        For intFor = 0 To UBound(vetParcelas, 2)
            
            If vetParcelas(1, intFor) = PkidAlfa Then
                
                xadbParcelas.ReDim 0, intPosition, 0, 13
                
                varAux = vetParcelas(0, intFor)
                xadbParcelas(intPosition, 0) = varAux
                xadbParcelas(intPosition, 1) = vetParcelas(11, intFor)
                varAux = vetParcelas(1, intFor)
                xadbParcelas(intPosition, 2) = varAux
                varAux = vetParcelas(2, intFor)
                xadbParcelas(intPosition, 3) = varAux
                varAux = vetParcelas(3, intFor)
                xadbParcelas(intPosition, 4) = varAux
                varAux = vetParcelas(4, intFor)
                xadbParcelas(intPosition, 5) = varAux
                varAux = vetParcelas(5, intFor)
                xadbParcelas(intPosition, 6) = varAux
                varAux = vetParcelas(6, intFor)
                xadbParcelas(intPosition, 7) = varAux
                varAux = vetParcelas(7, intFor)
                xadbParcelas(intPosition, 8) = varAux
                varAux = vetParcelas(8, intFor)
                xadbParcelas(intPosition, 9) = varAux
                varAux = vetParcelas(9, intFor)
                xadbParcelas(intPosition, 10) = varAux
                varAux = vetParcelas(10, intFor)
                xadbParcelas(intPosition, 11) = varAux
                varAux = vetParcelas(12, intFor)
                xadbParcelas(intPosition, 12) = varAux
                varAux = vetParcelas(19, intFor)
                xadbParcelas(intPosition, 13) = varAux
                intPosition = intPosition + 1
            End If
        Next
    Else
        'xadbParcelas.ReDim 0, 0, 0, 13
        'xadbParcelas(0, 0) = ""
        'xadbParcelas(0, 1) = ""
        'xadbParcelas(0, 2) = ""
        'xadbParcelas(0, 3) = ""
        'xadbParcelas(0, 4) = ""
        'xadbParcelas(0, 5) = ""
        'xadbParcelas(0, 6) = ""
        'xadbParcelas(0, 7) = ""
        'xadbParcelas(0, 8) = ""
        'xadbParcelas(0, 9) = ""
        'xadbParcelas(0, 10) = ""
        'xadbParcelas(0, 11) = ""
        'xadbParcelas(0, 12) = ""
        'xadbParcelas(0, 13) = ""
    End If
    
    Set tdbValoresParcelas.Array = xadbParcelas
    tdbValoresParcelas.ReBind
    tdbValoresParcelas.Refresh
    
End Sub

Private Sub ExibeParcelasOpcionais(PkidAlfa As Long)
Dim intFor            As Integer
Dim varAux            As Variant
Dim intPosition       As Integer
    
    Set xadbParcelasOpcionais = New XArrayDB
    xadbParcelasOpcionais.Clear
    
    If Len(vetParcelasOpcionais(0, 0)) > 0 Then
        
        For intFor = 0 To UBound(vetParcelasOpcionais, 2)
            
            If vetParcelasOpcionais(1, intFor) = PkidAlfa Then
                
                xadbParcelasOpcionais.ReDim 0, intPosition, 0, 13
                
                varAux = vetParcelasOpcionais(0, intFor)
                xadbParcelasOpcionais(intPosition, 0) = varAux
                xadbParcelasOpcionais(intPosition, 1) = vetParcelasOpcionais(11, intFor)
                varAux = vetParcelasOpcionais(1, intFor)
                xadbParcelasOpcionais(intPosition, 2) = varAux
                varAux = vetParcelasOpcionais(2, intFor)
                xadbParcelasOpcionais(intPosition, 3) = varAux
                varAux = vetParcelasOpcionais(3, intFor)
                xadbParcelasOpcionais(intPosition, 4) = varAux
                varAux = vetParcelasOpcionais(4, intFor)
                xadbParcelasOpcionais(intPosition, 5) = varAux
                varAux = vetParcelasOpcionais(5, intFor)
                xadbParcelasOpcionais(intPosition, 6) = varAux
                varAux = vetParcelasOpcionais(6, intFor)
                xadbParcelasOpcionais(intPosition, 7) = varAux
                varAux = vetParcelasOpcionais(7, intFor)
                xadbParcelasOpcionais(intPosition, 8) = varAux
                varAux = vetParcelasOpcionais(8, intFor)
                xadbParcelasOpcionais(intPosition, 9) = varAux
                varAux = vetParcelasOpcionais(9, intFor)
                xadbParcelasOpcionais(intPosition, 10) = varAux
                varAux = vetParcelasOpcionais(10, intFor)
                xadbParcelasOpcionais(intPosition, 11) = varAux
                varAux = vetParcelasOpcionais(12, intFor)
                xadbParcelasOpcionais(intPosition, 12) = varAux
                varAux = vetParcelasOpcionais(19, intFor)
                xadbParcelasOpcionais(intPosition, 13) = varAux
                intPosition = intPosition + 1
            End If
        Next
    Else
        'xadbParcelasOpcionais.ReDim 0, 0, 0, 13
        'xadbParcelasOpcionais(0, 0) = ""
        'xadbParcelasOpcionais(0, 1) = ""
        'xadbParcelasOpcionais(0, 2) = ""
        'xadbParcelasOpcionais(0, 3) = ""
        'xadbParcelasOpcionais(0, 4) = ""
        'xadbParcelasOpcionais(0, 5) = ""
        'xadbParcelasOpcionais(0, 6) = ""
        'xadbParcelasOpcionais(0, 7) = ""
        'xadbParcelasOpcionais(0, 8) = ""
        'xadbParcelasOpcionais(0, 9) = ""
        'xadbParcelasOpcionais(0, 10) = ""
        'xadbParcelasOpcionais(0, 11) = ""
        'xadbParcelasOpcionais(0, 12) = ""
        'xadbParcelasOpcionais(0, 13) = ""
    End If
    
    Set tdbValoresParcelasOpcionais.Array = xadbParcelasOpcionais
    tdbValoresParcelasOpcionais.ReBind
    tdbValoresParcelasOpcionais.Refresh
    
End Sub

Public Sub ImprimirGuia()
    Dim adoResultado        As ADODB.Recordset
    Dim strSQL              As String
    Dim VetWordGuia()       As String
    Dim VetWordParcelas()   As String
    Dim contLista           As Integer
    Dim contParcela         As Integer
    Dim contGuias           As Integer
    Dim intContador         As Integer
    
    'Variaveis para Anistia
    Dim dblPrincipal        As Double
    Dim dblMulta            As Double
    Dim dblJuros            As Double
    Dim blnAnistia          As Boolean
    
    ReDim VetWordGuia(12, 1)
    ReDim VetWordParcelas(13, 0)
        
    Screen.MousePointer = vbHourglass
    blnAnistia = False
    
    intContador = 1
    contGuias = 0
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strQueryAnistia, 5, adoResultado) Then 'Consulta para verificação de anistia no periodo
        If Not adoResultado.EOF Then
            blnAnistia = True
            dblPrincipal = CDbl(gstrConvVrDoSql(gstrENulo(adoResultado!Dblvalororiginal), , , True))
            dblMulta = CDbl(gstrConvVrDoSql(gstrENulo(adoResultado!dblMulta), , , True))
            dblJuros = CDbl(gstrConvVrDoSql(gstrENulo(adoResultado!dblJuros), , , True))
        End If
    End If
    
    'Vamos ordenar o grid por Conta, para podermos agrupar no array
    xadbLista.QuickSort 0, xadbLista.UpperBound(1), 23, XORDER_ASCEND, XTYPE_DATE

    For contLista = 0 To xadbLista.Count(1) - 1

        If xadbLista(contLista, 1) = -1 Then
            
            'Na primeira vez atribuiremo o valor 1
            If contGuias = 0 Then
                contGuias = 1
            'Caso seja F. Compensacao, vamos criar mais 1 registro no array
            ElseIf Not IsNull(xadbLista(contLista, 23)) Then
                'Caso seja contas diferentes, senao agruparemos
                If xadbLista(contLista, 23) <> VetWordGuia(12, contGuias) Then
                    contGuias = contGuias + 1
                    ReDim Preserve VetWordGuia(12, contGuias)
                End If
            End If
    
            VetWordGuia(0, contGuias) = strProprietario   ' Proprietario
            VetWordGuia(1, contGuias) = strLogradouro     ' Logradouro
            VetWordGuia(2, contGuias) = STRBAIRRO         ' Bairro
            VetWordGuia(3, contGuias) = ""                ' Limpa o número do aviso.
            'VetWordGuia(3, contGuias) = strNumeroAviso
        
            strSQL = ""
            strSQL = strSQL & "Select "
            strSQL = strSQL & gstrCONVERT(cdt_numeric, "LA.Strnumeroaviso") & ","
            strSQL = strSQL & "CR.Strsigla, "
            strSQL = strSQL & "strNumeroAviso "
            strSQL = strSQL & "From "
            strSQL = strSQL & gstrLancamentoAlfa & " LA, "
            strSQL = strSQL & gstrComposicaoDaReceita & " CR "
            strSQL = strSQL & "Where "
            strSQL = strSQL & "CR.Pkid = LA.Intcomposicaodareceita AND "
            strSQL = strSQL & "LA.Pkid = " & xadbLista(contLista, 0)  'Pkid tblLancamentoAlfa
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                'Composicao/Ano
                VetWordGuia(4, contGuias) = VetWordGuia(4, contGuias) & IIf(VetWordGuia(4, contGuias) = "", " ", " ; ") & gstrENulo(adoResultado!strsigla) & "-" & xadbLista(contLista, 5)
            End If
            VetWordGuia(3, contGuias) = VetWordGuia(3, contGuias) & adoResultado!strNumeroAviso & ", "
            VetWordGuia(5, contGuias) = CDbl(gstrConvVrDoSql(IIf(VetWordGuia(5, contGuias) <> "", VetWordGuia(5, contGuias), 0), 2)) + CDbl(gstrConvVrDoSql(xadbLista(contLista, 7), 2)) 'Principal
            VetWordGuia(6, contGuias) = CDbl(gstrConvVrDoSql(IIf(VetWordGuia(6, contGuias) <> "", VetWordGuia(6, contGuias), 0), 2)) + CDbl(gstrConvVrDoSql(xadbLista(contLista, 8), 2))  'Multa
            VetWordGuia(7, contGuias) = CDbl(gstrConvVrDoSql(IIf(VetWordGuia(7, contGuias) <> "", VetWordGuia(7, contGuias), 0), 2)) + CDbl(gstrConvVrDoSql(xadbLista(contLista, 9), 2))  'Juros
            VetWordGuia(8, contGuias) = CDbl(gstrConvVrDoSql(IIf(VetWordGuia(8, contGuias) <> "", VetWordGuia(8, contGuias), 0), 2)) + CDbl(gstrConvVrDoSql(xadbLista(contLista, 10), 2))  'Correcao
            VetWordGuia(9, contGuias) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(IIf(VetWordGuia(9, contGuias) <> "", VetWordGuia(9, contGuias), 0), 2)) + CDbl(gstrConvVrDoSql(xadbLista(contLista, 11), 2)), 2) 'Total
            VetWordGuia(10, contGuias) = VetWordGuia(10, contGuias) & IIf(Trim(VetWordGuia(10, contGuias)) = "", "", " ; ") & xadbLista(contLista, 5) & " : "
            VetWordGuia(11, contGuias) = xadbLista(contLista, 2) 'Inscricao
            VetWordGuia(12, contGuias) = Space$(0) & xadbLista(contLista, 23) 'Conta Corrente
            
            For contParcela = 0 To UBound(vetParcelas, 2) 'Parcelas
                If xadbLista(contLista, 0) = vetParcelas(1, contParcela) Then
                    If vetParcelas(11, contParcela) = True Then 'Or vetParcelas(11, contParcela) = -1 Then
                        ReDim Preserve VetWordParcelas(13, UBound(VetWordParcelas, 2) + 1)
                        
                        VetWordGuia(10, contGuias) = VetWordGuia(10, contGuias) & vetParcelas(2, contParcela) & ", "
                        VetWordParcelas(0, intContador) = vetParcelas(0, contParcela)  'Pkid tblLancamentoValor
                        If blnAnistia Then
                            VetWordParcelas(1, intContador) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(vetParcelas(4, contParcela), 2, , True)) * (100 - dblPrincipal) / 100, , , True) 'Principal
                            VetWordParcelas(2, intContador) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(vetParcelas(5, contParcela), 2, , True)) * (100 - dblMulta) / 100, , , True) 'Multa
                            VetWordParcelas(3, intContador) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(vetParcelas(6, contParcela), 2, , True)) * (100 - dblJuros) / 100, , , True) 'Juros
                        Else
                            VetWordParcelas(1, intContador) = gstrConvVrDoSql(vetParcelas(4, contParcela), 2)  'Principal
                            VetWordParcelas(2, intContador) = gstrConvVrDoSql(vetParcelas(5, contParcela), 2)  'Multa
                            VetWordParcelas(3, intContador) = gstrConvVrDoSql(vetParcelas(6, contParcela), 2)  'Juros
                        End If
                        VetWordParcelas(4, intContador) = gstrConvVrDoSql(vetParcelas(7, contParcela), 2)  'Correcao
                        
                        VetWordParcelas(13, intContador) = Space$(0) & xadbLista(contLista, 23) 'Conta Corrente
                        
                        intContador = intContador + 1
                    End If
                End If
            Next
            
            For contParcela = 0 To UBound(vetParcelasOpcionais, 2) 'Parcelas Opcionais
                If xadbLista(contLista, 0) = vetParcelasOpcionais(1, contParcela) Then
                    If vetParcelasOpcionais(11, contParcela) = True Then 'Or vetParcelas(11, contParcela) = -1 Then
                        ReDim Preserve VetWordParcelas(13, UBound(VetWordParcelas, 2) + 1)
                        
                        VetWordGuia(10, contGuias) = VetWordGuia(10, contGuias) & vetParcelasOpcionais(2, contParcela) & " ,"
                        VetWordParcelas(0, intContador) = vetParcelasOpcionais(0, contParcela)  'Pkid tblLancamentoValor
                        VetWordParcelas(1, intContador) = gstrConvVrDoSql(vetParcelasOpcionais(4, contParcela), 2)  'Principal
                        VetWordParcelas(2, intContador) = gstrConvVrDoSql(vetParcelasOpcionais(5, contParcela), 2)  'Multa
                        VetWordParcelas(3, intContador) = gstrConvVrDoSql(vetParcelasOpcionais(6, contParcela), 2)  'Juros
                        VetWordParcelas(4, intContador) = gstrConvVrDoSql(vetParcelasOpcionais(7, contParcela), 2)   'Correcao
                        
                        VetWordParcelas(13, intContador) = Space$(0) & xadbLista(contLista, 23) 'Conta Corrente
                        
                        intContador = intContador + 1
                    End If
                End If
            Next
            
            VetWordGuia(10, contGuias) = Mid(VetWordGuia(10, contGuias), 1, Len(VetWordGuia(10, contGuias)) - 2)
            
        End If
        
        'Vamos verificar se mudou de Conta
        If xadbLista.UpperBound(1) >= contLista + 1 Then
            'Caso seja Conta diferente
            If xadbLista(contLista, 23) <> IIf(IsNull(xadbLista(contLista + 1, 23)), 0, xadbLista(contLista + 1, 23)) Then
                
                If Len(VetWordGuia(3, contGuias)) > 0 Then 'Verificacao de array em branco - nenhum selecionado
                    VetWordGuia(3, contGuias) = Mid(VetWordGuia(3, contGuias), 1, Len(Trim(VetWordGuia(3, contGuias))) - 1)

                    If blnAnistia Then
                        VetWordGuia(5, contGuias) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(VetWordGuia(5, contGuias), 2, , True)) * (100 - dblPrincipal) / 100, , , True) 'Principal
                        VetWordGuia(6, contGuias) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(VetWordGuia(6, contGuias), 2, , True)) * (100 - dblMulta) / 100, , , True)     'Multa
                        VetWordGuia(7, contGuias) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(VetWordGuia(7, contGuias), 2, , True)) * (100 - dblJuros) / 100, , , True)     'Juros
                        VetWordGuia(9, contGuias) = gstrConvVrDoSql(CDbl(VetWordGuia(5, contGuias)) + CDbl(VetWordGuia(6, contGuias)) + CDbl(VetWordGuia(7, contGuias)) + CDbl(VetWordGuia(8, contGuias)), 2)
                    End If
                End If
                
                'Caso a proxima seja Febrabam, vamos descarregar as Fichas de Compensacao
                If IsNull(xadbLista(contLista + 1, 23)) Then
                    
                    If Len(VetWordGuia(3, 1)) > 0 Then 'Verificacao de array em branco - nenhum selecionado
                        If IsNull(xadbLista(contLista, 23)) Then 'Febraban
                            ImprimeGuiaFebraban VetWordGuia(), strQuadra, strLote, gstrDataFormatada(dtmVencimento), VetWordParcelas()
                        Else
                            ImprimeGuiaFichaCompensacao VetWordGuia(), strQuadra, strLote, gstrDataFormatada(dtmVencimento), VetWordParcelas()
                        End If
                    End If
                    
                    'Vamos zerar o array para armazenar a Febraban
                    ReDim VetWordGuia(12, 1)
                    contGuias = 1
                
                End If
                
            End If
        End If

    Next
       
    'Impressao da ultima Conta
    If Len(VetWordGuia(9, 1)) > 0 Then
        If Len(VetWordGuia(12, contGuias)) = 0 Then 'Febraban
            
            VetWordGuia(3, contGuias) = Mid(VetWordGuia(3, contGuias), 1, Len(Trim(VetWordGuia(3, contGuias))) - 1)

            If blnAnistia Then
                VetWordGuia(5, contGuias) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(VetWordGuia(5, contGuias), 2, , True)) * (100 - dblPrincipal) / 100, , , True) 'Principal
                VetWordGuia(6, contGuias) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(VetWordGuia(6, contGuias), 2, , True)) * (100 - dblMulta) / 100, , , True)     'Multa
                VetWordGuia(7, contGuias) = gstrConvVrDoSql(CDbl(gstrConvVrDoSql(VetWordGuia(7, contGuias), 2, , True)) * (100 - dblJuros) / 100, , , True)     'Juros
                VetWordGuia(9, contGuias) = gstrConvVrDoSql(CDbl(VetWordGuia(5, contGuias)) + CDbl(VetWordGuia(6, contGuias)) + CDbl(VetWordGuia(7, contGuias)) + CDbl(VetWordGuia(8, contGuias)), 2)
            End If
        
            ImprimeGuiaFebraban VetWordGuia(), strQuadra, strLote, gstrDataFormatada(dtmVencimento), VetWordParcelas()
        Else
            ImprimeGuiaFichaCompensacao VetWordGuia(), strQuadra, strLote, gstrDataFormatada(dtmVencimento), VetWordParcelas()
        End If
    End If
   Screen.MousePointer = vbDefault
   
End Sub

Private Function strQueryLogradouroImovel(bytTipoComposicao As Byte) As String
Dim strSQL As String
Dim adoRec As New ADODB.Recordset

    strQuadra = Space$(0)
    strLote = Space$(0)
    lngContribuinte = 0

    Select Case bytTipoComposicao
    
        Case Is = TYP_IMOBILIARIA, TYP_ISS_CONSTRUCAO, TYP_ECONOMICA
            
            If bytTipoComposicao = TYP_ECONOMICA Then
                strSQL = "SELECT intContribuinte FROM " & gstrEconomico & " WHERE strInscricaoCadastral = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "' "
            Else
                strSQL = "SELECT intContribuinte, strQuadra, strLote FROM " & gstrImobiliario & " WHERE strInscricao = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "' "
            End If
            
            If gobjBanco.CriaADO(strSQL, 5, adoRec) Then
            
                If bytTipoComposicao <> TYP_ECONOMICA Then
                    strQuadra = gstrENulo(adoRec!strQuadra)
                    strLote = gstrENulo(adoRec!strLote)
                End If
                
                lngContribuinte = Val(gstrENulo(adoRec!intContribuinte))
                
                strSQL = ""
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "BA.strDescricao AS strBairro, "
                strSQL = strSQL & " RTRIM(LTRIM(L.strDescricao)) "
                strSQL = strSQL & strCONCAT & gstrISNULL("TL.strSigla", "''", "', '")
                strSQL = strSQL & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''")
                strSQL = strSQL & strCONCAT & gstrISNULL("U.strDescricao", "' '", "', '")
                strSQL = strSQL & strCONCAT & gstrISNULL("U.strDescricao", "''") & ")) AS strLogradouro, "
                strSQL = strSQL & "CO.intNumero, "
                strSQL = strSQL & "L.Intcep, "
                strSQL = strSQL & "CO.strComplemento, "
                strSQL = strSQL & "(SELECT MU.strDescricao FROM tblMunicipio  MU WHERE BA.intMunicipio = MU.PKId ) AS strMunicipio, "
                strSQL = strSQL & "(SELECT UF.strSigla FROM " & gstrUF & " UF WHERE UF.PKId = "
                strSQL = strSQL & "(SELECT MU.intUF FROM tblMunicipio  MU WHERE BA.intMunicipio = MU.PKId )) AS strEstado, "
                strSQL = strSQL & "CO.strLogradouroC, "
                strSQL = strSQL & "CO.intNumeroC, "
                strSQL = strSQL & "CO.strBairroC, "
                strSQL = strSQL & "CO.IntcepC, "
                strSQL = strSQL & "CO.strComplementoC, "
                strSQL = strSQL & "(SELECT MU.strDescricao FROM " & gstrCidade & " MU WHERE MU.PKId = CO.intMunicipioC) strMunicipioC, "
                strSQL = strSQL & "(SELECT UF.strSigla FROM " & gstrUF & " UF WHERE UF.PKId = CO.intUFC) strEstadoC "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & gstrContribuinte & " CO, "
                strSQL = strSQL & gstrBairro & " BA, "
                strSQL = strSQL & gstrLogradouro & " L, "
                strSQL = strSQL & gstrTituloLogradouro & " U, "
                strSQL = strSQL & gstrTipoLogradouro & " TL "
                strSQL = strSQL & " WHERE"
                strSQL = strSQL & " L.pkid = CO.Intlogradouro "
                strSQL = strSQL & " AND L.intBairro " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle
                strSQL = strSQL & " AND L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle
                'strSql = strSql & " AND L.Dtmdtexclusao is null "
                strSQL = strSQL & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
                strSQL = strSQL & " AND CO.Pkid = " & Val(gstrENulo(adoRec!intContribuinte))
            End If
    
        Case Is = TYP_DIVIDA_ATIVA, TYP_ACORDO, TYP_PRECO_PUBLICO

            strSQL = ""
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "LA.strBairro, "
            strSQL = strSQL & "LA.strLogradouro, "
            strSQL = strSQL & "LA.strNumero intNumero, "
            strSQL = strSQL & "LA.Intcep, "
            strSQL = strSQL & "LA.strComplemento, "
            strSQL = strSQL & "LA. strMunicipio, "
            strSQL = strSQL & "LA.strUF AS strEstado, "
            strSQL = strSQL & "LA.strLogradouroC, "
            strSQL = strSQL & "LA.strNumeroC intNumeroC, "
            strSQL = strSQL & "LA.strBairroC, "
            strSQL = strSQL & "LA.IntcepC, "
            strSQL = strSQL & "LA.strComplementoC, "
            strSQL = strSQL & "LA.strMunicipioC, "
            strSQL = strSQL & "LA.strUFc strEstadoC "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrLancamentoAlfa & " LA "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "LA.Strinscricao = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "'"

    End Select
        
    strQueryLogradouroImovel = strSQL
    
End Function

Private Function strVerificaIPTU() As String
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "Select Distinct CR.Strdescricao From "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR "
    strSQL = strSQL & "Where CR.Pkid = LA.Intcomposicaodareceita AND "
    strSQL = strSQL & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "'"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Do While Not .EOF
                    strIPTU = strIPTU & IIf(strIPTU = "", "", " / ") & gstrENulo(!strDescricao)
                    .MoveNext
                Loop
            End If
        End With
    End If

    strVerificaIPTU = strIPTU
    
End Function

Public Sub ImprimiNegativa()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim strNumero       As String

    Screen.MousePointer = vbHourglass
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    strNumero = CLng(glngRetornaProximoNumeroGuia(gstrEmpresa, "intNumeroGuiaNegativa"))
    
    If Val(strNumero) = 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        gobjBanco.ExecutaCommitTrans
    End If
        
    'strSql = ""
    'strSql = strSql & "select (intNumeroGuiaNegativa + 1) AS Numero From " & gstrEmpresa
    'Set gobjBanco = New clsBanco
    'If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
    '    If Not adoResultado.EOF Then
    '        If Val(gstrENulo(adoResultado!Numero)) > "0" Then
    '            strNumero = gstrENulo(adoResultado!Numero)
    '        Else
    '            strNumero = "1"
    '        End If
    '        gobjBanco.Execute ("Update " & gstrEmpresa & " Set intNumeroGuiaNegativa = " & strNumero)
    '    Else
    '        Exit Sub
    '    End If
    'End If
    
    OpenWordDocumentCertidaoNegativa String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text), strNumero, strLogradouro, STRBAIRRO, strLote, strQuadra, "", txtstrContribuinte, strIPTU, gstrDataDoSistema, gstrDataDoSistema, strInscricaoAuxiliar, intUtilizacao
    
    Set gobjBanco = Nothing
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub ImprimiPositiva()
Dim strSQL              As String
Dim adoResultado        As ADODB.Recordset
Dim strAtividade        As String
Dim intForLista         As Integer
Dim intForParcela       As Integer
Dim dblTotal            As Double
Dim dblSubTotal         As Double
Dim XArrayTabela        As XArrayDB
Dim XArrayAlinhaColunas As XArrayDB
Dim strNumero           As String
Dim lngComposicao       As Long

    Screen.MousePointer = vbHourglass
    
    Set XArrayTabela = New XArrayDB
    Set XArrayAlinhaColunas = New XArrayDB
    
    XArrayTabela.Clear
    
    With XArrayAlinhaColunas 'Alinhamento
        .Clear
        .ReDim 0, 0, 0, 5
        .Value(0, 0) = WORDALIGNPARAGRAPHLEFT
        .Value(0, 1) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 2) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 3) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 4) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 5) = WORDALIGNPARAGRAPHRIGHT
    End With
    
    'cabecalho da tabela
    XArrayTabela.ReDim 0, 0, 0, 5
    XArrayTabela(XArrayTabela.UpperBound(1), 0) = "Inscrição"
    XArrayTabela(XArrayTabela.UpperBound(1), 1) = "Exercício"
    XArrayTabela(XArrayTabela.UpperBound(1), 2) = "Parcelas"
    XArrayTabela(XArrayTabela.UpperBound(1), 3) = "Total"
    XArrayTabela(XArrayTabela.UpperBound(1), 4) = "1º Vencimento"
    XArrayTabela(XArrayTabela.UpperBound(1), 5) = "Nº do Aviso"
    For intForLista = 0 To xadbLista.Count(1) - 1
        
        'So vamos considerar com valores maiores que zero
        If xadbLista(intForLista, 20) > 0 Then
        
            dblSubTotal = 0
            
            XArrayTabela.ReDim 0, XArrayTabela.UpperBound(1) + 1, 0, 5
            
            If lngComposicao <> xadbLista(intForLista, 3) Then
                'Titulo composicao
                XArrayTabela(XArrayTabela.UpperBound(1), 0) = FORMAT_NEGRITO & xadbLista(intForLista, 4)
                lngComposicao = xadbLista(intForLista, 3)
                
                XArrayTabela.ReDim 0, XArrayTabela.UpperBound(1) + 1, 0, 5
            End If
            
            XArrayTabela(XArrayTabela.UpperBound(1), 0) = "  " & xadbLista(intForLista, 2)  'Inscrição
            XArrayTabela(XArrayTabela.UpperBound(1), 1) = xadbLista(intForLista, 5)         'Exercício
            XArrayTabela(XArrayTabela.UpperBound(1), 2) = xadbLista(intForLista, 18)        'Parcelas
            
            For intForParcela = 0 To UBound(vetParcelas, 2)
                'Só vamos somar no valor total parcelas que nao estejam em acordo
                If xadbLista(intForLista, 0) = vetParcelas(1, intForParcela) And (vetParcelas(9, intForParcela) = "/" Or vetParcelas(9, intForParcela) = "") Then
                    'E só parcelas vencidas
                    If CDate(vetParcelas(12, intForParcela)) < CDate(gstrDataDoSistema) Then
                        dblSubTotal = dblSubTotal + gstrConvVrDoSql(vetParcelas(8, intForParcela), 2)
                    End If
                End If
            Next
    
            XArrayTabela(XArrayTabela.UpperBound(1), 3) = gstrConvVrDoSql(xadbLista(intForLista, 20), 2) 'Total
            XArrayTabela(XArrayTabela.UpperBound(1), 4) = xadbLista(intForLista, 19)        '1º Vencimento
            XArrayTabela(XArrayTabela.UpperBound(1), 5) = xadbLista(intForLista, 22)        'Nº Aviso
            
            'Acordos
            If Len(xadbLista(intForLista, 21)) > 0 Then
                XArrayTabela.ReDim 0, XArrayTabela.UpperBound(1) + 1, 0, 5
                XArrayTabela(XArrayTabela.UpperBound(1), 0) = "   Acordo(s):" & xadbLista(intForLista, 21)
            End If
            
            dblTotal = dblTotal + gstrConvVrDoSql(dblSubTotal, 2)
            
        End If
        
    Next
    
    'Vamos verificar as Atividades
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "AEC.strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEconomico & " EC, "
    strSQL = strSQL & gstrAtividadeEC & " AEC, "
    strSQL = strSQL & gstrAtividadeDaEmpresa & " AE "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "EC.Pkid = AE.Inteconomico AND "
    strSQL = strSQL & "AEC.Pkid = AE.Intatividade AND "
    strSQL = strSQL & "EC.Strinscricaoimobiliaria = '" & String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text) & "'"
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & "AE.Blnprincipal Desc, "
    strSQL = strSQL & "AEC.strDescricao "
        
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount >= 1 Then
            Do While Not adoResultado.EOF
                   strAtividade = strAtividade & gstrENulo(adoResultado!strDescricao) & ", "
                   adoResultado.MoveNext
            Loop
            strAtividade = Mid(strAtividade, 1, Len(Trim(strAtividade)) - 1)
        End If
    End If
    
    'Vamos pegar o número da guia
    strSQL = ""
    strSQL = strSQL & "select (IntNumeroGuiaPositiva + 1) as Numero From " & gstrEmpresa
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If Val(gstrENulo(adoResultado!Numero)) > "0" Then
                strNumero = gstrENulo(adoResultado!Numero)
            Else
                strNumero = "1"
            End If
            gobjBanco.Execute ("Update " & gstrEmpresa & " Set IntNumeroGuiaPositiva = " & strNumero)
        Else
            ExibeMensagem "Não foi possível conseguir o número da guia."
            Exit Sub
        End If
    Else
        ExibeMensagem "Não foi possível conseguir o número da guia."
        Exit Sub
    End If
    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrGuiaCertidaoPositiva
    
    OpenWordDocumentCertidaoPositiva String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text), strAtividade, strNumeroProcesso, strNumero, strLogradouro, txtstrContribuinte, gstrDataDoSistema, XArrayTabela, XArrayAlinhaColunas, dblTotal, strInscricaoAuxiliar, intUtilizacao
    
    Set gobjBanco = Nothing
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub ImprimiPositivaNegativo()
Dim strSQL              As String
Dim adoResultado        As ADODB.Recordset
Dim strNumero           As String
Dim lngComposicao       As Long

    Screen.MousePointer = vbHourglass
    
    'Vamos pegar o número da guia
    strSQL = ""
    strSQL = strSQL & "select (IntNumeroGuiaPositiva + 1) as Numero From " & gstrEmpresa
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If Val(gstrENulo(adoResultado!Numero)) > "0" Then
                strNumero = gstrENulo(adoResultado!Numero)
            Else
                strNumero = "1"
            End If
            gobjBanco.Execute ("Update " & gstrEmpresa & " Set IntNumeroGuiaPositiva = " & strNumero)
        Else
            ExibeMensagem "Não foi possível conseguir o número da guia."
            Exit Sub
        End If
    Else
        ExibeMensagem "Não foi possível conseguir o número da guia."
        Exit Sub
    End If
    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrGuiaCertidaoPositivaEfeitoNegativa
    OpenWordDocumentCertidaoPositivaNegativo String(gintLenInscricao - Len(Trim(txtstrInscricao.Text)), "0") & UCase(txtstrInscricao.Text), strNumeroProcesso, strNumero, strLogradouro, STRBAIRRO, STRMUNICIPIO, STRUF, txtstrContribuinte, gstrDataDoSistema, strInscricaoAuxiliar, intUtilizacao
    
    Set gobjBanco = Nothing
    Screen.MousePointer = vbDefault
    
End Sub

Private Function blnEmDividaAtiva() As Boolean
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
    
    blnEmDividaAtiva = False
    
    strSQL = strSQL & "Select * "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrDativa & " Da, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "LA.Pkid = Da.Intlancamentoalfa AND "
    strSQL = strSQL & "LA.Pkid = " & lngDividaAtiva
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                blnEmDividaAtiva = True
            End If
        End With
    End If
    
End Function

Private Function blnMontaArrayDividaAtiva() As Boolean
Dim intFor          As Integer
Dim intFor2         As Integer
Dim intPosition     As Integer
Dim blnCheck        As Boolean
Dim adoResultado    As ADODB.Recordset
Dim strSQL          As String
Dim varAux          As Variant
    
    blnMontaArrayDividaAtiva = False
    blnCheck = False
    
    Set xadbDividaAtiva = New XArrayDB
    xadbDividaAtiva.ReDim 0, 0, 0, 36
    xadbDividaAtiva.Clear

    
    For intFor = 0 To xadbLista.UpperBound(1)
        If Trim(xadbLista(intFor, 1)) = -1 Then
            blnCheck = True
            Exit For
        End If
    Next
    
    If blnCheck = False Then
        ExibeMensagem "É preciso estar selecionado algum registro no grid."
        Exit Function
    End If
    
    Set gobjBanco = New clsBanco
    blnCheck = False
    For intFor = 0 To xadbLista.UpperBound(1)
        If Trim(xadbLista(intFor, 1)) = -1 Then
            If Not gobjBanco.CriaADO(strDividaAtiva(CLng(xadbLista(intFor, 0))), 5, adoResultado) Then
                ExibeMensagem "Não foi possivel emitir a certidão de Dívida Ativa."
                Exit Function
            End If
            For intFor2 = 0 To xadbParcelas.UpperBound(1)
                If xadbLista(intFor, 0) = xadbParcelas(intFor2, 2) And (xadbParcelas(intFor2, 1) = "True" Or xadbParcelas(intFor2, 1) = "-1") Then
                  With adoResultado
                  
                    blnCheck = True
                    xadbDividaAtiva.ReDim 0, intPosition, 0, 36
                    
                    varAux = Space$(0) & !Pkid
                    xadbDividaAtiva(intPosition, 0) = varAux
                    
                    varAux = Space$(0) & !intLivro
                    xadbDividaAtiva(intPosition, 1) = varAux
                    
                    varAux = Space$(0) & !intFolha
                    xadbDividaAtiva(intPosition, 2) = varAux
                    
                    varAux = Space$(0) & !dtmdtinscricao
                    xadbDividaAtiva(intPosition, 3) = varAux
                    
                    varAux = Space$(0) & !strComposicaoDaReceita
                    xadbDividaAtiva(intPosition, 4) = varAux
                    
                    varAux = Space$(0) & !intExercicio
                    xadbDividaAtiva(intPosition, 5) = varAux
                    
                    varAux = Space$(0) & !strLogradouro
                    xadbDividaAtiva(intPosition, 6) = varAux
                    
                    varAux = Space$(0) & !strNumero
                    xadbDividaAtiva(intPosition, 7) = varAux
                    
                    varAux = Space$(0) & !STRCOMPLEMENTO
                    xadbDividaAtiva(intPosition, 8) = varAux
                    
                    varAux = Space$(0) & !STRBAIRRO
                    xadbDividaAtiva(intPosition, 9) = varAux
                    
                    varAux = Space$(0) & !STRMUNICIPIO
                    xadbDividaAtiva(intPosition, 10) = varAux
                    
                    varAux = Space$(0) & !STRUF
                    xadbDividaAtiva(intPosition, 11) = varAux
                    
                    varAux = Space$(0) & !INTCEP
                    xadbDividaAtiva(intPosition, 12) = varAux
                    
                    varAux = Space$(0) & gstrFormataInscricao(!strInscricao, !intUtilizacao)
                    xadbDividaAtiva(intPosition, 13) = varAux
                    
                    varAux = Space$(0) & !strnomeproprietario
                    xadbDividaAtiva(intPosition, 14) = varAux
                    
                    varAux = Space$(0) & !STRCNPJCPF
                    xadbDividaAtiva(intPosition, 15) = varAux
                    
                    varAux = Space$(0) & !STRIDENTIDADE
                    xadbDividaAtiva(intPosition, 16) = varAux
                    
                    varAux = Space$(0) & !strLogradouroC
                    xadbDividaAtiva(intPosition, 17) = varAux
                    
                    varAux = Space$(0) & !strNumeroC
                    xadbDividaAtiva(intPosition, 18) = varAux
                    
                    varAux = Space$(0) & !strComplementoC
                    xadbDividaAtiva(intPosition, 19) = varAux
                    
                    varAux = Space$(0) & !strBairroC
                    xadbDividaAtiva(intPosition, 20) = varAux
                    
                    varAux = Space$(0) & !strMunicipioC
                    xadbDividaAtiva(intPosition, 21) = varAux
                    
                    varAux = Space$(0) & !strUFC
                    xadbDividaAtiva(intPosition, 22) = varAux
                    
                    varAux = Space$(0) & !INTCEPC
                    xadbDividaAtiva(intPosition, 23) = varAux
                    
                    varAux = Space$(0) & !Strindexador
                    xadbDividaAtiva(intPosition, 24) = varAux
                    
                    varAux = Space$(0) & !dblvlIndexador
                    xadbDividaAtiva(intPosition, 25) = varAux
                    
                    varAux = Space$(0) & !intCertidao
                    xadbDividaAtiva(intPosition, 26) = varAux
                    
                    varAux = Space$(0) & xadbParcelas(intFor2, 3) 'Número da Parcela
                    xadbDividaAtiva(intPosition, 27) = varAux
                    
                    varAux = Space$(0) & xadbParcelas(intFor2, 12) 'Data de vencimento
                    xadbDividaAtiva(intPosition, 28) = varAux
                    
                    varAux = Space$(0) & xadbParcelas(intFor2, 5) 'Valor Principal
                    xadbDividaAtiva(intPosition, 29) = varAux
                    
                    varAux = Space$(0) & xadbParcelas(intFor2, 6) 'Valor Multa
                    xadbDividaAtiva(intPosition, 30) = varAux
                    
                    varAux = Space$(0) & xadbParcelas(intFor2, 7) 'Valor Juros
                    xadbDividaAtiva(intPosition, 31) = varAux
                    
                    varAux = Space$(0) & xadbParcelas(intFor2, 8) 'Valor Correção
                    xadbDividaAtiva(intPosition, 32) = varAux
                    
                    varAux = Space$(0) & xadbParcelas(intFor2, 8) + xadbParcelas(intFor2, 5) 'Valor Corrigido
                    xadbDividaAtiva(intPosition, 33) = varAux
                    
                    varAux = Space$(0) & xadbParcelas(intFor2, 9) 'Valor Total
                    xadbDividaAtiva(intPosition, 34) = varAux
                    
                    varAux = Space$(0) & !intUtilizacao           'Utilizacao da Composicao
                    xadbDividaAtiva(intPosition, 35) = varAux
                    
                    Select Case gstrENulo(!intUtilizacao)         'Descrição Utilizacao da Composicao
                        Case 1
                            varAux = "Imobiliário"
                            xadbDividaAtiva(intPosition, 36) = varAux
                        Case 2
                            varAux = "Econômico"
                            xadbDividaAtiva(intPosition, 36) = varAux
                        Case 3
                            varAux = "Dívida Ativa"
                            xadbDividaAtiva(intPosition, 36) = varAux
                        Case 4
                            varAux = "Acordo"
                            xadbDividaAtiva(intPosition, 36) = varAux
                        Case 5
                            varAux = "Preco Público"
                            xadbDividaAtiva(intPosition, 36) = varAux
                        Case Else
                            varAux = ""
                            xadbDividaAtiva(intPosition, 36) = varAux
                    End Select

                  End With
                  intPosition = intPosition + 1
                End If
            Next
        End If
    Next
    
    If blnCheck = False Then
        ExibeMensagem "É preciso estar selecionado alguma parcela no grid."
        Exit Function
    End If
    
    blnMontaArrayDividaAtiva = True
End Function

Private Function strDividaAtiva(lngPkid As Long) As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "LA.Pkid, "
    strSQL = strSQL & "CR.intUtilizacao, "
    strSQL = strSQL & "A.intlivro, "
    strSQL = strSQL & "A.intFolha, "
    strSQL = strSQL & "A.dtmdtinscricao, "
    strSQL = strSQL & "LA.Strcomposicaodareceita, "
    strSQL = strSQL & "LA.intExercicio, "
    strSQL = strSQL & "A.strlogradouro, "
    strSQL = strSQL & "A.strnumero, "
    strSQL = strSQL & "A.strcomplemento, "
    strSQL = strSQL & "A.strbairro, "
    strSQL = strSQL & "A.strmunicipio, "
    strSQL = strSQL & "A.struf, "
    strSQL = strSQL & "A.intcep, "
    strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
    strSQL = strSQL & "A.strnomeproprietario, "
    strSQL = strSQL & "A.strcnpjcpf, "
    strSQL = strSQL & "A.stridentidade, "
    strSQL = strSQL & "A.strlogradouroc, "
    strSQL = strSQL & "A.strnumeroc, "
    strSQL = strSQL & "A.strcomplementoc, "
    strSQL = strSQL & "A.strbairroc, "
    strSQL = strSQL & "A.strmunicipioc, "
    strSQL = strSQL & "A.strufc, "
    strSQL = strSQL & "A.intcepc, "
    strSQL = strSQL & "A.strindexador, "
    strSQL = strSQL & "A.dblvlindexador, "
    strSQL = strSQL & "A.intcertidao "
    strSQL = strSQL & "From "
    strSQL = strSQL & "tblLancamentoAlfa LA, "
    strSQL = strSQL & "tblComposicaoDaReceita CR, "
    strSQL = strSQL & "tblDativa A "
    strSQL = strSQL & " Where "
    strSQL = strSQL & "LA.Pkid = A.Intlancamentoalfa AND "
    strSQL = strSQL & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strSQL = strSQL & "LA.Pkid = " & lngPkid
        
    strDividaAtiva = strSQL
    
End Function

Private Sub VerificaAcordoEmDividaAtiva()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
    
    blnAcordoEmDividaAtiva = False
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "PT.bitAcordoDividaAtiva, PT.bitCreditaParcela "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " tblParametrosTributario PT "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            blnAcordoEmDividaAtiva = Val(gstrENulo(adoResultado("bitAcordoDividaAtiva").Value)) = 1
            blnAcordoCreditaParcela = Val(gstrENulo(adoResultado("bitCreditaParcela").Value)) = 1
        Else
            blnAcordoEmDividaAtiva = 0
            blnAcordoCreditaParcela = 0
        End If
    End If

End Sub

Private Function strQueryAnistia() As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select * from " & gstrDescontosProvisorios & " Where "
    strSQL = strSQL & "intParcela = 1 and "
    'If bytDBType = SQLServer Then
    '    strSql = strSql & gstrCONVERT(CDT_DATETIME, "LA.intExercicio") & " Between dtmdtinicial and dtmdtfinal"
    'Else
        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & " Between dtmdtinicial and dtmdtfinal"
    'End If
    
    strQueryAnistia = strSQL
    
End Function

Private Function RetornaExecutivo(intLancamentoAlfaDativa) As String
Dim strSQL As String
Dim adoResultado As ADODB.Recordset

    If IsNull(intLancamentoAlfaDativa) Then
        RetornaExecutivo = ""
        Exit Function
    End If

    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " INTEXECUTIVO "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrDativa
    strSQL = strSQL & " WHERE intLancamentoAlfa = " & intLancamentoAlfaDativa
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
    
        If Not adoResultado.EOF Then
            If Not IsNull(adoResultado!intExecutivo) Then
                strSQL = ""
                strSQL = strSQL & "SELECT "
                strSQL = strSQL & "intNumeroCartorioDistribuidor intNumDis, "
                strSQL = strSQL & "intSerieCartorioDistribuidor intSerDis "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & gstrExecutivo
                strSQL = strSQL & " WHERE PKID = " & adoResultado!intExecutivo
                
                Set gobjBanco = New clsBanco
    
                If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                    RetornaExecutivo = gstrENulo(adoResultado!intNumDis) & "/" & gstrENulo(adoResultado!intSerDis)
                End If
                
            End If
        End If
    
    End If
End Function

Private Function strQueryUtilizacao() As String
Dim strSQL As String

    strSQL = "SELECT * FROM ("
    strSQL = strSQL & "SELECT " & TYP_IMOBILIARIA & " AS PKID, 'Imobiliária' AS strDescricao "
    strSQL = strSQL & "UNION ALL "
    strSQL = strSQL & "SELECT " & TYP_ECONOMICA & " AS PKID, 'Econômica' AS strDescricao "
    strSQL = strSQL & "UNION ALL "
    strSQL = strSQL & "SELECT " & TYP_DIVIDA_ATIVA & " AS PKID, 'Dívida Ativa' AS strDescricao "
    strSQL = strSQL & "UNION ALL "
    strSQL = strSQL & "SELECT " & TYP_ACORDO & " AS PKID, 'Acordo' AS strDescricao "
    strSQL = strSQL & "UNION ALL "
    strSQL = strSQL & "SELECT " & TYP_PRECO_PUBLICO & " AS PKID, 'Preço Público' AS strDescricao "
    strSQL = strSQL & "UNION ALL "
    strSQL = strSQL & "SELECT " & TYP_ISS_CONSTRUCAO & " AS PKID, 'ISS Construção' AS strDescricao "
    strSQL = strSQL & "UNION ALL "
    strSQL = strSQL & "SELECT " & TYP_OUTROS & " AS PKID, 'Outros' AS strDescricao "
    strSQL = strSQL & "UNION ALL "
    strSQL = strSQL & "SELECT " & TYP_IMOBILIARIO_TAXAS & " AS PKID, 'Imobiliário Taxas' AS strDescricao "
    strSQL = strSQL & ") X "
    
    strQueryUtilizacao = strSQL

End Function

Private Function VerificaDiasNaoUteis(ByVal DTMDATA As Date) As Date
Dim strSQL          As String
Dim dtmdtData       As Date
Dim adoResultado    As Recordset

    Select Case Weekday(DTMDATA)
        Case 1 'domingo
            dtmdtData = DateAdd("d", 1, DTMDATA)
            dtmdtData = VerificaDiasNaoUteis(dtmdtData)
        Case 7 'sabado
            dtmdtData = DateAdd("d", 2, DTMDATA)
            dtmdtData = VerificaDiasNaoUteis(dtmdtData)
        Case Else
        
            strSQL = "select * from " & gstrDiasNaoUteis & " where dtmdata = " & gstrConvDtParaSql(DTMDATA)
        
            'Set gobjBanco = New clsBanco
        
            If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                If Not adoResultado.EOF Then
        
                    Select Case adoResultado("byttipo")
                        Case 0 'feriado
                            dtmdtData = DateAdd("d", 1, DTMDATA)
                            dtmdtData = VerificaDiasNaoUteis(dtmdtData)
                        Case 1 'sabado
                            dtmdtData = DateAdd("d", 2, DTMDATA)
                            dtmdtData = VerificaDiasNaoUteis(dtmdtData)
                        Case 2 'domingo
                            dtmdtData = DateAdd("d", 1, DTMDATA)
                            dtmdtData = VerificaDiasNaoUteis(dtmdtData)
                        Case Else
                            dtmdtData = DTMDATA
                    End Select
        
                Else
                    dtmdtData = DTMDATA
                End If
            End If
        
    End Select

    VerificaDiasNaoUteis = dtmdtData
    
End Function
