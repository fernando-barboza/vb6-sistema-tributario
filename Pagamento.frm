VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmPagamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagamentos"
   ClientHeight    =   7080
   ClientLeft      =   2145
   ClientTop       =   2460
   ClientWidth     =   9480
   HelpContextID   =   5
   Icon            =   "Pagamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9480
   Begin VB.ComboBox cbo_temp 
      Height          =   315
      Left            =   7800
      TabIndex        =   72
      Text            =   "Combo1"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtdtmDataAnulacao 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4860
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   67
      Top             =   -120
      Visible         =   0   'False
      Width           =   1245
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6975
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Pagamentos"
      TabPicture(0)   =   "Pagamento.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrProcesso"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_DataEmpenho"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_Saldo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_Total"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblTotalAPagar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_TotalResto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_TotalEmpenho"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblCodEventoContabil"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_TotalReceitaExtra"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDataAnulacao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblRecOrdem"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPKId"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtintProcesso"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtData"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cbo_HistoricoLiquidacao"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmd_HistoricoLiquidacao"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "fra_HistoricoLiquidacao"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkBordero"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "tab_3DPastaEmpenho"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "tdb_Lista"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtSaldo"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtTotal"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtTotalAPagar"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtTotalDespesa"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtTotalResto"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtTotalEmpenho"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkEstorno"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txt_codEvento"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cbo_intEvento"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmd_Evento"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtTotalReceitaExtra"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtTotalRecOrdem"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txt_strCodigoHistorico"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      Begin VB.TextBox txt_strCodigoHistorico 
         Height          =   320
         Left            =   2010
         TabIndex        =   13
         Top             =   1475
         Width           =   885
      End
      Begin VB.TextBox txtTotalRecOrdem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   3330
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   57
         Top             =   4905
         Width           =   1230
      End
      Begin VB.TextBox txtTotalReceitaExtra 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   56
         Top             =   4905
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Evento 
         Height          =   300
         Left            =   8880
         Picture         =   "Pagamento.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "247"
         ToolTipText     =   "Clique para cadastar convênio"
         Top             =   420
         Width           =   330
      End
      Begin VB.ComboBox cbo_intEvento 
         Height          =   315
         Left            =   4905
         TabIndex        =   7
         Top             =   420
         Width           =   4005
      End
      Begin VB.TextBox txt_codEvento 
         Height          =   315
         Left            =   4125
         MaxLength       =   15
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   420
         Width           =   765
      End
      Begin VB.CheckBox chkEstorno 
         Caption         =   "Estorno"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   945
         Width           =   885
      End
      Begin VB.TextBox txtTotalEmpenho 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   52
         Top             =   4560
         Width           =   1140
      End
      Begin VB.TextBox txtTotalResto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   3330
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   53
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtTotalDespesa 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   5565
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   54
         Top             =   4560
         Width           =   1230
      End
      Begin VB.TextBox txtTotalAPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   5565
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   58
         Top             =   4890
         Width           =   1230
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   7905
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   55
         Top             =   4560
         Width           =   1230
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   7905
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   59
         Top             =   4890
         Width           =   1230
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   1695
         Left            =   105
         TabIndex        =   43
         Top             =   5235
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   2990
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
         Columns(1).Caption=   "Artigo Caixa"
         Columns(1).DataField=   "intProcesso"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data"
         Columns(2).DataField=   "dtmData"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Empenho"
         Columns(3).DataField=   "dblTotalEmpenho"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Resto a Pagar"
         Columns(4).DataField=   "dblTotalResto"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Despesa"
         Columns(5).DataField=   "dblTotalDespesaExtra"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Anul. Receita"
         Columns(6).DataField=   "dblTotalAnulacaoReceita"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "bytBordero"
         Columns(7).DataField=   "bytBordero"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Borderô"
         Columns(8).DataField=   "strBordero"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Histórico"
         Columns(9).DataField=   "strHistorico"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Estorno"
         Columns(10).DataField=   "strEstorno"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "bytEstorno"
         Columns(11).DataField=   "bytEstorno"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "intEventoContabil"
         Columns(12).DataField=   "intEvento"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   13
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160664
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=13"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1693"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1693"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2037"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1958"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2037"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1958"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=2037"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=1958"
         Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=2037"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=1958"
         Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(41)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(45)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(47)=   "Column(8).Width=2064"
         Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=1984"
         Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(52)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(55)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(56)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(58)=   "Column(10).Width=1879"
         Splits(0)._ColumnProps(59)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(60)=   "Column(10)._WidthInPix=1799"
         Splits(0)._ColumnProps(61)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(62)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(63)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(64)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(65)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(66)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(67)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(68)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(69)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(70)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(72)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(73)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(74)=   "Column(12).Order=13"
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
         DeadAreaBackColor=   13160664
         RowDividerColor =   13160664
         RowSubDividerColor=   13160664
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=64,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
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
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=86,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=83,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=84,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=85,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
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
         _StyleDefs(89)  =   "Named:id=33:Normal"
         _StyleDefs(90)  =   ":id=33,.parent=0"
         _StyleDefs(91)  =   "Named:id=34:Heading"
         _StyleDefs(92)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(93)  =   ":id=34,.wraptext=-1"
         _StyleDefs(94)  =   "Named:id=35:Footing"
         _StyleDefs(95)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(96)  =   "Named:id=36:Selected"
         _StyleDefs(97)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(98)  =   "Named:id=37:Caption"
         _StyleDefs(99)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(100) =   "Named:id=38:HighlightRow"
         _StyleDefs(101) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(102) =   "Named:id=39:EvenRow"
         _StyleDefs(103) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(104) =   "Named:id=40:OddRow"
         _StyleDefs(105) =   ":id=40,.parent=33"
         _StyleDefs(106) =   "Named:id=41:RecordSelector"
         _StyleDefs(107) =   ":id=41,.parent=34"
         _StyleDefs(108) =   "Named:id=42:FilterBar"
         _StyleDefs(109) =   ":id=42,.parent=33"
      End
      Begin TabDlg.SSTab tab_3DPastaEmpenho 
         Height          =   2655
         Left            =   90
         TabIndex        =   16
         Top             =   1800
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   4683
         _Version        =   393216
         Style           =   1
         Tabs            =   6
         Tab             =   2
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "Empenho"
         TabPicture(0)   =   "Pagamento.frx":13E8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lbl_Empenho"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl_Parcela"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl_OrdemPagamentoEmpenho"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbl_credor"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "dcbOrdemPagamentoEmpenho"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "dcbParcela"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lvw_Empenho"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "dcbEmpenho"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmd_Empenho"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmd_OrdemPagamentoEmpenho"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txt_strnome"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txt_CDC"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "Resto a pagar"
         TabPicture(1)   =   "Pagamento.frx":1404
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label4"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lbl_OrdemPagamentoResto"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "dcbOrdemPagamentoResto"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "dcbParcelaResto"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "dcbResto"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lvw_Resto"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "cmd_Resto"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "cmd_OrdemPagamentoResto"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txt_intExercicioRP"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "Despesa extra-orçamentária"
         TabPicture(2)   =   "Pagamento.frx":1420
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label6"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "lbl_OrdemPagamentoDespesa"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "dcbOrdemPagamentoDespesa"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "lvw_Despesa"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "dcbDespesa"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "cmd_Despesa"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "cmd_OrdemPagamentoDespesa"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "txt_intExercicioDE"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).ControlCount=   8
         TabCaption(3)   =   "Lançamentos"
         TabPicture(3)   =   "Pagamento.frx":143C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lbl_NumCheque"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "lbl_Valor"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "lbl_Conta"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "lbldblSaldoAtual"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "lvw_Conta"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "cmd_PlanoConta"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "cbostrContaContabil"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "cbointContaContabil"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "txtNumCheque"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "txtValorLancamento"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "txt_saldoAtual"
         Tab(3).Control(10).Enabled=   0   'False
         Tab(3).ControlCount=   11
         TabCaption(4)   =   "Receita Extra"
         TabPicture(4)   =   "Pagamento.frx":1458
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lvw_Extra"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Anulação de Receita"
         TabPicture(5)   =   "Pagamento.frx":1474
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "cmd_OrdemPagamentoAnulacao"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "lvw_AnulacaoReceita"
         Tab(5).Control(2)=   "dcbOrdemPagamentoAnulacao"
         Tab(5).Control(3)=   "lbl_OrdemPagamentoAnulacao"
         Tab(5).ControlCount=   4
         Begin VB.TextBox txt_CDC 
            Height          =   315
            Left            =   -71040
            TabIndex        =   82
            Top             =   390
            Width           =   720
         End
         Begin VB.TextBox txt_strnome 
            Height          =   315
            Left            =   -70320
            TabIndex        =   81
            Top             =   390
            Width           =   4320
         End
         Begin VB.CommandButton cmd_OrdemPagamentoAnulacao 
            Height          =   300
            Left            =   -72240
            Picture         =   "Pagamento.frx":1490
            Style           =   1  'Graphical
            TabIndex        =   77
            TabStop         =   0   'False
            Tag             =   "241"
            ToolTipText     =   "Clique para cadastar empenho"
            Top             =   405
            Width           =   330
         End
         Begin VB.TextBox txt_saldoAtual 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73680
            MaxLength       =   15
            MultiLine       =   -1  'True
            OLEDropMode     =   2  'Automatic
            TabIndex        =   74
            Top             =   780
            Width           =   1725
         End
         Begin VB.TextBox txt_intExercicioDE 
            Height          =   300
            Left            =   1530
            TabIndex        =   30
            Top             =   390
            Width           =   765
         End
         Begin VB.TextBox txt_intExercicioRP 
            Height          =   300
            Left            =   -73470
            TabIndex        =   23
            Top             =   390
            Width           =   765
         End
         Begin VB.CommandButton cmd_OrdemPagamentoDespesa 
            Height          =   300
            Left            =   3480
            Picture         =   "Pagamento.frx":181A
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            Tag             =   "241"
            ToolTipText     =   "Clique para cadastar empenho"
            Top             =   405
            Width           =   330
         End
         Begin VB.CommandButton cmd_OrdemPagamentoResto 
            Height          =   300
            Left            =   -71520
            Picture         =   "Pagamento.frx":1BA4
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "241"
            ToolTipText     =   "Clique para cadastar empenho"
            Top             =   405
            Width           =   330
         End
         Begin VB.CommandButton cmd_OrdemPagamentoEmpenho 
            Height          =   300
            Left            =   -72360
            Picture         =   "Pagamento.frx":1F2E
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Tag             =   "241"
            ToolTipText     =   "Clique para cadastar empenho"
            Top             =   405
            Width           =   330
         End
         Begin VB.TextBox txtValorLancamento 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -70920
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   39
            Top             =   780
            Width           =   1725
         End
         Begin VB.TextBox txtNumCheque 
            Height          =   285
            Left            =   -67800
            MaxLength       =   6
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   40
            Top             =   780
            Width           =   1005
         End
         Begin VB.ComboBox cbointContaContabil 
            Height          =   315
            Left            =   -74400
            OLEDragMode     =   1  'Automatic
            TabIndex        =   36
            ToolTipText     =   "Histórico padrão"
            Top             =   390
            Width           =   1875
         End
         Begin VB.ComboBox cbostrContaContabil 
            Height          =   315
            Left            =   -72570
            OLEDragMode     =   1  'Automatic
            Sorted          =   -1  'True
            TabIndex        =   37
            ToolTipText     =   "Histórico padrão"
            Top             =   390
            Width           =   6315
         End
         Begin VB.CommandButton cmd_PlanoConta 
            Height          =   300
            Left            =   -66270
            Picture         =   "Pagamento.frx":22B8
            Style           =   1  'Graphical
            TabIndex        =   38
            TabStop         =   0   'False
            Tag             =   "322"
            ToolTipText     =   "Clique para cadastar conta"
            Top             =   390
            Width           =   330
         End
         Begin VB.CommandButton cmd_Despesa 
            Height          =   300
            Left            =   7770
            Picture         =   "Pagamento.frx":2642
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Clique para cadastar despesa"
            Top             =   390
            Width           =   330
         End
         Begin VB.CommandButton cmd_Resto 
            Height          =   300
            Left            =   -68190
            Picture         =   "Pagamento.frx":29CC
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Tag             =   "246"
            ToolTipText     =   "Clique para cadastar resto a pagar"
            Top             =   390
            Width           =   330
         End
         Begin VB.CommandButton cmd_Empenho 
            Height          =   300
            Left            =   -72360
            Picture         =   "Pagamento.frx":2D56
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Tag             =   "241"
            ToolTipText     =   "Clique para cadastar empenho"
            Top             =   870
            Width           =   330
         End
         Begin MSDataListLib.DataCombo dcbEmpenho 
            Height          =   315
            Left            =   -73530
            TabIndex        =   19
            Tag             =   "1"
            Top             =   870
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "dcbEmpenho"
         End
         Begin MSComctlLib.ListView lvw_Empenho 
            Height          =   1215
            Left            =   -74910
            TabIndex        =   22
            Top             =   1320
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   2143
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "EmpenhoParcela"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "O. Pagamento"
               Object.Width           =   2116
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Número"
               Object.Width           =   1808
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Parcela"
               Object.Width           =   1279
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Previsão"
               Object.Width           =   2011
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Liquidação"
               Object.Width           =   2011
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Valor"
               Object.Width           =   2171
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Desconto"
               Object.Width           =   2171
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Líquido"
               Object.Width           =   2171
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "bytAdiantamento"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Key             =   "Exercicio"
               Text            =   "Exercício O.P."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "DescontoOrc"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Liquidação"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "pkidOP"
               Object.Width           =   0
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcbParcela 
            Height          =   315
            Left            =   -71040
            TabIndex        =   21
            Top             =   870
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "dcbParcela"
         End
         Begin MSComctlLib.ListView lvw_Resto 
            Height          =   1785
            Left            =   -74910
            TabIndex        =   29
            Top             =   780
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   3149
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "RestoParcela"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "O Pagamento"
               Object.Width           =   2116
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Exercício"
               Object.Width           =   1843
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Numero"
               Object.Width           =   1939
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Parcela"
               Object.Width           =   1279
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Previsão"
               Object.Width           =   2063
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Valor"
               Object.Width           =   2171
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Desconto"
               Object.Width           =   2171
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Líquido"
               Object.Width           =   2171
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Exercício O.P."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Liquidado"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcbResto 
            Height          =   315
            Left            =   -69390
            TabIndex        =   26
            Tag             =   "1"
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcbParcelaResto 
            Height          =   315
            Left            =   -67110
            TabIndex        =   28
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcbDespesa 
            Height          =   315
            Left            =   6570
            TabIndex        =   33
            Tag             =   "1"
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvw_Despesa 
            Height          =   1785
            Left            =   90
            TabIndex        =   35
            Top             =   810
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   3149
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Numero"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "O. Pagamento"
               Object.Width           =   1808
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Número"
               Object.Width           =   1808
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Previsão"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Valor"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Credor"
               Object.Width           =   7690
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Exercício OP"
               Object.Width           =   2019
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Conta 
            Height          =   1455
            Left            =   -74910
            TabIndex        =   41
            Top             =   1110
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Conta"
               Object.Width           =   2911
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descrição"
               Object.Width           =   8944
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Valor"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Cheque"
               Object.Width           =   1940
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Extra 
            Height          =   2180
            Left            =   -74910
            TabIndex        =   66
            Top             =   390
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   3836
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Número"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Parcela"
               Object.Width           =   1279
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Conta"
               Object.Width           =   2911
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Descrição"
               Object.Width           =   7815
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "pkid"
               Object.Width           =   0
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcbOrdemPagamentoEmpenho 
            Height          =   315
            Left            =   -73530
            TabIndex        =   17
            Tag             =   "1"
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "dcbOrdemPagamento"
         End
         Begin MSDataListLib.DataCombo dcbOrdemPagamentoResto 
            Height          =   315
            Left            =   -72690
            TabIndex        =   24
            Tag             =   "1"
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "dcbOrdemPagamento"
         End
         Begin MSDataListLib.DataCombo dcbOrdemPagamentoDespesa 
            Height          =   315
            Left            =   2310
            TabIndex        =   31
            Tag             =   "1"
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "dcbOrdemPagamento"
         End
         Begin MSComctlLib.ListView lvw_AnulacaoReceita 
            Height          =   1785
            Left            =   -74910
            TabIndex        =   78
            Top             =   780
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   3149
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "pkid"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "O. Pagamento"
               Object.Width           =   2116
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Descrição"
               Object.Width           =   7938
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Data"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Valor"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Exercício O.P."
               Object.Width           =   2540
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcbOrdemPagamentoAnulacao 
            Height          =   315
            Left            =   -73410
            TabIndex        =   79
            Tag             =   "1"
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "dcbOrdemPagamento"
         End
         Begin VB.Label lbl_credor 
            AutoSize        =   -1  'True
            Caption         =   "Credor"
            Height          =   195
            Left            =   -71640
            TabIndex        =   83
            Top             =   450
            Width           =   465
         End
         Begin VB.Label lbl_OrdemPagamentoAnulacao 
            AutoSize        =   -1  'True
            Caption         =   "Ordem Pagamento"
            Height          =   195
            Left            =   -74790
            TabIndex        =   80
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label lbldblSaldoAtual 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Atual"
            Height          =   195
            Left            =   -74520
            TabIndex        =   75
            Top             =   870
            Width           =   810
         End
         Begin VB.Label lbl_OrdemPagamentoDespesa 
            AutoSize        =   -1  'True
            Caption         =   "Ordem Pagamento"
            Height          =   195
            Left            =   90
            TabIndex        =   71
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label lbl_OrdemPagamentoResto 
            AutoSize        =   -1  'True
            Caption         =   "Ordem Pagamento"
            Height          =   195
            Left            =   -74910
            TabIndex        =   70
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label lbl_OrdemPagamentoEmpenho 
            AutoSize        =   -1  'True
            Caption         =   "Ordem Pagamento"
            Height          =   195
            Left            =   -74910
            TabIndex        =   68
            Top             =   450
            Width           =   1320
         End
         Begin VB.Label lbl_Conta 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   -74910
            TabIndex        =   51
            Top             =   450
            Width           =   420
         End
         Begin VB.Label lbl_Valor 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   -71370
            TabIndex        =   50
            Top             =   870
            Width           =   360
         End
         Begin VB.Label lbl_NumCheque 
            AutoSize        =   -1  'True
            Caption         =   "Cheque"
            Height          =   195
            Left            =   -68400
            TabIndex        =   49
            Top             =   870
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   5940
            TabIndex        =   48
            Top             =   450
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   -70020
            TabIndex        =   47
            Top             =   450
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Parcela"
            Height          =   195
            Left            =   -67710
            TabIndex        =   46
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl_Parcela 
            AutoSize        =   -1  'True
            Caption         =   "Parcela"
            Height          =   195
            Left            =   -71715
            TabIndex        =   45
            Top             =   930
            Width           =   540
         End
         Begin VB.Label lbl_Empenho 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   -74145
            TabIndex        =   44
            Top             =   930
            Width           =   555
         End
      End
      Begin VB.CheckBox chkBordero 
         Caption         =   "Pag. por borderô"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1395
         Width           =   1485
      End
      Begin VB.Frame fra_HistoricoLiquidacao 
         Caption         =   " Histórico "
         Height          =   615
         Left            =   2010
         TabIndex        =   11
         Top             =   810
         Width           =   7245
         Begin VB.TextBox txtHistoricoLiquidacao 
            Height          =   435
            Left            =   -15
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   165
            Width           =   7245
         End
      End
      Begin VB.CommandButton cmd_HistoricoLiquidacao 
         Height          =   300
         Left            =   8910
         Picture         =   "Pagamento.frx":30E0
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "248"
         ToolTipText     =   "Clique para cadastar histórico"
         Top             =   1470
         Width           =   330
      End
      Begin VB.ComboBox cbo_HistoricoLiquidacao 
         Height          =   315
         Left            =   2940
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Histórico padrão"
         Top             =   1465
         Width           =   5955
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   2475
         OLEDragMode     =   1  'Automatic
         TabIndex        =   4
         Top             =   420
         Width           =   1005
      End
      Begin VB.TextBox txtintProcesso 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1020
         MaxLength       =   10
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         ToolTipText     =   "Artigo de caixa somente para pesquisa"
         Top             =   420
         Width           =   1005
      End
      Begin VB.TextBox txtPKId 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6690
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   42
         Top             =   90
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblRecOrdem 
         AutoSize        =   -1  'True
         Caption         =   "Rec. Orc"
         Height          =   195
         Left            =   2565
         TabIndex        =   76
         Top             =   4965
         Width           =   645
      End
      Begin VB.Label lblDataAnulacao 
         AutoSize        =   -1  'True
         Caption         =   "Data de cancelamento"
         Height          =   195
         Left            =   2940
         TabIndex        =   73
         Top             =   -30
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label lbl_TotalReceitaExtra 
         AutoSize        =   -1  'True
         Caption         =   "Receita Extra"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   4950
         Width           =   960
      End
      Begin VB.Label lblCodEventoContabil 
         AutoSize        =   -1  'True
         Caption         =   "Evento"
         Height          =   195
         Left            =   3570
         TabIndex        =   5
         Top             =   495
         Width           =   510
      End
      Begin VB.Label lbl_TotalEmpenho 
         AutoSize        =   -1  'True
         Caption         =   "Empenho"
         Height          =   195
         Left            =   405
         TabIndex        =   65
         Top             =   4590
         Width           =   675
      End
      Begin VB.Label lbl_TotalResto 
         AutoSize        =   -1  'True
         Caption         =   "Resto"
         Height          =   195
         Left            =   2790
         TabIndex        =   64
         Top             =   4590
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Despesa"
         Height          =   195
         Left            =   4890
         TabIndex        =   63
         Top             =   4590
         Width           =   630
      End
      Begin VB.Label lblTotalAPagar 
         AutoSize        =   -1  'True
         Caption         =   "Total a pagar"
         Height          =   195
         Left            =   4575
         TabIndex        =   62
         Top             =   4950
         Width           =   945
      End
      Begin VB.Label lbl_Total 
         AutoSize        =   -1  'True
         Caption         =   "Lançamento"
         Height          =   195
         Left            =   6960
         TabIndex        =   61
         Top             =   4590
         Width           =   885
      End
      Begin VB.Label lbl_Saldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Left            =   7440
         TabIndex        =   60
         Top             =   4950
         Width           =   405
      End
      Begin VB.Label lbl_DataEmpenho 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   2070
         TabIndex        =   3
         Top             =   480
         Width           =   345
      End
      Begin VB.Label lblstrProcesso 
         AutoSize        =   -1  'True
         Caption         =   "Artigo Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim mblnAlterando               As Boolean
    Dim mblnselecionou              As Boolean
    Dim mblnAlterandoEmpenho        As Boolean
    Dim mblnAlterandoResto          As Boolean
    Dim mblnAlterandoConta          As Boolean
    Dim mblnAlterandoAnulacao       As Boolean
    Dim mblnInclirConta             As Boolean
    Dim mblnAlterandoDespesa        As Boolean
    Dim mobjLista                   As Object
    Dim mobjAux                     As Object
    Dim mblnClickOk                 As Boolean
    Dim mblnCarregaFormConta        As Boolean
    Dim blnClick                    As Boolean
    Dim intTipoEventoSelecionado    As Integer
    Dim mblcodEventoMudou           As Boolean
    Dim recGrid                     As New ADODB.Recordset
    Dim blnPrimeiraVez              As Boolean
    Dim blnOrdem                    As Boolean
    Dim blnOrdenacaoAsc         As Boolean
    Dim bytOrdenacao            As Byte
    Dim intArquivo                                 As Integer    'Numero da Abertura da Porta Paralela
    Dim strRegistroAutent                          As String     'Linha a ser impressa



Private Sub CriaViewBordereaux()

'******************************************************************************************
' Data: 11/06/2003
' Alteração: - Incluída instrução IF fazendo com que, caso o banco de dados corrente seja o
'            Oracle, a função nada faça.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL     As String
    
    If (bytDBType = EDatabases.Oracle) Then Exit Sub
    
    'Removido na Pendencia orc1525.. e a criação das views abaixo
    'foi inseridas por meio de script, ou seja já se encontra na base
    Set gobjBanco = New clsBanco
    strSQL = "CREATE VIEW vw_Contribuinte AS SELECT * FROM " & gstrContribuinte
    gobjBanco.Execute strSQL
    strSQL = "CREATE VIEW vw_Banco AS SELECT * FROM " & gstrBanco
    gobjBanco.Execute strSQL
    strSQL = "CREATE VIEW vw_Agencia AS SELECT * FROM " & gstrAgencia
    gobjBanco.Execute strSQL
    strSQL = "CREATE VIEW vw_ContaBancaria AS SELECT * FROM " & gstrContaBancaria
    gobjBanco.Execute strSQL
End Sub

Private Sub ImprimeBordereaux(lngProcesso As Long)

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
'    CriaViewBordereaux
'    strSql = "sp_BordereauxPagamento " & lngProcesso
    strSQL = gstrStoredProcedure("sp_BordereauxPagamento", CStr(lngProcesso), True)
    ImprimeRelatorio rptBordereauxPagamento, strSQL
End Sub

Private Sub LePagamento()
    With tdb_Lista
        LeEmpenho
        txtintProcesso = .Columns("intProcesso")
        txtData = .Columns("dtmData")
        chkBordero = Val(.Columns("bytBordero"))
        chkEstorno = Val(.Columns("bytEstorno"))
    End With
    'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCancelar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar
    lblDataAnulacao.Enabled = True
    txtdtmDataAnulacao.OLEDropMode = 0
    TrocaCorObjeto txtdtmDataAnulacao, False
End Sub

Private Sub LeEmpenho()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL               As String
    Dim adoResultado         As ADODB.Recordset
    strSQL = ""
'    strSql = strSql & "sp_SubempenhoPago " & Val(txtPKId)
    strSQL = strSQL & gstrStoredProcedure("sp_SubempenhoPago", CStr(Val(txtPKID)), True)
    LeDaTabelaParaObj "", lvw_Empenho, strSQL
    Totaliza lvw_Empenho, txtTotalEmpenho
    
    preencheLiquidacaoExtra lvw_Empenho
End Sub

Private Sub LeResto()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL               As String
    Dim adoResultado         As ADODB.Recordset
    strSQL = ""
'    strSql = strSql & "sp_RestoPago " & Val(txtPKId)
    strSQL = strSQL & gstrStoredProcedure("sp_RestoPago", CStr(Val(txtPKID)), True)
    LeDaTabelaParaObj "", lvw_Resto, strSQL
    Totaliza lvw_Resto, txtTotalResto
    preencheLiquidacaoExtra lvw_Resto
End Sub

Private Sub LeDespesa()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL               As String
    Dim adoResultado         As ADODB.Recordset
    strSQL = ""
'    strSql = strSql & "sp_DespesaPaga " & Val(txtPKId)
    strSQL = strSQL & gstrStoredProcedure("sp_DespesaPaga", CStr(Val(txtPKID)), True)
    LeDaTabelaParaObj "", lvw_Despesa, strSQL
    Totaliza lvw_Despesa, txtTotalDespesa
End Sub


Private Sub LeAnulacaoReceita()

    Dim strSQL               As String
    Dim adoResultado         As ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT OPA.PKID, OPA.PKID, OP.INTNUMERO intOrdem, OPA.Strdescricao, "
    strSQL = strSQL & " OPA.Dtmdtatualizacao,  OPA.Dblvalor, OP.intExercicio intExercicioOP "
    strSQL = strSQL & "FROM " & gstrOrdemPagamento & " OP ,"
    strSQL = strSQL & gstrOrdemPagamentoAnulacaoReceita & " OPA "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " OP.PKID = OPA.intOrdemPagamento "
    strSQL = strSQL & " AND  OPA.PKID IN ( "

    strSQL = strSQL & " SELECT OPA1.PKID FROM  "
    strSQL = strSQL & gstrOrdemPagamentoAnulacaoReceita & " OPA1 "
    strSQL = strSQL & " WHERE OPA1.intProcesso = " & txtPKID.Text
    
    strSQL = strSQL & " UNION ALL "
    strSQL = strSQL & " SELECT APA.intOrdemPagamentoAnulacaoRec FROM  "
    strSQL = strSQL & gstrAnulacaoRecPagtoAnulado & " APA "
    strSQL = strSQL & " WHERE APA.intProcesso = " & txtPKID.Text
    
    strSQL = strSQL & " UNION ALL "
    strSQL = strSQL & " SELECT APA1.intOrdemPagamentoAnulacaoRec FROM  "
    strSQL = strSQL & gstrAnulacaoRecPagtoAnulado & " APA1 "
    strSQL = strSQL & " WHERE APA1.intProcessoOriginal = " & txtPKID.Text
    strSQL = strSQL & " ) "
    
    LeDaTabelaParaObj "", lvw_AnulacaoReceita, strSQL
    SomaTotalRecOrdemAnulacao
End Sub


Private Sub CriaView()
    Dim strSQL     As String
    strSQL = strSQL & "CREATE VIEW vw_Contribuinte AS SELECT PKId, "
    strSQL = strSQL & "strNome FROM " & gstrContribuinte
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSQL
End Sub

Private Sub LimpaTelaPagamento(Optional blnSetaData As Boolean, Optional bytTipoLimpeza As Byte)
    blnPrimeiraVez = False
    txtintProcesso = ""
    chkBordero = 0
    txtHistoricoLiquidacao.Text = ""
    cbo_HistoricoLiquidacao = ""
    
    txtTotal = "0,00"
    txtTotalDespesa = "0,00"
    txtTotalResto = "0,00"
    txtTotalRecOrdem = "0,00"
    txtTotalEmpenho = "0,00"
    txtTotalReceitaExtra = "0,00"
    txtTotalAPagar = "0,00"
    txtSaldo = "0,00"
    
    
    lvw_Conta.ListItems.Clear
    lvw_Empenho.ListItems.Clear
    lvw_Resto.ListItems.Clear
    lvw_Despesa.ListItems.Clear
    lvw_AnulacaoReceita.ListItems.Clear
    chkEstorno.Value = 0
    txt_Cdc.Text = ""
    txt_strNome.Text = ""
    
    TrocaCorObjeto txtData, False
    TrocaCorObjeto cmd_Evento, False
    TrocaCorObjeto cbo_intEvento, False
    TrocaCorObjeto txt_codEvento, False
    TrocaCorObjeto dcbEmpenho, False
    TrocaCorObjeto dcbParcela, False
    TrocaCorObjeto dcbResto, False
    TrocaCorObjeto dcbParcelaResto, False
    TrocaCorObjeto dcbDespesa, False
    txt_codEvento.BackColor = vbWindowBackground
    txt_codEvento.Enabled = True
    chkEstorno.Enabled = True
    chkBordero.Enabled = True
    'M4R
    TrocaCorObjeto cbointContaContabil, False
    TrocaCorObjeto cbostrContaContabil, False
    TrocaCorObjeto cmd_PlanoConta, False
    
    'Pen_773
    txtintProcesso.Enabled = True
    txtintProcesso.BackColor = -2147483643
    
    LimpaDadosEmpenho True
    LimpaDadosResto True
    LimpaDadosDespesa True
    LimpaDadosConta True
    LimpaDadosAnulacao True
    
    
    'AtualizaListas
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    lblDataAnulacao.Enabled = False
    TrocaCorObjeto txtdtmDataAnulacao, True
    If blnSetaData Then
        txtData.SetFocus
    End If
    mblnAlterando = False
    tab_3DPastaEmpenho.Tab = 3
    habilitaGuias 3
        
    If bytTipoLimpeza = 1 Then
       txtData = ""
       cbo_intEvento.Text = ""
       txt_codEvento.Text = ""
    Else
       cbo_intevento_Click
    End If
    
    
End Sub

Private Sub AtualizaListas()
    LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
    LeDaTabelaParaObj "", dcbResto, strQueryResto
    LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
End Sub

Private Function blnDadosPagamentoOk() As Boolean
    Dim intCont             As Integer
    Dim strSQL              As String
    Dim strTabelaVerificar  As String
    Dim strCampoVerificar   As String
    Dim intInd              As Integer
    Dim listaPKID           As String
    Dim adoResultado        As ADODB.Recordset

    If tab_3DPastaEmpenho.TabEnabled(0) = True Then
        If lvw_Empenho.ListItems.Count = 0 Then
            ExibeMensagem "Nenhum lançamento foi informado."
            Exit Function
        End If
    End If
    
    If tab_3DPastaEmpenho.TabEnabled(1) = True Then
        If lvw_Resto.ListItems.Count = 0 Then
            ExibeMensagem "Nenhum lançamento foi informado."
            Exit Function
        End If
    End If
    
    If tab_3DPastaEmpenho.TabEnabled(2) = True Then
        If lvw_Despesa.ListItems.Count = 0 Then
            ExibeMensagem "Nenhum lançamento foi informado."
            Exit Function
        End If
    End If
    
    'M4R
    If tab_3DPastaEmpenho.TabEnabled(1) = True Then
        If lvw_Resto.ListItems.Count > 0 Then
            
            strSQL = ""
            If cbo_intEvento.ListIndex >= 0 Then
                strSQL = "SELECT intExercicio FROM " & gstrEvento & " WHERE Pkid = " & cbo_intEvento.ItemData(cbo_intEvento.ListIndex)
                
                Set gobjBanco = New clsBanco
                
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    If adoResultado.RecordCount > 0 Then
                        For intCont = 1 To lvw_Resto.ListItems.Count
                            If Val(Trim(lvw_Resto.ListItems(intCont).ListSubItems(2).Text)) <> Val(adoResultado!intExercicio) Then
                                ExibeMensagem "Exite(m) parcela(s) que não pertence(m) ao evento contabil selecionado !!!"
                                Exit Function
                            End If
                            
                        Next intCont
                    End If
                End If
            End If
            For intCont = 1 To lvw_Resto.ListItems.Count
                strSQL = ""
                strSQL = strSQL & "Select dtmLiquidacao From " & gstrSubempenho
                strSQL = strSQL & " Where pkId = " & lvw_Resto.ListItems(intCont).Tag
                
                Set gobjBanco = New clsBanco
                
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    If CDate(adoResultado!dtmLiquidacao) > CDate(txtData) Then
                        ExibeMensagem "Data de pagamento inferior à Data de Liquidação da parcela " & lvw_Resto.ListItems(intCont).SubItems(3) & "/" & lvw_Resto.ListItems(intCont).SubItems(4) & ". Data de Liquidação: " & gstrDataFormatada(adoResultado!dtmLiquidacao)
                        
                        Exit Function
                    End If
                End If
            Next
        End If
    End If
    
    If gblnDataValida(txtData) = False Then
        ExibeMensagem "Data informada não é valida"
        txtData.SetFocus
        Exit Function
    ElseIf CDate(txtData) < CDate(strDataEncerramento) Then
        ExibeMensagem "Data informada não pode ser menor que Data do Encerramento."
        txtData.SetFocus
        Exit Function
    ElseIf Year(CDate(txtData)) <> gintExercicio Then
        ExibeMensagem "Data informada deve estar dentro do Exercício corrente."
        txtData.SetFocus
        Exit Function
    ElseIf cbo_intEvento.ListIndex = -1 Then
        ExibeMensagem "O Evento Contabil não foi informado."
        cbo_intEvento.SetFocus
        Exit Function
    ElseIf Year(txtData) <> gintExercicio Then
        ExibeMensagem "A data do pagamento tem que estar dentro do execício corrente."
        txtData.SetFocus
        Exit Function
    ElseIf Val(gstrConvVrParaSql(txtTotal)) <> Val(gstrConvVrParaSql(txtTotalAPagar)) Then
        'If intTipoEventoSelecionado <> 11 Then
            ExibeMensagem "Total de lançamentos está diferente do total a ser pago."
            cbointContaContabil.SetFocus
            Exit Function
        'End If
    ElseIf chkBordero And lvw_Conta.ListItems.Count > 1 Then
        ExibeMensagem "Pagamento através de boderô não ser lançado em mais de uma conta."
        cbointContaContabil.SetFocus
        Exit Function
    ElseIf lvw_Conta.ListItems.Count = 0 Then
        ExibeMensagem "Não há conta informada para lançamento."
        cbointContaContabil.SetFocus
        Exit Function
    ElseIf Not blnDataLiquidacaoOk Then
        
        Exit Function
    Else
        blnDadosPagamentoOk = True
    End If
    
    If chkEstorno.Value = 1 Then Exit Function
    
    
   'Verifica se algum lancamento já esta presente em alguma OP ou Pagamento
   If tab_3DPastaEmpenho.TabEnabled(0) = True Then
      
      For intInd = 1 To lvw_Empenho.ListItems.Count
          If lvw_Empenho.ListItems(intInd).SubItems(1) = "--" Then
            listaPKID = listaPKID & lvw_Empenho.ListItems(intInd).Tag & ","
          End If
      Next
      If Len(listaPKID) > 0 Then listaPKID = Mid(listaPKID, 1, Len(listaPKID) - 1)
   
      strSQL = ""
      strSQL = strSQL & "SELECT OP.intNumero ORDEM , E.intNumero " & strCONCAT & " ' \ ' " & strCONCAT & " SE.intnumero  lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE,"
      strSQL = strSQL & gstrOrdemPagamentoEmpenho & ","
      strSQL = strSQL & gstrEmpenho & " E,"
      strSQL = strSQL & gstrOrdemPagamento & " OP"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " intParcela in(" & listaPKID & ")"
      strSQL = strSQL & " AND OP.PKID = intordempagamento"
      strSQL = strSQL & " AND (OP.Bytcancelado = 0 OR OP.Bytcancelado IS NULL)"
      strSQL = strSQL & " AND SE.PKID = intParcela"
      strSQL = strSQL & " AND E.PKID = SE.intEmpenho"
   

      strSQL = strSQL & " UNION"

      strSQL = strSQL & " SELECT PP.intNumero ORDEM , E.intNumero " & strCONCAT & " ' \P ' " & strCONCAT & " SE.intnumero  lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE,"
      strSQL = strSQL & gstrProcessoPagamento & " PP,"
      strSQL = strSQL & gstrEmpenho & " E"
      
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " SE.PKID in(" & listaPKID & ") AND PP.PKID = SE.intProcesso"
      strSQL = strSQL & " AND SE.intEmpenho = E.PKID"
   
   ElseIf tab_3DPastaEmpenho.TabEnabled(1) = True Then
      
      For intInd = 1 To lvw_Resto.ListItems.Count
          If lvw_Resto.ListItems(intInd).SubItems(1) = "--" Then
            listaPKID = listaPKID & lvw_Resto.ListItems(intInd).Tag & ","
          End If
      Next
      If Len(listaPKID) > 0 Then listaPKID = Mid(listaPKID, 1, Len(listaPKID) - 1)
   
      strSQL = ""
      strSQL = strSQL & "SELECT OP.intNumero ORDEM , E.intNumero " & strCONCAT & "' \ '" & strCONCAT & " SE.intnumero  lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE,"
      strSQL = strSQL & gstrOrdemPagamentoResto & ","
      strSQL = strSQL & gstrEmpenho & " E,"
      strSQL = strSQL & gstrOrdemPagamento & " OP"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " intParcela in(" & listaPKID & ")"
      strSQL = strSQL & " AND OP.PKID = intordempagamento"
      strSQL = strSQL & " AND (OP.Bytcancelado = 0 OR OP.Bytcancelado IS NULL)"
      strSQL = strSQL & " AND SE.PKID = intParcela"
      strSQL = strSQL & " AND E.PKID = SE.intEmpenho"
   
      strSQL = strSQL & " UNION"

      strSQL = strSQL & " SELECT PP.intNumero ORDEM , E.intNumero " & strCONCAT & "' \P ' " & strCONCAT & " SE.intnumero  lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE,"
      strSQL = strSQL & gstrProcessoPagamento & " PP,"
      strSQL = strSQL & gstrEmpenho & " E"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " SE.PKID in(" & listaPKID & ") AND PP.PKID = SE.intProcesso"
      strSQL = strSQL & " AND SE.intEmpenho = E.PKID"
   
   ElseIf tab_3DPastaEmpenho.TabEnabled(2) = True Then
      
      For intInd = 1 To lvw_Despesa.ListItems.Count
        If lvw_Despesa.ListItems(intInd).SubItems(1) = "--" Then
          listaPKID = listaPKID & lvw_Despesa.ListItems(intInd).Tag & ","
        End If
      Next
      If Len(listaPKID) > 0 Then listaPKID = Mid(listaPKID, 1, Len(listaPKID) - 1)
      
      strSQL = ""
      strSQL = strSQL & "SELECT OP.intNumero ORDEM , '' " & strCONCAT & " DE.intnumero Lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DE,"
      strSQL = strSQL & gstrOrdemPagamentoDespesaExtra & ","
      strSQL = strSQL & gstrOrdemPagamento & " OP"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " intdespesaextraorcamentaria in(" & listaPKID & ")"
      strSQL = strSQL & " AND (OP.Bytcancelado = 0 OR OP.Bytcancelado IS NULL)"
      strSQL = strSQL & " AND OP.PKID = intordempagamento"
      strSQL = strSQL & " AND DE.PKID = intdespesaextraorcamentaria"
      
      strSQL = strSQL & " UNION"

      strSQL = strSQL & " SELECT PP.intNumero ORDEM , 'P' " & strCONCAT & " DE.intnumero  lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DE,"
      strSQL = strSQL & gstrProcessoPagamento & " PP"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " DE.PKID in(" & listaPKID & ") AND PP.PKID = DE.intProcesso"
   End If
       
   If Trim(listaPKID) = "" Then
        blnDadosPagamentoOk = True
        Exit Function
   End If

   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
      With adoResultado
          If .EOF = False Then
              If InStr(!Lancamento, "\P") <> 0 Then
                ExibeMensagem "O lançamento N°" & Replace(!Lancamento, "\P", "\") & " já está presente no Pagamento N°" & !Ordem
              ElseIf InStr(!Lancamento, "P") <> 0 Then
                ExibeMensagem "O lançamento N°" & Replace(!Lancamento, "P", "") & " já está presente no Pagamento N°" & !Ordem
              Else
                ExibeMensagem "O lançamento N°" & !Lancamento & " já está presente na Ordem de Pagamento N°" & !Ordem
              End If
              blnDadosPagamentoOk = False
              Exit Function
          End If
      End With
   End If
    
End Function

Private Function blnDataLiquidacaoOk() As Boolean
    Dim bytInd      As Byte
    Dim strDataAux  As String
    For bytInd = 1 To lvw_Empenho.ListItems.Count
        strDataAux = lvw_Empenho.ListItems(bytInd).SubItems(5)
        If CVDate(strDataAux) > CVDate(txtData) Then
            ExibeMensagem "Parcela de empenho com data de liquidação superior à data do pagamento."
            If gblnEncontroItemNoListView(lvw_Empenho, Trim$(lvw_Empenho.ListItems(bytInd).Text), lvwText) Then
                lvw_Empenho.SetFocus
                SendKeys " "
                mblnAlterandoEmpenho = True
            Else
                mblnAlterandoEmpenho = False
            End If
            Exit Function
        End If
    Next
    For bytInd = 1 To lvw_Despesa.ListItems.Count
        strDataAux = lvw_Despesa.ListItems(bytInd).SubItems(3)
        If CVDate(strDataAux) > CVDate(txtData) Then
            ExibeMensagem "Despesa extra-orçamentária com data prevista superior à data do pagamento."
            If gblnEncontroItemNoListView(lvw_Despesa, Trim$(lvw_Despesa.ListItems(bytInd).Text), lvwText) Then
                lvw_Despesa.SetFocus
                SendKeys " "
                mblnAlterandoDespesa = True
            Else
                mblnAlterandoDespesa = False
            End If
            Exit Function
        End If
    Next
    blnDataLiquidacaoOk = True
End Function

Private Sub GravaPagamento()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 11/06/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL
'            permitindo, assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL                  As String
    Dim strOrigemMovimento      As String
    Dim strTotal                As String
    Dim intInd                  As Integer
    Dim adoResultado            As ADODB.Recordset
    Dim mblnEstorno             As Boolean
    Dim lngProcesso             As Long
    Dim intIdxConta             As Integer
    Dim intIdxExtra             As Integer
    Dim IDOrdemPagamento        As Long
    
    
    If Not mblnAlterando Then
        If blnDadosPagamentoOk Then
            If gblnExclusaoGravacaoOk("I", "Confirma pagamento?", True) Then
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaBeginTrans
                
                mblnEstorno = IIf(chkEstorno.Value = 0, False, True)
                
                If tab_3DPastaEmpenho.TabEnabled(0) = False Then
                    lvw_Empenho.ListItems.Clear
                End If
                
                If tab_3DPastaEmpenho.TabEnabled(1) = False Then
                    lvw_Resto.ListItems.Clear
                End If
                
                If tab_3DPastaEmpenho.TabEnabled(2) = False Then
                    lvw_Despesa.ListItems.Clear
                End If
                
                strSQL = ""
                strSQL = strSQL & gstrStoredProcedure("sp_GravaProcessoPagamento", _
                    gstrConvDtParaSql(txtData) & ", " & _
                    "'" & Trim(txtHistoricoLiquidacao) & "', " & _
                    glngCodUsr & "," & CStr(chkBordero.Value) & "," & CStr(chkEstorno.Value) & _
                    "," & gstrItemData(cbo_intEvento) & ", " & gintExercicio, True)
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    With adoResultado
                        If .EOF = False Then
                            lngProcesso = !intProcesso
                        End If
                    End With
                End If
                If lngProcesso = 0 Then
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaRollbackTrans
                    Exit Sub
                End If
                strSQL = ""
                strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
                With lvw_Conta
                    For intInd = 1 To .ListItems.Count
                        If chkEstorno.Value = 0 Then
                            If SaldoContaContabilAtual(.ListItems(intInd).Tag, Month(CDate(txtData)), gintExercicio, CDbl(.ListItems(intInd).SubItems(2)), txtData) = Empty Then
                                Set gobjBanco = New clsBanco
                                gobjBanco.ExecutaRollbackTrans
                                Exit Sub
                            End If
                        End If
                        strSQL = strSQL & "INSERT INTO " & gstrLancamentoContabil & " ("
                        strSQL = strSQL & "intProcesso, intConta, dblValor, bytNatureza, "
                        strSQL = strSQL & "bytTipo, strDocumento, dtmDtAtualizacao, lngCodUsr)"
                        strSQL = strSQL & " VALUES "
                        strSQL = strSQL & "(" & lngProcesso & ", "
                        strSQL = strSQL & .ListItems(intInd).Tag & ", "
                        strSQL = strSQL & IIf(mblnEstorno, gstrConvVrParaSql(CDbl(lvw_Conta.ListItems(intInd).SubItems(2)) * -1), gstrConvVrParaSql(.ListItems(intInd).SubItems(2))) & ", "
                        strSQL = strSQL & "0, 0, '" & Trim(.ListItems(intInd).SubItems(3)) & "', "
                        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
                        strSQL = strSQL & glngCodUsr & ");"
                    
                    Next
                End With
    
                For intInd = 1 To lvw_Empenho.ListItems.Count
                    If mblnEstorno Then
                        'Registra o estorno do pagamento das parcela do empenho
                        strSQL = strSQL & "INSERT INTO " & gstrSubempenhoPagtoAnulado & " ("
                        strSQL = strSQL & "intSubEmpenho,intProcesso,intProcessoOriginal,"
                        strSQL = strSQL & "dtmDataEstorno,dtmDtAtualizacao,lngCodUsr)"
                        strSQL = strSQL & "SELECT PKId, "
                        strSQL = strSQL & CStr(lngProcesso) & " ,intProcesso,"
                        strSQL = strSQL & gstrConvDtParaSql(txtData) & ", "
                        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                        strSQL = strSQL & glngCodUsr & " FROM " & gstrSubempenho & " "
                        strSQL = strSQL & "WHERE PKID = "
                        strSQL = strSQL & lvw_Empenho.ListItems(intInd).Tag & ";"
                    End If
                    
                    strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
                    strSQL = strSQL & "dtmPagamento = " & IIf(mblnEstorno, "NULL", gstrConvDtParaSql(txtData)) & ", "
                    strSQL = strSQL & "bytSituacao = " & IIf(mblnEstorno, 2, 3) & ", "
                    strSQL = strSQL & "intProcesso = " & IIf(mblnEstorno, "NULL", lngProcesso) & " "
                    strSQL = strSQL & "WHERE PKId = " & lvw_Empenho.ListItems(intInd).Tag & "; "
                                
                    IDOrdemPagamento = gstrOrdemPagamentonoGrid(intInd)
                    
                    strSQL = strSQL & "INSERT INTO " & gstrPagamentoEstornoEmpenho & " ("
                    strSQL = strSQL & "intParcela,intProcesso,dblValor,intOrdemPagamento,"
                    strSQL = strSQL & "dtmData,dtmDtAtualizacao,lngCodUsr)"
                    strSQL = strSQL & "SELECT PKId, "
                    strSQL = strSQL & CStr(lngProcesso) & ","
                    strSQL = strSQL & IIf(mblnEstorno, gstrConvVrParaSql(CDbl(lvw_Empenho.ListItems(intInd).SubItems(6)) * -1), gstrConvVrParaSql(lvw_Empenho.ListItems(intInd).SubItems(6))) & ", "
                    strSQL = strSQL & IIf(IDOrdemPagamento = 0, "NULL", IDOrdemPagamento) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(txtData) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSQL = strSQL & glngCodUsr & " FROM " & gstrSubempenho & " "
                    strSQL = strSQL & "WHERE PKId = " & lvw_Empenho.ListItems(intInd).Tag & "; "
                    
                Next
                
                For intInd = 1 To lvw_Resto.ListItems.Count
                    If mblnEstorno Then
                        'Registra o estorno do pagamento das parcela do resto
                        strSQL = strSQL & "INSERT INTO " & gstrParcelaRestoPagtoAnulado & " ("
                        strSQL = strSQL & "intSubEmpenho,intProcesso,intProcessoOriginal,"
                        strSQL = strSQL & "dtmDataEstorno,dtmDtAtualizacao,lngCodUsr)"
                        strSQL = strSQL & "SELECT PKId, "
                        strSQL = strSQL & CStr(lngProcesso) & " ,intProcesso,"
                        strSQL = strSQL & gstrConvDtParaSql(txtData) & ", "
                        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                        strSQL = strSQL & glngCodUsr & " FROM " & gstrSubempenho & " "
                        strSQL = strSQL & "WHERE PKID = "
                        strSQL = strSQL & lvw_Resto.ListItems(intInd).Tag & ";"
                    End If
                    
                    strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
                    strSQL = strSQL & "dtmPagamento = " & IIf(mblnEstorno, "NULL", gstrConvDtParaSql(txtData)) & ", "
                    strSQL = strSQL & "bytSituacao = " & IIf(mblnEstorno, 2, 3) & ", "
                    strSQL = strSQL & "intProcesso = " & IIf(mblnEstorno, "NULL", lngProcesso) & " "
                    strSQL = strSQL & "WHERE PKId = " & lvw_Resto.ListItems(intInd).Tag & ";"
                    
                    IDOrdemPagamento = gstrOrdemPagamentonoGrid(intInd)
                    
                    strSQL = strSQL & "INSERT INTO " & gstrPagamentoEstornoEmpenho & " ("
                    strSQL = strSQL & "intParcela,intProcesso,dblValor,intOrdemPagamento,"
                    strSQL = strSQL & "dtmData,dtmDtAtualizacao,lngCodUsr)"
                    strSQL = strSQL & "SELECT PKId, "
                    strSQL = strSQL & CStr(lngProcesso) & ","
                    strSQL = strSQL & IIf(mblnEstorno, gstrConvVrParaSql(CDbl(lvw_Resto.ListItems(intInd).SubItems(6)) * -1), gstrConvVrParaSql(lvw_Resto.ListItems(intInd).SubItems(6))) & ", "
                    strSQL = strSQL & IIf(IDOrdemPagamento = 0, "NULL", IDOrdemPagamento) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(txtData) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSQL = strSQL & glngCodUsr & " FROM " & gstrSubempenho & " "
                    strSQL = strSQL & "WHERE PKId = " & lvw_Resto.ListItems(intInd).Tag & ";"
                    
                Next
                
                For intInd = 1 To lvw_Despesa.ListItems.Count
                    If mblnEstorno Then
                        'Registra o estorno do pagamento da Despesa Extra-orçamentária
                        strSQL = strSQL & "INSERT INTO " & gstrDespesaExtraOrcamPagtoAnulado & " ("
                        strSQL = strSQL & "intDespesa,intProcesso,intProcessoOriginal,"
                        strSQL = strSQL & "dtmDataEstorno,dtmDtAtualizacao,lngCodUsr)"
                        strSQL = strSQL & " SELECT PKId, "
                        strSQL = strSQL & CStr(lngProcesso) & " ,intProcesso,"
                        strSQL = strSQL & gstrConvDtParaSql(txtData) & ", "
                        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                        strSQL = strSQL & glngCodUsr & " FROM " & gstrDespesaExtraOrcamentaria & " "
                        strSQL = strSQL & "WHERE PKID = "
                        strSQL = strSQL & lvw_Despesa.ListItems(intInd).Tag & ";"
                    End If
                    
                    strSQL = strSQL & "UPDATE " & gstrDespesaExtraOrcamentaria & " SET "
                    strSQL = strSQL & "dtmPagamento = " & IIf(mblnEstorno, "NULL", gstrConvDtParaSql(txtData)) & ", "
                    strSQL = strSQL & "bytSituacao = " & IIf(mblnEstorno, 0, 2) & ", "
                    strSQL = strSQL & "intProcesso = " & IIf(mblnEstorno, "NULL", lngProcesso) & " "
                    strSQL = strSQL & "WHERE PKId = " & lvw_Despesa.ListItems(intInd).Tag & ";"
                    
                    IDOrdemPagamento = gstrOrdemPagamentonoGrid(intInd)
                    
                    strSQL = strSQL & "INSERT INTO " & gstrPagamentoEstornoEmpenho & " ("
                    strSQL = strSQL & "intDespesaExtra,intProcesso,dblValor,intOrdemPagamento,"
                    strSQL = strSQL & "dtmData,dtmDtAtualizacao,lngCodUsr)"
                    strSQL = strSQL & "SELECT PKId, "
                    strSQL = strSQL & CStr(lngProcesso) & ","
                    strSQL = strSQL & IIf(mblnEstorno, gstrConvVrParaSql(CDbl(lvw_Despesa.ListItems(intInd).SubItems(4)) * -1), gstrConvVrParaSql(lvw_Despesa.ListItems(intInd).SubItems(4))) & ", "
                    strSQL = strSQL & IIf(IDOrdemPagamento = 0, "NULL", IDOrdemPagamento) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(txtData) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSQL = strSQL & glngCodUsr & " FROM " & gstrDespesaExtraOrcamentaria & " "
                    strSQL = strSQL & "WHERE PKId = " & lvw_Despesa.ListItems(intInd).Tag & ";"

                Next
                
                
                For intInd = 1 To lvw_AnulacaoReceita.ListItems.Count
                    If mblnEstorno Then
                        'Registra o estorno do pagamento da Anulacao da Receita
                        strSQL = strSQL & "INSERT INTO " & gstrAnulacaoRecPagtoAnulado & " ("
                        strSQL = strSQL & "intOrdemPagamentoAnulacaoRec,intProcesso,intProcessoOriginal,"
                        strSQL = strSQL & "dtmDataEstorno,dtmDtAtualizacao,lngCodUsr)"
                        strSQL = strSQL & " SELECT PKId, "
                        strSQL = strSQL & CStr(lngProcesso) & " ,intProcesso,"
                        strSQL = strSQL & gstrConvDtParaSql(txtData) & ", "
                        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                        strSQL = strSQL & glngCodUsr & " FROM " & gstrOrdemPagamentoAnulacaoReceita & " "
                        strSQL = strSQL & "WHERE PKID = "
                        strSQL = strSQL & lvw_AnulacaoReceita.ListItems(intInd).Tag & ";"
                    End If
                    
                    strSQL = strSQL & "UPDATE " & gstrOrdemPagamentoAnulacaoReceita & " SET "
                    strSQL = strSQL & "dtmPagamento = " & IIf(mblnEstorno, "NULL", gstrConvDtParaSql(txtData)) & ", "
                    strSQL = strSQL & "bytSituacao = " & IIf(mblnEstorno, 0, 2) & ", "
                    strSQL = strSQL & "intProcesso = " & IIf(mblnEstorno, "NULL", lngProcesso) & " "
                    strSQL = strSQL & "WHERE PKId = " & lvw_AnulacaoReceita.ListItems(intInd).Tag & ";"
                Next
                
                
                
                For intInd = 1 To lvw_Extra.ListItems.Count
                    strSQL = strSQL & "INSERT INTO " & gstrLancamentoContabil & " ("
                    strSQL = strSQL & "intProcesso, intConta, intParcela, dblValor, bytNatureza, "
                    strSQL = strSQL & "bytTipo,dtmDtAtualizacao, lngCodUsr)"
                    strSQL = strSQL & " VALUES "
                    strSQL = strSQL & "(" & lngProcesso & ", "
                    strSQL = strSQL & gstrConta(lvw_Extra.ListItems(intInd).Tag, False) & ", "
                    strSQL = strSQL & gstrConta(lvw_Extra.ListItems(intInd).Tag, True) & ", "
                    'strSQL = strSQL & IIf(mblnEstorno, gstrConvVrParaSql(CDbl(lvw_Extra.ListItems(intInd).SubItems(4))) * -1, gstrConvVrParaSql(lvw_Extra.ListItems(intInd).SubItems(4))) & ", "
                    strSQL = strSQL & IIf(mblnEstorno, gstrConvVrParaSql(CDbl(lvw_Extra.ListItems(intInd).SubItems(4)) * -1), gstrConvVrParaSql(lvw_Extra.ListItems(intInd).SubItems(4))) & ", "
                    strSQL = strSQL & "0, 0,"
                    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
                    strSQL = strSQL & glngCodUsr & ");"
                Next
                'Removido por M4RC3LØ 12/04/2003
                'If Val(intTipoEventoSelecionado) <> 11 Then
                    strSQL = strSQL & "UPDATE " & gstrOrdemPagamento & " SET "
                    strSQL = strSQL & "blnPago = " & IIf(mblnEstorno, "0, ", "1 , ")
                    'Grava Data de Pagamento M6R
                    strSQL = strSQL & "dtmPagamento = " & IIf(mblnEstorno, "NULL", gstrConvDtParaSql(txtData))
                    strSQL = strSQL & " WHERE PKId IN( " & gstrOrdemPagamentonoGrid & "); "
                    
                    strSQL = strSQL & "UPDATE " & gstrcheque & " SET "
                    strSQL = strSQL & " strFlag = " & IIf(mblnEstorno, "0", "1")
                    strSQL = strSQL & " WHERE PKId IN( "
                    strSQL = strSQL & " Select CH.pkid FROM "
                    strSQL = strSQL & gstrcheque & " CH, "
                    strSQL = strSQL & gstrchequeOP & " CHOP "
                    strSQL = strSQL & " WHERE "
                    strSQL = strSQL & " CH.pkid = CHOP.intCheque AND "
                    strSQL = strSQL & " CHOP.intordemPagamento in ( " & gstrOrdemPagamentonoGrid & "));"
                    
                'End If
                
                strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
                
                Set gobjBanco = New clsBanco
                If gobjBanco.Execute(strSQL) Then
                    
                    ReDim aryContas(1 To lvw_Conta.ListItems.Count + lvw_Extra.ListItems.Count + lvw_Despesa.ListItems.Count)
                    ReDim aryTpMov(1 To lvw_Conta.ListItems.Count + lvw_Extra.ListItems.Count + lvw_Despesa.ListItems.Count)
                    ReDim aryValor(1 To lvw_Conta.ListItems.Count + lvw_Extra.ListItems.Count + lvw_Despesa.ListItems.Count)
                    
                    With lvw_Conta
                      
                      For intIdxConta = 1 To .ListItems.Count
                         .ListItems(intIdxConta).Selected = True
                         aryContas(intIdxConta) = Val(.SelectedItem.Tag)
                         aryTpMov(intIdxConta) = 0 'Crédito
                         
                         aryValor(intIdxConta) = Replace(Str( _
                         IIf(mblnEstorno, _
                         CDbl(.ListItems(intIdxConta).SubItems(2)) * -1, _
                         CDbl(.ListItems(intIdxConta).SubItems(2))) _
                         ), ".", ",")
                         
                      Next
                    End With
                    
                    intIdxConta = intIdxConta - 1
                    
                    With lvw_Extra
                      For intIdxExtra = 1 To .ListItems.Count
                         .ListItems(intIdxExtra).Selected = True
                         aryContas(intIdxExtra + intIdxConta) = Val(gstrConvVrParaSql(.ListItems(intIdxExtra).SubItems(6)))
                         aryTpMov(intIdxExtra + intIdxConta) = 0 'Crédito
                         aryValor(intIdxExtra + intIdxConta) = Replace(Str( _
                         IIf(mblnEstorno, _
                         CDbl(.ListItems(intIdxExtra).SubItems(4)) * -1, _
                         CDbl(.ListItems(intIdxExtra).SubItems(4))) _
                         ), ".", ",")
                      Next
                    End With
                    
                    intIdxConta = intIdxExtra + intIdxConta - 1
                    
                    With lvw_Despesa
                      For intIdxExtra = 1 To .ListItems.Count
                         .ListItems(intIdxExtra).Selected = True
                         aryContas(intIdxExtra + intIdxConta) = RetornaContaDespesa(.ListItems(intIdxExtra).Tag)
                         aryTpMov(intIdxExtra + intIdxConta) = 1 'Debito
                         aryValor(intIdxExtra + intIdxConta) = Replace(Str( _
                         IIf(mblnEstorno, _
                         CDbl(.ListItems(intIdxExtra).SubItems(4)) * -1, _
                         CDbl(.ListItems(intIdxExtra).SubItems(4))) _
                         ), ".", ",")
                      Next
                    End With
                    
                    
                    
                    If tab_3DPastaEmpenho.TabEnabled(0) = True Then
                       strTotal = txtTotalEmpenho
                    ElseIf tab_3DPastaEmpenho.TabEnabled(1) = True Then
                       strTotal = txtTotalResto
                    ElseIf tab_3DPastaEmpenho.TabEnabled(2) = True Then
                       strTotal = txtTotalDespesa
                    ElseIf tab_3DPastaEmpenho.TabEnabled(3) = True Then
                        strTotal = txtTotal
                    ElseIf tab_3DPastaEmpenho.TabEnabled(5) = True Then
                        strTotal = txtTotalRecOrdem
                    End If
                    
                    strTotal = gstrConvVrDoSql(Abs(Val(gstrConvVrParaSql(strTotal))))
                    
                    If Not GeraMovimentosByEvento(gstrItemData(cbo_intEvento), txtData, IIf(mblnEstorno, "-", "") & Str(CDbl(strTotal)), txtHistoricoLiquidacao, Str(lngProcesso), "7", aryContas, aryTpMov, IIf(lvw_Despesa.ListItems.Count > 0, True, False), aryValor, True) Then
                        ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                        Set gobjBanco = New clsBanco
                        gobjBanco.ExecutaRollbackTrans
                    Else
                        Set gobjBanco = New clsBanco
                        gobjBanco.ExecutaCommitTrans
                    
                        If chkBordero Then
                            ImprimeBordereaux lngProcesso
                        End If
    
                        LeDaTabelaParaObj "", tdb_Lista, gstrQueryLocalizar
                        
                        'Impressão da Autenticação de Movimentação
                        If Not mblnAlterando Then
                            If blnFitaAutenticadoraOK Then
                                'Inicio da Impressão
                                strRegistroAutent = Right(String(6, "0") & Me.txtintProcesso.Text, 6) & " " & Format(Me.txtData.Text, "dd/mm/yyyy") & " " & String(6, "0") & " "
                                'Pagamentos
                                If tab_3DPastaEmpenho.TabEnabled(0) = True Then
                                    'Empenho
                                    With Me.lvw_Empenho
                                        For intInd = 1 To .ListItems.Count
                                            ImprimeFitaAutenticadora strRegistroAutent & Right(String(13 - Len(.ListItems.Item(intInd).SubItems(8)), "0"), 13) & gstrConvVrParaSql(.ListItems.Item(intInd).SubItems(8)) & " OPO"
                                        Next
                                    End With
                                End If
                                If tab_3DPastaEmpenho.TabEnabled(1) = True Then
                                    'Restos à Pagar
                                    With Me.lvw_Resto
                                        For intInd = 1 To .ListItems.Count
                                            ImprimeFitaAutenticadora strRegistroAutent & Right(String(13 - Len(.ListItems.Item(intInd).SubItems(8)), "0"), 13) & gstrConvVrParaSql(.ListItems.Item(intInd).SubItems(8)) & " OPR"
                                        Next
                                    End With
                                End If
                                If tab_3DPastaEmpenho.TabEnabled(2) = True Then
                                    'Despesa Extra-Orçamentária
                                    With Me.lvw_Despesa
                                        For intInd = 1 To .ListItems.Count
                                            ImprimeFitaAutenticadora strRegistroAutent & Right(String(13 - Len(.ListItems.Item(intInd).SubItems(4)), "0"), 13) & gstrConvVrParaSql(.ListItems.Item(intInd).SubItems(4)) & " OPE"
                                        Next
                                    End With
                                End If
                                If tab_3DPastaEmpenho.TabEnabled(3) = True Then
                                    'Lançamento
                                    strRegistroAutent = Right(String(6, "0") & Me.txtintProcesso.Text, 6) & " " & Format(Me.txtData.Text, "dd/mm/yyyy") & " "
                                    'strRegistroAutent = strRegistroAutenticacao(Me.txtintProcesso.Text)
                                    With Me.lvw_Conta
                                        For intInd = 1 To .ListItems.Count
                                            ImprimeFitaAutenticadora strRegistroAutent & Right(String(6, "0") & lvw_Conta.ListItems.Item(intInd), 6) & " " & Right(String(13 - Len(.ListItems.Item(intInd).SubItems(2)), "0") & gstrConvVrParaSql(.ListItems.Item(intInd).SubItems(2)), 13) & " BCO"
                                        Next
                                    End With
                                End If
                            End If
                        End If
                        
                        'estas lista mudam de status durante a gravação, por isso são atualizadas
                        If tab_3DPastaEmpenho.TabEnabled(0) = True Then
                            LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
                            LeDaTabelaParaObj "", dcbOrdemPagamentoEmpenho, strQueryOrdemPagamentoEmpenho
                        End If
                        
                        If tab_3DPastaEmpenho.TabEnabled(1) = True Then
                            LeDaTabelaParaObj "", dcbResto, strQueryResto
                            If Trim(txt_intExercicioRP) <> "" Then
                                LeDaTabelaParaObj "", dcbOrdemPagamentoResto, strQueryOrdemPagamentoResto
                            End If
                        End If
                        
                        If tab_3DPastaEmpenho.TabEnabled(2) = True Then
                            LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
                            If Trim(txt_intExercicioDE) <> "" And cbo_intEvento.ListIndex <> -11 Then
                                LeDaTabelaParaObj "", dcbOrdemPagamentoDespesa, strQueryOrdemPagamentoDespesa
                            End If
                            
                        End If
                        
                        If tab_3DPastaEmpenho.TabEnabled(4) = True Then
                            LeDaTabelaParaObj "", dcbOrdemPagamentoAnulacao, strQueryOrdemPagamentoAnulacao
                        End If
                        
                        LimpaTelaPagamento True
                    End If
                Else
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaRollbackTrans
                End If
            End If
        End If ' if do blndadosok
        Else 'se mblalterando = true
            If gblnExclusaoGravacaoOk("A", "Confirma alteração do histórico do pagamento?", True) Then
                strSQL = ""
                strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
                
                strSQL = strSQL & "UPDATE " & gstrProcessoPagamento
                strSQL = strSQL & " SET strHistorico = '" & txtHistoricoLiquidacao.Text & "'"
                strSQL = strSQL & " WHERE PKID = " & txtPKID.Text & ";"
                If blnIncremtCheque Then
                    With lvw_Conta
                        For intInd = 1 To .ListItems.Count
                            strSQL = strSQL & " UPDATE " & gstrLancamentoContabil & " SET"
                            strSQL = strSQL & " strDocumento = '" & lvw_Conta.ListItems.Item(intInd).ListSubItems(3) & "'"
                            strSQL = strSQL & " WHERE intProcesso = " & Me.txtPKID.Text
                            strSQL = strSQL & " AND intConta = " & .ListItems(intInd).Tag & ";"
                        Next
                    End With
                End If
                
                Set gobjBanco = New clsBanco
                
                strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
                
                gobjBanco.Execute (strSQL)
            End If
        End If
    
End Sub

Private Function blnDadoContaOk() As Boolean
    If cbointContaContabil.ListIndex = -1 Then
        ExibeMensagem "A conta tem que ser informada corretamente."
        If cbointContaContabil.Enabled Then cbointContaContabil.SetFocus
        Exit Function
    ElseIf Val(gstrConvVrParaSql(txtValorLancamento)) = 0 Then
        ExibeMensagem "O valor do lançamento tem que ser informado corretamente."
        If txtValorLancamento.Enabled Then txtValorLancamento.SetFocus
        Exit Function
    ElseIf blnVerificaGrid Then
        ExibeMensagem "Não é permitido inclusão de mesma Conta com mesmo Número de cheque."
        Exit Function
    ElseIf cbostrContaContabil.ListIndex = -1 Then
        ExibeMensagem "A descrição da conta tem que ser informada corretamente."
        If cbostrContaContabil.Enabled Then cbostrContaContabil.SetFocus
        Exit Function
    End If
    
    If chkEstorno.Value = 0 Then
        If strValorAtualdaConta = -1 Then Exit Function
    End If

    blnDadoContaOk = True
End Function

Private Sub ProcuraConta()
    If gblnEncontroItemNoListView(lvw_Conta, gstrItemData(cbointContaContabil), lvwTag) Then
        mblnAlterandoConta = True
    Else
        mblnAlterandoConta = False
    End If


End Sub

Private Sub TotalLancado()
    Dim intInd      As Integer
    Dim dblTotal    As Double
    For intInd = 1 To lvw_Conta.ListItems.Count
        dblTotal = dblTotal + Val(gstrConvVrParaSql(lvw_Conta.ListItems(intInd).SubItems(2)))
    Next
    txtTotal = gstrConvVrDoSql(IIf(chkEstorno.Value = 1, dblTotal * -1, dblTotal))
End Sub

Private Sub LimpaDadosEmpenho(Optional blnNaoSetaFoco As Boolean)
    If blnNaoSetaFoco = False Then
        If dcbParcela.Enabled Then dcbParcela.SetFocus
    End If
    dcbParcela = ""
    dcbEmpenho.Text = ""
    mblnAlterandoEmpenho = False
    dcbOrdemPagamentoEmpenho.Text = ""
End Sub

Private Sub LimpaDadosResto(Optional blnNaoSetaFoco As Boolean)
    If blnNaoSetaFoco = False Then
        If dcbParcelaResto.Enabled Then dcbParcelaResto.SetFocus
    End If
    dcbParcelaResto = ""
    dcbOrdemPagamentoResto.Text = ""
    mblnAlterandoResto = False
End Sub

Private Sub LimpaDadosDespesa(Optional blnNaoSetaFoco As Boolean)
    If blnNaoSetaFoco = False Then
        If dcbDespesa.Enabled Then dcbDespesa.SetFocus
    End If
    dcbDespesa = ""
    mblnAlterandoDespesa = False
    dcbOrdemPagamentoDespesa.Text = ""
End Sub

Private Sub LimpaDadosConta(Optional blnNaoSetaFoco As Boolean)
    If blnNaoSetaFoco = False Then
        If cbointContaContabil.Enabled Then cbointContaContabil.SetFocus
    End If
    cbointContaContabil.ListIndex = -1
    txtValorLancamento = ""
    txtNumCheque = ""
    txt_saldoAtual = ""
    mblnAlterandoConta = False
End Sub

Private Sub LimpaDadosAnulacao(Optional blnNaoSetaFoco As Boolean)
    If blnNaoSetaFoco = False Then
        If dcbOrdemPagamentoAnulacao.Enabled Then dcbOrdemPagamentoAnulacao.SetFocus
    End If
    dcbOrdemPagamentoAnulacao.Text = ""
    mblnAlterandoAnulacao = False
End Sub


Private Sub IncluiAlteraConta()
    If blnDadoContaOk Then
        If mblnAlterandoConta Then
            lvw_Conta.ListItems(lvw_Conta.SelectedItem.Index).Text = cbointContaContabil
            lvw_Conta.SelectedItem.SubItems(1) = cbostrContaContabil
            lvw_Conta.SelectedItem.SubItems(2) = gstrConvVrDoSql(txtValorLancamento)
            lvw_Conta.SelectedItem.SubItems(3) = Trim(txtNumCheque)
            lvw_Conta.SelectedItem.Tag = gstrItemData(cbointContaContabil)
            mblnAlterandoConta = False
        Else
            Set mobjLista = lvw_Conta.ListItems.Add(, , cbointContaContabil)
            mobjLista.SubItems(1) = cbostrContaContabil
            mobjLista.SubItems(2) = gstrConvVrDoSql(txtValorLancamento)
            mobjLista.SubItems(3) = Trim(txtNumCheque)
            mobjLista.Tag = gstrItemData(cbointContaContabil)
        End If
        TotalLancado
        LimpaDadosConta
    End If
End Sub

Private Sub ExcluiContaCheque(ByVal strOPNumero As String, ByVal strOPExercicio As String)

    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset

    If chkEstorno.Value = 1 Then Exit Sub
    
    strSQL = ""
    strSQL = strSQL & "Select pc.pkid, "
    strSQL = strSQL & " cb.intnumeroconta, "
    strSQL = strSQL & " cb.strdescricao, "
    strSQL = strSQL & " CH.dblvalor,"
    strSQL = strSQL & " CH.strCheque "
    strSQL = strSQL & " FROM " & gstrPlanoConta & " PC,"
    strSQL = strSQL & gstrContaBancaria & " CB,"
    strSQL = strSQL & gstrchequeOP & " CHOP,"
    strSQL = strSQL & gstrcheque & " CH "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PC.intcontabancaria = CB.PKID AND "
    strSQL = strSQL & " CH.intcontabancaria = CB.PKID AND "
    strSQL = strSQL & " CH.PKID = CHOP.intCheque AND "
    strSQL = strSQL & " CH.strFlag = 0 AND "
    strSQL = strSQL & " CH.bytCancelado = 0 AND "
    strSQL = strSQL & " CHOP.intOrdemPagamento IN ( "
    strSQL = strSQL & " SELECT PKID FROM " & gstrOrdemPagamento
    strSQL = strSQL & " WHERE intnumero = " & strOPNumero & " AND "
    strSQL = strSQL & " intExercicio = " & strOPExercicio
    strSQL = strSQL & "  ) "
    
    

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
       With adoResultado
          If Not .EOF Then
                'checagem para ver: se o cheque já não foi removido
                If gblnEncontroItemNoListView(lvw_Conta, !Pkid, lvwTag) Then 'encontra o banco na lista
                    If gblnEncontroItemNoListView(lvw_Conta, !strCheque, lvwSubItem, lvwWhole) Then 'encontra o cheque na lista
                        ExcluiItemLista lvw_Conta, txtTotal
                        Exit Sub
                    End If
                End If

                TotalLancado
          End If
       End With
    End If


End Sub



Private Sub IncluiContaCheque(ByVal strOPpkid As String)

    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset

'    If chkEstorno.Value = 1 Then Exit Sub
    
    strSQL = ""
    strSQL = strSQL & "Select pc.pkid, "
    strSQL = strSQL & " cb.intnumeroconta, "
    strSQL = strSQL & " cb.strdescricao, "
    strSQL = strSQL & " CH.dblvalor,"
    strSQL = strSQL & " CH.strCheque "
    strSQL = strSQL & " FROM " & gstrPlanoConta & " PC,"
    strSQL = strSQL & gstrContaBancaria & " CB,"
    strSQL = strSQL & gstrchequeOP & " CHOP,"
    strSQL = strSQL & gstrcheque & " CH "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PC.intcontabancaria = CB.PKID AND "
    strSQL = strSQL & " CH.intcontabancaria = CB.PKID AND "
    strSQL = strSQL & " CH.PKID = CHOP.intCheque AND "
    strSQL = strSQL & " CH.strFlag = " & chkEstorno.Value & " AND "
    strSQL = strSQL & " CH.bytCancelado = 0 AND "
    strSQL = strSQL & " CHOP.intOrdemPagamento = '" & strOPpkid & "' "
    

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
       With adoResultado
          If Not .EOF Then
        
                IncluiOpsRelacionadas strOPpkid
                            
                'checagem para ver: se o cheque já tiver sido inserido nao é re-inserido
                If gblnEncontroItemNoListView(lvw_Conta, !Pkid, lvwTag) Then 'encontra o banco na lista
                    If gblnEncontroItemNoListView(lvw_Conta, !strCheque, lvwSubItem, lvwWhole) Then Exit Sub 'encontra o cheque na lista
                End If
                
                Screen.MousePointer = vbHourglass
                If chkEstorno.Value = 0 Then
                    If blnValorAtualdaContaCheque(!Pkid, !intNumeroConta, !dblValor) = False Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                Screen.MousePointer = vbDefault
        
                Set mobjLista = lvw_Conta.ListItems.Add(, , !intNumeroConta)
                mobjLista.SubItems(1) = !strDescricao
                mobjLista.SubItems(2) = gstrConvVrDoSql(!dblValor)
                mobjLista.SubItems(3) = Trim(!strCheque)
                mobjLista.Tag = !Pkid
                TotalLancado
            
        
          End If
       End With
    End If


End Sub

Private Function gstrQueryChequeEmpenho(ByVal strPKidOP As String, _
                                  Optional ByVal numeroOP As String = "", _
                                  Optional ByVal ExercicioOP As String = "") As String
    Dim strSQL As String
    
    If numeroOP <> "" Then strPKidOP = " SELECT PKID FROM " & gstrOrdemPagamento & " WHERE intnumero = " & numeroOP & " AND intExercicio = " & ExercicioOP
    
        strSQL = ""
        strSQL = "SELECT SE.Pkid Pkid, "
        strSQL = strSQL & "(SELECT OP.intNumero FROM " & gstrOrdemPagamentoEmpenho & " OPR, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OPR.intOrdemPagamento AND OPR.intParcela = SE.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intOrdem,"
        strSQL = strSQL & "(SELECT OP.intExercicio FROM " & gstrOrdemPagamentoEmpenho & " OPR, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OPR.intOrdemPagamento AND OPR.intParcela = SE.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intExercicioOP,"
        strSQL = strSQL & "E.IntNumero intEmpenho, "
        strSQL = strSQL & "SE.IntNumero intNumero, "
        strSQL = strSQL & "SE.dblvalor dblValor, "
        strSQL = strSQL & "SE.dtmData dtmPrevisao, "
        strSQL = strSQL & "SE.dtmLiquidacao dtmLiquidacao, "
        strSQL = strSQL & gstrISNULL("SUM(SL.dblvalor)", "0") & " dblDesconto, "
        strSQL = strSQL & gstrISNULL("SE.dblvalor", "0") & " - "
        strSQL = strSQL & gstrISNULL("SUM(Sl.dblValor)", "0") & " dblLiquido, "
        strSQL = strSQL & "TE.bytAdiantamento "
        strSQL = strSQL & "FROM " & gstrEmpenho & " E ," & gstrTipoEmpenho & " TE,"
        strSQL = strSQL & gstrSubempenho & " SE, " & gstrSubempenhoLiquidado & " SL"
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & "SE.Pkid IN ("

        strSQL = strSQL & "SELECT ope.intparcela "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrcheque & " CH , "
        strSQL = strSQL & gstrchequeOP & " CHOP, "
        strSQL = strSQL & gstrchequeOP & " CHOP1, "
        strSQL = strSQL & gstrOrdemPagamentoEmpenho & " ope "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "CH.PKID = CHOP.INTCHEQUE AND "
        strSQL = strSQL & "CH.PKID = CHOP1.INTCHEQUE AND "
        strSQL = strSQL & "ope.intordempagamento = CHOP.Intordempagamento AND "
        strSQL = strSQL & "CHOP1.INTORDEMPAGAMENTO in ( " & strPKidOP & ") AND "
        strSQL = strSQL & "NOT CHOP.INTORDEMPAGAMENTO in ( " & strPKidOP & ")"
        
        strSQL = strSQL & ") AND "
        strSQL = strSQL & " TE.PKID = E.Inttipo "
        strSQL = strSQL & " AND E.Pkid = SE.intEmpenho AND "
        strSQL = strSQL & "SE.Pkid " & strOUTJSQLServer & "= SL.intParcela" & strOUTJOracle
        strSQL = strSQL & " GROUP BY SE.Pkid, E.IntNumero, SE.IntNumero, SE.dblvalor,SE.dtmData,SE.dtmLiquidacao,TE.bytAdiantamento"

    gstrQueryChequeEmpenho = strSQL
End Function

Private Function gstrQueryChequeRestoAPagar(ByVal strPKidOP As String, _
                                  Optional ByVal numeroOP As String = "", _
                                  Optional ByVal ExercicioOP As String = "") As String
                                  
    Dim strSQL As String
    
    If numeroOP <> "" Then strPKidOP = " SELECT PKID FROM " & gstrOrdemPagamento & " WHERE intnumero = " & numeroOP & " AND intExercicio = " & ExercicioOP
    
        strSQL = ""
        strSQL = "SELECT SE.Pkid Pkid, "
        strSQL = strSQL & "(SELECT OP.intNumero FROM " & gstrOrdemPagamentoResto & " OPR, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OPR.intOrdemPagamento AND OPR.intParcela = SE.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intOrdem,"
        strSQL = strSQL & "(SELECT OP.intExercicio FROM " & gstrOrdemPagamentoResto & " OPR, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OPR.intOrdemPagamento AND OPR.intParcela = SE.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intExercicioOP,"
        strSQL = strSQL & "E.intExercicioRP, "
        strSQL = strSQL & gstrDATEPART(strYEAR, "E.dtmData") & " intExercicio, "
        strSQL = strSQL & "E.IntNumero intResto, "
        strSQL = strSQL & "SE.IntNumero intNumero, "
        strSQL = strSQL & "SE.dblvalor dblValor, "
        strSQL = strSQL & "SE.dtmData dtmPrevisao, "
        strSQL = strSQL & gstrISNULL("SUM(SL.dblvalor)", "0") & " dblDesconto, "
        strSQL = strSQL & gstrISNULL("SE.dblvalor", "0") & " - "
        strSQL = strSQL & gstrISNULL("SUM(Sl.dblValor)", "0") & " dblLiquido "
        strSQL = strSQL & "FROM " & gstrEmpenho & " E ,"
        strSQL = strSQL & gstrSubempenho & " SE, " & gstrSubempenhoLiquidado & " SL"
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & "SE.Pkid IN ("
        
        strSQL = strSQL & "SELECT ope.intparcela "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrcheque & " CH , "
        strSQL = strSQL & gstrchequeOP & " CHOP, "
        strSQL = strSQL & gstrchequeOP & " CHOP1, "
        strSQL = strSQL & gstrOrdemPagamentoResto & " ope "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "CH.PKID = CHOP.INTCHEQUE AND "
        strSQL = strSQL & "CH.PKID = CHOP1.INTCHEQUE AND "
        strSQL = strSQL & "ope.intordempagamento = CHOP.Intordempagamento AND "
        strSQL = strSQL & "CHOP1.INTORDEMPAGAMENTO in ( " & strPKidOP & ") AND "
        strSQL = strSQL & "NOT CHOP.INTORDEMPAGAMENTO in ( " & strPKidOP & ")"

        
        strSQL = strSQL & ") AND "
        strSQL = strSQL & "E.Pkid = SE.intEmpenho AND "
        strSQL = strSQL & "SE.Pkid " & strOUTJSQLServer & "= SL.intParcela" & strOUTJOracle & " And "
        strSQL = strSQL & "E.intExercicioRP =" & CStr(gintExercicio)
        strSQL = strSQL & " GROUP BY SE.Pkid, E.intExercicioRP, E.IntNumero, SE.IntNumero,SE.dblvalor ,SE.dtmData ," & gstrDATEPART(strYEAR, "E.dtmData")
    
    
    gstrQueryChequeRestoAPagar = strSQL
End Function

Private Function gstrQueryChequeExtra(ByVal strPKidOP As String, _
                                  Optional ByVal numeroOP As String = "", _
                                  Optional ByVal ExercicioOP As String = "") As String
                                    
    Dim strSQL As String
    
        If numeroOP <> "" Then strPKidOP = " SELECT PKID FROM " & gstrOrdemPagamento & " WHERE intnumero = " & numeroOP & " AND intExercicio = " & ExercicioOP
    
    
        strSQL = strSQL & "SELECT DP.PKId,"
        strSQL = strSQL & "(SELECT OP.intNumero FROM " & gstrOrdemPagamentoDespesaExtra & " OD, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OD.intOrdemPagamento AND OD.intDespesaExtraOrcamentaria = DP.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intOrdem,"
        strSQL = strSQL & "(SELECT OP.intExercicio FROM " & gstrOrdemPagamentoDespesaExtra & " OD, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OD.intOrdemPagamento AND OD.intDespesaExtraOrcamentaria = DP.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intExercicioOP,"
        strSQL = strSQL & " DP.intNumero , DP.dtmData, "
        strSQL = strSQL & "DP.dblValor, CT.strNome FROM "
        strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DP, "
        strSQL = strSQL & gstrContribuinte & " CT "
        strSQL = strSQL & "WHERE DP.intContribuinte = CT.PKId AND "
        strSQL = strSQL & "DP.Pkid IN ("
        
        strSQL = strSQL & "SELECT ope.intdespesaextraorcamentaria "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrcheque & " CH , "
        strSQL = strSQL & gstrchequeOP & " CHOP, "
        strSQL = strSQL & gstrchequeOP & " CHOP1, "
        strSQL = strSQL & gstrOrdemPagamentoDespesaExtra & " ope "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "CH.PKID = CHOP.INTCHEQUE AND "
        strSQL = strSQL & "CH.PKID = CHOP1.INTCHEQUE AND "
        strSQL = strSQL & "ope.intordempagamento = CHOP.Intordempagamento AND "
        strSQL = strSQL & "CHOP1.INTORDEMPAGAMENTO in ( " & strPKidOP & ") AND "
        strSQL = strSQL & "NOT CHOP.INTORDEMPAGAMENTO in ( " & strPKidOP & ")"
        strSQL = strSQL & ")"
    
    
    gstrQueryChequeExtra = strSQL
End Function
    
    
Private Sub ExcluiOpsRelacionadas(ByVal strPKidOP As String, _
                                  Optional ByVal numeroOP As String = "", _
                                  Optional ByVal ExercicioOP As String = "")
    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset
   
    ExcluiContaCheque numeroOP, ExercicioOP
   
    Select Case tab_3DPastaEmpenho.Tab
    Case 0
            
        strSQL = gstrQueryChequeEmpenho(strPKidOP, numeroOP, ExercicioOP)
                    
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF
                    If gblnEncontroItemNoListView(lvw_Empenho, !Pkid, lvwTag) Then 'encontra o banco na lista
                        ExcluiItemLista lvw_Empenho, txtTotalEmpenho
                    End If
                    .MoveNext
                Wend
            End With
        End If
    Case 1
        
        strSQL = gstrQueryChequeRestoAPagar(strPKidOP, numeroOP, ExercicioOP)
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF
                    If gblnEncontroItemNoListView(lvw_Resto, !Pkid, lvwTag) Then 'encontra o banco na lista
                        ExcluiItemLista lvw_Resto, txtTotalResto
                    End If
                    .MoveNext
                Wend
            End With
        End If
    Case 2
    
        strSQL = gstrQueryChequeExtra(strPKidOP, numeroOP, ExercicioOP)
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF
                    If gblnEncontroItemNoListView(lvw_Despesa, !Pkid, lvwTag) Then  'encontra o banco na lista
                        ExcluiItemLista lvw_Despesa, txtTotalDespesa
                    End If
                    .MoveNext
                Wend
            End With
        End If
    End Select
End Sub
    


Private Sub IncluiOpsRelacionadas(ByVal strPKidOP As String)
    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset
   
    Select Case tab_3DPastaEmpenho.Tab
    Case 0
            
        strSQL = gstrQueryChequeEmpenho(strPKidOP)
                    
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF
                    If lvw_Empenho.ListItems.Count <> 0 Then
                        If lvw_Empenho.ListItems(1).SubItems(9) <> !bytAdiantamento Then
                            ExibeMensagem "Não é possivel incluir Empenhos de " & _
                            IIf(!bytAdiantamento = 1, "adiantamento", "outros tipos") & _
                            " junto com de " & IIf(!bytAdiantamento = 1, "outros tipos", "adiantamento")
                            Exit Sub
                        End If
                    End If
    
                    Set mobjLista = lvw_Empenho.ListItems.Add(, , Trim$(!INTEMPENHO) & Trim$(!INTNUMERO))
                    mobjLista.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                    mobjLista.SubItems(2) = !INTEMPENHO
                    mobjLista.SubItems(3) = !INTNUMERO
                    mobjLista.SubItems(4) = gstrDataFormatada(!dtmPrevisao)
                    mobjLista.SubItems(5) = gstrDataFormatada(!dtmLiquidacao)
                    mobjLista.SubItems(6) = gstrConvVrDoSql(!dblValor)
                    mobjLista.SubItems(7) = gstrConvVrDoSql(!dblDesconto)
                    mobjLista.SubItems(8) = gstrConvVrDoSql(!dblLiquido)
                    mobjLista.SubItems(9) = IIf(IsNull(!bytAdiantamento), "", !bytAdiantamento)
                    mobjLista.SubItems(10) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                    mobjLista.Tag = !Pkid
    
                    .MoveNext
                Wend
    
            End With
        End If
    Case 1
        
        strSQL = gstrQueryChequeRestoAPagar(strPKidOP)
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF

                    Set mobjLista = lvw_Resto.ListItems.Add(, , gstrItemData(dcbResto) & _
                                                                gstrItemData(dcbParcelaResto))
                    mobjLista.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                    mobjLista.SubItems(2) = !intExercicio
                    mobjLista.SubItems(3) = !intResto
                    mobjLista.SubItems(4) = !INTNUMERO
                    mobjLista.SubItems(5) = gstrDataFormatada(!dtmPrevisao)
                    mobjLista.SubItems(6) = gstrConvVrDoSql(!dblValor)
                    mobjLista.SubItems(7) = gstrConvVrDoSql(!dblDesconto)
                    mobjLista.SubItems(8) = gstrConvVrDoSql(!dblLiquido)
                    mobjLista.SubItems(9) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                    mobjLista.Tag = !Pkid
                    .MoveNext
                Wend
            End With
        End If
    Case 2
    
        strSQL = gstrQueryChequeExtra(strPKidOP)
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF
                        Set mobjLista = lvw_Despesa.ListItems.Add(, , !INTNUMERO)
                        mobjLista.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                        mobjLista.SubItems(2) = !INTNUMERO
                        mobjLista.SubItems(3) = gstrDataFormatada(!DTMDATA)
                        mobjLista.SubItems(4) = gstrConvVrDoSql(!dblValor)
                        mobjLista.SubItems(5) = gstrENulo(!STRNOME)
                        mobjLista.SubItems(6) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                        mobjLista.Tag = !Pkid

                    .MoveNext
                Wend
            End With
        End If
    End Select
End Sub

Private Sub SomaTotalAPagar()
    Dim dblValor As Double
    dblValor = Val(gstrConvVrParaSql(txtTotalEmpenho)) + _
               Val(gstrConvVrParaSql(txtTotalResto)) + _
               Val(gstrConvVrParaSql(txtTotalDespesa))
               
    dblValor = dblValor - (Val(gstrConvVrParaSql(txtTotalReceitaExtra)) + Val(gstrConvVrParaSql(txtTotalRecOrdem)))
               
    'txtTotalAPagar = IIf(chkEstorno.Value = 1, gstrConvVrDoSql(DBLVALOR * -1), gstrConvVrDoSql(DBLVALOR))
    txtTotalAPagar = gstrConvVrDoSql(dblValor)
End Sub

Private Sub Totaliza(lvw_Lista As ListView, txtTotal As TextBox)
    Dim intInd      As Integer
    Dim dblTotal    As Double
    With lvw_Lista
        For intInd = 1 To .ListItems.Count
            Select Case UCase(lvw_Lista.Name)
            Case "LVW_CONTA"
                dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(2)))
            Case "LVW_DESPESA"
                dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(4)))
            Case Else
                dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(6)))
            End Select
        Next
    End With
    txtTotal = gstrConvVrDoSql(IIf(chkEstorno.Value = 1, dblTotal * -1, dblTotal))
End Sub

Private Sub TotalizaRecOrc()
Dim intInd      As Integer
Dim dblTotal    As Double
    With lvw_Empenho
        For intInd = 1 To .ListItems.Count
            dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(11)))
        Next
    End With
    txtTotalRecOrdem = gstrConvVrDoSql(dblTotal)
End Sub

Private Sub ProcuraParcelaEmpenho()
    If gblnEncontroItemNoListView(lvw_Empenho, Trim$(dcbEmpenho) & _
                                  Trim$(dcbParcela), lvwText) Then
        mblnAlterandoEmpenho = True
    Else
        mblnAlterandoEmpenho = False
    End If
End Sub

Private Sub ProcuraParcelaOrdemEmpenho()
    Dim i As Integer
    For i = 1 To lvw_Empenho.ListItems.Count
        If lvw_Empenho.ListItems(i).SubItems(1) = dcbOrdemPagamentoEmpenho Then
            mblnAlterandoEmpenho = True
            Exit For
        Else
            mblnAlterandoEmpenho = False
        End If
    Next
End Sub

Private Sub ProcuraParcelaOrdemAnulacao()
    
    Dim i As Integer
    mblnAlterandoAnulacao = False
    For i = 1 To lvw_AnulacaoReceita.ListItems.Count
        If lvw_AnulacaoReceita.ListItems(i).SubItems(1) = dcbOrdemPagamentoAnulacao Then
            mblnAlterandoAnulacao = True
            Exit For
        Else
            mblnAlterandoAnulacao = False
        End If
    Next
End Sub

Private Sub ProcuraParcelaOrdemDespesa()
    Dim i As Integer
    For i = 1 To lvw_Despesa.ListItems.Count
        If lvw_Despesa.ListItems(i).SubItems(1) = dcbOrdemPagamentoDespesa Then
            mblnAlterandoDespesa = True
            Exit For
        Else
            mblnAlterandoDespesa = False
        End If
    Next
End Sub

Private Sub ProcuraParcelaResto()
    If gblnEncontroItemNoListView(lvw_Resto, gstrItemData(dcbResto) & _
                                  gstrItemData(dcbParcelaResto), lvwText) Then
        mblnAlterandoResto = True
    Else
        mblnAlterandoResto = False
    End If
End Sub

Private Sub ProcuraDespesa()
    If gblnEncontroItemNoListView(lvw_Despesa, dcbDespesa.BoundText, lvwTag) Then
        mblnAlterandoDespesa = True
    Else
        mblnAlterandoDespesa = False
    End If
End Sub

Private Sub VerificaLista()
    Select Case tab_3DPastaEmpenho.Tab
    Case 0
        IncluiAlteraListaEmpenho
    Case 1
        IncluiAlteraListaResto
    Case 2
        IncluiAlteraListaDespesa
    Case 3
        IncluiAlteraConta
    Case 5
        IncluiAlteraListaAnulacao
    End Select
End Sub

Private Sub ExcluiItemLista(lvw_Lista As ListView, txtTotal As TextBox)
    Dim mstrOrdem As String
    Dim i As Integer
    
    With lvw_Lista
        If .ListItems.Count = 0 Then Exit Sub
        mstrOrdem = .SelectedItem.SubItems(1)
        If mstrOrdem = "--" Then
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
            If Len(Trim(dcbOrdemPagamentoDespesa)) > 0 Then
               MontaTotalRecOrdem False
            ElseIf Len(Trim(dcbOrdemPagamentoEmpenho)) > 0 Then
               MontaTotalRecOrdem False
            ElseIf Len(Trim(dcbOrdemPagamentoResto)) > 0 Then
               MontaTotalRecOrdem False
            End If
            Totaliza lvw_Lista, txtTotal
            
        Else
reinicia_exclusao:
            For i = 1 To .ListItems.Count
                If .ListItems(i).SubItems(1) = mstrOrdem Then
                    .ListItems.Remove .ListItems(i).Index
                    GoTo reinicia_exclusao
                End If
            Next
        End If
    End With
End Sub

Private Sub VerificaListaExcluir()
    Dim mstrNumeroOrdem As String
    Dim mstrExercicioOrdem As String

    Select Case tab_3DPastaEmpenho.Tab
    Case 0
        If Not lvw_Empenho.SelectedItem Is Nothing Then
            mstrNumeroOrdem = lvw_Empenho.SelectedItem.SubItems(1)
            mstrExercicioOrdem = lvw_Empenho.SelectedItem.SubItems(10)
        End If
        MontaTotalRecOrdem False
        ExcluiItemLista lvw_Empenho, txtTotalEmpenho
        If Not lvw_Empenho.SelectedItem Is Nothing Then ExcluiOpsRelacionadas "", mstrNumeroOrdem, mstrExercicioOrdem
        LimpaDadosEmpenho
        Totaliza lvw_Empenho, txtTotalEmpenho
        preencheLiquidacaoExtra lvw_Empenho
    Case 1
        If Not lvw_Resto.SelectedItem Is Nothing Then
            mstrNumeroOrdem = lvw_Resto.SelectedItem.SubItems(1)
            mstrExercicioOrdem = lvw_Resto.SelectedItem.SubItems(9)
        End If
        ExcluiItemLista lvw_Resto, txtTotalResto
        If Not lvw_Resto.SelectedItem Is Nothing Then ExcluiOpsRelacionadas "", mstrNumeroOrdem, mstrExercicioOrdem
        LimpaDadosResto
        MontaTotalRecOrdem True
        Totaliza lvw_Resto, txtTotalResto
        preencheLiquidacaoExtra lvw_Resto
    Case 2
        If Not lvw_Despesa.SelectedItem Is Nothing Then
            mstrNumeroOrdem = lvw_Despesa.SelectedItem.SubItems(1)
            mstrExercicioOrdem = lvw_Despesa.SelectedItem.SubItems(6)
        End If
        ExcluiItemLista lvw_Despesa, txtTotalDespesa
        If Not lvw_Despesa.SelectedItem Is Nothing Then ExcluiOpsRelacionadas "", mstrNumeroOrdem, mstrExercicioOrdem
        Totaliza lvw_Despesa, txtTotalDespesa
        LimpaDadosDespesa
        txtTotalReceitaExtra = "0,00"
    Case 3
        ExcluiItemLista lvw_Conta, txtTotal
        TotalLancado
        LimpaDadosConta
    Case 5
        ExcluiListaAnulacao
        LimpaDadosConta
        SomaTotalRecOrdemAnulacao
    End Select
End Sub

Private Sub ExcluiListaAnulacao()
    Dim strOrdem As String
    Dim i As Integer
    If Not lvw_AnulacaoReceita.SelectedItem Is Nothing Then
        strOrdem = lvw_AnulacaoReceita.SelectedItem.SubItems(1)
    End If
    
ReComeca:
    For i = 1 To lvw_AnulacaoReceita.ListItems.Count
        If strOrdem = lvw_AnulacaoReceita.ListItems(i).SubItems(1) Then
            lvw_AnulacaoReceita.ListItems.Remove (i)
            GoTo ReComeca:
        End If
    Next

End Sub

Private Function blnDadosEmpenhoOK() As Boolean

    Dim strSQL              As String
    Dim adoResultado        As New ADODB.Recordset
    Dim strTipoEmpenhoAtual As String
    Dim strPKIds            As String
    Dim strOrigemOP         As Boolean

    If dcbEmpenho.MatchedWithList = False And dcbOrdemPagamentoEmpenho.MatchedWithList = False Then
        ExibeMensagem "O número do empenho ou da Ordem de Pagamento tem que ser informado corretamente."
        'If dcbEmpenho.Enabled Then dcbEmpenho.SetFocus
        If dcbOrdemPagamentoEmpenho.Enabled Then dcbOrdemPagamentoEmpenho.SetFocus
        Exit Function
    End If
        
    If dcbEmpenho.MatchedWithList = True And dcbParcela.MatchedWithList = False Then
        ExibeMensagem "O número da parcela tem que ser informado corretamente."
        If dcbParcela.Enabled Then dcbParcela.SetFocus
        Exit Function
    End If
    
    
'    If lvw_Empenho.ListItems.Count > 0 Then
'        strTipoEmpenhoAtual = lvw_Empenho.ListItems(1).SubItems(8)
'    Else
'        blnDadosEmpenhoOK = True
'        Exit Function
'    End If
'
'    If dcbOrdemPagamentoEmpenho.BoundText <> "" Then
'        strPKIds = gstrEmpenhobySubEmpenho(gstrOrdemPagamentoItens(dcbOrdemPagamentoEmpenho, gstrOrdemPagamentoEmpenho, True))
'        strOrigemOP = True
'    Else
'        strPKIds = CStr(gstrItemData(dcbEmpenho))
'        strOrigemOP = False
'    End If
'
'    strSQL = ""
'    strSQL = strSQL & " SELECT E.intnumero , TE.bytAdiantamento FROM "
'    strSQL = strSQL & gstrEmpenho & " E,"
'    strSQL = strSQL & gstrTipoEmpenho & " TE"
'    strSQL = strSQL & " WHERE E.Pkid IN (" & strPKIds & ")"
'    strSQL = strSQL & " AND TE.PKID = E.IntTipo "
'
'   Set gobjBanco = New clsBanco
'   If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
'      With adoResultado
'         While Not .EOF
'            'If !bytAdiantamento <> strTipoEmpenhoAtual Then
'            If !bytAdiantamento = 1 Then
'                If strOrigemOP Then
'                    ExibeMensagem "Não é possível inserir Ordem de Pagamanto com empenhos de tipos diferentes."
'                    If dcbOrdemPagamentoEmpenho.Enabled Then dcbOrdemPagamentoEmpenho.SetFocus
'                Else
'                    ExibeMensagem "Não é possível inserir empenhos de tipos diferentes."
'                    If dcbEmpenho.Enabled Then dcbEmpenho.SetFocus
'                End If
'                Exit Function
'            End If
'            .MoveNext
'         Wend
'      End With
'   End If
    
    blnDadosEmpenhoOK = True
End Function

Private Function blnDadosDespesaOK() As Boolean
    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset



    If dcbDespesa.MatchedWithList = False And dcbOrdemPagamentoDespesa.MatchedWithList = False Then
        ExibeMensagem "O número da Despesa ou da Ordem de Pagamento tem que ser informado corretamente."
        'If dcbDespesa.Enabled Then dcbDespesa.SetFocus
        If dcbOrdemPagamentoDespesa.Enabled Then dcbOrdemPagamentoDespesa.SetFocus
        Exit Function
    End If
    
    
    If lvw_Despesa.ListItems.Count > 0 Then
    
        strSQL = ""
        strSQL = strSQL & "Select intContaContabil intContaContabilCombo ,"
        strSQL = strSQL & "(Select intContaContabil FROM " & gstrDespesaExtraOrcamentaria & " WHERE PKID = " & lvw_Despesa.ListItems(1).Tag & ") intContaContabilGrid "
        strSQL = strSQL & " FROM " & gstrDespesaExtraOrcamentaria
        strSQL = strSQL & " WHERE "


        If dcbDespesa.Text <> "" Then
            strSQL = strSQL & " PKID = " & gstrItemData(dcbDespesa)
        ElseIf dcbOrdemPagamentoDespesa.Text <> "" Then
            strSQL = strSQL & " Pkid IN (" & gstrOrdemPagamentoItens(dcbOrdemPagamentoDespesa, gstrOrdemPagamentoDespesaExtra, True, lvw_Despesa) & ")"
        End If
        

        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
           With adoResultado
              While Not .EOF
                 If !intContaContabilGrid <> !intContaContabilCombo Then
                     ExibeMensagem "Não é possivel inserir itens de Despesa Extra-Orçamentária com contas diferentes."
                     'If dcbDespesa.Enabled Then dcbDespesa.SetFocus
                     If dcbOrdemPagamentoDespesa.Enabled Then dcbOrdemPagamentoDespesa.SetFocus
                     Exit Function
                 End If
                 .MoveNext
              Wend
           End With
        End If
    End If

       
    blnDadosDespesaOK = True

End Function

Private Function blnDadosRestoOk() As Boolean
    If dcbResto.MatchedWithList = False And dcbOrdemPagamentoResto.MatchedWithList = False Then
        ExibeMensagem "O número do Resto ou da Ordem de Pagamento tem que ser informado corretamente."
        'If dcbResto.Enabled Then dcbResto.SetFocus
        If dcbOrdemPagamentoResto.Enabled Then dcbOrdemPagamentoResto.SetFocus
        Exit Function
    End If
        
    If dcbResto.MatchedWithList = True And dcbParcelaResto.MatchedWithList = False Then
        ExibeMensagem "O número da parcela tem que ser informado corretamente."
        If dcbParcelaResto.Enabled Then dcbParcelaResto.SetFocus
        Exit Function
    End If
    blnDadosRestoOk = True
End Function

Private Function blnDadosAnulacaoOk() As Boolean
    If dcbOrdemPagamentoAnulacao.MatchedWithList = False Then
        ExibeMensagem "O número da Ordem de Pagamento tem que ser informado corretamente."
         If dcbOrdemPagamentoAnulacao.Enabled Then dcbOrdemPagamentoAnulacao.SetFocus
        Exit Function
    End If
    
    blnDadosAnulacaoOk = True
End Function

Private Sub IncluiAlteraListaEmpenho()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
Dim strOrdemPagamentoEmpenho
Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
    If blnDadosEmpenhoOK() Then
        
        
            strSQL = ""
            strSQL = "SELECT SE.Pkid Pkid, "
            strSQL = strSQL & "(SELECT OP.intNumero FROM " & gstrOrdemPagamentoEmpenho & " OPR, " & gstrOrdemPagamento & " OP"
            strSQL = strSQL & " WHERE OP.PKID = OPR.intOrdemPagamento AND OPR.intParcela = SE.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intOrdem,"
            strSQL = strSQL & "(SELECT OP.intExercicio FROM " & gstrOrdemPagamentoEmpenho & " OPR, " & gstrOrdemPagamento & " OP"
            strSQL = strSQL & " WHERE OP.PKID = OPR.intOrdemPagamento AND OPR.intParcela = SE.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intExercicioOP,"

            strSQL = strSQL & "E.IntNumero intEmpenho, "
            strSQL = strSQL & "SE.IntNumero intNumero, "
            strSQL = strSQL & "SE.dblvalor dblValor, "
            strSQL = strSQL & "SE.dtmData dtmPrevisao, "
            strSQL = strSQL & "SE.dtmLiquidacao dtmLiquidacao, "
            strSQL = strSQL & gstrISNULL("SUM(SL.dblvalor)", "0") & " dblDesconto, "
            strSQL = strSQL & gstrISNULL("SE.dblvalor", "0") & " - "
            strSQL = strSQL & gstrISNULL("SUM(Sl.dblValor)", "0") & " dblLiquido, "
            strSQL = strSQL & "TE.bytAdiantamento "
            strSQL = strSQL & "FROM " & gstrEmpenho & " E ," & gstrTipoEmpenho & " TE,"
            strSQL = strSQL & gstrSubempenho & " SE, " & gstrSubempenhoLiquidado & " SL"
            strSQL = strSQL & " WHERE "
            
            If dcbEmpenho.Text <> "" Then
                strSQL = strSQL & "SE.Pkid =" & gstrItemData(dcbParcela) & " AND "
            ElseIf dcbOrdemPagamentoEmpenho.Text <> "" Then
                If mblnAlterandoEmpenho Then
                    dcbOrdemPagamentoEmpenho.Text = ""
                    Exit Sub
                End If
                strSQL = strSQL & "SE.Pkid IN (" & gstrOrdemPagamentoItens(dcbOrdemPagamentoEmpenho, gstrOrdemPagamentoEmpenho, True, lvw_Empenho) & ") AND "
            End If
            
            strSQL = strSQL & " TE.PKID = E.Inttipo "
            strSQL = strSQL & " AND E.Pkid = SE.intEmpenho AND "
            strSQL = strSQL & "SE.Pkid " & strOUTJSQLServer & "= SL.intParcela" & strOUTJOracle
            
            
            strSQL = strSQL & " GROUP BY SE.Pkid, E.IntNumero, SE.IntNumero, SE.dblvalor,SE.dtmData,SE.dtmLiquidacao,TE.bytAdiantamento"
        
'            If dcbEmpenho.Text <> "" Then
'                strSQL = gstrStoredProcedure("sp_EmpenhoParaPagar", CStr(Val(dcbParcela.BoundText)), True)
'            ElseIf dcbOrdemPagamentoEmpenho.Text <> "" Then
'                If mblnAlterandoEmpenho Then
'                    dcbOrdemPagamentoEmpenho.Text = ""
'                    Exit Sub
'                End If
'                strSQL = gstrStoredProcedure("sp_EmpenhoParaPagar", gstrOrdemPagamentoItens(dcbOrdemPagamentoEmpenho, gstrOrdemPagamentoEmpenho), True)
'            End If
        
                
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF
                    If lvw_Empenho.ListItems.Count <> 0 Then
                        If lvw_Empenho.ListItems(1).SubItems(9) <> !bytAdiantamento Then
                            ExibeMensagem "Não é possivel incluir Empenhos de " & _
                            IIf(!bytAdiantamento = 1, "adiantamento", "outros tipos") & _
                            " junto com de " & IIf(!bytAdiantamento = 1, "outros tipos", "adiantamento")
                            Exit Sub
                        End If
                    End If
                    If mblnAlterandoEmpenho Then
                        lvw_Empenho.ListItems(lvw_Empenho.SelectedItem.Index).Text = Trim$(!INTEMPENHO) & Trim$(!INTNUMERO)
                        lvw_Empenho.SelectedItem.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                        lvw_Empenho.SelectedItem.SubItems(2) = !INTEMPENHO
                        lvw_Empenho.SelectedItem.SubItems(3) = !INTNUMERO
                        lvw_Empenho.SelectedItem.SubItems(4) = gstrDataFormatada(!dtmPrevisao)
                        lvw_Empenho.SelectedItem.SubItems(5) = gstrDataFormatada(!dtmLiquidacao)
                        lvw_Empenho.SelectedItem.SubItems(6) = gstrConvVrDoSql(!dblValor)
                        lvw_Empenho.SelectedItem.SubItems(7) = gstrConvVrDoSql(!dblDesconto)
                        lvw_Empenho.SelectedItem.SubItems(8) = gstrConvVrDoSql(!dblLiquido)
                        lvw_Empenho.SelectedItem.SubItems(9) = !bytAdiantamento
                        lvw_Empenho.SelectedItem.SubItems(10) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                    Else
                        Set mobjLista = lvw_Empenho.ListItems.Add(, , Trim$(!INTEMPENHO) & Trim$(!INTNUMERO))
                        mobjLista.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                        mobjLista.SubItems(2) = !INTEMPENHO
                        mobjLista.SubItems(3) = !INTNUMERO
                        mobjLista.SubItems(4) = gstrDataFormatada(!dtmPrevisao)
                        mobjLista.SubItems(5) = gstrDataFormatada(!dtmLiquidacao)
                        mobjLista.SubItems(6) = gstrConvVrDoSql(!dblValor)
                        mobjLista.SubItems(7) = gstrConvVrDoSql(!dblDesconto)
                        mobjLista.SubItems(8) = gstrConvVrDoSql(!dblLiquido)
                        mobjLista.SubItems(9) = IIf(IsNull(!bytAdiantamento), "", !bytAdiantamento)
                        mobjLista.SubItems(10) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                        mobjLista.Tag = !Pkid
                    End If
                    .MoveNext
                Wend
                IncluiContaCheque dcbOrdemPagamentoEmpenho.BoundText
                MontaTotalRecOrdem True
                Totaliza lvw_Empenho, txtTotalEmpenho
                preencheLiquidacaoExtra lvw_Empenho
                LimpaDadosEmpenho
            End With
        End If
    End If
End Sub

Private Sub IncluiAlteraListaResto()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    If blnDadosRestoOk() Then

        'strSQL = gstrStoredProcedure("sp_RestoParaPagar", CStr(Val(dcbParcelaResto.BoundText)), True)

        strSQL = ""
        strSQL = "SELECT SE.Pkid Pkid, "
        strSQL = strSQL & "(SELECT OP.intNumero FROM " & gstrOrdemPagamentoResto & " OPR, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OPR.intOrdemPagamento AND OPR.intParcela = SE.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intOrdem,"
        strSQL = strSQL & "(SELECT OP.intExercicio FROM " & gstrOrdemPagamentoResto & " OPR, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OPR.intOrdemPagamento AND OPR.intParcela = SE.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intExercicioOP,"
        strSQL = strSQL & "E.intExercicioRP, "
        strSQL = strSQL & gstrDATEPART(strYEAR, "E.dtmData") & " intExercicio, "
        strSQL = strSQL & "E.IntNumero intResto, "
        strSQL = strSQL & "SE.IntNumero intNumero, "
        strSQL = strSQL & "SE.dblvalor dblValor, "
        strSQL = strSQL & "SE.dtmData dtmPrevisao, "
        strSQL = strSQL & gstrISNULL("SUM(SL.dblvalor)", "0") & " dblDesconto, "
        strSQL = strSQL & gstrISNULL("SE.dblvalor", "0") & " - "
        strSQL = strSQL & gstrISNULL("SUM(Sl.dblValor)", "0") & " dblLiquido "
        strSQL = strSQL & "FROM " & gstrEmpenho & " E ,"
        strSQL = strSQL & gstrSubempenho & " SE, " & gstrSubempenhoLiquidado & " SL"
        strSQL = strSQL & " WHERE "
        
        If dcbResto.Text <> "" Then
            strSQL = strSQL & "SE.Pkid =" & gstrItemData(dcbParcelaResto) & " AND "
        ElseIf dcbOrdemPagamentoResto.Text <> "" Then
            If mblnAlterandoResto Then
                dcbOrdemPagamentoResto.Text = ""
                Exit Sub
            End If
            strSQL = strSQL & "SE.Pkid IN (" & gstrOrdemPagamentoItens(dcbOrdemPagamentoResto, gstrOrdemPagamentoResto, True, lvw_Resto) & ") AND "
        End If
        
        strSQL = strSQL & "E.Pkid = SE.intEmpenho AND "
        strSQL = strSQL & "SE.Pkid " & strOUTJSQLServer & "= SL.intParcela" & strOUTJOracle & " And "
        
        strSQL = strSQL & "E.intExercicioRP =" & CStr(gintExercicio)
        
        strSQL = strSQL & " GROUP BY SE.Pkid, E.intExercicioRP, E.IntNumero, SE.IntNumero,SE.dblvalor ,SE.dtmData ," & gstrDATEPART(strYEAR, "E.dtmData")
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF
                    If mblnAlterandoResto Then
                        lvw_Resto.ListItems(lvw_Resto.SelectedItem.Index).Text = _
                                            gstrItemData(dcbResto) & gstrItemData(dcbParcelaResto)
                        lvw_Resto.SelectedItem.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                        lvw_Resto.SelectedItem.SubItems(2) = !intExercicio
                        lvw_Resto.SelectedItem.SubItems(3) = !intResto
                        lvw_Resto.SelectedItem.SubItems(4) = !INTNUMERO
                        lvw_Resto.SelectedItem.SubItems(5) = gstrDataFormatada(!dtmPrevisao)
                        lvw_Resto.SelectedItem.SubItems(6) = gstrConvVrDoSql(!dblValor)
                        lvw_Resto.SelectedItem.SubItems(7) = gstrConvVrDoSql(!dblDesconto)
                        lvw_Resto.SelectedItem.SubItems(8) = gstrConvVrDoSql(!dblLiquido)
                        lvw_Resto.SelectedItem.SubItems(9) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                        lvw_Resto.SelectedItem.Tag = !Pkid
                    Else
                        Set mobjLista = lvw_Resto.ListItems.Add(, , gstrItemData(dcbResto) & _
                                                                    gstrItemData(dcbParcelaResto))
                        mobjLista.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                        mobjLista.SubItems(2) = !intExercicio
                        mobjLista.SubItems(3) = !intResto
                        mobjLista.SubItems(4) = !INTNUMERO
                        mobjLista.SubItems(5) = gstrDataFormatada(!dtmPrevisao)
                        mobjLista.SubItems(6) = gstrConvVrDoSql(!dblValor)
                        mobjLista.SubItems(7) = gstrConvVrDoSql(!dblDesconto)
                        mobjLista.SubItems(8) = gstrConvVrDoSql(!dblLiquido)
                        mobjLista.SubItems(9) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                        mobjLista.Tag = !Pkid
                    End If
                        .MoveNext
                Wend
                IncluiContaCheque dcbOrdemPagamentoResto.BoundText
                MontaTotalRecOrdem True
                Totaliza lvw_Resto, txtTotalResto
                preencheLiquidacaoExtra lvw_Resto
                LimpaDadosResto
            End With
        End If
    End If
End Sub


Private Sub IncluiAlteraListaAnulacao()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    If blnDadosAnulacaoOk() Then

     If mblnAlterandoAnulacao Then Exit Sub

        strSQL = ""
        strSQL = "SELECT OPA.PKID, OP.INTNUMERO intOrdem, OPA.Strdescricao, "
        strSQL = strSQL & " OPA.Dblvalor , OPA.Dtmdtatualizacao, OP.intExercicio intExercicioOP "
        strSQL = strSQL & "FROM " & gstrOrdemPagamento & " OP ,"
        strSQL = strSQL & gstrOrdemPagamentoAnulacaoReceita & " OPA "
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " OP.PKID = " & gstrItemData(dcbOrdemPagamentoAnulacao)
        strSQL = strSQL & " AND OP.PKID = OPA.intOrdemPagamento "
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF
                    If mblnAlterandoAnulacao Then
                        lvw_AnulacaoReceita.ListItems(lvw_AnulacaoReceita.SelectedItem.Index).Text = _
                                            gstrItemData(dcbOrdemPagamentoAnulacao)
                        lvw_AnulacaoReceita.SelectedItem.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                        lvw_AnulacaoReceita.SelectedItem.SubItems(2) = !strDescricao
                        lvw_AnulacaoReceita.SelectedItem.SubItems(3) = gstrDataFormatada(!dtmDtAtualizacao)
                        lvw_AnulacaoReceita.SelectedItem.SubItems(4) = gstrConvVrDoSql(!dblValor)
                        lvw_AnulacaoReceita.SelectedItem.SubItems(5) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                        lvw_AnulacaoReceita.SelectedItem.Tag = !Pkid
                    Else
                        Set mobjLista = lvw_AnulacaoReceita.ListItems.Add(, , gstrItemData(dcbOrdemPagamentoAnulacao))
                        mobjLista.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                        mobjLista.SubItems(2) = !strDescricao
                        mobjLista.SubItems(3) = gstrDataFormatada(!dtmDtAtualizacao)
                        mobjLista.SubItems(4) = gstrConvVrDoSql(!dblValor)
                        mobjLista.SubItems(5) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                        mobjLista.Tag = !Pkid
                    End If
                        .MoveNext
                Wend
                
                SomaTotalRecOrdemAnulacao
                LimpaDadosAnulacao
            End With
        End If
    End If
End Sub

Private Sub SomaTotalRecOrdemAnulacao()
    Dim i As Integer
    txtTotalRecOrdem = gstrConvVrDoSql("0")
    For i = 1 To lvw_AnulacaoReceita.ListItems.Count
        txtTotalRecOrdem = gstrConvVrDoSql(CDbl(txtTotalRecOrdem) + CDbl(lvw_AnulacaoReceita.ListItems(i).ListSubItems(4)))
    Next
    
    txtTotalRecOrdem = gstrConvVrDoSql(IIf(chkEstorno.Value = 1, CDbl(txtTotalRecOrdem) * -1, CDbl(txtTotalRecOrdem)))
    
    txtTotalAPagar = txtTotalRecOrdem
End Sub

Private Sub IncluiAlteraListaDespesa()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    If blnDadosDespesaOK() Then
        strSQL = ""
        strSQL = strSQL & "SELECT DP.PKId,"
        strSQL = strSQL & "(SELECT OP.intNumero FROM " & gstrOrdemPagamentoDespesaExtra & " OD, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OD.intOrdemPagamento AND OD.intDespesaExtraOrcamentaria = DP.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intOrdem,"
        strSQL = strSQL & "(SELECT OP.intExercicio FROM " & gstrOrdemPagamentoDespesaExtra & " OD, " & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE OP.PKID = OD.intOrdemPagamento AND OD.intDespesaExtraOrcamentaria = DP.Pkid AND (OP.bytCancelado = 0 OR OP.bytCancelado is null)) intExercicioOP,"
        strSQL = strSQL & " DP.intNumero , DP.dtmData, "
        strSQL = strSQL & "DP.dblValor, CT.strNome FROM "
        strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DP, "
        strSQL = strSQL & gstrContribuinte & " CT "
        strSQL = strSQL & "WHERE DP.intContribuinte = CT.PKId AND "
        If dcbDespesa.Text <> "" Then
            strSQL = strSQL & "DP.Pkid =" & gstrItemData(dcbDespesa)
        ElseIf dcbOrdemPagamentoDespesa.Text <> "" Then
            If mblnAlterandoDespesa Then
                dcbOrdemPagamentoDespesa.Text = ""
                Exit Sub
            End If
            strSQL = strSQL & "DP.Pkid IN (" & gstrOrdemPagamentoItens(dcbOrdemPagamentoDespesa, gstrOrdemPagamentoDespesaExtra, True, lvw_Despesa) & ")"
        End If
        'strSQL = strSQL & " DP.PKId = " & Val(dcbDespesa.BoundText) & " "
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                While Not .EOF
                    If mblnAlterandoDespesa Then
                        lvw_Despesa.ListItems(lvw_Despesa.SelectedItem.Index).Text = Trim$(!INTNUMERO)
                        lvw_Despesa.SelectedItem.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                        lvw_Despesa.SelectedItem.SubItems(2) = !INTNUMERO
                        lvw_Despesa.SelectedItem.SubItems(3) = gstrDataFormatada(!DTMDATA)
                        lvw_Despesa.SelectedItem.SubItems(4) = gstrConvVrDoSql(!dblValor)
                        lvw_Despesa.SelectedItem.SubItems(5) = gstrENulo(!STRNOME)
                        lvw_Despesa.SelectedItem.SubItems(6) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                        lvw_Despesa.SelectedItem.Tag = !Pkid
                    Else
                        Set mobjLista = lvw_Despesa.ListItems.Add(, , !INTNUMERO)
                        mobjLista.SubItems(1) = IIf(IsNull(!intOrdem), "--", !intOrdem)
                        mobjLista.SubItems(2) = !INTNUMERO
                        mobjLista.SubItems(3) = gstrDataFormatada(!DTMDATA)
                        mobjLista.SubItems(4) = gstrConvVrDoSql(!dblValor)
                        mobjLista.SubItems(5) = gstrENulo(!STRNOME)
                        mobjLista.SubItems(6) = IIf(IsNull(!intExercicioOP), "--", !intExercicioOP)
                        mobjLista.Tag = !Pkid
                    End If
                    .MoveNext
                Wend
                IncluiContaCheque dcbOrdemPagamentoDespesa.BoundText
                MontaTotalRecOrdem True
                Totaliza lvw_Despesa, txtTotalDespesa
                LimpaDadosDespesa
                txtTotalReceitaExtra = "0,00"
            End With
        End If
    End If
End Sub

Private Function strSqlParcelaEmpenho() As String
    Dim strSQL  As String
    Dim mstrSituacao As String
    mstrSituacao = IIf(chkEstorno.Value = 0, "2", "3")
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, intNumero "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrSubempenho & " "
    strSQL = strSQL & "WHERE bytSituacao = " & mstrSituacao
    strSQL = strSQL & " AND intEmpenho = " & gstrItemData(dcbEmpenho)
    'strSql = strSql & " AND NOT PKID IN (SELECT IntParcela FROM " & gstrOrdemPagamentoEmpenho & " ) "
    strSQL = strSQL & " AND NOT PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela) "
    
    strSQL = strSQL & " ORDER BY intNumero"
    strSqlParcelaEmpenho = strSQL
End Function

Private Function strQuerySubempenho() As String

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL      As String
    Dim mstrSituacao As String
    mstrSituacao = IIf(chkEstorno.Value = 0, "2", "3")
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, intNumero "
    strSQL = strSQL & "FROM " & gstrSubempenho & " "
    strSQL = strSQL & "WHERE intEmpenho = " & CStr(gstrItemData(dcbResto, True))
    strSQL = strSQL & " AND bytSituacao = " & mstrSituacao
    'strSql = strSql & " AND NOT PKID IN (SELECT IntParcela FROM " & gstrOrdemPagamentoResto & " ) "
    strSQL = strSQL & " AND NOT PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoResto & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela)     "
    strSQL = strSQL & " ORDER BY intNumero"
    
    strQuerySubempenho = strSQL
End Function

Private Function strSqlParcelaResto() As String
    Dim strSQL  As String
    Dim mstrSituacao As String
    mstrSituacao = IIf(chkEstorno.Value = 0, "2", "3")
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, intNumero "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrParcelaRestoPagar & " "
    strSQL = strSQL & "WHERE bytSituacao = " & mstrSituacao 'Paga
    strSQL = strSQL & "AND intResto = " & Val(dcbResto.BoundText) & " "
'    strSql = strSql & "AND NOT PKID IN (SELECT IntParcela FROM " & gstrOrdemPagamentoResto & " ) "
    strSQL = strSQL & " AND NOT PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoResto & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela) "
    
    strSQL = strSQL & "ORDER BY intNumero"
    strSqlParcelaResto = strSQL
End Function

Private Function strQueryLancamento() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PC.PKId, CB.intNumeroConta, PC.strDescricao, "
    strSQL = strSQL & "ABS(LC.dblValor) dblValor, LC.strDocumento FROM "
    strSQL = strSQL & gstrLancamentoContabil & " LC, "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrContaBancaria & " CB "
    strSQL = strSQL & "WHERE LC.intConta = PC.PKId "
    strSQL = strSQL & "AND CB.PKId = PC.intContaBancaria "
    strSQL = strSQL & "AND LC.intProcesso = " & Val(txtPKID) & " "
    strSQL = strSQL & "AND PC.blnAnalitica = 1 "
    strSQL = strSQL & "AND PC.bytDisponibilidadeDeCaixa = 1 "
    strSQL = strSQL & "ORDER BY PC.strContaContabil"
    
'    strSql = strSql & "SELECT PC.PKId, "
'    strSql = strSql & "CB.intNumeroConta, "
'    strSql = strSql & "PC.strDescricao "
'    strSql = strSql & "FROM " & gstrContaBancaria & " CB, "
'    strSql = strSql & gstrPlanoConta & " PC "
'    strSql = strSql & "WHERE CB.PKId = PC.intContaBancaria AND "
'    strSql = strSql & "PC.blnAnalitica = 1 AND "
'    strSql = strSql & "PC.bytDisponibilidadeDeCaixa = 1 "
'    strSql = strSql & "ORDER BY PC.strDescricao"
    
    strQueryLancamento = strSQL
End Function

Private Function strQueryEmpenho() As String
    Dim strSQL       As String
    Dim mstrSituacao As String
    
    mstrSituacao = IIf(chkEstorno.Value = 0, "2", "3")
    
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT EP.PKId, EP.intNumero "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP, "
    strSQL = strSQL & gstrTipoEmpenho & " TE, "
    strSQL = strSQL & gstrSubempenho & " SE, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
    strSQL = strSQL & "WHERE EP.PKId = SE.intEmpenho "
    strSQL = strSQL & "AND SE.bytSituacao = " & mstrSituacao
    strSQL = strSQL & " AND " & gstrISNULL("EP.intExercicioRP", "0") & " = 0 "
    
    strSQL = strSQL & "AND NOT SE.PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela) "
    
    strSQL = strSQL & "AND PT.PKID = EP.intProgramaTrabalho "
    strSQL = strSQL & "AND PT.intExercicio = " & CStr(gintExercicio)
    strSQL = strSQL & " AND EP.intTipo = TE.PKID "
    
    If chkEstorno.Value = 1 Then
        strSQL = strSQL & " AND NOT EP.PKID IN ( "
        strSQL = strSQL & " SELECT EMP.PKID FROM "
        strSQL = strSQL & gstrSubempenho & " SUBE,"
        strSQL = strSQL & gstrEmpenho & " EMP, "
        strSQL = strSQL & gstrTipoEmpenho & " TIE "
        strSQL = strSQL & " WHERE SUBE.INTEMPENHO = EMP.Pkid"
        strSQL = strSQL & " AND EMP.inttipo = TIE.PKID "
        strSQL = strSQL & " AND TIE.Bytadiantamento = 1 "
        strSQL = strSQL & " AND SUBE.intnumero = 0 "
        strSQL = strSQL & " AND SUBE.bytsituacao = 4)"
    
        'strSQL = strSQL & " AND (TE.bytAdiantamento IS NULL OR TE.bytAdiantamento=0)"
    End If
    
    strSQL = strSQL & " ORDER BY EP.intNumero"
    strQueryEmpenho = strSQL
End Function

Private Function strQueryOrdemPagamentoEmpenho(Optional intOPE As Variant) As String
    Dim strSQL       As String
    Dim mstrSituacao As String
    
    mstrSituacao = IIf(chkEstorno.Value = 0, "0", "1")
    
    strSQL = ""
    'Select que abriga todos
    strSQL = strSQL & "SELECT TMP.PKId, TMP.intNumero FROM ("
    strSQL = strSQL & "SELECT OP.PKId, OP.intNumero ,"
    
    'Select equivalente ao campo Tipo
    strSQL = strSQL & "(SELECT SUM(" & gstrCONVERT(CDT_INT, "TE.bytAdiantamento") & ") FROM "
    strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OE, "
    strSQL = strSQL & gstrSubempenho & " SE, "
    strSQL = strSQL & gstrEmpenho & " E, " & gstrTipoEmpenho & " TE "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " OE.intOrdemPagamento = OP.Pkid"
    strSQL = strSQL & " AND OE.intParcela = SE.PKID"
    strSQL = strSQL & " AND E.PKID = SE.intEmpenho"
    strSQL = strSQL & " AND TE.pkid = E.intTipo) tipo "
    
    strSQL = strSQL & "FROM  "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OP.intExercicio = " & gintExercicio
    
    
    strSQL = strSQL & " AND OP.bytTipo = 0 "
    
    If chkEstorno.Value = 1 Then
        strSQL = strSQL & " AND NOT OP.PKID IN ( "
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & " OP.Pkid "
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrSubempenho & " SUBE,"
        strSQL = strSQL & gstrSubempenho & " SUBE1,"
        strSQL = strSQL & gstrOrdemPagamento & " OP ,"
        strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OPE,"
        strSQL = strSQL & gstrEmpenho & " EMP,"
        strSQL = strSQL & gstrEmpenho & " EMP1,"
        strSQL = strSQL & gstrTipoEmpenho & " TIE"
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " SUBE.INTEMPENHO = EMP.Pkid"
        strSQL = strSQL & " AND EMP.inttipo = TIE.PKID"
        strSQL = strSQL & " AND TIE.Bytadiantamento = 1"
        strSQL = strSQL & " AND SUBE.intnumero = 0"
        strSQL = strSQL & " AND SUBE.bytsituacao = 4"
        strSQL = strSQL & " AND SUBE1.PKID = OPE.INTPARCELA"
        strSQL = strSQL & " AND OP.PKID = OPE.INTORDEMPAGAMENTO"
        strSQL = strSQL & " AND EMP1.PKID = SUBE1.INTEMPENHO"
        strSQL = strSQL & " AND EMP1.PKID = EMP.PKID"
        strSQL = strSQL & " GROUP BY OP.PKID)"
    End If
    
    If Not IsMissing(intOPE) Then
       strSQL = strSQL & " AND OP.INTNUMERO = " & intOPE
    End If
    
    strSQL = strSQL & " AND OP.blnPago = " & mstrSituacao
    strSQL = strSQL & " AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "OP.dtmdata") & " = " & gintExercicio
    
    strSQL = strSQL & " ) TMP "
    
    'If chkEstorno.Value = 1 Then
    '    strSQL = strSQL & "WHERE TMP.Tipo = 0 Or TMP.Tipo IS NULL "
    'End If
    
    strSQL = strSQL & " ORDER BY TMP.intNumero"
    strQueryOrdemPagamentoEmpenho = strSQL
End Function

Private Function strQueryOrdemPagamentoDespesa(Optional intOPEX As Variant) As String
    Dim strSQL       As String
    Dim mstrSituacao As String
    
    mstrSituacao = IIf(chkEstorno.Value = 0, "0", "1")
    
    strSQL = ""
'    strSQL = strSQL & "SELECT OP.PKId, OP.intNumero "
'    strSQL = strSQL & "FROM "
'    strSQL = strSQL & gstrOrdemPagamento & " OP "
'    strSQL = strSQL & "WHERE OP.bytTipo = 2 "
'    strSQL = strSQL & "AND OP.blnPago = " & mstrSituacao
'    strSQL = strSQL & " ORDER BY OP.intNumero"

    strSQL = strSQL & " SELECT G.PKID, G.INTNUMERO FROM ("
    strSQL = strSQL & " SELECT TMP.PKID, TMP.INTNUMERO , COUNT(*) qdeDespesas, SUM (Evento) Evento FROM ("
    strSQL = strSQL & " SELECT OP.PKId, OP.intNumero,"
    strSQL = strSQL & " (SELECT 1"
    strSQL = strSQL & " FROM " & gstrEventoContaContabilDebito & " EC"
    strSQL = strSQL & " WHERE DP.intContaContabil = EC.intContaContabil"
    strSQL = strSQL & " AND EC.intEvento =" & gstrItemData(cbo_intEvento) & ") Evento"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrOrdemPagamento & " OP ,"
    strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DP,"
    strSQL = strSQL & gstrOrdemPagamentoDespesaExtra & " OD"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " DP.Pkid = OD.intDespesaExtraOrcamentaria"
    strSQL = strSQL & " AND OP.PKID = OD.intOrdemPagamento"
    strSQL = strSQL & " AND OP.intExercicio = " & txt_intExercicioDE
    
    If Not IsMissing(intOPEX) Then
       strSQL = strSQL & " AND OP.INTNUMERO = " & intOPEX
    End If
    
    strSQL = strSQL & " AND OP.bytTipo = 2 AND OP.blnPago =" & mstrSituacao
    strSQL = strSQL & " AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    'strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "OP.dtmdata") & " = " & gintExercicio
    strSQL = strSQL & " ) TMP"
    strSQL = strSQL & " GROUP BY TMP.PKId, TMP.intNumero) G"
    strSQL = strSQL & " WHERE G.qdeDespesas = G.Evento"
    strSQL = strSQL & " ORDER BY G.intNumero"
    strQueryOrdemPagamentoDespesa = strSQL
End Function


Private Function strQueryOrdemPagamentoAnulacao(Optional intOPEX As Variant) As String
    Dim strSQL       As String
    Dim mstrSituacao As String
    
    mstrSituacao = IIf(chkEstorno.Value = 0, "0", "1")
    
    strSQL = ""
    strSQL = strSQL & " SELECT OP.PKId, OP.intNumero "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " OP.bytTipo = 3 AND OP.blnPago =" & mstrSituacao
    
    If Not IsMissing(intOPEX) Then
       strSQL = strSQL & " AND OP.INTNUMERO = " & intOPEX
    End If
    
    strSQL = strSQL & " AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "OP.dtmdata") & " = " & gintExercicio
    strSQL = strSQL & " ORDER BY OP.intNumero"
    strQueryOrdemPagamentoAnulacao = strSQL
End Function


Private Function strQueryOrdemPagamentoResto(Optional intOPR As Variant) As String
    Dim strSQL       As String
    Dim mstrSituacao As String
    
    mstrSituacao = IIf(chkEstorno.Value = 0, "0", "1")
    
    strSQL = ""
    
    'Select que abriga todos
    strSQL = strSQL & "SELECT TMP.PKId, TMP.intNumero FROM ("
    strSQL = strSQL & "SELECT OP.PKId, OP.intNumero ,"
    
    'Select equivalente ao campo Tipo
    strSQL = strSQL & "(SELECT SUM(" & gstrCONVERT(CDT_INT, "TE.bytAdiantamento") & ") FROM "
    strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OE, "
    strSQL = strSQL & gstrSubempenho & " SE, "
    strSQL = strSQL & gstrEmpenho & " E, " & gstrTipoEmpenho & " TE "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " OE.intOrdemPagamento = OP.Pkid"
    strSQL = strSQL & " AND OE.intParcela = SE.PKID"
    strSQL = strSQL & " AND E.PKID = SE.intEmpenho"
    strSQL = strSQL & " AND TE.pkid = E.intTipo) tipo "
        
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OP.bytTipo = 1 "
    strSQL = strSQL & "AND OP.blnPago = " & mstrSituacao
    strSQL = strSQL & " AND OP.intExercicio = " & txt_intExercicioRP
    
    If Not IsMissing(intOPR) Then
       strSQL = strSQL & " AND OP.INTNUMERO = " & intOPR
    End If
    
    strSQL = strSQL & " AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "OP.dtmdata") & " = " & txt_intExercicioRP
    
    strSQL = strSQL & " ) TMP "
    strSQL = strSQL & "WHERE TMP.Tipo = 0 Or TMP.Tipo IS NULL "
    
    strSQL = strSQL & "ORDER BY TMP.intNumero"
    
    strQueryOrdemPagamentoResto = strSQL
    
End Function

Private Function strQueryDespesa() As String
    Dim strSQL          As String
    Dim mstrSituacao As String
    mstrSituacao = IIf(chkEstorno.Value = 0, "0", "2")
    
    strSQL = ""
    strSQL = strSQL & "SELECT DP.PKId, DP.intNumero "
    strSQL = strSQL & "FROM " & gstrDespesaExtraOrcamentaria & " DP "
    strSQL = strSQL & "WHERE bytSituacao = " & mstrSituacao
     
    strSQL = strSQL & " AND NOT DP.PKID IN "
    
    strSQL = strSQL & "(SELECT intDespesaExtraOrcamentaria FROM "
    strSQL = strSQL & gstrOrdemPagamentoDespesaExtra & " OPD, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPD.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intDespesaExtraOrcamentaria) "

    strSQL = strSQL & " AND intContaContabil IN ("
    strSQL = strSQL & " SELECT PC.PKId FROM " & gstrPlanoConta & " PC, " & gstrEventoContaContabilDebito & " EC"
    strSQL = strSQL & " WHERE PC.PKId = EC.intContaContabil AND EC.intEvento =" & gstrItemData(cbo_intEvento) & ")"

    strSQL = strSQL & " ORDER BY DP.intNumero"
    strQueryDespesa = strSQL
End Function


Private Function strQueryPlanoConta() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " "
    strSQL = strSQL & "WHERE ABS(blnFinanceira) = 1 "
    strSQL = strSQL & "AND ABS(blnAnalitica) = 1 "
    strQueryPlanoConta = strSQL
End Function

Private Sub cbo_HistoricoLiquidacao_Click()
Dim adoResultado As ADODB.Recordset
Dim strSQL As String

    If cbo_HistoricoLiquidacao.ListIndex > -1 Then
    
        strSQL = ""
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & " H.strcodigo "
        strSQL = strSQL & " FROM "
        strSQL = strSQL & " tblhistorico H "
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " H.PKID = " & gstrItemData(cbo_HistoricoLiquidacao) & ""
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
            With adoResultado
                If Not .EOF Then
                    txt_strCodigoHistorico.Text = gstrENulo(!strCodigo)
                    txtHistoricoLiquidacao.Text = cbo_HistoricoLiquidacao.Text
                End If
            End With
        End If
    End If
    
End Sub

Private Sub cbointContaContabil_Change()
    ProcuraConta
End Sub

Private Sub cbointContaContabil_Click()
    cbostrContaContabil.ListIndex = gintIndiceCBO(cbostrContaContabil, _
                                        gstrItemData(cbointContaContabil))
    If Len(Trim(cbointContaContabil.Text)) > 0 Then
        If blnIncremtCheque Then
            txtNumCheque.Text = gstrProximoCheque(gstrItemData(cbointContaContabil), cbointContaContabil.Text)
        End If
    End If
End Sub

Private Sub cbointContaContabil_GotFocus()
    If cbointContaContabil.ListCount = 0 Then PreencheComboConta
End Sub

Private Sub cbointContaContabil_Validate(Cancel As Boolean)
Dim intFor As Integer
    
    For intFor = 0 To cbointContaContabil.ListCount - 1
        If cbointContaContabil.Text = cbointContaContabil.list(intFor) Then
            cbointContaContabil.ListIndex = intFor
            Exit For
        End If
    Next

End Sub

Private Sub cbostrContaContabil_Click()
    Dim tempIndice As Integer
        
On Error GoTo Problema

    tempIndice = cbostrContaContabil.ListIndex
    
    cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, _
                                    gstrItemData(cbostrContaContabil))
                                    
   If cbointContaContabil.ListIndex = -1 Then
'        LePlanoContaGeral cbointContaContabil, cbostrContaContabil, "FN"
        PreencheComboConta
        cbostrContaContabil.ListIndex = tempIndice
        cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, _
                                        gstrItemData(cbostrContaContabil))
   End If
   
   If cbostrContaContabil.ListIndex = -1 Then
       cbointContaContabil.ListIndex = -1
   Else
       If cbointContaContabil.Text <> "" And gstrItemData(cbointContaContabil) <> 0 Then
          txt_saldoAtual = strValorAtualdaConta(True)
       End If
   End If

   
Problema:
    If Err.Number = 380 Then
        Exit Sub
    End If
End Sub

Private Sub cbostrContaContabil_GotFocus()
    If mblnCarregaFormConta = True Then
        mblnCarregaFormConta = False
        If cbostrContaContabil.ListIndex = -1 Then cbointContaContabil.ListIndex = -1
    End If
End Sub

Private Sub chkEstorno_Click()
    LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
    LeDaTabelaParaObj "", dcbResto, strQueryResto
    LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
    LeDaTabelaParaObj "", dcbOrdemPagamentoAnulacao, strQueryOrdemPagamentoAnulacao
    
    LeDaTabelaParaObj "", dcbOrdemPagamentoEmpenho, strQueryOrdemPagamentoEmpenho
    If txt_intExercicioRP <> "" Then
        LeDaTabelaParaObj "", dcbOrdemPagamentoResto, strQueryOrdemPagamentoResto
    End If
    
    If txt_intExercicioDE <> "" Then
        LeDaTabelaParaObj "", dcbOrdemPagamentoDespesa, strQueryOrdemPagamentoDespesa
    End If
    
    lvw_Empenho.ListItems.Clear
    lvw_Resto.ListItems.Clear
    lvw_Despesa.ListItems.Clear
    lvw_AnulacaoReceita.ListItems.Clear
    
    txtTotalEmpenho = "0,00"
    txtTotalResto = "0,00"
    txtTotalDespesa = "0,00"
    txtTotalReceitaExtra = "0,00"
    txtTotalRecOrdem = "0,00"
    txtTotalAPagar = "0,00"
    txtTotal = IIf(chkEstorno.Value = 1, Val(gstrConvVrParaSql(txtTotal)) * -1, txtTotal)
    
End Sub

Private Sub cmd_Despesa_Click()
    CarregaForm frmCadDespesaExtraOrcamentaria, dcbDespesa, strQueryDespesa
End Sub

Private Sub cmd_Empenho_Click()
    frmCadEmpenho.mblnRestosAPagar = False
    CarregaForm frmCadEmpenho, dcbEmpenho, strQueryEmpenho
End Sub

Private Sub cmd_HistoricoLiquidacao_Click()
    CarregaForm frmCadHistorico, cbo_HistoricoLiquidacao
End Sub

Private Sub cmd_OrdemPagamentoAnulacao_Click()
    CarregaForm frmCadOrdemPagamento, dcbOrdemPagamentoAnulacao, strQueryOrdemPagamentoAnulacao
End Sub

Private Sub cmd_OrdemPagamentoDespesa_Click()
    CarregaForm frmCadOrdemPagamento, dcbOrdemPagamentoDespesa, strQueryOrdemPagamentoDespesa
End Sub

Private Sub cmd_OrdemPagamentoEmpenho_Click()
    CarregaForm frmCadOrdemPagamento, dcbOrdemPagamentoEmpenho, strQueryOrdemPagamentoEmpenho
End Sub

Private Sub cmd_OrdemPagamentoResto_Click()
    CarregaForm frmCadOrdemPagamento, dcbOrdemPagamentoResto, strQueryOrdemPagamentoResto
End Sub

Private Sub cmd_PlanoConta_Click()
    mblnCarregaFormConta = True
    CarregaForm frmCadPlanoConta, cbostrContaContabil, strQueryPlanoConta
End Sub

Private Sub cmd_Resto_Click()
    'CarregaForm frmCadRestoAPagar1, dcbResto
    frmCadEmpenho.mblnRestosAPagar = True
    CarregaForm frmCadEmpenho, dcbResto
End Sub

Private Sub dcbDespesa_Change()
    ProcuraDespesa
    
    Dim strGuardaValor As String
    strGuardaValor = dcbDespesa.Text
    dcbOrdemPagamentoDespesa.Text = ""
    dcbDespesa.Text = strGuardaValor
    
End Sub

Private Sub dcbDespesa_Click(Area As Integer)
    ProcuraDespesa
End Sub

Private Sub dcbDespesa_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 2
End Sub

Private Sub dcbEmpenho_Change()
Dim strGuardaValor As String
    
    strGuardaValor = dcbEmpenho.Text
    LeDaTabelaParaObj "", dcbParcela, strSqlParcelaEmpenho
    dcbOrdemPagamentoEmpenho.Text = ""
    dcbEmpenho.Text = strGuardaValor
    dcbEmpenho.SelStart = Len(dcbEmpenho.Text)
    
End Sub

Private Sub dcbEmpenho_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 0
End Sub

Private Sub dcbOrdemPagamentoAnulacao_Change()
    Dim strGuardaValor As String
    strGuardaValor = dcbOrdemPagamentoAnulacao.Text
    dcbOrdemPagamentoAnulacao.Text = strGuardaValor
    ProcuraParcelaOrdemAnulacao
End Sub

Private Sub dcbOrdemPagamentoAnulacao_LostFocus()
   Dim strOPE As String
   If Len(Trim(dcbOrdemPagamentoAnulacao)) > 0 Then
      strOPE = dcbOrdemPagamentoAnulacao.Text
      LeDaTabelaParaObj "", dcbOrdemPagamentoAnulacao, strQueryOrdemPagamentoAnulacao(Val(dcbOrdemPagamentoAnulacao.Text))
      dcbOrdemPagamentoAnulacao.Text = strOPE
      DropDownDataCombo dcbOrdemPagamentoAnulacao, Me
   End If
End Sub

Private Sub dcbOrdemPagamentoDespesa_Change()
    Dim strGuardaValor As String
    strGuardaValor = dcbOrdemPagamentoDespesa.Text
    dcbDespesa.Text = ""
    dcbOrdemPagamentoDespesa.Text = strGuardaValor
    ProcuraParcelaOrdemDespesa
End Sub

Private Sub dcbOrdemPagamentoDespesa_LostFocus()
   Dim strOPEX As String
   
   If Len(Trim(dcbOrdemPagamentoDespesa)) > 0 Then
      strOPEX = dcbOrdemPagamentoDespesa.Text
      
      If Len(Trim(txt_intExercicioDE)) > 0 Then
         LeDaTabelaParaObj "", dcbOrdemPagamentoDespesa, strQueryOrdemPagamentoDespesa(Val(dcbOrdemPagamentoDespesa.Text))
      Else
         ExibeMensagem "Informe o exercicio antes de efetuar o preenchimento desta lista."
         txt_intExercicioDE.SetFocus
      End If
      dcbOrdemPagamentoDespesa.Text = strOPEX
      DropDownDataCombo dcbOrdemPagamentoDespesa, Me
   End If

End Sub

Private Sub dcbOrdemPagamentoEmpenho_Change()
Dim strGuardaValor As String
    
    strGuardaValor = dcbOrdemPagamentoEmpenho.Text
    dcbEmpenho.Text = ""
    dcbOrdemPagamentoEmpenho.Text = strGuardaValor
    
    ProcuraParcelaOrdemEmpenho
    
    If dcbOrdemPagamentoEmpenho.Text <> "" Then
        PegaCredor
    End If
    
    dcbOrdemPagamentoEmpenho.SelStart = Len(dcbOrdemPagamentoEmpenho.Text)
    
End Sub

Private Sub dcbOrdemPagamentoEmpenho_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 0
End Sub

Private Sub dcbOrdemPagamentoEmpenho_LostFocus()
Dim strOPE As String
   
   If Len(Trim(dcbOrdemPagamentoEmpenho)) > 0 Then
      strOPE = dcbOrdemPagamentoEmpenho.Text
      LeDaTabelaParaObj "", dcbOrdemPagamentoEmpenho, strQueryOrdemPagamentoEmpenho(Val(dcbOrdemPagamentoEmpenho.Text))
      dcbOrdemPagamentoEmpenho.Text = strOPE
      DropDownDataCombo dcbOrdemPagamentoEmpenho, Me
   End If

End Sub

Private Sub dcbOrdemPagamentoResto_Change()
    Dim strGuardaValor As String
    strGuardaValor = dcbOrdemPagamentoResto.Text
    dcbResto.Text = ""
    dcbOrdemPagamentoResto.Text = strGuardaValor
    ProcuraParcelaOrdemResto
End Sub

Private Sub dcbOrdemPagamentoResto_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 1
End Sub

Private Sub dcbOrdemPagamentoResto_LostFocus()
   Dim strOPR As String
   
   If Len(Trim(dcbOrdemPagamentoResto)) > 0 Then
      strOPR = dcbOrdemPagamentoResto.Text
      
      If Len(Trim(txt_intExercicioRP)) > 0 Then
         LeDaTabelaParaObj "", dcbOrdemPagamentoResto, strQueryOrdemPagamentoResto(Val(dcbOrdemPagamentoResto.Text))
      Else
         ExibeMensagem "Informe o exercicio antes de efetuar o preenchimento desta lista."
         txt_intExercicioRP.SetFocus
      End If
      dcbOrdemPagamentoResto.Text = strOPR
      DropDownDataCombo dcbOrdemPagamentoResto, Me
   End If

End Sub

Private Sub dcbParcela_Change()
    ProcuraParcelaEmpenho
End Sub

Private Sub dcbParcela_Click(Area As Integer)
    ProcuraParcelaEmpenho
End Sub

Private Sub dcbParcela_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 0
End Sub

Private Sub dcbParcelaResto_Change()
    ProcuraParcelaResto
End Sub

Private Sub dcbParcelaResto_Click(Area As Integer)
    ProcuraParcelaResto
End Sub

Private Sub dcbParcelaResto_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 1
End Sub

Private Sub dcbResto_Change()
    Dim strGuardaValor As String
    strGuardaValor = dcbResto.Text
    dcbOrdemPagamentoResto.Text = ""
    dcbResto.Text = strGuardaValor
    
    If dcbResto.Text <> "" Then
        LeDaTabelaParaObj "", dcbParcelaResto, strQuerySubempenho
    End If
End Sub

Private Sub dcbResto_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 1
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 286
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, _
                             gstrExcluirItem, gstrSalvar
    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar
    
    If mblnselecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar, gstrImprimir
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrImprimir
    End If
    
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
    If mblnAlterando = True Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
    TrocaCorObjeto cbointContaContabil, False
    TrocaCorObjeto cbostrContaContabil, False
    TrocaCorObjeto cmd_PlanoConta, False
    'Pen_773
    TrocaCorObjeto txt_Cdc, True
    TrocaCorObjeto txt_strNome, True
    TrocaCorObjeto txtintProcesso, False, True
    txtintProcesso.BackColor = -2147483643
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    mblnAlterando = False
    blnOrdenacaoAsc = True
    VerificaListaAutomatica gstrPagamento, tdb_Lista, gstrQueryLocalizar
    VerificaObjParaAplicar mobjAux

    LimpaTelaPagamento
    
    preencheCboevento
    
    If cbo_intEvento.ListCount = 1 Then
        cbo_intEvento.ListIndex = 0
    End If
    
    lvw_Despesa.ColumnHeaders(7).Position = 3
    
    habilitaGuias 3
    tab_3DPastaEmpenho.TabEnabled(4) = False
    TrocaCorObjeto txt_saldoAtual, True
    LocalDesabilitaNumeroParcela
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    blnPrimeiraVez = False
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lvw_AnulacaoReceita_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 5
End Sub

Private Sub lvw_AnulacaoReceita_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Not dcbOrdemPagamentoAnulacao.MatchedWithList Then
        LeDaTabelaParaObj "", dcbOrdemPagamentoAnulacao, strQueryOrdemPagamentoAnulacao
    End If
    
    If Not lvw_AnulacaoReceita.SelectedItem Is Nothing Then
        dcbOrdemPagamentoAnulacao.Text = lvw_AnulacaoReceita.SelectedItem.SubItems(1)
    End If
    
End Sub

Private Sub lvw_Conta_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_Conta
        If cbointContaContabil.ListCount = 0 Then
            PreencheComboConta
        End If
        cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, .SelectedItem.Tag)
'        cbointContaContabil.Text = .SelectedItem.Text
        txtValorLancamento = .ListItems(.SelectedItem.Index).SubItems(2)
        txtNumCheque = .ListItems(.SelectedItem.Index).SubItems(3)
    End With
    strValorAtualdaConta (True)
    mblnAlterandoConta = True
    
    'Incluido na pendência orc1585
    If blnIncremtCheque Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        TrocaCorObjeto cbointContaContabil, True, False
        TrocaCorObjeto cbostrContaContabil, True, False
        TrocaCorObjeto txtValorLancamento, True, False
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        TrocaCorObjeto cbointContaContabil, False, False
        TrocaCorObjeto cbostrContaContabil, False, False
        TrocaCorObjeto txtValorLancamento, False, False
    End If
    
End Sub

Private Sub lvw_Despesa_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 2
End Sub

Private Sub lvw_Despesa_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnAlterandoDespesa = True
    With lvw_Despesa
        If Not dcbDespesa.MatchedWithList Then
            LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
        End If
        dcbDespesa = .ListItems(.SelectedItem.Index).Text
    End With
    
    
    With lvw_Despesa
        If .ListItems(.SelectedItem.Index).SubItems(1) = "--" Then
            If Not dcbDespesa.MatchedWithList Then
                LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
            End If
            dcbDespesa = .ListItems(.SelectedItem.Index).Text
            
        Else
            txt_intExercicioDE.Text = .ListItems(.SelectedItem.Index).SubItems(6)
            If Not dcbOrdemPagamentoDespesa.MatchedWithList Then
                LeDaTabelaParaObj "", dcbOrdemPagamentoDespesa, strQueryOrdemPagamentoDespesa
            End If
            dcbOrdemPagamentoDespesa = .ListItems(.SelectedItem.Index).SubItems(1)
        End If
    End With
End Sub

Private Sub lvw_Empenho_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 0
End Sub

Private Sub lvw_Empenho_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnAlterandoEmpenho = True
    
    With lvw_Empenho
        If .ListItems(.SelectedItem.Index).SubItems(1) = "--" Then
            If Not dcbEmpenho.MatchedWithList Then
                LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
            End If
            dcbEmpenho = .ListItems(.SelectedItem.Index).SubItems(2)
            dcbParcela = .ListItems(.SelectedItem.Index).SubItems(3)
            
        Else
            If Not dcbOrdemPagamentoEmpenho.MatchedWithList Then
                LeDaTabelaParaObj "", dcbOrdemPagamentoEmpenho, strQueryOrdemPagamentoEmpenho
            End If
            dcbOrdemPagamentoEmpenho = .ListItems(.SelectedItem.Index).SubItems(1)
        End If
        
    End With
End Sub

Private Sub CancelaPagamento()

'******************************************************************************************
' Data: 11/06/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL
'            permitindo, assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL      As String
    If gblnDataValida(txtdtmDataAnulacao) Then
    
        'Orc677
        If Right(txtdtmDataAnulacao, 4) <> gintExercicio Then
            ExibeMensagem "A data de anulação não equivale a data do exercício corrente."
            If txtdtmDataAnulacao.Enabled Then txtdtmDataAnulacao.SetFocus
            Exit Sub
        End If
    
        If gblnExclusaoGravacaoOk("I", "Confirma cancelamento do pagamento?", True) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
        'Registra o cancelamento do processo de pagamento
            strSQL = ""
            
            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
            
            strSQL = strSQL & "INSERT INTO " & gstrProcessoPagtoAnulado & " ("
            strSQL = strSQL & "intProcesso, dtmData, strHistorico, "
            strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr"
            strSQL = strSQL & ") VALUES ("
            strSQL = strSQL & tdb_Lista.Columns("PKId") & ", "
            strSQL = strSQL & gstrConvDtParaSql(txtdtmDataAnulacao) & ", "
            strSQL = strSQL & "'" & txtHistoricoLiquidacao & "', "
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSQL = strSQL & glngCodUsr & ");"
            
        'Cancela o processo de pagamento
            strSQL = strSQL & "UPDATE " & gstrProcessoPagamento & " SET "
            strSQL = strSQL & "bytSituacao = 1 "
            strSQL = strSQL & "WHERE PKId = " & tdb_Lista.Columns("PKId") & ";"
            
        'Registra o cancelamento do pagamento das parcela do empenho
            strSQL = strSQL & "INSERT INTO " & gstrSubempenhoPagtoAnulado & " ("
            strSQL = strSQL & "intSubempenho, intEmpenho, intNumero, dtmData, "
            strSQL = strSQL & "intProcesso, dblValor, bytTipo, dtmLiquidacao, "
            strSQL = strSQL & "dtmPagamento, dtmAnulacaoPagamento, "
            strSQL = strSQL & "strHistorico, dtmDtAtualizacao, lngCodUsr) "
            strSQL = strSQL & "SELECT PKId, intEmpenho, intNumero, dtmData, "
            strSQL = strSQL & "intProcesso, dblValor, bytTipo, dtmLiquidacao, "
            strSQL = strSQL & "dtmPagamento, "
            strSQL = strSQL & gstrConvDtParaSql(txtdtmDataAnulacao) & ", "
            strSQL = strSQL & "strHistorico, "
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSQL = strSQL & glngCodUsr & " FROM " & gstrSubempenho & " "
            strSQL = strSQL & "WHERE intProcesso = "
            strSQL = strSQL & tdb_Lista.Columns("PKId") & ";"
            
        'cancela do pagamento das parcela do empenho
            strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
            strSQL = strSQL & "dtmPagamento = NULL, bytSituacao = 2, " 'Liquidada
            strSQL = strSQL & "intProcesso = NULL "
            strSQL = strSQL & "WHERE intProcesso = "
            strSQL = strSQL & tdb_Lista.Columns("PKId") & ";"
            
        'Registra o cancelamento do pagamento das parcela do resto a pagar
            strSQL = strSQL & "INSERT INTO " & gstrParcelaRestoPagtoAnulado & " "
            strSQL = strSQL & "SELECT intResto, intNumero, dtmData, "
            strSQL = strSQL & "dblValor, dtmPagamento, intProcesso, "
            strSQL = strSQL & gstrConvDtParaSql(txtdtmDataAnulacao) & ", "
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSQL = strSQL & glngCodUsr & " FROM " & gstrParcelaRestoPagar & " "
            strSQL = strSQL & "WHERE intProcesso = "
            strSQL = strSQL & tdb_Lista.Columns("PKId") & ";"
            
        'cancela do pagamento das parcela do resto a pagar
            strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
            strSQL = strSQL & "dtmPagamento = NULL, bytSituacao = 1, " 'Processada
            strSQL = strSQL & "intProcesso = NULL "
            strSQL = strSQL & "WHERE intProcesso = "
            strSQL = strSQL & tdb_Lista.Columns("PKId") & ";"
            
        'Registra o cancelamento do pagamento da Despesa Extra-orçamentária
            strSQL = strSQL & "INSERT INTO " & gstrDespesaExtraOrcamPagtoAnulado & " "
            strSQL = strSQL & "SELECT intNumero, intContribuinte, "
            strSQL = strSQL & "intContaContabil, dblValor, dtmData, "
            strSQL = strSQL & "dtmPagamento , intProcesso, strHistorico, "
            strSQL = strSQL & gstrConvDtParaSql(txtdtmDataAnulacao) & ", "
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSQL = strSQL & glngCodUsr & " FROM " & gstrDespesaExtraOrcamentaria & " "
            strSQL = strSQL & "WHERE intProcesso = "
            strSQL = strSQL & tdb_Lista.Columns("PKId") & ";"
        
        'cancela do pagamento da Despesa Extra-orçamentária
            strSQL = strSQL & "UPDATE " & gstrDespesaExtraOrcamentaria & " SET "
            strSQL = strSQL & "dtmPagamento = NULL, bytSituacao = 0, " 'Programada
            strSQL = strSQL & "intProcesso = NULL "
            strSQL = strSQL & "WHERE intProcesso = "
            strSQL = strSQL & tdb_Lista.Columns("PKId") & ";"
            
            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
            
            If gobjBanco.Execute(strSQL) Then
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaCommitTrans
                txtdtmDataAnulacao = ""
                txtdtmDataAnulacao.SetFocus
                tdb_Lista.Refresh
                'AtualizaListas
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
            End If
        End If
    Else
        ExibeMensagem "A data do cancelamento tem que ser informada corretamente."
        txtdtmDataAnulacao.SetFocus
    End If
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    Dim strSQL As String

    Select Case UCase(strModoOperacao)
        Case gstrNovo
            LimpaTelaPagamento True, 1
            LocalDesabilitaNumeroParcela
        Case gstrSalvar
            GravaPagamento
            
        Case UCase(gstrIncluirItem)
            
            If Me.ActiveControl.Name = dcbOrdemPagamentoEmpenho.Name Then
               dcbOrdemPagamentoEmpenho_LostFocus
            End If
            
            If Me.ActiveControl.Name = dcbOrdemPagamentoResto.Name Then
               dcbOrdemPagamentoResto_LostFocus
            End If
            
            If Me.ActiveControl.Name = dcbOrdemPagamentoDespesa.Name Then
               dcbOrdemPagamentoDespesa_LostFocus
            End If
            
            If Me.ActiveControl.Name = dcbOrdemPagamentoAnulacao.Name Then
               dcbOrdemPagamentoAnulacao_LostFocus
            End If
            
            VerificaLista
            
            
        Case UCase(gstrExcluirItem)
            VerificaListaExcluir
        Case UCase(gstrCancelar)
            CancelaPagamento
        Case UCase(gstrImprimir)
            ImprimeBordereaux Val(txtPKID)
        Case UCase(gstrPreencherLista)
            If ActiveControl.Name = cbo_HistoricoLiquidacao.Name Then
                LeDaTabelaParaObj gstrHistorico, cbo_HistoricoLiquidacao
            End If
            
            If ActiveControl.Name = cbointContaContabil.Name Or ActiveControl.Name = cbostrContaContabil.Name Then
    '            If cbo_intEvento.ListIndex = -1 Then
    '                ExibeMensagem "Para preencher a lista de Contas é necessário informar o Evento"
    '                If cbo_intEvento.Enabled Then cbo_intEvento.SetFocus
    '            Else
                    PreencheComboConta
    '            End If
            End If
            
            If ActiveControl.Name = dcbEmpenho.Name Then
                LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
            End If
            
            If ActiveControl.Name = dcbResto.Name Then
               LeDaTabelaParaObj "", dcbResto, strQueryResto
            End If
            
            If ActiveControl.Name = dcbDespesa.Name Then
               LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
            End If
            
            If ActiveControl.Name = dcbOrdemPagamentoEmpenho.Name Then
                LeDaTabelaParaObj "", dcbOrdemPagamentoEmpenho, strQueryOrdemPagamentoEmpenho
            End If
            
            If ActiveControl.Name = dcbOrdemPagamentoAnulacao.Name Then
                LeDaTabelaParaObj "", dcbOrdemPagamentoAnulacao, strQueryOrdemPagamentoAnulacao
            End If
            
            If ActiveControl.Name = dcbOrdemPagamentoResto.Name Then
                If Len(Trim(txt_intExercicioRP)) > 0 Then
                   LeDaTabelaParaObj "", dcbOrdemPagamentoResto, strQueryOrdemPagamentoResto
                Else
                   ExibeMensagem "Informe o exercicio antes de efetuar o preenchimento desta lista."
                   txt_intExercicioRP.SetFocus
                End If
            End If
            
            If ActiveControl.Name = dcbOrdemPagamentoDespesa.Name Then
               If Len(Trim(txt_intExercicioDE)) > 0 Then
                  LeDaTabelaParaObj "", dcbOrdemPagamentoDespesa, strQueryOrdemPagamentoDespesa
               Else
                  ExibeMensagem "Informe o exercicio antes de efetuar o preenchimento desta lista."
                  txt_intExercicioDE.SetFocus
               End If
            End If
                
            
            If Me.ActiveControl.Name = cbo_intEvento.Name Then
                preencheCboevento
            End If
            
            
        Case UCase(gstrLocalizar)
            

            
            LeDaTabelaParaObj "", tdb_Lista, gstrQueryLocalizar
            'PreencheGridPagamento
            'FiltraCampos tdb_Lista
            
    End Select
End Sub

Private Function gstrQueryLocalizar() As String

        If bytDBType = Oracle Then
            gstrQueryLocalizar = gstrStoredProcedure("sp_AnulacaoPagamento", CStr(gintExercicio) & ", 0," & IIf(Trim(txtintProcesso) <> "", Trim(txtintProcesso), 0) & "," & IIf(Trim(txtData) <> "", "'" & txtData & "',1", "'" & gstrDataDoSistema & "', 0"), True)
        Else
            gstrQueryLocalizar = gstrStoredProcedure("sp_AnulacaoPagamento", CStr(gintExercicio) & ", 0," & IIf(Trim(txtintProcesso) <> "", Trim(txtintProcesso), 0) & "," & IIf(Trim(txtData) <> "", gstrConvDtParaSql(txtData) & ",1", gstrConvDtParaSql(gstrDataDoSistema) & ", 0"), True)
        End If
End Function

Private Function strQueryResto()
    Dim strSQL  As String
    Dim mstrSituacao As String
    
    mstrSituacao = IIf(chkEstorno.Value = 0, "2", "3")
    
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT EP.PKId, EP.intNumero , "
    strSQL = strSQL & "EP.dtmData,EP.dblValor ,PT.strCodigo "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP, "
    strSQL = strSQL & gstrTipoEmpenho & " TE, "
    strSQL = strSQL & gstrSubempenho & " SE, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
    strSQL = strSQL & "WHERE EP.intProgramaTrabalho = PT.PKId "
    strSQL = strSQL & "AND EP.PKId = SE.intEmpenho "
    strSQL = strSQL & "AND SE.bytSituacao = " & mstrSituacao
    strSQL = strSQL & " AND intExercicioRP = " & gintExercicio
    
    strSQL = strSQL & " AND NOT SE.PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoResto & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela) "
    
    strSQL = strSQL & " AND EP.intTipo = TE.PKID "
    strSQL = strSQL & " AND (TE.bytAdiantamento IS NULL OR TE.bytAdiantamento=0)"
    
    strSQL = strSQL & " ORDER BY EP.intNumero, EP.dtmData, PT.strCodigo "
    strQueryResto = strSQL
End Function

Private Sub lvw_Resto_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 1
End Sub

Private Sub lvw_Resto_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_Resto
        If .ListItems(.SelectedItem.Index).SubItems(1) = "--" Then
            If Not dcbResto.MatchedWithList Then
                LeDaTabelaParaObj "", dcbResto, strQueryResto
            End If
            dcbResto = .ListItems(.SelectedItem.Index).SubItems(3)
            dcbParcelaResto = .ListItems(.SelectedItem.Index).SubItems(4)
            
            
        Else
            txt_intExercicioRP = .ListItems(.SelectedItem.Index).SubItems(9)
            If Not dcbOrdemPagamentoResto.MatchedWithList Then
                LeDaTabelaParaObj "", dcbOrdemPagamentoResto, strQueryOrdemPagamentoResto
            End If
            dcbOrdemPagamentoResto = .ListItems(.SelectedItem.Index).SubItems(1)
        End If
    End With
End Sub

Private Sub tab_3DPastaEmpenho_Click(PreviousTab As Integer)
'    If tab_3DPastaEmpenho.Tab = 4 Then

'        TrocaCorObjeto txtData, True
'        TrocaCorObjeto txtHistoricoLiquidacao, True
'        TrocaCorObjeto cbo_HistoricoLiquidacao, True
'        TrocaCorObjeto cmd_HistoricoLiquidacao, True
'    Else
'        TrocaCorObjeto txtintProcesso, False
'        TrocaCorObjeto txtData, False
'        TrocaCorObjeto txtHistoricoLiquidacao, False
'        TrocaCorObjeto cbo_HistoricoLiquidacao, False
'        TrocaCorObjeto cmd_HistoricoLiquidacao, False
'    End If
   If tab_3DPastaEmpenho.Tab = 3 Then
      txtValorLancamento = txtTotalAPagar
   End If
End Sub

Private Sub tdb_Lista_Click()
    blnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo saida
    
    Dim recGrid1 As New ADODB.Recordset
    
    Set recGrid = tdb_Lista.DataSource
    
    blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
    bytOrdenacao = ColIndex
    
    recGrid.Sort = tdb_Lista.Columns(ColIndex).DataField & IIf(blnOrdenacaoAsc, "  DESC", " ASC")

    

    If Not recGrid.EOF Then
        recGrid.MoveFirst
    End If
        
    Set tdb_Lista.DataSource = recGrid
    tdb_Lista.ReBind
    tdb_Lista.Refresh
saida:
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    If tdb_Lista.Col = 2 Then
        CaracterValido KeyAscii, "D", tdb_Lista
    End If
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk And blnPrimeiraVez Then
        
            TrocaCorObjeto txtintProcesso, True, False
            
            LimpaDadosConta True
            LimpaCamposOP
            mblnAlterando = True
            LocalDesabilitaNumeroParcela
            mblnClickOk = False
            txtPKID = .Columns("PKId")
            gCorLinhaSelecionada tdb_Lista
            LePagamento
            TotalizaRecOrc
            txtTotalEmpenho = IIf(chkEstorno.Value = 1 And .Columns("dblTotalEmpenho") <> "", "-", "") & .Columns("dblTotalEmpenho")
            txtTotalResto = IIf(chkEstorno.Value = 1 And .Columns("dblTotalResto") <> "", "-", "") & .Columns("dblTotalResto")
            txtTotalDespesa = IIf(chkEstorno.Value = 1 And .Columns("dblTotalDespesaExtra") <> "", "-", "") & .Columns("dblTotalDespesaExtra")
            txtHistoricoLiquidacao = .Columns("strHistorico")
            cbo_intEvento.ListIndex = gintIndiceCBO(cbo_intEvento, .Columns("intevento"))
            
            txtTotal = IIf(chkEstorno.Value = 1, "-", "") & txtTotalAPagar

            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
            HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
            
            TrocaCorObjeto txtData, True
            TrocaCorObjeto cmd_Evento, True
            TrocaCorObjeto cbo_intEvento, True
            TrocaCorObjeto txt_codEvento, True
            chkEstorno.Enabled = False
            chkBordero.Enabled = False
            
            LeDaTabelaParaObj "", lvw_Conta, strQueryLancamento
            Totaliza lvw_Conta, txtTotal
                    
            If Val(.Columns("dblTotalEmpenho")) <> 0 Then
                LeEmpenho
                AjustaFormatacaoEmpenho
                With lvw_Empenho
                If .ListItems(.SelectedItem.Index).SubItems(1) = "--" Then
                    If Not dcbEmpenho.MatchedWithList Then
                        LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
                    End If
                    dcbEmpenho = .ListItems(.SelectedItem.Index).SubItems(2)
                    dcbParcela = .ListItems(.SelectedItem.Index).SubItems(3)
            
                Else
                    If Not dcbOrdemPagamentoEmpenho.MatchedWithList Then
                        LeDaTabelaParaObj "", dcbOrdemPagamentoEmpenho, strQueryOrdemPagamentoEmpenho
                    End If
                    dcbOrdemPagamentoEmpenho = .ListItems(.SelectedItem.Index).SubItems(1)
                End If
        
                End With
            ElseIf Val(.Columns("dblTotalResto")) <> 0 Then
                LeResto
                AjustaFormatacaoResto
            ElseIf Val(.Columns("dblTotalDespesaExtra")) <> 0 Then
                LeDespesa
            ElseIf Val(.Columns("dblTotalAnulacaoReceita")) <> 0 Then
                LeAnulacaoReceita
            End If
            
            AjustaFormatacaoConta
            
            'lvw_Empenho_ItemClick (Nothing)
            
        End If
        'Pen_773
        
    End With
End Sub

Private Sub LimpaCamposOP()
    dcbOrdemPagamentoEmpenho.Text = ""
    dcbOrdemPagamentoEmpenho.BoundText = ""
    dcbEmpenho.Text = ""
    dcbParcela.Text = ""
    txt_intExercicioRP.Text = ""
    dcbOrdemPagamentoResto.Text = ""
    dcbResto.Text = ""
    dcbParcelaResto.Text = ""
    txt_intExercicioDE.Text = ""
    dcbOrdemPagamentoDespesa.Text = ""
    dcbDespesa.Text = ""
    cbointContaContabil.Text = ""
    cbostrContaContabil.Text = ""
    txt_saldoAtual.Text = ""
    txtValorLancamento.Text = ""
    txtNumCheque.Text = ""
    dcbOrdemPagamentoAnulacao.Text = ""
    txt_Cdc.Text = ""
    txt_strNome.Text = ""
End Sub



Private Sub txt_codEvento_Change()
    mblcodEventoMudou = True
End Sub

Private Sub txt_codEvento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_codEvento_LostFocus()
   ' If (txt_codEvento <> "" And mblcodEventoMudou = True) Or (cbo_intEvento.ListIndex = -1 Or cbo_intEvento.Text = "") Then
        PreencheEventobyCodigo txt_codEvento, cbo_intEvento, "3,4,5,11" '3,4,5,11,12
    '    mblcodEventoMudou = False
    'End If
End Sub

Private Sub txt_intExercicioDE_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "N", txt_intExercicioDE
End Sub

Private Sub txt_intExercicioRP_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "N", txt_intExercicioRP
End Sub

Private Sub txt_strCodigoHistorico_GotFocus()
    MarcaCampo txt_strCodigoHistorico
End Sub

Private Sub txt_strCodigoHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strCodigoHistorico, False
End Sub

Private Sub txt_strCodigoHistorico_LostFocus()
Dim adoResultado As ADODB.Recordset
Dim strSQL As String

    If Len(Trim(txt_strCodigoHistorico)) > 0 Then
        Set gobjBanco = New clsBanco
        
        strSQL = ""
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & " h.StrDescricao "
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrHistorico & " H "
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " H.STRCODIGO = '" & txt_strCodigoHistorico & "'"
        
        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
            With adoResultado
                If Not .EOF Then
                    cbo_HistoricoLiquidacao.Text = gstrENulo(!strDescricao)
                    txtHistoricoLiquidacao.Text = gstrENulo(!strDescricao)
                Else
                    cbo_HistoricoLiquidacao = ""
                    txt_strCodigoHistorico = ""
                End If
            End With
        End If
    End If
    
End Sub

Private Sub txtData_GotFocus()
    MarcaCampo txtData
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtData
End Sub

Private Sub txtData_LostFocus()

    txtData = gstrDataFormatada(txtData)
    
    'ORC677
    If IsDate(txtData) Then
        If Year(CDate(txtData)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data tem que estar no exercício de " & gintExercicio & "."
            If txtData.Enabled Then txtData.SetFocus
            Exit Sub
        End If
    End If
    
'    If Trim(txtData) = "" Then
'        TrocaCorObjeto cbointContaContabil, True
'        TrocaCorObjeto cbostrContaContabil, True
'        TrocaCorObjeto cmd_PlanoConta, True
'        LimpaDadosConta True
'        Exit Sub
'    ElseIf gblnDataValida(txtData) = False Then
'        ExibeMensagem "Data informada não é valida"
'        txtData.SetFocus
'        Exit Sub
'    ElseIf CDate(txtData) < CDate(strDataEncerramento) Then
'        ExibeMensagem "Data informada não pode ser menor que Data do Encerramento."
'        txtData.SetFocus
'        Exit Sub
'    ElseIf Year(CDate(txtData)) <> gintExercicio Then
'        ExibeMensagem "Data informada deve estar dentro do Exercício corrente."
'        txtData.SetFocus
'        Exit Sub
'    End If
'    TrocaCorObjeto cbointContaContabil, False
'    TrocaCorObjeto cbostrContaContabil, False
'    TrocaCorObjeto cmd_PlanoConta, False

End Sub

Private Sub txtdtmDataAnulacao_GotFocus()
    MarcaCampo txtdtmDataAnulacao
End Sub

Private Sub txtdtmDataAnulacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataAnulacao
End Sub

Private Sub txtdtmDataAnulacao_LostFocus()

    txtdtmDataAnulacao = gstrDataFormatada(txtdtmDataAnulacao)
    
    'ORC677
    If IsDate(txtdtmDataAnulacao) Then
        If Year(CDate(txtdtmDataAnulacao)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data de Anulação tem que estar no exercício de " & gintExercicio & "."
            If txtdtmDataAnulacao.Enabled Then txtdtmDataAnulacao.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txtNumCheque_GotFocus()
    MarcaCampo txtNumCheque
End Sub

Private Sub txtNumCheque_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtPKId_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Public Function strQueryRelatorio()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PR.intCodigoreduzido AS CodigoReduzido, CO.strDescricao AS CodigoOrcamentario, FR.strDescricao AS FonteRecurso, "
    strSQL = strSQL & "PR.blnEducacao AS Educacao, PR.blnFundef AS Fundef, PR.blnSaude AS Saude, PR.blnPessoal AS Pessoal, PR.dblValor  "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR, "
    strSQL = strSQL & gstrCodigoOrcamentario & " CO, "
    strSQL = strSQL & gstrFonteRecurso & " FR "
'    strSql = strSql & "WHERE PR.intCodigoOrcamentario *= CO.PKId "
    strSQL = strSQL & "WHERE PR.intCodigoOrcamentario " & strOUTJSQLServer & "= CO.PKId " & strOUTJOracle
'    strSql = strSql & "AND PR.intFonteRecurso *= FR.PKId "
    strSQL = strSQL & "AND PR.intFonteRecurso " & strOUTJSQLServer & "= FR.PKId " & strOUTJOracle
    strSQL = strSQL & "ORDER BY PR.intCodigoreduzido, CO.strDescricao, FR.strDescricao, "
    strSQL = strSQL & "PR.blnEducacao, PR.blnFundef, PR.blnSaude, PR.blnPessoal "
    strQueryRelatorio = strSQL
End Function

Private Sub txtSaldo_Change()
    txtSaldo = IIf(txtTotal = "0,00" And chkEstorno.Value = 1, txtTotalAPagar, txtSaldo)
End Sub

Private Sub txtTotal_Change()
    If txtTotal.Text <> "" And txtTotalAPagar.Text <> "" And txtTotalRecOrdem.Text <> "" Then
    
        txtSaldo = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtTotal)) - _
                               Val(gstrConvVrParaSql(txtTotalAPagar)))
    End If
End Sub

Private Sub txtTotalAPagar_Change()
If txtTotal.Text <> "" And txtTotalAPagar.Text <> "" And txtTotalRecOrdem.Text <> "" Then
    
        txtSaldo = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtTotal)) - _
                               Val(gstrConvVrParaSql(txtTotalAPagar)))
    End If
End Sub

Private Sub txtTotalDespesa_Change()
    SomaTotalAPagar
End Sub

Private Sub txtTotalEmpenho_Change()
    SomaTotalAPagar
End Sub

Private Sub txtTotalRecOrdem_Change()
    If txtTotal.Text <> "" And txtTotalAPagar.Text <> "" And txtTotalRecOrdem.Text <> "" Then
    
        txtSaldo = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtTotal)) - _
                               Val(gstrConvVrParaSql(txtTotalAPagar)))
    End If
End Sub

Private Sub txtTotalResto_Change()
    SomaTotalAPagar
End Sub

Private Sub txtValorLancamento_Change()
    Dim intPosAtual As Integer
    If InStr(1, txtValorLancamento, "-", vbTextCompare) Then
        intPosAtual = txtValorLancamento.SelStart
        txtValorLancamento = Replace(txtValorLancamento, "-", "")
        txtValorLancamento.SelStart = intPosAtual
    End If
End Sub

Private Sub txtValorLancamento_GotFocus()
    MarcaCampo txtValorLancamento
End Sub

Private Sub txtValorLancamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtValorLancamento
End Sub

Private Sub txtValorLancamento_LostFocus()
    txtValorLancamento = gstrConvVrDoSql(txtValorLancamento)
End Sub

Private Sub cbo_intevento_Click()
    leCodigoEvento txt_codEvento, cbo_intEvento
    VerificaTipoEvento
    
    Static intIndiceAnterior As Integer
    'glIgualaContas cbo_strEvento, cbointEvento
    leCodigoEvento txt_codEvento, cbo_intEvento
    mblcodEventoMudou = False
    If intIndiceAnterior = gstrItemData(cbo_intEvento) Then Exit Sub
    
    If cbo_intEvento.ListIndex <> -1 Then
        'lvw_Conta.ListItems.Clear
        'PreencheComboConta
        LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
    Else
        txt_codEvento.Text = ""
        lvw_Despesa.ListItems.Clear
        'lvw_Conta.ListItems.Clear
        'cbointContaContabil.Clear
        'cbostrContaContabil.Clear
    End If
    
    intIndiceAnterior = gstrItemData(cbo_intEvento)
    mblcodEventoMudou = False
End Sub

Private Sub cbo_intevento_GotFocus()
    If cbo_intEvento.Text = "" Then
        txt_codEvento.Text = ""
        habilitaGuias 3
    End If
    
End Sub

Private Sub cbo_intevento_LostFocus()
    If cbo_intEvento.Text = "" Then
        txt_codEvento.Text = ""
        habilitaGuias 3
        lvw_Conta.ListItems.Clear
        cbointContaContabil.Clear
        cbostrContaContabil.Clear
    End If
End Sub

Private Sub cmd_Evento_Click()
    CarregaForm frmCadEvento, cbo_intEvento, strQueryAplicarEvento
End Sub

Private Function strQueryAplicarEvento() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEvento & " "
    strSQL = strSQL & "WHERE intTipoEvento in (3,4,5) "
    strQueryAplicarEvento = strSQL
    
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra
End Function

Private Sub AjustaFormatacaoConta()
    Dim lstItem As ListItem
    For Each lstItem In lvw_Conta.ListItems
        lstItem.ListSubItems(2).Text = gstrConvVrDoSql(lstItem.ListSubItems(2).Text)
    Next
    
    If (bytDBType = EDatabases.SQLServer) Then Exit Sub
    For Each lstItem In lvw_Conta.ListItems
        lstItem.Text = gvntFormatacaoEspecifica(lstItem.Text)
    Next
End Sub

Private Sub AjustaFormatacaoEmpenho()

    Dim lstItem As ListItem
    If (bytDBType = EDatabases.SQLServer) Then Exit Sub
    For Each lstItem In lvw_Empenho.ListItems
        lstItem.ListSubItems(2).Text = Replace(lstItem.ListSubItems(2).Text, ".", "")
        lstItem.ListSubItems(6).Text = gstrConvVrDoSql(lstItem.ListSubItems(6).Text)
        lstItem.ListSubItems(7).Text = gstrConvVrDoSql(lstItem.ListSubItems(7).Text)
        lstItem.ListSubItems(8).Text = gstrConvVrDoSql(lstItem.ListSubItems(8).Text)
    Next
End Sub

Private Sub AjustaFormatacaoReceitaExtra()

    Dim lstItem As ListItem
    If (bytDBType = EDatabases.SQLServer) Then Exit Sub
    For Each lstItem In lvw_Extra.ListItems
        lstItem.Text = Replace(lstItem.Text, ".", "")
    Next
End Sub


Private Sub AjustaFormatacaoResto()
    Dim lstItem As ListItem
    If (bytDBType = EDatabases.SQLServer) Then Exit Sub
    For Each lstItem In lvw_Resto.ListItems
        lstItem.ListSubItems(3).Text = Replace(lstItem.ListSubItems(3).Text, ".", "")
        lstItem.ListSubItems(6).Text = gstrConvVrDoSql(lstItem.ListSubItems(6).Text)
        lstItem.ListSubItems(7).Text = gstrConvVrDoSql(lstItem.ListSubItems(7).Text)
        lstItem.ListSubItems(8).Text = gstrConvVrDoSql(lstItem.ListSubItems(8).Text)
    Next
End Sub

Private Sub preencheCboevento()
    LeDaTabelaParaObj gstrEvento, cbo_intEvento, "SELECT PKID, strDescricao FROM " & gstrEvento & " WHERE intTipoEvento in (3,4,5,11)"
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra
End Sub


Private Sub PreencheComboConta()
    'LePlanoContaGeral cbointContaContabil, cbostrContaContabil, "FN"
    PreenchePlanoContaGeral cbointContaContabil, cbostrContaContabil
    'LePlanoContaGeral1 cbo_temp, cbo_temp, cbointContaContabil, cbostrContaContabil, "EC" & gstrItemData(cbo_intEvento)
End Sub


Private Sub habilitaGuias(Optional ByVal intGuia As Integer)
    tab_3DPastaEmpenho.TabEnabled(0) = False
    tab_3DPastaEmpenho.TabEnabled(1) = False
    tab_3DPastaEmpenho.TabEnabled(2) = False
    tab_3DPastaEmpenho.TabEnabled(4) = False
    tab_3DPastaEmpenho.TabEnabled(5) = False
    lvw_Extra.ListItems.Clear
    
    txtTotalEmpenho.Text = "0,00"
    txtTotalResto.Text = "0,00"
    txtTotalDespesa.Text = "0,00"
    txtTotalRecOrdem.Text = "0,00"
    
    tab_3DPastaEmpenho.TabEnabled(intGuia) = True
    
    tab_3DPastaEmpenho.Tab = intGuia
    
    If intGuia = 2 Or intGuia = 3 Or intGuia = 5 Then
        tab_3DPastaEmpenho.TabEnabled(4) = False
    Else
        tab_3DPastaEmpenho.TabEnabled(4) = True
    End If
    
    If intGuia = 0 Then
        lvw_Resto.ListItems.Clear
        lvw_Despesa.ListItems.Clear
        lvw_AnulacaoReceita.ListItems.Clear
        Totaliza lvw_Empenho, txtTotalEmpenho
    End If
    
    If intGuia = 1 Then
        lvw_Empenho.ListItems.Clear
        lvw_Despesa.ListItems.Clear
        lvw_AnulacaoReceita.ListItems.Clear
        Totaliza lvw_Resto, txtTotalResto
    End If
    
    If intGuia = 2 Then
        lvw_Empenho.ListItems.Clear
        lvw_Resto.ListItems.Clear
        lvw_AnulacaoReceita.ListItems.Clear
        Totaliza lvw_Despesa, txtTotalDespesa
    End If
    
    If intGuia = 5 Then
        lvw_Empenho.ListItems.Clear
        lvw_Resto.ListItems.Clear
        lvw_Despesa.ListItems.Clear
        LeDaTabelaParaObj "", dcbOrdemPagamentoAnulacao, strQueryOrdemPagamentoAnulacao
        Totaliza lvw_Despesa, txtTotalDespesa
    End If
    
    TotalLancado
    
End Sub


Private Sub VerificaTipoEvento()
    Dim strSQL        As String
    Dim Pkid          As Integer
    Dim adoResultado  As New ADODB.Recordset
    
    Pkid = gstrItemData(cbo_intEvento)
    
    If Pkid = 0 Then
        habilitaGuias 3
        Exit Sub
    End If
    
    strSQL = ""
   
    strSQL = strSQL & "SELECT intTipoEvento FROM "
    strSQL = strSQL & gstrEvento
    strSQL = strSQL & " WHERE PKID = " & CStr(Pkid)
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            intTipoEventoSelecionado = adoResultado!intTipoEvento
         End If
      End With
   End If
   
    Select Case Val(intTipoEventoSelecionado)
        Case Is = 3
            habilitaGuias 0
        Case Is = 4
            habilitaGuias 1
        Case Is = 5
            habilitaGuias 2
        'Case Is = 11
        '    habilitaGuias 3
        Case Is = 11
            habilitaGuias 5
    End Select


End Sub


Private Function RetornaContaDespesa(ByVal strPKId As String) As String
    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset
    
    strSQL = "SELECT intContaContabil FROM "
    strSQL = strSQL & gstrDespesaExtraOrcamentaria
    strSQL = strSQL & " WHERE PKID = " & strPKId
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            RetornaContaDespesa = adoResultado!intContaContabil
         End If
      End With
   End If
   
End Function


Public Sub PreenchePlanoContaGeral(cboCodigo As ComboBox, _
                             cboDescricao As ComboBox)

    Dim strSQL              As String
    Dim strCondicao         As String
    Dim adoResultado        As ADODB.Recordset
    
    cboCodigo.Clear
    cboDescricao.Clear
    
    strSQL = ""
'    strSql = "SELECT * FROM ("
'    strSql = strSql & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao "
'    strSql = strSql & "FROM "
'    strSql = strSql & gstrPlanoConta & " PC "
'    strSql = strSql & "WHERE ABS(PC.blnFinanceira) = 1 AND blnAnalitica = 1 "
'
'    strSql = strSql & "UNION ALL "
'    strSql = strSql & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao "
'    strSql = strSql & "FROM "
'    strSql = strSql & gstrPlanoConta & " PC "
'    strSql = strSql & "WHERE blnextraorcamentaria = 1) TD "
'    strSql = strSql & "ORDER BY TD.strDescricao"
    
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 60, adoResultado) Then
'        With adoResultado
'            Do While .EOF = False
'                cboDescricao.AddItem !strDescricao
'                cboDescricao.ItemData(cboDescricao.NewIndex) = !Pkid
'                cboCodigo.AddItem gvntFormatacaoEspecifica(!strContaContabil)
'                cboCodigo.ItemData(cboCodigo.NewIndex) = !Pkid
'                .MoveNext
'            Loop
'        End With
'    End If
    
    strSQL = strSQL & "SELECT PC.PKId, "
    strSQL = strSQL & "CB.intNumeroConta, "
    strSQL = strSQL & "PC.strDescricao "
    strSQL = strSQL & "FROM " & gstrContaBancaria & " CB, "
    strSQL = strSQL & gstrPlanoConta & " PC "
    strSQL = strSQL & "WHERE CB.PKId = PC.intContaBancaria AND "
    strSQL = strSQL & "PC.blnAnalitica = 1 AND "
    strSQL = strSQL & "PC.bytDisponibilidadeDeCaixa = 1 "
    strSQL = strSQL & "ORDER BY CB.intNumeroConta"

    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
        With adoResultado
            Do While .EOF = False
                cboDescricao.AddItem !strDescricao
                cboDescricao.ItemData(cboDescricao.NewIndex) = !Pkid
                cboCodigo.AddItem !intNumeroConta
                cboCodigo.ItemData(cboCodigo.NewIndex) = !Pkid
                .MoveNext
            Loop
        End With
    End If

End Sub

Private Sub LeLiquidacaoExtra(ByVal intParcela As String, _
                              ByVal strEmpenho As String, _
                              ByVal strNumero As String)
    Dim objList         As Object
    Dim dblExtra        As Double
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    
    If Trim(intParcela) = "" Then Exit Sub
    
    strSQL = ""
    If mblnAlterando = True Then
        strSQL = strSQL & "SELECT PC.PKID CONTAPKID, PC.intExtraMaua, PC.strDescricao, "
        strSQL = strSQL & "LC.PKId, LC.dblValor, LC.intConta "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrPlanoConta & " PC, "
        strSQL = strSQL & gstrLancamentoContabil & " LC "
        strSQL = strSQL & "WHERE LC.intConta = PC.PKId "
        strSQL = strSQL & "AND LC.intProcesso = " & txtPKID.Text
        strSQL = strSQL & " AND LC.intParcela = " & intParcela & " "
        strSQL = strSQL & "ORDER BY PC.strContaContabil"
    Else
        strSQL = strSQL & "SELECT PC.PKID CONTAPKID, PC.intExtraMaua, PC.strDescricao, "
        strSQL = strSQL & "SL.PKId, SL.dblValor, SL.intConta "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrPlanoConta & " PC, "
        strSQL = strSQL & gstrSubempenhoLiquidado & " SL "
        strSQL = strSQL & "WHERE SL.intConta = PC.PKId "
        strSQL = strSQL & "AND SL.bytTipo = 1 "
        strSQL = strSQL & " AND SL.intParcela = " & intParcela & " "
        strSQL = strSQL & "ORDER BY PC.strContaContabil"
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set objList = lvw_Extra.ListItems.Add(, , _
                              strEmpenho)
                              
                objList.SubItems(1) = strNumero
                objList.SubItems(2) = gstrENulo(!intExtraMaua)
                objList.SubItems(3) = !strDescricao
                objList.SubItems(4) = gstrConvVrDoSql(IIf(!dblValor < 0, !dblValor * -1, !dblValor))
                objList.SubItems(5) = !intConta
                objList.SubItems(6) = !ContaPKID
                
                dblExtra = dblExtra + gstrConvVrDoSql(IIf(!dblValor < 0, !dblValor * -1, !dblValor))
                objList.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
    
    If lvw_Extra.ListItems.Count > 0 Then
        tab_3DPastaEmpenho.TabEnabled(4) = True
    Else
        tab_3DPastaEmpenho.TabEnabled(4) = False
    End If
    'lblExtra = gstrConvVrDoSql(dblExtra)
End Sub

Private Function gstrOrdemPagamentoItens(combo As DataCombo, _
                                         ByVal strTabela As String, _
                                         Optional semAspas As Boolean, _
                                         Optional objLvwEmpenho As ListView) As String

    Dim strSQL              As String
    Dim adoResultado        As ADODB.Recordset
    Dim strOrdemPagamentoEmpenho As String

    Dim strOP               As String
    Dim intInd              As Integer
        
    Set gobjBanco = New clsBanco
    
    If strTabela = gstrOrdemPagamentoDespesaExtra Then
        strSQL = ""
        strSQL = strSQL & "SELECT t.intDespesaExtraOrcamentaria intParcela FROM " & strTabela
        strSQL = strSQL & " T WHERE intOrdemPagamento in( " & CStr(gstrItemData(combo)) & ")"
    Else
        strSQL = ""
        strSQL = strSQL & "SELECT T.intParcela intParcela FROM " & strTabela
        strSQL = strSQL & " T WHERE intOrdemPagamento in( " & CStr(gstrItemData(combo)) & ")" '
    End If
    
    
'    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        While Not adoResultado.EOF
            strOrdemPagamentoEmpenho = strOrdemPagamentoEmpenho & adoResultado!intParcela & ","
            adoResultado.MoveNext
        Wend
        If strOrdemPagamentoEmpenho <> "" Then
            strOrdemPagamentoEmpenho = "'" & Mid(strOrdemPagamentoEmpenho, 1, Len(strOrdemPagamentoEmpenho) - 1) & "'"
        End If
    End If
    If semAspas Then
        strOrdemPagamentoEmpenho = Replace(strOrdemPagamentoEmpenho, "'", "")
    End If
    
    gstrOrdemPagamentoItens = strOrdemPagamentoEmpenho
    
End Function

Private Function gstrOrdemPagamentoItensArray(combo As DataCombo, _
                                         ByVal strTabela As String, _
                                         Optional semAspas As Boolean, _
                                         Optional objLvwEmpenho As ListView) As String

Dim strSQL              As String
Dim adoResultado        As ADODB.Recordset
Dim strOrdemPagamentoEmpenho As String

Dim strOP               As String
Dim intInd              As Integer
        
    Set gobjBanco = New clsBanco
    
    strOP = Space$(0)
    
    For intInd = 1 To objLvwEmpenho.ListItems.Count
    
        strSQL = "SELECT PKId FROM " & gstrOrdemPagamento & " WHERE intNumero = " & objLvwEmpenho.ListItems(intInd).SubItems(1)
        If UCase(objLvwEmpenho.Name) = "LVW_RESTO" Then
            strSQL = strSQL & " AND intExercicio = " & objLvwEmpenho.ListItems(intInd).SubItems(9)
        ElseIf UCase(objLvwEmpenho.Name) = "LVW_EMPENHO" Then
            strSQL = strSQL & " AND intExercicio = " & gintExercicio
        Else
            strSQL = strSQL & " AND intExercicio = " & objLvwEmpenho.ListItems(intInd).SubItems(6)
        End If
        
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                strOP = strOP & adoResultado("PKId") & ","
            End If
        End If
        
    Next
    
    If Len(Trim$(strOP)) > 0 Then
        strOP = Left$(strOP, Len(strOP) - 1)
    Else
        strOP = CStr(gstrItemData(combo))
    End If
    
    If strTabela = gstrOrdemPagamentoDespesaExtra Then
        strSQL = ""
        strSQL = strSQL & "SELECT t.intDespesaExtraOrcamentaria intParcela FROM " & strTabela
        strSQL = strSQL & " T WHERE intOrdemPagamento in( " & strOP & ")" 'CStr(gstrItemData(combo))
    Else
        strSQL = ""
        strSQL = strSQL & "SELECT T.intParcela intParcela FROM " & strTabela
        strSQL = strSQL & " T WHERE intOrdemPagamento in( " & strOP & ")" 'CStr(gstrItemData(combo))
    End If
    
    
'    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        While Not adoResultado.EOF
            strOrdemPagamentoEmpenho = strOrdemPagamentoEmpenho & adoResultado!intParcela & ","
            adoResultado.MoveNext
        Wend
        If strOrdemPagamentoEmpenho <> "" Then
            strOrdemPagamentoEmpenho = "'" & Mid(strOrdemPagamentoEmpenho, 1, Len(strOrdemPagamentoEmpenho) - 1) & "'"
        End If
    End If
    If semAspas Then
        strOrdemPagamentoEmpenho = Replace(strOrdemPagamentoEmpenho, "'", "")
    End If
    
    If Len(Trim(strOrdemPagamentoEmpenho)) = 0 Then strOrdemPagamentoEmpenho = "0"
    
    gstrOrdemPagamentoItensArray = strOrdemPagamentoEmpenho
    
End Function


Private Function gstrEmpenhobySubEmpenho(ByVal strSubEmpenhoPKIDs As String) As String

    Dim strSQL              As String
    Dim adoResultado        As ADODB.Recordset

    strSQL = ""
    strSQL = strSQL & "SELECT E.PKID FROM "
    strSQL = strSQL & gstrEmpenho & " E," & gstrSubempenho & " SE "
    strSQL = strSQL & " WHERE SE.PKID IN (" & strSubEmpenhoPKIDs & ")"
    strSQL = strSQL & " AND SE.intEmpenho = E.PKID "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        While Not adoResultado.EOF
            gstrEmpenhobySubEmpenho = gstrEmpenhobySubEmpenho & adoResultado!Pkid & ","
            adoResultado.MoveNext
        Wend
    End If
    
    If gstrEmpenhobySubEmpenho <> "" Then
        gstrEmpenhobySubEmpenho = Mid(gstrEmpenhobySubEmpenho, 1, Len(gstrEmpenhobySubEmpenho) - 1)
    End If
End Function

Private Sub ProcuraParcelaOrdemResto()
    Dim i As Integer
    For i = 1 To lvw_Resto.ListItems.Count
        If lvw_Resto.ListItems(i).SubItems(1) = dcbOrdemPagamentoResto Then
            mblnAlterandoResto = True
            Exit For
        Else
            mblnAlterandoResto = False
        End If
    Next
End Sub


Private Sub preencheLiquidacaoExtra(list As ListView)
    Dim i As Integer
    txtTotalReceitaExtra = "0,00"
    lvw_Extra.ListItems.Clear
    
    For i = 1 To list.ListItems.Count
        If UCase(list.Name) = "LVW_RESTO" Then
            LeLiquidacaoExtra list.ListItems(i).Tag, list.ListItems(i).SubItems(3), list.ListItems(i).SubItems(4)
        ElseIf UCase(list.Name) = "LVW_EMPENHO" Then
            LeLiquidacaoExtra list.ListItems(i).Tag, list.ListItems(i).SubItems(2), list.ListItems(i).SubItems(3)
        End If
    Next
    
    For i = 1 To lvw_Extra.ListItems.Count
        txtTotalReceitaExtra = gstrConvVrDoSql(Val(gstrConvVrParaSql(txtTotalReceitaExtra)) + Val(gstrConvVrParaSql(lvw_Extra.ListItems(i).SubItems(4))))
    Next
    
    txtTotalReceitaExtra = gstrConvVrDoSql(IIf(chkEstorno.Value = 1, Val(gstrConvVrParaSql(txtTotalReceitaExtra)) * -1, Val(gstrConvVrParaSql(txtTotalReceitaExtra))))
    
    SomaTotalAPagar
    AjustaFormatacaoReceitaExtra
End Sub

Private Function gstrOrdemPagamentonoGrid(Optional intLinha As Integer = 0) As String
    Dim lvw_list       As ListView
    Dim i              As Integer
    Dim strOrdemNumero As String
    Dim strSQL              As String
    Dim adoResultado        As ADODB.Recordset
    Dim bytColunaExercicio  As Byte
    
    If tab_3DPastaEmpenho.TabEnabled(0) = True Then
        Set lvw_list = lvw_Empenho
        bytColunaExercicio = 10
    ElseIf tab_3DPastaEmpenho.TabEnabled(1) = True Then
        Set lvw_list = lvw_Resto
        bytColunaExercicio = 9
    ElseIf tab_3DPastaEmpenho.TabEnabled(2) = True Then
        Set lvw_list = lvw_Despesa
        bytColunaExercicio = 6
    ElseIf tab_3DPastaEmpenho.TabEnabled(5) = True Then
        Set lvw_list = lvw_AnulacaoReceita
        bytColunaExercicio = 5
    End If
        
    strOrdemNumero = "("
    
    If intLinha = 0 Then
        For i = 1 To lvw_list.ListItems.Count
            If InStr(1, strOrdemNumero, "intNumero = " & lvw_list.ListItems(i).SubItems(1)) = 0 And lvw_list.ListItems(i).SubItems(1) <> "--" Then
                strOrdemNumero = strOrdemNumero & " intNumero = " & lvw_list.ListItems(i).SubItems(1) & " AND intExercicio = " & lvw_list.ListItems(i).SubItems(bytColunaExercicio) & ") OR("
            End If
        Next
    Else
        If lvw_list.ListItems(intLinha).SubItems(1) <> "--" Then
            strOrdemNumero = strOrdemNumero & " intNumero = " & lvw_list.ListItems(intLinha).SubItems(1) & " AND intExercicio = " & lvw_list.ListItems(intLinha).SubItems(bytColunaExercicio) & ") OR("
        End If
    End If
    
    If Len(strOrdemNumero) > 1 Then
        strOrdemNumero = Mid(strOrdemNumero, 1, Len(strOrdemNumero) - 3)
    Else
        strOrdemNumero = "0"
        gstrOrdemPagamentonoGrid = strOrdemNumero
        Exit Function
    End If
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKID FROM " & gstrOrdemPagamento
    strSQL = strSQL & " WHERE (" & strOrdemNumero & ")"
    
    strOrdemNumero = ""
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        While Not adoResultado.EOF
            strOrdemNumero = strOrdemNumero & CStr(adoResultado!Pkid) & ","
            adoResultado.MoveNext
        Wend
    End If
    
    If Len(strOrdemNumero) <> 0 Then
        strOrdemNumero = Mid(strOrdemNumero, 1, Len(strOrdemNumero) - 1)
    Else
        strOrdemNumero = "0"
    End If
    gstrOrdemPagamentonoGrid = strOrdemNumero
    
End Function

Private Function gstrConta(ByVal strPKId As String, blnEmpenho As Boolean) As String
    Dim strSQL              As String
    Dim adoResultado        As ADODB.Recordset
    Dim strCampoRetorno        As String
    
    strCampoRetorno = IIf(blnEmpenho, "intParcela", "intConta")
    
    strSQL = ""
    strSQL = strSQL & "SELECT " & strCampoRetorno & " FROM " & gstrSubempenhoLiquidado
    strSQL = strSQL & " WHERE PKID = " & strPKId

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        While Not adoResultado.EOF
            gstrConta = CStr(adoResultado(strCampoRetorno))
            adoResultado.MoveNext
        Wend
    End If
End Function

Private Sub FiltraCampos(tdb_Grid As TDBGrid)
'Procedimento para filtar os campos com a opção F5 por Nino

Dim adoGrid     As ADODB.Recordset
Dim strCondicao As String

    Set adoGrid = New ADODB.Recordset
    Set adoGrid = tdb_Grid.DataSource
    
    
    If adoGrid.BOF And adoGrid.EOF Then
        Exit Sub
    End If
    
    
    If gblnDataValida(txtData) Then
        strCondicao = strCondicao & " dtmData LIKE #" & txtData.Text & "#"
    End If
    
    If cbo_intEvento.ListIndex <> -1 Then
        If strCondicao <> "" Then
            strCondicao = strCondicao & " AND "
        End If
        strCondicao = strCondicao & "intEvento = " & gstrItemData(cbo_intEvento)
    End If
    
    
    If strCondicao <> "" Then
        adoGrid.Filter = strCondicao
        If adoGrid.EOF And adoGrid.BOF Then
            adoGrid.Filter = adFilterNone
        End If
    
    Else
        adoGrid.Filter = adFilterNone
    End If
        
        
    'SetArrayforGrid adoGrid
    tdb_Grid.DataSource = adoGrid
    
    tdb_Grid.ReBind
    tdb_Grid.Refresh

End Sub

Private Function blnVerificaGrid() As Boolean
    Dim intCont As Integer
    blnVerificaGrid = False
    With lvw_Conta
        If .ListItems.Count >= 1 Then
            For intCont = 1 To .ListItems.Count
                If mblnAlterandoConta Then
                    If cbointContaContabil.Text = .ListItems(intCont).Text And Trim(txtNumCheque) = .ListItems(intCont).SubItems(3) And .ListItems(intCont).Selected = False Then
                        blnVerificaGrid = True
                    End If
                Else
                    If cbointContaContabil.Text = .ListItems(intCont).Text And Trim(txtNumCheque) = .ListItems(intCont).SubItems(3) Then
                        blnVerificaGrid = True
                    End If
                End If
            Next
        End If
    End With
End Function

Private Function strDataEncerramento() As String
    Dim adoResultado As ADODB.Recordset
    Dim strSQL       As String
    strSQL = "SELECT"
    strSQL = strSQL & " dtmFechamento"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrFechamentoContabil
    strSQL = strSQL & " WHERE strCodigo = 'EF'"
    strSQL = strSQL & " AND intExercicio = " & gintExercicio
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
          strDataEncerramento = gstrDataFormatada(adoResultado!dtmFechamento)
    End If
    Set gobjBanco = Nothing
End Function

Private Function strValorAtualdaConta(Optional blnLista As Boolean) As String
    Dim intContador     As Integer
    Dim strConta        As String
    Dim strSaldo        As Double
    Dim strdtmData      As String
    Dim strSaldodoBanco As String
    
    If txtData = "" Then
        strdtmData = gstrDataDoSistema
    Else
        strdtmData = txtData
    End If
    
    strConta = cbointContaContabil
    strSaldo = SaldoContaContabilAtual(gstrItemData(cbointContaContabil), Month(CDate(strdtmData)), gintExercicio, , txtData)
    strSaldodoBanco = strSaldo
          
        If mblnAlterando = True Then
            strValorAtualdaConta = gstrConvVrDoSql(strSaldo)
            txt_saldoAtual = gstrConvVrDoSql(strSaldo)
            Exit Function
        End If
            
        If blnLista Then
            If lvw_Conta.ListItems.Count >= 1 Then
                With lvw_Conta
                     For intContador = 1 To .ListItems.Count
                        If strConta = .ListItems(intContador).Text Then
                            If chkEstorno.Value = 1 Then
                                strSaldo = CDbl(strSaldo) + CDbl(.ListItems(intContador).SubItems(2))
                            Else
                                strSaldo = CDbl(strSaldo) - CDbl(.ListItems(intContador).SubItems(2))
                            End If
                        End If
                    Next
                End With
            End If
            strValorAtualdaConta = gstrConvVrDoSql(strSaldo)
            txt_saldoAtual = gstrConvVrDoSql(strSaldo)
        Else
                If lvw_Conta.ListItems.Count >= 1 Then
                    With lvw_Conta
                         For intContador = 1 To .ListItems.Count
                            If strConta = .ListItems(intContador).Text Then
                                If chkEstorno.Value = 1 Then
                                    strSaldo = strSaldo + CDbl(.ListItems(intContador).SubItems(2))
                                Else
                                    strSaldo = strSaldo - CDbl(.ListItems(intContador).SubItems(2))
                                End If
                            End If
                        Next
                    End With
                End If
                strSaldo = Val(gstrConvVrParaSql(strSaldo)) - Val(gstrConvVrParaSql(CDbl(txtValorLancamento)))
                If CDbl(strSaldo) < 0 Then
                
                    If MsgBox("O saldo disponível para esta movimentação é insuficiente." & vbNewLine & "Você Deseja realizar a movimentação mesmo assim?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                        strValorAtualdaConta = "0"
                    Else
                        'ExibeMensagem "Não existe saldo para esta Movimentação"
                        strValorAtualdaConta = "-1"
                    End If

                    Exit Function
                End If
                strValorAtualdaConta = strSaldo
                txt_saldoAtual = gstrConvVrDoSql(strSaldo)
        End If
    
End Function

Private Function blnValorAtualdaContaCheque(ByVal strPkidConta As String, ByVal strNumConta As String, ByVal strValorCheque As String) As Boolean
    Dim intContador     As Integer
    Dim strConta        As String
    Dim dblSaldo        As Double
    Dim strdtmData      As String
    Dim strSaldodoBanco As String
    Dim dlbPkidConta    As Double
    
    If txtData = "" Then
        strdtmData = gstrDataDoSistema
    Else
        strdtmData = txtData
    End If
    
    dlbPkidConta = CDbl(strPkidConta)
    strConta = strNumConta
    dblSaldo = SaldoContaContabilAtual(dlbPkidConta, Month(CDate(strdtmData)), gintExercicio, , txtData)
    strSaldodoBanco = dblSaldo
          
    If lvw_Conta.ListItems.Count >= 1 Then
        With lvw_Conta
             For intContador = 1 To .ListItems.Count
                If strConta = .ListItems(intContador).Text Then
                    dblSaldo = dblSaldo - Val(gstrConvVrParaSql(.ListItems(intContador).SubItems(2)))
                End If
            Next
        End With
    End If
    dblSaldo = dblSaldo - Val(gstrConvVrParaSql(strValorCheque))
    
    If dblSaldo < 0 Then
    
        If MsgBox("O saldo disponível para esta movimentação é insuficiente." & vbNewLine & "Você Deseja realizar a movimentação mesmo assim?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            blnValorAtualdaContaCheque = True
        Else
            'ExibeMensagem "Não existe saldo para esta Movimentação"
            blnValorAtualdaContaCheque = False
        End If

        Exit Function
    End If
    blnValorAtualdaContaCheque = True
    
End Function

Private Sub MontaTotalRecOrdem(blnAdiciona As Boolean)
Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
    If tab_3DPastaEmpenho.Tab = 0 Then
      strSQL = "SELECT "
      strSQL = strSQL & gstrISNULL("SUM(SE.dblDesconto)", "0") & " dblDesconto"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE, "
      strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OPE, "
      strSQL = strSQL & gstrOrdemPagamento & " OP"
      strSQL = strSQL & " WHERE "
      'strSql = strSql & " OPE.intOrdemPagamento = " & Val(lngPkidOrdem) & " AND"
      strSQL = strSQL & "SE.Pkid IN (" & gstrOrdemPagamentoItensArray(dcbOrdemPagamentoEmpenho, gstrOrdemPagamentoEmpenho, True, lvw_Empenho) & ") AND "
      strSQL = strSQL & " SE.Pkid = OPE.intParcela AND "
      strSQL = strSQL & " OPE.intOrdemPagamento = OP.PKId AND "
      strSQL = strSQL & " (OP.bytCancelado IS NULL OR OP.bytCancelado = 0)"
      
    End If
    If tab_3DPastaEmpenho.Tab = 1 Then
      strSQL = "SELECT "
      strSQL = strSQL & gstrISNULL("SUM(SE.dblDesconto)", "0") & " dblDesconto"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE, "
      strSQL = strSQL & gstrOrdemPagamentoResto & " OPR, "
      strSQL = strSQL & gstrOrdemPagamento & " OP"
      strSQL = strSQL & " WHERE "
      'strSql = strSql & " OPR.intOrdemPagamento = " & Val(lngPkidOrdem) & " AND"
      strSQL = strSQL & "SE.Pkid IN (" & gstrOrdemPagamentoItensArray(dcbOrdemPagamentoResto, gstrOrdemPagamentoResto, True, lvw_Resto) & ") AND "
      strSQL = strSQL & " SE.Pkid = OPR.intParcela AND "
      strSQL = strSQL & " OPR.intOrdemPagamento = OP.PKId AND "
      strSQL = strSQL & " (OP.bytCancelado IS NULL OR OP.bytCancelado = 0)"
      
    End If
    If tab_3DPastaEmpenho.Tab = 2 Then
      strSQL = "SELECT "
      strSQL = strSQL & gstrISNULL("SUM(DEX.dblDesconto)", "0") & " dblDesconto"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DEX, "
      strSQL = strSQL & gstrOrdemPagamentoDespesaExtra & " OPEX, "
      strSQL = strSQL & gstrOrdemPagamento & " OP"
      strSQL = strSQL & " WHERE "
      'strSql = strSql & " OPEX.intOrdemPagamento = " & Val(lngPkidOrdem) & " AND"
      strSQL = strSQL & " DEX.Pkid IN (" & gstrOrdemPagamentoItensArray(dcbOrdemPagamentoDespesa, gstrOrdemPagamentoDespesaExtra, True, lvw_Despesa) & ") AND "
      strSQL = strSQL & " DEX.Pkid = OPEX.intDespesaExtraOrcamentaria AND"
      strSQL = strSQL & " OPEX.intOrdemPagamento = OP.PKId AND"
      strSQL = strSQL & " (OP.bytCancelado IS NULL OR OP.bytCancelado = 0)"
    
    End If
    
    txtTotalRecOrdem = gstrConvVrDoSql("0.00")
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            While Not adoResultado.EOF
               If blnAdiciona Then
                   txtTotalRecOrdem = gstrConvVrDoSql(CDbl(txtTotalRecOrdem) + adoResultado!dblDesconto, 2)
               Else
                   txtTotalRecOrdem = gstrConvVrDoSql(CDbl(txtTotalRecOrdem) - adoResultado!dblDesconto, 2)
               End If
               adoResultado.MoveNext
            Wend
            
            txtTotalRecOrdem = gstrConvVrDoSql(IIf(chkEstorno.Value = 1, CDbl(txtTotalRecOrdem) * -1, CDbl(txtTotalRecOrdem)))
        End If
    End If

End Sub
Public Function blnFitaAutenticadoraOK() As Boolean
'****************************************************************************************
' Create by:         Éder Henrique
' Módulos:           Orçamentário
' Data:              02/05/2006
' Ficha:             orc1340
' Comentários:       Verifica se é necessário imprimir a fita
'                    Autenticadora
'****************************************************************************************

    Dim strSQL                          As String
    Dim adoResult                       As New ADODB.Recordset

    Err.Clear

    On Local Error GoTo ERRO_blnFitaAutenticadoraOK

    Set gobjBanco = New clsBanco
    
    'Pego o parametro para autenticação
    strSQL = "SELECT bitUsaRotinaAutenticacao "
    strSQL = strSQL & "FROM " & gstrConfiguracaoGeral
    
    If gobjBanco.CriaADO(strSQL, 5, adoResult) Then
        If Not adoResult.EOF Then
            If gstrENulo(adoResult.Fields("bitUsaRotinaAutenticacao"), False, False) = 1 Then
                blnFitaAutenticadoraOK = True 'Sim
            Else
                blnFitaAutenticadoraOK = False 'Não
            End If
        End If
        adoResult.Close
    End If

ERRO_blnFitaAutenticadoraOK:
    If Err.Number <> 0 Then
        ExibeMensagem "Ocorreu o erro: " + Str(Err) + vbCrLf + Err.Description + vbCrLf + "Em: Função blnImprimeFitaAutenticadora "
        Err.Clear
        blnFitaAutenticadoraOK = False
    End If
End Function
Public Sub ImprimeFitaAutenticadora(strLinha_A_Imprimir As String)
'********************************************************************************************************************************************************************************
' Create by:         Éder Henrique
' Módulos:           Orçamentário
' Data:              02/05/2006
' Ficha:             orc1340
' Comentários:       Imprimi a fita Autenticadora
'********************************************************************************************************************************************************************************
    
Dim intLoop         As Integer
    
    On Local Error GoTo ERRO_ImprimeFitaAutenticadora
    
    intLoop = 1
    Do While intLoop <= 2
        If MsgBox("Autenticar Movimento " & strLinha_A_Imprimir & " ?", vbYesNo + vbInformation, "Orçamentário") = vbYes Then
            'Abre a Porta Paralela
            intArquivo = FreeFile
            Open "LPT1:" For Output As #intArquivo
                'Imprime a fita Autenticadora
                Print #intArquivo, Chr(15) + strLinha_A_Imprimir + Chr(18)
                'O comando acima  Chr(15) + "Conteudo a ser impresso" + Chr(18) comprime a impressão
            'Fecha a Porta Paralela
            Close #intArquivo
        End If
        intLoop = intLoop + 1
    Loop
    Exit Sub
    
ERRO_ImprimeFitaAutenticadora:
    If Err.Number <> 0 Then
        If Err.Number = 76 Then
            ExibeMensagem "Impressora não instalada corretamente ! "
        Else
            ExibeMensagem "Ocorreu o erro: " + Str(Err) + vbCrLf + Err.Description + vbCrLf + "Em: Função ImprimeFitaAutenticadora"
        End If
        Err.Clear
    End If
End Sub

Private Function strRegistroAutenticacao(intNumeroGuia As String) As String
'********************************************************************************************************************************************************************************
' Create by:         Éder Henrique
' Módulos:           Orçamentário
' Data:              02/05/2006
' Ficha:             orc1340
' Comentários:       Concatena o registro a ser impresso
'********************************************************************************************************************************************************************************
    Dim strSQL                          As String
    Dim strAux                          As String
    Dim adoResult                       As New ADODB.Recordset

    Err.Clear

    On Local Error GoTo ERRO_strRegistroAutenticacao

    Set gobjBanco = New clsBanco
    
    'Consulto autenticação
    strSQL = "SELECT AR.intNumero, AR.dtmData, CB.intNumeroConta" & _
           "  FROM tblPlanoConta PC, tblContaBancaria CB, tblarrecadacaoreceita AR" & _
           "  WHERE CB.PKId = PC.intContaBancaria" & _
           "  AND AR.intContaContabil = PC.Pkid" & _
           "  AND PC.Pkid = " & gstrItemData(cbointContaContabil) & _
           "  AND AR.intNumero = " & intNumeroGuia
    
    If gobjBanco.CriaADO(strSQL, 5, adoResult) Then
        If Not adoResult.EOF Then
             strAux = Right(String(6, "0") & adoResult.Fields("intNumero"), 6) & " " & Format(adoResult.Fields("dtmData"), "dd/mm/yyyy") & " " & Right(String(6, "0") & adoResult.Fields("intNumeroConta"), 6) & " "
        End If
        adoResult.Close
    End If
    
    strRegistroAutenticacao = strAux
    
ERRO_strRegistroAutenticacao:
    If Err.Number <> 0 Then
        ExibeMensagem "Ocorreu o erro: " + Str(Err) + vbCrLf + Err.Description + vbCrLf + "Em: Função strRegistroAutenticacao "
        Err.Clear
        strRegistroAutenticacao = ""
    End If
End Function
Private Function PegaCredor()
Dim strSQL      As String
Dim adoResult   As New ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT  CT.CDC, "
    strSQL = strSQL & "    CT.strNome "
    strSQL = strSQL & "FROM    tblOrdemPagamento   OP, "
    strSQL = strSQL & "    tblContribuinte     CT "
    strSQL = strSQL & "WHERE   OP.intnumero = " & dcbOrdemPagamentoEmpenho.Text & " AND "
    strSQL = strSQL & " OP.intexercicio = " & gintExercicio & " AND "
    strSQL = strSQL & "    OP.intContribuinte = CT.pkid"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResult) Then
        If Not adoResult.EOF Then
            txt_Cdc.Text = Space$(0) & adoResult.Fields("CDC")
            txt_strNome.Text = Space$(0) & adoResult.Fields("strNome")
        Else
            txt_Cdc.Text = Space$(0)
            txt_strNome.Text = Space$(0)
        End If
    End If
    
End Function
Sub LocalDesabilitaNumeroParcela()
'/*****************************************************************************************
' Programador:   Éder Henrique
' Módulos:       Orçamentário
' Data:          23/08/2006
' Ficha:         orc1562
' Objetivo:      Desabilitar as combos numero e parcelas das abas
'                empenho, restos a pagar e depesa extra.
'******************************************************************************************/
    TrocaCorObjeto dcbEmpenho, True, False
    TrocaCorObjeto dcbParcela, True, False
    TrocaCorObjeto dcbResto, True, False
    TrocaCorObjeto dcbParcelaResto, True, False
    TrocaCorObjeto dcbDespesa, True, False
End Sub

Private Function gstrProximoCheque(intIDContaBancaria As Integer, intContaBancaria As Integer) As String
'/*****************************************************************************************
' Programador:   Éder Henrique
' Módulos:       Orçamentário
' Data:          28/08/2006
' Ficha:         orc1585
' Objetivo:      Sugere o próximo numero de cheque de acordo com a conta
'                atraves do parametro intContaBancaria
'******************************************************************************************/

Dim strSQL                      As String
Dim strNumeroCheque             As String
Dim i                           As Integer
Dim lngDifCheque                As Long
Dim adoResultado                As New ADODB.Recordset

    'Vamos verificar no listView se já foi inserido uma conta
    'com algum cheque
    If lvw_Conta.ListItems.Count > 0 Then
        For i = 1 To lvw_Conta.ListItems.Count
            'Se for a mesma conta
            If Trim(lvw_Conta.ListItems.Item(i).Text) = Trim(intContaBancaria) Then
                'Se for Maior que o cheque anterior
                If Val(lvw_Conta.ListItems.Item(i).ListSubItems(3)) > Val(strNumeroCheque) Then
                    strNumeroCheque = lvw_Conta.ListItems.Item(i).ListSubItems(3)
                End If
            End If
        Next
    End If
    
    'Vamos verificar o maior numero de cheques cadastrados
    strSQL = "SELECT MAX(LC.strDocumento) Cheque "
    strSQL = strSQL & "FROM " & gstrContaBancaria & " CB, " & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrProcessoPagamento & " PP, "
    strSQL = strSQL & gstrLancamentoContabil & " LC "
    strSQL = strSQL & "WHERE CB.Pkid = PC.intContaBancaria "
    strSQL = strSQL & "AND PC.PKId = LC.intConta "
    strSQL = strSQL & "AND LC.intProcesso = PP.PKId "
    strSQL = strSQL & "AND PC.blnAnalitica = 1 "
    strSQL = strSQL & "AND PC.bytDisponibilidadeDeCaixa = 1 "
    strSQL = strSQL & "AND " & gstrDATEPART("YYYY", "PP.dtmdata") & " = " & gintExercicio
    strSQL = strSQL & "AND CB.intNumeroConta = " & intContaBancaria
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If Val(strNumeroCheque) > 0 Then
                'Se não for o primeiro Lancamento
                lngDifCheque = (Val(strNumeroCheque) - Val(gstrENulo(adoResultado.Fields("Cheque").Value))) + 1
            Else
                'É o primeiro lancamento
                lngDifCheque = 1
            End If
            gstrProximoCheque = CStr(Val(gstrENulo(adoResultado.Fields("Cheque").Value)) + Abs(lngDifCheque))
        End If
        adoResultado.Close
    End If
End Function
Private Function blnIncremtCheque() As Boolean
'/*****************************************************************************************
' Programador:   Éder Henrique
' Módulos:       Orçamentário
' Data:          28/08/2006
' Ficha:         orc1585
' Objetivo:      Verifica se o parametro bitNumeroChequeAutomatico libera
'                Numeração de cheque automatica
'******************************************************************************************/

Dim strSQL                      As String
Dim adoResultado                As New ADODB.Recordset


    blnIncremtCheque = False
    
    strSQL = "SELECT bitNumeroChequeAutomatico "
    strSQL = strSQL & "FROM " & gstrConfiguracaoGeral
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If Val(gstrENulo(adoResultado.Fields("bitNumeroChequeAutomatico"))) = 1 Then
                blnIncremtCheque = True
            End If
        End If
        adoResultado.Close
    End If

End Function
