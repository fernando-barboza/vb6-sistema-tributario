VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmMovimentoBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimento Bancário"
   ClientHeight    =   6045
   ClientLeft      =   2085
   ClientTop       =   3855
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9915
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5925
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   10451
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Movimento Bancário"
      TabPicture(0)   =   "frmMovimentoBancario.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAviso"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDataDoMovimento"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLote"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblContaBancaria"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbldblPrincipal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDataDoPagamento"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTributo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbldblMulta"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbldblJuros"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbldblCorrecao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblTotal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblCodigoDeBarras"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblCorreto"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblExercicio"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblcodigoBaixa"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblAgencia"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblBanco"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl_bytTipo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl_QtdeDocs"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "sha_Lote"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "dbc_strComposicaoDaReceita"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dbc_strDescricaoConta"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "dbc_intComposicaoDaReceita"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "dbcintlancamentovalor"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "dbc_intNumeroAviso"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "dbcintcodigobaixa"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "tdb_MovimentoBancario"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "dbcintContaBancaria"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtdblPrincipal"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtintLote"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtdtmDtMovimento"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmd_ContaCorrente"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtdtmDtPagamento"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmd_Composicao"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtdblMulta"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtdblJuros"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtdblCorrecao"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txt_dblTotal"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txt_strCodigoDeBarras"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtdblCorreto"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtintDigito"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt_intExercicio"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt_strBanco"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txt_strAgencia"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtPKId"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtbytTipo"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt_QtdeDocs"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).ControlCount=   47
      Begin VB.TextBox txt_QtdeDocs 
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
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   4
         Top             =   495
         Width           =   480
      End
      Begin VB.TextBox txtbytTipo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   4830
         MaxLength       =   10
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_strAgencia 
         Height          =   315
         Left            =   1665
         TabIndex        =   12
         Top             =   1215
         Width           =   2415
      End
      Begin VB.TextBox txt_strBanco 
         Height          =   315
         Left            =   4695
         TabIndex        =   14
         Top             =   1215
         Width           =   4935
      End
      Begin VB.TextBox txt_intExercicio 
         Alignment       =   1  'Right Justify
         DataField       =   ","
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1665
         MaxLength       =   4
         TabIndex        =   22
         Top             =   2055
         Width           =   570
      End
      Begin VB.TextBox txtintDigito 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   5250
         MaxLength       =   1
         TabIndex        =   26
         Top             =   2055
         Width           =   480
      End
      Begin VB.TextBox txtdblCorreto 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   42
         Top             =   2790
         Width           =   1230
      End
      Begin VB.TextBox txt_strCodigoDeBarras 
         Height          =   285
         Left            =   1665
         MaxLength       =   44
         TabIndex        =   40
         Top             =   2805
         Width           =   4095
      End
      Begin VB.TextBox txt_dblTotal 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   8400
         TabIndex        =   38
         Top             =   2430
         Width           =   1230
      End
      Begin VB.TextBox txtdblCorrecao 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6645
         TabIndex        =   36
         Top             =   2460
         Width           =   1230
      End
      Begin VB.TextBox txtdblJuros 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4950
         TabIndex        =   34
         Top             =   2445
         Width           =   975
      End
      Begin VB.TextBox txtdblMulta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   32
         Top             =   2445
         Width           =   975
      End
      Begin VB.CommandButton cmd_Composicao 
         Height          =   300
         Left            =   9270
         Picture         =   "frmMovimentoBancario.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Composição da Receita"
         Top             =   1680
         Width           =   360
      End
      Begin VB.TextBox txtdtmDtPagamento 
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
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1680
         Width           =   1125
      End
      Begin VB.CommandButton cmd_ContaCorrente 
         Height          =   315
         Left            =   8295
         Picture         =   "frmMovimentoBancario.frx":013A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "585"
         ToolTipText     =   "Ativa Cadastro de Conta Bancária"
         Top             =   840
         Width           =   360
      End
      Begin VB.TextBox txtdtmDtMovimento 
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
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   2
         Top             =   495
         Width           =   1125
      End
      Begin VB.TextBox txtintLote 
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
         Left            =   9150
         MaxLength       =   4
         TabIndex        =   10
         Top             =   870
         Width           =   480
      End
      Begin VB.TextBox txtdblPrincipal 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1665
         TabIndex        =   30
         Top             =   2445
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo dbcintContaBancaria 
         Height          =   315
         HelpContextID   =   1
         Left            =   1665
         TabIndex        =   6
         Top             =   840
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   -2147483643
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_MovimentoBancario 
         Height          =   2565
         Left            =   180
         TabIndex        =   43
         Top             =   3210
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   4524
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
         Columns(1).Caption=   "Movimento"
         Columns(1).DataField=   "Movimento"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Conta Corrente"
         Columns(2).DataField=   "ContaCorrente"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Lote"
         Columns(3).DataField=   "Lote"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Pagamento"
         Columns(4).DataField=   "Pagamento"
         Columns(4).NumberFormat=   "General Date"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Composição da Receita"
         Columns(5).DataField=   "Tributo"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Exercício"
         Columns(6).DataField=   "intExercicio"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Aviso"
         Columns(7).DataField=   "STRNUMEROAVISO"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Parcela"
         Columns(8).DataField=   "Intparcela"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Dígito"
         Columns(9).DataField=   "intDigito"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Tipo de Baixa"
         Columns(10).DataField=   "intCodigoBaixa"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Principal"
         Columns(11).DataField=   "Principal"
         Columns(11).NumberFormat=   "Standard"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Multa"
         Columns(12).DataField=   "Multa"
         Columns(12).NumberFormat=   "Standard"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Juros"
         Columns(13).DataField=   "Juros"
         Columns(13).NumberFormat=   "Standard"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Correção"
         Columns(14).DataField=   "Correcao"
         Columns(14).NumberFormat=   "Standard"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Total"
         Columns(15).DataField=   "Total"
         Columns(15).NumberFormat=   "Standard"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Tipo"
         Columns(16).DataField=   "Tipo"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Código de Barras"
         Columns(17).DataField=   "CodigoDeBarras"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "OK"
         Columns(18).DataField=   "bitGuia"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "Processado"
         Columns(19).DataField=   "bitProcessado"
         Columns(19).DefaultValue=   "0"
         Columns(19).DefaultValue.vt=   8
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   20
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=20"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1588"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1508"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3625"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3545"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=1138"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1058"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=1720"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1640"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=4921"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=4842"
         Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=1429"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1349"
         Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(40)=   "Column(7).Width=2223"
         Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2143"
         Splits(0)._ColumnProps(43)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(44)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(46)=   "Column(8).Width=1667"
         Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=1588"
         Splits(0)._ColumnProps(49)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(51)=   "Column(9).Width=979"
         Splits(0)._ColumnProps(52)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(9)._WidthInPix=900"
         Splits(0)._ColumnProps(54)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(55)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(56)=   "Column(10).Width=2937"
         Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=2858"
         Splits(0)._ColumnProps(59)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(61)=   "Column(11).Width=1879"
         Splits(0)._ColumnProps(62)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(11)._WidthInPix=1799"
         Splits(0)._ColumnProps(64)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(11)._ColStyle=2"
         Splits(0)._ColumnProps(66)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(67)=   "Column(12).Width=1958"
         Splits(0)._ColumnProps(68)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(12)._WidthInPix=1879"
         Splits(0)._ColumnProps(70)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(12)._ColStyle=2"
         Splits(0)._ColumnProps(72)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(73)=   "Column(13).Width=1799"
         Splits(0)._ColumnProps(74)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(13)._WidthInPix=1720"
         Splits(0)._ColumnProps(76)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(77)=   "Column(13)._ColStyle=2"
         Splits(0)._ColumnProps(78)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(79)=   "Column(14).Width=1879"
         Splits(0)._ColumnProps(80)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(14)._WidthInPix=1799"
         Splits(0)._ColumnProps(82)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(83)=   "Column(14)._ColStyle=2"
         Splits(0)._ColumnProps(84)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(85)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(86)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(88)=   "Column(15)._EditAlways=0"
         Splits(0)._ColumnProps(89)=   "Column(15)._ColStyle=2"
         Splits(0)._ColumnProps(90)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(91)=   "Column(16).Width=714"
         Splits(0)._ColumnProps(92)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(93)=   "Column(16)._WidthInPix=635"
         Splits(0)._ColumnProps(94)=   "Column(16)._EditAlways=0"
         Splits(0)._ColumnProps(95)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(96)=   "Column(17).Width=12250"
         Splits(0)._ColumnProps(97)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(98)=   "Column(17)._WidthInPix=12171"
         Splits(0)._ColumnProps(99)=   "Column(17)._EditAlways=0"
         Splits(0)._ColumnProps(100)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(101)=   "Column(18).Width=1191"
         Splits(0)._ColumnProps(102)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(103)=   "Column(18)._WidthInPix=1111"
         Splits(0)._ColumnProps(104)=   "Column(18)._EditAlways=0"
         Splits(0)._ColumnProps(105)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(106)=   "Column(19).Width=1667"
         Splits(0)._ColumnProps(107)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(108)=   "Column(19)._WidthInPix=1588"
         Splits(0)._ColumnProps(109)=   "Column(19)._EditAlways=0"
         Splits(0)._ColumnProps(110)=   "Column(19).Order=20"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=94,.parent=13,.alignment=2"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=91,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=92,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=93,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=102,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=99,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=100,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=101,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=98,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=95,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=96,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=97,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=106,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=103,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=104,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=105,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=63,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=64,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=65,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=67,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=68,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=69,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=71,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=72,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=73,.parent=17"
         _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=78,.parent=13,.alignment=1"
         _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=75,.parent=14"
         _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=76,.parent=15"
         _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=77,.parent=17"
         _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=82,.parent=13,.alignment=1"
         _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=79,.parent=14"
         _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=80,.parent=15"
         _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=81,.parent=17"
         _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=86,.parent=13"
         _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=83,.parent=14"
         _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=84,.parent=15"
         _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=85,.parent=17"
         _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=90,.parent=13"
         _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=87,.parent=14"
         _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=88,.parent=15"
         _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=89,.parent=17"
         _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=110,.parent=13"
         _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
         _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
         _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
         _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=114,.parent=13"
         _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=14"
         _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=15"
         _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=17"
         _StyleDefs(116) =   "Named:id=33:Normal"
         _StyleDefs(117) =   ":id=33,.parent=0"
         _StyleDefs(118) =   "Named:id=34:Heading"
         _StyleDefs(119) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(120) =   ":id=34,.wraptext=-1"
         _StyleDefs(121) =   "Named:id=35:Footing"
         _StyleDefs(122) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(123) =   "Named:id=36:Selected"
         _StyleDefs(124) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(125) =   "Named:id=37:Caption"
         _StyleDefs(126) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(127) =   "Named:id=38:HighlightRow"
         _StyleDefs(128) =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(129) =   "Named:id=39:EvenRow"
         _StyleDefs(130) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(131) =   "Named:id=40:OddRow"
         _StyleDefs(132) =   ":id=40,.parent=33"
         _StyleDefs(133) =   "Named:id=41:RecordSelector"
         _StyleDefs(134) =   ":id=41,.parent=34"
         _StyleDefs(135) =   "Named:id=42:FilterBar"
         _StyleDefs(136) =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintcodigobaixa 
         Height          =   315
         HelpContextID   =   1
         Left            =   6960
         TabIndex        =   28
         Top             =   2070
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intNumeroAviso 
         Height          =   315
         HelpContextID   =   1
         Left            =   2865
         TabIndex        =   24
         Top             =   2055
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintlancamentovalor 
         Height          =   315
         HelpContextID   =   1
         Left            =   4290
         TabIndex        =   25
         Top             =   2055
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
         Height          =   315
         HelpContextID   =   1
         Left            =   5250
         TabIndex        =   18
         Top             =   1680
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strDescricaoConta 
         Height          =   315
         Left            =   4125
         TabIndex        =   7
         Top             =   840
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strComposicaoDaReceita 
         Height          =   315
         HelpContextID   =   1
         Left            =   6270
         TabIndex        =   19
         Top             =   1680
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Shape sha_Lote 
         BackColor       =   &H00FFFFFF&
         Height          =   1200
         Left            =   105
         Top             =   420
         Width           =   9600
      End
      Begin VB.Label lbl_QtdeDocs 
         AutoSize        =   -1  'True
         Caption         =   "Qtde. Docs."
         Height          =   195
         Left            =   3210
         TabIndex        =   3
         Top             =   585
         Width           =   855
      End
      Begin VB.Label lbl_bytTipo 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   5715
         TabIndex        =   46
         Top             =   495
         Width           =   3765
      End
      Begin VB.Label lblBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   4185
         TabIndex        =   13
         Top             =   1305
         Width           =   465
      End
      Begin VB.Label lblAgencia 
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   990
         TabIndex        =   11
         Top             =   1305
         Width           =   585
      End
      Begin VB.Label lblcodigoBaixa 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Baixa"
         Height          =   195
         Left            =   5850
         TabIndex        =   27
         Top             =   2145
         Width           =   975
      End
      Begin VB.Label lblExercicio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   900
         TabIndex        =   21
         Top             =   2145
         Width           =   675
      End
      Begin VB.Label lblCorreto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Correto"
         Height          =   195
         Left            =   7770
         TabIndex        =   41
         Top             =   2925
         Width           =   510
      End
      Begin VB.Label lblCodigoDeBarras 
         AutoSize        =   -1  'True
         Caption         =   "Código de Barras"
         Height          =   195
         Left            =   360
         TabIndex        =   39
         Top             =   2925
         Width           =   1215
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   7920
         TabIndex        =   37
         Top             =   2505
         Width           =   360
      End
      Begin VB.Label lbldblCorrecao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Correção"
         Height          =   195
         Left            =   5970
         TabIndex        =   35
         Top             =   2505
         Width           =   645
      End
      Begin VB.Label lbldblJuros 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Juros"
         Height          =   195
         Left            =   4500
         TabIndex        =   33
         Top             =   2505
         Width           =   375
      End
      Begin VB.Label lbldblMulta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Multa"
         Height          =   195
         Left            =   3030
         TabIndex        =   31
         Top             =   2505
         Width           =   390
      End
      Begin VB.Label lblTributo 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   3375
         TabIndex        =   17
         Top             =   1755
         Width           =   1695
      End
      Begin VB.Label lblDataDoPagamento 
         AutoSize        =   -1  'True
         Caption         =   "Data do Pagamento"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   1755
         Width           =   1425
      End
      Begin VB.Label lbldblPrincipal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Principal"
         Height          =   195
         Left            =   975
         TabIndex        =   29
         Top             =   2505
         Width           =   600
      End
      Begin VB.Label lblContaBancaria 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente"
         Height          =   195
         Left            =   495
         TabIndex        =   5
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label lblLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   8745
         TabIndex        =   9
         Top             =   960
         Width           =   315
      End
      Begin VB.Label lblDataDoMovimento 
         AutoSize        =   -1  'True
         Caption         =   "Data do Movimento"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   585
         Width           =   1395
      End
      Begin VB.Label lblAviso 
         AutoSize        =   -1  'True
         Caption         =   "Aviso"
         Height          =   195
         Left            =   2415
         TabIndex        =   23
         Top             =   2145
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmMovimentoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnAlterando            As Boolean
Dim bytOrdenacao            As Byte
Dim blnOrdenacaoAsc         As Boolean
Dim blnPrimeiraVez          As Boolean

Dim intQtdeDocsSalvos       As Integer
Dim blnNovoPorOutraOperacao As Boolean

Private Sub cmd_Composicao_Click()
    CarregaForm frmCadComposicaoDaReceita, dbc_intComposicaoDaReceita
End Sub

Private Sub cmd_ContaCorrente_Click()
    CarregaForm frmCadContasBancarias, dbcintContaBancaria
End Sub

Private Sub dbc_intComposicaoDaReceita_Change()
    If dbc_intComposicaoDaReceita.MatchedWithList Then
        If dbc_strComposicaoDaReceita.BoundText <> dbc_intComposicaoDaReceita.BoundText Then
            PreencherListaDeOpcoes dbc_strComposicaoDaReceita, dbc_intComposicaoDaReceita.BoundText
            
            txt_intExercicio.Text = ""
            dbc_intNumeroAviso.BoundText = ""
            Set dbc_intNumeroAviso.RowSource = Nothing
            dbcintlancamentovalor.BoundText = ""
            Set dbcintlancamentovalor.RowSource = Nothing
            txtintDigito.Text = ""

        End If
     End If
End Sub

Private Sub dbc_intComposicaoDaReceita_LostFocus()
    LeDaTabelaParaObj "", dbc_intComposicaoDaReceita, strQueryComposicao
    If Not dbc_intComposicaoDaReceita.MatchedWithList Then
        dbc_strComposicaoDaReceita.BoundText = ""
        Set dbc_strComposicaoDaReceita.RowSource = Nothing
    End If
    
End Sub

Private Sub dbc_strComposicaoDaReceita_Change()
    If dbc_strComposicaoDaReceita.MatchedWithList Then
        If dbc_strComposicaoDaReceita.BoundText <> dbc_intComposicaoDaReceita.BoundText Then
            PreencherListaDeOpcoes dbc_intComposicaoDaReceita, dbc_strComposicaoDaReceita.BoundText
            
            txt_intExercicio.Text = ""
            dbc_intNumeroAviso.BoundText = ""
            Set dbc_intNumeroAviso.RowSource = Nothing
            dbcintlancamentovalor.BoundText = ""
            Set dbcintlancamentovalor.RowSource = Nothing
            txtintDigito.Text = ""

        End If
    End If
End Sub

Private Sub dbc_strComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_strComposicaoDaReceita, Me, Area
End Sub

Private Sub dbc_strComposicaoDaReceita_GotFocus()
    MarcaCampo dbc_strComposicaoDaReceita
End Sub

Private Sub dbc_strComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strComposicaoDaReceita
End Sub

Private Sub dbc_strDescricaoConta_Change()
    If dbc_strDescricaoConta.MatchedWithList Then
        If dbc_strDescricaoConta.BoundText <> dbcintContaBancaria.BoundText Then
            PreencherListaDeOpcoes dbcintContaBancaria, dbc_strDescricaoConta.BoundText
        End If
    End If
End Sub

Private Sub dbc_strDescricaoConta_Click(Area As Integer)
    DropDownDataCombo dbc_strDescricaoConta, Me, Area
End Sub

Private Sub dbc_strDescricaoConta_GotFocus()
    MarcaCampo dbc_strDescricaoConta
End Sub

Private Sub dbc_strDescricaoConta_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strDescricaoConta, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strDescricaoConta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strDescricaoConta
End Sub

Private Sub dbcintcodigobaixa_Change()
Dim adoResultado    As New ADODB.Recordset
Dim blnCancelamento As Boolean
    
    blnCancelamento = False
    
    If dbcintcodigobaixa.MatchedWithList Then
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO("SELECT bytTipo FROM " & gstrCodigoDeBaixa & " WHERE Pkid = " & dbcintcodigobaixa.BoundText, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                blnCancelamento = adoResultado("bytTipo").Value = 2
            End If
        End If
    End If
    
    If blnCancelamento Then
        txtdblPrincipal = "0,00"
        txtdblMulta = "0,00"
        txtdblJuros = "0,00"
        txtdblCorrecao = "0,00"
        txt_dblTotal = "0,00"
        txtdblCorreto = "0,00"
    End If
    
    TrocaCorObjeto txtdblPrincipal, blnCancelamento
    TrocaCorObjeto txtdblMulta, blnCancelamento
    TrocaCorObjeto txtdblJuros, blnCancelamento
    TrocaCorObjeto txtdblCorrecao, blnCancelamento
    TrocaCorObjeto txt_dblTotal, blnCancelamento
    TrocaCorObjeto txtdblCorreto, True
        
End Sub

Private Sub dbcintContaBancaria_LostFocus()
    LeDaTabelaParaObj "", dbcintContaBancaria, strQueryContaCorrente
    If Not dbcintContaBancaria.MatchedWithList Then
        dbc_strDescricaoConta.BoundText = ""
        Set dbc_strDescricaoConta.RowSource = Nothing
        txt_strAgencia.Text = ""
        txt_strBanco.Text = ""
    End If
End Sub

Private Sub dbcintLancamentoValor_Click(Area As Integer)
    DropDownDataCombo dbcintlancamentovalor, Me, Area
End Sub

Private Sub dbcintLancamentoValor_GotFocus()
    MarcaCampo dbcintlancamentovalor
End Sub

Private Sub dbcintLancamentoValor_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintlancamentovalor, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLancamentoValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintlancamentovalor
End Sub

Private Sub dbcintLancamentoValor_LostFocus()
Dim adoResultado As New ADODB.Recordset

    If dbc_intNumeroAviso.MatchedWithList And Trim(dbcintlancamentovalor.Text) <> "" Then
        
        LeDaTabelaParaObj "", dbcintlancamentovalor, strQueryParcela
        
        txtintDigito = gstrCalculaDigitoModulo10(Trim(dbc_intNumeroAviso.Text) & Format$(Trim(dbcintlancamentovalor.Text), "000"))
        
        'Vamos carregar a combo de tipos de baixa de acordo com o vencimento da parcela
        If gblnDataValida(txtdtmDtPagamento.Text) Then
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strQueryParcela, 5, adoResultado) Then
                If Not adoResultado.EOF Then
                    LeDaTabelaParaObj "", dbcintcodigobaixa, strQueryCodigoBaixa(adoResultado("dtmDtVencimento").Value)
                End If
            End If
        End If
        
        If dbcintlancamentovalor.MatchedWithList = False Then
            dbcintlancamentovalor.SetFocus
        End If
        
        'Ja vamos fazer o calculo dos valores
        If Not blnDadosCalculoOk Then Exit Sub
            
        CalculaReajuste

    Else
        Set dbcintlancamentovalor.RowSource = Nothing
            dbcintlancamentovalor.Text = ""
    End If
End Sub

Private Sub dbc_intNumeroAviso_Click(Area As Integer)
    DropDownDataCombo dbc_intNumeroAviso, Me, Area
End Sub

Private Sub dbc_intNumeroAviso_GotFocus()
    MarcaCampo dbc_intNumeroAviso
End Sub

Private Sub dbc_intNumeroAviso_KeyDown(KeyCode As Integer, Shift As Integer)
     DropDownDataCombo dbc_intNumeroAviso, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intNumeroAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intNumeroAviso
End Sub

Private Sub dbc_intNumeroAviso_LostFocus()
    If dbc_intComposicaoDaReceita.MatchedWithList And Len(Trim(txt_intExercicio)) = 4 And Trim(dbc_intNumeroAviso.Text) <> "" Then
        
        LeDaTabelaParaObj "", dbc_intNumeroAviso, strQueryAviso
        If dbc_intNumeroAviso.MatchedWithList = False Then
            dbc_intNumeroAviso.SetFocus
        End If
    Else
        Set dbc_intNumeroAviso.RowSource = Nothing
        dbc_intNumeroAviso.Text = ""
    End If
End Sub

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
End Sub

Private Sub dbc_intComposicaoDaReceita_GotFocus()
    MarcaCampo dbc_intComposicaoDaReceita
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intComposicaoDaReceita
End Sub

Private Sub dbcintcodigobaixa_Click(Area As Integer)
    DropDownDataCombo dbcintcodigobaixa, Me, Area
End Sub

Private Sub dbcintcodigobaixa_GotFocus()
    MarcaCampo dbcintcodigobaixa
End Sub

Private Sub dbcintcodigobaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContaBancaria, Me, , KeyCode, Shift
End Sub

Private Sub dbcintcodigobaixa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintcodigobaixa
End Sub

Private Sub dbcintcontabancaria_Change()
    If dbcintContaBancaria.MatchedWithList Then
        If dbc_strDescricaoConta.BoundText <> dbcintContaBancaria.BoundText Then
            PreencherListaDeOpcoes dbc_strDescricaoConta, dbcintContaBancaria.BoundText
        End If
        PreencheAgBanco (dbcintContaBancaria.BoundText)
    End If
End Sub

Private Sub dbcintcontabancaria_Click(Area As Integer)
    DropDownDataCombo dbcintContaBancaria, Me, Area
End Sub

Private Sub dbcintContaBancaria_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContaBancaria, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContaBancaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContaBancaria
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1110
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    dbc_intComposicaoDaReceita.Tag = strQueryComposicao(True) & ";strDescricao"
    dbc_strComposicaoDaReceita.Tag = strQueryComposicaoDescricao & ";strdescricao"
    dbcintContaBancaria.Tag = strQueryContaCorrente(True) & ";intNumeroConta"
    dbcintcodigobaixa.Tag = strQueryCodigoBaixa & ";strAbreviatura"
    dbc_strDescricaoConta.Tag = strQueryContaDescricao & ";strdescricao"
    TrocaCorObjeto txt_strCodigoDeBarras, True
    TrocaCorObjeto txt_strAgencia, True
    TrocaCorObjeto txt_strBanco, True
    TrocaCorObjeto txtdblCorreto, True
    blnAlterando = False
    txtbytTipo.Text = 0
    lbl_bytTipo.Caption = "Manual"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub tdb_MovimentoBancario_Click()
   blnPrimeiraVez = True
End Sub

Private Sub tdb_MovimentoBancario_FilterChange()
    gblnFilraCampos tdb_MovimentoBancario
    blnPrimeiraVez = False
End Sub

Private Sub tdb_MovimentoBancario_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_MovimentoBancario, ColIndex
    blnPrimeiraVez = False
End Sub

Private Sub tdb_MovimentoBancario_KeyDown(KeyCode As Integer, Shift As Integer)
    blnPrimeiraVez = True
End Sub

Private Sub tdb_MovimentoBancario_KeyPress(KeyAscii As Integer)
    Select Case tdb_MovimentoBancario.Col
        Case 1
            CaracterValido KeyAscii, "D", tdb_MovimentoBancario
        Case 2
            CaracterValido KeyAscii, "A", tdb_MovimentoBancario
        Case 3
            CaracterValido KeyAscii, "N", tdb_MovimentoBancario
        Case 4
            CaracterValido KeyAscii, "D", tdb_MovimentoBancario
        Case 5
            CaracterValido KeyAscii, "A", tdb_MovimentoBancario
        Case 6
            CaracterValido KeyAscii, "N", tdb_MovimentoBancario
        Case 7
            CaracterValido KeyAscii, "V", tdb_MovimentoBancario
        Case 8
            CaracterValido KeyAscii, "V", tdb_MovimentoBancario
        Case 9
            CaracterValido KeyAscii, "V", tdb_MovimentoBancario
        Case 10
            CaracterValido KeyAscii, "V", tdb_MovimentoBancario
        Case 11
            CaracterValido KeyAscii, "V", tdb_MovimentoBancario
    End Select
End Sub

Private Sub tdb_MovimentoBancario_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
With tdb_MovimentoBancario
    If Not .EOF And blnPrimeiraVez Then
        txtPKId.Text = .Columns("PKID").Value
        LeDaTabelaParaObj gstrMovimentoBancario, Me
        txt_strCodigoDeBarras.Text = .Columns("CodigoDeBarras").Value
        Select Case Trim(txtbytTipo)
            Case Is = 0
                lbl_bytTipo.Caption = "Manual"
            Case Is = 1
                lbl_bytTipo.Caption = "Boca de Caixa"
            Case Is = 2
                lbl_bytTipo.Caption = "Arrecadação Eletrônica"
            Case Is = 3
                lbl_bytTipo.Caption = "Código de Barras"
            Case Is = 9
                lbl_bytTipo.Caption = "Outros"
        End Select
        'HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
        blnAlterando = True
        PreencheDados
        dblValorTotal
    End If
End With
End Sub
Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case UCase(strModoOperacao)
        
        Case Is = UCase(gstrSalvar)
            If Not blnDadosOk Then Exit Sub
            VerificaValores
            If ToolBarGeral(strModoOperacao, gstrMovimentoBancario, blnAlterando, tdb_MovimentoBancario, Me, , strQuery(gstrSalvar), , , , False) Then
                'Vamos somar os docs salvos
                intQtdeDocsSalvos = intQtdeDocsSalvos + 1
                blnNovoPorOutraOperacao = True
                MantemForm gstrNovo
            End If
            
            If Not blnAlterando Then
                blnPrimeiraVez = False
            End If
            
        Case Is = UCase(gstrPreencherLista)
            If Me.ActiveControl.Name = "dbc_intNumeroAviso" Then
                If dbc_intComposicaoDaReceita.MatchedWithList And Len(Trim(txt_intExercicio)) = 4 Then
                    LeDaTabelaParaObj "", dbc_intNumeroAviso, strQueryAviso(True)
                End If
            ElseIf Me.ActiveControl.Name = "dbcintlancamentovalor" Then
                If dbc_intNumeroAviso.MatchedWithList Then
                    LeDaTabelaParaObj "", dbcintlancamentovalor, strQueryParcela(True)
                End If
            Else
                PreencherListaDeOpcoes Me.ActiveControl
            End If
            
        Case Is = UCase(gstrNovo)
        
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
            
            txt_intExercicio.Text = ""
            dbc_intNumeroAviso.BoundText = ""
            Set dbc_intNumeroAviso.RowSource = Nothing
            dbcintlancamentovalor.BoundText = ""
            Set dbcintlancamentovalor.RowSource = Nothing
            txtintDigito.Text = ""
            dbcintcodigobaixa.BoundText = ""
            Set dbcintcodigobaixa.RowSource = Nothing
            txtdblPrincipal.Text = ""
            txtdblMulta.Text = ""
            txtdblJuros.Text = ""
            txtdblCorrecao.Text = ""
            txt_dblTotal.Text = ""
            txt_strCodigoDeBarras.Text = ""
            txtdblCorreto.Text = ""
            
            If txt_QtdeDocs <> Trim(Str(intQtdeDocsSalvos)) And Trim(txt_QtdeDocs) <> "" And Not blnNovoPorOutraOperacao And intQtdeDocsSalvos <> 0 Then
                If MsgBox("A quantidade de movimentos lançados não é a mesma da quantidade informada. Deseja limpar o formulário?", vbYesNo, "Mensagem ao Usuário") = vbYes Then
                    intQtdeDocsSalvos = 0
                End If
            End If
            
            'Caso a quantidade cadastrada seja igual a qtde informada vamos limpar o form todo
            If txt_QtdeDocs = Trim(Str(intQtdeDocsSalvos)) Or intQtdeDocsSalvos = 0 Or Trim(txt_QtdeDocs) = "" Then
                txtPKId.Text = ""
                txtbytTipo.Text = 0
                lbl_bytTipo.Caption = "Manual"
                txtdtmDtMovimento_GotFocus
                txtdtmDtMovimento.SetFocus
                txt_QtdeDocs.Text = ""
                dbcintContaBancaria.BoundText = ""
                Set dbcintContaBancaria.RowSource = Nothing
                dbc_strDescricaoConta.BoundText = ""
                Set dbc_strDescricaoConta.RowSource = Nothing
                txtintLote.Text = ""
                txt_strAgencia.Text = ""
                txt_strBanco.Text = ""
                txtdtmDtPagamento.Text = ""
                dbc_intComposicaoDaReceita.BoundText = ""
                Set dbc_intComposicaoDaReceita.RowSource = Nothing
                dbc_strComposicaoDaReceita.BoundText = ""
                Set dbc_strComposicaoDaReceita.RowSource = Nothing
                
                intQtdeDocsSalvos = 0
                
            Else
                txtdtmDtPagamento_GotFocus
                txtdtmDtPagamento.SetFocus
            End If
            
            'Limpa_Controles Me, True, False, False, True, False
            blnNovoPorOutraOperacao = False
            
            blnAlterando = False
            blnPrimeiraVez = False
            
        Case Is = UCase(gstrLocalizar)
            ToolBarGeral strModoOperacao, gstrMovimentoBancario, blnAlterando, tdb_MovimentoBancario, Me, , strQuery(gstrLocalizar)
            
        Case Is = UCase(gstrDeletar)
        
            If ApagaMovimentoBancario Then
                intQtdeDocsSalvos = 0
                blnNovoPorOutraOperacao = True
                blnPrimeiraVez = False
                MantemForm gstrLocalizar
                MantemForm gstrNovo
            End If

'            If ToolBarGeral(strModoOperacao, gstrMovimentoBancario, blnAlterando, tdb_MovimentoBancario, Me, , strQuery) Then
'                intQtdeDocsSalvos = 0
'                blnNovoPorOutraOperacao = True
'                blnPrimeiraVez = False
'                MantemForm gstrLocalizar
'                MantemForm gstrNovo
'            End If
            
        Case Else
            ToolBarGeral strModoOperacao, gstrMovimentoBancario, blnAlterando, tdb_MovimentoBancario, Me, , strQuery
            
    End Select
                 
End Sub

Private Function blnDadosOk()
    
    blnDadosOk = False
    
    If Not gblnDataValida(txtdtmDtMovimento) Then
        ExibeMensagem "A Data informada não é válida."
        txtdtmDtMovimento.SetFocus
        Exit Function
    End If
    
    If Not dbcintContaBancaria.MatchedWithList Then
        ExibeMensagem "Selecione uma Conta Corrente válida."
        dbcintContaBancaria.SetFocus
        Exit Function
    End If
    
    If txtintLote.Text = "" Then
        ExibeMensagem "O Lote deve ser preenchido."
        txtintLote.SetFocus
        Exit Function
    End If
    
    If Not gblnDataValida(txtdtmDtPagamento.Text) Then
       ExibeMensagem "A Data de pagamento informada não é válida."
       txtdtmDtPagamento.SetFocus
       Exit Function
    End If
    
    If dbc_intComposicaoDaReceita.MatchedWithList = False Then
        ExibeMensagem "A Composição da Receita deve ser preenchida corretamente."
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intExercicio) = "" Then
        ExibeMensagem "O execício deve ser preenchido corretamente."
        txt_intExercicio.SetFocus
        Exit Function
    End If
    
    If dbc_intNumeroAviso.MatchedWithList = False Then
        ExibeMensagem "O Aviso deve ser preenchido corretamente."
        dbc_intNumeroAviso.SetFocus
        Exit Function
    End If
    
    If dbcintlancamentovalor.MatchedWithList = False Then
        ExibeMensagem "A Parcela deve ser preenchida corretamente."
        dbcintlancamentovalor.SetFocus
        Exit Function
    End If
    
    If Trim(txtintDigito) = "" Then
        ExibeMensagem "O Dígito deve ser preenchido corretamente."
        txtintDigito.SetFocus
        Exit Function
    End If
    
    If gstrCalculaDigitoModulo10(Trim(dbc_intNumeroAviso.Text) & Format$(Trim(dbcintlancamentovalor.Text), "000")) <> Trim(txtintDigito) Then
        ExibeMensagem "Aviso inválido."
        txtintDigito = ""
        txtintDigito.SetFocus
        Exit Function
    End If

    If dbcintcodigobaixa.MatchedWithList = False Then
        ExibeMensagem "O Tipo de Baixa deve ser preenchido corretamente."
        dbcintcodigobaixa.SetFocus
        Exit Function
    End If
    
    
    If CCur(txtdblPrincipal.Text) = 0 Then
        ExibeMensagem "O Campo Principal deve ser informado."
        txtdblPrincipal.SetFocus
        Exit Function
    End If

    blnDadosOk = True
    
End Function

Private Function blnDadosCalculoOk()

    blnDadosCalculoOk = False
    
    If dbc_intComposicaoDaReceita.MatchedWithList = False Then
        'ExibeMensagem "A Composição da Recita deve ser preenchida corretamente."
        'dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    ElseIf Trim(txt_intExercicio) = "" Then
        'ExibeMensagem "O Exercício deve ser preenchido corretamente."
        'txt_intExercicio.SetFocus
        Exit Function
    ElseIf dbc_intNumeroAviso.MatchedWithList = False Then
        'ExibeMensagem "O número do aviso deve ser preenchido corretamente."
        'dbc_intNumeroAviso.SetFocus
        Exit Function
    ElseIf dbcintlancamentovalor.MatchedWithList = False Then
        'ExibeMensagem "A parcela deve ser preenchida corretamente."
        'dbcintlancamentovalor.SetFocus
        Exit Function
    'ElseIf Trim(txtintDigito.Text) = "" Then
    '    ExibeMensagem "O Digito deve ser preenchido corretamente."
    '    txtintDigito.SetFocus
    '    Exit Function
    End If
    
    blnDadosCalculoOk = True
    
End Function

Private Function strQuery(Optional strModoOperacao As String) As String
Dim strsql As String

    strsql = "SELECT MB.Pkid,"
    strsql = strsql & " MB.dtmDtMovimento Movimento,"
    strsql = strsql & " CB.strConta " & strCONCAT & "'-'" & strCONCAT & " strDigitoVerificador ContaCorrente,"
    strsql = strsql & " MB.intLote Lote,"
    strsql = strsql & " MB.dtmDtPagamento Pagamento,"
    strsql = strsql & " MB.byttipo,"
    strsql = strsql & " CR.strDescricao Tributo,"
    strsql = strsql & " LA.intExercicio,"
    strsql = strsql & gstrCONVERT(CDT_numeric, "LA.STRNUMEROAVISO") & " STRNUMEROAVISO,"
    strsql = strsql & " LV.Intparcela,"
    strsql = strsql & " MB.Intdigito,"
    strsql = strsql & " B.strabreviatura as IntCodigoBaixa,"
    strsql = strsql & " MB.dblPrincipal Principal,"
    strsql = strsql & " MB.dblMulta Multa,"
    strsql = strsql & " MB.dblJuros Juros,"
    strsql = strsql & " MB.dblCorrecao Correcao,"
    strsql = strsql & gstrCASEWHEN("MB.bitGuia", "0, 'Não'", "'Sim'") & " bitGuia, "
    strsql = strsql & gstrCASEWHEN("MB.bitProcessado", "1, 'Sim'", "'Não'") & " bitProcessado, "
    strsql = strsql & " (MB.dblPrincipal + MB.dblMulta + MB.dblJuros + MB.dblCorrecao) Total, "
    strsql = strsql & strSUBSTRING & "(LTRIM(RTRIM(MB.strCodigoDeBarras)),1,1) Tipo,"
    strsql = strsql & " MB.strCodigoDeBarras CodigoDeBarras"
    
    If bytDBType = EDatabases.Oracle Then
        strsql = strsql & " FROM "
        strsql = strsql & gstrMovimentoBancario & " MB, "
        strsql = strsql & gstrContaBancaria & " CB, "
        strsql = strsql & gstrLancamentoAlfa & " LA, "
        strsql = strsql & gstrLancamentoValor & " LV, "
        strsql = strsql & gstrComposicaoDaReceita & " CR, "
        strsql = strsql & gstrCodigoDeBaixa & " B "
        strsql = strsql & " WHERE "
        strsql = strsql & "CB.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " MB.intContaBancaria    AND "
        strsql = strsql & "LV.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " MB.Intlancamentovalor  AND "
        strsql = strsql & "LA.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " LV.Intlancamentoalfa   AND "
        strsql = strsql & "CR.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.Intcomposicaodareceita AND "
        strsql = strsql & "B.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " MB.Intcodigobaixa"
    Else
        strsql = strsql & " FROM " & gstrMovimentoBancario & " MB LEFT OUTER JOIN " & _
                       gstrContaBancaria & " CB ON MB.intContaBancaria = CB.PKId LEFT OUTER JOIN " & _
                       gstrLancamentoValor & " LV ON MB.intlancamentovalor = LV.PKId LEFT OUTER JOIN " & _
                       gstrLancamentoAlfa & " LA ON LV.intLancamentoAlfa = LA.PKId LEFT OUTER JOIN " & _
                       gstrComposicaoDaReceita & " CR ON LA.INTCOMPOSICAODARECEITA = CR.PKId LEFT OUTER JOIN " & _
                       gstrCodigoDeBaixa & " B ON MB.INTCODIGOBAIXA = B.PKID "
        strsql = strsql & "WHERE MB.Pkid > 1 "
    End If
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If Not blnAlterando Then
            strsql = strsql & " AND MB.Pkid = " & glngPegaUltimaChave(gstrMovimentoBancario, "Pkid") + 1
        Else
            strsql = strsql & " AND MB.Pkid = " & txtPKId.Text
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrLocalizar) Then
        If dbc_intComposicaoDaReceita.MatchedWithList Then
            strsql = strsql & " AND CR.Pkid = " & dbc_intComposicaoDaReceita.BoundText
        End If
        
        If Trim(txt_intExercicio.Text) <> "" Then
            strsql = strsql & " AND LA.intExercicio = " & txt_intExercicio.Text
        End If
    End If
    
    Select Case bytOrdenacao
        Case Is = 1
            strsql = strsql & " ORDER BY MB.dtmDtMovimento " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 2
            strsql = strsql & " ORDER BY CB.intNumeroConta " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 3
            strsql = strsql & " ORDER BY MB.intLote" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 4
            strsql = strsql & " ORDER BY MB.dtmDtPagamento" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 5
            strsql = strsql & " ORDER BY CR.strDescricao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 6
            strsql = strsql & " ORDER BY LA.intExercicio" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 7
            strsql = strsql & " ORDER BY LA.STRNUMEROAVISO" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 8
            strsql = strsql & " ORDER BY LV.Intparcela" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 9
            strsql = strsql & " ORDER BY MB.Intdigito" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 10
            strsql = strsql & " ORDER BY MB.dblPrincipal" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 11
            strsql = strsql & " ORDER BY MB.dblMulta" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 12
            strsql = strsql & " ORDER BY MB.dblJuros" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 13
            strsql = strsql & " ORDER BY MB.dblCorrecao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case Is = 14
            strsql = strsql & " ORDER BY Total" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strsql

End Function

Private Sub txt_intExercicio_Change()
    Set dbc_intNumeroAviso.RowSource = Nothing
    dbc_intNumeroAviso.Text = ""
    Set dbcintlancamentovalor.RowSource = Nothing
    dbcintlancamentovalor.Text = ""
    txtintDigito.Text = ""
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_QtdeDocs_GotFocus()
    MarcaCampo txt_QtdeDocs
End Sub

Private Sub txt_QtdeDocs_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_QtdeDocs
End Sub

Private Sub txtdblCorrecao_GotFocus()
    MarcaCampo txtdblCorrecao
End Sub

Private Sub txtdblCorrecao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblCorrecao
End Sub

Private Sub txtdblCorrecao_LostFocus()
    txtdblCorrecao = gstrConvVrDoSql(txtdblCorrecao, 2)
    dblValorTotal
End Sub

Private Sub txtdblCorreto_GotFocus()
    MarcaCampo txtdblCorreto
End Sub

Private Sub txtdblCorreto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblCorreto
End Sub

Private Sub txtdblCorreto_LostFocus()
    txtdblCorreto = gstrConvVrDoSql(txtdblCorreto, 2)
End Sub

Private Sub txtdblJuros_GotFocus()
    MarcaCampo txtdblJuros
End Sub

Private Sub txtdblJuros_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblJuros
End Sub

Private Sub txtdblJuros_LostFocus()
    txtdblJuros = gstrConvVrDoSql(txtdblJuros, 2)
    dblValorTotal
End Sub

Private Sub txtdblMulta_GotFocus()
    MarcaCampo txtdblMulta
End Sub

Private Sub txtdblMulta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblMulta
End Sub

Private Sub txtdblMulta_LostFocus()
    txtdblMulta = gstrConvVrDoSql(txtdblMulta, 2)
    dblValorTotal
End Sub

Private Sub txtdblPrincipal_GotFocus()
    MarcaCampo txtdblPrincipal
End Sub

Private Sub txtdblPrincipal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblPrincipal
End Sub

Private Sub txtdblPrincipal_LostFocus()
    txtdblPrincipal = gstrConvVrDoSql(txtdblPrincipal, 2)
    dblValorTotal
End Sub

Private Sub txt_dblTotal_GotFocus()
    MarcaCampo txt_dblTotal
End Sub

Private Sub txt_dblTotal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblTotal
End Sub

Private Sub txt_dblTotal_LostFocus()
    txt_dblTotal = gstrConvVrDoSql(txt_dblTotal, 2)
    dblValorPrincipal
End Sub

Private Sub txtdtmDtMovimento_GotFocus()
    If txtdtmDtMovimento.Text = "" Then txtdtmDtMovimento = gstrDataDoSistema
    MarcaCampo txtdtmDtMovimento
End Sub

Private Sub txtdtmDtMovimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtMovimento
End Sub

Private Sub txtdtmDtMovimento_LostFocus()
    txtdtmDtMovimento = gstrDataFormatada(txtdtmDtMovimento)
End Sub

Private Sub txtdtmDtPagamento_GotFocus()
    If txtdtmDtPagamento.Text = "" Then txtdtmDtPagamento = txtdtmDtMovimento
    MarcaCampo txtdtmDtPagamento
End Sub

Private Sub txtdtmDtPagamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtPagamento
End Sub

Private Sub txtdtmDtPagamento_LostFocus()
    txtdtmDtPagamento = gstrDataFormatada(txtdtmDtPagamento)
End Sub

Private Sub txtintDigito_GotFocus()
    MarcaCampo txtintDigito
End Sub

Private Sub txtintDigito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintDigito
End Sub

Private Sub txtintDigito_LostFocus()
    If dbc_intNumeroAviso.MatchedWithList And dbcintlancamentovalor.MatchedWithList Then
        If gstrCalculaDigitoModulo10(Trim(dbc_intNumeroAviso.Text) & Format$(Trim(dbcintlancamentovalor.Text), "000")) <> Trim(txtintDigito) Then
            ExibeMensagem "Aviso inválido."
            txtintDigito = ""
            'txtintDigito.SetFocus
        End If
    End If
End Sub

Private Sub txtintLote_GotFocus()
    MarcaCampo txtintLote
End Sub

Private Sub txtintLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintLote
End Sub

Private Function strQueryContaCorrente(Optional blnF5 As Boolean) As String
    Dim strsql As String

    strsql = "SELECT CB.Pkid, "
    strsql = strsql & "intNumeroConta ContaCorrente"
    strsql = strsql & " FROM " & gstrContaBancaria & " CB, "
    strsql = strsql & gstrPlanoConta & " PC"
    strsql = strsql & " Where"
    strsql = strsql & " CB.Pkid = PC.Intcontabancaria"
    If blnF5 = False Then
        strsql = strsql & " AND CB.intNumeroConta = " & Val(dbcintContaBancaria.Text)
    End If
    strsql = strsql & " ORDER BY intNumeroConta, strDigitoVerificador"
    
    strQueryContaCorrente = strsql

End Function

Private Function strCompletaNumero(strValor As String, intNumeroCasas As Integer) As String
            
Dim intI       As Integer
Dim strDigito  As String
    
    For intI = 1 To gstrENulo(intNumeroCasas) - Len(strValor)
        strDigito = strDigito & "0"
    Next intI
    strCompletaNumero = strDigito
    
End Function

Private Function strQueryComposicao(Optional blnF5 As Boolean) As String
    Dim strsql As String
    
    strsql = "SELECT Pkid,"
    strsql = strsql & "intCodigo "
    strsql = strsql & " FROM "
    strsql = strsql & gstrComposicaoDaReceita
    If blnF5 = False Then
        strsql = strsql & " WHERE intCodigo = " & Val(dbc_intComposicaoDaReceita.Text)
    End If
    strsql = strsql & " ORDER BY intCodigo"
    
    strQueryComposicao = strsql

End Function

Private Function strQueryComposicaoDescricao() As String
    Dim strsql As String
    
    strsql = "SELECT Pkid,"
    strsql = strsql & " strDescricao Descricao "
    strsql = strsql & " FROM "
    strsql = strsql & gstrComposicaoDaReceita
    strsql = strsql & " ORDER BY strDescricao"
    
    strQueryComposicaoDescricao = strsql

End Function

Private Sub dblValorTotal()
    Dim dblValorTotal As Variant
    
    dblValorTotal = CDbl(gstrConvVrDoSql(txtdblPrincipal.Text, 2, , True)) + _
                    CDbl(gstrConvVrDoSql(txtdblMulta.Text, 2, , True)) + _
                    CDbl(gstrConvVrDoSql(txtdblJuros.Text, 2, , True)) + _
                    CDbl(gstrConvVrDoSql(txtdblCorrecao.Text, 2, , True))
    
    txt_dblTotal = gstrConvVrDoSql(dblValorTotal, 2)
    
End Sub

Private Sub dblValorPrincipal()
    Dim dblValorPrincipal As Variant
    
    dblValorPrincipal = CDbl(gstrConvVrDoSql(txt_dblTotal.Text, 2, , True)) - _
                        CDbl(gstrConvVrDoSql(txtdblMulta.Text, 2, , True)) - _
                        CDbl(gstrConvVrDoSql(txtdblJuros.Text, 2, , True)) - _
                        CDbl(gstrConvVrDoSql(txtdblCorrecao.Text, 2, , True))
    
    txtdblPrincipal = gstrConvVrDoSql(dblValorPrincipal, 2)
    
End Sub


Private Function strQueryAviso(Optional blnF5 As Boolean) As String
   
    Dim strsql As String
    
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "LA.Pkid, "
    strsql = strsql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso "
    strsql = strsql & "From "
    strsql = strsql & gstrLancamentoAlfa & " LA "
    strsql = strsql & "Where "
    strsql = strsql & "LA.Intcomposicaodareceita = " & dbc_intComposicaoDaReceita.BoundText & " AND "
    strsql = strsql & "LA.intExercicio = " & Trim(txt_intExercicio)
    If blnF5 = False Or Trim(dbc_intNumeroAviso) <> "" Then
        strsql = strsql & " AND LA.strNumeroAviso = '" & String(gintLenNumAviso - Len(Trim(dbc_intNumeroAviso.Text)), "0") & Val(dbc_intNumeroAviso.Text) & "'"
    End If
    strQueryAviso = strsql
End Function

Private Function strQueryParcela(Optional blnF5 As Boolean) As String
    Dim strsql As String
    
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "LV.Pkid, "
    strsql = strsql & "LV.Intparcela, "
    strsql = strsql & "LV.dtmDtVencimento "
    strsql = strsql & "From "
    strsql = strsql & gstrLancamentoValor & " LV "
    strsql = strsql & "Where "
    strsql = strsql & "LV.Intlancamentoalfa = " & dbc_intNumeroAviso.BoundText
    If blnF5 = False Then
        strsql = strsql & " AND LV.Intparcela = " & dbcintlancamentovalor.Text
    End If
    strQueryParcela = strsql
End Function

Private Sub PreencheAgBanco(lngPkidContaBancaria As Long)
Dim adoResultado    As ADODB.Recordset
Dim strsql          As String
    strsql = "SELECT BA.strDescricao Banco,"
    strsql = strsql & " AG.strDescricao Agencia"
    strsql = strsql & " FROM "
    strsql = strsql & gstrContaBancaria & " CB, "
    strsql = strsql & gstrBanco & " BA, "
    strsql = strsql & gstrAgencia & " AG"
    strsql = strsql & " WHERE"
    strsql = strsql & " BA.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CB.intBanco AND"
    strsql = strsql & " AG.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " CB.intAgencia AND"
    strsql = strsql & " CB.Pkid = " & lngPkidContaBancaria
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_strAgencia.Text = gstrENulo(adoResultado!Agencia)
            txt_strBanco.Text = gstrENulo(adoResultado!Banco)
        Else
            txt_strAgencia.Text = ""
            txt_strBanco.Text = ""
        End If
    End If
End Sub

Private Function strQueryCodigoBaixa(Optional dtmVencimento As Date) As String
    Dim strsql As String
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "Pkid, "
    strsql = strsql & "strAbreviatura "
    strsql = strsql & "From "
    strsql = strsql & gstrCodigoDeBaixa & " "
    If Trim(txtdtmDtMovimento) <> "" And Trim(dtmVencimento) <> "" Then
        If gblnDataValida(txtdtmDtPagamento) Then
            If CDate(dtmVencimento) < CDate(txtdtmDtPagamento) Then
                strsql = strsql & " WHERE BytTipo = 4 "
            Else
                strsql = strsql & " WHERE BytTipo = 0 "
            End If
        Else
            strsql = strsql & " WHERE BytTipo = 0 "
        End If
        strsql = strsql & "Order By Pkid "
    Else
        strsql = strsql & "Order By strabreviatura "
    End If
    
    strQueryCodigoBaixa = strsql
    
End Function

Private Sub PreencheDados()
    Dim strsql As String
    Dim strDigito As String
    
    strDigito = txtintDigito
    
    dbc_strComposicaoDaReceita.BoundText = ""
    Set dbc_strComposicaoDaReceita.RowSource = Nothing

    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "CPR.Pkid, "
    strsql = strsql & "CPR.intCodigo "
    strsql = strsql & "From "
    strsql = strsql & gstrMovimentoBancario & " MB, "
    strsql = strsql & gstrLancamentoValor & " LV, "
    strsql = strsql & gstrLancamentoAlfa & " LA, "
    strsql = strsql & gstrComposicaoDaReceita & " CPR "
    strsql = strsql & "Where "
    strsql = strsql & "LV.Pkid = MB.INTLANCAMENTOVALOR AND "
    strsql = strsql & "LA.Pkid = LV.INTLANCAMENTOALFA AND "
    strsql = strsql & "CPR.Pkid = LA.Intcomposicaodareceita AND "
    strsql = strsql & "MB.Pkid = " & tdb_MovimentoBancario.Columns("PKID").Value
    strsql = strsql & " Order By CPR.STRDESCRICAO"
    
    LeDaTabelaParaObj "", dbc_intComposicaoDaReceita, strsql
    
    txt_intExercicio = tdb_MovimentoBancario.Columns("intExercicio").Value
    
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "LA.Pkid, "
    strsql = strsql & "LA.Strnumeroaviso "
    strsql = strsql & "From "
    strsql = strsql & gstrMovimentoBancario & " MB, "
    strsql = strsql & gstrLancamentoValor & " LV, "
    strsql = strsql & gstrLancamentoAlfa & " LA "
    strsql = strsql & "Where "
    strsql = strsql & "LV.Pkid = MB.INTLANCAMENTOVALOR AND "
    strsql = strsql & "LA.Pkid = LV.INTLANCAMENTOALFA AND "
    strsql = strsql & "MB.Pkid = " & tdb_MovimentoBancario.Columns("PKID").Value
    strsql = strsql & " Order By LA.Strnumeroaviso"
    
    LeDaTabelaParaObj "", dbc_intNumeroAviso, strsql
    
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "LV.Pkid, "
    strsql = strsql & "LV.Intparcela "
    strsql = strsql & "From "
    strsql = strsql & gstrMovimentoBancario & " MB, "
    strsql = strsql & gstrLancamentoValor & " LV "
    strsql = strsql & "Where "
    strsql = strsql & "LV.Pkid = MB.INTLANCAMENTOVALOR AND "
    strsql = strsql & "MB.Pkid = " & tdb_MovimentoBancario.Columns("PKID").Value
    strsql = strsql & " Order By LV.Intparcela"
    
    LeDaTabelaParaObj "", dbcintlancamentovalor, strsql
    
    txtintDigito = strDigito
    
End Sub

Sub VerificaValores()
    If Trim(txtdblMulta) = "" Then
        txtdblMulta = 0
    End If
    If Trim(txtdblJuros) = "" Then
        txtdblJuros = 0
    End If
    If Trim(txtdblCorrecao) = "" Then
        txtdblCorrecao = 0
    End If
    
    If Val(gstrConvVrParaSql(txt_dblTotal)) = Val(gstrConvVrParaSql(txtdblCorreto)) * -1 Then
        txtdblCorreto = txtdblCorreto * -1
    End If
    
End Sub

Private Function strQueryContaDescricao() As String
    Dim strsql As String

    strsql = strsql & "SELECT "
    strsql = strsql & "CB.Pkid, "
    strsql = strsql & "CB.strdescricao " & strCONCAT & "'('" & strCONCAT & " CB.strConta " & strCONCAT & " CB.strdigitoverificador" & strCONCAT & "')' strdescricao"
    strsql = strsql & " FROM " & gstrContaBancaria & " CB, "
    strsql = strsql & gstrPlanoConta & " PC"
    strsql = strsql & " Where"
    strsql = strsql & " CB.Pkid = PC.Intcontabancaria"
    strsql = strsql & " ORDER BY CB.strdescricao"
    
    strQueryContaDescricao = strsql

End Function

Private Sub CalculaReajuste()
Dim strsql       As String
Dim adoResultado As New ADODB.Recordset
Dim adoParcelas  As New ADODB.Recordset
    
    If Not gblnDataValida(txtdtmDtPagamento) Then
        txtdblCorreto = Space$(0)
        Exit Sub
    End If
    
    Set gobjBanco = New clsBanco
    
    strsql = "SELECT LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda " & _
             "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA " & _
             "WHERE LV.intLancamentoAlfa = LA.pkid " & _
             " AND LV.Pkid not in(SELECT Intlancamentovalor FROM " & gstrLancamentoPagamento & ") AND LA.Pkid = " & dbc_intNumeroAviso.BoundText & _
             " AND LV.intParcela = " & dbcintlancamentovalor.Text

    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            
        If Not adoResultado.EOF Then

            strsql = gstrStoredProcedure("sp_AtualizaParcela", dbc_intComposicaoDaReceita.BoundText & ", " & txt_intExercicio & ", " & dbcintlancamentovalor.Text & ", " & gstrConvDtParaSql(!Dtmdtvencimento) & ", " & gstrConvDtParaSql(txtdtmDtPagamento.Text) & ", " & gstrConvVrParaSql(!ValorOrig) & ", " & !intMoeda, True)
            If gobjBanco.CriaADO(strsql, 80, adoParcelas) Then
                txtdblPrincipal = Space$(0) & gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)
                txtdblMulta = Space$(0) & gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)
                txtdblJuros = Space$(0) & gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)
                txtdblCorrecao = Space$(0) & gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)
                txtdblCorreto = Space$(0) & gstrConvVrDoSql(Val(gstrConvVrParaSql(adoParcelas("dblValorPrincipal").Value)) + Val(gstrConvVrParaSql(adoParcelas("dblValorMulta").Value)) + Val(gstrConvVrParaSql(adoParcelas("dblValorJuros").Value)) + Val(gstrConvVrParaSql(adoParcelas("dblValorCorrecao").Value)))
                dblValorTotal
            Else
                txtdblPrincipal = Space$(0)
                txtdblMulta = Space$(0)
                txtdblJuros = Space$(0)
                txtdblCorrecao = Space$(0)
                txtdblCorreto = Space$(0)
            End If
        
        Else
            ExibeMensagem "Não foram encontrados lançamentos para esta Inscrição."
            Exit Sub
        End If
        
        End With
    End If
    
End Sub


Private Function ApagaMovimentoBancario() As Boolean

    Dim strsql          As String
    Dim strMensagem     As String
    
    Dim adoResultado    As ADODB.Recordset
    
    ApagaMovimentoBancario = False
    
    If Trim(txtPKId.Text) = "" Then
        ExibeMensagem "Não existem dados para excluir."
    Else
        
        strMensagem = "Confirma Exclusão do registro?"
    
        If gblnExclusaoGravacaoOk("E", strMensagem, True) Then
        
            strsql = ""
            strsql = strsql & "SELECT bitProcessado "
            strsql = strsql & "  FROM tblMovimentoBancario "
            strsql = strsql & " WHERE PKID = " & txtPKId.Text
            
            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strsql, 5, adoResultado
            
            If Not adoResultado.EOF Then
            
                If adoResultado!bitProcessado = 0 Then
                
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaBeginTrans
                    
                    ' Apaga primeiramente, os dados na tabela relacionada...
                    strsql = ""
                    strsql = strsql & "DELETE tblCriticaBaixa "
                    strsql = strsql & " WHERE intMovimentoBancario = " & txtPKId.Text
                
                    If gobjBanco.Execute(strsql) Then
                    
                        ' Apaga os dados na tabela...
                        strsql = ""
                        strsql = strsql & "DELETE tblMovimentoBancario "
                        strsql = strsql & " WHERE PKID = " & txtPKId.Text
                        
                        If gobjBanco.Execute(strsql) Then
                            Set gobjBanco = New clsBanco
                            gobjBanco.ExecutaCommitTrans
                            ApagaMovimentoBancario = True
                        Else
                            Set gobjBanco = New clsBanco
                            gobjBanco.ExecutaRollbackTrans
                        End If
                        
                    Else
                        Set gobjBanco = New clsBanco
                        gobjBanco.ExecutaRollbackTrans
                    End If
                    
                Else
                    ExibeMensagem "O registro não pode ser excluído pois já foi processada a baixa definitiva."
                End If
                
            Else
            
                ExibeMensagem "Não foi possível excluir o registro."
            
            End If
                
        End If
        
    End If

End Function
