VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadAcordos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acordos"
   ClientHeight    =   8880
   ClientLeft      =   1260
   ClientTop       =   2040
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6945
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   12250
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Acordos"
      TabPicture(0)   =   "frmCadAcordos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNumero"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblData"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLeiDecreto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblValor"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAcrescimos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblExpressasEm"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblstrDescricao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblRequerimento"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblRequerente"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblstrRG"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblstrCNPJCPF"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl_dblVlIndexador"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl_strIndexador"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl_Indexador"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl_anistia"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbldtmDtCancelamento"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbldtmDtUtilizacao"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "dbc_intDescProvisorios"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "mskstrInscricao"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "tdb_Parcelas"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "tab_3dEnderecos"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "mskstrCNPJCPF"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "dbcstrNomeProprietario"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "dbcintMoeda"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtdtmData"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtstrLeiDecreto"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtdblValor"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtdblAcrescimos"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmd_intMoeda"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtbitDigitoProcesso"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtintExercicioProcesso"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtstrCodigoProcesso"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtintRequerimento"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtintExercicioRequerimento"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmd_Contribuinte"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtstrIdentidade"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtPkid"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "fra_Parcelamento"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtDblVlIndexador"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtdblTotalIndexador"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmd_Indexador"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "fra_Impressoes"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtintExercicio"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt_strIndexador"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt_strAnistia"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt_strAnistiaLegislacao"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtdtmDtCancelamento"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtdtmDtUtilizacao"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "Débitos"
      TabPicture(1)   =   "frmCadAcordos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Débitos"
      Tab(1).Control(1)=   "fra_Parcelas"
      Tab(1).Control(2)=   "txtintExercicioDebito(1)"
      Tab(1).Control(3)=   "txtintExercicioDebito(0)"
      Tab(1).Control(4)=   "txtdtmDataDebito"
      Tab(1).Control(5)=   "txtstrLeiDecretoDebito"
      Tab(1).Control(6)=   "txtdblValorDebito"
      Tab(1).Control(7)=   "txtdblAcrescimosDebito"
      Tab(1).Control(8)=   "txtstrComposicaoDaReceita"
      Tab(1).Control(9)=   "mskstrInscricaoCadastralDebito"
      Tab(1).Control(10)=   "mskstrInscricaoDebito"
      Tab(1).Control(11)=   "dbcintMoedaDebito"
      Tab(1).Control(12)=   "lbl_Exercicio"
      Tab(1).Control(13)=   "Label6"
      Tab(1).Control(14)=   "Label5"
      Tab(1).Control(15)=   "Label4"
      Tab(1).Control(16)=   "Label3"
      Tab(1).Control(17)=   "Label2"
      Tab(1).Control(18)=   "Label1"
      Tab(1).Control(19)=   "lbl_strInscricaoAnterior(0)"
      Tab(1).Control(20)=   "lblintComposicao"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Acerto de Receitas"
      TabPicture(2)   =   "frmCadAcordos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmd_Finalizar"
      Tab(2).Control(1)=   "cmd_Executar"
      Tab(2).Control(2)=   "prg_Status"
      Tab(2).Control(3)=   "txt_datafinalprov"
      Tab(2).Control(4)=   "txt_datainicialprov"
      Tab(2).Control(5)=   "dbc_intAcordo"
      Tab(2).Control(6)=   "Label8"
      Tab(2).Control(7)=   "lbl_status"
      Tab(2).Control(8)=   "lbl_datafim"
      Tab(2).Control(9)=   "lbl_datainicio"
      Tab(2).ControlCount=   10
      Begin VB.TextBox txtdtmDtUtilizacao 
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
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1110
         Width           =   1275
      End
      Begin VB.TextBox txtdtmDtCancelamento 
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1110
         Width           =   1245
      End
      Begin VB.CommandButton cmd_Finalizar 
         Caption         =   "Finalizar"
         Height          =   495
         Left            =   -71880
         TabIndex        =   119
         Top             =   2070
         Width           =   1725
      End
      Begin VB.CommandButton cmd_Executar 
         Caption         =   "Executar"
         Height          =   495
         Left            =   -74700
         TabIndex        =   118
         Top             =   2070
         Width           =   1725
      End
      Begin MSComctlLib.ProgressBar prg_Status 
         Height          =   345
         Left            =   -74700
         TabIndex        =   117
         Top             =   1650
         Visible         =   0   'False
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.TextBox txt_datafinalprov 
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
         Left            =   -71430
         MaxLength       =   10
         TabIndex        =   114
         Top             =   570
         Width           =   1245
      End
      Begin VB.TextBox txt_datainicialprov 
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
         Left            =   -73770
         MaxLength       =   10
         TabIndex        =   112
         Top             =   570
         Width           =   1245
      End
      Begin VB.TextBox txt_strAnistiaLegislacao 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   109
         Top             =   3540
         Width           =   4740
      End
      Begin VB.TextBox txt_strAnistia 
         Height          =   285
         Left            =   1200
         MaxLength       =   70
         TabIndex        =   108
         Top             =   3240
         Width           =   4740
      End
      Begin VB.TextBox txt_strIndexador 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   106
         Top             =   1440
         Width           =   1245
      End
      Begin VB.TextBox txtintExercicio 
         Height          =   285
         Left            =   2505
         MaxLength       =   4
         TabIndex        =   3
         Top             =   405
         Width           =   705
      End
      Begin VB.Frame fra_Impressoes 
         Caption         =   " Opções de impressão : "
         Height          =   555
         Left            =   6165
         TabIndex        =   45
         Top             =   3075
         Width           =   2970
         Begin VB.CheckBox chk_Carne 
            Caption         =   "Carnê"
            Height          =   195
            Left            =   1740
            TabIndex        =   47
            Top             =   255
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chk_Termo 
            Caption         =   "Termo"
            Height          =   195
            Left            =   600
            TabIndex        =   46
            Top             =   255
            Value           =   1  'Checked
            Width           =   975
         End
      End
      Begin VB.CommandButton cmd_Indexador 
         Height          =   315
         Left            =   2490
         Picture         =   "frmCadAcordos.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Ativa Cadastro de Indexador Econômico"
         Top             =   1440
         Width           =   360
      End
      Begin VB.TextBox txtdblTotalIndexador 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   6855
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         Top             =   1440
         Width           =   1875
      End
      Begin VB.TextBox txtDblVlIndexador 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4455
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Frame fra_Parcelamento 
         Height          =   1365
         Left            =   6150
         TabIndex        =   34
         Top             =   1695
         Visible         =   0   'False
         Width           =   2985
         Begin VB.TextBox txt_dblValorParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1530
            TabIndex        =   40
            Top             =   960
            Width           =   1245
         End
         Begin VB.TextBox txt_dtmVencimento 
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
            Left            =   1515
            MaxLength       =   10
            TabIndex        =   38
            Top             =   585
            Width           =   1245
         End
         Begin VB.TextBox txt_QtdeParcelas 
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
            Left            =   1515
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   210
            Width           =   675
         End
         Begin VB.Label lbl_dblValorParcela 
            AutoSize        =   -1  'True
            Caption         =   "Valor da Parcela"
            Height          =   195
            Left            =   270
            TabIndex        =   39
            Top             =   1020
            Width           =   1170
         End
         Begin VB.Label lbl_dtmVencimento 
            AutoSize        =   -1  'True
            Caption         =   "1º Vencimento"
            Height          =   195
            Left            =   390
            TabIndex        =   37
            Top             =   645
            Width           =   1035
         End
         Begin VB.Label lbl_QtdeParcelas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Qtde Parcelas"
            Height          =   195
            Left            =   420
            TabIndex        =   35
            Top             =   285
            Width           =   1005
         End
      End
      Begin VB.Frame fra_Débitos 
         Caption         =   "Débitos"
         Height          =   1815
         Left            =   -74910
         TabIndex        =   100
         Top             =   1680
         Width           =   9015
         Begin TrueOleDBGrid70.TDBGrid tdb_Debitos 
            Height          =   1560
            Left            =   60
            TabIndex        =   101
            Top             =   180
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   2752
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
            Columns(1).Caption=   "Inscrição Cadastral"
            Columns(1).DataField=   "Inscricao"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Composição da Receita"
            Columns(2).DataField=   "ComposicaoDaReceita"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Numero do Aviso"
            Columns(3).DataField=   "strNumeroAviso"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Exercício"
            Columns(4).DataField=   "Exercicio"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "intUtilizacao"
            Columns(5).DataField=   "intUtilizacao"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=3440"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3360"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=8070"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=7990"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2355"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2275"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1588"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1508"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5).AllowSizing=0"
            Splits(0)._ColumnProps(36)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(30)  =   ":id=18,.fgcolor=&H8000000E&"
            _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(33)  =   ":id=19,.fgcolor=&H8000000E&"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(62)  =   "Named:id=33:Normal"
            _StyleDefs(63)  =   ":id=33,.parent=0"
            _StyleDefs(64)  =   "Named:id=34:Heading"
            _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(66)  =   ":id=34,.wraptext=-1"
            _StyleDefs(67)  =   "Named:id=35:Footing"
            _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   "Named:id=36:Selected"
            _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=37:Caption"
            _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(73)  =   "Named:id=38:HighlightRow"
            _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=39:EvenRow"
            _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(77)  =   "Named:id=40:OddRow"
            _StyleDefs(78)  =   ":id=40,.parent=33"
            _StyleDefs(79)  =   "Named:id=41:RecordSelector"
            _StyleDefs(80)  =   ":id=41,.parent=34"
            _StyleDefs(81)  =   "Named:id=42:FilterBar"
            _StyleDefs(82)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_Parcelas 
         Caption         =   "Parcelas"
         Height          =   2445
         Left            =   -74910
         TabIndex        =   102
         Top             =   3525
         Width           =   9015
         Begin TrueOleDBGrid70.TDBGrid tdb_ParcelasDebitos 
            Height          =   2190
            Left            =   75
            TabIndex        =   103
            Top             =   180
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   3863
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
            Columns(1).DataField=   "Parcela"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Vencimento"
            Columns(2).DataField=   "Vencimento"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Valor"
            Columns(3).DataField=   "Valor"
            Columns(3).NumberFormat=   "FormatText Event"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Juros"
            Columns(4).DataField=   "Juros"
            Columns(4).NumberFormat=   "FormatText Event"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Multa"
            Columns(5).DataField=   "Multa"
            Columns(5).NumberFormat=   "FormatText Event"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Correção"
            Columns(6).DataField=   "Correcao"
            Columns(6).NumberFormat=   "FormatText Event"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1296"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1217"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=1"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2011"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1931"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=1"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2064"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1984"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2249"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2170"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2223"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2143"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=2249"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2170"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(20)  =   ":id=8,.fgcolor=&H8000000E&"
            _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
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
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(65)  =   "Named:id=33:Normal"
            _StyleDefs(66)  =   ":id=33,.parent=0"
            _StyleDefs(67)  =   "Named:id=34:Heading"
            _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   ":id=34,.wraptext=-1"
            _StyleDefs(70)  =   "Named:id=35:Footing"
            _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   "Named:id=36:Selected"
            _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(74)  =   "Named:id=37:Caption"
            _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(76)  =   "Named:id=38:HighlightRow"
            _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(78)  =   "Named:id=39:EvenRow"
            _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(80)  =   "Named:id=40:OddRow"
            _StyleDefs(81)  =   ":id=40,.parent=33"
            _StyleDefs(82)  =   "Named:id=41:RecordSelector"
            _StyleDefs(83)  =   ":id=41,.parent=34"
            _StyleDefs(84)  =   "Named:id=42:FilterBar"
            _StyleDefs(85)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.TextBox txtPkid 
         Height          =   315
         Left            =   8130
         TabIndex        =   81
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtintExercicioDebito 
         Height          =   315
         Index           =   1
         Left            =   -66600
         MaxLength       =   4
         TabIndex        =   80
         Top             =   1260
         Width           =   705
      End
      Begin VB.TextBox txtintExercicioDebito 
         Height          =   285
         Index           =   0
         Left            =   -72645
         MaxLength       =   4
         TabIndex        =   72
         Top             =   525
         Width           =   705
      End
      Begin VB.TextBox txtdtmDataDebito 
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
         Left            =   -70965
         MaxLength       =   10
         TabIndex        =   73
         Top             =   525
         Width           =   1245
      End
      Begin VB.TextBox txtstrLeiDecretoDebito 
         Height          =   285
         Left            =   -68145
         MaxLength       =   20
         TabIndex        =   74
         Top             =   525
         Width           =   2250
      End
      Begin VB.TextBox txtdblValorDebito 
         Height          =   285
         Left            =   -73935
         TabIndex        =   75
         Top             =   885
         Width           =   1245
      End
      Begin VB.TextBox txtdblAcrescimosDebito 
         Height          =   285
         Left            =   -70965
         TabIndex        =   76
         Top             =   885
         Width           =   1245
      End
      Begin VB.TextBox txtstrComposicaoDaReceita 
         Height          =   285
         Left            =   -70965
         TabIndex        =   79
         Top             =   1260
         Width           =   3150
      End
      Begin VB.TextBox txtstrIdentidade 
         Height          =   285
         Left            =   4215
         MaxLength       =   20
         TabIndex        =   44
         Top             =   2535
         Width           =   1275
      End
      Begin VB.CommandButton cmd_Contribuinte 
         Height          =   315
         Left            =   5655
         Picture         =   "frmCadAcordos.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Ativa Cadastro de Contribuintes"
         Top             =   2145
         Width           =   360
      End
      Begin VB.TextBox txtintExercicioRequerimento 
         Height          =   285
         Left            =   4980
         MaxLength       =   4
         TabIndex        =   30
         Top             =   1800
         Width           =   705
      End
      Begin VB.TextBox txtintRequerimento 
         Height          =   285
         Left            =   4035
         MaxLength       =   8
         TabIndex        =   29
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtstrCodigoProcesso 
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
         Left            =   1200
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   1800
         Width           =   825
      End
      Begin VB.TextBox txtintExercicioProcesso 
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
         Left            =   2070
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1800
         Width           =   465
      End
      Begin VB.TextBox txtbitDigitoProcesso 
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
         Left            =   2580
         MaxLength       =   2
         TabIndex        =   27
         Top             =   1800
         Width           =   285
      End
      Begin VB.CommandButton cmd_intMoeda 
         Height          =   300
         Left            =   8745
         Picture         =   "frmCadAcordos.frx":0290
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Moeda"
         Top             =   750
         Width           =   360
      End
      Begin VB.TextBox txtdblAcrescimos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Top             =   765
         Width           =   1275
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   765
         Width           =   1245
      End
      Begin VB.TextBox txtstrLeiDecreto 
         Height          =   285
         Left            =   6855
         MaxLength       =   20
         TabIndex        =   7
         Top             =   405
         Width           =   2250
      End
      Begin VB.TextBox txtdtmData 
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
         Left            =   4455
         MaxLength       =   10
         TabIndex        =   5
         Top             =   405
         Width           =   1245
      End
      Begin MSDataListLib.DataCombo dbcintMoeda 
         Height          =   315
         Left            =   6975
         TabIndex        =   13
         Top             =   735
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcstrNomeProprietario 
         Height          =   315
         Left            =   1200
         TabIndex        =   32
         Top             =   2145
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSMask.MaskEdBox mskstrCNPJCPF 
         Height          =   300
         Left            =   1200
         TabIndex        =   42
         Top             =   2535
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "0"
         PromptChar      =   " "
      End
      Begin TabDlg.SSTab tab_3dEnderecos 
         Height          =   1350
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3870
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   2381
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Endereço"
         TabPicture(0)   =   "frmCadAcordos.frx":03AE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_EndImobiliario"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Endereço de Notificação"
         TabPicture(1)   =   "frmCadAcordos.frx":03CA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).ControlCount=   1
         Begin VB.Frame fra_EndImobiliario 
            Height          =   930
            Left            =   150
            TabIndex        =   49
            Top             =   315
            Width           =   8655
            Begin VB.TextBox txtstrUf 
               Height          =   300
               Left            =   6765
               MaxLength       =   2
               TabIndex        =   61
               Top             =   540
               Width           =   375
            End
            Begin VB.TextBox txtstrMunicipio 
               Height          =   300
               Left            =   4185
               MaxLength       =   50
               TabIndex        =   59
               Top             =   540
               Width           =   2235
            End
            Begin VB.TextBox txtstrNumero 
               Height          =   300
               Left            =   5475
               MaxLength       =   10
               TabIndex        =   53
               Top             =   180
               Width           =   825
            End
            Begin VB.TextBox txtintCep 
               Height          =   300
               Left            =   7560
               MaxLength       =   9
               TabIndex        =   63
               Top             =   525
               Width           =   1005
            End
            Begin VB.TextBox txtstrComplemento 
               Height          =   300
               Left            =   6960
               MaxLength       =   10
               TabIndex        =   55
               Top             =   180
               Width           =   1590
            End
            Begin VB.TextBox txtstrLogradouro 
               Height          =   300
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   51
               Top             =   180
               Width           =   4065
            End
            Begin VB.TextBox txtstrBairro 
               Height          =   300
               Left            =   675
               MaxLength       =   50
               TabIndex        =   57
               Top             =   540
               Width           =   2670
            End
            Begin VB.Label lblstrUf 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   6495
               TabIndex        =   60
               Top             =   630
               Width           =   210
            End
            Begin VB.Label lblstrMunicipio 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   3435
               TabIndex        =   58
               Top             =   615
               Width           =   705
            End
            Begin VB.Label lblintCep 
               AutoSize        =   -1  'True
               Caption         =   "CEP"
               Height          =   195
               Left            =   7200
               TabIndex        =   62
               Top             =   615
               Width           =   315
            End
            Begin VB.Label lblstrComplemento 
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6435
               TabIndex        =   54
               Top             =   270
               Width           =   480
            End
            Begin VB.Label lblintNumero 
               AutoSize        =   -1  'True
               Caption         =   "N°"
               Height          =   195
               Left            =   5250
               TabIndex        =   52
               Top             =   270
               Width           =   180
            End
            Begin VB.Label lblintBairro 
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   210
               TabIndex        =   56
               Top             =   630
               Width           =   405
            End
            Begin VB.Label lblintLogradouro 
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   195
               TabIndex        =   50
               Top             =   270
               Width           =   810
            End
         End
         Begin VB.Frame Frame1 
            Height          =   930
            Left            =   -74850
            TabIndex        =   82
            Top             =   315
            Width           =   8655
            Begin VB.TextBox txtstrBairroC 
               Height          =   300
               Left            =   675
               MaxLength       =   50
               TabIndex        =   67
               Top             =   540
               Width           =   2670
            End
            Begin VB.TextBox txtstrLogradouroC 
               Height          =   300
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   64
               Top             =   180
               Width           =   4065
            End
            Begin VB.TextBox txtstrComplementoC 
               Height          =   300
               Left            =   6960
               MaxLength       =   10
               TabIndex        =   66
               Top             =   180
               Width           =   1590
            End
            Begin VB.TextBox txtintCepC 
               Height          =   300
               Left            =   7560
               MaxLength       =   9
               TabIndex        =   70
               Top             =   525
               Width           =   1005
            End
            Begin VB.TextBox txtstrNumeroC 
               Height          =   300
               Left            =   5475
               MaxLength       =   10
               TabIndex        =   65
               Top             =   180
               Width           =   825
            End
            Begin VB.TextBox txtstrUFC 
               Height          =   300
               Left            =   6765
               MaxLength       =   2
               TabIndex        =   69
               Top             =   540
               Width           =   375
            End
            Begin VB.TextBox txtstrMunicipioC 
               Height          =   300
               Left            =   4185
               MaxLength       =   50
               TabIndex        =   68
               Top             =   540
               Width           =   2235
            End
            Begin VB.Label lbl_LogradouroC 
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   195
               TabIndex        =   83
               Top             =   270
               Width           =   810
            End
            Begin VB.Label lbl_BairroC 
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   210
               TabIndex        =   86
               Top             =   630
               Width           =   405
            End
            Begin VB.Label lbl_NumeroC 
               AutoSize        =   -1  'True
               Caption         =   "N°"
               Height          =   195
               Left            =   5250
               TabIndex        =   84
               Top             =   270
               Width           =   180
            End
            Begin VB.Label lbl_ComplementoC 
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6435
               TabIndex        =   85
               Top             =   270
               Width           =   480
            End
            Begin VB.Label lbl_CepC 
               AutoSize        =   -1  'True
               Caption         =   "CEP"
               Height          =   195
               Left            =   7200
               TabIndex        =   89
               Top             =   615
               Width           =   315
            End
            Begin VB.Label lbl_UFC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   6495
               TabIndex        =   88
               Top             =   630
               Width           =   210
            End
            Begin VB.Label lbl_MunicipioC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   3435
               TabIndex        =   87
               Top             =   615
               Width           =   705
            End
         End
      End
      Begin MSMask.MaskEdBox mskstrInscricaoCadastralDebito 
         Height          =   285
         Left            =   -73380
         TabIndex        =   78
         Top             =   1260
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskstrInscricaoDebito 
         Height          =   285
         Left            =   -73935
         TabIndex        =   71
         Top             =   525
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo dbcintMoedaDebito 
         Height          =   315
         Left            =   -68145
         TabIndex        =   77
         Top             =   900
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Parcelas 
         Height          =   1620
         Left            =   135
         TabIndex        =   90
         Top             =   5235
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   2858
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
         Columns(1).DataField=   "Parcela"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Vencimento"
         Columns(2).DataField=   "Vencimento"
         Columns(2).NumberFormat=   "General Date"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Valor"
         Columns(3).DataField=   "Valor"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Baixa"
         Columns(4).DataField=   "Baixa"
         Columns(4).NumberFormat=   "General Date"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Descrição"
         Columns(5).DataField=   "Descricao"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Observação"
         Columns(6).DataField=   "Observacao"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=1852"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1773"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=1958"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1879"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=2328"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2249"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=6826"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=6747"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=6033"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=5953"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=94,.parent=13"
         _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=91,.parent=14"
         _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=92,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=93,.parent=17"
         _StyleDefs(67)  =   "Named:id=33:Normal"
         _StyleDefs(68)  =   ":id=33,.parent=0"
         _StyleDefs(69)  =   "Named:id=34:Heading"
         _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   ":id=34,.wraptext=-1"
         _StyleDefs(72)  =   "Named:id=35:Footing"
         _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(74)  =   "Named:id=36:Selected"
         _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(76)  =   "Named:id=37:Caption"
         _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(78)  =   "Named:id=38:HighlightRow"
         _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   "Named:id=39:EvenRow"
         _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(82)  =   "Named:id=40:OddRow"
         _StyleDefs(83)  =   ":id=40,.parent=33"
         _StyleDefs(84)  =   "Named:id=41:RecordSelector"
         _StyleDefs(85)  =   ":id=41,.parent=34"
         _StyleDefs(86)  =   "Named:id=42:FilterBar"
         _StyleDefs(87)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox mskstrInscricao 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   405
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo dbc_intDescProvisorios 
         Height          =   315
         Left            =   1200
         TabIndex        =   107
         Top             =   2880
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intAcordo 
         Height          =   315
         Left            =   -73770
         TabIndex        =   120
         Top             =   990
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Acordo"
         Height          =   195
         Left            =   -74370
         TabIndex        =   121
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label lbldtmDtUtilizacao 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   3690
         TabIndex        =   17
         Top             =   1185
         Width           =   690
      End
      Begin VB.Label lbldtmDtCancelamento 
         AutoSize        =   -1  'True
         Caption         =   "Cancelamento"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   1185
         Width           =   1020
      End
      Begin VB.Label lbl_status 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74640
         TabIndex        =   116
         Top             =   1320
         Visible         =   0   'False
         Width           =   4425
      End
      Begin VB.Label lbl_datafim 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   -72330
         TabIndex        =   115
         Top             =   645
         Width           =   795
      End
      Begin VB.Label lbl_datainicio 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   -74670
         TabIndex        =   113
         Top             =   645
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Legislação"
         Height          =   195
         Left            =   405
         TabIndex        =   111
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lbl_anistia 
         AutoSize        =   -1  'True
         Caption         =   "Anistia"
         Height          =   195
         Left            =   705
         TabIndex        =   110
         Top             =   3300
         Width           =   465
      End
      Begin VB.Label lbl_Indexador 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   6765
         TabIndex        =   105
         Top             =   1185
         Width           =   45
      End
      Begin VB.Label lbl_strIndexador 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   195
         Left            =   405
         TabIndex        =   19
         Top             =   1515
         Width           =   705
      End
      Begin VB.Label lbl_dblVlIndexador 
         AutoSize        =   -1  'True
         Caption         =   "Valor Indexador"
         Height          =   195
         Left            =   3300
         TabIndex        =   21
         Top             =   1515
         Width           =   1110
      End
      Begin VB.Label lbl_Exercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   -67350
         TabIndex        =   99
         Top             =   1350
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   -74595
         TabIndex        =   91
         Top             =   615
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   -71400
         TabIndex        =   92
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lei/Decreto"
         Height          =   195
         Left            =   -69045
         TabIndex        =   93
         Top             =   615
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -74400
         TabIndex        =   94
         Top             =   975
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Acréscimos"
         Height          =   195
         Left            =   -71850
         TabIndex        =   95
         Top             =   975
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Expressos em"
         Height          =   195
         Left            =   -69165
         TabIndex        =   96
         Top             =   975
         Width           =   975
      End
      Begin VB.Label lbl_strInscricaoAnterior 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Index           =   0
         Left            =   -74835
         TabIndex        =   97
         Top             =   1350
         Width           =   1350
      End
      Begin VB.Label lblintComposicao 
         AutoSize        =   -1  'True
         Caption         =   "Composição"
         Height          =   195
         Left            =   -71910
         TabIndex        =   98
         Top             =   1350
         Width           =   870
      End
      Begin VB.Label lblstrCNPJCPF 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ / CPF"
         Height          =   195
         Left            =   300
         TabIndex        =   41
         Top             =   2610
         Width           =   870
      End
      Begin VB.Label lblstrRG 
         AutoSize        =   -1  'True
         Caption         =   "RG"
         Height          =   195
         Left            =   3900
         TabIndex        =   43
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label lblRequerente 
         AutoSize        =   -1  'True
         Caption         =   "Requerente"
         Height          =   195
         Left            =   330
         TabIndex        =   31
         Top             =   2235
         Width           =   840
      End
      Begin VB.Label lblRequerimento 
         AutoSize        =   -1  'True
         Caption         =   "Requerimento"
         Height          =   195
         Left            =   2970
         TabIndex        =   28
         Top             =   1890
         Width           =   990
      End
      Begin VB.Label lblstrDescricao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Processo"
         Height          =   195
         Left            =   510
         TabIndex        =   24
         Top             =   1890
         Width           =   660
      End
      Begin VB.Label lblExpressasEm 
         AutoSize        =   -1  'True
         Caption         =   "Expressos em"
         Height          =   195
         Left            =   5955
         TabIndex        =   12
         Top             =   825
         Width           =   975
      End
      Begin VB.Label lblAcrescimos 
         AutoSize        =   -1  'True
         Caption         =   "Acréscimos"
         Height          =   195
         Left            =   3570
         TabIndex        =   10
         Top             =   855
         Width           =   810
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   810
         TabIndex        =   8
         Top             =   855
         Width           =   360
      End
      Begin VB.Label lblLeiDecreto 
         AutoSize        =   -1  'True
         Caption         =   "Lei/Decreto"
         Height          =   195
         Left            =   5955
         TabIndex        =   6
         Top             =   510
         Width           =   855
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   4035
         TabIndex        =   4
         Top             =   510
         Width           =   345
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   615
         TabIndex        =   1
         Top             =   495
         Width           =   555
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Acordos 
      Height          =   1845
      Left            =   30
      TabIndex        =   104
      Top             =   6990
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   3254
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
      Columns(1).Caption=   "Número"
      Columns(1).DataField=   "Numero"
      Columns(1).NumberFormat=   "FormatText Event"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Data"
      Columns(2).DataField=   "Data"
      Columns(2).NumberFormat=   "General Date"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Lei/Decreto"
      Columns(3).DataField=   "LeiDecreto"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Número do Aviso"
      Columns(4).DataField=   "strNumeroAviso"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Valor"
      Columns(5).DataField=   "Valor"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Acréscimos"
      Columns(6).DataField=   "Acrescimos"
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Expressas Em"
      Columns(7).DataField=   "Moeda"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Processo"
      Columns(8).DataField=   "Processo"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Requerimento"
      Columns(9).DataField=   "Requerimento"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Requerente"
      Columns(10).DataField=   "NomeProprietario"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "CNPJ/CPF"
      Columns(11).DataField=   "CNPJCPF"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "RG"
      Columns(12).DataField=   "Identidade"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "intLancamentoAlfa"
      Columns(13).DataField=   "intLancamentoAlfa"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
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
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2275"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2196"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1931"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1852"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=4842"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4763"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=2487"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2408"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=2170"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2090"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=2328"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2249"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=2037"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1958"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(50)=   "Column(8).Width=2355"
      Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=2275"
      Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=1"
      Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(56)=   "Column(9).Width=2461"
      Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=2381"
      Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=1"
      Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(62)=   "Column(10).Width=7250"
      Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=7170"
      Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=0"
      Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(68)=   "Column(11).Width=3307"
      Splits(0)._ColumnProps(69)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(11)._WidthInPix=3228"
      Splits(0)._ColumnProps(71)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(72)=   "Column(11)._ColStyle=1"
      Splits(0)._ColumnProps(73)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(74)=   "Column(12).Width=2937"
      Splits(0)._ColumnProps(75)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(12)._WidthInPix=2858"
      Splits(0)._ColumnProps(77)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(78)=   "Column(12)._ColStyle=1"
      Splits(0)._ColumnProps(79)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(80)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(81)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(82)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(83)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(84)=   "Column(13).AllowSizing=0"
      Splits(0)._ColumnProps(85)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(86)=   "Column(13).Order=14"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(20)  =   ":id=8,.fgcolor=&H8000000E&"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(31)  =   ":id=18,.fgcolor=&H8000000E&"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(34)  =   ":id=19,.fgcolor=&H8000000E&"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=78,.parent=13,.alignment=1"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=94,.parent=13,.alignment=2"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=91,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=92,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=93,.parent=17"
      _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(75)  =   "Splits(0).Columns(9).Style:id=102,.parent=13,.alignment=2"
      _StyleDefs(76)  =   "Splits(0).Columns(9).HeadingStyle:id=99,.parent=14"
      _StyleDefs(77)  =   "Splits(0).Columns(9).FooterStyle:id=100,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(9).EditorStyle:id=101,.parent=17"
      _StyleDefs(79)  =   "Splits(0).Columns(10).Style:id=66,.parent=13,.alignment=0"
      _StyleDefs(80)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
      _StyleDefs(81)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
      _StyleDefs(83)  =   "Splits(0).Columns(11).Style:id=98,.parent=13,.alignment=2"
      _StyleDefs(84)  =   "Splits(0).Columns(11).HeadingStyle:id=95,.parent=14"
      _StyleDefs(85)  =   "Splits(0).Columns(11).FooterStyle:id=96,.parent=15"
      _StyleDefs(86)  =   "Splits(0).Columns(11).EditorStyle:id=97,.parent=17"
      _StyleDefs(87)  =   "Splits(0).Columns(12).Style:id=106,.parent=13,.alignment=2"
      _StyleDefs(88)  =   "Splits(0).Columns(12).HeadingStyle:id=103,.parent=14"
      _StyleDefs(89)  =   "Splits(0).Columns(12).FooterStyle:id=104,.parent=15"
      _StyleDefs(90)  =   "Splits(0).Columns(12).EditorStyle:id=105,.parent=17"
      _StyleDefs(91)  =   "Splits(0).Columns(13).Style:id=70,.parent=13"
      _StyleDefs(92)  =   "Splits(0).Columns(13).HeadingStyle:id=67,.parent=14"
      _StyleDefs(93)  =   "Splits(0).Columns(13).FooterStyle:id=68,.parent=15"
      _StyleDefs(94)  =   "Splits(0).Columns(13).EditorStyle:id=69,.parent=17"
      _StyleDefs(95)  =   "Named:id=33:Normal"
      _StyleDefs(96)  =   ":id=33,.parent=0"
      _StyleDefs(97)  =   "Named:id=34:Heading"
      _StyleDefs(98)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(99)  =   ":id=34,.wraptext=-1"
      _StyleDefs(100) =   "Named:id=35:Footing"
      _StyleDefs(101) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(102) =   "Named:id=36:Selected"
      _StyleDefs(103) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(104) =   "Named:id=37:Caption"
      _StyleDefs(105) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(106) =   "Named:id=38:HighlightRow"
      _StyleDefs(107) =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(108) =   "Named:id=39:EvenRow"
      _StyleDefs(109) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(110) =   "Named:id=40:OddRow"
      _StyleDefs(111) =   ":id=40,.parent=33"
      _StyleDefs(112) =   "Named:id=41:RecordSelector"
      _StyleDefs(113) =   ":id=41,.parent=34"
      _StyleDefs(114) =   "Named:id=42:FilterBar"
      _StyleDefs(115) =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadAcordos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnPrimeiraVez           As Boolean
Dim mblnAlterando            As Boolean
Dim blnPrimeiraVezDebito     As Boolean

Dim strInscricaoAtual        As String

'Armazena o valor original do acordo, para podermos fazer calculo de acrscimo em cima dele
Public dblValorAcordoOriginal   As Double

Public dblValorHonorarios       As Double

'Verifica se o acordo foi gerado pela Atualização de Débitos
Dim blnAtualizacao           As Boolean

'Array que armazena parcelas para geracao do acordo
Dim vetParcelasParaAcordo()  As String
'Array que armazena parcelas para geracao do acordo mas para atualização se for anistia
Dim vetParcelasParaAcordoAux()  As String

'Vamos armazenar o acrescimo de desctos providorios
Dim dblAcrescimoDescProv         As Double
Dim intAcrescimoParcIniDescProv  As Integer
Dim intQtdeParcelasAcrescimo     As Integer

'Cosntantes referentes a cada coluna do array de parcelas para acordo
Const PKID_LANCAMENTO_VALOR  As Byte = 0
Const VALOR_ORIGINAL         As Byte = 1
Const VALOR_PRINCIPAL        As Byte = 2
Const VALOR_MULTA            As Byte = 3
Const VALOR_JUROS            As Byte = 4
Const VALOR_CORRECAO         As Byte = 5
Const VALOR_TOTAL            As Byte = 6
Const NUMERO_INSCRICAO       As Byte = 7
Const NUMERO_AVISO           As Byte = 8
Const DATA_VENCIMENTO        As Byte = 9
Const EXERCICIO              As Byte = 10
Const PARCELA                As Byte = 11
Const COMPOSICAO_RECEITA     As Byte = 12
Const NUMERO_INSCRICAO_PURA  As Byte = 13
Const Utilizacao             As Byte = 14
Const EXECUTIVO              As Byte = 15

'Array que armazena parcelas para geracao do acordo
Dim vetReceitas()            As String
'Cosntantes referentes a cada coluna do array de receitas
Const RECEITA                As Byte = 0
Const VALOR_RECEITA          As Byte = 1

Public blnParcelamentoDebito As Boolean 'TRUE quando o form for chamado de Atualizacao de Debitos
Public blnVBModal            As Boolean
Dim blnAutoNumeracao         As Boolean

Private Sub cmd_Contribuinte_Click()
    blnVBModal = False
    ChamaFormCadastro frmCadContribuinte, dbcstrNomeProprietario
    DoEvents
    blnVBModal = True
End Sub

Private Sub cmd_Executar_Click()
    
    'Acerto de receitas Provisorio
    If Not dbc_intAcordo.MatchedWithList Then
        If Len(txt_datainicialprov) = 0 Or Len(txt_datafinalprov) = 0 Then
            ExibeMensagem "Preencha as datas corretamente"
            Exit Sub
        End If
    End If
    
    cmd_Executar.Enabled = False
    'GravarAcordoProvisorio txt_datainicialprov, txt_datafinalprov
    GravarAcordoProvisorio2 txt_datainicialprov, txt_datafinalprov
    cmd_Executar.Enabled = True
    
End Sub

Private Sub cmd_Finalizar_Click()
    tab_3dPasta.Tab = 0
    tab_3dPasta.TabVisible(2) = False
End Sub

Private Sub cmd_indexador_Click()
    blnVBModal = False
    ChamaFormCadastro frmIndexadorEconomico, txt_strIndexador
    DoEvents
    blnVBModal = True
End Sub

Private Sub cmd_intMoeda_Click()
    blnVBModal = False
    ChamaFormCadastro frmCadMoedas, dbcintMoeda
    DoEvents
    blnVBModal = True
End Sub

Private Sub dbc_intAcordo_Click(Area As Integer)
    DropDownDataCombo dbc_intAcordo, Me, Area
End Sub

Private Sub dbc_intAcordo_GotFocus()
    MarcaCampo dbc_intAcordo
    dbc_intAcordo.Tag = "SELECT Pkid, strinscricao FROM tblLancamentoAlfa WHERE intUtilizacao = 4" & ";strInscricao"
End Sub

Private Sub dbc_intAcordo_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intAcordo, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intAcordo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intAcordo
End Sub

Private Sub dbc_intDescProvisorios_Change()
    Dim intFor          As Integer
    Dim dblvalorTot     As Double
    Dim dblValorTotHon  As Double
    
    If dbc_intDescProvisorios.MatchedWithList And Val(dbc_intDescProvisorios.BoundText) > 0 Then
        PreencheAnistia dbc_intDescProvisorios.BoundText
    Else
        txt_strAnistia = ""
        txt_strAnistiaLegislacao = ""
        vetParcelasParaAcordo = vetParcelasParaAcordoAux
        For intFor = 0 To UBound(vetParcelasParaAcordo, 2)
            dblvalorTot = dblvalorTot + vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor) + vetParcelasParaAcordo(VALOR_MULTA, intFor) + vetParcelasParaAcordo(VALOR_JUROS, intFor) + vetParcelasParaAcordo(VALOR_CORRECAO, intFor)
            'Vamos somar o total de parcelas com honorarios para aplicar o desconto nele tambem
            If vetParcelasParaAcordo(EXECUTIVO, intFor) = True Then
                dblValorTotHon = dblValorTotHon + vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor) + vetParcelasParaAcordo(VALOR_MULTA, intFor) + vetParcelasParaAcordo(VALOR_JUROS, intFor) + vetParcelasParaAcordo(VALOR_CORRECAO, intFor)
            End If
        Next
        
        'Vamos redefinir o valor dos honorarios com os valores ja com desconto
        dblValorHonorarios = dblCalculaEncargos(BIT_HONORARIOS, dblValorTotHon, "Acordo")

        'Vamos somar o Honorario no total
        dblvalorTot = dblvalorTot + dblValorHonorarios
        
        txtdblValor = gstrConvVrDoSql(dblvalorTot, , , True)
        
        'Vamos aplicar a anistia no valor original
        dblValorAcordoOriginal = dblvalorTot
        
        dblAcrescimoDescProv = 0
        dblAcrescimoDescProv = 0
        intAcrescimoParcIniDescProv = 0
        intQtdeParcelasAcrescimo = 0
        
        txt_QtdeParcelas = ""
        txt_dblValorParcela = ""
        txt_dtmVencimento = ""
    End If
End Sub

Private Sub dbc_intDescProvisorios_Click(Area As Integer)
    Dim intFor          As Integer
    Dim dblvalorTot     As Double
    
    If Area = 2 Then
        If dbc_intDescProvisorios.MatchedWithList And Val(dbc_intDescProvisorios.BoundText) > 0 Then
            PreencheAnistia dbc_intDescProvisorios.BoundText
        Else
            txt_strAnistia = ""
            txt_strAnistiaLegislacao = ""
            vetParcelasParaAcordo = vetParcelasParaAcordoAux
            
            For intFor = 0 To UBound(vetParcelasParaAcordo, 2)
                dblvalorTot = dblvalorTot + vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor) + vetParcelasParaAcordo(VALOR_MULTA, intFor) + vetParcelasParaAcordo(VALOR_JUROS, intFor) + vetParcelasParaAcordo(VALOR_CORRECAO, intFor)
            Next
            txtdblValor = gstrConvVrDoSql(dblvalorTot, , , True)
            
            dblAcrescimoDescProv = 0
            dblAcrescimoDescProv = 0
            intAcrescimoParcIniDescProv = 0
            intQtdeParcelasAcrescimo = 0
        
            txt_QtdeParcelas = ""
            txt_dblValorParcela = ""
            txt_dtmVencimento = ""
        End If
    End If
End Sub

Private Sub dbc_intDescProvisorios_GotFocus()
    MarcaCampo dbc_intDescProvisorios
End Sub

Private Sub dbc_intDescProvisorios_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intDescProvisorios
End Sub

Private Sub dbcintMoeda_Click(Area As Integer)
    DropDownDataCombo dbcintMoeda, Me, Area
End Sub

Private Sub dbcintMoeda_GotFocus()
    MarcaCampo dbcintMoeda
End Sub

Private Sub dbcintMoeda_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintMoeda, Me, , KeyCode, Shift
End Sub

Private Sub dbcintMoeda_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintMoeda
End Sub

Private Sub dbcstrNomeProprietario_Change()
    If dbcstrNomeProprietario.MatchedWithList Then
        PreencheDadosProprietario
        CarregaEndereco
    End If
End Sub

Private Sub dbcstrNomeProprietario_Click(Area As Integer)
    DropDownDataCombo dbcstrNomeProprietario, Me, Area
    If dbcstrNomeProprietario.MatchedWithList Then
        PreencheDadosProprietario
        CarregaEndereco
    End If
End Sub

Private Sub dbcstrNomeProprietario_GotFocus()
    MarcaCampo dbcstrNomeProprietario
End Sub

Private Sub dbcstrNomeProprietario_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcstrNomeProprietario, Me, , KeyCode, Shift
End Sub

Private Sub dbcstrNomeProprietario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcstrNomeProprietario
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1147
    blnAutoNumeracao = IIf(blnParcelamentoDebito, 1, blnVerificaAutoNumeracao)
    tab_3dPasta.TabVisible(1) = Not blnParcelamentoDebito
    tab_3dPasta.TabEnabled(1) = Not blnParcelamentoDebito
    fra_Parcelamento.Visible = blnParcelamentoDebito
    fra_Impressoes.Visible = blnParcelamentoDebito
    dbc_intDescProvisorios.Enabled = blnParcelamentoDebito
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrLocalizar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
    TrocaCorObjeto txtdtmDtCancelamento, blnParcelamentoDebito
    TrocaCorObjeto txtdtmDtUtilizacao, blnParcelamentoDebito
    
    If blnAtualizacao = False Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar
    Else
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo, gstrImprimir, gstrLocalizar
       mskstrInscricao.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrLocalizar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    If blnVBModal Then
       Me.SetFocus
    End If
End Sub

Private Sub Form_Load()
    gintCodSeguranca = 1147
    mblnAlterando = False
    blnParcelamentoDebito = False
    
    ReDim vetParcelasParaAcordo(15, 0)
    ReDim vetParcelasParaAcordoAux(15, 0)
    
    dbcintMoeda.Tag = strQueryMoedas & ";strAbreviatura"
    dbcintMoedaDebito.Tag = strQueryMoedas & ";strAbreviatura"
    dbcstrNomeProprietario.Tag = strQueryRequerente & ";strNome"
    dbc_intDescProvisorios.Tag = strDesctoProvisorio & ";strdescricao"
    
    TrocaCorObjeto txt_strAnistia, True
    TrocaCorObjeto txt_strAnistiaLegislacao, True
    TrocaCorObjeto mskstrInscricaoDebito, True
    TrocaCorObjeto txtintExercicioDebito(0), True
    TrocaCorObjeto txtdtmDataDebito, True
    TrocaCorObjeto txtstrLeiDecretoDebito, True
    TrocaCorObjeto txtdblValorDebito, True
    TrocaCorObjeto txtdblAcrescimosDebito, True
    TrocaCorObjeto dbcintMoedaDebito, True
    TrocaCorObjeto mskstrCNPJCPF, True
    TrocaCorObjeto txtstridentidade, True
    TrocaCorObjeto txt_dblValorParcela, True
    TrocaCorObjeto txt_strIndexador, True
    TrocaCorObjeto txtdblVlIndexador, True
        
    'Acerto de receitas Provisorio
    tab_3dPasta.TabVisible(2) = False
    
    'Verfica qual máscara usar
    VerificaMascaraInscricao
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrLocalizar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnVBModal = False
    blnAtualizacao = False
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    frmAtualizacaoDebitos.blnAcordoAtivo = False
End Sub

Private Sub lblLeiDecreto_DblClick()
    'Acerto de receitas provisorio
    tab_3dPasta.TabVisible(2) = True
End Sub

Private Sub mskstrInscricao_GotFocus()
    If Len(Trim(mskstrInscricao)) = 0 And blnAutoNumeracao = True Then
        Screen.MousePointer = vbArrowHourglass
        mskstrInscricao = ProximaInscricaoAcordo
        txtintExercicio = Year(gstrDataDoSistema)
        Screen.MousePointer = vbDefault
    End If
    tab_3dPasta.Tab = 0
    MarcaCampo mskstrInscricao
End Sub

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricao
End Sub

Private Sub mskstrInscricao_LostFocus()
'    If Len(Trim(mskstrInscricao)) > 0 Then
        'mskstrInscricao = Format(mskstrInscricao, "000.000")
'    End If
End Sub

Private Sub VerificaMascaraInscricao()
Dim strSql As String
Dim adoResultado As ADODB.Recordset
Dim strMascara   As String

    strMascara = ""
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & TYP_ACORDO
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

End Sub

Private Sub mskstrInscricaoCadastralDebito_GotFocus()
    MarcaCampo mskstrInscricaoCadastralDebito
End Sub

Private Sub tdb_Acordos_Click()
    blnPrimeiraVez = True
End Sub

Private Sub tdb_Acordos_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_Acordos
End Sub

Private Sub tdb_Acordos_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = tdb_Acordos.Columns("Número").ColIndex Then
        Value = gstrFormataInscricao(CStr(Value), TYP_ACORDO)
    End If
End Sub

Private Sub tdb_Acordos_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Acordos, ColIndex
End Sub

Private Sub tdb_Acordos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Not tdb_Acordos.EOF And blnPrimeiraVez Then
        lbl_Indexador.Caption = ""
        txtdblTotalIndexador.Text = ""
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrImprimir, gstrAplicar
        mblnAlterando = True
        txtPKId = tdb_Acordos.Columns("Pkid").Value
        PreecheCampos
        PreencheGrdParcelas Val(txtPKId)
        PreencheGrdDebitos Val(txtPKId)
        GeraTotalIndexador
    End If
    
End Sub

Private Sub tdb_Debitos_Click()
    blnPrimeiraVezDebito = True
End Sub

Private Sub tdb_Debitos_FilterChange()
    blnPrimeiraVezDebito = False
    gblnFilraCampos tdb_Parcelas
End Sub

Private Sub tdb_Debitos_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Debitos, ColIndex
End Sub

Private Sub tdb_Debitos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tdb_Debitos.EOF Then
        'mskstrInscricaoCadastralDebito = gstrFormataInscricao(tdb_Debitos.Columns("Inscrição Cadastral").Value, tdb_Debitos.Columns("intUtilizacao"))
        mskstrInscricaoCadastralDebito = gstrFormataInscricao(Right(gstrENulo(tdb_Debitos.Columns("Inscrição Cadastral").Value), gintRetornaTamanhoMascara(Val(gstrENulo(tdb_Debitos.Columns("intUtilizacao"))))), Val(gstrENulo(tdb_Debitos.Columns("intUtilizacao"))))
        txtStrcomposicaodareceita = gstrENulo(tdb_Debitos.Columns("Composição da Receita").Value)
        txtintExercicioDebito(1) = gstrENulo(tdb_Debitos.Columns("Exercício").Value)
        PreencheGrdParcelasDebitos Val(txtPKId), gstrENulo(tdb_Debitos.Columns("Composição da Receita").Value), Val(gstrENulo(tdb_Debitos.Columns("Exercicio").Value))
    End If
End Sub

Private Sub tdb_Parcelas_FilterChange()
    gblnFilraCampos tdb_Parcelas
End Sub

Private Sub tdb_Parcelas_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Parcelas, ColIndex
End Sub

Private Sub tdb_ParcelasDebitos_FilterChange()
    gblnFilraCampos tdb_ParcelasDebitos
End Sub

Private Sub tdb_ParcelasDebitos_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Select Case ColIndex
        Case Is = 3, 4, 5, 6
            Value = gstrConvVrDoSql(Value, 2)
    End Select
End Sub

Private Sub tdb_ParcelasDebitos_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_ParcelasDebitos, ColIndex
End Sub

Private Sub txt_dtmVencimento_GotFocus()
    MarcaCampo txt_dtmVencimento
    If Len(txt_dtmVencimento) = 0 Then
        txt_dtmVencimento.Text = gstrDataDoSistema
    End If
End Sub

Private Sub txt_dtmVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmVencimento
End Sub

Private Sub txt_dtmVencimento_LostFocus()
    
    txt_dtmVencimento = gstrDataFormatada(txt_dtmVencimento)
    
    'Vamos aplicar o acrescimo de juros do desconto provisorio
    AplicarAcrescimoDesctoProvisorio
    
End Sub

Private Sub txt_QtdeParcelas_LostFocus()
    
    If Len(txt_QtdeParcelas) > 0 And Len(txtdblValor) > 0 Then
        If Val(txt_QtdeParcelas) = 0 Then
            ExibeMensagem "Não é possível parcelar em 0 vezes"
            If txt_QtdeParcelas.Enabled Then txt_QtdeParcelas.SetFocus
            Exit Sub
        ElseIf CCur(txtdblValor) = 0 Then
            ExibeMensagem "Não é possível parcelar valor 0,00"
            If txt_QtdeParcelas.Enabled Then txt_QtdeParcelas.SetFocus
            Exit Sub
        End If
        
        'Vamos aplicar o acrescimo por parcela, caso seja parametrizado e nao esteja com desconto provisorio
        If Not dbc_intDescProvisorios.MatchedWithList Then
            AplicarAcrescimoPorParcela
        End If
        
        txt_dblValorParcela = gstrConvVrDoSql(((CCur(txtdblValor) - dblAcrescimoDescProv) + CCur(gstrConvVrDoSql(txtdblAcrescimos, , , True))) / txt_QtdeParcelas, 2)
        
        'Vamos aplicar o acrescimo de juros do desconto provisorio
        AplicarAcrescimoDesctoProvisorio
        
    Else
        txt_dblValorParcela = ""
    End If

    If Not blnVerificaParametrosParaParcelamento(CCur(gstrConvVrDoSql(txtdblValor, , , True)) + CCur(gstrConvVrDoSql(txtdblAcrescimos, , , True)), gstrConvVrDoSql(txt_dblValorParcela, , , True), Val(txt_QtdeParcelas)) Then
        txt_QtdeParcelas = Space$(0)
        txt_dblValorParcela = Space$(0)
        If txt_QtdeParcelas.Enabled Then txt_QtdeParcelas.SetFocus
    End If
End Sub

Private Sub txtbitDigitoProcesso_GotFocus()
    MarcaCampo txtbitDigitoProcesso
End Sub

Private Sub txtbitDigitoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigitoProcesso
End Sub

Private Sub txt_Datainicialprov_GotFocus()
    MarcaCampo txt_datainicialprov
End Sub

Private Sub txt_Datainicialprov_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_datainicialprov
End Sub

Private Sub txt_Datainicialprov_LostFocus()
    txt_datainicialprov = gstrDataFormatada(txt_datainicialprov)
End Sub

Private Sub txt_Datafinalprov_GotFocus()
    MarcaCampo txt_datafinalprov
End Sub

Private Sub txt_Datafinalprov_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_datafinalprov
End Sub

Private Sub txt_Datafinalprov_LostFocus()
    txt_datafinalprov = gstrDataFormatada(txt_datafinalprov)
End Sub

Private Sub txtdblAcrescimos_GotFocus()
    MarcaCampo txtdblAcrescimos
End Sub

Private Sub txtdblAcrescimos_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblAcrescimos
End Sub

Private Sub txtdblAcrescimos_LostFocus()
Dim dblValor, dblAcrescimos, dblValorParcela As Double

    txtdblAcrescimos = gstrConvVrDoSql(txtdblAcrescimos, 2)
    
    If Len(txt_QtdeParcelas) > 0 And Len(txtdblValor) > 0 Then
        If Val(txt_QtdeParcelas) = 0 Then
            ExibeMensagem "Não é possível parcelar em 0 vezes"
            Exit Sub
        ElseIf CCur(txtdblValor) = 0 Then
            ExibeMensagem "Não é possível parcelar valor 0,00"
            Exit Sub
        End If
        
        'Vamos aplicar o acrescimo por parcela, caso seja parametrizado e nao esteja com desconto provisorio
        If Not dbc_intDescProvisorios.MatchedWithList Then
            AplicarAcrescimoPorParcela
        End If
        
        txt_dblValorParcela = gstrConvVrDoSql(((CCur(txtdblValor) - dblAcrescimoDescProv) + CCur(IIf(Trim(txtdblAcrescimos) = "", "0", txtdblAcrescimos))) / txt_QtdeParcelas, 2)
    Else
        txt_dblValorParcela = ""
    End If
    
        
    dblValor = CDbl(gstrConvVrDoSql(txtdblValor.Text, , , True))
    dblAcrescimos = CDbl(gstrConvVrDoSql(txtdblAcrescimos.Text, , , True))
    dblValorParcela = CDbl(gstrConvVrDoSql(txt_dblValorParcela, , , True))
    
    If Not blnVerificaParametrosParaParcelamento(dblValor + dblAcrescimos, dblValorParcela, Val(txt_QtdeParcelas)) Then
        txt_QtdeParcelas = Space$(0)
        txt_dblValorParcela = Space$(0)
    End If
    
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
Dim dblValor, dblAcrescimos, dblValorParcela As Double

    txtdblValor = gstrConvVrDoSql(txtdblValor, 2)
    
    If Len(txt_QtdeParcelas) > 0 And Len(txtdblValor) > 0 Then
        If Val(txt_QtdeParcelas) = 0 Then
            ExibeMensagem "Não é possível parcelar em 0 vezes"
            Exit Sub
        ElseIf CCur(txtdblValor) = 0 Then
            ExibeMensagem "Não é possível parcelar valor 0,00"
            Exit Sub
        End If
        
        'Vamos aplicar o acrescimo por parcela, caso seja parametrizado e nao esteja com desconto provisorio
        If Not dbc_intDescProvisorios.MatchedWithList Then
            AplicarAcrescimoPorParcela
        End If
        
        txt_dblValorParcela = gstrConvVrDoSql(((CCur(txtdblValor) - dblAcrescimoDescProv) + CCur(IIf(Trim(txtdblAcrescimos) = "", "0", txtdblAcrescimos))) / txt_QtdeParcelas, 2)
    Else
        txt_dblValorParcela = ""
    End If
    
        
    dblValor = CDbl(gstrConvVrDoSql(txtdblValor.Text, , , True))
    dblAcrescimos = CDbl(gstrConvVrDoSql(txtdblAcrescimos.Text, , , True))
    dblValorParcela = CDbl(gstrConvVrDoSql(txt_dblValorParcela, , , True))
    
    If Not blnVerificaParametrosParaParcelamento(dblValor + dblAcrescimos, dblValorParcela, Val(txt_QtdeParcelas)) Then
        txt_QtdeParcelas = Space$(0)
        txt_dblValorParcela = Space$(0)
    End If

End Sub

Private Sub txtDblVlIndexador_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtdtmData_GotFocus()
    MarcaCampo txtdtmData
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmData
End Sub

Private Sub txtdtmData_LostFocus()
    txtdtmData = gstrDataFormatada(txtdtmData)
End Sub

Private Sub txtdtmDtCancelamento_GotFocus()
    MarcaCampo txtdtmDtCancelamento
End Sub

Private Sub txtdtmDtCancelamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtCancelamento
End Sub

Private Sub txtdtmDtCancelamento_LostFocus()
    txtdtmDtCancelamento = gstrDataFormatada(txtdtmDtCancelamento)
End Sub

Private Sub txtdtmDtUtilizacao_GotFocus()
    MarcaCampo txtdtmDtUtilizacao
End Sub

Private Sub txtdtmDtUtilizacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtUtilizacao
End Sub

Private Sub txtdtmDtUtilizacao_LostFocus()
    txtdtmDtUtilizacao = gstrDataFormatada(txtdtmDtUtilizacao)
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub

Private Sub txtintCepC_GotFocus()
    MarcaCampo txtintCEPC
End Sub

Private Sub txtintCepC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCEPC
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintExercicioProcesso_GotFocus()
    MarcaCampo txtintExercicioProcesso
End Sub

Private Sub txtintExercicioProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicioProcesso
End Sub

Private Sub txtintExercicioRequerimento_GotFocus()
    MarcaCampo txtintExercicioRequerimento
End Sub

Private Sub txtintExercicioRequerimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicioRequerimento
End Sub

Private Sub txtintRequerimento_GotFocus()
    MarcaCampo txtintRequerimento
End Sub

Private Sub txtintRequerimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintRequerimento
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

Private Sub txtstrCodigoProcesso_GotFocus()
    MarcaCampo txtstrCodigoProcesso
End Sub

Private Sub txtstrCodigoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigoProcesso
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

Private Sub txtstrLeiDecreto_GotFocus()
    MarcaCampo txtstrLeiDecreto
End Sub

Private Sub txtstrLeiDecreto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLeiDecreto
End Sub

Private Function strQueryMoedas() As String
Dim strSql As String

    strSql = "SELECT Pkid,"
    strSql = strSql & " strAbreviatura"
    strSql = strSql & " FROM "
    strSql = strSql & gstrMoedas
    strSql = strSql & " ORDER BY"
    strSql = strSql & " strAbreviatura"
    
    strQueryMoedas = strSql

End Function

Private Function strQueryRequerente() As String
Dim strSql As String

    strSql = "SELECT Pkid,"
    strSql = strSql & " strNome"
    strSql = strSql & " FROM "
    strSql = strSql & gstrContribuinte
    strSql = strSql & " ORDER BY"
    strSql = strSql & " strNome"
    
    strQueryRequerente = strSql
    
End Function

Private Sub PreencheDadosRequerente()
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset

    strSql = "SELECT bytNaturezaJuridica,"
    strSql = strSql & " strIdentidade,"
    strSql = strSql & " strCNPJCPF"
    strSql = strSql & " FROM "
    strSql = strSql & gstrContribuinte
    strSql = strSql & " WHERE Pkid =" & Val(dbcstrNomeProprietario.BoundText)
       
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            mskstrCNPJCPF.Mask = IIf(adoResultado!bytNaturezaJuridica = 0, "###\.###\.###\-##", "")
            mskstrCNPJCPF.Text = gstrENulo(adoResultado!StrCnpjCpf)
            txtstridentidade.Text = gstrENulo(adoResultado!STRIDENTIDADE)
        End If
    
    End If

End Sub

Private Sub txtstrLogradouro_GotFocus()
    MarcaCampo txtstrLogradouro
    tab_3dEnderecos.Tab = 0
End Sub

Private Sub txtstrLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLogradouro
End Sub

Private Sub txtstrLogradouroC_GotFocus()
    MarcaCampo txtstrLogradouroC
    tab_3dEnderecos.Tab = 1
End Sub

Private Sub txtstrLogradouroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLogradouroC
End Sub

Private Sub txtstrMunicipio_GotFocus()
    MarcaCampo txtstrMunicipio
End Sub

Private Sub txtstrMunicipio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrMunicipio
End Sub

Private Sub txtstrMunicipioC_GotFocus()
    MarcaCampo txtstrMunicipioC
End Sub

Private Sub txtstrMunicipioC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrMunicipioC
End Sub

Private Sub txtstrNumero_GotFocus()
    MarcaCampo txtstrNumero
End Sub

Private Sub txtstrNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNumero
End Sub

Private Sub txtstrNumeroC_GotFocus()
    MarcaCampo txtstrNumeroC
End Sub

Private Sub txtstrNumeroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNumeroC
End Sub

Private Sub txtstrUF_GotFocus()
    MarcaCampo txtstrUf
End Sub

Private Sub txtstrUF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "U", txtstrUf
End Sub

Private Sub txtstrUFC_GotFocus()
    MarcaCampo txtstrUFC
End Sub

Private Sub txtstrUFC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "U", txtstrUFC
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrNovo)
            
            LimpaObjeto Me
            txt_QtdeParcelas = ""
            txt_dtmVencimento = ""
            lbl_Indexador.Caption = ""
            txtdblTotalIndexador.Text = ""
            txt_strIndexador.Text = ""
            txt_dblValorParcela.Text = ""
            Set tdb_Parcelas.DataSource = Nothing
            Set tdb_Debitos.DataSource = Nothing
            Set tdb_ParcelasDebitos.DataSource = Nothing
            HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrAplicar
        Case Is = UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        Case Is = UCase(gstrLocalizar)
            blnPrimeiraVez = True
            LeDaTabelaParaObj "", tdb_Acordos, strQuery
        Case Is = UCase(gstrSalvar)
            If blnDadosOk Then
                If Not mblnAlterando Then
                    GravarAcordo
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
                End If
            End If
        Case Is = UCase(gstrImprimir)
            ImprimeRelatorio rptAcordos, strQueryRelatorio
        Case Is = UCase(gstrDeletar)
        
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
                        
            
            If blnExluiAcordo Then
                gobjBanco.ExecutaCommitTrans
                MantemForm gstrNovo
            Else
                gobjBanco.ExecutaRollbackTrans
            End If
            
            
    End Select
    
End Sub

Private Function strQueryRelatorio() As String
' RESPONSAVEL LEANDRO  30/06/2004
Dim strSql As String
        
strSql = "SELECT AC.Pkid,"
    strSql = strSql & " LA.strNumeroAviso " & strCONCAT & "'/'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "LA.intExercicio") & " Numero,"
    strSql = strSql & " AC.dtmData Data,"
    strSql = strSql & " AC.strLeiDecreto LeiDecreto,"
    strSql = strSql & " AC.dblValor Valor,"
    strSql = strSql & " AC.dblAcrescimos Acrescimos,"
    strSql = strSql & " MO.strAbreviatura Moeda,"
    strSql = strSql & " AC.strCodigoProcesso " & strCONCAT & "'-'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "AC.bitDigitoProcesso") & strCONCAT & "'/'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "AC.intExercicioProcesso") & " Processo,"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "AC.intRequerimento") & strCONCAT & "'/'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "AC.intExercicioRequerimento") & " Requerimento,"
    strSql = strSql & " LA.strNomeProprietario Requerente,"
    strSql = strSql & " LA.strCNPJCPF CNPJCPF,"
    strSql = strSql & " LA.strIdentidade RG "
    
strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrAcordo & " AC, "
    strSql = strSql & gstrMoedas & " MO"
    
strSql = strSql & " WHERE"
    strSql = strSql & " AC.intLancamentoAlfa = LA.Pkid AND"
    strSql = strSql & " AC.intMoedas = MO.Pkid"

strSql = strSql & " ORDER BY LA.strInscricao"

strQueryRelatorio = strSql
    


End Function

Private Function strQuery() As String
Dim strSql     As String
Dim strSqlSub  As String

    strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " strIdentificacao " & _
                "FROM " & gstrAcordoDebitos & " AD WHERE intAcordo = AC.Pkid " & _
                "ORDER BY strIdentificacao, strComposicaoDaReceita, intExercicio"
    strSqlSub = gstrTOPnOracle(strSqlSub, 1, "intAcordo", "AC.Pkid", "strIdentificacao")
    
    strSql = "SELECT AC.Pkid, "
    
    'strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1) & strCONCAT & " '/'" & strCONCAT
    strSql = strSql & strSUBSTRING & "(LA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "LA.intExercicio") & " Numero,"
    strSql = strSql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso,"
    'strSql = strSql & " LA.strNumeroAviso " & strCONCAT & "'/'" & strCONCAT
    'strSql = strSql & gstrCONVERT(CDT_VARCHAR, "LA.intExercicio") & " Numero,"
    strSql = strSql & " AC.dtmData Data,"
    strSql = strSql & " AC.strLeiDecreto LeiDecreto,"
    strSql = strSql & " AC.dblValor Valor,"
    strSql = strSql & " AC.dblAcrescimos Acrescimos,"
    'strSql = strSql & "(" & strSqlSub & ") strIdentificacao,"
    strSql = strSql & " MO.strAbreviatura Moeda,"
    strSql = strSql & " AC.strCodigoProcesso " & strCONCAT & "'-'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "AC.bitDigitoProcesso") & strCONCAT & "'/'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "AC.intExercicioProcesso") & " Processo,"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "AC.intRequerimento") & strCONCAT & "'/'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "AC.intExercicioRequerimento") & " Requerimento,"
    strSql = strSql & " LA.strNomeProprietario NomeProprietario,"
    strSql = strSql & " LA.strCNPJCPF CNPJCPF,"
    strSql = strSql & " LA.strIdentidade Identidade, "
    strSql = strSql & " AC.intLancamentoAlfa "
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrAcordo & " AC, "
    strSql = strSql & gstrMoedas & " MO"
    strSql = strSql & " WHERE"
    strSql = strSql & " AC.intLancamentoAlfa = LA.Pkid AND"
    strSql = strSql & " AC.intMoedas = MO.Pkid"
    
    strSql = strSql & strFiltros
    
    strSql = strSql & " ORDER BY LA.strNumeroAviso, LA.intExercicio"
    
    strQuery = strSql

End Function

Private Function strFiltros() As String
Dim strSql As String

    If mskstrInscricao.Text <> "" Then
        strSql = " AND strInscricao LIKE " & "'" & UCase(String(gintLenInscricao - Len(mskstrInscricao) - 4, "0") & mskstrInscricao) & "%'"
    End If

    If txtintExercicio.Text <> "" Then
        strSql = strSql & " AND LA.intExercicio = " & Val(txtintExercicio.Text)
    End If
    
    If txtdtmData.Text <> "" Then
        If gblnDataValida(txtdtmData.Text) Then
            strSql = strSql & " AND AC.dtmData =" & gstrConvDtParaSql(txtdtmData.Text)
        End If
    End If
    
    If txtstrLeiDecreto.Text <> "" Then
        strSql = strSql & " AND AC.strLeiDecreto LIKE '" & txtstrLeiDecreto.Text & "%'"
    End If
    
    If txtdblValor.Text <> "" Then
        strSql = strSql & " AND AC.dblValor = " & gstrConvVrParaSql(txtdblValor.Text)
    End If
    
    If txtdblAcrescimos.Text <> "" Then
        strSql = strSql & " AND AC.dblAcrescimos = " & gstrConvVrParaSql(txtdblAcrescimos.Text)
    End If
    
    If dbcintMoeda.MatchedWithList Then
        strSql = strSql & " AND AC.intMoedas = " & Val(dbcintMoeda.BoundText)
    End If
    
    If txtstrCodigoProcesso.Text <> "" Then
        strSql = strSql & " AND AC.strCodigoProcesso LIKE '" & txtstrCodigoProcesso.Text & "%'"
    End If
    
    If txtintExercicioProcesso.Text <> "" Then
        strSql = strSql & " AND AC.intExercicioProcesso = " & txtintExercicioProcesso.Text
    End If
    
    If txtbitDigitoProcesso.Text <> "" Then
        strSql = strSql & " AND AC.bitDigitoProcesso = " & txtbitDigitoProcesso.Text
    End If
        
    If txtintRequerimento.Text <> "" Then
        strSql = strSql & " AND AC.intRequerimento = " & txtintRequerimento.Text
    End If
    
    If txtintExercicioRequerimento.Text <> "" Then
        strSql = strSql & " AND AC.intExercicioRequerimento = " & txtintExercicioRequerimento.Text
    End If
           
    If dbcstrNomeProprietario.Text <> "" Then
        strSql = strSql & " AND LA.strNomeProprietario LIKE '" & dbcstrNomeProprietario.Text & "%'"
    End If
    
    strFiltros = strSql

End Function

Private Sub PreencheGrdParcelas(lngPkid As Long)
Dim strSql As String
  
  If bytDBType = SQLServer Then
    
    strSql = "SELECT LV.PKId, "
    strSql = strSql & "LV.intParcela AS Parcela, "
    strSql = strSql & "LV.dtmDtVencimento AS Vencimento, "
    strSql = strSql & "LP.DTMDTPAGAMENTO AS Baixa, "
    strSql = strSql & "CB.STRDESCRICAO AS Descricao, "
    strSql = strSql & "LP.STROBSERVACAO AS Observacao, "
    strSql = strSql & "LV.dblValor AS Valor "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoValor & " LV LEFT OUTER JOIN "
    strSql = strSql & gstrLancamentoPagamento & " LP ON LV.PKId = LP.INTLANCAMENTOVALOR LEFT OUTER JOIN "
    strSql = strSql & gstrCodigoDeBaixa & " CB ON LP.INTCODIGOBAIXA = CB.PKID "
    strSql = strSql & "WHERE LV.intLancamentoAlfa = " & Val(tdb_Acordos.Columns("intLancamentoAlfa").Value)
    
  Else
  
    strSql = "SELECT LV.Pkid, "
    strSql = strSql & "LV.intParcela Parcela, "
    strSql = strSql & "LV.dtmDtVencimento Vencimento, "
    strSql = strSql & "LP.dtmDtPagamento Baixa, "
    strSql = strSql & "CB.strDescricao Descricao, "
    strSql = strSql & "LP.strObservacao Observacao, "
    strSql = strSql & "LV.dblValor Valor "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrCodigoDeBaixa & " CB, "
    strSql = strSql & gstrLancamentoPagamento & " LP "
    strSql = strSql & "WHERE "
    strSql = strSql & "LV.Pkid " & strOUTJSQLServer & "= " & "LP.intLancamentoValor " & strOUTJOracle & " AND "
    strSql = strSql & "CB.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " LP.Intcodigobaixa  AND "
    strSql = strSql & "LV.intLancamentoAlfa = " & Val(tdb_Acordos.Columns("intLancamentoAlfa").Value) & " "
  End If
    
    strSql = strSql & " ORDER BY LV.intParcela"
    
    LeDaTabelaParaObj "", tdb_Parcelas, strSql

End Sub

Private Sub PreecheCampos()
Dim adoResultado As ADODB.Recordset

    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(QueryPreencheCampos, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            mskstrInscricao.Text = gstrENulo(adoResultado!NumeroInscricao)
            mskstrInscricaoDebito = gstrENulo(adoResultado!NumeroInscricao)
            txtintExercicio = gstrENulo(adoResultado!EXERCICIO)
            txtintExercicioDebito(0) = gstrENulo(adoResultado!EXERCICIO)
            txtdtmData = gstrENulo(adoResultado!DataAcordo)
            txtdtmDataDebito = gstrENulo(adoResultado!DataAcordo)
            txtdtmDtCancelamento = gstrENulo(adoResultado!Dtmdtcancelamento)
            txtdtmDtUtilizacao = gstrENulo(adoResultado!dtmDtUtilizacao)
            txtstrLeiDecreto = gstrENulo(adoResultado!LeiDecreto)
            txtstrLeiDecretoDebito = gstrENulo(adoResultado!LeiDecreto)
            txtdblValor = gstrConvVrDoSql(gstrENulo(adoResultado!Valor), 2)
            txtdblValorDebito = gstrConvVrDoSql(gstrENulo(adoResultado!Valor), 2)
            txtdblAcrescimos = gstrConvVrDoSql(gstrENulo(adoResultado!Acrescimos), 2)
            txtdblAcrescimosDebito = gstrConvVrDoSql(gstrENulo(adoResultado!Acrescimos), 2)
            PreencherListaDeOpcoes dbcintMoeda, gstrENulo(adoResultado!Moeda)
            PreencherListaDeOpcoes dbcintMoedaDebito, gstrENulo(adoResultado!Moeda)
            txtdblVlIndexador = gstrConvVrDoSql(gstrENulo(adoResultado!dblvlIndexador), 6, , True)
            txt_strIndexador.Text = gstrENulo(adoResultado!Strindexador)
            txtstrCodigoProcesso = gstrENulo(adoResultado!CodigoProcesso)
            txtbitDigitoProcesso = gstrENulo(adoResultado!DigitoProcesso)
            txtintExercicioProcesso = gstrENulo(adoResultado!ExercicioProcesso)
            txtintRequerimento = gstrENulo(adoResultado!Requerimento)
            txtintExercicioRequerimento = gstrENulo(adoResultado!ExercicioRequerimento)
            dbcstrNomeProprietario = gstrENulo(adoResultado!NomeProprietario)
            mskstrCNPJCPF = gstrENulo(adoResultado!cnpjcpf)
            txtstridentidade = gstrENulo(adoResultado!Identidade)
            txtstrLogradouro = gstrENulo(adoResultado!Logradouro)
            txtstrNumero = gstrENulo(adoResultado!Numero)
            txtstrComplemento = gstrENulo(adoResultado!Complemento)
            txtstrBairro = gstrENulo(adoResultado!Bairro)
            txtstrMunicipio = gstrENulo(adoResultado!Municipio)
            txtstrUf = gstrENulo(adoResultado!UF)
            txtintCep = Format(gstrENulo(adoResultado!CEP), "00000-000")
            txtstrLogradouroC = gstrENulo(adoResultado!LogradouroC)
            txtstrNumeroC = gstrENulo(adoResultado!NumeroC)
            txtstrComplementoC = gstrENulo(adoResultado!ComplementoC)
            txtstrBairroC = gstrENulo(adoResultado!BairroC)
            txtstrMunicipioC = gstrENulo(adoResultado!MunicipioC)
            txtstrUFC = gstrENulo(adoResultado!UFC)
            txtintCEPC = Format(gstrENulo(adoResultado!CEPC), "00000-000")
            txt_strAnistia = gstrENulo(adoResultado!strAnistia)
            txt_strAnistiaLegislacao = gstrENulo(adoResultado!stranistialegislacao)
            strInscricaoAtual = gstrENulo(adoResultado!NumeroInscricao)
        End If
    
    End If

End Sub

Private Sub PreencheGrdParcelasDebitos(lngPkidAcordo As Long, strComposicao As String, intExercicio As Integer)
Dim strSql As String
       
    strSql = "SELECT AD.Pkid,"
    strSql = strSql & " AD.intParcela Parcela,"
    strSql = strSql & " AD.dtmDtVencimento Vencimento,"
    strSql = strSql & " AD.dblPrincipal Valor,"
    strSql = strSql & " AD.dblJuros Juros, "
    strSql = strSql & " AD.dblMulta Multa, "
    strSql = strSql & " AD.dblCorrecaoMonetaria Correcao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrAcordoDebitos & " AD "
'    strSql = strSql & " WHERE AD.strIdentificacao = '" & tdb_Debitos.Columns("Inscricao").Value & "' AND AD.intExercicio = " & tdb_Debitos.Columns("Exercicio").Value
    strSql = strSql & "WHERE AD.intAcordo = " & lngPkidAcordo & " AND lTrim(rTrim(strComposicaoDaReceita)) = '" & strComposicao & "' And intExercicio = " & intExercicio
    strSql = strSql & " ORDER BY AD.intParcela"
       
    LeDaTabelaParaObj "", tdb_ParcelasDebitos, strSql
       
End Sub

Private Sub PreencheGrdDebitos(lngPkidAcordo As Long)
Dim strSql As String
    
    strSql = "SELECT AD.STRIDENTIFICACAO Inscricao, " & _
                    "AD.STRCOMPOSICAODARECEITA ComposicaoDaReceita, " & _
                    "AD.Intexercicio Exercicio, " & _
                    gstrCONVERT(CDT_numeric, "AD.strNumeroAviso") & " strNumeroAviso," & _
                    "AD.intUtilizacao " & _
             "FROM " & gstrAcordoDebitos & " AD " & _
             "WHERE intAcordo = " & lngPkidAcordo & " " & _
             "GROUP BY AD.strIdentificacao, AD.strComposicaoDaReceita, AD.intExercicio,AD.intUtilizacao,AD.strNumeroAviso " & _
             "ORDER BY AD.strIdentificacao, AD.strComposicaoDaReceita, AD.Intexercicio,AD.intUtilizacao"

    LeDaTabelaParaObj "", tdb_Debitos, strSql

End Sub

Private Sub PreencheDadosProprietario()
Dim strSql As String
Dim adoResultado As ADODB.Recordset

    strSql = "SELECT bytNaturezaJuridica,"
    strSql = strSql & " strCNPJCPF CNPJCPF,"
    strSql = strSql & " strIdentidade Identidade"
    strSql = strSql & " FROM "
    strSql = strSql & gstrContribuinte
    strSql = strSql & " WHERE Pkid = " & Val(dbcstrNomeProprietario.BoundText)

    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            mskstrCNPJCPF.Mask = IIf(Val(gstrENulo(adoResultado!bytNaturezaJuridica)) = 0, "###\.###\.###\-##", "")
            mskstrCNPJCPF = gstrENulo(adoResultado!cnpjcpf)
            txtstridentidade = gstrENulo(adoResultado!Identidade)
        End If
    
    End If

End Sub

Private Function QueryPreencheCampos() As String
Dim strSql As String
    
    strSql = "SELECT AC.Pkid,"
    'strSQL = strSQL & " LA.strInscricao NumeroInscricao,"
    'strSQL = strSQL & strSUBSTRING & "(LA.strInscricao, " & gintRetornaTamanhoMascara(TYP_ACORDO) * -1 & ") NumeroInscricao, "
    strSql = strSql & strSUBSTRING & "(LA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ") NumeroInscricao, "
    strSql = strSql & " LA.intExercicio Exercicio,"
    strSql = strSql & " AC.dtmData DataAcordo,"
    strSql = strSql & " AC.strLeiDecreto LeiDecreto,"
    strSql = strSql & " AC.dblValor Valor,"
    strSql = strSql & " AC.dblAcrescimos Acrescimos,"
    strSql = strSql & " AC.intMoedas Moeda,"
    strSql = strSql & " AC.strCodigoProcesso CodigoProcesso,"
    strSql = strSql & " AC.bitDigitoProcesso DigitoProcesso,"
    strSql = strSql & " AC.intExercicioProcesso ExercicioProcesso,"
    strSql = strSql & " AC.intRequerimento Requerimento,"
    strSql = strSql & " AC.intExercicioRequerimento ExercicioRequerimento,"
    strSql = strSql & " AC.stranistia,"
    strSql = strSql & " AC.stranistialegislacao,"
    strSql = strSql & " AC.dtmDtCancelamento,"
    strSql = strSql & " AC.dtmDtUtilizacao,"
    strSql = strSql & " LA.strIndexador,"
    strSql = strSql & " LA.dblvlindexador,"
    strSql = strSql & " LA.strNomeProprietario NomeProprietario,"
    strSql = strSql & " LA.strCNPJCPF CNPJCPF,"
    strSql = strSql & " LA.strIdentidade Identidade,"
    strSql = strSql & " LA.strLogradouro Logradouro,"
    strSql = strSql & " LA.strNumero Numero,"
    strSql = strSql & " LA.strComplemento Complemento,"
    strSql = strSql & " LA.strBairro Bairro,"
    strSql = strSql & " LA.strMunicipio Municipio,"
    strSql = strSql & " LA.strUF UF,"
    strSql = strSql & " LA.intCep CEP,"
    strSql = strSql & " LA.strLogradouroC LogradouroC,"
    strSql = strSql & " LA.strNumeroC NumeroC,"
    strSql = strSql & " LA.strComplementoC ComplementoC,"
    strSql = strSql & " LA.strBairroC BairroC,"
    strSql = strSql & " LA.strMunicipioC MunicipioC,"
    strSql = strSql & " LA.strUFC UFC,"
    strSql = strSql & " LA.intCepC CEPC"
    strSql = strSql & " FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrAcordo & " AC"
    strSql = strSql & " WHERE"
    strSql = strSql & " AC.intLancamentoAlfa = LA.Pkid"
        
    strSql = strSql & " AND AC.Pkid = " & Val(txtPKId.Text)
    
    strSql = strSql & " ORDER BY LA.strInscricao"

    QueryPreencheCampos = strSql

End Function

Private Function blnDadosOk() As Boolean

    blnDadosOk = False
    
    If Len(Trim(mskstrInscricao)) = 0 Or Len(Trim(txtintExercicio)) = 0 Then
        ExibeMensagem "É necessário preencher o Número."
        mskstrInscricao.SetFocus
        Exit Function
    End If

    If Len(Trim(txtdtmData)) = 0 Then
        ExibeMensagem "É necessário preencher a Data."
        txtdtmData.SetFocus
        Exit Function
    End If

    If Len(Trim(txtdblValor)) = 0 Then
        ExibeMensagem "É necessário preencher o Valor."
        txtdblValor.SetFocus
        Exit Function
    End If

    If Len(dbcintMoeda.BoundText) = 0 Then
        ExibeMensagem "É necessário preencher a Moeda."
        dbcintMoeda.SetFocus
        Exit Function
    End If

    If Len(dbcstrNomeProprietario.BoundText) = 0 Then
        ExibeMensagem "É necessário preencher o Requerente."
        dbcstrNomeProprietario.SetFocus
        Exit Function
    End If
    
    If blnParcelamentoDebito Then
        If Len(Trim(txt_QtdeParcelas)) = 0 Then
            ExibeMensagem "É necessário preencher a Qtde de Parcelas."
            txt_QtdeParcelas.SetFocus
            Exit Function
        End If
        
        If Len(Trim(txt_dtmVencimento)) = 0 Then
            ExibeMensagem "É necessário preencher a 1ª Data de Vencimento."
            txt_dtmVencimento.SetFocus
            Exit Function
        End If
    End If
    
    If Len(Trim(txtstrLogradouroC)) = 0 Then
        ExibeMensagem "Endereço de Notificação incompleto."
        txtstrLogradouroC.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtstrNumeroC)) = 0 Then
        ExibeMensagem "Endereço de Notificação incompleto."
        txtstrNumeroC.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtstrBairroC)) = 0 Then
        ExibeMensagem "Endereço de Notificação incompleto."
        txtstrBairroC.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtstrMunicipioC)) = 0 Then
        ExibeMensagem "Endereço de Notificação incompleto."
        txtstrMunicipioC.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtstrUFC)) = 0 Then
        ExibeMensagem "Endereço de Notificação incompleto."
        txtstrUFC.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtintCEPC)) = 0 Then
        ExibeMensagem "Endereço de Notificação incompleto."
        txtintCEPC.SetFocus
        Exit Function
    End If
    
    'If Trim(cbo_strIndexador.Text) <> "" Then
    '    If Trim(txtDblvlindexador) = "" Then
    '        ExibeMensagem "O valor do indexador deve ser preenchido."
    '    End If
    'End If
    
    If Len(Trim(txtstrCodigoProcesso)) > 0 Or Len(Trim(txtbitDigitoProcesso)) > 0 Or Len(Trim(txtintExercicioProcesso)) > 0 Then
        If Len(Trim(txtstrCodigoProcesso)) = 0 Or Len(Trim(txtbitDigitoProcesso)) = 0 Or Len(Trim(txtintExercicioProcesso)) = 0 Then
            ExibeMensagem "Processo incompleto."
            txtstrCodigoProcesso.SetFocus
            Exit Function
        Else
            If Not VerificaEmpenhoProcesso(txtstrCodigoProcesso, txtbitDigitoProcesso, txtintExercicioProcesso) Then
                ExibeMensagem "O Processo informado não é válido."
                txtstrCodigoProcesso.SetFocus
                Exit Function
            End If
        End If
    End If
    
    'If Not mblnAlterando Or (mblnAlterando And UCase$(strInscricaoAtual) <> UCase$(mskstrInscricao)) Then
    '    If gblnExisteCodigo(1, gstrLancamentoAlfa, "strInscricao", mskstrInscricao) Then
    '        ExibeMensagem "Já existe acordo com este Número."
    '        Exit Function
    '    End If
    'End If
    
    blnDadosOk = True

End Function

Private Function GravarAcordo() As Boolean
NovaGravacao:
    Dim adoResultado        As New ADODB.Recordset
    Dim adoReceitas         As New ADODB.Recordset
    
    Dim strSql              As String
    
    Dim intFor                    As Integer
    Dim intForReceitas            As Integer
    Dim intForReceitasArray       As Integer
    
    Dim lngPkidLAAcordo           As Long
    Dim lngPkidAcordo             As Long
    
    Dim intExisteReceita          As Integer
    Dim dblProporcao              As Double
    
    Dim dblValorParcela           As Double
    Dim dblValorDiferenca         As Double
    Dim dblValorReceita           As Double
    Dim dblValorTotalReceitas     As Double
    Dim dblValorDiferencaReceitas As Double
    
    Dim lngReceitaMulta           As Long
    Dim lngReceitaJuros           As Long
    Dim lngReceitaCorrecao        As Long
    Dim lngReceitaHonorario       As Long
    
    Dim lngComposicaoDaReceita    As Long
    Dim strComposicaoDaReceita    As String
    Dim intUtilizacao             As Integer
    
    Dim strParcelasAcordo         As String 'Variavel utilizada na impressao do carne
    Dim strInscricaoAux           As String

On Error GoTo Problema_Na_Rotina
    
    Set gobjBanco = New clsBanco
        
    gobjBanco.ExecutaBeginTrans
    
    'Vamos obter a primeira Composição do tipo Acordo
    If gobjBanco.CriaADO("SELECT Pkid, strDescricao, intUtilizacao FROM " & gstrComposicaoDaReceita & " WHERE intUtilizacao = " & TYP_ACORDO, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngComposicaoDaReceita = Space$(0) & adoResultado("Pkid").Value
            strComposicaoDaReceita = Space$(0) & adoResultado("strDescricao").Value
            intUtilizacao = Space$(0) & adoResultado("intUtilizacao").Value
        Else
            ExibeMensagem "Não foi encontrada nenhuma Composição do Tipo de Acordo. A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    Else
        ExibeMensagem "Não foi encontrada nenhuma Composição do Tipo de Acordo. A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    'Vamos obter as receitas de Multas, Juros e Correcao da Composicao de Receita
    If gobjBanco.CriaADO("SELECT intReceitaMulta, intReceitaJuros, intReceitaCorrecao FROM " & gstrParametroAtualizacao & " WHERE intExercicio = " & Year(gstrDataDoSistema) & " AND intComposicaoReceita = " & lngComposicaoDaReceita, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngReceitaMulta = Space$(0) & adoResultado("intReceitaMulta").Value
            lngReceitaJuros = Space$(0) & adoResultado("intReceitaJuros").Value
            lngReceitaCorrecao = Space$(0) & adoResultado("intReceitaCorrecao").Value
        Else
            ExibeMensagem "Não foi(ram) encontrada(s) receita(s) de Multa, Juros para a Composição de Receita. A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    Else
        ExibeMensagem "Não foi(ram) encontrada(s) receita(s) de Multa, Juros para a Composição de Receita. A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    'Vamos obter s receita de Honorarios
    If gobjBanco.CriaADO("SELECT intReceitaHonorarios FROM " & gstrParametrosTributario, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngReceitaHonorario = Space$(0) & Val(gstrENulo(adoResultado("intReceitaHonorarios").Value))
        Else
            ExibeMensagem "Não foi encontrada receita de Honorário. A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    Else
        ExibeMensagem "Não foi encontrada receita de Honorário. A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    ReDim vetReceitas(1, 0)
    
    'Vamos gravar os dados da tabela TBLLANCAMENTOALFA
    strSql = "INSERT INTO " & gstrLancamentoAlfa & " (" & _
             "strInscricao, strComposicaoDaReceita, strOcorrencia, strNomeProprietario, strCnpjCpf, " & _
             "strIdentidade, strLogradouro, strNumero, strComplemento, strBairro, " & _
             "strMunicipio, strUF, intCep, strLogradouroC, strNumeroC, " & _
             "strComplementoC, strBairroC, strMunicipioC, strUFC, intCepC, " & _
             "strNumeroAviso, strPromissario, strEmissao, intExercicio, intComposicaoDaReceita, strindexador, dblvlindexador, intUtilizacao, dtmDtAtualizacao, lngCodUsr) VALUES ('" & _
             mskstrInscricao & txtintExercicio & "','" & strComposicaoDaReceita & "', '','" & dbcstrNomeProprietario.Text & "','" & mskstrCNPJCPF & "','" & _
             txtstridentidade & "','" & txtstrLogradouro & "','" & txtstrNumero & "','" & txtstrComplemento & "','" & txtstrBairro & "','" & _
             txtstrMunicipio & "','" & txtstrUf & "'," & gstrENulo(Replace(txtintCep, "-", ""), , True) & ",'" & txtstrLogradouroC & "','" & txtstrNumeroC & "','" & _
             txtstrComplementoC & "','" & txtstrBairroC & "','" & txtstrMunicipioC & "','" & txtstrUFC & "'," & gstrENulo(Replace(txtintCEPC, "-", ""), , True) & ",'" & _
             mskstrInscricao & "'," & "'','00'," & txtintExercicio & "," & lngComposicaoDaReceita & ",'" & Trim(txt_strIndexador.Text) & "'," & gstrConvVrParaSql(txtdblVlIndexador) & "," & intUtilizacao & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"

    If Not gobjBanco.Execute(strSql) Then
        ExibeMensagem "A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    lngPkidLAAcordo = glngRetornaPkidTabelaPai("seqtblLancamentoAlfa", gstrLancamentoAlfa)
    
    'Vamos gravar os dados da tabela TBLACORDO
    strSql = "INSERT INTO " & gstrAcordo & " (" & _
             "dblValor, dblAcrescimos, intMoedas, strLeiDecreto, strCodigoProcesso, " & _
             "bitDigitoProcesso, intExercicioProcesso, dtmData, strObservacao, intRequerimento, " & _
             "intExercicioRequerimento, intLancamentoAlfa, dtmDtAtualizacao, lngCodUsr, stranistia, stranistialegislacao) VALUES (" & _
             gstrConvVrParaSql(txtdblValor) & "," & gstrConvVrParaSql(txtdblAcrescimos) & "," & dbcintMoeda.BoundText & ",'" & txtstrLeiDecreto.Text & "','" & txtstrCodigoProcesso & "'," & _
             gstrENulo(txtbitDigitoProcesso, , True) & "," & gstrENulo(txtintExercicioProcesso, , True) & "," & gstrConvDtParaSql(txtdtmData) & ",''," & gstrENulo(txtintRequerimento, , True) & "," & _
             gstrENulo(txtintExercicioRequerimento, , True) & "," & lngPkidLAAcordo & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ",'" & Trim(txt_strAnistia) & "','" & Trim(txt_strAnistiaLegislacao) & "')"

    If Not gobjBanco.Execute(strSql) Then
        ExibeMensagem "A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    lngPkidAcordo = glngRetornaPkidTabelaPai("seqtblAcordo", gstrAcordo)
    
    'Vamos gravar nas parcelas de TBLLANCAMENTOVALOR utilizadas para criar o Acordo, a referencia do Acordo
    For intFor = 0 To UBound(vetParcelasParaAcordo, 2)
        
        strSql = "UPDATE " & gstrLancamentoValor & " SET intLancamentoAlfaAcordo = " & lngPkidLAAcordo & " WHERE Pkid = " & vetParcelasParaAcordo(PKID_LANCAMENTO_VALOR, intFor)
        
        If Not gobjBanco.Execute(strSql) Then
            ExibeMensagem "A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
        
        'Vamos gravar as parcelas pre-definidas na tabela TBLACORDODEBITOS
        strSql = "INSERT INTO " & gstrAcordoDebitos & " (" & _
                 "intAcordo, strComposicaoDaReceita, strIdentificacao, intExercicio, intParcela, dtmDtVencimento, dblPrincipal, dblMulta, dblJuros, dblCorrecaoMonetaria, intExecutivoNumero, intExecutivoSerie, intCertidao, dblPrincipalOriginal, strPrincipalOriginalMoeda, strNumeroAviso, dtmDtDataAtualizacao, lngCodUsr, intUtilizacao) VALUES (" & _
                 lngPkidAcordo & ",'" & vetParcelasParaAcordo(COMPOSICAO_RECEITA, intFor) & "','" & vetParcelasParaAcordo(NUMERO_INSCRICAO_PURA, intFor) & "'," & vetParcelasParaAcordo(EXERCICIO, intFor) & "," & vetParcelasParaAcordo(PARCELA, intFor) & "," & gstrConvDtParaSql(vetParcelasParaAcordo(DATA_VENCIMENTO, intFor)) & "," & gstrConvVrParaSql(vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor)) & "," & gstrConvVrParaSql(vetParcelasParaAcordo(VALOR_MULTA, intFor)) & "," & gstrConvVrParaSql(vetParcelasParaAcordo(VALOR_JUROS, intFor)) & "," & gstrConvVrParaSql(vetParcelasParaAcordo(VALOR_CORRECAO, intFor)) & ",NULL,NULL,NULL," & gstrConvVrParaSql(vetParcelasParaAcordo(VALOR_ORIGINAL, intFor)) & ",'','" & vetParcelasParaAcordo(NUMERO_AVISO, intFor) & "'," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & "," & vetParcelasParaAcordo(Utilizacao, intFor) & ")"
        'Debug.Print vetParcelasParaAcordo(COMPOSICAO_RECEITA, intFor) & " - " & vetParcelasParaAcordo(Utilizacao, intFor)
                 
        If Not gobjBanco.Execute(strSql) Then
            ExibeMensagem "A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
        
        'Vamos montar o array que armazenara as receitas do acordo
        If gobjBanco.CriaADO("SELECT intReceita, dblValor, (SELECT Sum(dblValor) FROM tblLancamentoReceita WHERE intlancamentovalor = " & vetParcelasParaAcordo(PKID_LANCAMENTO_VALOR, intFor) & ") dblTotalReceitas FROM " & gstrLancamentoReceita & " WHERE intLancamentoValor = " & vetParcelasParaAcordo(PKID_LANCAMENTO_VALOR, intFor), 5, adoReceitas) Then
            
            If Not adoReceitas.EOF Then
            
                If adoReceitas.RecordCount = 1 Then
                    
                    intExisteReceita = -1
                    
                    'Vamos verificar se a receita ja existe no array
                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                        If vetReceitas(RECEITA, intForReceitasArray) = adoReceitas("intReceita").Value Then
                            intExisteReceita = intForReceitasArray
                            Exit For
                        End If
                    Next
                                            
                    'Caso ja exista a receita no array vamos somar, senao a criaremos
                    If intExisteReceita > -1 Then
                        vetReceitas(1, intExisteReceita) = vetReceitas(1, intExisteReceita) + CDbl((vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor)))
                    Else
                        'Nao é preciso redimensionar, pois o array é criado com 0
                        If Val(vetReceitas(0, 0)) > 0 Then
                            ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                        End If
                        vetReceitas(0, UBound(vetReceitas, 2)) = adoReceitas("intReceita").Value
                        vetReceitas(1, UBound(vetReceitas, 2)) = vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor)
                    End If
                        
                    'Caso seja apenas uma receita vamos atribuir o valor total a ela
                    'vetReceitas(0, UBound(vetReceitas, 2)) = adoReceitas("intReceita").Value
                    'vetReceitas(1, UBound(vetReceitas, 2)) = vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor)
                    
                    dblValorTotalReceitas = vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor)
                    
                Else
                    
                    'Vamos atribuir ao array receita a receita da parcela
                    For intForReceitas = 0 To adoReceitas.RecordCount - 1
                        
                        intExisteReceita = -1
                        
                        'Vamos verificar se a receita ja existe no array
                        For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                            If vetReceitas(RECEITA, intForReceitasArray) = adoReceitas("intReceita").Value Then
                                intExisteReceita = intForReceitasArray
                                Exit For
                            End If
                        Next
                        
                        'Vamos calcular a proporcao da receita ao valor original
                        'dblProporcao = adoReceitas("dblValor").Value / vetParcelasParaAcordo(VALOR_ORIGINAL, intFor)
                        dblProporcao = adoReceitas("dblValor").Value / adoReceitas("dblTotalReceitas")
                        
                        'Caso ja exista a receita no array vamos somar, senao a criaremos
                        If intExisteReceita > -1 Then
                            vetReceitas(1, intExisteReceita) = vetReceitas(1, intExisteReceita) + (vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor) * dblProporcao)
                        Else
                            'Nao é preciso redimensionar, pois o array é criado com 0
                            If Val(vetReceitas(0, 0)) > 0 Then
                                ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                            End If
                            vetReceitas(0, UBound(vetReceitas, 2)) = adoReceitas("intReceita").Value
                            vetReceitas(1, UBound(vetReceitas, 2)) = vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor) * dblProporcao
                        End If
                        
                        'Soma total das receitas da parcela
                        dblValorTotalReceitas = dblValorTotalReceitas + FormatCurrency(vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor) * dblProporcao, 2)
                        
                        adoReceitas.MoveNext
                        
                    Next
                
                End If
                
                intExisteReceita = -1
                
                'Vamos criar as receitas Multa, Juros e Correcao, caso nao seja zero
                If vetParcelasParaAcordo(VALOR_MULTA, intFor) > 0 Then
                    'Vamos verificar se a receita ja existe no array
                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                        If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaMulta Then
                            intExisteReceita = intForReceitasArray
                            Exit For
                        End If
                    Next
                    If intExisteReceita > -1 Then
                        vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + CCur(vetParcelasParaAcordo(VALOR_MULTA, intFor))
                        intExisteReceita = -1
                    Else
                        ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                        vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaMulta
                        vetReceitas(1, UBound(vetReceitas, 2)) = vetParcelasParaAcordo(VALOR_MULTA, intFor)
                    End If
                End If
                If vetParcelasParaAcordo(VALOR_JUROS, intFor) > 0 Then
                    'Vamos verificar se a receita ja existe no array
                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                        If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaJuros Then
                            intExisteReceita = intForReceitasArray
                            Exit For
                        End If
                    Next
                    If intExisteReceita > -1 Then
                        vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + CCur(vetParcelasParaAcordo(VALOR_JUROS, intFor))
                        intExisteReceita = -1
                    Else
                        ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                        vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaJuros
                        vetReceitas(1, UBound(vetReceitas, 2)) = vetParcelasParaAcordo(VALOR_JUROS, intFor)
                    End If
                End If
                If vetParcelasParaAcordo(VALOR_CORRECAO, intFor) > 0 Then
                    'Vamos verificar se a receita ja existe no array
                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                        If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaCorrecao Then
                            intExisteReceita = intForReceitasArray
                            Exit For
                        End If
                    Next
                    If intExisteReceita > -1 Then
                        vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + CCur(vetParcelasParaAcordo(VALOR_CORRECAO, intFor))
                        intExisteReceita = -1
                    Else
                        ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                        vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaCorrecao
                        vetReceitas(1, UBound(vetReceitas, 2)) = vetParcelasParaAcordo(VALOR_CORRECAO, intFor)
                    End If
                End If
                
                'Soma total das receitas da parcela + multa, juros e correcao
                dblValorTotalReceitas = dblValorTotalReceitas + vetParcelasParaAcordo(VALOR_MULTA, intFor) + vetParcelasParaAcordo(VALOR_JUROS, intFor) + vetParcelasParaAcordo(VALOR_CORRECAO, intFor)
                
                'Vamos verificar se a diferenca do valor total das receitas com o valor total da parcela ja atualizado
                dblValorDiferencaReceitas = CCur(vetParcelasParaAcordo(VALOR_TOTAL, intFor)) - CCur(dblValorTotalReceitas)
                
                'Caso exista diferenca vamos jogar na primeira receita com valor maior que zero
                If dblValorDiferencaReceitas <> 0 Then
                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                        If vetReceitas(VALOR_RECEITA, intForReceitasArray) > 0 Then
                            vetReceitas(VALOR_RECEITA, intForReceitasArray) = vetReceitas(VALOR_RECEITA, intForReceitasArray) + dblValorDiferencaReceitas
                            Exit For
                        End If
                    Next
                End If
                
                dblValorTotalReceitas = 0
                
            Else
                ExibeMensagem "Não foi(ram) encontrada(s) receita(s) para uma das parcelas originárias do acordo. A gravação não foi concluída."
                gobjBanco.ExecutaRollbackTrans
                Exit Function
            End If
            
        End If
        
    Next
                                                                                                                                                                         
    'Vamos somar o acrescimo na receita de Juros
    'Vamos verificar se a receita ja existe no array
    If Val(gstrConvVrParaSql(gstrConvVrDoSql(txtdblAcrescimos, 2, , True))) > 0 Then
        For intForReceitasArray = 0 To UBound(vetReceitas, 2)
            If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaJuros Then
                intExisteReceita = intForReceitasArray
                Exit For
            End If
        Next
        If intExisteReceita > -1 Then
            vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + Val(gstrConvVrParaSql(gstrConvVrDoSql(txtdblAcrescimos)))
            intExisteReceita = -1
        Else
            ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
            vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaJuros
            vetReceitas(1, UBound(vetReceitas, 2)) = Val(gstrConvVrParaSql(gstrConvVrDoSql(txtdblAcrescimos, 2, , True)))
        End If
    End If
    
    'Vamos somar o acrescimo por parcela na receita de Juros
    'Vamos verificar se a receita ja existe no array
    If Val(gstrConvVrParaSql(gstrConvVrDoSql(txtdblValor - dblValorAcordoOriginal, 2, , True))) > 0 Then
        For intForReceitasArray = 0 To UBound(vetReceitas, 2)
            If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaJuros Then
                intExisteReceita = intForReceitasArray
                Exit For
            End If
        Next
        If intExisteReceita > -1 Then
            vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + Val(gstrConvVrParaSql(gstrConvVrDoSql(txtdblValor - dblValorAcordoOriginal)))
            intExisteReceita = -1
        Else
            ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
            vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaJuros
            vetReceitas(1, UBound(vetReceitas, 2)) = Val(gstrConvVrParaSql(gstrConvVrDoSql(txtdblValor - dblValorAcordoOriginal, 2, , True)))
        End If
    End If
    
    'Vamos definir a receita de Honorarios
    'Vamos verificar se a receita ja existe no array
    If Val(gstrConvVrParaSql(gstrConvVrDoSql(dblValorHonorarios, 2, , True))) > 0 Then
        For intForReceitasArray = 0 To UBound(vetReceitas, 2)
            If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaHonorario Then
                intExisteReceita = intForReceitasArray
                Exit For
            End If
        Next
        If intExisteReceita > -1 Then
            vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + Val(gstrConvVrParaSql(gstrConvVrDoSql(dblValorHonorarios)))
            intExisteReceita = -1
        Else
            ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
            vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaHonorario
            vetReceitas(1, UBound(vetReceitas, 2)) = Val(gstrConvVrParaSql(gstrConvVrDoSql(dblValorHonorarios, 2, , True)))
        End If
    End If
    
    'Vamos definir os valores das parcelas do acordo
    dblValorParcela = FormatCurrency(((CCur(txtdblValor) - dblAcrescimoDescProv) + Val(gstrConvVrParaSql(gstrConvVrDoSql(txtdblAcrescimos)))) / txt_QtdeParcelas, 2)
    
    'Vamos verificar se ha diferenca
    dblValorDiferenca = FormatCurrency(((CCur(txtdblValor) - dblAcrescimoDescProv) + Val(gstrConvVrParaSql(gstrConvVrDoSql(txtdblAcrescimos)))) - (dblValorParcela * txt_QtdeParcelas), 2)
    
    'Vamos gravar as parcelas pre-definidas na tabela TBLLANCAMENTOVALOR
    For intFor = 1 To txt_QtdeParcelas
    
        'Vamos adicionar o valor do acrescimo na parcela
        If intFor >= intAcrescimoParcIniDescProv And dblAcrescimoDescProv > 0 Then
            dblValorParcela = dblValorParcela + (dblAcrescimoDescProv / intQtdeParcelasAcrescimo)
        End If
        
        strSql = "INSERT INTO " & gstrLancamentoValor & " (" & _
                 "intLancamentoAlfa, intParcela, dtmDtVencimento, dblValor, intMoeda, bitParcelaValida, dtmDtAtualizacao, lngCodUsr) VALUES (" & _
                 lngPkidLAAcordo & "," & intFor & "," & gstrConvDtParaSql(DateAdd("M", intFor - 1, txt_dtmVencimento)) & "," & gstrConvVrParaSql(dblValorParcela + dblValorDiferenca) & "," & dbcintMoeda.BoundText & ",1," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                 
        If Not gobjBanco.Execute(strSql) Then
            ExibeMensagem "A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
                     
        'Vamos preencher o parametro de todas as parcelas para a impressao do carne
        strParcelasAcordo = strParcelasAcordo & intFor & ","
        
         'Vamos gravar as receitas na tabela TBLLANCAMENTORECEITA
         For intForReceitasArray = 0 To UBound(vetReceitas, 2)
             
             'Valores sem participacao de acrescimos dos descontos provisorios
             dblValorReceita = gstrConvVrDoSql(vetReceitas(VALOR_RECEITA, intForReceitasArray), 2)
             
             'Vamos verificar se existem acrescimos no desconto provisorio
             If dblAcrescimoDescProv > 0 And vetReceitas(RECEITA, intForReceitasArray) = lngReceitaJuros Then
                 'Vamos subtrair o acrescimo de desconto provisorio da receita de multa, nas parcelas que nao foram acrescentadas
                 If intFor < intAcrescimoParcIniDescProv Then
                     dblValorReceita = gstrConvVrDoSql(dblValorReceita - (dblAcrescimoDescProv), 2)
                 Else
                     'Vamos somar na receita o acrescimo das parcelas que nao possuem
                     dblValorReceita = gstrConvVrDoSql(dblValorReceita + ((dblAcrescimoDescProv / intQtdeParcelasAcrescimo) * (txt_QtdeParcelas - intQtdeParcelasAcrescimo)), 2)
                 End If
             End If
             
             strSql = "INSERT INTO " & gstrLancamentoReceita & " (" & _
                      "intLancamentoValor, intReceita, dblValor, dtmDtAtualizacao, lngCodUsr) VALUES (" & _
                      glngRetornaPkidTabelaPai("seqtblLancamentoValor", gstrLancamentoValor) & "," & vetReceitas(RECEITA, intForReceitasArray) & "," & gstrConvVrParaSql(gstrConvVrDoSql(dblValorReceita / txt_QtdeParcelas, 2, , True)) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
             
             If Not gobjBanco.Execute(strSql) Then
                 ExibeMensagem "A gravação não foi concluída."
                 gobjBanco.ExecutaRollbackTrans
                 Exit Function
             End If
                                 
             'Vamos somar todas as receitas para depois verificar diferenca por parcela do acordo
             dblValorTotalReceitas = dblValorTotalReceitas + gstrConvVrDoSql(dblValorReceita / txt_QtdeParcelas, 2)
             
         Next
        
         'Vamos verificar se ha diferenca da soma das receitas com o valor da parcela
         dblValorDiferencaReceitas = CCur((dblValorParcela + dblValorDiferenca)) - CCur(dblValorTotalReceitas)
         
         'Caso exista diferenca vamos jogar na primeira receita
         If dblValorDiferencaReceitas <> 0 Then
             If Not gobjBanco.Execute("UPDATE " & gstrLancamentoReceita & " SET dblValor = dblValor + " & gstrConvVrParaSql(dblValorDiferencaReceitas) & " WHERE Pkid = " & glngRetornaPkidTabelaPai("seqtblLancamentoReceita", gstrLancamentoReceita)) Then
                 ExibeMensagem "Não foi possível atualizar diferença de valor na Receita. A gravação não foi concluída."
                 gobjBanco.ExecutaRollbackTrans
                 Exit Function
             End If
         End If
         
         dblValorTotalReceitas = 0
            
        'So vamos aplicar a diferenca na primeira parcela
        dblValorDiferenca = 0
        
        'Vamos redefinir os valores das parcelas do acordo
        dblValorParcela = FormatCurrency(((CCur(txtdblValor) - dblAcrescimoDescProv) + Val(gstrConvVrParaSql(gstrConvVrDoSql(txtdblAcrescimos)))) / txt_QtdeParcelas, 2)
        
    Next
    
    strParcelasAcordo = Mid(strParcelasAcordo, 1, Len(strParcelasAcordo) - 1)
    
    If blnVerificaDuplicidade(lngPkidLAAcordo, strComposicaoDaReceita) Then
        gobjBanco.ExecutaCommitTrans
    Else
        strInscricaoAux = mskstrInscricao
        mskstrInscricao = ProximaInscricaoAcordo
        ExibeMensagem "O acordo " & strInscricaoAux & "/" & txtintExercicio & " acabou de ser cadastrado por outro usuário." & _
                      Chr(13) & "O novo nº do acordo é " & mskstrInscricao & "/" & txtintExercicio
        gobjBanco.ExecutaRollbackTrans
        GoTo NovaGravacao
    End If
    
    LeDaTabelaParaObj "", tdb_Acordos, strQuery
     
    mblnAlterando = True
    
    Set gobjBanco = New clsBanco

    blnVBModal = False
    
    If chk_Carne.Value = vbChecked Then
         rptCapaCarneAcordo.strParcelasSelecionadas = strParcelasAcordo
         ImprimeRelatorio rptCapaCarneAcordo, gstrQueryCarneAcordo(mskstrInscricao.Text & txtintExercicio, "", strParcelasAcordo, False)
    End If
    
    DoEvents
    
    If chk_Termo.Value = vbChecked Then
        ImprimirTermo lngPkidLAAcordo
    End If
    
    DoEvents
    
    blnVBModal = True
    
    Exit Function
    
Problema_Na_Rotina:

   ExibeDetalheErro "Erro na rotina de Gravação do Acordo."
   gobjBanco.ExecutaRollbackTrans
   
End Function

Private Function ProximaInscricaoAcordo() As String
    Dim adoResultado As New ADODB.Recordset
    
    If bytDBType = SQLServer Then
        If gobjBanco.CriaADO(strQueryProximoAcordo, 20, adoResultado) Then
            If Not adoResultado.EOF Then
                If Not IsNull(adoResultado("ProximaInscricao").Value) Then
                    ProximaInscricaoAcordo = Format(adoResultado("ProximaInscricao").Value + 1, String(gintRetornaTamanhoMascara(TYP_ACORDO) - 4, "0"))
                Else
                    ProximaInscricaoAcordo = Format(1, String(gintRetornaTamanhoMascara(TYP_ACORDO) - 4, "0"))
                End If
            Else
                ProximaInscricaoAcordo = Format(1, String(gintRetornaTamanhoMascara(TYP_ACORDO) - 4, "0"))
            End If
        Else
            ProximaInscricaoAcordo = "0"
        End If
    Else
        If gobjBanco.CriaADO("SELECT seqNumeroAcordo.NextVal as ProximaInscricao FROM dual", 5, adoResultado) Then
            If Not adoResultado.EOF And Not IsNull(adoResultado("ProximaInscricao").Value) Then
                ProximaInscricaoAcordo = Format(adoResultado("ProximaInscricao").Value, String(gintRetornaTamanhoMascara(TYP_ACORDO) - 4, "0"))
            Else
                ProximaInscricaoAcordo = Format(1, String(gintRetornaTamanhoMascara(TYP_ACORDO) - 4, "0"))
            End If
        End If
    End If
End Function

Public Sub InicializaArrayParcelas(vetArray() As String)
    vetParcelasParaAcordo = vetArray
    vetParcelasParaAcordoAux = vetArray
    blnAtualizacao = True
End Sub

Private Sub GeraTotalIndexador()
    If (CDbl(gstrConvVrDoSql(txtdblVlIndexador, , , True)) > 0) And (Trim(txt_strIndexador.Text) <> "") Then
        lbl_Indexador.Caption = "Total em " & Trim(txt_strIndexador.Text)
        txtdblTotalIndexador.Text = gstrConvVrDoSql((CDbl(gstrConvVrDoSql(txtdblValor, , , True)) + CDbl(gstrConvVrDoSql(txtdblAcrescimos, , , True))) / CDbl(gstrConvVrDoSql(txtdblVlIndexador, 6, , True)), 6)
    End If
End Sub

Private Sub PreencheAbreviaturaIndexador()
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select Pkid, strAbreviatura From " & gstrIndexadorEconomico & " Order By strAbreviatura"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Do While Not adoResultado.EOF
                'cbo_strIndexador.AddItem gstrENulo(adoResultado!strAbreviatura)
                'cbo_strIndexador.ItemData(cbo_strIndexador.NewIndex) = gstrENulo(adoResultado!Pkid)
                'cbo_strIndexador.Tag = gstrENulo(adoResultado!Pkid)
                adoResultado.MoveNext
            Loop
        End If
    End If
End Sub

Public Sub PreencheValorIndexador()
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Dim intWhile        As Integer
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "FAV.DBLVALOR, IE.strAbreviatura "
    strSql = strSql & "From "
    strSql = strSql & gstrIndexadorEconomico & " IE, "
    strSql = strSql & gstrFormaAtualizacaoValor & " FAV, "
    strSql = strSql & "(SELECT intIndexadorEconomico "
    strSql = strSql & "FROM " & gstrParametrosTributario & " "
    strSql = strSql & "WHERE pkID = (SELECT MIN(pkID) FROM "
    strSql = strSql & gstrParametrosTributario & ")) IEID "
    strSql = strSql & "WHERE "
    strSql = strSql & "IE.Pkid = FAV.Intindexadoreconomico AND "
    
    If bytDBType = EDatabases.SQLServer Then
       strSql = strSql & gstrCONVERT(CDT_VARCHAR, "FAV.Dtmdata" & ",103") & " = " & gstrCONVERT(CDT_VARCHAR, strGETDATE & ",103") & " AND "
    Else
       strSql = strSql & "FAV.Dtmdata = " & gstrCONVERT(CDT_VARCHAR, strGETDATE & ", 'DD/MM/YYYY'") & " AND "
    End If
    
    strSql = strSql & "IE.Pkid = IEID.intIndexadorEconomico "
    

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txtdblVlIndexador = gstrConvVrDoSql(gstrENulo(adoResultado!dblValor), 6)
            txt_strIndexador.Text = IIf(IsNull(adoResultado!Strabreviatura), "", adoResultado!Strabreviatura)
            GeraTotalIndexador
        Else
            strSql = ""
            strSql = strSql & "Select "
            strSql = strSql & "FAV.DBLVALOR, IE.strAbreviatura "
            strSql = strSql & "From "
            strSql = strSql & gstrIndexadorEconomico & " IE, "
            strSql = strSql & gstrFormaAtualizacaoValor & " FAV, "
            strSql = strSql & "(SELECT intIndexadorEconomico "
            strSql = strSql & "FROM " & gstrParametrosTributario & " "
            strSql = strSql & "WHERE pkID = (SELECT MIN(pkID) FROM "
            strSql = strSql & gstrParametrosTributario & ")) IEID "
            strSql = strSql & "WHERE "
            strSql = strSql & "IE.Pkid = FAV.Intindexadoreconomico AND "
            strSql = strSql & "FAV.Dtmdata = " & gstrConvDtParaSql(gstrDataFormatada("01/" & Month(Date) & "/" & Year(Date))) & " AND "
            strSql = strSql & "IE.Pkid = IEID.intIndexadorEconomico "
            
            If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
                If Not adoResultado.EOF Then
                    txtdblVlIndexador = gstrConvVrDoSql(gstrENulo(adoResultado!dblValor), 6)
                    txt_strIndexador.Text = IIf(IsNull(adoResultado!Strabreviatura), "", adoResultado!Strabreviatura)
                    GeraTotalIndexador
                Else
                    txt_strIndexador.Text = ""
                    txtdblVlIndexador.Text = gstrConvVrDoSql("", 6, , True)
                    txtdblTotalIndexador.Text = ""
                    ExibeMensagem "Não encontrado valor do indexador para data atual."
                End If
            End If
        End If
    End If
    
    
End Sub

Private Function strQueryProximoAcordo() As String

    Dim strSql As String
    
    'Alteração feita por Hugo

'    strSQL = "SELECT MAX(" & gstrCONVERT(CDT_INT, strSUBSTRING & "(LA.strInscricao," & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & (gintRetornaTamanhoMascara(TYP_ACORDO) - 4) & ")") & ") ProximaInscricao "
'    strSQL = strSQL & "FROM " & gstrLancamentoAlfa & " LA "
'    strSQL = strSQL & "WHERE LA.intExercicio = " & Year(gstrDataDoSistema)
'    strSQL = strSQL & " AND LA.intUtilizacao = " & TYP_ACORDO
    
    
    strSql = "SELECT "
    'strSql = strSql & " TOP 1 CONVERT(int,SUBSTRING(LA.strInscricao,11,6)) ProximaInscricao "
    strSql = strSql & " TOP 1 " & gstrCONVERT(CDT_numeric, strSUBSTRING & "(LA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4) & ") ProximaInscricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA "
    strSql = strSql & "WHERE "
    strSql = strSql & "LA.intExercicio = " & Year(gstrDataDoSistema)
    strSql = strSql & " AND LA.intUtilizacao = " & TYP_ACORDO
    strSql = strSql & " Order by "
    strSql = strSql & " LA.strInscricao Desc "

    strQueryProximoAcordo = strSql

End Function

Private Function blnVerificaParametrosParaParcelamento(Optional dblValorParaParcelar As Double, Optional dblValorPorParcela As Double, Optional intQtdeParcelas As Integer) As Boolean
    Dim adoResultado As ADODB.Recordset
    Dim strSql       As String
    Dim strMsg       As String
    
    blnVerificaParametrosParaParcelamento = False
        
    strMsg = Space$(0)
    
    If dbc_intDescProvisorios.MatchedWithList = False Then
        strSql = ""
        strSql = strSql & "SELECT "
        strSql = strSql & "PT.dblValorMinimoParaParcelamento, PT.dblValorMinimoPorParcela, PT.intQtdeMaximaParcelas "
        strSql = strSql & "FROM "
        strSql = strSql & " tblParametrosTributario PT "
    
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
            
                'Vamos verificar o valor a parcelar
                If dblValorParaParcelar > 0 And Not IsNull(adoResultado("dblValorMinimoParaParcelamento").Value) Then
                    If dblValorParaParcelar < adoResultado("dblValorMinimoParaParcelamento").Value Then
                        strMsg = "O valor para parcelamento é inferior ao valor mínimo parametrizado, que é de: " & gstrConvVrDoSql(adoResultado("dblValorMinimoParaParcelamento").Value) & "."
                    End If
                End If
                
                'Vamos verificar o valor por parcela
                If dblValorPorParcela > 0 And Not IsNull(adoResultado("dblValorMinimoPorParcela").Value) Then
                    If dblValorPorParcela < adoResultado("dblValorMinimoPorParcela").Value Then
                        strMsg = strMsg & Chr(13) & "O valor da parcela é inferior ao valor mínimo parametrizado, que é de: " & gstrConvVrDoSql(adoResultado("dblValorMinimoPorParcela").Value) & "."
                    End If
                End If
                
                'Vamos verificar a qtde de parcelas
                If intQtdeParcelas > 0 And Not IsNull(adoResultado("intQtdeMaximaParcelas").Value) Then
                    If intQtdeParcelas > adoResultado("intQtdeMaximaParcelas").Value Then
                        strMsg = strMsg & Chr(13) & "A quantidade de parcelas é superior à quantidade máxima parametrizada, que é de: " & adoResultado("intQtdeMaximaParcelas").Value & "."
                    End If
                End If
                
                If Len(strMsg) > 0 Then
                    ExibeMensagem strMsg
                    Exit Function
                End If
                
            End If
        End If
        
        adoResultado.Close
    Else
        strSql = "Select * From " & gstrDescontosProvisorios & " Where pkid = " & dbc_intDescProvisorios.BoundText
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
            
                'Vamos verificar o valor por parcela
                If dblValorPorParcela > 0 And Not IsNull(adoResultado("dblvalorminimo").Value) Then
                    If dblValorPorParcela < adoResultado("dblvalorminimo").Value Then
                        strMsg = strMsg & Chr(13) & "O valor da parcela é inferior ao valor mínimo parametrizado, que é de: " & gstrConvVrDoSql(adoResultado("dblvalorminimo").Value) & "."
                    End If
                End If
                
                'Vamos verificar a qtde de parcelas
                If intQtdeParcelas > 0 And Not IsNull(adoResultado("intparcela").Value) Then
                    If intQtdeParcelas > adoResultado("intparcela").Value Then
                        strMsg = strMsg & Chr(13) & "A quantidade de parcelas é superior à quantidade máxima parametrizada, que é de: " & adoResultado("intparcela").Value & "."
                    End If
                End If
                
                If Len(strMsg) > 0 Then
                    ExibeMensagem strMsg
                    Exit Function
                End If
            
            End If
        End If
    End If
    
    blnVerificaParametrosParaParcelamento = True
    
End Function

Private Function blnVerificaAutoNumeracao() As Boolean
    Dim adoResultado As ADODB.Recordset
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Select bitAutoNumeracao From " & gstrItens & " WHERE intcodigo = " & gintCodSeguranca
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        blnVerificaAutoNumeracao = Val(gstrENulo(adoResultado!bitAutoNumeracao))
    End If
    
End Function

Private Sub CarregaEndereco()
    Dim adoResultado As ADODB.Recordset
    Dim strSql As String

    strSql = strSql & "SELECT "
    strSql = strSql & "C.strLogradouroC, "
    strSql = strSql & "C.intNumeroC, "
    strSql = strSql & "C.strComplementoC, "
    strSql = strSql & "C.strBairroC, "
    strSql = strSql & "(Select strdescricao From tblmunicipio where pkid = C.Intmunicipioc) strMunicipio, "
    strSql = strSql & "(Select Strsigla from tblUF where pkid = C.Intufc) strUf, "
    strSql = strSql & "C.intCepC "
    strSql = strSql & "FROM "
    strSql = strSql & "Tblcontribuinte C "
    strSql = strSql & "WHERE "
    strSql = strSql & "PKId = " & dbcstrNomeProprietario.BoundText
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            'txtstrLogradouro = gstrENulo(adoResultado!strlogradouroc)
            'txtstrNumero = gstrENulo(adoResultado!intNumeroC)
            'txtstrComplemento = gstrENulo(adoResultado!strComplementoC)
            'txtstrBairro = gstrENulo(adoResultado!strBairroC)
            'txtstrMunicipio = gstrENulo(adoResultado!STRMUNICIPIO)
            'txtstrUf = gstrENulo(adoResultado!STRUF)
            'txtintCep = gstrCEPFormatado(gstrENulo(adoResultado!intcepc))
            
            txtstrLogradouroC = gstrENulo(adoResultado!strLogradouroC)
            txtstrNumeroC = gstrENulo(adoResultado!intNumeroC)
            txtstrComplementoC = gstrENulo(adoResultado!strComplementoC)
            txtstrBairroC = gstrENulo(adoResultado!strBairroC)
            txtstrMunicipioC = gstrENulo(adoResultado!STRMUNICIPIO)
            txtstrUFC = gstrENulo(adoResultado!STRUF)
            txtintCEPC = gstrCEPFormatado(gstrENulo(adoResultado!INTCEPC))
        End If
    End If
    
End Sub

Private Function strDesctoProvisorio() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Select Pkid, strdescricao from " & gstrDescontosProvisorios & " Where "
    strSql = strSql & "intParcela > 1 and "
    
    'If bytDBType = SQLServer Then
        'strSql = strSql & gstrCONVERT(CDT_DATETIME, "LA.intExercicio") & " Between dtmdtinicial and dtmdtfinal"
        strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & " Between dtmdtinicial and dtmdtfinal"
    'Else
    '    strSql = strSql & "'" & gstrDataDoSistema & "' Between dtmdtinicial and dtmdtfinal"
    'End If
        
    strDesctoProvisorio = strSql
End Function
Private Sub PreencheAnistia(lngPkid As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Dim intFor          As Integer
    Dim dblvalorTot     As Double
    Dim dblValorTotHon  As Double
    
    strSql = strSql & "Select * From " & gstrDescontosProvisorios & " Where pkid = " & lngPkid
    vetParcelasParaAcordo = vetParcelasParaAcordoAux
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_strAnistia = gstrENulo(adoResultado!strDescricao)
            txt_strAnistiaLegislacao = gstrENulo(adoResultado!strLegislacao)
            For intFor = 0 To UBound(vetParcelasParaAcordo, 2)
                vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor) = CDbl(vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor)) * (100 - adoResultado!Dblvalororiginal) / 100
                vetParcelasParaAcordo(VALOR_MULTA, intFor) = CDbl(vetParcelasParaAcordo(VALOR_MULTA, intFor)) * (100 - adoResultado!dblMulta) / 100
                vetParcelasParaAcordo(VALOR_JUROS, intFor) = CDbl(vetParcelasParaAcordo(VALOR_JUROS, intFor)) * (100 - adoResultado!dblJuros) / 100
                vetParcelasParaAcordo(VALOR_TOTAL, intFor) = CCur(vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor)) + CCur(vetParcelasParaAcordo(VALOR_MULTA, intFor)) + CCur(vetParcelasParaAcordo(VALOR_JUROS, intFor)) + CCur(vetParcelasParaAcordo(VALOR_CORRECAO, intFor))
                dblvalorTot = dblvalorTot + vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor) + vetParcelasParaAcordo(VALOR_MULTA, intFor) + vetParcelasParaAcordo(VALOR_JUROS, intFor) + vetParcelasParaAcordo(VALOR_CORRECAO, intFor)
                'Vamos somar o total de parcelas com honorarios para aplicar o desconto nele tambem
                If vetParcelasParaAcordo(EXECUTIVO, intFor) = True Then
                    dblValorTotHon = dblValorTotHon + vetParcelasParaAcordo(VALOR_PRINCIPAL, intFor) + vetParcelasParaAcordo(VALOR_MULTA, intFor) + vetParcelasParaAcordo(VALOR_JUROS, intFor) + vetParcelasParaAcordo(VALOR_CORRECAO, intFor)
                End If
            Next
        End If
    End If
    
    'Vamos redefinir o valor dos honorarios com os valores ja com desconto
    dblValorHonorarios = dblCalculaEncargos(BIT_HONORARIOS, dblValorTotHon, "Acordo")
    
    'Vamos somar o Honorario no total, sem aplicar o desconto provisorio
    dblvalorTot = dblvalorTot + dblValorHonorarios
    
    txtdblValor = gstrConvVrDoSql(dblvalorTot, 2, , True)
    
    'Vamos aplicar a anistia no valor original
    dblValorAcordoOriginal = dblvalorTot
    txt_QtdeParcelas = adoResultado!intParcela
    
ReaplicarAcrescimo:

    'Vamos aplicar o acrescimo por parcela, caso seja parametrizado e nao esteja com desconto provisorio
    If Not dbc_intDescProvisorios.MatchedWithList Then
        AplicarAcrescimoPorParcela
    End If
    
    If (CDbl(CCur(txtdblValor) + CCur(gstrConvVrDoSql(txtdblAcrescimos, , , True))) / txt_QtdeParcelas) < adoResultado!dblvalorminimo Then
        txt_QtdeParcelas = Int(CDbl(txtdblValor) / adoResultado!dblvalorminimo)
        GoTo ReaplicarAcrescimo: 'Neste caso temos que reaplicar o acrescimo pois mudou o numero de parcelas
        'txt_dblValorParcela = gstrConvVrDoSql((CDbl(txtdblValor) / txt_QtdeParcelas), , , True)
    Else
        'txt_QtdeParcelas = adoResultado!intParcela
        txt_dblValorParcela = gstrConvVrDoSql((CDbl(CCur(txtdblValor) + CCur(gstrConvVrDoSql(txtdblAcrescimos, , , True))) / txt_QtdeParcelas), , , True)
    End If
    txt_dtmVencimento.Text = gstrDataDoSistema
    
    'Vamos aplicar o acrescimo de juros do desconto provisorio
    AplicarAcrescimoDesctoProvisorio
    
End Sub

Private Function blnExluiAcordo() As Boolean
    Dim adoResultado            As New ADODB.Recordset
    Dim adoAcordo               As New ADODB.Recordset
    Dim adoCritica              As New ADODB.Recordset
    Dim strSql                  As String
    Dim strAcordosParaConsulta  As String
    Dim strInscricoes           As String

    On Error GoTo Problema_Na_Rotina

    blnExluiAcordo = False
    
    'Vamos obter os valores das parcelas da inscricao selecionada
    Set gobjBanco = New clsBanco
        
    If Val(txtPKId) <= 0 Then
        ExibeMensagem "Não foi selecionado Acordo no Grid."
        Exit Function
    End If
            
    'Vamos obter os Pkids das inscricoes para fazer consulta de acordos
    strSql = strSql & "SELECT "
    strSql = strSql & "LA.Pkid PkidLA, "
    strSql = strSql & "AC.Pkid PkidAC, "
    strSql = strSql & "LV.Intlancamentoalfaacordo "
    strSql = strSql & "FROM "
    strSql = strSql & gstrLancamentoAlfa & " LA, "
    strSql = strSql & gstrLancamentoValor & " LV, "
    strSql = strSql & gstrAcordo & " AC "
    strSql = strSql & "WHERE "
    strSql = strSql & "LA.Pkid = AC.intLancamentoAlfa AND "
    strSql = strSql & "LA.Pkid = LV.Intlancamentoalfa AND "
    strSql = strSql & "AC.Pkid = " & Val(txtPKId) & " "
    strSql = strSql & "Group By "
    strSql = strSql & "LA.Pkid, "
    strSql = strSql & "AC.Pkid, "
    strSql = strSql & "LV.Intlancamentoalfaacordo "
    strSql = strSql & "Order by LV.Intlancamentoalfaacordo "
    strSql = strSql & IIf(bytDBType = Oracle, "ASC ", "DESC ")

    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            'Verifica se o acordo informado faz parte de outro acordo
            If Trim(gstrENulo(adoResultado!intlancamentoalfaacordo)) <> "" Then
                strSql = "Select "
                strSql = strSql & strSUBSTRING & "(strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ") NumeroInscricao, "
                strSql = strSql & "intExercicio Exercicio "
                strSql = strSql & "From " & gstrLancamentoAlfa & " Where Pkid = " & gstrENulo(adoResultado!intlancamentoalfaacordo)
                If gobjBanco.CriaADO(strSql, 5, adoCritica) Then
                    If Not adoCritica.EOF Then
                        ExibeMensagem "O Acordo selecionado é composição do Acordo " & gstrENulo(adoCritica!NumeroInscricao) & "/" & gstrENulo(adoCritica!EXERCICIO) & " não podendo completar a operação."
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
            
                'Vamos verificar se o acordo tem movimentos bancarios
                strSql = "Select "
                strSql = strSql & strSUBSTRING & "(strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ") NumeroInscricao, " & " LA.intExercicio Exercicio "
                strSql = strSql & "From " & gstrLancamentoAlfa & " LA, "
                strSql = strSql & gstrLancamentoValor & " LV, "
                strSql = strSql & gstrMovimentoBancario & " MB "
                strSql = strSql & "Where LA.Pkid = LV.Intlancamentoalfa AND "
                strSql = strSql & "LV.Pkid = MB.Intlancamentovalor AND "
                strSql = strSql & "LA.Pkid = " & Val(gstrENulo(adoResultado!PkidLA))
                If gobjBanco.CriaADO(strSql, 5, adoCritica) Then
                    If adoCritica.RecordCount > 0 Then
                        ExibeMensagem "Não é possível cancelar o acordo " & gstrENulo(adoCritica!NumeroInscricao) & "/" & gstrENulo(adoCritica!EXERCICIO) & " pois há movimentos bancários para o mesmo."
                        Exit Function
                    Else
                        'Vamos verificar se o acordo tem pagamentos
                        strSql = "Select "
                        strSql = strSql & strSUBSTRING & "(strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ") NumeroInscricao, " & " LA.intExercicio Exercicio "
                        strSql = strSql & "From " & gstrLancamentoAlfa & " LA, "
                        strSql = strSql & gstrLancamentoValor & " LV, "
                        strSql = strSql & gstrLancamentoPagamento & " LP "
                        strSql = strSql & "Where LA.Pkid = LV.Intlancamentoalfa AND "
                        strSql = strSql & "LV.Pkid = LP.Intlancamentovalor AND "
                        strSql = strSql & "LA.Pkid = " & Val(gstrENulo(adoResultado!PkidLA))
                        If gobjBanco.CriaADO(strSql, 5, adoCritica) Then
                            If adoCritica.RecordCount > 0 Then
                                If gblnExclusaoGravacaoOk("E", "O Acordo " & gstrENulo(adoCritica!NumeroInscricao) & "/" & gstrENulo(adoCritica!EXERCICIO) & " possui pagamentos." & Chr(13) & "Deseja excluir o acordo?", True) Then
                                    GoTo Gravar
                                Else
                                    Exit Function
                                End If
                            End If
                        Else
                            Exit Function
                        End If
                    End If
                Else
                    Exit Function
                End If
            End If
        Else
            ExibeMensagem "Não foi encontrado nenhum acordo com esta Inscrição."
            Exit Function
        End If
    Else
        Exit Function
    End If
        
    If gblnExclusaoGravacaoOk("E", "Deseja realmente cancelar o acordo " & mskstrInscricao.Text & "/" & txtintExercicio, True) Then
Gravar:
        'Vamos excluir da tabela Lancamento Receita
        If Not gobjBanco.Execute(" DELETE FROM " & gstrLancamentoReceita & " WHERE intLancamentoValor IN (SELECT Pkid FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & adoResultado("PkidLA").Value & ")") Then
            ExibeMensagem "Não foi possível excluir registro(s) referente a Inscrição selecionada."
            Exit Function
        End If
        
        'Vamos excluir da tabela Lancamento Guias
        If Not gobjBanco.Execute(" DELETE FROM " & gstrLancamentoGuias & " WHERE intLancamentoValor IN (SELECT Pkid FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & adoResultado("PkidLA").Value & ")") Then
            ExibeMensagem "Não foi possível excluir registro(s) referente a Inscrição selecionada."
            Exit Function
        End If
        
        'Vamos excluir da tabela Lancamento Pagamento
        If Not gobjBanco.Execute(" DELETE FROM " & gstrLancamentoPagamento & " WHERE intLancamentoValor IN (SELECT Pkid FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & adoResultado("PkidLA").Value & ")") Then
            ExibeMensagem "Não foi possível excluir registro(s) referente a Inscrição selecionada."
            Exit Function
        End If
        
        'Vamos excluir da tabela Lancamento Valor
        If Not gobjBanco.Execute(" DELETE FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & adoResultado("PkidLA").Value) Then
            ExibeMensagem "Não foi possível excluir registro(s) referente a Inscrição selecionada."
            Exit Function
        End If
        
        'Vamos excluir da tabela Acordo Debitos
        If Not gobjBanco.Execute(" DELETE FROM " & gstrAcordoDebitos & " WHERE intAcordo = " & adoResultado("PkidAC").Value) Then
            ExibeMensagem "Não foi possível excluir registro(s) referente a Inscrição selecionada."
            Exit Function
        End If
        
        'Vamos excluir da tabela Acordos
        If Not gobjBanco.Execute(" DELETE FROM " & gstrAcordo & " WHERE Pkid = " & adoResultado("PkidAC").Value) Then
            ExibeMensagem "Não foi possível excluir registro(s) referente a Inscrição selecionada."
            Exit Function
        End If
        
        'Vamos desvincular o acordo da tabela Lancamento Valor
        If Not gobjBanco.Execute(" UPDATE " & gstrLancamentoValor & " SET intLancamentoAlfaAcordo = Null WHERE intLancamentoAlfaAcordo = " & adoResultado("PkidLA").Value) Then
            ExibeMensagem "Não foi possível excluir registro(s) referente a Inscrição selecionada."
            Exit Function
        End If
        
        'Vamos excluir da tabela Lancamento Alfa
        If Not gobjBanco.Execute(" DELETE FROM " & gstrLancamentoAlfa & " WHERE Pkid = " & adoResultado("PkidLA").Value) Then
            ExibeMensagem "Não foi possível excluir registro(s) referente a Inscrição selecionada."
            Exit Function
        End If
    Else
        Exit Function
    End If
    blnExluiAcordo = True
    
    Exit Function

Problema_Na_Rotina:
    ExibeMensagem "Não foi possível concluir a operação."
    
End Function

Private Function blnVerificaDuplicidade(lngPkid As Long, strComposicao As String) As Boolean
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    blnVerificaDuplicidade = True
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "La.Strinscricao, "
    strSql = strSql & "La.Strcomposicaodareceita, "
    strSql = strSql & "La.Intexercicio, "
    strSql = strSql & "La.strnumeroaviso "
    strSql = strSql & "From "
    strSql = strSql & gstrLancamentoAlfa & " LA "
    strSql = strSql & "Where "
    strSql = strSql & "strInscricao ='" & UCase(String(gintLenInscricao - Len(mskstrInscricao & txtintExercicio), "0") & mskstrInscricao & txtintExercicio) & "' AND "
    strSql = strSql & "LA.strcomposicaodareceita ='" & strComposicao & "' AND "
    strSql = strSql & "LA.Intexercicio = " & txtintExercicio & " AND "
    strSql = strSql & "LA.Strnumeroaviso = '" & UCase(String(gintLenNumAviso - Len(mskstrInscricao), "0") & mskstrInscricao) & "' "
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.RecordCount > 1 Then
            blnVerificaDuplicidade = False
        End If
    Else
        blnVerificaDuplicidade = False
    End If
    
End Function

Private Function GravarAcordoProvisorio(dtmDataInicial As Date, dtmDataFinal As Date) As Boolean

    Dim adoResultado                As New ADODB.Recordset
    Dim adoReceitas                 As New ADODB.Recordset
    Dim adoParcelas                 As New ADODB.Recordset
    Dim adoAnistia                  As New ADODB.Recordset
    Dim adoAcordo                   As New ADODB.Recordset
    
    Dim strSql                      As String
    
    Dim lngLancamentoAlfaAcordo     As Long
    
    Dim intFor                      As Integer
    Dim intForReceitas              As Integer
    Dim intForReceitasArray         As Integer
    
    Dim intExisteReceita            As Integer
    Dim dblProporcao                As Double
    
    Dim dblValorParcela             As Double
    Dim dblValorTotalReceitas       As Double
    
    Dim dblValorMulta               As Double
    Dim dblValorJuros               As Double
    Dim dblHonorarios               As Double
    Dim dblPorcHonorarios           As Double
    
    Dim lngReceitaMulta             As Long
    Dim lngReceitaJuros             As Long
    Dim lngReceitaCorrecao          As Long
    Dim lngReceitaHonorario         As Long
    
    Dim lngComposicaoDaReceita      As Long
    Dim strComposicaoDaReceita      As String
    Dim intUtilizacao               As Integer
    
On Error GoTo Problema_Na_Rotina
    
    Set gobjBanco = New clsBanco
        
    'Vamos obter a composicao de acordo
    If gobjBanco.CriaADO("SELECT Pkid, strDescricao, intUtilizacao FROM " & gstrComposicaoDaReceita & " WHERE intUtilizacao = " & TYP_ACORDO, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngComposicaoDaReceita = Space$(0) & adoResultado("Pkid").Value
            strComposicaoDaReceita = Space$(0) & adoResultado("strDescricao").Value
            intUtilizacao = Space$(0) & adoResultado("intUtilizacao").Value
        Else
            ExibeMensagem "Não foi encontrada nenhuma Composição do Tipo de Acordo. A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    Else
        ExibeMensagem "Não foi encontrada nenhuma Composição do Tipo de Acordo. A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    'Vamos obter as receitas de Multas, Juros e Correcao da Composicao de Receita
    If gobjBanco.CriaADO("SELECT intReceitaMulta, intReceitaJuros, intReceitaCorrecao FROM " & gstrParametroAtualizacao & " WHERE intExercicio = " & Year(gstrDataDoSistema) & " AND intComposicaoReceita = " & lngComposicaoDaReceita, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngReceitaMulta = Space$(0) & adoResultado("intReceitaMulta").Value
            lngReceitaJuros = Space$(0) & adoResultado("intReceitaJuros").Value
            lngReceitaCorrecao = Space$(0) & adoResultado("intReceitaCorrecao").Value
        Else
            ExibeMensagem "Não foi(ram) encontrada(s) receita(s) de Multa, Juros para a Composição de Receita. A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    Else
        ExibeMensagem "Não foi(ram) encontrada(s) receita(s) de Multa, Juros para a Composição de Receita. A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    'Vamos obter s receita de Honorarios
    If gobjBanco.CriaADO("SELECT intReceitaHonorarios FROM " & gstrParametrosTributario, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngReceitaHonorario = Space$(0) & adoResultado("intReceitaHonorarios").Value
        Else
            ExibeMensagem "Não foi encontrada receita de Honorário. A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    Else
        ExibeMensagem "Não foi encontrada receita de Honorário. A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & gstrISNULL("dblPorcHonorarios", 0) & " dblPorcHonorarios "
    strSql = strSql & "FROM "
    strSql = strSql & " tblParametrosTributario PT "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dblPorcHonorarios = adoResultado("dblPorcHonorarios").Value
        Else
            dblPorcHonorarios = 0
        End If
    End If
    
    strSql = ""
    strSql = strSql & "Select LA.PKID intLancamentoAlfaOrigem, "
    strSql = strSql & "LV.INTLANCAMENTOALFAACORDO, "
    strSql = strSql & "(Select Count(*) from tbllancamentoValor where intLancamentoAlfa = LV.INTLANCAMENTOALFAACORDO) dblTotParcelaAcordo, "
    strSql = strSql & "LA.STRINSCRICAO, "
    strSql = strSql & "LV.PKID intLancamentoValorOrigem, "
    strSql = strSql & "LV.intParcela, "
    strSql = strSql & "LA.intComposicaoDaReceita, "
    strSql = strSql & "LA.intExercicio, "
    strSql = strSql & "LV.DTMDTVENCIMENTO, "
    strSql = strSql & "AC.DTMDATA, "
    strSql = strSql & "AC.strAnistia, "
    strSql = strSql & "LV.DBLVALOR dblValorOrig, "
    strSql = strSql & "LV.Intmoeda, "
    strSql = strSql & "(Select STRNUMDISTRIBUIDOR " & strCONCAT & " '/' " & strCONCAT & " STRSERIEDISTRIBUIDOR From tblexecutivo Where pkid = DA.intExecutivo) strExecutivo "
    strSql = strSql & "From tblacordo AC, tbllancamentovalor LV, tbllancamentoalfa LA, tblDativa DA, "
    strSql = strSql & "(Select intLancamentoAlfa Pkid "
    strSql = strSql & "From tblLAncamentoValor LV "
    strSql = strSql & "Where Not LV.pkid in(Select intLancamentoValor From tblLancamentoReceita) "
    'strsql = strsql & " AND Not LV.pkid in(Select intLancamentoValor From tblLancamentoPagamento) "
    strSql = strSql & "Group By intLancamentoAlfa ) A "
    If dbc_intAcordo.MatchedWithList Then
        strSql = strSql & "Where LV.INTLANCAMENTOALFAACORDO = " & dbc_intAcordo.BoundText
    Else
        strSql = strSql & "Where AC.DTMDATA BETWEEN " & gstrConvDtParaSql(dtmDataInicial) & " and " & gstrConvDtParaSql(dtmDataFinal) & " "
    End If
    strSql = strSql & " and LV.INTLANCAMENTOALFAACORDO = AC.INTLANCAMENTOALFA "
    strSql = strSql & "and LA.Pkid = LV.intLancamentoAlfa "
    strSql = strSql & "and LV.INTLANCAMENTOALFAACORDO  = A.Pkid "
    strSql = strSql & "and LA.pkid *= DA.intLancamentoAlfa "
    'strsql = strsql & " and not LA.intComposicaoDaReceita = 37 "
    'strsql = strsql & " and LA.intComposicaoDaReceita = 37 "
    strSql = strSql & "Group By LA.PKID, "
    strSql = strSql & "LV.INTLANCAMENTOALFAACORDO, "
    strSql = strSql & "LA.STRINSCRICAO, "
    strSql = strSql & "LV.PKID, "
    strSql = strSql & "LV.intParcela, "
    strSql = strSql & "LA.intComposicaoDaReceita, "
    strSql = strSql & "LA.intExercicio, "
    strSql = strSql & "LV.DTMDTVENCIMENTO, "
    strSql = strSql & "AC.DTMDATA, "
    strSql = strSql & "AC.strAnistia, "
    strSql = strSql & "LV.DBLVALOR, "
    strSql = strSql & "DA.intExecutivo, "
    strSql = strSql & "LV.Intmoeda "
    strSql = strSql & "Order by LV.INTLANCAMENTOALFAACORDO, LA.PKID, LV.INTPARCELA "
    
    If gobjBanco.CriaADO(strSql, 400, adoResultado) Then
    If Not adoResultado.EOF Then
        ReDim vetReceitas(1, 0)
        
        prg_Status.Visible = True
        prg_Status.Max = adoResultado.RecordCount
        prg_Status.Value = 0
        
        lbl_Status.Visible = True
        lbl_Status.Caption = ""
        
        'Vamos gravar nas parcelas de TBLLANCAMENTOVALOR utilizadas para criar o Acordo, a referencia do Acordo
        Do While Not adoResultado.EOF
            If gstrConvVrDoSql(gstrENulo(adoResultado!dblValororig)) <= 0 Then GoTo zerado

            gobjBanco.ExecutaBeginTrans
            
            lngLancamentoAlfaAcordo = adoResultado("intLancamentoAlfaAcordo")
            
            'Vamos montar o array que armazenara as receitas do acordo
            If gobjBanco.CriaADO("SELECT intReceita, dblValor, (SELECT Sum(dblValor) FROM tblLancamentoReceita WHERE intlancamentovalor = " & adoResultado("intLancamentoValorOrigem") & ") dblTotalReceitas FROM " & gstrLancamentoReceita & " WHERE intLancamentoValor = " & adoResultado("intLancamentoValorOrigem"), 5, adoReceitas) Then
                
                If Not adoReceitas.EOF Then
                                
                    'Vamos atualizar as parcelas
                    'strsql = gstrStoredProcedure("sp_AtualizaParcela", adoResultado("intComposicaoDaReceita") & ", " & adoResultado("intExercicio") & ", " & adoResultado("intParcela") & ", " & gstrConvDtParaSql(adoResultado("dtmDtVencimento")) & ", " & gstrConvDtParaSql(adoResultado("dtmData")) & ", " & gstrConvVrParaSql(adoResultado("dblValorOrig")) & ", " & adoResultado("intMoeda"), True)
                    strSql = gstrStoredProcedure("sp_AtualizaParcela", adoResultado("intComposicaoDaReceita") & ", " & adoResultado("intExercicio") & ", " & adoResultado("intParcela") & ", " & gstrConvDtParaSql(adoResultado("dtmDtVencimento")) & ", " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & gstrConvVrParaSql(adoResultado("dblValorOrig")) & ", " & adoResultado("intMoeda"), True)
                    
                    If gobjBanco.CriaADO(strSql, 80, adoParcelas) Then
                    
                        'Caso exista executivo vamos calcular Honorarios e Custas
                        If Len(Trim(gstrENulo(adoResultado!strExecutivo))) <> 0 Then
                            dblHonorarios = FormatCurrency((adoParcelas("dblValorPrincipal").Value + adoParcelas("dblValorMulta").Value + adoParcelas("dblValorJuros").Value + adoParcelas("dblValorCorrecao").Value) * dblPorcHonorarios, 2)
                        End If

                        'Vamos aplicar a anistia necessaria
                        If Len(Trim(adoResultado("strAnistia"))) > 0 Then
                            If Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) >= 1 And Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) < 13 Then
                                strSql = "SELECT dblMulta, dblJuros FROM tbldesctosprovisorios WHERE intPArcela = 12 "
                            ElseIf Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) >= 12 And Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) < 31 Then
                                strSql = "SELECT dblMulta, dblJuros FROM tbldesctosprovisorios WHERE intPArcela = 30 "
                            ElseIf Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) >= 30 And Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) < 61 Then
                                strSql = "SELECT dblMulta, dblJuros FROM tbldesctosprovisorios WHERE intPArcela = 60 "
                            ElseIf Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) >= 60 And Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) < 121 Then
                                strSql = "SELECT dblMulta, dblJuros FROM tbldesctosprovisorios WHERE intPArcela = 120 "
                            End If

                            If gobjBanco.CriaADO(strSql, 5, adoAnistia) Then
                                If Not adoAnistia.EOF Then
                                    dblValorMulta = adoParcelas("dblValorMulta").Value - ((adoAnistia("dblmulta").Value / 100) * adoParcelas("dblValorMulta").Value)
                                    dblValorJuros = adoParcelas("dblValorJuros").Value - ((adoAnistia("dbljuros").Value / 100) * adoParcelas("dblValorJuros").Value)
                                End If
                            End If
                            adoAnistia.Close
                        Else
                            dblValorMulta = adoParcelas("dblValorMulta").Value
                            dblValorJuros = adoParcelas("dblValorJuros").Value
                        End If
                        
                        If adoReceitas.RecordCount = 1 Then
                            
                            intExisteReceita = -1
                            
                            'Vamos verificar se a receita ja existe no array
                            For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                If vetReceitas(RECEITA, intForReceitasArray) = adoReceitas("intReceita").Value Then
                                    intExisteReceita = intForReceitasArray
                                    Exit For
                                End If
                            Next
                                                    
                            'Caso ja exista a receita no array vamos somar, senao a criaremos
                            If intExisteReceita > -1 Then
                                vetReceitas(1, intExisteReceita) = vetReceitas(1, intExisteReceita) + CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                            Else
                                'Nao é preciso redimensionar, pois o array é criado com 0
                                If Val(vetReceitas(0, 0)) > 0 Then
                                    ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                End If
                                vetReceitas(0, UBound(vetReceitas, 2)) = adoReceitas("intReceita").Value
                                vetReceitas(1, UBound(vetReceitas, 2)) = CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                            End If
                            
                            dblValorTotalReceitas = CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                            
                        Else
                            
                            'Vamos atribuir ao array receita a receita da parcela
                            For intForReceitas = 0 To adoReceitas.RecordCount - 1
                                
                                intExisteReceita = -1
                                
                                'Vamos verificar se a receita ja existe no array
                                For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                    If vetReceitas(RECEITA, intForReceitasArray) = adoReceitas("intReceita").Value Then
                                        intExisteReceita = intForReceitasArray
                                        Exit For
                                    End If
                                Next
                                
                                'Vamos calcular a proporcao da receita ao valor original
                                dblProporcao = adoReceitas("dblValor").Value / adoReceitas("dblTotalReceitas") 'adoResultado("dblValorOrig")
                                
                                'Caso ja exista a receita no array vamos somar, senao a criaremos
                                If intExisteReceita > -1 Then
                                    vetReceitas(1, intExisteReceita) = vetReceitas(1, intExisteReceita) + (adoParcelas("dblValorPrincipal").Value * dblProporcao)
                                Else
                                    'Nao é preciso redimensionar, pois o array é criado com 0
                                    If Val(vetReceitas(0, 0)) > 0 Then
                                        ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                    End If
                                    vetReceitas(0, UBound(vetReceitas, 2)) = adoReceitas("intReceita").Value
                                    vetReceitas(1, UBound(vetReceitas, 2)) = adoParcelas("dblValorPrincipal").Value * dblProporcao
                                End If
                                
                                'Soma total das receitas da parcela
                                dblValorTotalReceitas = dblValorTotalReceitas + FormatCurrency(adoParcelas("dblValorPrincipal").Value * dblProporcao, 2)
                                
                                adoReceitas.MoveNext
                                
                            Next
                        
                        End If
                        
                        intExisteReceita = -1
                        
                        'Vamos criar as receitas Multa, Juros e Correcao, caso nao seja zero
                        If dblValorMulta > 0 Then
                            'Vamos verificar se a receita ja existe no array
                            For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaMulta Then
                                    intExisteReceita = intForReceitasArray
                                    Exit For
                                End If
                            Next
                            If intExisteReceita > -1 Then
                                vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + CCur(dblValorMulta)
                                intExisteReceita = -1
                            Else
                                ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaMulta
                                vetReceitas(1, UBound(vetReceitas, 2)) = dblValorMulta
                            End If
                        End If
                        
                        If dblValorJuros > 0 Then
                            'Vamos verificar se a receita ja existe no array
                            For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaJuros Then
                                    intExisteReceita = intForReceitasArray
                                    Exit For
                                End If
                            Next
                            If intExisteReceita > -1 Then
                                vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + CCur(dblValorJuros)
                                intExisteReceita = -1
                            Else
                                ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaJuros
                                vetReceitas(1, UBound(vetReceitas, 2)) = dblValorJuros
                            End If
                        End If
                        
                        If adoParcelas("dblValorCorrecao").Value > 0 Then
                            'Vamos verificar se a receita ja existe no array
                            For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaCorrecao Then
                                    intExisteReceita = intForReceitasArray
                                    Exit For
                                End If
                            Next
                            If intExisteReceita > -1 Then
                                vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + CCur(adoParcelas("dblValorCorrecao").Value)
                                intExisteReceita = -1
                            Else
                                ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaCorrecao
                                vetReceitas(1, UBound(vetReceitas, 2)) = adoParcelas("dblValorCorrecao").Value
                            End If
                        End If
                        
                        'Vamos definir a receita de Honorarios
                        'Vamos verificar se a receita ja existe no array
                        If Val(gstrConvVrParaSql(gstrConvVrDoSql(dblHonorarios, 2, , True))) > 0 Then
                            For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaHonorario Then
                                    intExisteReceita = intForReceitasArray
                                    Exit For
                                End If
                            Next
                            If intExisteReceita > -1 Then
                                vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + Val(gstrConvVrParaSql(gstrConvVrDoSql(dblHonorarios)))
                                intExisteReceita = -1
                            Else
                                ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaHonorario
                                vetReceitas(1, UBound(vetReceitas, 2)) = Val(gstrConvVrParaSql(gstrConvVrDoSql(dblHonorarios, 2, , True)))
                            End If
                        End If

                        
                        'Soma total das receitas da parcela + multa, juros e correcao
                        dblValorTotalReceitas = dblValorTotalReceitas + dblValorMulta + dblValorJuros + adoParcelas("dblValorCorrecao").Value
                        
                        dblValorTotalReceitas = 0
                    
                    End If
                    
                Else
                    adoResultado.MoveNext
                    GoTo Proximo
                    
                    ExibeMensagem "Não foi(ram) encontrada(s) receita(s) para uma das parcelas originárias do acordo. A gravação não foi concluída."
                    gobjBanco.ExecutaRollbackTrans
                    Exit Function
                End If
                
            End If
zerado:
            adoResultado.MoveNext
            
            If adoResultado.EOF Then GoTo FinalizaOperacao
            
            'Caso tenha mudado de acordo vamos atualizar as receitas
            If lngLancamentoAlfaAcordo <> adoResultado("intLancamentoAlfaAcordo") Then
                    
FinalizaOperacao:

                'Vamos obter as parcelas do acordo
                If gobjBanco.CriaADO("Select Pkid, dblValor " & _
                                     "From tbllancamentovalor LV " & _
                                     "Where LV.INTLANCAMENTOALFA = " & lngLancamentoAlfaAcordo & _
                                     "Order by INTPARCELA", 5, adoAcordo) Then
                    
                    'Vamos apagar as receitas antigas
                    gobjBanco.Execute "DELETE FROM tblLancamentoReceita WHERE intLancamentoValor in (select Pkid from tbllancamentovalor where intlancamentoalfa = " & lngLancamentoAlfaAcordo & ")"
                    
                    'Vamos passar por todas as parcelas do acordo para atualizar as receitas
                    Do While Not adoAcordo.EOF
                
                        'Vamos atualizar as receitas na tabela TBLLANCAMENTORECEITA
                        For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                            If Trim(vetReceitas(RECEITA, intForReceitasArray)) <> "" Then
                                strSql = "INSERT INTO " & gstrLancamentoReceita & "(intLancamentoValor, intReceita, dblValor, dtmDtAtualizacao, lngCodusr)  " & _
                                         " VALUES (" & adoAcordo("Pkid").Value & "," & vetReceitas(RECEITA, intForReceitasArray) & "," & gstrConvVrParaSql(gstrConvVrDoSql(vetReceitas(VALOR_RECEITA, intForReceitasArray) / adoAcordo.RecordCount, 2, , True)) & "," & strGETDATE & ",1)"

                            
                                If Not gobjBanco.Execute(strSql) Then
                                    ExibeMensagem "A atualização não foi concluída."
                                    gobjBanco.ExecutaRollbackTrans
                                    Exit Function
                                   
                                End If
                            End If
                        Next
                        
                        
                        adoAcordo.MoveNext
                        
                    Loop
                
                End If
                
                ReDim vetReceitas(1, 0)
                
            End If
Proximo:
            prg_Status.Value = prg_Status.Value + 1
            lbl_Status.Caption = adoResultado.AbsolutePosition & " de " & adoResultado.RecordCount
            DoEvents
            Me.Refresh
            gobjBanco.ExecutaCommitTrans
        Loop
        
        
        
        prg_Status.Visible = False
        lbl_Status.Visible = False
        
        DoEvents
        
        Exit Function
    Else
        Exit Function
    End If
    
    End If
    
Problema_Na_Rotina:

   ExibeDetalheErro "Erro na rotina de Gravação do Acordo."
   gobjBanco.ExecutaRollbackTrans
   
End Function

Private Function GravarAcordoProvisorio2(dtmDataInicial As Date, dtmDataFinal As Date) As Boolean

    Dim adoResultado                As New ADODB.Recordset
    Dim adoAuxiliar                 As New ADODB.Recordset
    Dim adoReceitas                 As New ADODB.Recordset
    Dim adoParcelas                 As New ADODB.Recordset
    Dim adoAnistia                  As New ADODB.Recordset
    Dim adoAcordo                   As New ADODB.Recordset
    
    Dim strSql                      As String
    
    Dim lngLancamentoAlfaAcordo     As Long
    
    Dim intFor                      As Integer
    Dim intForReceitas              As Integer
    Dim intForReceitasArray         As Integer
    
    Dim intExisteReceita            As Integer
    Dim dblProporcao                As Double
    
    Dim dblValorParcela             As Double
    Dim dblValorTotalReceitas       As Double
    
    Dim dblValorMulta               As Double
    Dim dblValorJuros               As Double
    Dim dblHonorarios               As Double
    Dim dblPorcHonorarios           As Double
    
    Dim lngReceitaMulta             As Long
    Dim lngReceitaJuros             As Long
    Dim lngReceitaCorrecao          As Long
    Dim lngReceitaHonorario         As Long
    
    Dim lngComposicaoDaReceita      As Long
    Dim strComposicaoDaReceita      As String
    Dim intUtilizacao               As Integer
    
On Error GoTo Problema_Na_Rotina
    
    Set gobjBanco = New clsBanco
        
    'Vamos obter a composicao de acordo
    If gobjBanco.CriaADO("SELECT Pkid, strDescricao, intUtilizacao FROM " & gstrComposicaoDaReceita & " WHERE intUtilizacao = " & TYP_ACORDO, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngComposicaoDaReceita = Space$(0) & adoResultado("Pkid").Value
            strComposicaoDaReceita = Space$(0) & adoResultado("strDescricao").Value
            intUtilizacao = Space$(0) & adoResultado("intUtilizacao").Value
        Else
            ExibeMensagem "Não foi encontrada nenhuma Composição do Tipo de Acordo. A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    Else
        ExibeMensagem "Não foi encontrada nenhuma Composição do Tipo de Acordo. A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    'Vamos obter as receitas de Multas, Juros e Correcao da Composicao de Receita
    If gobjBanco.CriaADO("SELECT intReceitaMulta, intReceitaJuros, intReceitaCorrecao FROM " & gstrParametroAtualizacao & " WHERE intExercicio = " & Year(gstrDataDoSistema) & " AND intComposicaoReceita = " & lngComposicaoDaReceita, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngReceitaMulta = Space$(0) & adoResultado("intReceitaMulta").Value
            lngReceitaJuros = Space$(0) & adoResultado("intReceitaJuros").Value
            lngReceitaCorrecao = Space$(0) & adoResultado("intReceitaCorrecao").Value
        Else
            ExibeMensagem "Não foi(ram) encontrada(s) receita(s) de Multa, Juros para a Composição de Receita. A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    Else
        ExibeMensagem "Não foi(ram) encontrada(s) receita(s) de Multa, Juros para a Composição de Receita. A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    'Vamos obter s receita de Honorarios
    If gobjBanco.CriaADO("SELECT intReceitaHonorarios FROM " & gstrParametrosTributario, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            lngReceitaHonorario = Space$(0) & adoResultado("intReceitaHonorarios").Value
        Else
            ExibeMensagem "Não foi encontrada receita de Honorário. A gravação não foi concluída."
            gobjBanco.ExecutaRollbackTrans
            Exit Function
        End If
    Else
        ExibeMensagem "Não foi encontrada receita de Honorário. A gravação não foi concluída."
        gobjBanco.ExecutaRollbackTrans
        Exit Function
    End If
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & gstrISNULL("dblPorcHonorarios", 0) & " dblPorcHonorarios "
    strSql = strSql & "FROM "
    strSql = strSql & " tblParametrosTributario PT "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoAuxiliar) Then
        If Not adoAuxiliar.EOF Then
            dblPorcHonorarios = adoAuxiliar("dblPorcHonorarios").Value
        Else
            dblPorcHonorarios = 0
        End If
    End If
    
    strSql = ""
    strSql = strSql & "SELECT LA.PKId, LA.strInscricao, LA.intExercicio "
    strSql = strSql & "FROM tblLancamentoAlfa LA INNER JOIN "
    strSql = strSql & "tblLancamentoValor LV ON LA.PKId = LV.intLancamentoAlfa INNER JOIN "
    strSql = strSql & "TblAcordo AC ON LA.PKId = AC.INTLANCAMENTOALFA LEFT OUTER JOIN "
    strSql = strSql & "tblLancamentoReceita LR ON LV.PKId = LR.intlancamentoValor "
    If dbc_intAcordo.MatchedWithList Then
        strSql = strSql & "Where LA.PKID = " & dbc_intAcordo.BoundText
    Else
        strSql = strSql & "Where AC.DTMDATA BETWEEN " & gstrConvDtParaSql(dtmDataInicial) & " and " & gstrConvDtParaSql(dtmDataFinal) & " "
    End If
    strSql = strSql & " and (LA.intUtilizacao = 4) And (lR.Pkid Is Null) "
    strSql = strSql & "GROUP BY LA.PKId, LA.strInscricao, LA.intExercicio "

    If gobjBanco.CriaADO(strSql, 400, adoAuxiliar) Then
    If Not adoAuxiliar.EOF Then
        
        ReDim vetReceitas(1, 0)
        
        prg_Status.Visible = True
        prg_Status.Max = adoAuxiliar.RecordCount
        prg_Status.Value = 0
        
        lbl_Status.Visible = True
        lbl_Status.Caption = ""
        
        DoEvents
        Me.Refresh
        
        'Vamos gravar nas parcelas de TBLLANCAMENTOVALOR utilizadas para criar o Acordo, a referencia do Acordo
        Do While Not adoAuxiliar.EOF
        
            'Vamos buscar o acordo mais atual desta divida
            strSql = "SELECT d2.NumeroAcordo, d2.Exercicio "
            strSql = strSql & "FROM admgrjorigem.dbo.tbl_dividasta_novo d, admgrjorigem.dbo.tbl_dividasta_novo d2 "
            strSql = strSql & "WHERE " & gstrCONVERT(CDT_INT, "D.NumeroAcordo") & " = '" & Mid(adoAuxiliar("strInscricao"), 1, Len(adoAuxiliar("strInscricao")) - 4) & "' and "
            strSql = strSql & "D.exercicio = '" & adoAuxiliar("intExercicio") & "' and "
            strSql = strSql & " convert(bigint,d2.inscricaoda) = convert(bigint,d.inscricaoda) and convert(bigint,d2.anobase) = convert(bigint,d.anobase)"
            strSql = strSql & "order by d2.exercicio desc, d2.numeroacordo desc"
            
            If Not gobjBanco.CriaADO(strSql, 5, adoAnistia) Then Exit Function
                        
            If adoAnistia.RecordCount = 0 Then
                adoAuxiliar.MoveNext
                GoTo Proximo
            End If
            
            'Vamos buscar as origens
            strSql = "SELECT LV.dblValor dblValorOrig, "
            strSql = strSql & "LV.intLancamentoAlfaAcordo, LV.Pkid intLancamentoValorOrigem, "
            strSql = strSql & "LV.intParcela, LV.dtmDtVencimento, LV.intMoeda, "
            strSql = strSql & "LA.intComposicaoDaReceita, LA.intExercicio, "
            strSql = strSql & "AC.strAnistia, "
            strSql = strSql & "(Select Count(*) from tbllancamentoValor where intLancamentoAlfa = " & adoAuxiliar("Pkid") & ") dblTotParcelaAcordo, "
            strSql = strSql & "(Select STRNUMDISTRIBUIDOR " & strCONCAT & " '/' " & strCONCAT & " STRSERIEDISTRIBUIDOR From tblexecutivo Where pkid = DA.intExecutivo) strExecutivo "
            strSql = strSql & "FROM admgrjOrigem.dbo.tbl_dividasta_novo D, "
            'strsql = strsql & "(SELECT inscricaoDA, anobase, numeroacordo, exercicio FROM admgrjOrigem.dbo.tbl_dividasta_novo X where RTRIM(LTRIM(X.numeroacordo)) + RTRIM(LTRIM(X.exercicio)) <> REPLICATE('0', 10 - Len('" & Val(adoAuxiliar("strInscricao")) & "')) +  '" & Val(adoAuxiliar("strInscricao")) & "' and " & gstrCONVERT(CDT_INT, "Exercicio") & " >= " & adoAuxiliar("intExercicio") & ") D2, "
            strSql = strSql & "(SELECT inscricaoDA, anobase, numeroacordo, exercicio FROM admgrjOrigem.dbo.tbl_dividasta_novo X where RTRIM(LTRIM(X.numeroacordo)) = '" & adoAnistia("NumeroAcordo") & "' and RTRIM(LTRIM(X.exercicio)) = '" & adoAnistia("Exercicio") & "') D2, "
            strSql = strSql & "tbllancamentoalfa LA, "
            strSql = strSql & "tbllancamentovalor LV, "
            strSql = strSql & "tblAcordo AC, "
            strSql = strSql & "tblDativa DA "
            strSql = strSql & " WHERE " & gstrCONVERT(CDT_INT, "D.NumeroAcordo") & " = '" & Mid(adoAuxiliar("strInscricao"), 1, Len(adoAuxiliar("strInscricao")) - 4) & "' and "
            strSql = strSql & "D.exercicio = '" & adoAuxiliar("intExercicio") & "' and "
            strSql = strSql & "D2.inscricaoDA = D.inscricaoDA and "
            strSql = strSql & "D2.anobase = D.AnoBase and "
            strSql = strSql & "LA.strinscricao = REPLICATE('0', 20 - Len(D2.numeroacordo + D2.Exercicio)) +  convert(varchar,D2.NUMEROACORDO) + convert(varchar,D2.EXERCICIO) and "
            'strsql = strsql & "LV.intlancamentoalfaacordo = La.Pkid "
            strSql = strSql & "LV.intlancamentoalfa = La.Pkid "
            'strsql = strsql & " and LV.INTLANCAMENTOALFAACORDO = AC.INTLANCAMENTOALFA "
            strSql = strSql & " and LV.INTLANCAMENTOALFA = AC.INTLANCAMENTOALFA "
            strSql = strSql & " and LA.pkid *= DA.intLancamentoAlfa "
            strSql = strSql & " GROUP BY D2.numeroacordo, D2.exercicio, LV.dblValor, LV.intLancamentoAlfaAcordo, LV.Pkid , LV.intParcela, "
            strSql = strSql & " LV.Dtmdtvencimento , LV.intMoeda, LA.intComposicaoDaReceita, LA.intExercicio, AC.strAnistia, DA.intExecutivo "
            
            adoAnistia.Close: Set adoAnistia = Nothing
            
            If gobjBanco.CriaADO(strSql, 400, adoResultado) Then
            If Not adoResultado.EOF Then
                
                Do While Not adoResultado.EOF
                
                    If gstrConvVrDoSql(gstrENulo(adoResultado!dblValororig)) <= 0 Then GoTo zerado
        
                    gobjBanco.ExecutaBeginTrans
                    
                    lngLancamentoAlfaAcordo = adoAuxiliar("Pkid")
                    
                    'Vamos montar o array que armazenara as receitas do acordo
                    If gobjBanco.CriaADO("SELECT intReceita, dblValor, (SELECT Sum(dblValor) FROM tblLancamentoReceita WHERE intlancamentovalor = " & adoResultado("intLancamentoValorOrigem") & ") dblTotalReceitas FROM " & gstrLancamentoReceita & " WHERE intLancamentoValor = " & adoResultado("intLancamentoValorOrigem"), 5, adoReceitas) Then
                        
                        If Not adoReceitas.EOF Then
                                        
                            'Vamos atualizar as parcelas
                            'strsql = gstrStoredProcedure("sp_AtualizaParcela", adoResultado("intComposicaoDaReceita") & ", " & adoResultado("intExercicio") & ", " & adoResultado("intParcela") & ", " & gstrConvDtParaSql(adoResultado("dtmDtVencimento")) & ", " & gstrConvDtParaSql(adoResultado("dtmData")) & ", " & gstrConvVrParaSql(adoResultado("dblValorOrig")) & ", " & adoResultado("intMoeda"), True)
                            strSql = gstrStoredProcedure("sp_AtualizaParcela", adoResultado("intComposicaoDaReceita") & ", " & adoResultado("intExercicio") & ", " & adoResultado("intParcela") & ", " & gstrConvDtParaSql(adoResultado("dtmDtVencimento")) & ", " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & gstrConvVrParaSql(adoResultado("dblValorOrig")) & ", " & adoResultado("intMoeda"), True)
                            
                            If gobjBanco.CriaADO(strSql, 80, adoParcelas) Then
                            
                                'Caso exista executivo vamos calcular Honorarios e Custas
                                If Len(Trim(gstrENulo(adoResultado!strExecutivo))) <> 0 Then
                                    dblHonorarios = FormatCurrency((adoParcelas("dblValorPrincipal").Value + adoParcelas("dblValorMulta").Value + adoParcelas("dblValorJuros").Value + adoParcelas("dblValorCorrecao").Value) * dblPorcHonorarios, 2)
                                End If
        
                                'Vamos aplicar a anistia necessaria
                                If Len(Trim(adoResultado("strAnistia"))) > 0 Then
                                    If Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) >= 1 And Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) < 13 Then
                                        strSql = "SELECT dblMulta, dblJuros FROM tbldesctosprovisorios WHERE intPArcela = 12 "
                                    ElseIf Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) >= 12 And Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) < 31 Then
                                        strSql = "SELECT dblMulta, dblJuros FROM tbldesctosprovisorios WHERE intPArcela = 30 "
                                    ElseIf Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) >= 30 And Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) < 61 Then
                                        strSql = "SELECT dblMulta, dblJuros FROM tbldesctosprovisorios WHERE intPArcela = 60 "
                                    ElseIf Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) >= 60 And Val(gstrENulo(adoResultado!dblTotParcelaAcordo)) < 121 Then
                                        strSql = "SELECT dblMulta, dblJuros FROM tbldesctosprovisorios WHERE intPArcela = 120 "
                                    End If
        
                                    If gobjBanco.CriaADO(strSql, 5, adoAnistia) Then
                                        If Not adoAnistia.EOF Then
                                            dblValorMulta = adoParcelas("dblValorMulta").Value - ((adoAnistia("dblmulta").Value / 100) * adoParcelas("dblValorMulta").Value)
                                            dblValorJuros = adoParcelas("dblValorJuros").Value - ((adoAnistia("dbljuros").Value / 100) * adoParcelas("dblValorJuros").Value)
                                        End If
                                    End If
                                    adoAnistia.Close
                                Else
                                    dblValorMulta = adoParcelas("dblValorMulta").Value
                                    dblValorJuros = adoParcelas("dblValorJuros").Value
                                End If
                                
                                If adoReceitas.RecordCount = 1 Then
                                    
                                    intExisteReceita = -1
                                    
                                    'Vamos verificar se a receita ja existe no array
                                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                        If vetReceitas(RECEITA, intForReceitasArray) = adoReceitas("intReceita").Value Then
                                            intExisteReceita = intForReceitasArray
                                            Exit For
                                        End If
                                    Next
                                                            
                                    'Caso ja exista a receita no array vamos somar, senao a criaremos
                                    If intExisteReceita > -1 Then
                                        vetReceitas(1, intExisteReceita) = vetReceitas(1, intExisteReceita) + CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                                    Else
                                        'Nao é preciso redimensionar, pois o array é criado com 0
                                        If Val(vetReceitas(0, 0)) > 0 Then
                                            ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                        End If
                                        vetReceitas(0, UBound(vetReceitas, 2)) = adoReceitas("intReceita").Value
                                        vetReceitas(1, UBound(vetReceitas, 2)) = CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                                    End If
                                    
                                    dblValorTotalReceitas = CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                                    
                                Else
                                    
                                    'Vamos atribuir ao array receita a receita da parcela
                                    For intForReceitas = 0 To adoReceitas.RecordCount - 1
                                        
                                        intExisteReceita = -1
                                        
                                        'Vamos verificar se a receita ja existe no array
                                        For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                            If vetReceitas(RECEITA, intForReceitasArray) = adoReceitas("intReceita").Value Then
                                                intExisteReceita = intForReceitasArray
                                                Exit For
                                            End If
                                        Next
                                        
                                        'Vamos calcular a proporcao da receita ao valor original
                                        dblProporcao = adoReceitas("dblValor").Value / adoReceitas("dblTotalReceitas") 'adoResultado("dblValorOrig")
                                        
                                        'Caso ja exista a receita no array vamos somar, senao a criaremos
                                        If intExisteReceita > -1 Then
                                            vetReceitas(1, intExisteReceita) = vetReceitas(1, intExisteReceita) + (adoParcelas("dblValorPrincipal").Value * dblProporcao)
                                        Else
                                            'Nao é preciso redimensionar, pois o array é criado com 0
                                            If Val(vetReceitas(0, 0)) > 0 Then
                                                ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                            End If
                                            vetReceitas(0, UBound(vetReceitas, 2)) = adoReceitas("intReceita").Value
                                            vetReceitas(1, UBound(vetReceitas, 2)) = adoParcelas("dblValorPrincipal").Value * dblProporcao
                                        End If
                                        
                                        'Soma total das receitas da parcela
                                        dblValorTotalReceitas = dblValorTotalReceitas + FormatCurrency(adoParcelas("dblValorPrincipal").Value * dblProporcao, 2)
                                        
                                        adoReceitas.MoveNext
                                        
                                    Next
                                
                                End If
                                
                                intExisteReceita = -1
                                
                                'Vamos criar as receitas Multa, Juros e Correcao, caso nao seja zero
                                If dblValorMulta > 0 Then
                                    'Vamos verificar se a receita ja existe no array
                                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                        If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaMulta Then
                                            intExisteReceita = intForReceitasArray
                                            Exit For
                                        End If
                                    Next
                                    If intExisteReceita > -1 Then
                                        vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + CCur(dblValorMulta)
                                        intExisteReceita = -1
                                    Else
                                        ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                        vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaMulta
                                        vetReceitas(1, UBound(vetReceitas, 2)) = dblValorMulta
                                    End If
                                End If
                                
                                If dblValorJuros > 0 Then
                                    'Vamos verificar se a receita ja existe no array
                                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                        If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaJuros Then
                                            intExisteReceita = intForReceitasArray
                                            Exit For
                                        End If
                                    Next
                                    If intExisteReceita > -1 Then
                                        vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + CCur(dblValorJuros)
                                        intExisteReceita = -1
                                    Else
                                        ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                        vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaJuros
                                        vetReceitas(1, UBound(vetReceitas, 2)) = dblValorJuros
                                    End If
                                End If
                                
                                If adoParcelas("dblValorCorrecao").Value > 0 Then
                                    'Vamos verificar se a receita ja existe no array
                                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                        If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaCorrecao Then
                                            intExisteReceita = intForReceitasArray
                                            Exit For
                                        End If
                                    Next
                                    If intExisteReceita > -1 Then
                                        vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + CCur(adoParcelas("dblValorCorrecao").Value)
                                        intExisteReceita = -1
                                    Else
                                        ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                        vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaCorrecao
                                        vetReceitas(1, UBound(vetReceitas, 2)) = adoParcelas("dblValorCorrecao").Value
                                    End If
                                End If
                                
                                'Vamos definir a receita de Honorarios
                                'Vamos verificar se a receita ja existe no array
                                If Val(gstrConvVrParaSql(gstrConvVrDoSql(dblHonorarios, 2, , True))) > 0 Then
                                    For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                        If vetReceitas(RECEITA, intForReceitasArray) = lngReceitaHonorario Then
                                            intExisteReceita = intForReceitasArray
                                            Exit For
                                        End If
                                    Next
                                    If intExisteReceita > -1 Then
                                        vetReceitas(1, intExisteReceita) = CCur(vetReceitas(1, intExisteReceita)) + Val(gstrConvVrParaSql(gstrConvVrDoSql(dblHonorarios)))
                                        intExisteReceita = -1
                                    Else
                                        ReDim Preserve vetReceitas(1, UBound(vetReceitas, 2) + 1)
                                        vetReceitas(0, UBound(vetReceitas, 2)) = lngReceitaHonorario
                                        vetReceitas(1, UBound(vetReceitas, 2)) = Val(gstrConvVrParaSql(gstrConvVrDoSql(dblHonorarios, 2, , True)))
                                    End If
                                End If
        
                                
                                'Soma total das receitas da parcela + multa, juros e correcao
                                dblValorTotalReceitas = dblValorTotalReceitas + dblValorMulta + dblValorJuros + adoParcelas("dblValorCorrecao").Value
                                
                                dblValorTotalReceitas = 0
                            
                            End If
                            
                        Else
                            adoAuxiliar.MoveNext
                            GoTo Proximo
                            
                            ExibeMensagem "Não foi(ram) encontrada(s) receita(s) para uma das parcelas originárias do acordo. A gravação não foi concluída."
                            gobjBanco.ExecutaRollbackTrans
                            Exit Function
                        End If
                        
                    End If
zerado:
                    adoResultado.MoveNext
                    
                    If adoResultado.EOF Then
        
                        'Vamos obter as parcelas do acordo
                        If gobjBanco.CriaADO("Select Pkid, dblValor " & _
                                             "From tbllancamentovalor LV " & _
                                             "Where LV.INTLANCAMENTOALFA = " & lngLancamentoAlfaAcordo & _
                                             "Order by INTPARCELA", 5, adoAcordo) Then
                            
                            'Vamos apagar as receitas antigas
                            gobjBanco.Execute "DELETE FROM tblLancamentoReceita WHERE intLancamentoValor in (select Pkid from tbllancamentovalor where intlancamentoalfa = " & lngLancamentoAlfaAcordo & ")"
                            
                            'Vamos passar por todas as parcelas do acordo para atualizar as receitas
                            Do While Not adoAcordo.EOF
                        
                                'Vamos atualizar as receitas na tabela TBLLANCAMENTORECEITA
                                For intForReceitasArray = 0 To UBound(vetReceitas, 2)
                                    If Trim(vetReceitas(RECEITA, intForReceitasArray)) <> "" Then
                                        strSql = "INSERT INTO " & gstrLancamentoReceita & "(intLancamentoValor, intReceita, dblValor, dtmDtAtualizacao, lngCodusr)  " & _
                                                 " VALUES (" & adoAcordo("Pkid").Value & "," & vetReceitas(RECEITA, intForReceitasArray) & "," & gstrConvVrParaSql(gstrConvVrDoSql(vetReceitas(VALOR_RECEITA, intForReceitasArray) / adoAcordo.RecordCount, 2, , True)) & "," & strGETDATE & ",1)"
        
                                    
                                        If Not gobjBanco.Execute(strSql) Then
                                            ExibeMensagem "A atualização não foi concluída."
                                            gobjBanco.ExecutaRollbackTrans
                                            Exit Function
                                           
                                        End If
                                    End If
                                Next
                                
                                
                                adoAcordo.MoveNext
                                
                            Loop
                        
                        End If
                        
                        ReDim vetReceitas(1, 0)
                        
                    End If
                    
                Loop
                
            End If
            End If
            
            adoAuxiliar.MoveNext
            
Proximo:
            
            prg_Status.Value = prg_Status.Value + 1
            lbl_Status.Caption = adoAuxiliar.AbsolutePosition & " de " & adoAuxiliar.RecordCount
            DoEvents
            Me.Refresh
            gobjBanco.ExecutaCommitTrans
            
        Loop
        
        prg_Status.Visible = False
        lbl_Status.Visible = False
        
        DoEvents
        
        Exit Function
        
    Else
        Exit Function
    End If
    
    End If
    
Problema_Na_Rotina:

   ExibeDetalheErro "Erro na rotina de Gravação do Acordo."
   gobjBanco.ExecutaRollbackTrans
   
End Function

Private Sub AplicarAcrescimoPorParcela()
Dim adoResultado As New ADODB.Recordset
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "PT.dblAcrescimoPorParcela "
    strSql = strSql & "FROM "
    strSql = strSql & " tblParametrosTributario PT "

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
        
            'Vamos verificar o valor do acrescimo
            If Not IsNull(adoResultado("dblAcrescimoPorParcela").Value) And adoResultado("dblAcrescimoPorParcela").Value > 0 Then
                'Valor Total * acrescimo por parcela - acrescimo
                txtdblValor = gstrConvVrDoSql((CCur(dblValorAcordoOriginal) + CCur(gstrConvVrDoSql(txtdblAcrescimos, , , True))) * (1 + ((adoResultado("dblAcrescimoPorParcela").Value) * (txt_QtdeParcelas - 1))) - CCur(gstrConvVrDoSql(txtdblAcrescimos, , , True)), 2)
            End If
            
        End If
    End If
    
End Sub

Private Sub AplicarAcrescimoDesctoProvisorio()
Dim dblPorcJuros             As Double
Dim adoResultado             As New ADODB.Recordset
Dim strSql                   As String
    
    dblAcrescimoDescProv = 0
    intAcrescimoParcIniDescProv = 0
    
    'Vamos verificar se existe desconto provisorio selecionado
    If dbc_intDescProvisorios.MatchedWithList Then
        
        strSql = ""
        strSql = strSql & "SELECT DP.dblJurosParcelamento, DP.intParcelaInicialJuros "
        strSql = strSql & "FROM "
        strSql = strSql & gstrDescontosProvisorios & " DP "
        strSql = strSql & "WHERE Pkid = " & dbc_intDescProvisorios.BoundText
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                
                If Not IsNull(adoResultado("dblJurosParcelamento").Value) And Not IsNull(adoResultado("intParcelaInicialJuros").Value) Then
                    'Fator de juros a ser aplicado a partir da parcela inicial
                    dblPorcJuros = ((DateDiff("M", DateAdd("M", adoResultado("intParcelaInicialJuros").Value - 1, txt_dtmVencimento), DateAdd("M", txt_QtdeParcelas - 1, txt_dtmVencimento)) + 2) * (adoResultado("dblJurosParcelamento").Value / 100)) / 2
                    'Qtde de parcelas a ser aplicado o acrescimo
                    intQtdeParcelasAcrescimo = txt_QtdeParcelas - (adoResultado("intParcelaInicialJuros").Value - 1)
                    'Total de acrescimo do desconto provisorio
                    dblAcrescimoDescProv = FormatCurrency(((dblValorAcordoOriginal / txt_QtdeParcelas) * dblPorcJuros) * intQtdeParcelasAcrescimo, 2)
                    
                    intAcrescimoParcIniDescProv = adoResultado("intParcelaInicialJuros").Value
                    
                End If
                
                txtdblValor = gstrConvVrDoSql(txtdblValor + dblAcrescimoDescProv, 2, , True)
                
            End If
        End If

    End If
    
End Sub
