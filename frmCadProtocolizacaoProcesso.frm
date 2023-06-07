VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadProtocolizacaoProcesso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastramento de Processos"
   ClientHeight    =   9285
   ClientLeft      =   2730
   ClientTop       =   1380
   ClientWidth     =   9645
   HelpContextID   =   4
   Icon            =   "frmCadProtocolizacaoProcesso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKId 
      Height          =   285
      Left            =   7200
      TabIndex        =   54
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtdtmDtData 
      Height          =   285
      Left            =   6540
      TabIndex        =   53
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   9075
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   16007
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cadastramento"
      TabPicture(0)   =   "frmCadProtocolizacaoProcesso.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_TotalVolume"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbldtmdthora"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrsumula"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintcodassunto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintcodcentrocusto"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbldtmdtdata"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblstrDescricao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_intCepC"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_intUFC"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_intLogradouroC"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl_intBairroC"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl_intMunicipioC"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl_intTipoProcesso"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dbcintTipoProcesso"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkbitEmpenho"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkbitRevisaoCalculo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "fra_EnderecoAcao"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "dbcintCodAssunto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "tdb_Protocolo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_dtmDtHora"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_dtmDtdata"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtstrCodigo"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtstrSumula"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmd_Assunto"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txt_strLogradouroC"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txt_intBairroC"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txt_intMunicipioC"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txt_intUFC"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txt_intCepC"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txt_intCodCentroCusto"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtintExercicio"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtbitDigito"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "opt_req(1)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "opt_req(0)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "fra_Requerente"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtintCentroCusto"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cbo_Volume"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmd_TipoProcesso"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      Begin VB.CommandButton cmd_TipoProcesso 
         Height          =   300
         Left            =   4260
         Picture         =   "frmCadProtocolizacaoProcesso.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Clique para cadastar requerente."
         Top             =   405
         Width           =   360
      End
      Begin VB.ComboBox cbo_Volume 
         Height          =   315
         Left            =   5250
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox txtintCentroCusto 
         Height          =   285
         Left            =   1995
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   4440
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Frame fra_Requerente 
         Height          =   675
         Left            =   180
         TabIndex        =   13
         Top             =   735
         Width           =   9135
         Begin VB.OptionButton opt_Requerente 
            Caption         =   "Requerente"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   150
            Width           =   1875
         End
         Begin VB.OptionButton opt_Requerente 
            Caption         =   "Unidade Centro de Custo"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   390
            Width           =   2205
         End
         Begin VB.TextBox txt_CodigoContribuinte 
            Height          =   315
            Left            =   2325
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   210
            Width           =   1035
         End
         Begin VB.CommandButton cmd_Requerente 
            Height          =   300
            Left            =   8640
            Picture         =   "frmCadProtocolizacaoProcesso.frx":117C
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Clique para cadastar requerente."
            Top             =   225
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintCodCentroCusto 
            Height          =   315
            Left            =   3420
            TabIndex        =   17
            Top             =   210
            Width           =   5220
            _ExtentX        =   9208
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintCodContribuinte 
            Height          =   315
            Left            =   3420
            TabIndex        =   18
            Top             =   210
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
      End
      Begin VB.OptionButton opt_req 
         Caption         =   "End. Residência"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   5805
         TabIndex        =   22
         Top             =   1530
         Width           =   1485
      End
      Begin VB.OptionButton opt_req 
         Caption         =   "End. Correspondência"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   7395
         TabIndex        =   23
         Top             =   1530
         Width           =   1875
      End
      Begin VB.TextBox txtbitDigito 
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
         Left            =   2265
         MaxLength       =   2
         TabIndex        =   4
         Top             =   420
         Width           =   285
      End
      Begin VB.TextBox txtintExercicio 
         CausesValidation=   0   'False
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
         Left            =   1785
         MaxLength       =   4
         TabIndex        =   3
         Top             =   420
         Width           =   465
      End
      Begin VB.TextBox txt_intCodCentroCusto 
         Height          =   285
         Left            =   2010
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   4440
         Width           =   7290
      End
      Begin VB.TextBox txt_intCepC 
         Height          =   285
         Left            =   7800
         TabIndex        =   29
         Top             =   1860
         Width           =   1485
      End
      Begin VB.TextBox txt_intUFC 
         Height          =   285
         Left            =   6390
         TabIndex        =   27
         Top             =   1860
         Width           =   465
      End
      Begin VB.TextBox txt_intMunicipioC 
         Height          =   285
         Left            =   1380
         TabIndex        =   31
         Top             =   2220
         Width           =   7905
      End
      Begin VB.TextBox txt_intBairroC 
         Height          =   285
         Left            =   1380
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1860
         Width           =   4275
      End
      Begin VB.TextBox txt_strLogradouroC 
         Height          =   285
         Left            =   1380
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1500
         Width           =   4275
      End
      Begin VB.CommandButton cmd_Assunto 
         Height          =   300
         Left            =   8940
         Picture         =   "frmCadProtocolizacaoProcesso.frx":129A
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Clique para cadastar assunto."
         Top             =   4050
         Width           =   360
      End
      Begin VB.TextBox txtstrSumula 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   615
         Left            =   2010
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   4800
         Width           =   7275
      End
      Begin VB.TextBox txtstrCodigo 
         CausesValidation=   0   'False
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
         Left            =   945
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   825
      End
      Begin VB.TextBox txt_dtmDtdata 
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
         Left            =   6765
         MaxLength       =   19
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   420
         Width           =   1035
      End
      Begin VB.TextBox txt_dtmDtHora 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   8355
         MaxLength       =   19
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   420
         Width           =   945
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Protocolo 
         Height          =   3450
         Left            =   180
         TabIndex        =   50
         Top             =   5535
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   6085
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "pkid"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Processo"
         Columns(1).DataField=   "strCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data "
         Columns(2).DataField=   "dtmDtData"
         Columns(2).NumberFormat=   "Short Date"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Requerente"
         Columns(3).DataField=   "strContribuinte"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).DataField=   "strCentroCusto"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Unidade Centro de Custo"
         Columns(5).DataField=   "intCodCentroCusto"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Assunto"
         Columns(6).DataField=   "strCodAssunto"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).DataField=   "intCodContribuinte"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).DataField=   "intCodAssunto"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2355"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2275"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2302"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2223"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(29)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=4683"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=4604"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=5477"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=5398"
         Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(41)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(45)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(46)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(48)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(49)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(51)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(52)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(53)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=70,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
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
      Begin MSDataListLib.DataCombo dbcintCodAssunto 
         Height          =   315
         Left            =   2010
         TabIndex        =   45
         Top             =   4050
         Width           =   6930
         _ExtentX        =   12224
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin Threed.SSFrame fra_EnderecoAcao 
         Height          =   1395
         Left            =   180
         TabIndex        =   32
         Top             =   2550
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   2461
         _StockProps     =   14
         Caption         =   " Endereço de Ação "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton cmd_Logradouros 
            Height          =   300
            Left            =   8655
            Picture         =   "frmCadProtocolizacaoProcesso.frx":13B8
            Style           =   1  'Graphical
            TabIndex        =   58
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Clique para cadastar requerente."
            Top             =   270
            Width           =   360
         End
         Begin VB.TextBox txt_intBairroA 
            CausesValidation=   0   'False
            Height          =   285
            Left            =   4200
            TabIndex        =   57
            Top             =   675
            Width           =   4830
         End
         Begin VB.TextBox txtintCepA 
            Height          =   285
            Left            =   1185
            TabIndex        =   41
            Top             =   990
            Width           =   1905
         End
         Begin VB.TextBox txtstrReferenciaA 
            Height          =   285
            Left            =   4200
            MaxLength       =   100
            TabIndex        =   43
            Top             =   990
            Width           =   4830
         End
         Begin VB.TextBox txtintNumeroA 
            Height          =   285
            Left            =   1185
            MaxLength       =   6
            TabIndex        =   36
            Top             =   660
            Width           =   645
         End
         Begin VB.TextBox txtstrComplementoA 
            Height          =   285
            Left            =   2415
            MaxLength       =   16
            TabIndex        =   38
            Top             =   660
            Width           =   1155
         End
         Begin MSDataListLib.DataCombo dbcintLogradouroA 
            Height          =   315
            Left            =   1185
            TabIndex        =   34
            Top             =   270
            Width           =   7440
            _ExtentX        =   13123
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   3720
            TabIndex        =   39
            Top             =   735
            Width           =   405
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   315
            TabIndex        =   33
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   825
            TabIndex        =   40
            Top             =   1065
            Width           =   285
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Referência"
            Height          =   195
            Left            =   3345
            TabIndex        =   42
            Top             =   1050
            Width           =   780
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   570
            TabIndex        =   35
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   1935
            TabIndex        =   37
            Top             =   720
            Width           =   480
         End
      End
      Begin VB.CheckBox chkbitRevisaoCalculo 
         Caption         =   "Revisão de Cálculo"
         Height          =   195
         Left            =   3450
         TabIndex        =   51
         Top             =   4935
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CheckBox chkbitEmpenho 
         Caption         =   "Empenho"
         Height          =   195
         Left            =   3450
         TabIndex        =   52
         Top             =   5205
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dbcintTipoProcesso 
         Height          =   315
         Left            =   3030
         TabIndex        =   5
         Top             =   420
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbl_intTipoProcesso 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   2655
         TabIndex        =   59
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Volume"
         Height          =   195
         Left            =   4680
         TabIndex        =   11
         Top             =   510
         Width           =   525
      End
      Begin VB.Label lbl_intMunicipioC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Município"
         Height          =   195
         Left            =   645
         TabIndex        =   30
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label lbl_intBairroC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   930
         TabIndex        =   24
         Top             =   1935
         Width           =   405
      End
      Begin VB.Label lbl_intLogradouroC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Logradouro"
         Height          =   195
         Left            =   540
         TabIndex        =   20
         Top             =   1545
         Width           =   810
      End
      Begin VB.Label lbl_intUFC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "UF"
         Height          =   195
         Left            =   6120
         TabIndex        =   26
         Top             =   1935
         Width           =   210
      End
      Begin VB.Label lbl_intCepC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cep"
         Height          =   195
         Left            =   7395
         TabIndex        =   28
         Top             =   1935
         Width           =   285
      End
      Begin VB.Label lblstrDescricao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Processo"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   510
         Width           =   660
      End
      Begin VB.Label lbldtmdtdata 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   6345
         TabIndex        =   9
         Top             =   510
         Width           =   435
      End
      Begin VB.Label lblintcodcentrocusto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Unidade Centro de Custo"
         Height          =   195
         Left            =   165
         TabIndex        =   46
         Top             =   4470
         Width           =   1785
      End
      Begin VB.Label lblintcodassunto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assunto"
         Height          =   195
         Left            =   1380
         TabIndex        =   44
         Top             =   4140
         Width           =   570
      End
      Begin VB.Label lblstrsumula 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Súmula"
         Height          =   195
         Left            =   1410
         TabIndex        =   48
         Top             =   4830
         Width           =   525
      End
      Begin VB.Label lbldtmdthora 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   195
         Left            =   7905
         TabIndex        =   10
         Top             =   510
         Width           =   345
      End
      Begin VB.Label lbl_TotalVolume 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5850
         TabIndex        =   12
         Top             =   420
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmCadProtocolizacaoProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando              As Boolean
Dim mobjAux                    As Object
Dim mblnSelecionou             As Boolean
Dim mblnPrimeiraVez            As Boolean
Dim FlagOperacao               As String
Dim adoResultado               As ADODB.Recordset
Dim strCodigoAtual             As String
Dim strCodigo                  As String
Dim strSQLCCDominio            As String  'Armazena os centros de custo referentes ao dominio do usuario
Dim vetProtocolizacaoVolumes() As Volumes
Dim bytOrdenacao               As Byte
Dim blnOrdenacaoAsc            As Boolean
Dim blnContinua                As Boolean
Private Sub cbo_Volume_Click()
    If cbo_Volume.ListIndex = -1 Then
        cbo_Volume.ListIndex = 0
    Else
        txt_dtmDtdata.Text = Format(gstrENulo(vetProtocolizacaoVolumes(cbo_Volume.ListIndex).DTMDATA), "dd/mm/yyyy")
        txt_dtmDtHora.Text = Format(gstrENulo(vetProtocolizacaoVolumes(cbo_Volume.ListIndex).DTMDATA), "hh:mm")
        txtstrSumula.Text = gstrENulo(vetProtocolizacaoVolumes(cbo_Volume.ListIndex).strSumula)
    End If
    HabilitaDesabilitaBotao1 IIf(ProcessoArquivado(cbo_Volume.ItemData(cbo_Volume.ListIndex)), False, True), gstrBtnArquivo, gstrCriarVolume
End Sub

Private Sub cmd_Assunto_Click()
    CarregaForm frmCadCatalogoAssunto, dbcintCodAssunto
End Sub

Private Sub cmd_Logradouros_Click()
    CarregaForm frmCadLogradouro, dbcintLogradouroA
End Sub

Private Sub cmd_Requerente_Click()
    If opt_Requerente(0).Value Then
        frmCadContribuinte.Caption = "Requerentes"
        frmCadContribuinte.Tag = "Requerentes"
        CarregaForm frmCadContribuinte, dbcintCodContribuinte
    Else
        CarregaForm frmCadLocais, dbcintCodCentroCusto
    End If
End Sub

Private Sub cmd_TipoProcesso_Click()
    CarregaForm frmCadTipoProcesso, dbcintTipoProcesso, strQueryTipoProcesso
End Sub

Private Sub dbcintCodCentroCusto_Change()
    If dbcintCodCentroCusto.MatchedWithList Then
        'txt_CodigoContribuinte = Mid(dbcintCodCentroCusto.Text, InStr(1, dbcintCodCentroCusto.Text, "(") + 1, Len(dbcintCodCentroCusto.Text) - InStr(1, dbcintCodCentroCusto.Text, "(") - 1)
        txt_CodigoContribuinte.Tag = dbcintCodCentroCusto.BoundText
    End If
End Sub

Private Sub dbcintCodCentroCusto_Click(Area As Integer)
   DropDownDataCombo dbcintCodCentroCusto, Me, Area
End Sub

Private Sub dbcintCodCentroCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintCodCentroCusto, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCodCentroCusto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintCodCentroCusto
End Sub

Private Sub dbcintCodAssunto_Change()
    If dbcintCodAssunto.MatchedWithList Then
        CarregaDadosAssunto
    End If
End Sub

Private Sub dbcintCodAssunto_Click(Area As Integer)
    If Area = 0 Then
       DropDownDataCombo dbcintCodAssunto, Me, Area
    ElseIf Area = 2 Then
       CarregaDadosAssunto
    End If
End Sub

Private Sub dbcintCodAssunto_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintCodAssunto, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCodContribuinte_Change()
    If dbcintCodContribuinte.MatchedWithList Then
        txt_CodigoContribuinte = dbcintCodContribuinte.BoundText
        CarregaDadosContribuinte
    End If
End Sub

Private Sub dbcintCodContribuinte_Click(Area As Integer)
    If Area = 0 Then
        dbcintCodContribuinte.DataChanged = False
        DropDownDataCombo dbcintCodContribuinte, Me, Area
    ElseIf Area = 2 And dbcintCodContribuinte.MatchedWithList Then
        CarregaDadosContribuinte
    End If
End Sub

Private Sub dbcintCodContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintCodContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouroA_Click(Area As Integer)
    If Area = 0 Then
        DropDownDataCombo dbcintLogradouroA, Me, Area
    End If
    
    If dbcintLogradouroA.MatchedWithList Then
        LogradouroCep dbcintLogradouroA.BoundText, txt_intBairroA, False, , , txtintCepA
    End If

End Sub

Private Sub dbcintLogradouroA_GotFocus()
Dim adoRec   As ADODB.Recordset
 
    If mblnAlterando = False And (dbcintLogradouroA.Text = Space$(0) And txt_strLogradouroC.Text <> Space$(0)) Then
        If opt_req(1).Value = True Then Exit Sub
        If MsgBox("O endereço de ação é o mesmo que o endereço residencial?", vbYesNo + vbQuestion) = vbYes Then
            If dbcintLogradouroA.MatchedWithList = False Then PreencherListaDeOpcoes dbcintLogradouroA
            
            If dbcintCodContribuinte.MatchedWithList Then
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strQueryEnderecoContrib, 10, adoRec) Then
                    dbcintLogradouroA.BoundText = gstrENulo(adoRec("intLogradouro"))
                    txtintNumeroA = gstrENulo(adoRec("intNumero"))
                    txtstrComplementoA = gstrENulo(adoRec("strComplemento"))
                    txtintCepA = gstrCEPFormatado(adoRec("intcep"))
                End If
            End If
        End If
    End If
    
End Sub

Private Sub dbcintLogradouroA_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLogradouroA, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouroA_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 452
    VirificaGradeListView Me
    
    If mblnAlterando Then HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCriarVolume
    
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    Dim strSql As String
    
    VerificaObjParaAplicar mobjAux
    
    TrocaCorObjeto txt_CodigoContribuinte, True
    TrocaCorObjeto txt_strLogradouroC, True
    TrocaCorObjeto txt_intBairroC, True
    TrocaCorObjeto txt_intUFC, True
    TrocaCorObjeto txt_intCepC, True
    TrocaCorObjeto txt_intMunicipioC, True
    TrocaCorObjeto txt_intCodCentroCusto, True
    TrocaCorObjeto txt_intBairroA, True
    TrocaCorObjeto txtbitDigito, True
    
    opt_Requerente(0).Value = True
    
    'LeDaTabelaParaObj gstrContribuinte, dbcintCodContribuinte, "SELECT PKId, strNome FROM " & gstrContribuinte & " ORDER BY strNome"
    dbcintCodContribuinte.Tag = gstrQueryDataComboContribuinte & ";strNome"
    dbcintCodCentroCusto.Tag = strQueryLocais & ";A.strDescricao"
    dbcintTipoProcesso.Tag = strQueryTipoProcesso & ";intCodigo"
    
    'LeDaTabelaParaObj gstrCatalogoAssunto, dbcintCodAssunto, strQueryAssunto
    dbcintCodAssunto.Tag = gstrQueryDataComboAssunto & ";strDescricao"
    
    dbcintLogradouroA.Tag = strQueryDataComboLogradouroA & ";L.strDescricao"
    
    
    strSql = "SELECT PKid, strDescricao FROM " & gstrBairro & " A"
    txt_intBairroA.Tag = strSql & ";A.strDescricao"
    
    
    'LeDaTabelaParaObj gstrProtocolizacaoProcesso, tdb_Protocolo, strQueryProtocolo
    
    mblnPrimeiraVez = False
    
    'LimpaCampos
    txtintExercicio = Year(gstrDataDoSistema)

End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrCriarVolume
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCriarVolume
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrCriarVolume
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCriarVolume
    mblnSelecionou = False
    mblnPrimeiraVez = False
    mblnAlterando = False
End Sub

Private Sub opt_Requerente_Click(Index As Integer)
    If Index = 0 Then
        dbcintCodCentroCusto.BoundText = Space$(0)
        dbcintCodCentroCusto.Visible = False
        dbcintCodContribuinte.Visible = True
    Else
        dbcintCodContribuinte.BoundText = Space$(0)
        dbcintCodCentroCusto.Visible = True
        dbcintCodContribuinte.Visible = False
    End If
End Sub

Private Sub tdb_Protocolo_Click()
    mblnPrimeiraVez = True
End Sub

Sub tdb_Protocolo_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Protocolo_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Protocolo
End Sub

Private Sub tdb_Protocolo_HeadClick(ByVal ColIndex As Integer)

blnOrdenacaoAsc = IIf(bytOrdenacao = ColIndex, Not blnOrdenacaoAsc, True)
bytOrdenacao = ColIndex: MantemForm gstrRefresh

End Sub

Private Sub tdb_Protocolo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Protocolo
End Sub

Private Function TemProcesso() As Boolean
Dim strSql As String
Dim ADOTemp As ADODB.Recordset
        strSql = "SELECT COUNT(*) AS TemProcesso FROM " & gstrTramiteProtocolo
        strSql = strSql & " WHERE intProtocolizacaoVolume = " & cbo_Volume.ItemData(cbo_Volume.ListIndex)
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, ADOTemp) Then
            If ADOTemp!TemProcesso > 0 Then
                TemProcesso = True
            End If
        End If
End Function

Private Sub tdb_Protocolo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Protocolo
        
        If Not .EOF And Not .BOF And mblnPrimeiraVez Then

            txtPKId.Text = .Columns("PKID").Value
                    
            'Cláudio
            txt_CodigoContribuinte.Text = ""
                    
            'Vamos preencher as combos somente com o registro selecionado
            If .Columns("intCodContribuinte").Value = Space$(0) Then
                opt_Requerente(1).Value = True
                dbcintCodCentroCusto.Text = gstrENulo(.Columns("strCentroCusto").Value)
                dbcintCodCentroCusto.SetFocus: dbcintCodCentroCusto_Click 0
            Else
                opt_Requerente(0).Value = True
                dbcintCodContribuinte.Text = gstrENulo(.Columns("strContribuinte").Value)
                'dbcintCodContribuinte.SetFocus: dbcintCodContribuinte_Click 0
                'Cláudio
                'PreencherListaDeOpcoes dbcintCodContribuinte
                PreencherListaDeOpcoes dbcintCodContribuinte, tdb_Protocolo.Columns("intCodContribuinte")
            End If
            
            dbcintCodAssunto.Text = gstrENulo(.Columns("strCodAssunto").Value)
            'dbcintCodAssunto.SetFocus: dbcintCodAssunto_Click 0
            PreencherListaDeOpcoes dbcintCodAssunto
            LeDaTabelaParaObj gstrProtocolizacaoProcesso, Me
            
            PreencheBairroA
            
            gCorLinhaSelecionada tdb_Protocolo
            
            txtintCepA = gstrCEPFormatado(txtintCepA)
            
            'If dbcintLogradouroA.MatchedWithList Then
            '    LogradouroCep dbcintLogradouroA.BoundText, txt_intBairroA, False, , , txtintCepA
            'Else
            '    txt_intBairroA.Text = ""
            'End If
            
            'Vamos preencher a combo de volumes
            PreencherVolumes
            
            If TemProcesso Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar, gstrSalvar
            End If
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            mblnSelecionou = True
            mblnAlterando = True
                 
            'txt_dtmDtdata.Text = Format(txtdtmDtData, "dd/mm/yy")
            txt_dtmDtHora.Text = Format(txtdtmDtData, "Short Time")
                
            'If dbcintCodContribuinte.BoundText <> Space$(0) Then dbcintCodContribuinte_Click 2
            If .Columns("intCodContribuinte") <> Space$(0) Then
                dbcintCodContribuinte.BoundText = gstrENulo(.Columns("intCodContribuinte").Value)
                'Cláudio
                If dbcintCodContribuinte.MatchedWithList Then 'Cadastrado nesta aplicação
                    dbcintCodContribuinte_Click 2
                Else
                    LeDaTabelaParaObj "", dbcintCodContribuinte, strQueryContribuinteEspecifico(.Columns("intCodContribuinte").Value)
                    dbcintCodContribuinte.BoundText = gstrENulo(.Columns("intCodContribuinte").Value)
                End If
            Else
                txt_strLogradouroC = Space$(0)
                txt_intBairroC = Space$(0)
                txt_intUFC = Space$(0)
                txt_intCepC = Space$(0)
                txt_intMunicipioC = Space$(0)
            End If
            
            'If dbcintCodAssunto.BoundText <> Space$(0) Then dbcintCodAssunto_Click 2
            dbcintCodAssunto.BoundText = gstrENulo(.Columns("intCodAssunto").Value)
            dbcintCodAssunto_Click 2
            
            strCodigoAtual = txtstrCodigo.Text
    
            If tdb_Protocolo.Enabled Then tdb_Protocolo.SetFocus
            
            txtstrCodigo.SelLength = 0
    
            TrocaCorObjeto txtstrCodigo, mblnAlterando
            TrocaCorObjeto txtintExercicio, mblnAlterando
            
            
            HabilitaDesabilitaBotao1 IIf(ProcessoArquivado(cbo_Volume.ItemData(cbo_Volume.ListIndex)), False, True), gstrBtnArquivo, gstrCriarVolume
     
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSql       As String
Dim strSqlVolume As String
Dim bytVolume    As Byte
Dim blnAlteracao As Boolean

'******************************************************************************************
' Data: 24/06/2003
' Alteração: - Retirada a tabela tblUnidadeCentroDeCusto, pois será usada tblLocais
' Responsável: Gustavo Monteiro
'******************************************************************************************
    
    blnAlteracao = mblnAlterando
    
    If UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = gstrImprimir Then
        If Len(tdb_Protocolo.Columns("PKId")) > 0 Then
            ImprimeRelatorio rptRelEtiquetasProcessoMod2, strqueryrelatorio, "Etiquetas de Processo"
        End If
        Exit Sub
    End If
    
    If strModoOperacao = gstrCriarVolume Then

        If MsgBox("Você deseja criar volume para este processo?", vbYesNo + vbQuestion) = vbYes Then
            strSqlVolume = "INSERT INTO " & gstrProtocolizacaoVolume & " (intProtocolizacaoProcesso, intVolume, strSumula, dtmDtData, dtmDtAtualizacao, lngCodUsr) VALUES (" & _
                               txtPKId & ", " & glngPegaUltimaChave(gstrProtocolizacaoVolume, "intVolume", "intProtocolizacaoProcesso", txtPKId.Text) + 1 & ", '" & txtstrSumula.Text & "', " & strGETDATE & ", " & strGETDATE & ", " & glngCodUsr & ")"

            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSqlVolume
            Set gobjBanco = Nothing
            
        
            'Chamamos agora o formulário de relatório de etiquetas
            
            frmRelEtiquetadeProcesso.fra_FaixaData.Enabled = False
            
            frmRelEtiquetadeProcesso.txtintProtocolizacaoInicial = txtstrCodigo
            frmRelEtiquetadeProcesso.txtintProtocolizacaoFinal = txtstrCodigo
            frmRelEtiquetadeProcesso.txtintExercicio = txtintExercicio
            frmRelEtiquetadeProcesso.fra_FaixaNum.Enabled = False
            frmRelEtiquetadeProcesso.txtVolume = glngPegaUltimaChave(gstrProtocolizacaoVolume, "intVolume", "intProtocolizacaoProcesso", txtPKId.Text)
            frmRelEtiquetadeProcesso.chkTodosVolumes.Value = 0
            frmRelEtiquetadeProcesso.chkTodosVolumes.Enabled = False
            frmRelEtiquetadeProcesso.fraVolume.Enabled = False
            
            CriaPrimeiroTramite
            
            CarregaForm frmRelEtiquetadeProcesso
            
        End If
        PreencherVolumes
        cbo_Volume.ListIndex = cbo_Volume.ListCount - 1
        
        Exit Sub
        
    End If
    
    strSql = strQuery

    If UCase(strModoOperacao) = UCase(gstrSalvar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
        mblnPrimeiraVez = False
    End If
        
    'Vamos forcar a data do servidor
    If strModoOperacao = gstrSalvar And Not mblnAlterando = True Then
        txt_dtmDtdata.Text = gstrDataDoSistema
        txtdtmDtData = gstrDataDoSistema(True, False, False)
    End If
    
    If strModoOperacao = gstrLocalizar Then
        'txtdtmDtData = Trim(txt_dtmDtdata) & " " & Trim(txt_dtmDtHora)
        txtdtmDtData = Space$(0)
        txtintCepA = Space$(0)
    End If
    
    If strModoOperacao = gstrSalvar And blnAlteracao = True Then
        If CInt(cbo_Volume.Text) = 1 Then
            txtdtmDtData = gstrDataFormatada(txt_dtmDtdata) & " " & gstrDataFormatada(gstrDataDoSistema(True), , False, True)
        End If
    End If
    
    If ToolBarGeral(strModoOperacao, gstrProtocolizacaoProcesso, mblnAlterando, _
                    tdb_Protocolo, Me, mobjAux, strQueryProtocolo(True), , _
                    rptRelProtoclizacaoProcessos, strqueryrelatorio, False) Then
        
        If strModoOperacao = gstrSalvar And blnAlteracao = True Then
            'Vamos gravar os dados referentes a alteracao do volume
            strSqlVolume = "UPDATE " & gstrProtocolizacaoVolume & " SET strSumula = '" & txtstrSumula.Text & "', dtmDtData = " & gstrConvDtParaSql(txt_dtmDtdata) & " WHERE PkID = " & cbo_Volume.ItemData(cbo_Volume.ListIndex)
                
            Set gobjBanco = New clsBanco
            gobjBanco.Execute strSqlVolume
            Set gobjBanco = Nothing
            
            'Vamos atribuir a variavel o volume alterado para mante-lo em exibicao
            bytVolume = cbo_Volume.ListIndex
        End If
    
    End If
    
    If UCase(strModoOperacao) = gstrNovo Then LimpaCampos
    
    If Not blnContinua Then Exit Sub
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) And Not blnAlteracao Then
        CriaPrimeiroTramite
    End If
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) And blnAlteracao = False And gblnCancelarInclusao = False Then
        frmRelEtiquetadeProcesso.fra_FaixaData.Enabled = False
        frmRelEtiquetadeProcesso.txtintProtocolizacaoInicial = txtstrCodigo
        frmRelEtiquetadeProcesso.txtintProtocolizacaoFinal = txtstrCodigo
        frmRelEtiquetadeProcesso.txtintExercicio = txtintExercicio
        frmRelEtiquetadeProcesso.fra_FaixaNum.Enabled = False
        frmRelEtiquetadeProcesso.chkTodosVolumes.Value = 1
        frmRelEtiquetadeProcesso.chkTodosVolumes.Enabled = False
        CarregaForm frmRelEtiquetadeProcesso
    End If
    
    LimpaObjeto Me
    
    DoEvents
    
    If UCase(strModoOperacao) = gstrSalvar Then
        
        CalculaDigito txtstrCodigo.Text & txtintExercicio ' recalcula o dígito
        tdb_Protocolo_Click
        tdb_Protocolo_RowColChange 1, 1
        
        If blnAlteracao = False Then
            If Not gblnCancelarInclusao Then frmRelEtiquetadeProcesso.SetFocus
            Exit Sub
        Else
            cbo_Volume.ListIndex = bytVolume
            Exit Sub
        End If
    End If
    
    If UCase(strModoOperacao) = gstrDeletar Then
       LimpaCampos
       Exit Sub
    End If
    
End Sub

Private Function strQueryContribuinte() As String
Dim strSql As String
Dim adoRec As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO("SELECT blnResidenteNoMunicipio FROM " & gstrContribuinte & " WHERE PkID = " & dbcintCodContribuinte.BoundText, 10, adoRec) Then
        
        With adoRec
            
            If Not (.BOF And .EOF) Then
               
                If gstrENulo(!blnResidenteNoMunicipio) = False Then
 
                    strSql = ""
                '    strSql = strSql & "SELECT RTRIM(LTRIM(ISNULL(TLC.strSigla, '') + ' ' + ISNULL(UC.strDescricao,'') + ' ' + A.strLogradouroC)) AS LogradouroCorrespondencia "
                    strSql = strSql & "SELECT RTRIM(LTRIM(" & gstrISNULL("TLC.strSigla", "''") & strCONCAT & " ' ' " & strCONCAT & gstrISNULL("UC.strDescricao", "''") & strCONCAT & " ' ' " & strCONCAT & " A.strLogradouroC)) AS LogradouroCorrespondencia "
                    strSql = strSql & ", A.intNumeroC "
                    strSql = strSql & ", A.strComplementoC "
                    strSql = strSql & ", A.intCEPC "
                    strSql = strSql & ", A.strCNPJCPF "
                    strSql = strSql & ", A.blnResidenteNoMunicipio "
                    strSql = strSql & ", A.strBairroC "
                    strSql = strSql & ", UFC.strSigla AS UFC "
                    strSql = strSql & ", MC.strDescricao AS MunicipioC "
                    
                    '................... Tabelas
                    strSql = strSql & " FROM "
                    strSql = strSql & gstrLogradouro & " L, "
                    strSql = strSql & gstrTituloLogradouro & " UC, "
                    strSql = strSql & gstrTipoLogradouro & " TLC, "
                    strSql = strSql & gstrCidade & " MC, "
                    strSql = strSql & gstrUF & " UFC, "
                    strSql = strSql & gstrContribuinte & " A"
                    
                    '................... Condição
                    strSql = strSql & " WHERE "
                    '    strSql = strSql & " L.intTituloLogradouro *= UC.PKId "
                    strSql = strSql & " A.intTituloLogradouro " & strOUTJSQLServer & "= UC.PKId " & strOUTJOracle
                    '    strSql = strSql & " AND L.intTipoLogradouro *= TLC.PKId "
                    strSql = strSql & " AND A.intTipoLogradouro " & strOUTJSQLServer & "= TLC.PKId " & strOUTJOracle
                    strSql = strSql & " AND A.intUFC " & strOUTJSQLServer & "= UFC.PKId" & strOUTJOracle
                    strSql = strSql & " AND A.intMunicipioC " & strOUTJSQLServer & "= MC.PKID" & strOUTJOracle
                    strSql = strSql & " AND A.PKId = " & dbcintCodContribuinte.BoundText
                    
                Else

                    strSql = ""
                    '    strSql = strSql & "SELECT L.PKId, RTRIM(LTRIM(ISNULL(TL.strSigla, '') + ' ' + ISNULL(U.strDescricao,'') + ' ' + L.strDescricao)) AS LogradouroResidencia "
                    strSql = strSql & "SELECT L.PKId AS intLogradouro, RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & strCONCAT & " ' ' " & strCONCAT & gstrISNULL("U.strDescricao", "''") & strCONCAT & " ' ' " & strCONCAT & " L.strDescricao)) AS LogradouroResidencia "
                    strSql = strSql & ", A.blnResidenteNoMunicipio "
                    strSql = strSql & ", A.intNumero, A.strCNPJCPF "
                    strSql = strSql & ", A.strComplemento "
                    'strSQL = strSQL & ", B.strDescricao "
                    'strSQL = strSQL & ", A.intCEP "
                    'strSQL = strSQL & ", UF.strSigla "
                    'strSQL = strSQL & ", M.strDescricao as strMunicipio "
    
                    '................... Tabelas
                    strSql = strSql & " FROM "
                    strSql = strSql & gstrLogradouro & " L, "
                    strSql = strSql & gstrContribuinte & " A, "
                    strSql = strSql & gstrTipoLogradouro & " TL,"
                    strSql = strSql & gstrTituloLogradouro & " U "
                    'strSQL = strSQL & gstrBairro & " B,"
                    'strSQL = strSQL & gstrUF & " UF, "
                    'strSQL = strSQL & gstrCidade & " M "
    
                    '................... Condição
                    strSql = strSql & " WHERE "
                    strSql = strSql & " A.intLogradouro = L.PKId "
                    '    strSql = strSql & " AND L.intTituloLogradouro *= U.PKId "
                    strSql = strSql & " AND L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle
                    '    strSql = strSql & " AND L.intTipoLogradouro *= TL.PKId "
                    strSql = strSql & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
                    'strSQL = strSQL & " AND A.intBairro " & strOUTJSQLServer & "= B.PKId" & strOUTJOracle
                    'strSQL = strSQL & " AND A.intUF " & strOUTJSQLServer & "= UF.PKId" & strOUTJOracle
                    'strSQL = strSQL & " AND A.intMunicipio " & strOUTJSQLServer & "= M.PKID" & strOUTJOracle
                    strSql = strSql & " AND A.PKId = " & dbcintCodContribuinte.BoundText

                End If
                
                '................... Ordem
                strSql = strSql & " ORDER BY A.PKID"
            
            End If
            
        End With
            
    End If
    
    strQueryContribuinte = strSql
    
End Function

Private Function strQueryLocaiss() As String

'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 06/03/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função.
'            - Adaptação dos outer joins.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 07/03/2003
' Alteração: - Retirada a palavra chave "AS" das cláusulas FROM, pois o Oracle não permite
'            a utilização desta palavra chave nesta cláusula.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 26/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql  As String
    strSql = ""
'''Query materiais relatório demanda reprimida centro custos
'        strSql = strSql & " SELECT A.PKID , CONVERT(VARCHAR,C.PKID) + '.' + "
        strSql = strSql & " SELECT A.PKID , " & gstrCONVERT(CDT_VARCHAR, "c.PKID") & strCONCAT & "'.'" & strCONCAT
'        strSql = strSql & " CONVERT(VARCHAR,D.PKID) + '.' + "
        strSql = strSql & gstrCONVERT(CDT_VARCHAR, "D.PKID") & strCONCAT & "'.'" & strCONCAT
'        strSql = strSql & " CASE ISNULL(E.PKId,0) WHEN 0 THEN '' ELSE CONVERT(VARCHAR,E.PKID) + '.' END + "
'        strSQL = strSQL & gstrCASEWHEN(strISNULL & "(E.PKId,0)", "0,''", gstrCONVERT(CDT_VARCHAR, "E.PKID") & strCONCAT & "'.'") & strCONCAT
        strSql = strSql & gstrISNULL("E.PKId", "''", gstrCONVERT(CDT_VARCHAR, "E.PKID") & strCONCAT & "'.'") & strCONCAT
'        strSql = strSql & " CONVERT(VARCHAR,B.PKID) + '.' + CONVERT(VARCHAR,A.intCodigo) + ' ' + A.strDescricao AS Codigo, A.* "
        strSql = strSql & gstrCONVERT(CDT_VARCHAR, "B.PKID") & strCONCAT & "'.'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "A.intCodigo") & strCONCAT & "' '" & strCONCAT & "A.strDescricao AS Codigo, A.* "
        strSql = strSql & " FROM "
'        strSql = strSql & gstrUnidadeCentroDeCusto2 & " AS A,"
        strSql = strSql & gstrUnidadeCentroDeCusto2 & " A,"
'        strSql = strSql & gstrUnidadeCentroDeCusto1 & " AS B,"
        strSql = strSql & gstrUnidadeCentroDeCusto1 & " B,"
'        strSql = strSql & gstrOrgao & " AS C,"
        strSql = strSql & gstrOrgao & " C,"
'        strSql = strSql & gstrUnidadeOrcamentaria & " AS D,"
        strSql = strSql & gstrUnidadeOrcamentaria & " D,"
'        strSql = strSql & gstrSubUnidade & " AS E"
        strSql = strSql & gstrSubUnidade & " E"
        
        strSql = strSql & " WHERE "
        'Orgao
        'X Unidade Orcamentaria
        strSql = strSql & " C.PKID = D.intOrgao"
        'X SubUnidade
'        strSql = strSql & " AND C.PKID *= E.intOrgao"
        strSql = strSql & " AND C.PKID " & strOUTJSQLServer & strOUTJOracle & "= E.intOrgao"
        'X Centro de Custo 1
        strSql = strSql & " AND C.PKID = B.intOrgao"
        'X Centro de Custo 2
        strSql = strSql & " AND C.PKID = A.intOrgao"
        
        'Unidade Orcamentaria
        'X SubUnidde
'        strSql = strSql & " AND D.PKID *= E.intUnidadeOrcamentaria"
        strSql = strSql & " AND D.PKID " & strOUTJSQLServer & strOUTJOracle & "= E.intUnidadeOrcamentaria"
        'X Centro de Custo1
        strSql = strSql & " AND D.PKID = B.intUnidadeOrcamentaria"
        'X Centro de Custo2
        strSql = strSql & " AND D.PKID = A.intUnidadeOrcamentaria"
        
        'SubUnidade
        'X Centro de Custo1
'        strSql = strSql & " AND E.PKID =* B.intSubUnidade"
        strSql = strSql & " AND E.PKID =" & strOUTJSQLServer & "B.intSubUnidade" & strOUTJOracle
        'X Centro de Custo2
'        strSql = strSql & " AND E.PKID =* A.intSubUnidade"
        strSql = strSql & " AND E.PKID =" & strOUTJSQLServer & "A.intSubUnidade" & strOUTJOracle
        
        'Centro de Custo1 X Centro de Custo2
        strSql = strSql & " AND B.PKID = A.intUnidadeCentrodeCusto1 "
        
        strSql = strSql & " ORDER BY C.PKId "
    
    strQueryLocaiss = strSql
End Function

Private Function CarregaDadosAssunto()
'
'******************************************************************************************
' Data: 24/06/2003
' Alteração: - Retirada a tabela tblUnidadeCentroDeCusto, pois será usada tblLocais
' Responsável: Gustavo Monteiro
'******************************************************************************************

Dim adoRec As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO("SELECT PkID ,strDescricao FROM " & gstrLocais & " WHERE PkID = (SELECT intUnidadeCentroDeCusto FROM " & gstrCatalogoAssunto & " WHERE PkID = " & dbcintCodAssunto.BoundText & ")", 10, adoRec) Then
    
        With adoRec
            If Not (.BOF And .EOF) Then
                'Preenche dados do logradouro
                txt_intCodCentroCusto = gstrENulo(!strDescricao)
                txtintCentroCusto = gstrENulo(!Pkid)
            Else
                txt_intCodCentroCusto = Space$(0)
                txtintCentroCusto = Space$(0)
            End If
        End With
        
    End If
    
    Set gobjBanco = Nothing
    
End Function

Private Function CarregaDadosContribuinte()
Dim adoRec As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strQueryContribuinte, 10, adoRec) Then

        With adoRec
            If Not (.BOF And .EOF) Then
                'Preenche dados do logradouro
                If !blnResidenteNoMunicipio Then
                    opt_req(0) = True
                    LogradouroCep !intLogradouro, txt_intBairroC, , txt_intMunicipioC, txt_intUFC, txt_intCepC
                    txt_strLogradouroC = gstrEnderecoConcatenado(!LogradouroResidencia, , !INTNUMERO, !STRCOMPLEMENTO)
                    'txt_intBairroC = gstrENulo(!strDescricao)
                    'txt_intMunicipioC = gstrENulo(!strMunicipio)
                    'txt_intUFC = gstrENulo(!strSigla)
                    'txt_intCepC = gstrCEPFormatado(gstrENulo(!intCep))
                Else
                    'CepLogradouro !intCepC, txt_strLogradouroC, txt_intBairroC, txt_intMunicipioC, txt_intUFC, , , , , , , , , , True, True, "BA.strDescricao = '" & gstrENulo(!strBairroC) & "'"
                    opt_req(1) = True
                    txt_strLogradouroC = gstrEnderecoConcatenado(!LogradouroCorrespondencia, , !intNumeroC, !strComplementoC)
                    txt_intBairroC = gstrENulo(!strBairroC)
                    txt_intMunicipioC = gstrENulo(!MunicipioC)
                    txt_intUFC = gstrENulo(!UFC)
                    txt_intCepC = gstrCEPFormatado(gstrENulo(!intCepC))
                End If
            Else
                txt_strLogradouroC = Space$(0)
                txt_intBairroC = Space$(0)
                txt_intUFC = Space$(0)
                txt_intCepC = Space$(0)
                txt_intMunicipioC = Space$(0)
            End If
        End With
        
    End If
    
    Set gobjBanco = Nothing
    
End Function

Private Function gstrQueryDataComboAssunto()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrCatalogoAssunto & " "
    strSql = strSql & " WHERE dtmDtCancelamento IS NULL"
    strSql = strSql & " ORDER BY strDescricao"
    gstrQueryDataComboAssunto = strSql
End Function

Private Function gstrQueryDataComboContribuinte()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNome "
    strSql = strSql & "FROM " & gstrContribuinte & " "
    strSql = strSql & "ORDER BY strNome"
    gstrQueryDataComboContribuinte = strSql
End Function

Private Function strQueryAssunto() As String
Dim strSql  As String
    strSql = ""
''' query ouvidoria relatório cartas resposta
    strSql = strSql & " SELECT CA.PKID, CA.strDescricao AS CatalogoAssunto "
    strSql = strSql & " FROM "
    strSql = strSql & gstrCatalogoAssunto & " CA "
    strSql = strSql & " ORDER BY CatalogoAssunto "
    strQueryAssunto = strSql
End Function

Private Function strQuery() As String
'
'******************************************************************************************
' Data: 24/06/2003
' Alteração: - Retirada a tabela tblUnidadeCentroDeCusto, pois será usada tblLocais
' Responsável: Gustavo Monteiro
'******************************************************************************************
Dim strSql  As String
    
    strSql = ""
    strSql = strSql & " SELECT PP.PKID, PP.strCodigo, PP.dtmDtData, "
    strSql = strSql & " CO.strNome, CC.strDescricao AS intCodCentroCusto, "
    strSql = strSql & " CA.strDescricao AS intCodAssunto "
    strSql = strSql & " FROM "
    strSql = strSql & gstrProtocolizacaoProcesso & " PP, "
    strSql = strSql & gstrContribuinte & " CO, "
'    strSql = strSql & gstrUnidadeCentroDeCusto2 & " CC, "
    strSql = strSql & gstrLocais & " CC, "
    strSql = strSql & gstrCatalogoAssunto & " CA "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = PP.intCodContribuinte "
    strSql = strSql & " AND CA.PKId = PP.intCodAssunto "
    strSql = strSql & " AND CC.PKId = CA.intUnidadeCentroDeCusto "
    strSql = strSql & " AND PP.intCodContribuinte = " & gstrItemData(dbcintCodContribuinte)
    strQuery = strSql
End Function

Private Function strQueryProtocolo(Optional blnFiltrar As Boolean) As String
'
'******************************************************************************************
' Data: 24/06/2003
' Alteração: - Retirada a tabela tblUnidadeCentroDeCusto, pois será usada tblLocais
' Responsável: Gustavo Monteiro
'******************************************************************************************
Dim strSql  As String
Dim adoRec  As ADODB.Recordset
    
    'strSQLCCDominio = Space$(0)
    '
    ''Vamos obter os Locais referentes ao Dominio
    'Set gobjBanco = New clsBanco
    'If gobjBanco.CriaADO("SELECT intUnidadeCentroDeCusto FROM " & gstrDominios & " WHERE intUsuario = " & glngCodUsr & " AND intModulo = (SELECT PkID FROM " & gstrItens & " WHERE strCodItem = 'H')", 10, adoRec) Then
    '    With adoRec
    '        Do While Not .EOF
    '            If strSQLCCDominio = Space$(0) Then
    '                strSQLCCDominio = gstrENulo(!intUnidadeCentroDeCusto)
    '            Else
    '                strSQLCCDominio = strSQLCCDominio & ", " & gstrENulo(!intUnidadeCentroDeCusto)
    '            End If
    '            .MoveNext
    '        Loop
    '    End With
    'End If
    'Set gobjBanco = Nothing

    strSql = ""
    strSql = strSql & " SELECT " & gstrProtocolizacaoProcesso & ".PKID, " & gstrCONVERT(CDT_INT, gstrProtocolizacaoProcesso & ".strCodigo") & " NumCodigo, " & gstrProtocolizacaoProcesso & ".strCodigo" & strCONCAT & "'-'" & strCONCAT & "LTrim(" & gstrCONVERT(CDT_VARCHAR, gstrProtocolizacaoProcesso & ".bitDigito)") & strCONCAT & "'/'" & strCONCAT & "LTrim(" & gstrCONVERT(CDT_VARCHAR, gstrProtocolizacaoProcesso & ".intExercicio)") & " As strCodigo, " & gstrProtocolizacaoProcesso & ".dtmDtData, " & gstrProtocolizacaoProcesso & ".bitEmpenho, "
    strSql = strSql & " " & gstrProtocolizacaoProcesso & ".bitRevisaoCalculo, " & gstrProtocolizacaoProcesso & ".strSumula, " & gstrProtocolizacaoProcesso & ".intCodContribuinte, " & gstrProtocolizacaoProcesso & ".intCodAssunto, "
    strSql = strSql & gstrISNULL("CO.strNome", "CC2.strDescricao", "CO.strNome") & " AS strContribuinte, CC.strDescricao AS intCodCentroCusto, CC2.strDescricao AS strCentroCusto, "
    strSql = strSql & " CA.strDescricao AS strCodAssunto "
    strSql = strSql & " FROM "
    strSql = strSql & gstrProtocolizacaoProcesso & ", "
    strSql = strSql & gstrContribuinte & " CO, "
    'strSql = strSql & gstrUnidadeCentroDeCusto2 & " CC, "
    'strSql = strSql & gstrUnidadeCentroDeCusto2 & " CC2, "
    strSql = strSql & gstrLocais & " CC, "
    strSql = strSql & gstrLocais & " CC2, "
    strSql = strSql & gstrCatalogoAssunto & " CA "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & gstrProtocolizacaoProcesso & ".intCodContribuinte "
    strSql = strSql & " AND CC2.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & gstrProtocolizacaoProcesso & ".intCodCentroCusto "
    strSql = strSql & " AND CA.PKId = " & gstrProtocolizacaoProcesso & ".intCodAssunto "
    strSql = strSql & " AND CC.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & " CA.intUnidadeCentroDeCusto "
    
    'If strSQLCCDominio <> Space$(0) Then
    '    strSQL = strSQL & " AND (" & gstrProtocolizacaoProcesso & ".intCentroCusto IN (" & strSQLCCDominio & ") OR " & gstrProtocolizacaoProcesso & ".intCentroCusto IS NULL)"
    'Else
    '    'Caso não esteja relacionado a nenhum dominio, não vamos retornar registros
    '    strSQL = strSQL & " AND (" & gstrProtocolizacaoProcesso & ".intCentroCusto IS NULL)"
    'End If
    
    If blnFiltrar Then
        If txtstrCodigo.Text <> "" Then strSql = strSql & " AND " & gstrProtocolizacaoProcesso & ".strCodigo = " & txtstrCodigo.Text
        If txtintExercicio.Text <> "" Then strSql = strSql & " AND " & gstrProtocolizacaoProcesso & ".intExercicio = " & txtintExercicio.Text
    End If
    
    Select Case bytOrdenacao
      Case Is = 1
            strSql = strSql & " ORDER BY strCodigo" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 2
         strSql = strSql & " ORDER BY dtmDtData" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 3
         strSql = strSql & " ORDER BY strContribuinte" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 4
         strSql = strSql & " ORDER BY intCodCentrocusto" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 5
         strSql = strSql & " ORDER BY strCodAssunto" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Else
         strSql = strSql & " ORDER BY " & gstrProtocolizacaoProcesso & ".intExercicio, NumCodigo DESC"
    End Select
    
    strQueryProtocolo = strSql
    
End Function

Private Function strQueryMantem() As String
'
'******************************************************************************************
' Data: 24/06/2003
' Alteração: - Retirada a tabela tblUnidadeCentroDeCusto, pois será usada tblLocais
' Responsável: Gustavo Monteiro
'******************************************************************************************
    Dim strSql  As String
    Dim DataHora As String
    DataHora = txt_dtmDtdata + " " + txt_dtmDtHora
    strSql = ""
    strSql = strSql & " SELECT " & gstrProtocolizacaoProcesso & ".PKID, " & gstrProtocolizacaoProcesso & ".strCodigo, dtmDtData, "
    strSql = strSql & " CO.strNome, CC.strDescricao AS intCodCentroCusto, "
    strSql = strSql & " CA.strDescricao AS intCodAssunto "
    strSql = strSql & " FROM "
    strSql = strSql & gstrProtocolizacaoProcesso & ", "
    strSql = strSql & gstrContribuinte & " CO, "
'    strSql = strSql & gstrUnidadeCentroDeCusto2 & " CC, "
    strSql = strSql & gstrLocais & " CC, "
    strSql = strSql & gstrCatalogoAssunto & " CA "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = " & gstrProtocolizacaoProcesso & ".intCodContribuinte "
    strSql = strSql & " AND CA.PKId = " & gstrProtocolizacaoProcesso & ".intCodAssunto "
    strSql = strSql & " AND CC.PKId = CA.intUnidadeCentroDeCusto "
    strQueryMantem = strSql
End Function

Private Function strqueryrelatorio() As String

Dim strSql  As String
strSql = ""
strSql = gstrStoredProcedure("sp_EtiquetaProcessoMod2", _
             "'" & txtstrCodigo & "'" & ", " & "'" & txtstrCodigo & "'" & ", " & "'" & txtintExercicio & "'" & ", " & gstrConvDtParaSql(txtdtmDtData) & ", " & _
             gstrConvDtParaSql(txtdtmDtData) & ", 0", True)
            
strqueryrelatorio = strSql
            
End Function

Public Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    blnContinua = False
    
    If txtstrCodigo = Space$(0) Then
       MsgBox "O campo Protocolo deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
       txtstrCodigo.SetFocus
       Exit Function
    End If
    
    If txtbitDigito = "" Then
       MsgBox "O campo Dígito de Protocolo esta vazio.", vbOKOnly, "Mensagem ao Usuário"
       txtstrCodigo.SetFocus
       Exit Function
    End If
    
    If txtintExercicio = Space$(0) Then
       MsgBox "O campo Exercício de Protocolo deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
       txtintExercicio.SetFocus
       Exit Function
    End If
    
    If Not dbcintTipoProcesso.MatchedWithList Then
        MsgBox "O campo Tipo do Processo deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
        dbcintTipoProcesso.SetFocus
        Exit Function
    End If
    
    If txt_dtmDtdata = Space$(0) Then
       MsgBox "O campo Data deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
       txt_dtmDtdata.SetFocus
       Exit Function
    End If
    
    If gblnDataValida(txt_dtmDtdata.Text) = False Then
       MsgBox "Data inválida.", vbOKOnly, "Mensagem ao Usuário"
       txt_dtmDtdata.SetFocus
       Exit Function
    End If
    
    If dbcintCodContribuinte.BoundText = Space$(0) And dbcintCodCentroCusto.BoundText = Space$(0) Then
       MsgBox "O campo Requerente deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
       If opt_Requerente(0).Value Then
           dbcintCodContribuinte.SetFocus
       Else
           dbcintCodCentroCusto.SetFocus
       End If
       Exit Function
    End If

    If dbcintCodAssunto.BoundText = Space$(0) Then
       MsgBox "O campo Assunto deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
       dbcintCodAssunto.SetFocus
       Exit Function
    End If

    If txtstrSumula = Space$(0) Then
       MsgBox "O campo Súmula deve ser preenchido.", vbOKOnly, "Mensagem ao Usuário"
       txtstrSumula.SetFocus
       Exit Function
    End If
    
    If Not mblnAlterando Or (mblnAlterando And UCase$(strCodigoAtual) <> UCase$(txtstrCodigo.Text)) Then

ProximoCodigo:

        If gblnExisteCodigo(2, gstrProtocolizacaoProcesso, "strCodigo", "'" & txtstrCodigo.Text & "'", "intExercicio", txtintExercicio.Text) Then
            strCodigo = (gstrProximoCodigo(txtstrCodigo, gstrProtocolizacaoProcesso, "strCodigo", gintCodSeguranca, "intExercicio", txtintExercicio.Text, , True))
            If MsgBox("O código informado já se encontra cadastrado. Deseja usar o código " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                txtstrCodigo.SetFocus
                Exit Function
            Else
                txtstrCodigo.Text = strCodigo
                CalculaDigito txtstrCodigo.Text & txtintExercicio
                GoTo ProximoCodigo
            End If
        End If
    End If
    
    blnDadosOk = True
    blnContinua = True
       
End Function

Private Function RetornaUltimoProcesso() As Double

Dim strSql As String
Dim adoRec As ADODB.Recordset

   strSql = Space$(0)
   
   strSql = "SELECT MAX(strCodigo) As UltProc "
   strSql = strSql & "FROM " & gstrProtocolizacaoProcesso & " "
   strSql = strSql & "WHERE intExercicio = " & txtintExercicio
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSql, 10, adoRec) Then
   
       With adoRec
           If Not (.BOF And .EOF) And Not IsNull(!UltProc) Then
              RetornaUltimoProcesso = !UltProc + 1
           Else
              RetornaUltimoProcesso = 1
           End If
       End With
        
   End If
   
   Set gobjBanco = Nothing
   
End Function

Private Sub PreencherVolumes()
Dim adoRec As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
    cbo_Volume.Clear
    
    If gobjBanco.CriaADO("SELECT PkId, intVolume, dtmDtData, strSumula FROM " & gstrProtocolizacaoVolume & " WHERE intProtocolizacaoProcesso = " & txtPKId.Text & " ORDER BY intVolume ", 10, adoRec) Then
        
        With adoRec
            
            ReDim vetProtocolizacaoVolumes(.RecordCount)
           
            Do While Not .EOF
                cbo_Volume.AddItem (!intVolume)
                cbo_Volume.ItemData(cbo_Volume.NewIndex) = (!Pkid)
                
                vetProtocolizacaoVolumes(cbo_Volume.NewIndex).DTMDATA = !dtmDtData
                vetProtocolizacaoVolumes(cbo_Volume.NewIndex).Pkid = !Pkid
                vetProtocolizacaoVolumes(cbo_Volume.NewIndex).strSumula = Space$(0) & !strSumula
                
                .MoveNext
            Loop
            
            If cbo_Volume.ListCount <> 0 Then
                cbo_Volume.ListIndex = 0
            End If
            lbl_TotalVolume.Caption = "/ " & .RecordCount
            
        End With
        
    End If

    Set gobjBanco = Nothing
    
End Sub

Sub LimpaCampos()
        
    cbo_Volume.Clear
    lbl_TotalVolume.Caption = Space$(0)
    dbcintCodAssunto.BoundText = Space$(0)
    txt_intCodCentroCusto = Space$(0)
    txt_CodigoContribuinte = Space$(0)
    txtstrSumula = Space$(0)
    txtbitDigito = Space$(0)
    'txt_dtmDtdata = gstrDataDoSistema
    'txt_dtmDtHora = Format(Time, "Short Time")
    'txtdtmDtData = txt_dtmDtdata + " " + txt_dtmDtHora
    TrocaCorObjeto txt_dtmDtdata, False
    TrocaCorObjeto txt_dtmDtHora, False
    txt_dtmDtdata = ""
    txt_dtmDtHora = "" 'Format(Time, "Short Time")
    'txtintExercicio = Year(Date)
    txtintExercicio = ""
    txtstrCodigo = Space$(0)
    txt_strLogradouroC = Space$(0)
    txt_intBairroC = Space$(0)
    txt_intUFC = Space$(0)
    txt_intCepC = Space$(0)
    txt_intMunicipioC = Space$(0)
    dbcintLogradouroA.BoundText = Space$(0)
    txtintNumeroA = Space$(0)
    txtstrComplementoA = Space$(0)
    txtintCepA = Space$(0)
    txtstrReferenciaA = Space$(0)
    txt_intBairroA = Space$(0)
    opt_Requerente(0).Value = -1
    chkbitEmpenho.Value = 0
    chkbitRevisaoCalculo.Value = 0
    
    txtPKId.Text = ""
    
    TrocaCorObjeto txtstrCodigo, False
    TrocaCorObjeto txtintExercicio, False
    
    
    txtstrCodigo.SetFocus
    mblnAlterando = False
    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCriarVolume
    
End Sub

Private Function strQueryDataComboLogradouroA()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT L.PKId, RTRIM(LTRIM(L.strDescricao)) " & strCONCAT & " ', ' " & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & " " & strCONCAT & "  " & gstrISNULL("U.strDescricao", "' '", "', '") & " " & strCONCAT & " " & gstrISNULL("U.strDescricao", "''") & strCONCAT & " ' - Bairro: ' " & strCONCAT & " BA.strDescricao " & strCONCAT & " ' - Cep: '" & strCONCAT & gstrCONVERT(CDT_VARCHAR, "L.intCep") & ")) AS Logradouro "
    strSql = strSql & "FROM " & gstrLogradouro & " L, " & gstrTituloLogradouro & " U, " & gstrTipoLogradouro & " TL, " & gstrBairro & " BA "
    strSql = strSql & "WHERE L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle & " AND L.intBairro " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle & " "
    strSql = strSql & " AND L.Dtmdtexclusao is null "
    strSql = strSql & "ORDER BY L.strDescricao"
    strQueryDataComboLogradouroA = strSql
End Function

Private Function strQueryDataComboBairroA()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrBairro & " "
    strSql = strSql & "ORDER BY strDescricao"
    strQueryDataComboBairroA = strSql
End Function

Private Function strQueryLocais() As String

'******************************************************************************************
' Data: 21/03/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela variável
'            strISNULL.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 21/03/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 21/03/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 21/03/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 21/03/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 21/03/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 24/03/2003
' Alteração: - Retirado o comando CONVERT da cláusula SELECT uma vez que este não era
'            necessário. Conversão das colunas ORG.strCodigo, UOR.strCodigo, SUB.strCodigo.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 25/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'******************************************************************************************
'
'******************************************************************************************
' Data: 24/06/2003
' Alteração: - Retirada a tabela tblUnidadeCentroDeCusto, pois será usada tblLocais
' Responsável: Gustavo Monteiro
'******************************************************************************************
    Dim strSql As String

'    strSql = " SELECT UC2.PKID, CONVERT(VARCHAR,ORG.strCodigo) + '.' + CONVERT(VARCHAR,UOR.strCodigo) + '.' +"
'    strSql = " SELECT UC2.PKID, RTrim(UC2.strDescricao) " & strCONCAT & " ' (' " & strCONCAT & " ORG.strCodigo " & strCONCAT & " '.' " & strCONCAT & " UOR.strCodigo " & strCONCAT & " '.' " & strCONCAT
'    " CASE ISNULL(SUB.strCodigo,0) WHEN 0 THEN '' ELSE CONVERT(VARCHAR,SUB.strCodigo) + '.' END +"
'    strSQL = strSQL & gstrCASEWHEN(strISNULL & "(SUB.strCodigo,0)", "0,''", "SUB.strCodigo" & strCONCAT & " '.'") & strCONCAT
'    strSql = strSql & gstrISNULL("SUB.strCodigo", "''", "SUB.strCodigo" & strCONCAT & " '.' ") & strCONCAT
'    " CONVERT(VARCHAR, UC1.intCodigo) + '.' +"
'    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "UC1.intCodigo") & strCONCAT & " '.' " & strCONCAT
'    " CONVERT(VARCHAR, UC2.intCodigo) + ' ' +"
'    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "UC2.intCodigo") & strCONCAT & "')' AS CodigoCentroCusto"
'    " FROM " & gstrUnidadeCentroDeCusto2 & " AS UC2, "
'    strSql = strSql & " FROM " & gstrUnidadeCentroDeCusto2 & " UC2, "
'    gstrUnidadeCentroDeCusto1 & " AS UC1, "
'    strSql = strSql & gstrUnidadeCentroDeCusto1 & " UC1, "
'    gstrOrgao & " AS ORG, "
'    strSql = strSql & gstrOrgao & " ORG, "
'    gstrUnidadeOrcamentaria & " AS UOR, "
'    strSql = strSql & gstrUnidadeOrcamentaria & " UOR, "
'    gstrSubUnidade & " AS SUB"
'    strSql = strSql & gstrSubUnidade & " SUB"

'    strSql = strSql & " WHERE  ORG.PKID = UOR.intOrgao AND"
'    " ORG.PKID *= SUB.intOrgao AND"
'    strSql = strSql & " ORG.PKID " & strOUTJSQLServer & strOUTJOracle & "= SUB.intOrgao AND" & _
'    " ORG.PKID = UC1.intOrgao AND" & _
'    " ORG.PKID = UC2.intOrgao AND"
'    " UOR.PKID *= SUB.intUnidadeOrcamentaria AND"
'    strSql = strSql & " UOR.PKID " & strOUTJSQLServer & strOUTJOracle & "= SUB.intUnidadeOrcamentaria AND" & _
'    " UOR.PKID = UC1.intUnidadeOrcamentaria AND" & _
'    " UOR.PKID = UC2.intUnidadeOrcamentaria AND"
'    " SUB.PKID =* UC1.intSubUnidade AND"
'    strSql = strSql & " SUB.PKID =" & strOUTJSQLServer & " UC1.intSubUnidade " & strOUTJOracle & " AND"
'    " SUB.PKID =* UC2.intSubUnidade AND"
'    strSql = strSql & " SUB.PKID =" & strOUTJSQLServer & " UC2.intSubUnidade " & strOUTJOracle & " AND" & _
'    " UC1.PKID = UC2.intUnidadeCentrodeCusto1" & _
'    " ORDER BY CodigoCentroCusto"

    strSql = ""
    strSql = strSql & " SELECT A.PkId, A.strDescricao"
    strSql = strSql & " FROM"
    strSql = strSql & " " & gstrLocais & " A"
    
    strQueryLocais = strSql

End Function

Private Function strQueryEnderecoContrib()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT intLogradouro, intNumero, strComplemento, intBairro, intCep "
    strSql = strSql & " FROM " & gstrContribuinte & " "
    strSql = strSql & " WHERE PKId=" & dbcintCodContribuinte.BoundText
    strQueryEnderecoContrib = strSql
End Function

Private Sub txt_dtmdtdata_LostFocus()
    txt_dtmDtdata.Text = gstrDataFormatada(txt_dtmDtdata.Text)
    txtdtmDtData = Trim(txt_dtmDtdata) & " " & Trim(txt_dtmDtHora)
End Sub

Private Sub txt_dtmDtHora_GotFocus()
    MarcaCampo txt_dtmDtHora
    If txt_dtmDtHora = "" Then txt_dtmDtHora = Mid(gstrDataDoSistema(True, False, False), 12, 8)
    'txtdtmDtData = txt_dtmDtdata + " " + txt_dtmDtHora
End Sub

Private Sub txt_dtmDtHora_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "H", txt_dtmDtHora
End Sub

Private Sub txt_dtmdtdata_GotFocus()
    MarcaCampo txt_dtmDtdata
    If txt_dtmDtdata = "" Then txt_dtmDtdata = gstrDataDoSistema
End Sub

Private Sub txt_dtmdtdata_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDtdata
End Sub

Private Sub txt_dtmDtHora_LostFocus()
    txtdtmDtData = Trim(txt_dtmDtdata) & " " & Trim(txt_dtmDtHora)
End Sub


Private Sub txtbitDigito_GotFocus()
    MarcaCampo txtbitDigito
End Sub

Private Sub txtbitDigito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigito
End Sub

Private Sub txtintCepA_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCepA
End Sub

Private Sub txtintCepA_LostFocus()
    txtintCepA = gstrCEPFormatado(txtintCepA)
    CepLogradouro txtintCepA, dbcintLogradouroA, txt_intBairroA, , , , , , True, False, , , , , True, False
    DoEvents
End Sub

Private Sub txtintExercicio_Change()
    If txtPKId.Text = "" And txtstrCodigo.Text <> "" And txtintExercicio.Text <> "" Then
        If Len(txtintExercicio) = 4 Then CalculaDigito txtstrCodigo.Text & txtintExercicio
    End If
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
    If txtintExercicio = "" Then txtintExercicio = Year(gstrDataDoSistema)
    
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintExercicio_LostFocus()
   If Len(txtintExercicio) > 0 Then
      txtintExercicio = Format("01/01/" & txtintExercicio, "yyyy")
   End If
   If txtPKId.Text = "" And txtstrCodigo.Text <> "" And txtintExercicio.Text <> "" Then
      CalculaDigito txtstrCodigo.Text & txtintExercicio
   End If
   
End Sub

Private Sub txtstrCodigo_GotFocus()
    If txtintExercicio = "" Then txtintExercicio = Year(gstrDataDoSistema)
    gstrProximoCodigo txtstrCodigo, gstrProtocolizacaoProcesso, "strCodigo", gintCodSeguranca, "intExercicio", txtintExercicio
    MarcaCampo txtstrCodigo
    If txtPKId.Text = "" And txtstrCodigo.Text <> "" And txtintExercicio.Text <> "" Then
        CalculaDigito txtstrCodigo.Text & txtintExercicio
    End If

End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub

Private Sub dbcintCodContribuinte_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintCodAssunto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintCodAssunto
End Sub


Private Sub txtstrCodigo_LostFocus()
If txtPKId.Text = "" And txtstrCodigo.Text <> "" And txtintExercicio.Text <> "" Then
    CalculaDigito txtstrCodigo.Text & txtintExercicio
End If
   
End Sub

Private Sub txtstrSumula_GotFocus()
    MarcaCampo txtstrSumula
End Sub

Private Sub txtstrSumula_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrSumula
End Sub

Private Sub chkbitEmpenho_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub chkbitRevisaoCalculo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub


Private Sub CalculaDigito(strCodigo As String)

Dim intTam      As Integer
Dim intDv       As Integer
Dim intMULT     As Integer
Dim intResult   As Integer
Dim intSoma     As Integer
Dim intResto    As Integer
Dim blnV_Resp   As Boolean
Dim strChar2    As String

intTam = Len(strCodigo)
intMULT = 1

blnV_Resp = True

Do While intTam >= 1
    intMULT = intMULT + 1
    If intMULT > 2 Then
        intMULT = 1
    End If
    intResult = Mid(strCodigo, intTam, 1) * intMULT
    
    If intResult >= 10 Then
        strChar2 = intResult
        intSoma = intSoma + Val(Mid(strChar2, 1, 1)) + Val(Mid(strChar2, 2, 1))
    Else
        intSoma = intSoma + intResult
    End If
    intTam = intTam - 1

Loop
intResto = intSoma Mod 10
If intResto = 0 Then
    intDv = 0
Else
    intDv = 10 - intResto
End If

txtbitDigito.Text = intDv

End Sub

Private Function PreencheBairroA() As String
Dim strSql As String
Dim adoRec As ADODB.Recordset

strSql = ""
strSql = "SELECT BA.strDescricao "
strSql = strSql & " FROM " & gstrBairro & " BA , " & gstrProtocolizacaoProcesso & " PP "
strSql = strSql & " WHERE PP.Pkid = " & txtPKId
strSql = strSql & " AND BA.Pkid = PP.intBairroA"
        
Set gobjBanco = New clsBanco
            
If gobjBanco.CriaADO(strSql, 5, adoRec) Then
    If Not adoRec.EOF Then
        txt_intBairroA = adoRec!strDescricao
    Else
        txt_intBairroA = ""
    End If
End If

PreencheBairroA = strSql

Set gobjBanco = Nothing

End Function

Private Sub CriaPrimeiroTramite()
        
    Load frmSelecionaLocal
    frmSelecionaLocal.intExercicio = txtintExercicio.Text
    
    frmSelecionaLocal.intProtocolo = txtstrCodigo.Text
    frmSelecionaLocal.dbcintCustoDestino.BoundText = txtintCentroCusto.Text
    'frmSelecionaLocal.intvolume = glngPegaUltimaChave(gstrProtocolizacaoVolume, "intVolume", "intProtocolizacaoProcesso", txtPKId.Text)
   
    frmSelecionaLocal.Show vbModal

End Sub

Private Function strQueryTipoProcesso() As String

    Dim strSql          As String
    
    strSql = "SELECT PKId, strCodigo FROM " & gstrTipoProcesso
    strSql = strSql & " ORDER BY strCodigo"

    strQueryTipoProcesso = strSql
    
End Function

Private Function strQueryContribuinteEspecifico(lngPKId As Long) As String
Dim strSql As String

strSql = "SELECT Pkid, strNome"
strSql = strSql & " FROM "
strSql = strSql & gstrContribuinte
strSql = strSql & " WHERE Pkid = " & lngPKId

strQueryContribuinteEspecifico = strSql

End Function

