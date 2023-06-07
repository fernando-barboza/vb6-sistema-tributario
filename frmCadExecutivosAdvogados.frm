VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadExecutivosAdvogados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advogados"
   ClientHeight    =   7200
   ClientLeft      =   1530
   ClientTop       =   3810
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8925
   Begin TabDlg.SSTab tab_Advogados 
      Height          =   4215
      Left            =   90
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Advogados"
      TabPicture(0)   =   "frmCadExecutivosAdvogados.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintCodigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintContribuinte"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_CNPJCPF"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcintContribuinte"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtintCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmd_Contribuinte"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt_CNPJCPF"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_OAB"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_Endereco"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtPKId"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fra_LogoBanco"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.Frame fra_LogoBanco 
         Caption         =   "Assinatura Digitalizada"
         Height          =   1365
         Left            =   6690
         TabIndex        =   30
         Top             =   360
         Width           =   1905
         Begin VB.TextBox txtintImagem 
            Height          =   285
            Left            =   480
            TabIndex        =   31
            Top             =   810
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Image img_Assinatura 
            BorderStyle     =   1  'Fixed Single
            Height          =   1050
            Left            =   90
            MouseIcon       =   "frmCadExecutivosAdvogados.frx":001C
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1740
         End
      End
      Begin VB.TextBox txtPKId 
         Enabled         =   0   'False
         Height          =   225
         Left            =   2610
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   " Endereço residencial"
         Height          =   1515
         Left            =   120
         TabIndex        =   12
         Top             =   1710
         Width           =   8490
         Begin VB.TextBox txt_strLogradouro 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1050
            TabIndex        =   28
            Top             =   300
            Width           =   4245
         End
         Begin VB.TextBox txt_intNumero 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            MaxLength       =   8
            TabIndex        =   20
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox txt_strComplemento 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7110
            MaxLength       =   20
            TabIndex        =   19
            Top             =   300
            Width           =   1230
         End
         Begin VB.TextBox txt_intCep 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6630
            MaxLength       =   9
            TabIndex        =   18
            Top             =   1050
            Width           =   1080
         End
         Begin VB.TextBox txt_strBairro 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1050
            TabIndex        =   17
            Top             =   675
            Width           =   3975
         End
         Begin VB.TextBox txt_strMunicipio 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1050
            TabIndex        =   16
            Top             =   1050
            Width           =   3975
         End
         Begin VB.TextBox txt_strUF 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5385
            TabIndex        =   15
            Top             =   1050
            Width           =   705
         End
         Begin VB.Label lbl_strtMunicipio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   255
            TabIndex        =   27
            Top             =   1110
            Width           =   705
         End
         Begin VB.Label lbl_strBairro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   540
            TabIndex        =   26
            Top             =   780
            Width           =   405
         End
         Begin VB.Label lbl_strLogradouro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   150
            TabIndex        =   25
            Top             =   420
            Width           =   810
         End
         Begin VB.Label lbl_intNumero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5370
            TabIndex        =   24
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lblstrComplemento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6570
            TabIndex        =   23
            Top             =   390
            Width           =   480
         End
         Begin VB.Label lbl_strUf 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   5100
            TabIndex        =   22
            Top             =   1140
            Width           =   210
         End
         Begin VB.Label lbl_intCep 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   6255
            TabIndex        =   21
            Top             =   1140
            Width           =   285
         End
      End
      Begin VB.Frame fra_OAB 
         Caption         =   " OAB"
         Height          =   765
         Left            =   120
         TabIndex        =   10
         Top             =   3300
         Width           =   8505
         Begin VB.TextBox txtstrOABnumero 
            Height          =   285
            Left            =   2910
            MaxLength       =   7
            TabIndex        =   4
            Top             =   270
            Width           =   885
         End
         Begin MSDataListLib.DataCombo dbcintUf 
            Height          =   315
            Left            =   960
            TabIndex        =   3
            Top             =   240
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblstrOABnumero 
            Caption         =   "OAB número"
            Height          =   225
            Left            =   1890
            TabIndex        =   13
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label lblintUf 
            Caption         =   "Unid. Fed."
            Height          =   315
            Left            =   150
            TabIndex        =   11
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.TextBox txt_CNPJCPF 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1230
         MaxLength       =   19
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1290
         Width           =   1600
      End
      Begin VB.CommandButton cmd_Contribuinte 
         Height          =   315
         Left            =   6090
         Picture         =   "frmCadExecutivosAdvogados.frx":0326
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Ativa Cadastro de Contribuintes"
         Top             =   900
         Width           =   360
      End
      Begin VB.TextBox txtintCodigo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1230
         MaxLength       =   9
         TabIndex        =   0
         Top             =   540
         Width           =   1170
      End
      Begin MSDataListLib.DataCombo dbcintContribuinte 
         Height          =   315
         Left            =   1230
         TabIndex        =   1
         Top             =   900
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lbl_CNPJCPF 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ / CPF"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   1380
         Width           =   870
      End
      Begin VB.Label lblintContribuinte 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   645
         TabIndex        =   7
         Top             =   660
         Width           =   495
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Advogados 
      Height          =   2775
      Left            =   90
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4320
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   4895
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "pkid"
      Columns(0).DataField=   "PKId"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Código"
      Columns(1).DataField=   "intCodigo"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nome"
      Columns(2).DataField=   "strNome"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   16
      Columns(3)._MaxComboItems=   5
      Columns(3).ValueItems(0)._DefaultItem=   0
      Columns(3).ValueItems(0).Value=   "0"
      Columns(3).ValueItems(0).Value.vt=   8
      Columns(3).ValueItems(0).DisplayValue=   "0"
      Columns(3).ValueItems(0).DisplayValue.vt=   8
      Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems.Count=   1
      Columns(3).Caption=   "CPF / CNPJ"
      Columns(3).DataField=   "strCNPJCPF"
      Columns(3).NumberFormat=   "FormatText Event"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Unid. Fed."
      Columns(4).DataField=   "strSigla"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "OAB número"
      Columns(5).DataField=   "strOABnumero"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=820"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=741"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2037"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1958"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=7038"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6959"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=3122"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=3043"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=1508"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1429"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTips        =   1
      CellTipsWidth   =   0
      MultiSelect     =   0
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadExecutivosAdvogados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando             As Boolean
Dim mobjAux                   As Object
Dim mblnSelecionou            As Boolean
Dim mblnPrimeiraVez           As Boolean
Dim mblnClickOk               As Boolean

Private Sub dbcintContribuinte_Click(Area As Integer)
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
     
    DropDownDataCombo dbcintContribuinte, Me, Area

    If Area = 2 Then
        If dbcintContribuinte.BoundText = "" Then
            Exit Sub
        End If
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strQueryLogradouro, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                With adoResultado
                    txt_CNPJCPF = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!StrCnpjCpf))
                    txt_strLogradouro = gstrVerificaCampoNulo(!strLogradouro)
                    txt_intNumero = gstrVerificaCampoNulo(!INTNUMERO)
                    txt_strComplemento = gstrVerificaCampoNulo(!STRCOMPLEMENTO)
                    txt_strBairro = gstrVerificaCampoNulo(!strBairro)
                    txt_strMunicipio = gstrVerificaCampoNulo(!STRMUNICIPIO)
                    txt_strUF = gstrVerificaCampoNulo(!STRUF)
                    txt_intCep = gstrCEPFormatado(gstrVerificaCampoNulo(!INTCEP))
                End With
            End If
        End If
    End If
End Sub

Private Sub cmd_Contribuinte_Click()
    ChamaFormCadastro frmCadContribuinte, dbcintContribuinte
End Sub


''Private Sub Form_Activate()
''    gintCodSeguranca = 626
''    VirificaGradeListView Me
''    If mblnSelecionou Then
''        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
''    Else
''        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
''    End If
'''    If mobjAux Is Nothing Then
''        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
''    Else
''        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
''    End If
''End Sub
'
Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    mblnPrimeiraVez = False
    mblnAlterando = False
    dbcintContribuinte.Tag = strQueryDataComboContribuinte & ";strNome"
    dbcintUf.Tag = gstrQueryDataComboUF & ";strSigla"
    
    TrocaCorObjeto txt_CNPJCPF, True
    TrocaCorObjeto txt_strLogradouro, True
    TrocaCorObjeto txt_intNumero, True
    TrocaCorObjeto txt_strComplemento, True
    TrocaCorObjeto txt_strBairro, True
    TrocaCorObjeto txt_strMunicipio, True
    TrocaCorObjeto txt_strUF, True
    TrocaCorObjeto txt_intCep, True
    
End Sub
Private Function blnDadosOk() As Boolean

blnDadosOk = False

    If txtintCodigo.Text = "" Then
        ExibeMensagem "O código deve ser informado."
        txtintCodigo.SetFocus
        Exit Function
    End If
        
    If dbcintContribuinte.BoundText = "" Then
        ExibeMensagem "O nome deve ser informado."
        dbcintContribuinte.SetFocus
        Exit Function
    End If
    
    If dbcintUf.BoundText = "" Then
        ExibeMensagem "A unidade da federação deve ser informada."
        dbcintUf.SetFocus
        Exit Function
    End If
    
    If txtstrOABnumero.Text = "" Then
        ExibeMensagem "O OAB número deve ser informado."
        txtstrOABnumero.SetFocus
        Exit Function
    End If

blnDadosOk = True

End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSql As String
    
    Select Case UCase(strModoOperacao)
          
           Case UCase(gstrSalvar)
                If blnDadosOk Then
                    ToolBarGeral strModoOperacao, gstrExecutivoAdvogados, mblnAlterando, tdb_Advogados, Me, , strQuery
                    ToolBarGeral gstrLocalizar, gstrExecutivoAdvogados, mblnAlterando, tdb_Advogados, Me, , strQuery, , , , True
                    MantemForm gstrNovo
                    Exit Sub
                End If
           Case UCase(gstrLogotipo)
                frmCadImagem.CadastraFoto img_Assinatura, txtintImagem
                frmCadImagem.Caption = "Assinatura Digitalizada"
                Exit Sub
           Case UCase(gstrNovo)
                    Limpa_Controles frmCadExecutivosAdvogados, True, False, False, True, False
                    dbcintContribuinte.Text = ""
                    txtintCodigo_GotFocus
                    mblnAlterando = False
                    mblnPrimeiraVez = False
                    Set img_Assinatura.Picture = Nothing
           Case UCase(gstrDeletar)
                ToolBarGeral strModoOperacao, gstrExecutivoAdvogados, mblnAlterando, tdb_Advogados, Me, , strQuery, , , , True
                ToolBarGeral gstrLocalizar, gstrExecutivoAdvogados, mblnAlterando, tdb_Advogados, Me, , strQuery, , , , True
           Case Else
                ToolBarGeral strModoOperacao, gstrExecutivoAdvogados, mblnAlterando, tdb_Advogados, Me, , strQuery
    End Select
      

End Sub

Private Function strQueryDataComboContribuinte()
    Dim strSql As String
        
        strSql = ""
        strSql = strSql & "SELECT PKId, strNome "
        strSql = strSql & "FROM " & gstrContribuinte & " "
        strSql = strSql & "ORDER BY strNome"
        
        strQueryDataComboContribuinte = strSql

End Function

Private Function strQuery() As String
    Dim strSql As String
    
     strSql = strSql & "SELECT "
     strSql = strSql & " EA.pkid pkid, "
     strSql = strSql & " EA.intCodigo intCodigo, "
     strSql = strSql & " CO.strNome strNome, "
     strSql = strSql & " CO.strCNPJCPF strCNPJCPF, "
     strSql = strSql & " UF.strSigla strSigla, "
     strSql = strSql & " EA.strOABnumero strOABnumero "
     strSql = strSql & "FROM "
     strSql = strSql & gstrExecutivoAdvogados & " EA, "
     strSql = strSql & gstrContribuinte & " CO, "
     strSql = strSql & gstrUF & " UF "
     strSql = strSql & "WHERE "
     strSql = strSql & " EA.intContribuinte = CO.Pkid AND"
     strSql = strSql & " EA.intUF = UF.pkid "
     
     strQuery = strSql
     
End Function

Private Function strQueryLogradouro() As String
    Dim strSql As String
    
        strSql = ""
        strSql = strSql & "SELECT "
        strSql = strSql & " CO.Pkid , "
        strSql = strSql & " CO.strNome, "
        strSql = strSql & " CO.strCNPJCPF, "
        strSql = strSql & " MU.strDescricao strMunicipio, "
        strSql = strSql & " BA.strDescricao strBairro, "
        strSql = strSql & " CO.intNumero intNumero, "
        strSql = strSql & " CO.strComplemento strComplemento, "
        strSql = strSql & " UF.strSigla strUF, "
        strSql = strSql & " CO.intCep intCep, "
        strSql = strSql & " RTRIM(LTRIM(LO.strDescricao)) " & strCONCAT & gstrISNULL("TL.strSigla", "''", "', '") & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''")
        strSql = strSql & strCONCAT & gstrISNULL("TI.strDescricao", "' '", "', '") & strCONCAT & gstrISNULL("TI.strDescricao", "''") & ")) " & strCONCAT & "' ( '" & strCONCAT & gstrISNULL("BA.strDescricao", "''") & strCONCAT & "' ) '" & " AS strLogradouro "
        strSql = strSql & "FROM "
        strSql = strSql & gstrContribuinte & " CO, "
        strSql = strSql & gstrBairro & " BA, "
        strSql = strSql & gstrLogradouro & " LO, "
        strSql = strSql & gstrUF & " UF, "
        strSql = strSql & gstrTituloLogradouro & " TI, "
        strSql = strSql & gstrTipoLogradouro & " TL, "
        strSql = strSql & " tblMunicipio MU "
        strSql = strSql & "WHERE "
        strSql = strSql & " CO.intMunicipio " & strOUTJSQLServer & "= MU.PKId " & strOUTJOracle
        strSql = strSql & " AND LO.intBairro " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle
        strSql = strSql & " AND LO.pkid = CO.intLogradouro "
        strSql = strSql & " AND CO.intUf " & strOUTJSQLServer & "= UF.PKId " & strOUTJOracle
        strSql = strSql & " AND LO.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
        strSql = strSql & " AND LO.intTituloLogradouro " & strOUTJSQLServer & "= TI.PKId " & strOUTJOracle
        strSql = strSql & " AND CO.pkid = " & dbcintContribuinte.BoundText

    strQueryLogradouro = strSql

End Function

Private Sub img_Assinatura_DblClick()
    MantemForm gstrLogotipo
End Sub

Private Sub img_Assinatura_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Set img_Assinatura.Picture = Nothing
        txtintImagem = ""
    End If
End Sub

Private Sub tdb_Advogados_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Advogados) = 1 Then
        tdb_Advogados_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Advogados_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Advogados
End Sub

Private Sub tdb_Advogados_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Advogados, ColIndex
End Sub

Private Sub tdb_Advogados_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Advogados_KeyPress(KeyAscii As Integer)
    Select Case tdb_Advogados.Col
        Case 0
            CaracterValido KeyAscii, "N", tdb_Advogados
        Case Else
            CaracterValido KeyAscii, "A", tdb_Advogados
    End Select
End Sub

Private Sub tdb_Advogados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Advogados_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Advogados
        If Not .EOF And Not .BOF And mblnClickOk = True Then
            If mblnPrimeiraVez Then
                mblnAlterando = True
                txtPKId.Text = .Columns("PKID").Value
                LeDaTabelaParaObj gstrExecutivoAdvogados, Me
                 gCorLinhaSelecionada tdb_Advogados
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                dbcintContribuinte_Click 2
                PreencheImagem txtPKId
            End If
        End If
    End With
End Sub

Private Sub txtintCodigo_GotFocus()
    gstrProximoCodigo txtintCodigo, gstrExecutivoAdvogados, "intCodigo", 1393
    MarcaCampo txtintCodigo
End Sub

Private Sub txtintCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintCodigo
End Sub

Private Sub PreencheImagem(intPkid As Long)
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT * FROM " & gstrExecutivoAdvogados & " WHERE pkid = " & intPkid
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            LeImagem Val(gstrENulo(adoResultado("intImagem").Value)), img_Assinatura
        End If
    End If

End Sub
