VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadEconomico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Econômico"
   ClientHeight    =   7410
   ClientLeft      =   1155
   ClientTop       =   2280
   ClientWidth     =   11805
   HelpContextID   =   6
   Icon            =   "CadEconomico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11805
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1965
      Left            =   30
      TabIndex        =   199
      Top             =   5415
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   3466
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Inscrição Cadastral"
      Columns(0).DataField=   "strInscricaoCadastral"
      Columns(0).NumberFormat=   "FormatText Event"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Razão Social"
      Columns(1).DataField=   "strNome"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Código"
      Columns(2).DataField=   "PKId"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4895"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4815"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=2"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=12991"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=12912"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=688"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=609"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5355
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   9446
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "CadEconomico.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Endereco"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Inscricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datas, Horários e Áreas"
      TabPicture(1)   =   "CadEconomico.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Processos"
      Tab(1).Control(1)=   "fra_Abertura"
      Tab(1).Control(2)=   "fra_Encerramento"
      Tab(1).Control(3)=   "fra_Outros"
      Tab(1).Control(4)=   "fra_Horarios"
      Tab(1).Control(5)=   "fra_Ocorrencia"
      Tab(1).Control(6)=   "fra_ValorArbitrado"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Atividades e ISSQN"
      TabPicture(2)   =   "CadEconomico.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_ISS"
      Tab(2).Control(1)=   "fra_Atividades"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Tributos e  Faixas"
      TabPicture(3)   =   "CadEconomico.frx":1096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Histórico"
      TabPicture(4)   =   "CadEconomico.frx":10B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fra_Historico"
      Tab(4).Control(1)=   "Fra_OcorrenciaProcesso"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Sócios e Contadores"
      TabPicture(5)   =   "CadEconomico.frx":10CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fra_Socios"
      Tab(5).Control(1)=   "fra_Contador"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Características"
      TabPicture(6)   =   "CadEconomico.frx":10EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "tab_3dCaracteristicas"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Publicidade"
      TabPicture(7)   =   "CadEconomico.frx":1106
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fra_Publicidades"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Feiras"
      TabPicture(8)   =   "CadEconomico.frx":1122
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Fra_Feiras"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Histórico Complementar"
      TabPicture(9)   =   "CadEconomico.frx":113E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Fra_TipoHistorico"
      Tab(9).Control(1)=   "tdb_HistOcorrencias"
      Tab(9).ControlCount=   2
      Begin VB.Frame fra_Processos 
         Caption         =   "Ocorrências de Processos"
         Height          =   1635
         Left            =   -74820
         TabIndex        =   76
         Top             =   1560
         Width           =   11370
         Begin VB.TextBox txtstrhistoricoprocesso 
            Height          =   600
            Left            =   120
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Top             =   930
            Width           =   11160
         End
         Begin VB.TextBox txtdtmdataprocesso 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   495
            TabIndex        =   78
            Top             =   210
            Width           =   1005
         End
         Begin VB.TextBox txtbitdigprocesso 
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
            Left            =   6570
            MaxLength       =   2
            TabIndex        =   82
            Top             =   210
            Width           =   285
         End
         Begin VB.TextBox txtintexerprocesso 
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
            Left            =   6090
            MaxLength       =   4
            TabIndex        =   81
            Top             =   210
            Width           =   465
         End
         Begin VB.TextBox txtstrcodprocesso 
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
            Left            =   5250
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   80
            Top             =   210
            Width           =   825
         End
         Begin MSDataListLib.DataCombo dbcintocorrenciadoeconomico 
            Height          =   315
            Left            =   2025
            TabIndex        =   84
            Top             =   540
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrências de Processo"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   270
            Width           =   345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nº do Processo"
            Height          =   195
            Left            =   4095
            TabIndex        =   79
            Top             =   270
            Width           =   1110
         End
      End
      Begin VB.Frame Fra_OcorrenciaProcesso 
         Caption         =   "Ocorrêcias de Processos"
         Height          =   2205
         Left            =   -74850
         TabIndex        =   146
         Top             =   3030
         Width           =   11355
         Begin TrueOleDBGrid70.TDBGrid tdb_Processo 
            Height          =   1785
            Left            =   300
            TabIndex        =   148
            Top             =   300
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   3149
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Data do Processo"
            Columns(0).DataField=   "Dtmdtprocesso"
            Columns(0).NumberFormat=   "FormatText Event"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Processo"
            Columns(1).DataField=   "strProcesso"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Ocorrência"
            Columns(2).DataField=   "strOcorrencia"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Observação"
            Columns(3).DataField=   "strObservacao"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=2"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=4392"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=4313"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=18256"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=18177"
            Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame Fra_TipoHistorico 
         Caption         =   "Tipo"
         Height          =   765
         Left            =   -71070
         TabIndex        =   207
         Top             =   450
         Width           =   3735
         Begin VB.ComboBox cbo_intTipo 
            Height          =   315
            ItemData        =   "CadEconomico.frx":115A
            Left            =   150
            List            =   "CadEconomico.frx":115C
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   270
            Width           =   3465
         End
      End
      Begin VB.Frame Fra_Feiras 
         Caption         =   "Feiras"
         Height          =   4335
         Left            =   -74228
         TabIndex        =   194
         Top             =   525
         Width           =   9570
         Begin VB.CommandButton cmd_TipoFeira 
            Height          =   315
            Left            =   7965
            Picture         =   "CadEconomico.frx":115E
            Style           =   1  'Graphical
            TabIndex        =   201
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro Único"
            Top             =   630
            Width           =   360
         End
         Begin VB.CommandButton cmd_Feira 
            Height          =   315
            Left            =   4005
            Picture         =   "CadEconomico.frx":14E8
            Style           =   1  'Graphical
            TabIndex        =   197
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro Único"
            Top             =   630
            Width           =   360
         End
         Begin VB.TextBox txt_areaFeira 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1365
            MaxLength       =   10
            TabIndex        =   203
            Top             =   1080
            Width           =   1350
         End
         Begin VB.TextBox txt_strnrbox 
            Height          =   285
            Left            =   5325
            MaxLength       =   13
            TabIndex        =   205
            Top             =   1080
            Width           =   1350
         End
         Begin MSComctlLib.ListView lvw_Itens 
            Height          =   2235
            Left            =   240
            TabIndex        =   206
            Top             =   1920
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   3942
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
               Text            =   "IntFeira"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Feira"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "intTipoFeira"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Tipo da Feira"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Área"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Número do Box"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbc_intTipoFeira 
            Height          =   315
            Left            =   5325
            TabIndex        =   200
            Top             =   630
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intFeira 
            Height          =   315
            Left            =   1365
            TabIndex        =   196
            Top             =   630
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Feira"
            Height          =   195
            Left            =   4560
            TabIndex        =   198
            Top             =   720
            Width           =   705
         End
         Begin VB.Label lblSTRNRBOX 
            AutoSize        =   -1  'True
            Caption         =   "Número do Box"
            Height          =   195
            Left            =   4170
            TabIndex        =   204
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Feira"
            Height          =   195
            Left            =   960
            TabIndex        =   195
            Top             =   720
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Área"
            Height          =   195
            Left            =   975
            TabIndex        =   202
            Top             =   1140
            Width           =   330
         End
      End
      Begin VB.Frame fra_Abertura 
         Caption         =   " Abertura "
         Height          =   1005
         Left            =   -74835
         TabIndex        =   52
         Top             =   510
         Width           =   3000
         Begin VB.TextBox txtstrCodProcAbertura 
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
            Left            =   1260
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   600
            Width           =   825
         End
         Begin VB.TextBox txtintExerProcAbertura 
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
            Left            =   2100
            MaxLength       =   4
            TabIndex        =   57
            Top             =   600
            Width           =   465
         End
         Begin VB.TextBox txtbitDigProcAbertura 
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
            MaxLength       =   4
            TabIndex        =   58
            Top             =   600
            Width           =   285
         End
         Begin VB.TextBox txtdtmDataAbertura 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1260
            TabIndex        =   54
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label lblstrNumeroProcessoAbertura 
            AutoSize        =   -1  'True
            Caption         =   "Nº do Processo"
            Height          =   195
            Left            =   105
            TabIndex        =   55
            Top             =   660
            Width           =   1110
         End
         Begin VB.Label lbldtmDataAbertura 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   855
            TabIndex        =   53
            Top             =   270
            Width           =   345
         End
      End
      Begin VB.Frame fra_Encerramento 
         Caption         =   " Encerramento "
         Height          =   1005
         Left            =   -71790
         TabIndex        =   59
         Top             =   510
         Width           =   2985
         Begin VB.TextBox txtstrCodProcEncerramento 
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
            Left            =   1260
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   63
            Top             =   600
            Width           =   825
         End
         Begin VB.TextBox txtintExerProcEncerramento 
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
            Left            =   2100
            MaxLength       =   4
            TabIndex        =   64
            Top             =   600
            Width           =   465
         End
         Begin VB.TextBox txtbitDigProcEncerramento 
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
            TabIndex        =   65
            Top             =   600
            Width           =   285
         End
         Begin VB.TextBox txtdtmDataEncerramento 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1245
            TabIndex        =   61
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label lblstrNumeroProcessoEncerramento 
            AutoSize        =   -1  'True
            Caption         =   "Nº do Processo"
            Height          =   195
            Left            =   105
            TabIndex        =   62
            Top             =   660
            Width           =   1110
         End
         Begin VB.Label lbldtmDataEncerramento 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   855
            TabIndex        =   60
            Top             =   270
            Width           =   345
         End
      End
      Begin VB.Frame fra_Outros 
         Caption         =   " Áreas "
         Height          =   1005
         Left            =   -68745
         TabIndex        =   66
         Top             =   510
         Width           =   2235
         Begin VB.TextBox txtdblAreaAnuncio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   870
            MaxLength       =   9
            TabIndex        =   70
            Top             =   600
            Width           =   1245
         End
         Begin VB.TextBox txtdblAreaOcupada 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   870
            MaxLength       =   9
            TabIndex        =   68
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label lbldblAreaAnuncio 
            AutoSize        =   -1  'True
            Caption         =   "Anúncio"
            Height          =   195
            Left            =   210
            TabIndex        =   69
            Top             =   690
            Width           =   585
         End
         Begin VB.Label lbldblAreaOcupada 
            AutoSize        =   -1  'True
            Caption         =   "Ocupada"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   300
            Width           =   660
         End
      End
      Begin VB.Frame fra_Horarios 
         Caption         =   " Horários"
         Height          =   1920
         Left            =   -74820
         TabIndex        =   86
         Top             =   3255
         Width           =   5400
         Begin VB.TextBox txtstrManhaDe 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1830
            TabIndex        =   90
            Top             =   555
            Width           =   975
         End
         Begin VB.TextBox txtstrTardeDe 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1830
            TabIndex        =   95
            Top             =   885
            Width           =   975
         End
         Begin VB.TextBox txtstrNoiteDe 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1830
            TabIndex        =   100
            Top             =   1215
            Width           =   975
         End
         Begin VB.TextBox txtstrMadrugadaDe 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1830
            TabIndex        =   105
            Top             =   1545
            Width           =   975
         End
         Begin VB.TextBox txtstrMadrugadaAte 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3990
            TabIndex        =   107
            Top             =   1545
            Width           =   975
         End
         Begin VB.TextBox txtstrNoiteAte 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3990
            TabIndex        =   102
            Top             =   1215
            Width           =   975
         End
         Begin VB.TextBox txtstrTardeAte 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3990
            TabIndex        =   97
            Top             =   885
            Width           =   975
         End
         Begin VB.TextBox txtstrManhaAte 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3990
            TabIndex        =   92
            Top             =   555
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dbcintHorarioFuncionamento 
            Height          =   315
            Left            =   1830
            TabIndex        =   88
            Top             =   180
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_HorarioEsp 
            AutoSize        =   -1  'True
            Caption         =   "Horário Funcionamento"
            Height          =   195
            Left            =   150
            TabIndex        =   87
            Top             =   300
            Width           =   1650
         End
         Begin VB.Label lblstrManha 
            AutoSize        =   -1  'True
            Caption         =   "Manhã de"
            Height          =   195
            Left            =   1065
            TabIndex        =   89
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lblstrTarde 
            AutoSize        =   -1  'True
            Caption         =   "Tarde de"
            Height          =   195
            Left            =   1140
            TabIndex        =   94
            Top             =   975
            Width           =   645
         End
         Begin VB.Label lblstrNoite 
            AutoSize        =   -1  'True
            Caption         =   "Noite de"
            Height          =   195
            Left            =   1185
            TabIndex        =   99
            Top             =   1305
            Width           =   600
         End
         Begin VB.Label lblstrMadrugada 
            AutoSize        =   -1  'True
            Caption         =   "Madrugada de"
            Height          =   195
            Left            =   750
            TabIndex        =   104
            Top             =   1635
            Width           =   1035
         End
         Begin VB.Label lbl_Ate1 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   3300
            TabIndex        =   91
            Top             =   645
            Width           =   225
         End
         Begin VB.Label lbl_Ate4 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   3300
            TabIndex        =   106
            Top             =   1635
            Width           =   225
         End
         Begin VB.Label lbl_Ate3 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   3300
            TabIndex        =   101
            Top             =   1305
            Width           =   225
         End
         Begin VB.Label lbl_Ate2 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   3300
            TabIndex        =   96
            Top             =   975
            Width           =   225
         End
         Begin VB.Label lbl_Hs1 
            AutoSize        =   -1  'True
            Caption         =   "Hs."
            Height          =   195
            Left            =   5010
            TabIndex        =   93
            Top             =   645
            Width           =   240
         End
         Begin VB.Label lbl_Hs4 
            AutoSize        =   -1  'True
            Caption         =   "Hs."
            Height          =   195
            Left            =   5010
            TabIndex        =   108
            Top             =   1635
            Width           =   240
         End
         Begin VB.Label lbl_Hs3 
            AutoSize        =   -1  'True
            Caption         =   "Hs."
            Height          =   195
            Left            =   5010
            TabIndex        =   103
            Top             =   1305
            Width           =   240
         End
         Begin VB.Label lbl_Hs2 
            AutoSize        =   -1  'True
            Caption         =   "Hs."
            Height          =   195
            Left            =   5010
            TabIndex        =   98
            Top             =   975
            Width           =   240
         End
      End
      Begin VB.Frame fra_Ocorrencia 
         Height          =   1935
         Left            =   -69360
         TabIndex        =   109
         Top             =   3260
         Width           =   5910
         Begin VB.TextBox txtintComponentes 
            Height          =   300
            Left            =   5040
            TabIndex        =   115
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtintNumDeEmpregados 
            Height          =   300
            Left            =   1560
            TabIndex        =   114
            Top             =   615
            Width           =   735
         End
         Begin VB.CommandButton cmd_Ocorrencia 
            Height          =   315
            Left            =   5430
            Picture         =   "CadEconomico.frx":1872
            Style           =   1  'Graphical
            TabIndex        =   112
            TabStop         =   0   'False
            Tag             =   "584"
            ToolTipText     =   "Ativa Cadastro de  Ocorrências"
            Top             =   240
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintOcorrencia 
            Height          =   315
            Left            =   1545
            TabIndex        =   111
            Top             =   240
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "N. de Componentes"
            Height          =   195
            Left            =   3570
            TabIndex        =   210
            Top             =   660
            Width           =   1410
         End
         Begin VB.Label lblintOcorrencia 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrência"
            Height          =   195
            Left            =   720
            TabIndex        =   110
            Top             =   330
            Width           =   780
         End
         Begin VB.Label lbl_NumeroDeEmpregados 
            AutoSize        =   -1  'True
            Caption         =   "N. de Empregados"
            Height          =   195
            Left            =   180
            TabIndex        =   113
            Top             =   675
            Width           =   1320
         End
      End
      Begin VB.Frame fra_ValorArbitrado 
         Height          =   1005
         Left            =   -66450
         TabIndex        =   71
         Top             =   510
         Width           =   3000
         Begin VB.TextBox txtdblValorEstimado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            MaxLength       =   9
            TabIndex        =   73
            Top             =   210
            Width           =   1275
         End
         Begin VB.TextBox txtdtmDataEstimativa 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1620
            TabIndex        =   75
            Top             =   615
            Width           =   1275
         End
         Begin VB.Label lblValorEstimado 
            AutoSize        =   -1  'True
            Caption         =   "Valor da Est. Mensal"
            Height          =   195
            Left            =   90
            TabIndex        =   72
            Top             =   270
            Width           =   1455
         End
         Begin VB.Label lblDatadaEstimativa 
            AutoSize        =   -1  'True
            Caption         =   "Data da Estimativa"
            Height          =   195
            Left            =   210
            TabIndex        =   74
            Top             =   675
            Width           =   1335
         End
      End
      Begin VB.Frame fra_Publicidades 
         Caption         =   "Publicidades"
         Height          =   4335
         Left            =   -74228
         TabIndex        =   179
         Top             =   525
         Width           =   9570
         Begin VB.TextBox txt_dtmPublicidadeInicio 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2070
            TabIndex        =   190
            Top             =   1830
            Width           =   1095
         End
         Begin VB.TextBox txt_dtmPublicidadeFim 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7920
            TabIndex        =   192
            Top             =   1830
            Width           =   1095
         End
         Begin VB.CommandButton cmd_TipoDePublicidade 
            Height          =   315
            Left            =   8640
            Picture         =   "CadEconomico.frx":1990
            Style           =   1  'Graphical
            TabIndex        =   182
            TabStop         =   0   'False
            Tag             =   "584"
            ToolTipText     =   "Ativa Cadastro de Tipos de Publicidade"
            Top             =   525
            Width           =   360
         End
         Begin VB.TextBox txt_intQuantidade 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2100
            TabIndex        =   184
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox txt_dblArea 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7815
            TabIndex        =   186
            Top             =   975
            Width           =   1185
         End
         Begin VB.TextBox txt_strObservacao 
            Height          =   315
            Left            =   2085
            MaxLength       =   40
            TabIndex        =   188
            Top             =   1395
            Width           =   6915
         End
         Begin MSComctlLib.ListView lvw_ItensPublicidade 
            Height          =   1965
            Left            =   150
            TabIndex        =   193
            Top             =   2220
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3466
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PkidHistPublicidade"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "intTributo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Publicidade"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Quantidade"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Área"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Observação"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Data Início"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Data Fim"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbc_intTributoPublicidade 
            Height          =   315
            Left            =   2100
            TabIndex        =   181
            Top             =   525
            Width           =   6555
            _ExtentX        =   11562
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_dtmPublicidadeInicio 
            AutoSize        =   -1  'True
            Caption         =   "Data Início"
            Height          =   195
            Left            =   1170
            TabIndex        =   189
            Top             =   1920
            Width           =   795
         End
         Begin VB.Label lbl_dtmPublicidadeFim 
            AutoSize        =   -1  'True
            Caption         =   "Data Fim"
            Height          =   195
            Left            =   7170
            TabIndex        =   191
            Top             =   1920
            Width           =   630
         End
         Begin VB.Label lbl_TiposDePublicidade 
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Publicidade"
            Height          =   195
            Left            =   525
            TabIndex        =   180
            Top             =   615
            Width           =   1485
         End
         Begin VB.Label lbl_Quantidade 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            Height          =   195
            Left            =   1185
            TabIndex        =   183
            Top             =   1050
            Width           =   825
         End
         Begin VB.Label lbl_Area 
            AutoSize        =   -1  'True
            Caption         =   "Área"
            Height          =   195
            Left            =   7365
            TabIndex        =   185
            Top             =   1050
            Width           =   330
         End
         Begin VB.Label lbl_Observacao 
            AutoSize        =   -1  'True
            Caption         =   "Observação"
            Height          =   195
            Left            =   1110
            TabIndex        =   187
            Top             =   1470
            Width           =   870
         End
      End
      Begin TabDlg.SSTab tab_3dCaracteristicas 
         Height          =   3945
         Left            =   -74190
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   6959
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "Geral"
         TabPicture(0)   =   "CadEconomico.frx":1AAE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "tdb_caracteristica(4)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "tdb_Detalhe(4)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Física"
         TabPicture(1)   =   "CadEconomico.frx":1ACA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "tdb_Detalhe(5)"
         Tab(1).Control(1)=   "tdb_caracteristica(5)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Jurídica"
         TabPicture(2)   =   "CadEconomico.frx":1AE6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "tdb_caracteristica(6)"
         Tab(2).Control(1)=   "tdb_Detalhe(6)"
         Tab(2).ControlCount=   2
         Begin TrueOleDBGrid70.TDBGrid tdb_Detalhe 
            Height          =   3255
            Index           =   4
            Left            =   4830
            TabIndex        =   174
            Top             =   510
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   5741
            _LayoutType     =   4
            _RowHeight      =   195
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   "PKId"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Detalhes"
            Columns(1).DataField=   "strNomeDaCaracteristica"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   1
            Splits(0).MarqueeStyle=   5
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7805"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7726"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
            Enabled         =   0   'False
            HeadLines       =   1
            FootLines       =   1
            MarqueeUnique   =   0   'False
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=38"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
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
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_caracteristica 
            Height          =   3255
            Index           =   4
            Left            =   180
            TabIndex        =   173
            Top             =   510
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   5741
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
            Columns(1).Caption=   "Característica"
            Columns(1).DataField=   "strNomeDaCaracteristica"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7805"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7726"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
            Enabled         =   0   'False
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=38"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
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
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Detalhe 
            Height          =   3255
            Index           =   5
            Left            =   -70170
            TabIndex        =   176
            Top             =   510
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   5741
            _LayoutType     =   4
            _RowHeight      =   195
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   "PKId"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Detalhes"
            Columns(1).DataField=   "strNomeDaCaracteristica"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   1
            Splits(0).MarqueeStyle=   5
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7805"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7726"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
            MarqueeUnique   =   0   'False
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=38"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
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
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_caracteristica 
            Height          =   3255
            Index           =   5
            Left            =   -74820
            TabIndex        =   175
            Top             =   510
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   5741
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
            Columns(1).Caption=   "Característica"
            Columns(1).DataField=   "strNomeDaCaracteristica"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7805"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7726"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=38"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
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
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Detalhe 
            Height          =   3255
            Index           =   6
            Left            =   -70170
            TabIndex        =   178
            Top             =   510
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   5741
            _LayoutType     =   4
            _RowHeight      =   195
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKId"
            Columns(0).DataField=   "PKId"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Detalhes"
            Columns(1).DataField=   "strNomeDaCaracteristica"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   1
            Splits(0).MarqueeStyle=   5
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7805"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7726"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
            MarqueeUnique   =   0   'False
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=38"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
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
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_caracteristica 
            Height          =   3255
            Index           =   6
            Left            =   -74820
            TabIndex        =   177
            Top             =   510
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   5741
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
            Columns(1).Caption=   "Característica"
            Columns(1).DataField=   "strNomeDaCaracteristica"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7805"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7726"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=144,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=38"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
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
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame fra_Socios 
         Caption         =   " Sócios "
         Height          =   2865
         Left            =   -74790
         TabIndex        =   149
         Top             =   420
         Width           =   11325
         Begin VB.TextBox txt_dtmSocioInicio 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7470
            TabIndex        =   156
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txt_dtmSocioFim 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8670
            TabIndex        =   158
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txt_strCotas 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9870
            TabIndex        =   160
            Top             =   540
            Width           =   1305
         End
         Begin VB.TextBox txt_strCnpjCpf 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   5220
            TabIndex        =   154
            Top             =   540
            Width           =   2145
         End
         Begin VB.CommandButton cmd_Socios 
            Height          =   315
            Left            =   4755
            Picture         =   "CadEconomico.frx":1B02
            Style           =   1  'Graphical
            TabIndex        =   152
            TabStop         =   0   'False
            Tag             =   "584"
            ToolTipText     =   "Ativa Cadastro de Sócios"
            Top             =   540
            Width           =   360
         End
         Begin VB.TextBox txt_TotalDeCotas 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9570
            Locked          =   -1  'True
            TabIndex        =   163
            Top             =   2430
            Width           =   1635
         End
         Begin MSDataListLib.DataCombo dbc_intSocios 
            Height          =   315
            Left            =   105
            TabIndex        =   151
            Top             =   540
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvw_Socios 
            Height          =   1425
            Left            =   90
            TabIndex        =   161
            Top             =   930
            Width           =   11145
            _ExtentX        =   19659
            _ExtentY        =   2514
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
               Text            =   "PkidSocioEconomico"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "intSocio"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Nome do Sócio"
               Object.Width           =   9172
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "CNPJ / CPF"
               Object.Width           =   3882
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Cotas"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Data Início"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Data Fim"
               Object.Width           =   2469
            EndProperty
         End
         Begin VB.Label lbl_dtmSocioInicio 
            AutoSize        =   -1  'True
            Caption         =   "Data Início"
            Height          =   195
            Left            =   7470
            TabIndex        =   155
            Top             =   300
            Width           =   795
         End
         Begin VB.Label lbl_dtmSocioFim 
            AutoSize        =   -1  'True
            Caption         =   "Data Fim"
            Height          =   195
            Left            =   8700
            TabIndex        =   157
            Top             =   300
            Width           =   630
         End
         Begin VB.Label lbl_strCotas 
            AutoSize        =   -1  'True
            Caption         =   "Número de Cotas"
            Height          =   195
            Left            =   9900
            TabIndex        =   159
            Top             =   300
            Width           =   1230
         End
         Begin VB.Label lbl_strCnpjCpf 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ / CPF"
            Height          =   195
            Left            =   5220
            TabIndex        =   153
            Top             =   300
            Width           =   870
         End
         Begin VB.Label lbl_strSocio 
            AutoSize        =   -1  'True
            Caption         =   "Nome do Sócio"
            Height          =   195
            Left            =   120
            TabIndex        =   150
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lbl_Cotas 
            AutoSize        =   -1  'True
            Caption         =   "Total de cotas"
            Height          =   225
            Left            =   8490
            TabIndex        =   162
            Top             =   2490
            Width           =   1020
         End
      End
      Begin VB.Frame fra_Inscricao 
         Height          =   2985
         Left            =   1170
         TabIndex        =   1
         Top             =   390
         Width           =   9345
         Begin VB.Frame fra_Razao 
            Height          =   975
            Left            =   360
            TabIndex        =   11
            Top             =   840
            Width           =   8865
            Begin VB.TextBox txt_PKIdContribuinte 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1170
               Locked          =   -1  'True
               OLEDragMode     =   1  'Automatic
               OLEDropMode     =   1  'Manual
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   180
               Width           =   1095
            End
            Begin VB.CommandButton cmd_Contribuinte 
               Height          =   315
               Left            =   8430
               Picture         =   "CadEconomico.frx":1C20
               Style           =   1  'Graphical
               TabIndex        =   15
               TabStop         =   0   'False
               Tag             =   "584"
               ToolTipText     =   "Ativa Cadastro de Razão Social"
               Top             =   180
               Width           =   360
            End
            Begin VB.TextBox txtdtmRazaoInicio 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   2295
               TabIndex        =   17
               Top             =   555
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbcintContribuinte 
               Height          =   315
               Left            =   2295
               TabIndex        =   14
               Top             =   180
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lblintContribuinte 
               AutoSize        =   -1  'True
               Caption         =   "Razão Social"
               Height          =   195
               Left            =   120
               TabIndex        =   12
               Top             =   300
               Width           =   945
            End
            Begin VB.Label lbldtmRazaoInicio 
               AutoSize        =   -1  'True
               Caption         =   "Data Início"
               Height          =   195
               Left            =   1395
               TabIndex        =   16
               Top             =   645
               Width           =   795
            End
         End
         Begin VB.CheckBox chkbitDefinitivo 
            Caption         =   "Definitivo"
            Height          =   195
            Left            =   8250
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtstrEmissao 
            Height          =   285
            Left            =   3900
            MaxLength       =   3
            TabIndex        =   8
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtPKId 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1545
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txt_strNomeFantasia 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1545
            MaxLength       =   100
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1875
            Width           =   5355
         End
         Begin VB.TextBox txt_strInscricaoEstadual 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1545
            MaxLength       =   100
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2205
            Width           =   2745
         End
         Begin VB.ComboBox cbointAtividadeBasica 
            Height          =   315
            Left            =   1545
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   2550
            Width           =   2745
         End
         Begin VB.CheckBox chkblnMicroEmpresa 
            Caption         =   "Micro-Empresa"
            Height          =   195
            Left            =   5535
            TabIndex        =   22
            Top             =   2280
            Width           =   1395
         End
         Begin VB.TextBox txt_CNPJCPF 
            Height          =   285
            Left            =   5295
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   2580
            Width           =   1605
         End
         Begin VB.Frame fra_NaturezaJuridica 
            Enabled         =   0   'False
            Height          =   945
            Left            =   7050
            TabIndex        =   24
            Top             =   1920
            Width           =   2175
            Begin VB.OptionButton opt_NaturezaJuridica 
               Caption         =   "Jurídica"
               Height          =   195
               Index           =   1
               Left            =   150
               TabIndex        =   27
               Top             =   600
               Width           =   1035
            End
            Begin VB.OptionButton opt_NaturezaJuridica 
               Caption         =   "Física"
               Height          =   195
               Index           =   0
               Left            =   150
               TabIndex        =   25
               Top             =   270
               Width           =   915
            End
            Begin VB.OptionButton opt_NaturezaJuridica 
               Caption         =   "SC"
               Height          =   195
               Index           =   2
               Left            =   1230
               TabIndex        =   26
               Top             =   270
               Width           =   705
            End
            Begin VB.OptionButton opt_NaturezaJuridica 
               Caption         =   "Outros"
               Height          =   195
               Index           =   3
               Left            =   1230
               TabIndex        =   28
               Top             =   600
               Width           =   795
            End
            Begin VB.Label lbl_Natureza 
               AutoSize        =   -1  'True
               Caption         =   " Natureza Jurídica "
               Height          =   195
               Left            =   150
               TabIndex        =   23
               Top             =   0
               Width           =   1350
            End
         End
         Begin MSMask.MaskEdBox mskstrInscricaoCadastral 
            Height          =   285
            Left            =   1545
            TabIndex        =   3
            Top             =   210
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskstrInscricaoImobiliaria 
            Height          =   285
            Left            =   6720
            TabIndex        =   10
            Top             =   540
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin VB.Label lbl_Emissao 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Index           =   1
            Left            =   3270
            TabIndex        =   7
            Top             =   600
            Width           =   585
         End
         Begin VB.Label lblstrInscricaoCadastral 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   90
            TabIndex        =   2
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label lbl_PKId 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   945
            TabIndex        =   5
            Top             =   630
            Width           =   495
         End
         Begin VB.Label lblstrNomeFantasia 
            AutoSize        =   -1  'True
            Caption         =   "Nome Fantasia"
            Height          =   195
            Left            =   375
            TabIndex        =   18
            Top             =   1950
            Width           =   1065
         End
         Begin VB.Label lblstrInscricaoEstadual 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Estadual"
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   2295
            Width           =   1305
         End
         Begin VB.Label lblintAtividade 
            AutoSize        =   -1  'True
            Caption         =   "Atividade Básica"
            Height          =   195
            Left            =   255
            TabIndex        =   29
            Top             =   2670
            Width           =   1185
         End
         Begin VB.Label lbl_CNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ / CPF"
            Height          =   195
            Left            =   4365
            TabIndex        =   31
            Top             =   2670
            Width           =   870
         End
         Begin VB.Label lblstrInscricaoImobiliaria 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Imobiliária"
            Height          =   195
            Left            =   5280
            TabIndex        =   9
            Top             =   600
            Width           =   1380
         End
      End
      Begin VB.Frame fra_Historico 
         Caption         =   " Históricos "
         Height          =   2640
         Left            =   -74100
         TabIndex        =   143
         Top             =   360
         Width           =   9315
         Begin VB.TextBox txt_Historico 
            Height          =   720
            Left            =   150
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   145
            Top             =   555
            Width           =   9060
         End
         Begin MSComctlLib.Toolbar tlb_Historico 
            Height          =   330
            Left            =   60
            TabIndex        =   144
            Top             =   210
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "img_Aux"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Novo"
                  Object.ToolTipText     =   "Novo"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Salvar"
                  Object.ToolTipText     =   "Adicionar / Alterar"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Deletar"
                  Object.ToolTipText     =   "Remover"
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList img_Aux 
            Left            =   1350
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CadEconomico.frx":1D3E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CadEconomico.frx":1E9E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "CadEconomico.frx":1FFA
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Historico 
            Height          =   1170
            Left            =   120
            TabIndex        =   147
            Top             =   1350
            Width           =   9060
            _ExtentX        =   15981
            _ExtentY        =   2064
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame fra_Atividades 
         Caption         =   " Atividades "
         Height          =   3345
         Left            =   -74880
         TabIndex        =   116
         Top             =   360
         Width           =   11535
         Begin VB.TextBox txt_intQtd 
            Height          =   315
            Left            =   10380
            TabIndex        =   128
            Top             =   780
            Width           =   645
         End
         Begin VB.TextBox txt_dtmAtividadeFim 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9960
            TabIndex        =   122
            Top             =   210
            Width           =   1095
         End
         Begin VB.TextBox txt_dtmAtividadeInicio 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8070
            TabIndex        =   120
            Top             =   210
            Width           =   1095
         End
         Begin VB.CheckBox chk_blnPrincipal 
            Caption         =   "Principal"
            Height          =   195
            Left            =   90
            TabIndex        =   118
            Top             =   540
            Width           =   1005
         End
         Begin MSComctlLib.ListView lvw_Atividades 
            Height          =   1080
            Left            =   120
            TabIndex        =   142
            Top             =   1110
            Width           =   11325
            _ExtentX        =   19976
            _ExtentY        =   1905
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
               Text            =   "PkidAtivEmpresa"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PkidAtividade"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "blnPrincipal"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Status"
               Object.Width           =   1773
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Descrição"
               Object.Width           =   11747
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Data Início"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Data Fim"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbc_intAtividadePrincipal 
            Height          =   315
            Left            =   90
            TabIndex        =   117
            Top             =   210
            Width           =   7020
            _ExtentX        =   12383
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intTipoTributo 
            Height          =   315
            HelpContextID   =   1
            Left            =   1350
            TabIndex        =   124
            Top             =   780
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intTributo 
            Height          =   315
            HelpContextID   =   1
            Left            =   6090
            TabIndex        =   126
            Top             =   780
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvw_ItensAtivTrib 
            Height          =   1035
            Left            =   120
            TabIndex        =   129
            Top             =   2220
            Width           =   11325
            _ExtentX        =   19976
            _ExtentY        =   1826
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IntAtivEmpresa"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Atividade"
               Object.Width           =   8114
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "intTributo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Tributo"
               Object.Width           =   8624
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Quantidade"
               Object.Width           =   1852
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbc_intAtividades 
            Height          =   315
            HelpContextID   =   1
            Left            =   10920
            TabIndex        =   211
            Top             =   210
            Visible         =   0   'False
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label lbl_Tributo 
            AutoSize        =   -1  'True
            Caption         =   "Tributos"
            Height          =   195
            Left            =   5490
            TabIndex        =   125
            Top             =   900
            Width           =   570
         End
         Begin VB.Label lbl_TipoTributo 
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Tributos"
            Height          =   195
            Left            =   90
            TabIndex        =   123
            Top             =   900
            Width           =   1230
         End
         Begin VB.Label lbl_intQtd 
            AutoSize        =   -1  'True
            Caption         =   "Qtde"
            Height          =   195
            Left            =   9930
            TabIndex        =   127
            Top             =   900
            Width           =   345
         End
         Begin VB.Label lbl_dtmAtividadeFim 
            AutoSize        =   -1  'True
            Caption         =   "Dt Fim"
            Height          =   195
            Left            =   9420
            TabIndex        =   121
            Top             =   330
            Width           =   450
         End
         Begin VB.Label lbl_dtmAtividadeInicio 
            AutoSize        =   -1  'True
            Caption         =   "Dt Início"
            Height          =   195
            Left            =   7350
            TabIndex        =   119
            Top             =   330
            Width           =   615
         End
      End
      Begin VB.Frame fra_ISS 
         Caption         =   "ISSQN"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   130
         Top             =   3720
         Width           =   11535
         Begin VB.TextBox txt_intQuantidadeIss 
            Height          =   315
            Left            =   11070
            MaxLength       =   2
            TabIndex        =   140
            Top             =   210
            Width           =   345
         End
         Begin VB.TextBox txt_dtmissinicio 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8070
            TabIndex        =   136
            Top             =   210
            Width           =   975
         End
         Begin VB.TextBox txt_dtmissfim 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9630
            TabIndex        =   138
            Top             =   210
            Width           =   975
         End
         Begin MSComctlLib.ListView lvw_ISS 
            Height          =   945
            Left            =   90
            TabIndex        =   141
            Top             =   570
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   1667
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PkidIssEmpresa"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "inttipoiss"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Tipo ISS"
               Object.Width           =   4895
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "intlistaservico"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Lista Serviço"
               Object.Width           =   7408
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Data Início"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Data Fim"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Quantidade"
               Object.Width           =   1411
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbc_intTipoISS 
            Height          =   315
            Left            =   450
            TabIndex        =   132
            Top             =   210
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intListaServico 
            Height          =   315
            Left            =   3480
            TabIndex        =   134
            Top             =   210
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_IssQtde 
            AutoSize        =   -1  'True
            Caption         =   "Qtde"
            Height          =   195
            Left            =   10680
            TabIndex        =   139
            Top             =   330
            Width           =   345
         End
         Begin VB.Label lbl_ISSLista 
            AutoSize        =   -1  'True
            Caption         =   "Lista Serviço"
            Height          =   195
            Left            =   2520
            TabIndex        =   133
            Top             =   330
            Width           =   915
         End
         Begin VB.Label lbl_ISSTipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   90
            TabIndex        =   131
            Top             =   330
            Width           =   315
         End
         Begin VB.Label lbl_ISSDtInicio 
            AutoSize        =   -1  'True
            Caption         =   "Dt Início"
            Height          =   195
            Left            =   7410
            TabIndex        =   135
            Top             =   330
            Width           =   615
         End
         Begin VB.Label lbl_ISSDtFim 
            AutoSize        =   -1  'True
            Caption         =   "Dt Fim"
            Height          =   195
            Left            =   9120
            TabIndex        =   137
            Top             =   330
            Width           =   450
         End
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   " Endereço do Estabelecimento "
         Height          =   1785
         Left            =   1170
         TabIndex        =   33
         Top             =   3420
         Width           =   9345
         Begin VB.TextBox txtdtmEnderecoInicio 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1530
            TabIndex        =   51
            Top             =   1350
            Width           =   1095
         End
         Begin VB.TextBox txt_strMunicipio 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   990
            Width           =   4185
         End
         Begin VB.TextBox txtintNumero 
            Height          =   285
            Left            =   6540
            MaxLength       =   6
            TabIndex        =   38
            Top             =   240
            Width           =   885
         End
         Begin VB.TextBox txtstrComplemento 
            Height          =   285
            Left            =   8340
            MaxLength       =   20
            TabIndex        =   40
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmd_Bairro 
            Height          =   315
            Left            =   5760
            Picture         =   "CadEconomico.frx":2156
            Style           =   1  'Graphical
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro de Bairro"
            Top             =   600
            Width           =   360
         End
         Begin VB.CommandButton cmd_Logradouro 
            Height          =   315
            Left            =   5760
            Picture         =   "CadEconomico.frx":2274
            Style           =   1  'Graphical
            TabIndex        =   36
            TabStop         =   0   'False
            Tag             =   "584"
            ToolTipText     =   "Ativa Cadastro de Logradouro"
            Top             =   210
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintLogradouro 
            Height          =   315
            Left            =   1530
            TabIndex        =   35
            Top             =   210
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.TextBox txt_strUF 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   7050
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   990
            Width           =   375
         End
         Begin VB.TextBox txtintCep 
            Height          =   285
            Left            =   8040
            MaxLength       =   9
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   990
            Width           =   1155
         End
         Begin MSDataListLib.DataCombo dbcintBairro 
            Height          =   315
            Left            =   1530
            TabIndex        =   42
            Top             =   600
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbldtmEnderecoInicio 
            AutoSize        =   -1  'True
            Caption         =   "Data Início"
            Height          =   195
            Left            =   660
            TabIndex        =   50
            Top             =   1440
            Width           =   795
         End
         Begin VB.Label lbl_Complemento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   7770
            TabIndex        =   39
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lblintNumero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   6270
            TabIndex        =   37
            Top             =   330
            Width           =   180
         End
         Begin VB.Label lblstrUF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   6750
            TabIndex        =   46
            Top             =   1080
            Width           =   210
         End
         Begin VB.Label lblintCep 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   7635
            TabIndex        =   48
            Top             =   1080
            Width           =   285
         End
         Begin VB.Label lblstrMunicipio 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   735
            TabIndex        =   44
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label lblintBairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   1035
            TabIndex        =   41
            Top             =   720
            Width           =   405
         End
         Begin VB.Label lblintLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   630
            TabIndex        =   34
            Top             =   330
            Width           =   810
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_HistOcorrencias 
         Height          =   3735
         Left            =   -74850
         TabIndex        =   209
         Top             =   1470
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   6588
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tipo"
         Columns(0).DataField=   "strTipo"
         Columns(0).NumberFormat=   "FormatText Event"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Data Inicial"
         Columns(1).DataField=   "dtmdtInicial"
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data Final"
         Columns(2).DataField=   "dtmdtfinal"
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Ocorrência"
         Columns(3).DataField=   "strOcorrencia"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3704"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3625"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=4180"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4101"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=18256"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=18177"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=0"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin VB.Frame fra_Contador 
         Caption         =   " Contador "
         Height          =   975
         Left            =   -74775
         TabIndex        =   164
         Top             =   3465
         Width           =   11325
         Begin VB.TextBox txt_CRC 
            Height          =   285
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   171
            Top             =   600
            Width           =   1605
         End
         Begin VB.TextBox txt_CNPJCPFContador 
            Height          =   285
            Left            =   1860
            Locked          =   -1  'True
            TabIndex        =   169
            Top             =   600
            Width           =   1965
         End
         Begin VB.TextBox txt_PKIdContador 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   1860
            Locked          =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   225
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dbcintContador 
            Height          =   315
            Left            =   2970
            TabIndex        =   167
            Top             =   225
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_CRC 
            AutoSize        =   -1  'True
            Caption         =   "CRC"
            Height          =   195
            Left            =   8100
            TabIndex        =   170
            Top             =   690
            Width           =   330
         End
         Begin VB.Label lbl_CNPJCPFContador 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ / CPF"
            Height          =   195
            Left            =   960
            TabIndex        =   168
            Top             =   690
            Width           =   870
         End
         Begin VB.Label lbl_Contador 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   1350
            TabIndex        =   165
            Top             =   330
            Width           =   420
         End
      End
   End
End
Attribute VB_Name = "frmCadEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando                As Boolean
    Dim mblnAlterandoH               As Boolean
    Dim mobjAux                      As Object
    Dim mblnClickOk                  As Boolean
    Dim mblnLoading                  As Boolean
    Dim oList                        As Object
    'Dim X                            As New XArrayDB 'Grid Livros de ISS
    
    Dim intCaractEconomico           As Integer
    Dim mblnAlterandoLista           As Boolean 'Verificação para alteração dos itens de feiras
    Dim mblnAlterandoListaTributos   As Boolean 'Verificação para alteração dos itens de Tributos
    Dim mblnAlterandoListaPubli      As Boolean 'Verificação para alteração dos itens de Publicidade
    Dim mblnAlterandoListaSocios     As Boolean 'Verificação para alteração dos itens de Socios
    Dim mblnAlterandoListaAtividade  As Boolean 'Verificação para alteração dos itens de Atividades
    Dim mblnAlterandoListaISS        As Boolean 'Verificação para alteração dos itens de ISSQN
    
    Dim mobjLista                    As Object
    Dim blnAlterando                 As Boolean
    Dim blnPrimeiraVez               As Boolean
    Dim blnCarregaSocio              As Boolean
    
    Dim blnConsultaProcesso          As Boolean
    Dim blnOkAtividade               As Boolean
    Dim blnIssORAtividades           As Boolean 'Faz Verificação em qual grid esta ativo ISS ou Atividades / True = ISS / False = Atividades
    
Private Sub cbo_intTipo_Click()
    If Val(gstrItemData(cbo_intTipo, False)) > 0 And Val(txtPKId) > 0 Then
        PreencheGridHistoricoOcorrencias
    End If
End Sub

Private Sub cbo_intTipo_GotFocus()
    tab_3dPasta.Tab = 9
End Sub

Private Sub cbointAtividadeBasica_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub cbointAtividadeBasica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", cbointAtividadeBasica
End Sub

Private Sub chk_blnPrincipal_Click()
    blnIssORAtividades = False
End Sub

Private Sub cmd_Feira_Click()
    ChamaFormCadastro frmCadFeira, dbc_intFeira
End Sub

Private Sub cmd_Socios_Click()
    ChamaFormCadastro frmCadSocio, dbc_intSocios
End Sub

Private Sub cmd_TipoFeira_Click()
    ChamaFormCadastro frmCadTipoFeira, dbc_intTipoFeira
End Sub

Private Sub cmd_Ocorrencia_Click()
    ChamaFormCadastro frmCadOcorrencia, dbcintOcorrencia
End Sub

Private Sub cmd_TipoDePublicidade_Click()
Dim adoTipo As ADODB.Recordset
Dim strSql As String
    
    strSql = ""
    strSql = "SELECT strDescricao, pkID "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTributoTipo & " "
    strSql = strSql & "WHERE "
    strSql = strSql & "bytTipo = " & TRIBUTO_TIPO_PUBLICIDADE & " "
    strSql = strSql & "ORDER BY strDescricao "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoTipo) Then
       If adoTipo.RecordCount > 0 Then adoTipo.MoveFirst
       ChamaFormCadastro frmCadTributos, dbc_intTributoPublicidade
       If adoTipo.RecordCount > 0 Then frmCadTributos.dbcinttributotipo.Text = adoTipo!strDescricao
       frmCadTributos.dbcinttributotipo.SetFocus
       frmCadTributos.MantemForm (gstrPreencherLista)
       If adoTipo.RecordCount > 0 Then frmCadTributos.dbcinttributotipo.Text = adoTipo!strDescricao
       frmCadTributos.MantemForm (gstrLocalizar)
    Else
       ChamaFormCadastro frmCadTributos, dbc_intTributoPublicidade
    End If
    
End Sub

Private Sub dbc_intAtividadePrincipal_Change()
    Dim strSql As String
    
    If Val(dbc_intAtividadePrincipal.BoundText) > 0 And dbc_intAtividadePrincipal.MatchedWithList Then
        strSql = "SELECT A.Pkid, Rtrim(Ltrim(A.STRDESCRICAO)) as STRDESCRICAO "
        strSql = strSql & "FROM "
        strSql = strSql & gstrAtividadeEC & " A "
        strSql = strSql & "Where "
        strSql = strSql & "A.Pkid in(" & dbc_intAtividadePrincipal.BoundText & ")"
        
        dbc_intAtividades.Tag = strSql & ";strDescricao"
    
        LeDaTabelaParaObj "", dbc_intAtividades, strSql
    End If
End Sub

Private Sub dbc_intAtividadePrincipal_Click(Area As Integer)
    blnIssORAtividades = False
    If Area = 0 Then DropDownDataCombo dbc_intAtividadePrincipal, Me, Area
End Sub

Private Sub dbc_intAtividadePrincipal_GotFocus()
    blnIssORAtividades = False
    MarcaCampo dbc_intAtividadePrincipal
    tab_3dPasta.Tab = 2
End Sub

Private Sub dbc_intAtividadePrincipal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intAtividadePrincipal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intAtividadePrincipal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intAtividadePrincipal
End Sub

Private Sub dbc_intAtividades_Change()
    Dim strSql As String
    
    If dbc_intAtividades.MatchedWithList Then
        strSql = ""
        strSql = strSql & "Select TT.Pkid, TT.strDescricao From " & gstrTributoTipo & " TT, " & gstrAtividadeTributo & " AT Where AT.intAtividadeEc = " & dbc_intAtividades.BoundText & " AND TT.Pkid = AT.intTributoTipo AND bytTipo in (" & TRIBUTO_TIPO_OUTROS & ", " & TRIBUTO_TIPO_HORARIO_ESPECIAL & ") "
        LeDaTabelaParaObj "", dbc_intTipoTributo, strSql
        
        If dbc_intTipoTributo.MatchedWithList Then
            
            strSql = ""
            strSql = strSql & "Select "
            strSql = strSql & " T.PKID, "
            strSql = strSql & " T.STRDESCRICAO "
            strSql = strSql & "From "
            strSql = strSql & gstrAtividadeTributo & " AT, "
            strSql = strSql & gstrAtividadeTributoTributo & " ATT, "
            strSql = strSql & gstrTributo & " T "
            strSql = strSql & "Where "
            strSql = strSql & " AT.INTATIVIDADEEC = " & dbc_intAtividades.BoundText
            strSql = strSql & " AND AT.INTTRIBUTOTIPO = " & dbc_intTipoTributo.BoundText
            strSql = strSql & " AND ATT.INTATIVIDADETRIBUTO = AT.PKID "
            strSql = strSql & " AND T.Pkid = ATT.intTributo "
            strSql = strSql & "ORDER BY T.Pkid "
            
            LeDaTabelaParaObj "", dbc_intTributo, strSql
        Else
            Set dbc_intTributo.RowSource = Nothing
            dbc_intTributo.Text = ""
        
        End If
               
    End If

End Sub

Private Sub dbc_intAtividades_Click(Area As Integer)
Dim strSql       As String
Dim adoResultado As ADODB.Recordset
    
    If Area = 2 Then
        If dbc_intAtividades.MatchedWithList Then
            strSql = ""
            strSql = strSql & "Select TT.Pkid, TT.strDescricao From " & gstrTributoTipo & " TT, " & gstrAtividadeTributo & " AT Where AT.intAtividadeEc = " & dbc_intAtividades.BoundText & " AND TT.Pkid = AT.intTributoTipo AND bytTipo in (" & TRIBUTO_TIPO_OUTROS & ", " & TRIBUTO_TIPO_HORARIO_ESPECIAL & ") "
            LeDaTabelaParaObj "", dbc_intTipoTributo, strSql
            
            If dbc_intTipoTributo.MatchedWithList Then
                
                strSql = ""
                strSql = strSql & "Select "
                strSql = strSql & " T.PKID, "
                strSql = strSql & " T.STRDESCRICAO "
                strSql = strSql & "From "
                strSql = strSql & gstrAtividadeTributo & " AT, "
                strSql = strSql & gstrAtividadeTributoTributo & " ATT, "
                strSql = strSql & gstrTributo & " T "
                strSql = strSql & "Where "
                strSql = strSql & " AT.INTATIVIDADEEC = " & dbc_intAtividades.BoundText
                strSql = strSql & " AND AT.INTTRIBUTOTIPO = " & dbc_intTipoTributo.BoundText
                strSql = strSql & " AND ATT.INTATIVIDADETRIBUTO = AT.PKID "
                strSql = strSql & " AND T.Pkid = ATT.intTributo "
                strSql = strSql & "ORDER BY T.Pkid "
                
                LeDaTabelaParaObj "", dbc_intTributo, strSql
            Else
                txt_intQtd.Text = "1"
            End If
        End If

    End If
End Sub

Private Sub dbc_intAtividades_GotFocus()
    tab_3dPasta.Tab = 2
End Sub

Private Sub dbc_intFeira_GotFocus()
    tab_3dPasta.Tab = 8
    MarcaCampo dbc_intFeira
End Sub

Private Sub dbc_intFeira_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intFeira, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intFeira_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intFeira
End Sub

Private Sub dbc_intListaServico_Click(Area As Integer)
    blnIssORAtividades = True
End Sub

Private Sub dbc_intListaServico_GotFocus()
    blnIssORAtividades = True
    MarcaCampo dbc_intListaServico
End Sub

Private Sub dbc_intListaServico_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intListaServico, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intListaServico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intListaServico
End Sub

Private Sub dbc_intSocios_Change()
    If dbc_intSocios.MatchedWithList Then
        If dbc_intSocios.BoundText > 0 Then
            txt_strCNPJCPF = gstrPreencherCNPJCPF(dbc_intSocios.BoundText)
        End If
    End If
End Sub

Private Sub dbc_intSocios_Click(Area As Integer)
    DropDownDataCombo dbc_intSocios, Me, Area
End Sub

Private Sub dbc_intSocios_GotFocus()
    MarcaCampo dbc_intSocios
    tab_3dPasta.Tab = 5
End Sub

Private Sub dbc_intSocios_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intSocios
End Sub

Private Sub dbc_intTipoFeira_GotFocus()
    MarcaCampo dbc_intTipoFeira
End Sub

Private Sub dbc_intTipoFeira_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intTipoFeira, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intTipoFeira_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intTipoFeira
End Sub

Private Sub dbc_intTipoISS_Click(Area As Integer)
    blnIssORAtividades = True
End Sub

Private Sub dbc_intTipoISS_GotFocus()
    blnIssORAtividades = True
    MarcaCampo dbc_intTipoISS
End Sub

Private Sub dbc_intTipoISS_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intTipoISS, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intTipoISS_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intTipoISS
End Sub

Private Sub dbc_intTipoTributo_Change()
    Dim strSql As String
    
    blnIssORAtividades = False
    If dbc_intTipoTributo.MatchedWithList And dbc_intAtividades.MatchedWithList Then
        strSql = ""
        strSql = strSql & "Select T.PKID, T.STRDESCRICAO From " & gstrAtividadeTributo & " AT, " & gstrAtividadeTributoTributo & " ATT, " & gstrTributo & " T Where AT.INTATIVIDADEEC = " & dbc_intAtividades.BoundText & " and AT.INTTRIBUTOTIPO = " & dbc_intTipoTributo.BoundText & " and ATT.INTATIVIDADETRIBUTO = AT.PKID and T.Pkid = ATT.intTributo"
        LeDaTabelaParaObj "", dbc_intTributo, strSql
    End If

End Sub

Private Sub dbc_intTipoTributo_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intTipoTributo, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intTributo_Click(Area As Integer)
    blnIssORAtividades = False
End Sub

Private Sub dbc_intTributoPublicidade_Click(Area As Integer)
    DropDownDataCombo dbc_intTributoPublicidade, Me, Area
End Sub

Private Sub dbc_intTributoPublicidade_GotFocus()
    tab_3dPasta.Tab = 7
    MarcaCampo dbc_intTributoPublicidade
End Sub

Private Sub dbc_intTributoPublicidade_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intTributoPublicidade, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intTributoPublicidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intTributoPublicidade
End Sub

Private Sub dbc_intTipoTributo_Click(Area As Integer)
    Dim strSql As String
    
    blnIssORAtividades = False
    If Area = 2 Then
        If dbc_intTipoTributo.MatchedWithList And dbc_intAtividades.MatchedWithList Then
            strSql = ""
            strSql = strSql & "Select T.PKID, T.STRDESCRICAO From " & gstrAtividadeTributo & " AT, " & gstrAtividadeTributoTributo & " ATT, " & gstrTributo & " T Where AT.INTATIVIDADEEC = " & dbc_intAtividades.BoundText & " and AT.INTTRIBUTOTIPO = " & dbc_intTipoTributo.BoundText & " and ATT.INTATIVIDADETRIBUTO = AT.PKID and T.Pkid = ATT.intTributo"
            LeDaTabelaParaObj "", dbc_intTributo, strSql
        End If
    End If
End Sub

Private Sub dbc_intTipoTributo_GotFocus()
    blnIssORAtividades = False
    MarcaCampo dbc_intTipoTributo
End Sub

Private Sub dbc_intTipoTributo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intTipoTributo
End Sub

Private Sub dbc_intTributo_GotFocus()
    blnIssORAtividades = False
    MarcaCampo dbc_intTributo
    tab_3dPasta.Tab = 2
End Sub

Private Sub dbc_intTributo_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intTipoTributo, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intTributo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intTributo
End Sub

Private Sub dbcintBairro_Change()
   TrocaCorObjeto txtdtmEnderecoInicio, False
   If Not mblnLoading Then txtdtmEnderecoInicio = ""
End Sub

Private Sub dbcintBairro_Click(Area As Integer)
   DropDownDataCombo dbcintBairro, Me, Area
End Sub

Private Sub dbcintBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintBairro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContador_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintContador, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinte_Change()
    If dbcintContribuinte.MatchedWithList Then
        If dbcintContribuinte.BoundText <> "" Then
            ExibeDadosContribuinte
            TrocaCorObjeto txtdtmRazaoInicio, False
            If Not mblnLoading Then txtdtmRazaoInicio = ""
        End If
    End If
End Sub

Private Sub dbcintContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbcintHorarioFuncionamento_Change()
    txtstrManhaDe.Text = ""
    txtstrManhaAte.Text = ""
    txtstrTardeDe.Text = ""
    txtstrTardeAte.Text = ""
    txtstrNoiteDe.Text = ""
    txtstrNoiteAte.Text = ""
    txtstrMadrugadaDe.Text = ""
    txtstrMadrugadaAte.Text = ""
End Sub

Private Sub dbcintHorarioFuncionamento_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintHorarioFuncionamento, Me, Area
End Sub

Private Sub dbcintHorarioFuncionamento_GotFocus()
    MarcaCampo dbcintHorarioFuncionamento
End Sub

Private Sub dbcintHorarioFuncionamento_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintHorarioFuncionamento, Me, , KeyCode, Shift
End Sub

Private Sub dbcintHorarioFuncionamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintHorarioFuncionamento
End Sub

Private Sub dbcintLogradouro_Change()
    If dbcintLogradouro.MatchedWithList Then
        LogradouroCep Val(dbcintLogradouro.BoundText), dbcintBairro, True, txt_strMunicipio, txt_strUF, txtintCep, False, False
        TrocaCorObjeto txtdtmEnderecoInicio, False
        If Not mblnLoading Then txtdtmEnderecoInicio = ""
    End If
End Sub

Private Sub dbcintLogradouro_Click(Area As Integer)
   DropDownDataCombo dbcintLogradouro, Me, Area
    If dbcintLogradouro.MatchedWithList Then
        LogradouroCep Val(dbcintLogradouro.BoundText), dbcintBairro, True, txt_strMunicipio, txt_strUF, txtintCep, False, False
    End If
End Sub

Private Sub dbcintLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrencia_Click(Area As Integer)
   DropDownDataCombo dbcintOcorrencia, Me, Area
End Sub

Private Sub dbcintOcorrencia_GotFocus()
   tab_3dPasta.Tab = 1
End Sub

Private Sub dbcintOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintOcorrencia
End Sub

Private Sub chkblnMicroEmpresa_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub chkblnMicroEmpresa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", chkblnMicroEmpresa
End Sub

Private Sub cmd_Bairro_Click()
    ChamaFormCadastro frmCadBairro, dbcintBairro
End Sub

Private Sub cmd_Contribuinte_Click()
    ChamaFormCadastro frmCadContribuinte, dbcintContribuinte
End Sub

Private Sub cmd_Logradouro_Click()
    ChamaFormCadastro frmCadLogradouro, dbcintLogradouro, gstrQueryLogradouro
End Sub

Private Sub dbcintBairro_GotFocus()
   tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintBairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintBairro
End Sub

Private Sub dbcintContador_Click(Area As Integer)
   DropDownDataCombo dbcintContador, Me, Area
   If Area = 2 Then
       ExibeDadosContador
   End If
End Sub

Sub ExibeDadosContador()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    txt_CNPJCPFContador = ""
    txt_CRC = ""
    txt_PKIdContador = ""
    
    If dbcintContador.BoundText = "" Then Exit Sub
    
    txt_PKIdContador = dbcintContador.BoundText
    
    strSql = ""
    strSql = strSql & "SELECT CO.strCNPJCPF, CT.strCRC "
    strSql = strSql & "FROM " & gstrContribuinte & " CO, " & gstrContador & " CT "
    strSql = strSql & "WHERE CT.intContribuinte = CO.PKId "
    strSql = strSql & "AND CT.PKId = " & dbcintContador.BoundText
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 4, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                txt_CNPJCPFContador = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!StrCnpjCpf))
                txt_CRC = gstrVerificaCampoNulo(!strCRC)
            End If
        End With
    End If
End Sub

Private Sub dbcintContador_GotFocus()
    tab_3dPasta.Tab = 5
End Sub

Private Sub dbcintContador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContador
End Sub

Private Sub dbcintContribuinte_Click(Area As Integer)
   DropDownDataCombo dbcintContribuinte, Me, Area
   If Area = 2 Then
       ExibeDadosContribuinte
   End If
End Sub

Private Sub dbcintContribuinte_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintContribuinte_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            dbcintContribuinte_Click 2
    End Select
    CaracterValido KeyAscii, "A", dbcintContribuinte
End Sub

Private Sub dbcintLogradouro_GotFocus()
   tab_3dPasta.Tab = 0
End Sub

Private Sub dbcintLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintLogradouro
End Sub

Private Sub dbcintocorrenciadoeconomico_GotFocus()
    MarcaCampo dbcintocorrenciadoeconomico
End Sub

Private Sub dbcintocorrenciadoeconomico_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintocorrenciadoeconomico, Me, , KeyCode, Shift
End Sub

Private Sub dbcintocorrenciadoeconomico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintLogradouro
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = 737
    
    If mblnAlterando Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    
    If tab_3dPasta.Tab = 8 Or tab_3dPasta.Tab = 2 Or tab_3dPasta.Tab = 7 Or tab_3dPasta.Tab = 5 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
    
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
    If tab_3dPasta.Tab = 0 Then
        VirificaGradeListView Me
        mskstrInscricaoCadastral.SetFocus
    End If
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrLocalizar
        
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
End Sub

Private Sub Form_Load()
Dim adoParametros As New ADODB.Recordset

    mblnAlterandoLista = False
    mblnAlterandoListaTributos = False
    mblnAlterandoListaPubli = False
    mblnAlterandoListaSocios = False
    mblnAlterandoListaAtividade = False


    blnCarregaSocio = False
        
    tab_3dPasta.TabVisible(3) = False
    tab_3dPasta.TabVisible(6) = False
    tab_3dPasta.TabEnabled(3) = False
    tab_3dPasta.TabEnabled(6) = False
    
    TrocaCorObjeto txt_strInscricaoEstadual, True
    TrocaCorObjeto txt_strNomeFantasia, True
    TrocaCorObjeto txt_CNPJCPF, True
    TrocaCorObjeto txt_PKIdContribuinte, True
    TrocaCorObjeto txtPKId, True
    TrocaCorObjeto txt_PKIdContador, True
    TrocaCorObjeto txt_CRC, True
    TrocaCorObjeto txt_CNPJCPFContador, True
    TrocaCorObjeto txt_strMunicipio, True
    TrocaCorObjeto txt_strUF, True
    TrocaCorObjeto txt_TotalDeCotas, True
    TrocaCorObjeto txt_strCNPJCPF, True
    
    MontaColumnHeaders
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT bytControleProcesso FROM " & gstrParametrosTributario, 5, adoParametros) Then
        blnConsultaProcesso = adoParametros("bytControleProcesso").Value = 1
    End If
    
    LeDaTabelaParaObj gstrAtividadeBasica, cbointAtividadeBasica
    
    dbc_intTipoISS.Tag = strQueryTipoIss & ";strdescricao"
    dbc_intListaServico.Tag = strQueryListaServico & ";strdescricao"
    dbcintContribuinte.Tag = strQueryDataComboContribuinte & ";strNome"
    dbcintContador.Tag = strQueryContador & ";strNome"
    dbcintOcorrencia.Tag = strQueryOcorrencia & ";strDescricao"
    dbc_intSocios.Tag = strQuerySocios & ";strNome"
    
    dbcintLogradouro.Tag = gstrQueryLogradouro(, , , True) & ";L.strDescricao"
    dbcintBairro.Tag = gstrQueryDataComboBairro & ";strDescricao"
    
    VerificaMascaraInscricao
        
    PreencheComboHistorico
        
    VerificaObjParaAplicar mobjAux
    txt_strMunicipio = gstrCidadeEmpresa
    txt_strUF = gstrUFEmpresa
    
    dbcintHorarioFuncionamento.Tag = strQueryHorarioFuncionamento & ";strDescricao"
    
    dbc_intFeira.Tag = strQueryFeira & ";strDescricao"
    dbc_intTipoFeira.Tag = strQueryTipoFeira & ";strDescricao"
    dbc_intTributoPublicidade.Tag = strQueryComboPublicidade & ";T.strDescricao"
    dbcintocorrenciadoeconomico.Tag = strQueryOcorrenciaProcesso & ";strDescricao"
    dbc_intAtividadePrincipal.Tag = strQueryAtividade & ";AEC.strDescricao;AEC.intCodigo"

End Sub

Private Function strQueryListaServico() As String
    Dim strSql As String
     
    strSql = ""
    strSql = "SELECT Pkid,"
    
    strSql = strSql
    
    If bytDBType = EDatabases.SQLServer Then
       strSql = strSql & "REPLICATE('0',5 - " & strLen & "(strCodigo))" & strCONCAT & "RTRIM(LTRIM(strCodigo)) " & _
                         strCONCAT & "' - '" & strCONCAT & _
                         " RTRIM(LTRIM(strDescricao)) strDescricao "
    Else
       strSql = strSql & "RTRIM(LTRIM( " & gstrCONVERT(CDT_VARCHAR, "strCodigo,'00000'") & ")) " & _
                         strCONCAT & "' - '" & strCONCAT & _
                         " RTRIM(LTRIM(strDescricao)) strDescricao "
    End If
    
    strSql = strSql & "FROM "
    strSql = strSql & gstrListaServico & " "
    
    strSql = strSql & "ORDER BY " & gstrCONVERT(CDT_INT, "strCodigo")
    
    strQueryListaServico = strSql
End Function


Private Function strQueryDataComboContribuinte() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strNome "
    strSql = strSql & "FROM " & gstrContribuinte & " "
    strSql = strSql & "ORDER BY strNome"
    strQueryDataComboContribuinte = strSql
End Function

Function strQuery() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT EC.PKId, CO.strNome, " & gstrRIGHT("EC.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricaoCadastral "
    strSql = strSql & " FROM " & gstrEconomico & " EC,"
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE EC.intContribuinte = CO.PKId "
    strSql = strSql & " ORDER BY EC.strInscricaoCadastral "
    
    strQuery = strSql
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    gobjBanco.ExecutaRollbackTrans
    Set gobjBanco = Nothing
End Sub

Function strQueryOcorrencia() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrOcorrencia & " "
    strSql = strSql & "WHERE intUtilizacaoDaOcorrencia = 5 "
    strSql = strSql & "ORDER BY strDescricao"
    strQueryOcorrencia = strSql
End Function

Function strQueryContador() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT CT.PKId, CO.strNome AS Contador "
    strSql = strSql & "FROM " & gstrContador & " CT, " & gstrContribuinte & " CO "
    strSql = strSql & "WHERE CT.intContribuinte = CO.PKId "
    strSql = strSql & "ORDER BY CO.strNome"
    strQueryContador = strSql
End Function

Sub VerificaMascaraInscricao()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
    
    strMascara = ""
    
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & TYP_ECONOMICA
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
    mskstrInscricaoCadastral.Mask = strMascara
    
    'Inscrição Imobiliaria
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
    mskstrInscricaoImobiliaria.Mask = strMascara
End Sub

Private Sub fra_Atividades_Click()
    blnIssORAtividades = False
End Sub

Private Sub fra_ISS_Click()
    blnIssORAtividades = True
End Sub

Private Sub lbl_dtmAtividadeFim_Click()
    blnIssORAtividades = False
End Sub

Private Sub lbl_dtmAtividadeInicio_Click()
    blnIssORAtividades = False
End Sub

Private Sub lbl_intQtd_Click()
        blnIssORAtividades = False
End Sub

Private Sub lbl_ISSDtFim_Click()
    blnIssORAtividades = True
End Sub

Private Sub lbl_ISSDtInicio_Click()
    blnIssORAtividades = True
End Sub

Private Sub lbl_ISSLista_Click()
    blnIssORAtividades = True
End Sub

Private Sub lbl_IssQtde_Click()
    blnIssORAtividades = True
End Sub

Private Sub lbl_ISSTipo_Click()
    blnIssORAtividades = True
End Sub

Private Sub lbl_TipoTributo_Click()
    blnIssORAtividades = False
End Sub

Private Sub lbl_Tributo_Click()
    blnIssORAtividades = False
End Sub

Private Sub lvw_Atividades_Click()
    blnIssORAtividades = False
    With lvw_Atividades
        If .ListItems.Count > 0 Then
            txt_intQtd = "1"
            PreencherListaDeOpcoes dbc_intAtividadePrincipal, .SelectedItem.SubItems(1)
            chk_blnPrincipal.Value = Abs(CInt(CBool(.SelectedItem.SubItems(2))))
            txt_dtmAtividadeInicio.Text = .SelectedItem.SubItems(5)
            txt_dtmAtividadeFim.Text = .SelectedItem.SubItems(6)
            mblnAlterandoListaAtividade = True
            
            PreencherListaDeOpcoes dbc_intAtividades, .SelectedItem.SubItems(1)
            
        End If
    End With
End Sub

Private Sub lvw_Atividades_GotFocus()
    blnIssORAtividades = False
End Sub

Private Sub lvw_Atividades_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnAlterandoListaAtividade = True
End Sub

Private Sub lvw_Historico_GotFocus()
    tab_3dPasta.Tab = 4
End Sub

Private Sub lvw_Historico_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txt_Historico = lvw_Historico.SelectedItem.Text
    mblnAlterandoH = True
End Sub

Private Sub lvw_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", lvw_Historico
End Sub

Private Sub lvw_ISS_Click()
    blnIssORAtividades = True
    
    With lvw_ISS
        If .ListItems.Count > 0 Then
            PreencherListaDeOpcoes dbc_intTipoISS, .SelectedItem.SubItems(1)
            PreencherListaDeOpcoes dbc_intListaServico, .SelectedItem.SubItems(3)
            txt_dtmissinicio.Text = gstrDataFormatada(.SelectedItem.SubItems(5))
            txt_dtmissfim.Text = gstrDataFormatada(.SelectedItem.SubItems(6))
            txt_intQuantidadeIss.Text = .SelectedItem.SubItems(7)
            mblnAlterandoListaISS = True
        End If
    End With
    
End Sub

Private Sub lvw_ISS_GotFocus()
    blnIssORAtividades = True
End Sub

Private Sub lvw_Itens_Click()

    With lvw_Itens
        If .ListItems.Count > 0 Then
            PreencherListaDeOpcoes dbc_intFeira, .SelectedItem.Text
            PreencherListaDeOpcoes dbc_intTipoFeira, .SelectedItem.SubItems(2)
            txt_areaFeira.Text = .SelectedItem.SubItems(4)
            txt_strnrbox.Text = .SelectedItem.SubItems(5)
            mblnAlterandoLista = True
        End If
    End With
End Sub

Private Sub lvw_Itens_GotFocus()
    tab_3dPasta.Tab = 8
End Sub

Private Sub lvw_ItensAtivTrib_Click()
    blnIssORAtividades = False
End Sub

Private Sub lvw_ItensAtivTrib_GotFocus()
    blnIssORAtividades = False
    tab_3dPasta.Tab = 2
End Sub

Private Sub lvw_ItensAtivTrib_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strSql As String
    
    mblnAlterandoListaTributos = True
    
    If Val(lvw_ItensAtivTrib.SelectedItem.Text) > 0 Then
        'Preenche a quantidade
        txt_intQtd = Val(lvw_ItensAtivTrib.SelectedItem.SubItems(4))
        'Preenche a atividade
        strSql = "SELECT A.Pkid, Rtrim(Ltrim(A.STRDESCRICAO)) as STRDESCRICAO "
        strSql = strSql & "FROM "
        strSql = strSql & gstrAtividadeEC & " A "
        strSql = strSql & "Where "
        strSql = strSql & "A.Pkid in(" & Val(lvw_ItensAtivTrib.SelectedItem.Text) & ")"
        
        dbc_intAtividades.Tag = strSql & ";strDescricao"
        PreencherListaDeOpcoes dbc_intAtividadePrincipal, Val(lvw_ItensAtivTrib.SelectedItem.Text)
        LeDaTabelaParaObj "", dbc_intAtividades, strSql
        
        If dbc_intAtividades.MatchedWithList Then
            'Preenche o Tipo De Tributo
            strSql = "Select TT.Pkid, TT.strDescricao From " & gstrTributoTipo & " TT, " & gstrAtividadeTributo & " AT Where AT.intAtividadeEc = " & dbc_intAtividades.BoundText & " AND TT.Pkid = AT.intTributoTipo AND bytTipo in (" & TRIBUTO_TIPO_OUTROS & ", " & TRIBUTO_TIPO_HORARIO_ESPECIAL & ") "
            LeDaTabelaParaObj "", dbc_intTipoTributo, strSql
        
            If dbc_intTipoTributo.MatchedWithList Then
                'Preenche o Tributo
                strSql = ""
                strSql = strSql & "Select "
                strSql = strSql & " T.PKID, "
                strSql = strSql & " T.STRDESCRICAO "
                strSql = strSql & "From "
                strSql = strSql & gstrTributo & " T "
                strSql = strSql & "Where "
                strSql = strSql & "T.pkid = " & Val(lvw_ItensAtivTrib.SelectedItem.SubItems(2))
                strSql = strSql & " ORDER BY T.Pkid "
                
                LeDaTabelaParaObj "", dbc_intTributo, strSql
            
            End If
        End If
    End If
End Sub

Private Sub lvw_ItensPublicidade_Click()
    With lvw_ItensPublicidade
        If .ListItems.Count > 0 Then
            PreencherListaDeOpcoes dbc_intTributoPublicidade, .SelectedItem.SubItems(1)
            txt_intQuantidade.Text = .SelectedItem.SubItems(3)
            txt_dblArea.Text = .SelectedItem.SubItems(4)
            txt_strObservacao.Text = .SelectedItem.SubItems(5)
            txt_dtmPublicidadeInicio.Text = .SelectedItem.SubItems(6)
            txt_dtmPublicidadeFim.Text = .SelectedItem.SubItems(7)
            mblnAlterandoListaPubli = True
        End If
    End With
    
End Sub

Private Sub lvw_Socios_Click()
    With lvw_Socios
        If .ListItems.Count > 0 Then
            PreencherListaDeOpcoes dbc_intSocios, .SelectedItem.SubItems(1)
            txt_strCNPJCPF.Text = .SelectedItem.SubItems(3)
            txt_strCotas.Text = .SelectedItem.SubItems(4)
            txt_dtmSocioInicio.Text = .SelectedItem.SubItems(5)
            txt_dtmSocioFim.Text = .SelectedItem.SubItems(6)
            mblnAlterandoListaSocios = True
        End If
    End With
End Sub

Private Sub mskstrInscricaoCadastral_GotFocus()
    MarcaCampo mskstrInscricaoCadastral
End Sub

Private Sub mskstrInscricaoCadastral_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricaoCadastral
End Sub

Private Sub mskstrInscricaoImobiliaria_GotFocus()
    tab_3dPasta.Tab = 0
    MarcaCampo mskstrInscricaoImobiliaria
End Sub

Private Sub mskstrInscricaoImobiliaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricaoImobiliaria
End Sub

Private Sub mskstrInscricaoImobiliaria_LostFocus()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
          
    strSql = ""
    strSql = strSql & "Select intlogradouro, intbairro, intnumero, strcomplemento, intcep "
    strSql = strSql & "FROM " & gstrImobiliario & " where strinscricao = '" & String(gintLenInscricao - Len(mskstrInscricaoImobiliaria.Text), "0") & mskstrInscricaoImobiliaria.Text & "'"
       
    Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 30, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    PreencherListaDeOpcoes dbcintLogradouro, gstrVerificaCampoNulo(adoResultado!intLogradouro)
                    PreencherListaDeOpcoes dbcintBairro, gstrVerificaCampoNulo(adoResultado!intBairro)
                    txtintNumero.Text = gstrVerificaCampoNulo(adoResultado!INTNUMERO)
                    txtstrComplemento.Text = gstrVerificaCampoNulo(adoResultado!STRCOMPLEMENTO)
                    PreencherListaDeOpcoes txtintCep, gstrVerificaCampoNulo(adoResultado!INTCEP)
                    .MoveNext
                Loop
            End With
        End If
        
        blnInscricaoImobiliaria
End Sub

Private Sub opt_NaturezaJuridica_GotFocus(Index As Integer)
    tab_3dPasta.Tab = 0
End Sub

Private Sub opt_NaturezaJuridica_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", opt_NaturezaJuridica(Index)
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3dPasta.Tab = 8 Or tab_3dPasta.Tab = 2 Or tab_3dPasta.Tab = 7 Or tab_3dPasta.Tab = 5 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
End Sub

Private Sub tdb_Detalhe_GotFocus(Index As Integer)
    If tab_3dCaracteristicas.TabEnabled(Index - 4) Then
        tab_3dCaracteristicas.Tab = Index - 4
    End If
End Sub

Private Sub tdb_Detalhe_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Detalhe(Index)
End Sub

Private Sub tdb_HistOcorrencias_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 1 Or ColIndex = 2 Then
        Value = gstrDataFormatada(Value)
    End If
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Value = gstrFormataInscricao(CStr(Value), TYP_ECONOMICA)
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
   gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    If tdb_Lista.Col = 2 Then
        CaracterValido KeyAscii, "N", tdb_Lista
    Else
        CaracterValido KeyAscii, "A", tdb_Lista
    End If
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Trim(txtPKId) <> "" Then
        Set gobjBanco = New clsBanco
        'gobjBanco.ExecutaRollbackTrans
    End If
    
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            Screen.MousePointer = vbHourglass
            'gobjBanco.ExecutaBeginTrans
            blnPrimeiraVez = False
            mblnAlterando = True
            mblnClickOk = False
            
            mblnLoading = True
            
            txtPKId = .Columns("PKID").Value

            LeDaTabelaParaObj gstrEconomico, Me, strQueryEconomico
            PreencheProcessos
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            dbcintContribuinte_Click 2
            dbcintContador_Click 2
            LimpaTabISS
            PreencheGrdAtividades txtPKId
            PreencheGrdISS txtPKId
            LeDaTabelaParaObj "", tdb_Processo, strQueryHistoricoProcessoGrid
            CarregaHistoricos txtPKId
            PreencheListaAtividadeTributo txtPKId
            dbc_intAtividades_Click 2
            tab_3dCaracteristicas.Tab = 0
            PreencheGrdPublicidade txtPKId
            PreencheListItens
            Set tdb_HistOcorrencias.DataSource = Nothing
            cbo_intTipo.ListIndex = -1
            PreencheGrdSocios txtPKId
            TrocaCorObjeto txtdtmRazaoInicio, True
            TrocaCorObjeto txtdtmEnderecoInicio, True
            Screen.MousePointer = vbDefault
        End If
    End With
    
    mblnLoading = False
    
End Sub

Private Sub tdb_Processo_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 0 Then
        Value = gstrDataFormatada(Value)
    End If
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark         As Variant
    Dim intPKIdEconomico    As Long
    Dim strDtEncerramento   As String
    Dim strSql              As String
    
    intPKIdEconomico = Val(txtPKId)
    Screen.MousePointer = vbHourglass
    
    Select Case UCase(strModoOperacao)
        Case UCase(gstrNovo)
            If tab_3dPasta.Tab = 8 Then
                LimpaTabFeira
            ElseIf tab_3dPasta.Tab = 2 Then
                If blnIssORAtividades Then
                    LimpaISS
                    dbc_intTipoISS.SetFocus
                Else
                    LimpaTabAtividade False
                    LimpaTabTributo
                    dbc_intAtividadePrincipal.SetFocus
                End If
            ElseIf tab_3dPasta.Tab = 5 Then
                LimpaTabSocios False
                dbc_intSocios.SetFocus
            ElseIf tab_3dPasta.Tab <> 7 Then
                LimpaObjeto Me, mblnAlterando
                LimpaTabTributo
                NovoEconomico
                Set tdb_HistOcorrencias.DataSource = Nothing
                cbo_intTipo.ListIndex = -1
                mskstrInscricaoCadastral.SetFocus
                gobjBanco.ExecutaRollbackTrans
            Else
                LimpaTabPubli False
            End If
        Case UCase(gstrSalvar)
            If blnDadosOk Then
                
                blnAlterando = mblnAlterando
                
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaBeginTrans
                
                If Trim(txtdtmDataEncerramento) <> "" Then
                    strDtEncerramento = txtdtmDataEncerramento
                    txtdtmRazaoInicio = txtdtmDataEncerramento
                    txtdtmEnderecoInicio = txtdtmDataEncerramento
                End If
                
                
                If ToolBarGeral(strModoOperacao, gstrEconomico, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery) Then
                    If Not blnAlterando Then
                        intPKIdEconomico = glngPegaUltimaChave(gstrEconomico, "PKId")
                    End If
                    If blnSalvaISS(intPKIdEconomico) Then
                        If blnGravaHistoricos(intPKIdEconomico) Then
                            If blnGravaSocios(intPKIdEconomico) Then
                                If StrSalvaItem(intPKIdEconomico) Then
                                    If GravaPublicidades(intPKIdEconomico) Then
                                        If blnGravaAtividades(intPKIdEconomico) Then
                                            If lvw_ItensAtivTrib.ListItems.Count >= 0 Then
                                                If SalvaItensAtividadesTributo(CLng(intPKIdEconomico), blnAlterando) Then
                                                    If Trim(strDtEncerramento) <> "" Then
                                                        If blnEncerraEmpresa(intPKIdEconomico, strDtEncerramento) Then
                                                            gobjBanco.ExecutaCommitTrans
                                                        Else
                                                            gobjBanco.ExecutaRollbackTrans
                                                        End If
                                                    Else
                                                        gobjBanco.ExecutaCommitTrans
                                                    End If
                                                Else
                                                    gobjBanco.ExecutaRollbackTrans
                                                End If
                                            Else
                                                gobjBanco.ExecutaCommitTrans
                                            End If
                                        Else
                                            gobjBanco.ExecutaRollbackTrans
                                        End If
                                    Else
                                        gobjBanco.ExecutaRollbackTrans
                                    End If
                                Else
                                    gobjBanco.ExecutaRollbackTrans
                                End If
                            Else
                                gobjBanco.ExecutaRollbackTrans
                            End If
                        Else
                            gobjBanco.ExecutaRollbackTrans
                        End If
                    Else
                        gobjBanco.ExecutaRollbackTrans
                    End If
                    NovoEconomico
                    Set tdb_HistOcorrencias.DataSource = Nothing
                    tab_3dPasta.Tab = 0
                    mskstrInscricaoCadastral.SetFocus
                Else
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaRollbackTrans
                    tab_3dPasta.Tab = 0
                    mskstrInscricaoCadastral.SetFocus
                End If
                blnAlterando = mblnAlterando
                
            End If
        Case UCase(gstrImprimir)
            
        Case UCase(gstrDeletar)
            If mblnAlterando Then
                If Not TemLancamento Then
                    If blnDeletaEconomico Then
                        LimpaObjeto Me, mblnAlterando
                        NovoEconomico
                    End If
                End If
            End If
        Case UCase(gstrLocalizar)
            mblnClickOk = True
            LocalizarEconomico
        Case UCase(gstrPreencherLista)
            If Me.ActiveControl.Name = "dbcintListaServico" Then
               LeDaTabelaParaObj "", Me.ActiveControl, strQueryListaServico
            ElseIf Me.ActiveControl.Name = "dbcintLogradouro" Then
                dbcintLogradouro.Tag = gstrQueryLogradouro(, , , False) & ";L.strDescricao"
                PreencherListaDeOpcoes Me.ActiveControl
                dbcintLogradouro.Tag = gstrQueryLogradouro(, , , True) & ";L.strDescricao"
            ElseIf Me.ActiveControl.Name = "dbc_intTributo" Then
                If dbc_intAtividades.MatchedWithList Then
                    If dbc_intTipoTributo.MatchedWithList Then
                        strSql = strSql & "Select T.PKID, T.STRDESCRICAO From " & gstrAtividadeTributo & " AT, " & gstrAtividadeTributoTributo & " ATT, " & gstrTributo & " T Where AT.INTATIVIDADEEC = " & dbc_intAtividades.BoundText & " and AT.INTTRIBUTOTIPO = " & dbc_intTipoTributo.BoundText & " and ATT.INTATIVIDADETRIBUTO = AT.PKID and T.Pkid = ATT.intTributo and Upper(T.strdescricao) like '" & UCase(dbc_intTributo.Text) & "%'"
                        LeDaTabelaParaObj "", dbc_intTributo, strSql
                    End If
                End If
            ElseIf Me.ActiveControl.Name = "dbc_intTipoTributo" Then
                If dbc_intAtividades.MatchedWithList Then
                    strSql = strSql & "Select TT.Pkid, TT.strDescricao From " & gstrTributoTipo & " TT, " & gstrAtividadeTributo & " AT Where AT.intAtividadeEc = " & dbc_intAtividades.BoundText & " AND TT.Pkid = AT.intTributoTipo AND bytTipo in (" & TRIBUTO_TIPO_OUTROS & ", " & TRIBUTO_TIPO_HORARIO_ESPECIAL & ") and Upper(TT.strDescricao) like '" & UCase(dbc_intTipoTributo.Text) & "%'"
                    LeDaTabelaParaObj "", dbc_intTipoTributo, strSql
                End If
            Else
               PreencherListaDeOpcoes Me.ActiveControl
            End If
        Case UCase(gstrFechar)
            Unload Me
        Case UCase(gstrRefresh)
            ToolBarGeral strModoOperacao, gstrEconomico, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery
        Case UCase(gstrIncluirItem)
            If tab_3dPasta.Tab = 2 Then 'Atividades
                If blnIssORAtividades Then
                    IncluirItemNoGrid 6
                Else
                    If Not mblnAlterandoListaTributos Then
                        IncluirItemNoGrid 5
                        If blnOkAtividade Then IncluirItemNoGrid 2
                    Else
                        IncluirItemNoGrid 2
                    End If
                End If
            ElseIf tab_3dPasta.Tab = 8 Then 'Feiras
                IncluirItemNoGrid 1
            ElseIf tab_3dPasta.Tab = 7 Then 'Publicidades
                IncluirItemNoGrid 3
            ElseIf tab_3dPasta.Tab = 5 Then 'Socios
                IncluirItemNoGrid 4
            End If
        Case UCase(gstrExcluirItem)
            If tab_3dPasta.Tab = 2 Then 'Atividades
                If blnIssORAtividades Then
                    ExcluirItemNoGrid 6
                Else
                    If Not mblnAlterandoListaTributos Then
                        ExcluirItemNoGrid 5
                    Else
                        ExcluirItemNoGrid 2
                    End If
                    LimpaTabTributo
                End If
            ElseIf tab_3dPasta.Tab = 7 Then 'Publicidades
                ExcluirItemNoGrid 3
            ElseIf tab_3dPasta.Tab = 8 Then 'Feiras
                ExcluirItemNoGrid 1
            ElseIf tab_3dPasta.Tab = 5 Then 'Socios
                ExcluirItemNoGrid 4
            End If
    End Select
    Screen.MousePointer = vbDefault

End Sub

Private Function TemLancamento() As Boolean
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = "SELECT PKID "
    strSql = strSql & "FROM " & gstrLancamentoCalculo
    strSql = strSql & " WHERE intContribuinte = " & dbcintContribuinte.BoundText
    strSql = strSql & " AND strInscricaoCadastral = '" & tdb_Lista.Columns("strInscricaoCadastral")
    strSql = strSql & "' AND bytOrigem = 2 "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
               ExibeMensagem " Para Inscrição Cadastral " & tdb_Lista.Columns("strInscricaoCadastral") & " existem lançamentos pendentes que impedem a exclusão "
               TemLancamento = True
            End If
        End With
    End If
End Function

Private Function blnDeletaEconomico() As Boolean
    Dim strSql As String

    If MsgBox("Confirma exclusão do registro de '" & dbcintContribuinte.Text & "' ?", vbQuestion + vbYesNo) = vbYes Then

        strSql = IIf(bytDBType = Oracle, "Begin ", "")

        strSql = strSql & "DELETE FROM " & gstrHistoricoPublicidades
        strSql = strSql & " WHERE intEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", " ")

        strSql = strSql & "Delete From " & gstrCaracteristicaDoEconomico & " "
        strSql = strSql & "Where Intcodigoeconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
        
        strSql = strSql & "Delete From " & gstrEconomicoFeira & " "
        strSql = strSql & "Where intEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")

        strSql = strSql & "Delete "
        strSql = strSql & "From "
        strSql = strSql & gstrAtivEmpresaTributo & " "
        strSql = strSql & "Where "
        strSql = strSql & "Pkid in(Select AET.pkid From "
        strSql = strSql & gstrAtivEmpresaTributo & " AET, "
        strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
        strSql = strSql & gstrEconomico & " E "
        strSql = strSql & "Where E.Pkid = AE.intEconomico And AE.Pkid = AET.INTATIVIDADEDAEMPRESA And "
        strSql = strSql & "E.Pkid = " & txtPKId & ")"
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        
        strSql = strSql & "Delete From " & gstrAtividadeDaEmpresa & " "
        strSql = strSql & "Where intEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")

        strSql = strSql & "DELETE FROM " & gstrHistoricoEconomico & " "
        strSql = strSql & "WHERE intEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")

        strSql = strSql & "DELETE FROM " & gstrHistoricoPublicidades
        strSql = strSql & " WHERE intEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")

        strSql = strSql & "Delete From " & gstrLivrosDeISS & " "
        strSql = strSql & "Where intEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")

        strSql = strSql & "Delete From " & gstrSocioEconomico & " "
        strSql = strSql & "Where intCodEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        
        strSql = strSql & "Delete From " & gstrHistoricoEconVariavel & " "
        strSql = strSql & "Where intEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")

        strSql = strSql & "Delete From " & gstrProcessoEconomico & " "
        strSql = strSql & "Where intEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        
        strSql = strSql & "Delete From " & gstrIssEmpresa
        strSql = strSql & " Where intEconomico = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")

        strSql = strSql & "Delete From " & gstrEconomico & " "
        strSql = strSql & "Where PKId = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
                    
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSql) Then
            gobjBanco.ExecutaCommitTrans
            blnDeletaEconomico = True
            VerificaListaAutomatica gstrEconomico, tdb_Lista, strQuery
        Else
            gobjBanco.ExecutaRollbackTrans
            blnDeletaEconomico = False
            ExibeMensagem "Ocorreu um erro ao excluir o registro."
        End If
    End If
    
End Function

Sub ExibeDadosContribuinte()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    Dim adoEndereco  As ADODB.Recordset
    
    txt_CNPJCPF = ""
    txt_PKIdContribuinte = ""
    opt_NaturezaJuridica(0).Value = False
    opt_NaturezaJuridica(1).Value = False
    opt_NaturezaJuridica(2).Value = False
    opt_NaturezaJuridica(3).Value = False
    
    If dbcintContribuinte.BoundText = "" Then Exit Sub
    
    txt_PKIdContribuinte = dbcintContribuinte.BoundText
    
    strSql = ""
    strSql = strSql & "SELECT CO.bytNaturezaJuridica, CO.strInscricaoEstadual, CO.strCNPJCPF, CO.intNumero, "
    strSql = strSql & "CO.strNomeFantasia Fantasia, CO.strComplemento, CO.intCEP, CO.intLogradouro, "
    strSql = strSql & "BA.strDescricao AS Bairro, "
    strSql = strSql & "CI.strDescricao AS Municipio, UF.strSigla AS UF "
    strSql = strSql & "FROM " & gstrContribuinte & " CO , " & gstrCidade & " CI, "
    strSql = strSql & gstrBairro & " BA, " & gstrUF & " UF "
    strSql = strSql & "Where CO.intMunicipio " & strOUTJSQLServer & "= CI.PKId " & strOUTJOracle
    strSql = strSql & "AND CO.intUF " & strOUTJSQLServer & "= UF.PKId " & strOUTJOracle
    strSql = strSql & "AND CO.intBairro " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle
    
    strSql = strSql & "AND CO.PKId = " & dbcintContribuinte.BoundText
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 4, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                txt_strNomeFantasia = gstrVerificaCampoNulo(!Fantasia)
                txt_strInscricaoEstadual = gstrVerificaCampoNulo(!strInscricaoEstadual)
                txt_CNPJCPF = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!StrCnpjCpf))
                opt_NaturezaJuridica(!bytNaturezaJuridica).Value = True
            End If
        End With
    End If
    
    tab_3dCaracteristicas.Tab = 0
    If opt_NaturezaJuridica(0).Value = True Then
        tab_3dCaracteristicas.TabEnabled(1) = True
        tab_3dCaracteristicas.TabEnabled(2) = False
    Else
        tab_3dCaracteristicas.TabEnabled(1) = False
        tab_3dCaracteristicas.TabEnabled(2) = True
    End If
    
End Sub

Private Sub txt_areaFeira_GotFocus()
    MarcaCampo txt_areaFeira
End Sub

Private Sub txt_areaFeira_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_areaFeira
End Sub

Private Sub txt_areaFeira_LostFocus()
    txt_areaFeira = gstrConvVrDoSql(txt_areaFeira, 2)
End Sub

Private Sub txt_CNPJCPF_GotFocus()
    tab_3dPasta.Tab = 0
End Sub

Private Sub txt_CNPJCPF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_CNPJCPF
End Sub

Private Sub txt_CNPJCPFContador_GotFocus()
    tab_3dPasta.Tab = 5
End Sub

Private Sub txt_CNPJCPFContador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CNPJCPFContador
End Sub

Private Sub txt_CRC_GotFocus()
    tab_3dPasta.Tab = 5
End Sub

Private Sub txt_CRC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CRC
End Sub

Private Sub txt_dblArea_GotFocus()
    If Trim(txt_dblArea) = "" Then
        txt_dblArea = "1"
    End If
    MarcaCampo txt_dblArea
End Sub

Private Sub txt_dblArea_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblArea
End Sub

Private Sub txt_dblArea_LostFocus()
    If Trim(txt_dblArea) = "" Then
        txt_dblArea = "1"
    ElseIf CDbl(Val(gstrConvVrParaSql(txt_dblArea))) = 0 Then
        txt_dblArea = "1"
    End If
    txt_dblArea = gstrConvVrDoSql(txt_dblArea.Text, 5)
    
End Sub

Private Sub txt_dtmAtividadeFim_Click()
    blnIssORAtividades = False
End Sub

Private Sub txt_dtmAtividadeFim_GotFocus()
    blnIssORAtividades = False
    MarcaCampo txt_dtmAtividadeFim
End Sub

Private Sub txt_dtmAtividadeFim_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmAtividadeFim
End Sub

Private Sub txt_dtmAtividadeFim_LostFocus()
    txt_dtmAtividadeFim = gstrDataFormatada(txt_dtmAtividadeFim)
    If Len(txt_dtmAtividadeFim) > 0 Then
        If Len(txt_dtmAtividadeInicio) = 0 Then
            ExibeMensagem "A data de início tem que ser informada."
            txt_dtmAtividadeFim.Text = ""
            txt_dtmAtividadeInicio.SetFocus
        ElseIf CDate(txt_dtmAtividadeFim) < CDate(txt_dtmAtividadeInicio) Then
            ExibeMensagem "A data de término não pode ser menor que a data de início."
            txt_dtmAtividadeFim.SetFocus
        End If
    End If
End Sub

Private Sub txt_dtmAtividadeInicio_Click()
    blnIssORAtividades = False
End Sub

Private Sub txt_dtmAtividadeInicio_GotFocus()
    blnIssORAtividades = False
    If Len(txt_dtmAtividadeInicio) = 0 Then txt_dtmAtividadeInicio = gstrDataDoSistema
    MarcaCampo txt_dtmAtividadeInicio
End Sub

Private Sub txt_dtmAtividadeInicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmAtividadeInicio
End Sub

Private Sub txt_dtmAtividadeInicio_LostFocus()
    txt_dtmAtividadeInicio = gstrDataFormatada(txt_dtmAtividadeInicio)
    If Len(txt_dtmAtividadeInicio) = 0 Then
        txt_dtmAtividadeFim.Text = ""
    End If
End Sub

Private Sub txt_dtmissfim_Click()
    blnIssORAtividades = True
End Sub

Private Sub txt_dtmissfim_GotFocus()
    blnIssORAtividades = True
    MarcaCampo txt_dtmissfim
End Sub

Private Sub txt_dtmissfim_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmissfim
End Sub

Private Sub txt_dtmissfim_LostFocus()
    txt_dtmissfim = gstrDataFormatada(txt_dtmissfim)
End Sub

Private Sub txt_dtmissinicio_Click()
    blnIssORAtividades = True
End Sub

Private Sub txt_dtmissinicio_GotFocus()
    blnIssORAtividades = True
    txt_dtmissinicio = gstrDataDoSistema
    MarcaCampo txt_dtmissinicio
End Sub

Private Sub txt_dtmissinicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmissinicio
End Sub

Private Sub txt_dtmissinicio_LostFocus()
    txt_dtmissinicio = gstrDataFormatada(txt_dtmissinicio)
End Sub

Private Sub txt_dtmPublicidadeFim_GotFocus()
    MarcaCampo txt_dtmPublicidadeFim
End Sub

Private Sub txt_dtmPublicidadeFim_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmPublicidadeFim
End Sub

Private Sub txt_dtmPublicidadeFim_LostFocus()
    txt_dtmPublicidadeFim = gstrDataFormatada(txt_dtmPublicidadeFim)
    If Len(txt_dtmPublicidadeFim) > 0 Then
        If Len(txt_dtmPublicidadeInicio) = 0 Then
            ExibeMensagem "A data de início tem que ser informada."
            txt_dtmPublicidadeFim.Text = ""
            txt_dtmPublicidadeInicio.SetFocus
        ElseIf CDate(txt_dtmPublicidadeFim) < CDate(txt_dtmPublicidadeInicio) Then
            ExibeMensagem "A data de término não pode ser menor que a data de início."
            txt_dtmPublicidadeFim.SetFocus
        End If
    End If
End Sub

Private Sub txt_dtmPublicidadeInicio_GotFocus()
    If Len(txt_dtmPublicidadeInicio) = 0 Then txt_dtmPublicidadeInicio = gstrDataDoSistema
    MarcaCampo txt_dtmPublicidadeInicio
End Sub

Private Sub txt_dtmPublicidadeInicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmPublicidadeInicio
End Sub

Private Sub txt_dtmPublicidadeInicio_LostFocus()
    txt_dtmPublicidadeInicio = gstrDataFormatada(txt_dtmPublicidadeInicio)
    If Len(txt_dtmPublicidadeInicio) = 0 Then
        txt_dtmPublicidadeFim.Text = ""
    End If
End Sub

Private Sub txt_dtmSocioFim_GotFocus()
    MarcaCampo txt_dtmSocioFim
End Sub

Private Sub txt_dtmSocioFim_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmSocioFim
End Sub

Private Sub txt_dtmSocioFim_LostFocus()
    txt_dtmSocioFim = gstrDataFormatada(txt_dtmSocioFim)
    If Len(txt_dtmSocioFim) > 0 Then
        If Len(txt_dtmSocioInicio) = 0 Then
            ExibeMensagem "A data de início tem que ser informada."
            txt_dtmSocioFim.Text = ""
            txt_dtmSocioInicio.SetFocus
        ElseIf CDate(txt_dtmSocioFim) < CDate(txt_dtmSocioInicio) Then
            ExibeMensagem "A data de término não pode ser menor que a data de início."
            txt_dtmSocioFim.SetFocus
        End If
    End If
End Sub

Private Sub txt_dtmSocioInicio_GotFocus()
    If Len(txt_dtmSocioInicio) = 0 Then txt_dtmSocioInicio = gstrDataDoSistema
    MarcaCampo txt_dtmSocioInicio
End Sub

Private Sub txt_dtmSocioInicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmSocioInicio
End Sub

Private Sub txt_dtmSocioInicio_LostFocus()
    txt_dtmSocioInicio = gstrDataFormatada(txt_dtmSocioInicio)
    If Len(txt_dtmSocioInicio) = 0 Then
        txt_dtmSocioFim.Text = ""
    End If
End Sub

Private Sub txt_intQtd_Click()
    blnIssORAtividades = False
End Sub

Private Sub txt_intQtd_GotFocus()
    blnIssORAtividades = False
    txt_intQtd = IIf(Val(txt_intQtd) = 0, 1, txt_intQtd)
    MarcaCampo txt_intQtd
End Sub

Private Sub txt_intQtd_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intQtd
End Sub

Private Sub txt_intquantidadeiss_Click()
    blnIssORAtividades = True
End Sub

Private Sub txt_intquantidadeiss_GotFocus()
    blnIssORAtividades = True
    MarcaCampo txt_intQuantidadeIss
End Sub

Private Sub txt_intQuantidadeIss_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intQuantidadeIss
End Sub

Private Sub txt_strCotas_GotFocus()
    MarcaCampo txt_strCotas
End Sub

Private Sub txt_strCotas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strCotas
End Sub

Private Sub txt_strnrbox_GotFocus()
    MarcaCampo txt_strnrbox
    tab_3dPasta.Tab = 8
End Sub

Private Sub txt_strnrbox_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strnrbox
End Sub

Private Sub txt_strObservacao_GotFocus()
    MarcaCampo txt_strObservacao
    tab_3dPasta.Tab = 7
End Sub

Private Sub txtbitDigProcAbertura_GotFocus()
    MarcaCampo txtbitDigProcAbertura
End Sub

Private Sub txtbitDigProcAbertura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigProcAbertura
End Sub

Private Sub txtbitDigProcEncerramento_GotFocus()
    MarcaCampo txtbitDigProcEncerramento
End Sub

Private Sub txtbitDigProcEncerramento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigProcEncerramento
End Sub

Private Sub txtbitDigProcesso_GotFocus()
    MarcaCampo txtbitdigprocesso
End Sub

Private Sub txtbitDigProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitdigprocesso
End Sub

Private Sub txtdblAreaAnuncio_GotFocus()
    MarcaCampo txtdblAreaAnuncio
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtdblAreaAnuncio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblAreaAnuncio
End Sub

Private Sub txtdblAreaAnuncio_LostFocus()
    txtdblAreaAnuncio = gvntConvVrDoSql(txtdblAreaAnuncio)
End Sub

Private Sub txtdblAreaOcupada_GotFocus()
    MarcaCampo txtDblareaocupada
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtdblAreaOcupada_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtDblareaocupada
End Sub

Private Sub txtdblAreaOcupada_LostFocus()
    txtDblareaocupada = gvntConvVrDoSql(txtDblareaocupada)
End Sub
Private Sub txtdblValorEstimado_GotFocus()
    MarcaCampo txtdblValorEstimado
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtdblValorEstimado_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValorEstimado
End Sub

Private Sub txtdblValorEstimado_LostFocus()
    txtdblValorEstimado = gvntConvVrDoSql(txtdblValorEstimado)
End Sub

Private Sub txtdtmDataEstimativa_GotFocus()
    MarcaCampo txtdtmDataEstimativa
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtdtmDataEstimativa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataEstimativa
End Sub

Private Sub txtdtmDataEstimativa_LostFocus()
    txtdtmDataEstimativa = gstrDataFormatada(txtdtmDataEstimativa)
End Sub

Private Sub txtdtmdataprocesso_GotFocus()
    MarcaCampo txtdtmdataprocesso
End Sub

Private Sub txtdtmdataprocesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdataprocesso
End Sub

Private Sub txtdtmdataprocesso_LostFocus()
    txtdtmdataprocesso = gstrDataFormatada(txtdtmdataprocesso)
End Sub

Private Sub txtdtmEnderecoInicio_GotFocus()
    If Len(txtdtmEnderecoInicio) = 0 Then txtdtmEnderecoInicio = gstrDataDoSistema
    MarcaCampo txtdtmEnderecoInicio
End Sub

Private Sub txtdtmEnderecoInicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmEnderecoInicio
End Sub

Private Sub txtdtmEnderecoInicio_LostFocus()
    txtdtmEnderecoInicio = gstrDataFormatada(txtdtmEnderecoInicio)
End Sub

Private Sub txtdtmRazaoInicio_GotFocus()
    If Len(txtdtmRazaoInicio) = 0 Then txtdtmRazaoInicio = gstrDataDoSistema
    MarcaCampo txtdtmRazaoInicio
End Sub

Private Sub txtdtmRazaoInicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmRazaoInicio
End Sub

Private Sub txtdtmRazaoInicio_LostFocus()
    txtdtmRazaoInicio = gstrDataFormatada(txtdtmRazaoInicio)
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub

Private Sub txtintCEP_LostFocus()
    txtintCep = gstrCEPFormatado(txtintCep)
    CepLogradouro txtintCep, dbcintLogradouro, dbcintBairro, txt_strMunicipio, txt_strUF, , , , True, True, False, False, , , True, False, ""
    If Trim(txt_strMunicipio) = "" Then
        txt_strMunicipio = gstrCidadeEmpresa
    End If
End Sub

Private Sub txtintComponentes_GotFocus()
    tab_3dPasta.Tab = 1
    MarcaCampo txtintComponentes
End Sub

Private Sub txtintExerProcAbertura_GotFocus()
    MarcaCampo txtintExerProcAbertura
End Sub

Private Sub txtintExerProcAbertura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExerProcAbertura
End Sub

Private Sub txtintExerProcEncerramento_GotFocus()
    MarcaCampo txtintExerProcEncerramento
End Sub

Private Sub txtintExerProcEncerramento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExerProcEncerramento
End Sub

Private Sub txtintExerProcesso_GotFocus()
    MarcaCampo txtintexerprocesso
End Sub

Private Sub txtintExerProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintexerprocesso
End Sub

Private Sub txtintNumDeEmpregados_GotFocus()
    tab_3dPasta.Tab = 1
    MarcaCampo txtintNumDeEmpregados
End Sub

Private Sub txtintNumDeEmpregados_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumDeEmpregados
End Sub

Private Sub txtstrCodProcAbertura_GotFocus()
    MarcaCampo txtstrCodProcAbertura
End Sub

Private Sub txtstrCodProcAbertura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodProcAbertura
End Sub

Private Sub txtstrCodProcEncerramento_GotFocus()
    MarcaCampo txtstrCodProcEncerramento
End Sub

Private Sub txtstrCodProcEncerramento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodProcEncerramento
End Sub


Private Sub txtstrCodProcesso_GotFocus()
    MarcaCampo txtstrcodprocesso
End Sub

Private Sub txtstrCodProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrcodprocesso
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub txtdtmDataAbertura_GotFocus()
    MarcaCampo txtDtmdataabertura
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtdtmDataAbertura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDtmdataabertura
End Sub

Private Sub txtdtmDataAbertura_LostFocus()
    txtDtmdataabertura = gstrDataFormatada(txtDtmdataabertura)
End Sub

Private Sub txtdtmDataEncerramento_GotFocus()
    MarcaCampo txtdtmDataEncerramento
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtdtmDataEncerramento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataEncerramento
End Sub

Private Sub txtdtmDataEncerramento_LostFocus()
    txtdtmDataEncerramento = gstrDataFormatada(txtdtmDataEncerramento)
End Sub

Private Sub txtstrEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrEmissao
End Sub

Private Sub txtstrhistoricoprocesso_GotFocus()
    MarcaCampo txtstrhistoricoprocesso
End Sub

Private Sub txtstrMadrugadaAte_KeyPress(KeyAscii As Integer)
    dbcintHorarioFuncionamento.BoundText = ""
    CaracterValido KeyAscii, "H", txtstrMadrugadaAte
End Sub

Private Sub txtstrMadrugadaDe_KeyPress(KeyAscii As Integer)
    dbcintHorarioFuncionamento.BoundText = ""
    CaracterValido KeyAscii, "H", txtstrMadrugadaDe
End Sub

Private Sub txtstrManhaAte_KeyPress(KeyAscii As Integer)
    dbcintHorarioFuncionamento.BoundText = ""
    CaracterValido KeyAscii, "H", txtstrManhaAte
End Sub

Private Sub txtstrManhaDe_GotFocus()
    MarcaCampo txtstrManhaDe
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtstrManhaAte_GotFocus()
    MarcaCampo txtstrManhaAte
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtintNumero_GotFocus()
    MarcaCampo txtintNumero
    tab_3dPasta.Tab = 0
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumero
    TrocaCorObjeto txtdtmEnderecoInicio, False
    If Not mblnLoading Then txtdtmEnderecoInicio = ""
End Sub

Private Sub txtstrManhaDe_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "H", txtstrManhaDe
    dbcintHorarioFuncionamento.BoundText = ""
End Sub

Private Sub txtstrNoiteAte_KeyPress(KeyAscii As Integer)
    dbcintHorarioFuncionamento.BoundText = ""
    CaracterValido KeyAscii, "H", txtstrNoiteAte
End Sub

Private Sub txtstrNoiteDe_KeyPress(KeyAscii As Integer)
    dbcintHorarioFuncionamento.BoundText = ""
    CaracterValido KeyAscii, "H", txtstrNoiteDe
End Sub

Private Sub txtstrTardeAte_KeyPress(KeyAscii As Integer)
    dbcintHorarioFuncionamento.BoundText = ""
    CaracterValido KeyAscii, "H", txtstrTardeAte
End Sub

Private Sub txtstrTardeDe_GotFocus()
    MarcaCampo txtstrTardeDe
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtstrTardeAte_GotFocus()
    MarcaCampo txtstrTardeAte
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtstrNoiteDe_GotFocus()
    MarcaCampo txtstrNoiteDe
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtstrNoiteAte_GotFocus()
    MarcaCampo txtstrNoiteAte
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtstrMadrugadaDe_GotFocus()
    MarcaCampo txtstrMadrugadaDe
    tab_3dPasta.Tab = 1
End Sub

Private Sub txtstrMadrugadaAte_GotFocus()
    MarcaCampo txtstrMadrugadaAte
    tab_3dPasta.Tab = 1
End Sub

Private Function blnDadosOk() As Boolean
    Dim i As Integer
    
    If mblnAlterando = False Then
        If gblnExisteCodigo(1, gstrEconomico, "Strinscricaocadastral", "'" & String(gintLenInscricao - Len(mskstrInscricaoCadastral), "0") & mskstrInscricaoCadastral & "'") Then
            ExibeMensagem "Inscrição cadastral já se encontra cadastrada."
            mskstrInscricaoCadastral.SetFocus
            Exit Function
        End If
    End If
    
    
    If blnInscricaoImobiliaria = False Then
        Exit Function
    End If
    If dbcintContribuinte.BoundText = "" Then
        ExibeMensagem "A razão social tem que ser selecionada."
        dbcintContribuinte.SetFocus
        Exit Function
    End If
    If dbcintLogradouro.BoundText = "" Then
        ExibeMensagem "O logradouro do estabelecimento tem que ser informado."
        dbcintLogradouro.SetFocus
        Exit Function
    End If
    If Trim(txtintNumero.Text) = "" Then
        ExibeMensagem "O número do estabelecimento tem que ser informado."
        txtintNumero.SetFocus
        Exit Function
    End If
    If dbcintBairro.BoundText = "" Then
        ExibeMensagem "O bairro do estabelecimento tem que ser informado."
        dbcintBairro.SetFocus
        Exit Function
    End If
    
    If Not gblnCepValido(txtintCep, dbcintLogradouro) Then
        ExibeMensagem "O cep do estabelecimento é inválido para o logradouro informado."
        txtintCep.SetFocus
        Exit Function
    End If
    
    If Trim(txtdtmRazaoInicio) <> "" Then
        If Not gblnDataValida(txtdtmRazaoInicio) Then
            ExibeMensagem "Data de início da Razão inválida."
            txtdtmRazaoInicio.SetFocus
            Exit Function
        End If
    Else
        ExibeMensagem "Data de início da Razão é obrigatória."
        txtdtmRazaoInicio.SetFocus
        Exit Function
    End If
    
    If Trim(txtdtmEnderecoInicio) <> "" Then
        If Not gblnDataValida(txtdtmEnderecoInicio) Then
            ExibeMensagem "Data de início do Endereço inválida."
            txtdtmEnderecoInicio.SetFocus
            Exit Function
        End If
    Else
        ExibeMensagem "Data de início do Endereço é obrigatória."
        txtdtmEnderecoInicio.SetFocus
        Exit Function
    End If
    
    If Trim(txtDtmdataabertura) <> "" Then
        If Not gblnDataValida(txtDtmdataabertura) Then
            ExibeMensagem "Data de abertura inválida."
            txtDtmdataabertura.SetFocus
            Exit Function
        End If
    Else
        ExibeMensagem "Data de abertura é obrigatória."
        txtDtmdataabertura.SetFocus
        Exit Function
    End If
    
    If (Trim(txtstrCodProcAbertura.Text) <> "" And Trim(txtintExerProcAbertura.Text) <> "" And Trim(txtbitDigProcAbertura.Text) <> "") Then
       If blnConsultaProcesso Then 'Variavel que indica se havera consulta no numero do processo
          If gblnExisteCodigo(2, gstrProtocolizacaoProcesso, "strCodigo", "'" & Trim(txtstrCodProcAbertura.Text) & "'", _
             "intExercicio", Trim(txtintExerProcAbertura.Text), "bitDigito", Trim(txtbitDigProcAbertura.Text)) = False Then
             ExibeMensagem "O Processo de Abertura informado não existe."
             tab_3dPasta.Tab = 1
             txtstrCodProcAbertura.SetFocus
             Exit Function
          End If
       End If
    Else
       If (Trim(txtstrCodProcAbertura.Text) = "" And Trim(txtintExerProcAbertura.Text) = "" And Trim(txtbitDigProcAbertura.Text) = "") Then
          If Not mblnAlterando Then
            ExibeMensagem "O Processo de Abertura deve ser informado."
            tab_3dPasta.Tab = 1
            txtstrCodProcAbertura.SetFocus
            Exit Function
          End If
       Else
          ExibeMensagem "O Processo de Abertura deve ser preenchido corretamente."
          tab_3dPasta.Tab = 1
          txtstrCodProcAbertura.SetFocus
          Exit Function
       End If
    End If
    
    If Trim(txtdtmDataEncerramento.Text) <> "" Then
        If Not gblnDataValida(txtdtmDataEncerramento) Then
            ExibeMensagem "Data de Encerramento inválida."
            tab_3dPasta.Tab = 1
            txtdtmDataEncerramento.SetFocus
            Exit Function
        End If
        If (CDate(txtdtmDataEncerramento) > CDate(gstrDataDoSistema)) Then
            ExibeMensagem "A Data de Encerramento não deve ser maior que a Data Atual."
            tab_3dPasta.Tab = 1
            txtdtmDataEncerramento.SetFocus
            Exit Function
        End If
        If CDate(txtdtmDataEncerramento) < CDate(txtDtmdataabertura) Then
            ExibeMensagem "A data de encerramento não pode ser menor que a data de abertura."
            tab_3dPasta.Tab = 1
            txtdtmDataEncerramento.SetFocus
            Exit Function
        End If
        
        If (Trim(txtstrCodProcEncerramento.Text) = "" Or Trim(txtintExerProcEncerramento.Text) = "" Or Trim(txtbitDigProcEncerramento.Text) = "") Then
            ExibeMensagem "O Processo de Encerramento deve ser preenchido corretamente quando a data de encerramento estiver preenchida."
            tab_3dPasta.Tab = 1
            txtstrCodProcEncerramento.SetFocus
            Exit Function
        End If
        
    Else
        If (Trim(txtstrCodProcEncerramento.Text) <> "" Or Trim(txtintExerProcEncerramento.Text) <> "" Or Trim(txtbitDigProcEncerramento.Text) <> "") Then
            ExibeMensagem "A data de encerramento deve ser preenchida corretamente quando os campos de processos estiverem preenchidos."
            tab_3dPasta.Tab = 1
            txtdtmDataEncerramento.SetFocus
            Exit Function
        End If
    End If
    
    If (Trim(txtstrCodProcEncerramento.Text) <> "" Or Trim(txtintExerProcEncerramento.Text) <> "" Or Trim(txtbitDigProcEncerramento.Text) <> "") Then
       If (Trim(txtstrCodProcEncerramento.Text) <> "" And Trim(txtintExerProcEncerramento.Text) <> "" And Trim(txtbitDigProcEncerramento.Text) <> "") Then
          If blnConsultaProcesso Then 'Variavel que indica se havera consulta no numero do processo
             If gblnExisteCodigo(2, gstrProtocolizacaoProcesso, "strCodigo", "'" & Trim(txtstrCodProcEncerramento.Text) & "'", _
                "intExercicio", Trim(txtintExerProcEncerramento.Text), "bitDigito", Trim(txtbitDigProcEncerramento.Text)) = False Then
                ExibeMensagem "O Processo de Encerramento informado não existe."
                tab_3dPasta.Tab = 1
                txtstrCodProcEncerramento.SetFocus
                Exit Function
             End If
          End If
       Else
          ExibeMensagem "O Processo de Encerramento deve ser preenchido corretamente."
          tab_3dPasta.Tab = 1
          txtstrCodProcEncerramento.SetFocus
          Exit Function
       End If
    End If
    
    If Trim(txtdtmdataprocesso.Text) <> "" Then
        If Not gblnDataValida(txtdtmdataprocesso) Then
            ExibeMensagem "Data de Ocorrências de Processos inválida."
            tab_3dPasta.Tab = 1
            txtdtmdataprocesso.SetFocus
            Exit Function
        End If
        
        If (Trim(txtstrcodprocesso.Text) = "" Or Trim(txtintexerprocesso.Text) = "" Or Trim(txtbitdigprocesso.Text) = "") Then
            ExibeMensagem "O Processo de Ocorrências de Processos deve ser preenchido corretamente quando a data estiver preenchida."
            tab_3dPasta.Tab = 1
            txtstrcodprocesso.SetFocus
            Exit Function
        End If
        
    Else
        If (Trim(txtstrcodprocesso.Text) <> "" Or Trim(txtintexerprocesso.Text) <> "" Or Trim(txtbitdigprocesso.Text) <> "") Then
            ExibeMensagem "A data de Ocorrências de Processos deve ser preenchida corretamente quando os campos de processos estiverem preenchidos."
            tab_3dPasta.Tab = 1
            txtdtmdataprocesso.SetFocus
            Exit Function
        End If
    End If
    
    If (Trim(txtstrcodprocesso.Text) <> "" Or Trim(txtintexerprocesso.Text) <> "" Or Trim(txtbitdigprocesso.Text) <> "") Then
       If (Trim(txtstrcodprocesso.Text) <> "" And Trim(txtintexerprocesso.Text) <> "" And Trim(txtbitdigprocesso.Text) <> "") Then
          If blnConsultaProcesso Then 'Variavel que indica se havera consulta no numero do processo
             If gblnExisteCodigo(2, gstrProtocolizacaoProcesso, "strCodigo", "'" & Trim(txtstrcodprocesso.Text) & "'", _
                "intExercicio", Trim(txtintexerprocesso.Text), "bitDigito", Trim(txtbitdigprocesso.Text)) = False Then
                ExibeMensagem "O processo de Ocorrências de Processos informado não existe."
                tab_3dPasta.Tab = 1
                txtstrcodprocesso.SetFocus
                Exit Function
             End If
          End If
       Else
          ExibeMensagem "O processo de Ocorrências de Processos deve ser preenchido corretamente."
          tab_3dPasta.Tab = 1
          txtstrcodprocesso.SetFocus
          Exit Function
       End If
    End If
    
    If Trim(dbcintocorrenciadoeconomico.Text) <> "" Then
        If Not dbcintocorrenciadoeconomico.MatchedWithList Then
            ExibeMensagem "O campo Ocorrência de Processo deva ser preenchido corretamente."
            tab_3dPasta.Tab = 1
            dbcintocorrenciadoeconomico.SetFocus
            Exit Function
        End If
    End If
    
    With lvw_Atividades
        If .ListItems.Count > 0 Then
            For i = 1 To .ListItems.Count
                If CBool(.ListItems(i).SubItems(2)) Then
                    GoTo Gatividade
                End If
            Next
            ExibeMensagem "É necessário incluir no mínimo uma atividade principal."
            tab_3dPasta.Tab = 2
            dbc_intAtividadePrincipal.SetFocus
            Exit Function
        Else
            ExibeMensagem "É necessário incluir no mínimo uma atividade."
            tab_3dPasta.Tab = 2
            dbc_intAtividadePrincipal.SetFocus
            Exit Function
        End If
    End With
Gatividade:
    ' *** TIMIM - 14/04/2003 ***
    Dim stpMensMn As String
    Dim stpMensTd As String
    Dim stpMensNt As String
    Dim stpMensMd As String
    
    stpMensMn = "Faixa de horário incorreta para o período da manhã."
    stpMensTd = "Faixa de horário incorreta para o período da tarde."
    stpMensNt = "Faixa de horário incorreta para o período da noite."
    stpMensMd = "Faixa de horário incorreta para o período da madrugada."
    
    If Trim$(txtstrManhaDe.Text) <> Space$(0) And Trim$(txtstrManhaAte.Text) <> Space$(0) Then
        If IsDate(Trim$(txtstrManhaDe.Text)) And IsDate(Trim$(txtstrManhaAte.Text)) Then
            If CDate(Trim$(txtstrManhaDe.Text)) <= CDate(Trim$(txtstrManhaAte.Text)) Then
                stpMensMn = Space$(0)
            End If
        End If
    Else
        If Trim$(txtstrManhaDe.Text) = Space$(0) And Trim$(txtstrManhaAte.Text) = Space$(0) Then
            stpMensMn = Space$(0)
        End If
    End If
    
    If Trim$(txtstrTardeDe.Text) <> Space$(0) And Trim$(txtstrTardeAte.Text) <> Space$(0) Then
        If IsDate(Trim$(txtstrTardeDe.Text)) And IsDate(Trim$(txtstrTardeAte.Text)) Then
            If CDate(Trim$(txtstrTardeDe.Text)) <= CDate(Trim$(txtstrTardeAte.Text)) Then
                stpMensTd = Space$(0)
            End If
        End If
    Else
        If Trim$(txtstrTardeDe.Text) = Space$(0) And Trim$(txtstrTardeAte.Text) = Space$(0) Then
            stpMensTd = Space$(0)
        End If
    End If
    
    If Trim$(txtstrNoiteDe.Text) <> Space$(0) And Trim$(txtstrNoiteAte.Text) <> Space$(0) Then
        If IsDate(Trim$(txtstrNoiteDe.Text)) And IsDate(Trim$(txtstrNoiteAte.Text)) Then
            If CDate(Trim$(txtstrNoiteDe.Text)) <= CDate(Trim$(txtstrNoiteAte.Text)) Then
                stpMensNt = Space$(0)
            End If
        End If
    Else
        If Trim$(txtstrNoiteDe.Text) = Space$(0) And Trim$(txtstrNoiteAte.Text) = Space$(0) Then
            stpMensNt = Space$(0)
        End If
    End If
    
    If Trim$(txtstrMadrugadaDe.Text) <> Space$(0) And Trim$(txtstrMadrugadaAte.Text) <> Space$(0) Then
        If IsDate(Trim$(txtstrMadrugadaDe.Text)) And IsDate(Trim$(txtstrMadrugadaAte.Text)) Then
            If CDate(Trim$(txtstrMadrugadaDe.Text)) <= CDate(Trim$(txtstrMadrugadaAte.Text)) Then
                stpMensMd = Space$(0)
            End If
        End If
    Else
        If Trim$(txtstrMadrugadaDe.Text) = Space$(0) And Trim$(txtstrMadrugadaAte.Text) = Space$(0) Then
            stpMensMd = Space$(0)
        End If
    End If
    
    If stpMensMn <> Space$(0) Then
        ExibeMensagem stpMensMn
        txtstrManhaDe.SetFocus
        Exit Function
    End If
    
    If stpMensTd <> Space$(0) Then
        ExibeMensagem stpMensTd
        txtstrTardeDe.SetFocus
        Exit Function
    End If
    
    If stpMensNt <> Space$(0) Then
        ExibeMensagem stpMensNt
        txtstrNoiteDe.SetFocus
        Exit Function
    End If
    
    If stpMensMd <> Space$(0) Then
        ExibeMensagem stpMensMd
        txtstrMadrugadaDe.SetFocus
        Exit Function
    End If
    
    If dbcintOcorrencia.BoundText = "" Then
        ExibeMensagem "A ocorrência tem que ser informada."
        dbcintOcorrencia.SetFocus
        Exit Function
    End If
   
    blnDadosOk = True

End Function

Private Sub NovoEconomico()
    Dim adoResultado As ADODB.Recordset
    Dim strSql       As String
    
    On Error GoTo err_NovoEconomico
    LimpaTabFeira
    LimpaTabPubli True
    LimpaTabSocios True
    LimpaTabAtividade True
    LimpaISS True
    lvw_Itens.ListItems.Clear
    LimpaTabISS
    lvw_Historico.ListItems.Clear
    txt_Historico = ""
    txt_CNPJCPF = ""
    txt_strInscricaoEstadual = ""
    txt_strNomeFantasia = ""
    txt_PKIdContribuinte = ""
    txt_TotalDeCotas = ""
    
    txt_strMunicipio = Space$(0)
    txt_strUF = Space$(0)
    
    txt_PKIdContador = ""
    txt_CNPJCPFContador = ""
    txt_CRC = ""
    
    txt_strMunicipio = gstrCidadeEmpresa
    txt_strUF = gstrUFEmpresa
      
    opt_NaturezaJuridica(0).Value = False
    opt_NaturezaJuridica(1).Value = False
    opt_NaturezaJuridica(2).Value = False
    opt_NaturezaJuridica(3).Value = False

    tab_3dPasta.Tab = 0
    mskstrInscricaoCadastral.SetFocus
    Set tdb_Processo.DataSource = Nothing
    
    mblnAlterando = False
    mblnAlterandoH = False
    tab_3dPasta.Tab = 0
    
err_NovoEconomico:

End Sub

Function blnGravaLivrosISS(intEconomico As Long) As Boolean
    Dim strSql           As String
    Dim i                As Integer

    
    blnGravaLivrosISS = True
End Function

Function blnGravaAtividades(intEconomico As Long) As Boolean
    Dim strSql              As String
    Dim strSql1             As String
    Dim strSql2             As String
    Dim intFor              As Integer
    Dim strPkidHistorico    As String
    
    blnGravaAtividades = False

    strSql = ""
    strSql2 = ""
    strPkidHistorico = ""
    
    If lvw_Atividades.ListItems.Count <= 0 Then
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
        
        strSql = strSql & "Delete From " & gstrAtivEmpresaTributo & " Where pkid In"
        strSql = strSql & "(Select "
        strSql = strSql & "C.Pkid "
        strSql = strSql & "From "
        strSql = strSql & gstrEconomico & " A, "
        strSql = strSql & gstrAtividadeDaEmpresa & " B, "
        strSql = strSql & gstrAtivEmpresaTributo & " C "
        strSql = strSql & "Where "
        strSql = strSql & "A.Pkid = B.Inteconomico AND "
        strSql = strSql & "B.Pkid = C.Intatividadedaempresa AND "
        strSql = strSql & "A.Pkid = " & intEconomico & ")"
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        
        strSql = strSql & "DELETE FROM " & gstrAtividadeDaEmpresa
        strSql = strSql & " WHERE Inteconomico = " & intEconomico
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    Else
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
        
        For intFor = 1 To lvw_Atividades.ListItems.Count
            With lvw_Atividades
                If .ListItems(intFor).Text <> "" Then
                
                    strPkidHistorico = strPkidHistorico & .ListItems(intFor).Text & ","
                    
                    strSql = strSql & " UPDATE " & gstrAtividadeDaEmpresa
                    strSql = strSql & " SET intatividade = " & .ListItems(intFor).SubItems(1) & ", "
                    strSql = strSql & " blnprincipal = " & Abs(CInt(CBool(.ListItems(intFor).SubItems(2)))) & ", "
                    If .ListItems(intFor).Selected And dbc_intAtividadePrincipal.MatchedWithList And mblnAlterandoListaAtividade Then
                        strSql = strSql & " dtmAtividadeInicio = " & gstrConvDtParaSql(txt_dtmAtividadeInicio) & ", "
                        strSql = strSql & " dtmAtividadeFim = " & gstrConvDtParaSql(txt_dtmAtividadeFim) & ", "
                    Else
                        strSql = strSql & " dtmAtividadeInicio = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(5)) & ", "
                        strSql = strSql & " dtmAtividadeFim = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(6)) & ", "
                    End If
                    strSql = strSql & " dtmDtAtualizacao = " & strGETDATE & ", "
                    strSql = strSql & " lngCodUsr = " & glngCodUsr
                    strSql = strSql & " WHERE Pkid = " & .ListItems(intFor).Text
                    strSql = strSql & IIf(bytDBType = Oracle, ";", "")
                
                Else
                    
                    strSql2 = strSql2 & " INSERT INTO " & gstrAtividadeDaEmpresa
                    strSql2 = strSql2 & " (inteconomico,"
                    strSql2 = strSql2 & " intatividade,"
                    strSql2 = strSql2 & " blnprincipal,"
                    strSql2 = strSql2 & " dtmDtAtualizacao,"
                    strSql2 = strSql2 & " dtmAtividadeInicio,"
                    strSql2 = strSql2 & " dtmAtividadeFim,"
                    strSql2 = strSql2 & " lngCodUsr)"
                    strSql2 = strSql2 & " VALUES( "
                    strSql2 = strSql2 & Val(intEconomico) & ", "
                    strSql2 = strSql2 & .ListItems(intFor).SubItems(1) & ", "
                    strSql2 = strSql2 & Abs(CInt(CBool(.ListItems(intFor).SubItems(2)))) & ", "
                    strSql2 = strSql2 & strGETDATE & ", "
                    strSql2 = strSql2 & gstrConvDtParaSql(.ListItems(intFor).SubItems(5)) & ", "
                    strSql2 = strSql2 & gstrConvDtParaSql(.ListItems(intFor).SubItems(6)) & ", "
                    strSql2 = strSql2 & glngCodUsr
                    strSql2 = strSql2 & ")"
                    strSql2 = strSql2 & IIf(bytDBType = Oracle, ";", "")
                End If
            End With
        Next
        
        If strPkidHistorico <> "" Then
            strPkidHistorico = Mid(strPkidHistorico, 1, Len(strPkidHistorico) - 1)
            
            strSql1 = strSql1 & " Delete From " & gstrAtivEmpresaTributo & " Where pkid In"
            strSql1 = strSql1 & "(Select "
            strSql1 = strSql1 & "C.Pkid "
            strSql1 = strSql1 & "From "
            strSql1 = strSql1 & gstrEconomico & " A, "
            strSql1 = strSql1 & gstrAtividadeDaEmpresa & " B, "
            strSql1 = strSql1 & gstrAtivEmpresaTributo & " C "
            strSql1 = strSql1 & "Where "
            strSql1 = strSql1 & "A.Pkid = B.Inteconomico AND "
            strSql1 = strSql1 & "B.Pkid = C.Intatividadedaempresa AND "
            strSql1 = strSql1 & "B.Pkid NOT in(" & strPkidHistorico & ") and "
            strSql1 = strSql1 & "A.Pkid = " & intEconomico & ")"
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", "")
            
            strSql1 = strSql1 & " DELETE FROM " & gstrAtividadeDaEmpresa
            strSql1 = strSql1 & " WHERE Pkid NOT in(" & strPkidHistorico & ") and "
            strSql1 = strSql1 & " inteconomico = " & intEconomico
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", " ")
            strSql = strSql & " " & strSql1
            
        Else
            strSql1 = strSql1 & " Delete From " & gstrAtivEmpresaTributo & " Where pkid In"
            strSql1 = strSql1 & "(Select "
            strSql1 = strSql1 & "C.Pkid "
            strSql1 = strSql1 & "From "
            strSql1 = strSql1 & gstrEconomico & " A, "
            strSql1 = strSql1 & gstrAtividadeDaEmpresa & " B, "
            strSql1 = strSql1 & gstrAtivEmpresaTributo & " C "
            strSql1 = strSql1 & "Where "
            strSql1 = strSql1 & "A.Pkid = B.Inteconomico AND "
            strSql1 = strSql1 & "B.Pkid = C.Intatividadedaempresa AND "
            strSql1 = strSql1 & "A.Pkid = " & intEconomico & ")"
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", "")
            
            strSql1 = strSql1 & " DELETE FROM " & gstrAtividadeDaEmpresa
            strSql1 = strSql1 & " WHERE Inteconomico = " & intEconomico
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", "")
            
            strSql = strSql & " " & strSql1
        End If
        
        strSql = strSql & " " & strSql2
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    End If
    
    If gobjBanco.Execute(strSql) = False Then
        ExibeMensagem "Ocorreu um erro ao gravar as Atividades. Os dados não foram gravados."
        blnGravaAtividades = False
        Exit Function
    Else
        blnGravaAtividades = True
    End If
    
End Function

Sub DeletaTributos(intPKIdEconomico As Long)
    Dim strSql As String

    strSql = ""
    strSql = strSql & "Delete From " & gstrTributoEmpresa & " "
    strSql = strSql & "Where intEconomico = " & intPKIdEconomico
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
End Sub

Sub DeletaFaixas(intPKIdEconomico As Long)
Dim strSql As String

    strSql = ""
    strSql = strSql & "Delete From " & gstrValorFaixaEmpresa & " "
    strSql = strSql & "Where intEconomico = " & intPKIdEconomico
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
End Sub

Private Sub tlb_Historico_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case gstrSalvar
            If Trim(txt_Historico) = "" Then Exit Sub
            If mblnAlterandoH Then
                lvw_Historico.SelectedItem.Text = txt_Historico
            Else
                lvw_Historico.ListItems.Add , , txt_Historico
            End If
        Case gstrNovo
            txt_Historico.SetFocus
        Case gstrDeletar
            If lvw_Historico.ListItems.Count = 0 Then Exit Sub
            If lvw_Historico.SelectedItem.Selected Then
                lvw_Historico.ListItems.Remove (lvw_Historico.SelectedItem.Index)
            End If
    End Select
    If lvw_Historico.ListItems.Count <> 0 Then
        lvw_Historico.SelectedItem.Selected = False
    End If
    mblnAlterandoH = False
    txt_Historico = ""
    txt_Historico.SetFocus
End Sub

Private Sub txt_Historico_GotFocus()
    MarcaCampo txt_Historico
    tab_3dPasta.Tab = 4
End Sub

Private Sub txt_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Historico
End Sub

Sub MontaColumnHeaders()
    With lvw_Historico
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Histórico", 8970
    End With
End Sub

Function blnGravaHistoricos(intEconomico As Long) As Boolean
    Dim strSql           As String
    Dim intI             As Integer
    
    blnGravaHistoricos = False
    
    strSql = ""
    strSql = strSql & "DELETE FROM " & gstrHistoricoEconomico & " "
    strSql = strSql & "WHERE intEconomico = " & intEconomico
    
    If Not gobjBanco.Execute(strSql) Then
        ExibeMensagem "Ocorreu um erro ao gravar o histórico. Os dados não foram gravados."
        Exit Function
    End If
    
    With lvw_Historico
        For intI = 1 To .ListItems.Count
            strSql = ""
            strSql = strSql & "Insert Into " & gstrHistoricoEconomico & " "
            strSql = strSql & "(intEconomico, strDescricao "
            strSql = strSql & ") Values ("
            strSql = strSql & intEconomico & ",'"
            strSql = strSql & .ListItems(intI).Text & "'"
            strSql = strSql & ")"
            If Not gobjBanco.Execute(strSql) Then
                ExibeMensagem "Ocorreu um erro ao gravar o histórico. Os dados não foram gravados."
                Exit Function
            End If
        Next
    End With

    blnGravaHistoricos = True
    
End Function

Private Sub CarregaHistoricos(intCodEconomico As Long)
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    
    lvw_Historico.ListItems.Clear
    txt_Historico = ""
    
    strSql = ""
    strSql = strSql & "SELECT HI.strDescricao AS Historico "
    strSql = strSql & "FROM " & gstrHistoricoEconomico & " HI "
    strSql = strSql & "WHERE HI.intEconomico = " & intCodEconomico
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set oList = lvw_Historico.ListItems.Add(, , Trim(!Historico))
                .MoveNext
            Loop
        End With
    End If
    If lvw_Historico.ListItems.Count <> 0 Then
        lvw_Historico.SelectedItem.Selected = False
    End If
End Sub


Function blnGravaSocios(intEconomico As Long) As Boolean
    Dim strSql              As String
    Dim strSql1             As String
    Dim strSql2             As String
    Dim intFor              As Integer
    Dim strPkidHistorico    As String
    
    blnGravaSocios = False
    
    strSql = ""
    strSql2 = ""
    strPkidHistorico = ""
    
    If lvw_Socios.ListItems.Count <= 0 Then
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrSocioEconomico
        strSql = strSql & " WHERE intCodEconomico = " & intEconomico
    Else
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
        
        For intFor = 1 To lvw_Socios.ListItems.Count
            With lvw_Socios
                If .ListItems(intFor).Text <> "" Then
                
                    strPkidHistorico = strPkidHistorico & .ListItems(intFor).Text & ","
                    
                    strSql = strSql & "UPDATE " & gstrSocioEconomico
                    strSql = strSql & " SET intsocio = " & .ListItems(intFor).SubItems(1) & ", "
                    strSql = strSql & " intnumerodecotas = " & gstrConvVrParaSql(.ListItems(intFor).SubItems(4)) & ", "
                    If .ListItems(intFor).Selected And dbc_intSocios.MatchedWithList Then
                        strSql = strSql & " dtmSocioInicio = " & gstrConvDtParaSql(txt_dtmSocioInicio) & ", "
                        strSql = strSql & " dtmSocioFim = " & gstrConvDtParaSql(txt_dtmSocioFim) & ", "
                    Else
                        strSql = strSql & " dtmSocioInicio = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(5)) & ", "
                        strSql = strSql & " dtmSocioFim = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(6)) & ", "
                    End If
                    strSql = strSql & " dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSql = strSql & " lngCodUsr = " & glngCodUsr
                    strSql = strSql & " WHERE Pkid = " & .ListItems(intFor).Text
                    strSql = strSql & IIf(bytDBType = Oracle, ";", "")
                Else
                    strSql2 = strSql2 & "INSERT INTO " & gstrSocioEconomico
                    strSql2 = strSql2 & " (intcodeconomico,"
                    strSql2 = strSql2 & " intsocio,"
                    strSql2 = strSql2 & " intnumerodecotas,"
                    strSql2 = strSql2 & " dtmSocioInicio,"
                    strSql2 = strSql2 & " dtmSocioFim,"
                    strSql2 = strSql2 & " dtmDtAtualizacao,"
                    strSql2 = strSql2 & " lngCodUsr)"
                    strSql2 = strSql2 & " VALUES( "
                    strSql2 = strSql2 & Val(intEconomico) & ", "
                    strSql2 = strSql2 & .ListItems(intFor).SubItems(1) & ", "
                    strSql2 = strSql2 & gstrConvVrParaSql(.ListItems(intFor).SubItems(4)) & ", "
                    strSql2 = strSql2 & gstrConvDtParaSql(.ListItems(intFor).SubItems(5)) & ", "
                    strSql2 = strSql2 & gstrConvDtParaSql(.ListItems(intFor).SubItems(6)) & ", "
                    strSql2 = strSql2 & strGETDATE & ", "
                    strSql2 = strSql2 & glngCodUsr
                    strSql2 = strSql2 & ")"
                    strSql2 = strSql2 & IIf(bytDBType = Oracle, ";", "")
                End If
            End With
        Next
        
        If strPkidHistorico <> "" Then
            strPkidHistorico = Mid(strPkidHistorico, 1, Len(strPkidHistorico) - 1)
            strSql1 = ""
            strSql1 = strSql1 & "DELETE FROM " & gstrSocioEconomico
            strSql1 = strSql1 & " WHERE Pkid NOT in(" & strPkidHistorico & ")and"
            strSql1 = strSql1 & " intcodeconomico = " & intEconomico
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", " ")
            strSql = strSql & " " & strSql1
        Else
            strSql1 = strSql1 & "DELETE FROM " & gstrSocioEconomico
            strSql1 = strSql1 & " WHERE intcodeconomico = " & intEconomico
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", "")
            strSql = strSql & " " & strSql1
        End If
        strSql = strSql & " " & strSql2
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    End If
    
    If gobjBanco.Execute(strSql) = False Then
        ExibeMensagem "Ocorreu um erro ao gravar os Sócios. Os dados não foram gravados."
        blnGravaSocios = False
        Exit Function
    Else
        blnGravaSocios = True
    End If
End Function

Private Sub txtstrTardeDe_KeyPress(KeyAscii As Integer)
    dbcintHorarioFuncionamento.BoundText = ""
    CaracterValido KeyAscii, "H", txtstrTardeDe
End Sub

Private Sub LocalizarEconomico()
    Dim strSql      As String
    Dim strCondicao As String
    Dim strValor    As String
    Dim strCampo    As String
    Dim i           As Integer
    
    strCondicao = ""
    With Me
        For i = 0 To .Controls.Count - 1
            If Not TypeOf .Controls(i) Is Label Then 'Elimina os Label's da pesquisa
                'Elimina objetos indesejáveis
                If UCase(.Controls(i).Name) <> "TXTPKId" _
                And UCase(Left(.Controls(i).Name, 3)) <> "IMG" _
                And UCase(Left(.Controls(i).Name, 3)) <> "LVW" _
                And UCase(Left(.Controls(i).Name, 3)) <> "TLB" _
                And UCase(Left(.Controls(i).Name, 3)) <> "TDD" _
                And UCase(Left(.Controls(i).Name, 3)) <> "GRD" _
                And UCase(Left(.Controls(i).Name, 4)) <> "TXT_" _
                And UCase(Left(.Controls(i).Name, 4)) <> "DBC_" _
                And UCase(Left(.Controls(i).Name, 4)) <> "CHK_" _
                And UCase(Left(.Controls(i).Name, 4)) <> "OPT_" Then
                
                    If Not (TypeOf .Controls(i) Is OptionButton) Or .Controls(i) = True Then 'Elimina OptionButton desmarcado
                        If TypeOf .Controls(i) Is TextBox Then
                            If Trim(.Controls(i).Text) <> "" Then
                                If InStr(1, .Controls(i).Name, "Cep") > 0 Then
                                    strValor = Val(gstrValorSemMascara(Trim(.Controls(i).Text)))
                                Else
                                    strValor = Trim(.Controls(i).Text)
                                End If
                                strCampo = Trim(.Controls(i).Name)
                                
                                If InStr(1, "_", strCampo) > 0 Then
                                    strCampo = Mid(strCampo, 5, Len(strCampo))
                                Else
                                    strCampo = Mid(strCampo, 4, Len(strCampo))
                                End If
                                strCampo = "EC." & strCampo
                                
                                If InStr(1, "%", strValor) > 0 Then
                                    If strCondicao <> "" Then
                                        strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & strValor & "'"
                                    Else
                                        strCondicao = strCampo & " LIKE '" & strValor & "'"
                                    End If
                                ElseIf InStr(1, UCase(.Controls(i).Name), "DTM") > 0 Then
                                    If strCondicao <> "" Then
                                        strCondicao = strCondicao & " AND " & strCampo & " = " & gstrConvDtParaSql(strValor)
                                    Else
                                        strCondicao = strCampo & " = " & gstrConvDtParaSql(strValor)
                                    End If
                                ElseIf InStr(1, UCase(.Controls(i).Name), "DBL") > 0 Then
                                    If strCondicao <> "" Then
                                        strCondicao = strCondicao & " AND " & strCampo & " = " & gstrConvVrParaSql(strValor)
                                    Else
                                        strCondicao = strCampo & " = " & gstrConvVrParaSql(strValor)
                                    End If
                                Else
                                    If strCondicao <> "" Then
                                        strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & strValor & "%'"
                                    Else
                                        strCondicao = strCampo & " LIKE '" & strValor & "%'"
                                    End If
                                End If
                            End If
                        ElseIf TypeOf .Controls(i) Is OptionButton Then
                            strValor = .Controls(i).Index
                            strCampo = Trim(.Controls(i).Name)
                            If InStr(1, "_", strCampo) > 0 Then
                                strCampo = Mid(strCampo, 5, Len(strCampo))
                            Else
                                strCampo = Mid(strCampo, 4, Len(strCampo))
                            End If
                            strCampo = "EC." & strCampo
                            
                            If strCondicao <> "" Then
                                strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                            Else
                                strCondicao = strCampo & " = " & strValor
                            End If
                        ElseIf TypeOf .Controls(i) Is CheckBox Then
                            If .Controls(i).Value = 1 Then
                                strValor = .Controls(i).Value
                                strCampo = Trim(.Controls(i).Name)
                                
                                If InStr(1, "_", strCampo) > 0 Then
                                    strCampo = Mid(strCampo, 5, Len(strCampo))
                                Else
                                    strCampo = Mid(strCampo, 4, Len(strCampo))
                                End If
                                strCampo = "EC." & strCampo
                                
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                                Else
                                    strCondicao = strCampo & " = " & strValor
                                End If
                            End If
                        ElseIf TypeOf .Controls(i) Is DataCombo Then
                            If .Controls(i).MatchedWithList Then
                                strValor = .Controls(i).BoundText
                                strCampo = Trim(.Controls(i).Name)
                                
                                If InStr(1, "_", strCampo) > 0 Then
                                    strCampo = Mid(strCampo, 5, Len(strCampo))
                                Else
                                    strCampo = Mid(strCampo, 4, Len(strCampo))
                                End If
                                strCampo = "EC." & strCampo
                                
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " = " & strValor
                                Else
                                    strCondicao = strCampo & " = " & strValor
                                End If
                            End If
                        ElseIf TypeOf .Controls(i) Is MaskEdBox Then
                            If Trim(.Controls(i).ClipText) <> "" Then
                                strValor = .Controls(i).ClipText
                                strCampo = Trim(.Controls(i).Name)
                                
                                If InStr(1, "_", strCampo) > 0 Then
                                    strCampo = Mid(strCampo, 5, Len(strCampo))
                                Else
                                    strCampo = Mid(strCampo, 4, Len(strCampo))
                                End If
                                strCampo = "EC." & strCampo
                                
                                If strCampo = "EC.strInscricaoCadastral" Or strCampo = "EC.strInscricaoImobiliaria" Then strValor = String(gintLenInscricao - Len(strValor), "0") & strValor
                                
                                If strCondicao <> "" Then
                                    strCondicao = strCondicao & " AND " & strCampo & " LIKE '" & strValor & "%'"
                                Else
                                    strCondicao = strCampo & " LIKE '" & strValor & "%'"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End With

    strSql = ""
    If strCondicao <> "" Then
        strSql = strSql & "SELECT EC.PKId, CO.strNome, " & gstrRIGHT("EC.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricaoCadastral "
        strSql = strSql & " FROM " & gstrEconomico & " EC,"
        strSql = strSql & gstrContribuinte & " CO "
        strSql = strSql & " WHERE EC.intContribuinte = CO.PKId AND "
        strSql = strSql & strCondicao
        strSql = strSql & " ORDER BY EC.strInscricaoCadastral "
    Else
        strSql = strQuery
    End If
    LeDaTabelaParaObj gstrEconomico, tdb_Lista, strSql

End Sub

Private Function strQueryProtocolizacaoProcesso() As String
    Dim strSql As String
    
    strSql = "SELECT PKId, "
    strSql = strSql & " strCodigo " & strCONCAT & "'/'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "intExercicio ") & strCONCAT & " '-' " & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_NVARCHAR, "bitDigito") & " AS Processo"
    strSql = strSql & " FROM "
    strSql = strSql & gstrProtocolizacaoProcesso
    strSql = strSql & " ORDER BY " & gstrCONVERT(cdt_numeric, "strCodigo") & " , intExercicio, bitDigito"
    
    strQueryProtocolizacaoProcesso = strSql

End Function
Private Function strQueryHistoricoProcessoGrid() As String
    Dim strSql As String
   
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "PEC.Dtmdtprocesso, "
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "PEC.Strcodigoprocesso") & strCONCAT & " '/' " & strCONCAT & gstrCONVERT(CDT_VARCHAR, " PEC.Intexercicioprocesso") & strCONCAT & " ' - ' " & strCONCAT & gstrCONVERT(CDT_VARCHAR, "PEC.BITDIGITOPROCESSO") & " As strProcesso, "
    strSql = strSql & "OEC.Strdescricao as strOcorrencia, PEC.strObservacao "
    strSql = strSql & "From " & gstrEconomico & " EC, "
    strSql = strSql & gstrProcessoEconomico & " PEC, "
    strSql = strSql & gstrOcorrenciaDoEconomico & " OEC "
    strSql = strSql & "Where EC.Pkid = PEC.Inteconomico AND "
    strSql = strSql & "OEC.Pkid = PEC.Intocorrenciadoeconomico AND "
    strSql = strSql & "EC.Pkid = " & txtPKId & " "
    strSql = strSql & "Order by dtmdtProcesso Desc "
    strSql = strSql & IIf(bytDBType = Oracle, "NULLS LAST ", "")
    strQueryHistoricoProcessoGrid = strSql

End Function

Private Function strQueryTributos() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT PKId, (" & gstrCONVERT(CDT_NVARCHAR, "intCodigo") & strCONCAT & " '  -  ' " & strCONCAT & " strDescricao) AS Inscricao "
    strSql = strSql & "FROM " & gstrComposicaoDaReceita & " "
    strSql = strSql & "WHERE intUtilizacao = 2 "
    strSql = strSql & "ORDER BY strDescricao"
    
    strQueryTributos = strSql
            
End Function

Private Sub PreencheGrdPublicidade(lngPkid As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "SELECT "
    strSql = strSql & "HPU.Pkid AS intHistorico, "
    strSql = strSql & "T.pkid AS intTributo, "
    strSql = strSql & "T.strDescricao Publicidade, "
    strSql = strSql & "HPU.intQuantidade Quantidade, "
    strSql = strSql & "HPU.dblArea Area, HPU.dtmPublicidadeInicio, HPU.dtmPublicidadeFim, "
    strSql = strSql & "HPU.strObservacao Observacao "
    strSql = strSql & "From "
    strSql = strSql & gstrTributo & " T, "
    strSql = strSql & gstrHistoricoPublicidades & " HPU "
    strSql = strSql & "Where "
    strSql = strSql & "T.PKID = HPU.INTTRIBUTO AND "
    strSql = strSql & "HPU.intEconomico = " & lngPkid
    
    Set gobjBanco = New clsBanco
    lvw_ItensPublicidade.ListItems.Clear
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                If .EOF = False Then
                    Do While Not .EOF
                        Set mobjLista = lvw_ItensPublicidade.ListItems.Add(, , gstrENulo(!intHistorico))
                        mobjLista.SubItems(1) = gstrENulo(!intTributo)
                        mobjLista.SubItems(2) = gstrENulo(!Publicidade)
                        mobjLista.SubItems(3) = gstrENulo(!Quantidade)
                        mobjLista.SubItems(4) = gstrConvVrDoSql(gstrENulo(!Area), 5)
                        mobjLista.SubItems(5) = gstrENulo(!Observacao)
                        mobjLista.SubItems(6) = gstrENulo(!dtmPublicidadeInicio)
                        mobjLista.SubItems(7) = gstrENulo(!dtmPublicidadeFim)
                        .MoveNext
                    Loop
                End If
            End With
        End If
    End If
End Sub

Private Function strQueryComboPublicidade() As String
    Dim strSql As String
    
    strSql = "SELECT T.Pkid, T.strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTributo & " T, "
    strSql = strSql & gstrTributoTipo & " TT "
    strSql = strSql & "Where "
    strSql = strSql & "T.INTTRIBUTOTIPO = TT.Pkid AND "
    strSql = strSql & "TT.bytTipo = " & TRIBUTO_TIPO_PUBLICIDADE
    strSql = strSql & " ORDER BY T.strDescricao"
    
    strQueryComboPublicidade = strSql
    
End Function

Private Function GravaPublicidades(lngPkid As Long) As Boolean
    Dim strSql              As String
    Dim strSql1             As String
    Dim strSql2             As String
    Dim intFor              As Integer
    Dim strPkidHistorico    As String
    
    GravaPublicidades = False

    strSql = ""
    strSql2 = ""
    strPkidHistorico = ""
    If lvw_ItensPublicidade.ListItems.Count <= 0 Then
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrHistoricoPublicidades
        strSql = strSql & " WHERE intEconomico = " & lngPkid
    Else
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
        
        For intFor = 1 To lvw_ItensPublicidade.ListItems.Count
            With lvw_ItensPublicidade
                If .ListItems(intFor).Text <> "" Then
                
                    strPkidHistorico = strPkidHistorico & .ListItems(intFor).Text & ","
                    
                    strSql = strSql & "UPDATE " & gstrHistoricoPublicidades
                    strSql = strSql & " SET inttributo = " & .ListItems(intFor).SubItems(1) & ", "
                    strSql = strSql & " intQuantidade = " & .ListItems(intFor).SubItems(3) & ", "
                    strSql = strSql & " dblArea = " & gstrConvVrParaSql(.ListItems(intFor).SubItems(4)) & ", "
                    strSql = strSql & " strObservacao = '" & .ListItems(intFor).SubItems(5) & "', "
                    If .ListItems(intFor).Selected And dbc_intTributoPublicidade.MatchedWithList Then
                        strSql = strSql & " dtmPublicidadeInicio = " & gstrConvDtParaSql(txt_dtmPublicidadeInicio) & ", "
                        strSql = strSql & " dtmPublicidadeFim = " & gstrConvDtParaSql(txt_dtmPublicidadeFim) & ", "
                    Else
                        strSql = strSql & " dtmPublicidadeInicio = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(6)) & ", "
                        strSql = strSql & " dtmPublicidadeFim = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(7)) & ", "
                    End If
                    strSql = strSql & " dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSql = strSql & " lngCodUsr = " & glngCodUsr
                    strSql = strSql & " WHERE Pkid = " & .ListItems(intFor).Text
                    strSql = strSql & IIf(bytDBType = Oracle, ";", "")
                Else
                    strSql2 = strSql2 & "INSERT INTO " & gstrHistoricoPublicidades
                    strSql2 = strSql2 & " (intEconomico,"
                    strSql2 = strSql2 & " inttributo,"
                    strSql2 = strSql2 & " intQuantidade,"
                    strSql2 = strSql2 & " dblArea,"
                    strSql2 = strSql2 & " strObservacao,"
                    strSql2 = strSql2 & " dtmPublicidadeInicio,"
                    strSql2 = strSql2 & " dtmPublicidadeFim,"
                    strSql2 = strSql2 & " dtmDtAtualizacao,"
                    strSql2 = strSql2 & " lngCodUsr)"
                    strSql2 = strSql2 & " VALUES( "
                    strSql2 = strSql2 & Val(lngPkid) & ", "
                    strSql2 = strSql2 & .ListItems(intFor).SubItems(1) & ", "
                    strSql2 = strSql2 & .ListItems(intFor).SubItems(3) & ", "
                    strSql2 = strSql2 & gstrConvVrParaSql(.ListItems(intFor).SubItems(4)) & ", "
                    strSql2 = strSql2 & "'" & .ListItems(intFor).SubItems(5) & "', "
                    strSql2 = strSql2 & gstrConvDtParaSql(.ListItems(intFor).SubItems(6)) & ", "
                    strSql2 = strSql2 & gstrConvDtParaSql(.ListItems(intFor).SubItems(7)) & ", "
                    strSql2 = strSql2 & strGETDATE & ", "
                    strSql2 = strSql2 & glngCodUsr
                    strSql2 = strSql2 & ")"
                    strSql2 = strSql2 & IIf(bytDBType = Oracle, ";", "")
                End If
            End With
        Next
        
        If strPkidHistorico <> "" Then
            strPkidHistorico = Mid(strPkidHistorico, 1, Len(strPkidHistorico) - 1)
            strSql1 = ""
            strSql1 = strSql1 & "DELETE FROM " & gstrHistoricoPublicidades
            strSql1 = strSql1 & " WHERE Pkid NOT in(" & strPkidHistorico & ")and"
            strSql1 = strSql1 & " intEconomico = " & lngPkid
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", " ")
            strSql = strSql & " " & strSql1
        Else
            strSql1 = strSql1 & "DELETE FROM " & gstrHistoricoPublicidades
            strSql1 = strSql1 & " WHERE intEconomico = " & lngPkid
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", "")
            strSql = strSql & " " & strSql1
        End If
        strSql = strSql & " " & strSql2
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    End If
    
    If gobjBanco.Execute(strSql) = False Then
        ExibeMensagem "Ocorreu um erro ao gravar as Publicidades. Os dados não foram gravados."
        GravaPublicidades = False
        Exit Function
    Else
        GravaPublicidades = True
    End If
End Function

Private Sub LimpaTabPubli(Optional blnLimpaGrid As Boolean)
    dbc_intTributoPublicidade.ListField = ""
    dbc_intTributoPublicidade.Text = ""
    txt_intQuantidade.Text = ""
    txt_dblArea.Text = ""
    txt_strObservacao.Text = ""
    txt_dtmPublicidadeInicio.Text = ""
    txt_dtmPublicidadeFim.Text = ""
    mblnAlterandoListaPubli = False
    If blnLimpaGrid Then
        lvw_ItensPublicidade.ListItems.Clear
    End If
End Sub

Private Sub PreencheListItens()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = strSql & "Select ECF.Intfeira, F.Strdescricao as strFeira, ECF.Inttipofeira, TF.Strdescricao as strTipoFeira, ECF.DBLAREA, ECF.STRNRBOX "
    strSql = strSql & "From "
    strSql = strSql & gstrEconomicoFeira & " ECF, "
    strSql = strSql & gstrEconomico & " EC, "
    strSql = strSql & gstrFeira & " F, "
    strSql = strSql & gstrTipoFeira & " TF "
    strSql = strSql & "Where "
    strSql = strSql & "EC.Pkid = ECF.Inteconomico AND "
    strSql = strSql & "F.Pkid = ECF.Intfeira AND "
    strSql = strSql & "TF.Pkid = ECF.Inttipofeira AND "
    strSql = strSql & "EC.Pkid = " & txtPKId
    strSql = strSql & " Order By strFeira"
    lvw_Itens.ListItems.Clear
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                Do While Not .EOF
                    Set mobjLista = lvw_Itens.ListItems.Add(, , gstrENulo(!Intfeira))
                    mobjLista.SubItems(1) = gstrENulo(!strFeira)
                    mobjLista.SubItems(2) = gstrENulo(!Inttipofeira)
                    mobjLista.SubItems(3) = gstrENulo(!strTipoFeira)
                    mobjLista.SubItems(4) = gstrConvVrDoSql(gstrENulo(!dblArea), 2)
                    mobjLista.SubItems(5) = gstrENulo(!STRNRBOX)
                    .MoveNext
                Loop
            End If
        End With
    End If
End Sub

Private Function ExcluirItemNoGrid(bytGuia As Byte) '1 - Feiras / 2 - Tributos / 3 - Publicidades / 4 - Sócios
    Dim intFor As Integer
    
    If bytGuia = 1 Then
        With lvw_Itens
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
        End With
        mblnAlterandoLista = False
    ElseIf bytGuia = 2 Then
        With lvw_ItensAtivTrib
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
        End With
        mblnAlterandoListaTributos = False
    ElseIf bytGuia = 3 Then
        With lvw_ItensPublicidade
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
        End With
        LimpaTabPubli False
        dbc_intTributoPublicidade.SetFocus
        mblnAlterandoListaPubli = False
    ElseIf bytGuia = 4 Then
        With lvw_Socios
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
        End With
        TotalizacaoCotas
        LimpaTabSocios False
        dbc_intSocios.SetFocus
    ElseIf bytGuia = 5 Then
        With lvw_Atividades
            If .ListItems.Count > 0 Then
                For intFor = 1 To lvw_ItensAtivTrib.ListItems.Count
                    If lvw_ItensAtivTrib.ListItems(intFor).Text = .ListItems(.SelectedItem.Index).SubItems(1) Then
                        ExibeMensagem "Atividade não pode ser excluída, pois a mesma se encontra relacionada nos tributos."
                        LimpaTabAtividade False
                        dbc_intAtividadePrincipal.SetFocus
                        mblnAlterandoListaAtividade = False
                        Exit Function
                    End If
                Next
                .ListItems.Remove .SelectedItem.Index
            End If
        End With
        LimpaTabAtividade False
        dbc_intAtividadePrincipal.SetFocus
        mblnAlterandoListaAtividade = False
    ElseIf bytGuia = 6 Then
        With lvw_ISS
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
        End With
        LimpaISS False
        dbc_intTipoISS.SetFocus
        mblnAlterandoListaISS = False
    End If
End Function

Private Function IncluirItemNoGrid(bytGuia As Byte) '1 - Feiras / 2 - Tributos / 3 - Publicidades / 4 - Socios / 5 - Atividades / 6 - ISS
    Dim intInd          As Integer
    
    If bytGuia = 1 Then
        If blnDadosItens = False Then Exit Function
        With lvw_Itens
            If mblnAlterandoLista Then
                For intInd = 1 To .ListItems.Count
                    If .SelectedItem.Index <> intInd Then
                        If Trim(dbc_intFeira.BoundText) = .ListItems(intInd).Text Then
                            ExibeMensagem "Não é possível incluir feiras iguais."
                            Exit Function
                        End If
                    End If
                Next
                .SelectedItem.Text = dbc_intFeira.BoundText
                .SelectedItem.SubItems(1) = dbc_intFeira.Text
                .SelectedItem.SubItems(2) = dbc_intTipoFeira.BoundText
                .SelectedItem.SubItems(3) = dbc_intTipoFeira.Text
                .SelectedItem.SubItems(4) = txt_areaFeira.Text
                .SelectedItem.SubItems(5) = txt_strnrbox
                mblnAlterandoLista = False
            Else
                For intInd = 1 To .ListItems.Count
                    If Trim(dbc_intFeira.BoundText) = .ListItems(intInd).Text Then
                        ExibeMensagem "Não é possível incluir feiras iguais."
                        Exit Function
                    End If
                Next
    
                Set mobjLista = .ListItems.Add(, , dbc_intFeira.BoundText)
                mobjLista.SubItems(1) = dbc_intFeira.Text
                mobjLista.SubItems(2) = dbc_intTipoFeira.BoundText
                mobjLista.SubItems(3) = dbc_intTipoFeira.Text
                mobjLista.SubItems(4) = txt_areaFeira.Text
                mobjLista.SubItems(5) = txt_strnrbox
            End If
        End With
        LimpaTabFeira
    ElseIf bytGuia = 2 Then
        If blnDadosItensAtiv = False Then Exit Function
        
        With lvw_ItensAtivTrib
            If mblnAlterandoListaTributos Then
                For intInd = 1 To .ListItems.Count
                    If .SelectedItem.Index <> intInd Then
                        If Trim(dbc_intAtividades.BoundText) = .ListItems(intInd).Text And Trim(dbc_intTributo.BoundText) = .ListItems(intInd).SubItems(2) Then
                            ExibeMensagem "Não é possível incluir Atividades iguais."
                            dbc_intAtividades.Text = ""
                            dbc_intTipoTributo.Text = ""
                            dbc_intTributo.Text = ""
                            txt_intQtd.Text = ""
                            Exit Function
                        End If
                    End If
                Next
                .SelectedItem.Text = dbc_intAtividades.BoundText
                .SelectedItem.SubItems(1) = dbc_intAtividades.Text
                .SelectedItem.SubItems(2) = dbc_intTributo.BoundText
                .SelectedItem.SubItems(3) = dbc_intTributo.Text
                .SelectedItem.SubItems(4) = IIf(Trim(txt_intQtd.Text) = "", 1, txt_intQtd)
                mblnAlterandoListaTributos = False
            Else
                For intInd = 1 To .ListItems.Count
                    If Trim(dbc_intAtividades.BoundText) = .ListItems(intInd).Text And Trim(dbc_intTributo.BoundText) = .ListItems(intInd).SubItems(2) Then
                        ExibeMensagem "Não é possível incluir Atividades com Tributos iguais."
                        dbc_intAtividades.Text = ""
                        dbc_intTipoTributo.Text = ""
                        dbc_intTributo.Text = ""
                        Exit Function
                    End If
                Next
                Set mobjLista = .ListItems.Add(, , dbc_intAtividades.BoundText)
                mobjLista.SubItems(1) = dbc_intAtividades.Text
                mobjLista.SubItems(2) = dbc_intTributo.BoundText
                mobjLista.SubItems(3) = dbc_intTributo.Text
                mobjLista.SubItems(4) = IIf(Trim(txt_intQtd.Text) = "", 1, txt_intQtd)
            End If
        End With
        LimpaTabTributo
        
    ElseIf bytGuia = 3 Then
        If blnDadosItenPublicidade = False Then Exit Function
                    
        With lvw_ItensPublicidade
            If mblnAlterandoListaPubli Then
                .SelectedItem.SubItems(1) = dbc_intTributoPublicidade.BoundText
                .SelectedItem.SubItems(2) = dbc_intTributoPublicidade.Text
                .SelectedItem.SubItems(3) = IIf(Trim(txt_intQuantidade.Text) = "", "1", IIf(CDbl(Val(gstrConvVrParaSql(txt_intQuantidade))) = 0, 1, txt_intQuantidade))
                .SelectedItem.SubItems(4) = IIf(Trim(txt_dblArea.Text) = "", "1", IIf(CDbl(Val(gstrConvVrParaSql(txt_dblArea.Text))) = 0, 1, txt_dblArea))
                .SelectedItem.SubItems(5) = txt_strObservacao.Text
                .SelectedItem.SubItems(6) = txt_dtmPublicidadeInicio.Text
                .SelectedItem.SubItems(7) = txt_dtmPublicidadeFim.Text
                mblnAlterandoListaPubli = False
                LimpaTabPubli False
                dbc_intTributoPublicidade.SetFocus
            Else
                Set mobjLista = .ListItems.Add(, , "")
                mobjLista.SubItems(1) = dbc_intTributoPublicidade.BoundText
                mobjLista.SubItems(2) = dbc_intTributoPublicidade.Text
                mobjLista.SubItems(3) = IIf(Trim(txt_intQuantidade.Text) = "", "1", IIf(CDbl(Val(gstrConvVrParaSql(txt_intQuantidade))) = 0, 1, txt_intQuantidade))
                mobjLista.SubItems(4) = IIf(Trim(txt_dblArea.Text) = "", "1", IIf(CDbl(Val(gstrConvVrParaSql(txt_dblArea.Text))) = 0, 1, txt_dblArea))
                mobjLista.SubItems(5) = txt_strObservacao.Text
                mobjLista.SubItems(6) = txt_dtmPublicidadeInicio.Text
                mobjLista.SubItems(7) = txt_dtmPublicidadeFim.Text
                LimpaTabPubli False
                dbc_intTributoPublicidade.SetFocus
            End If
        End With
    ElseIf bytGuia = 4 Then
        If blnDadoSocios = False Then Exit Function
                    
        With lvw_Socios
            If mblnAlterandoListaSocios Then
                .SelectedItem.SubItems(1) = dbc_intSocios.BoundText
                .SelectedItem.SubItems(2) = dbc_intSocios.Text
                .SelectedItem.SubItems(3) = txt_strCNPJCPF.Text
                .SelectedItem.SubItems(4) = txt_strCotas.Text
                .SelectedItem.SubItems(5) = txt_dtmSocioInicio.Text
                .SelectedItem.SubItems(6) = txt_dtmSocioFim.Text
                mblnAlterandoListaSocios = False
                LimpaTabSocios False
                dbc_intSocios.SetFocus
            Else
                Set mobjLista = .ListItems.Add(, , "")
                mobjLista.SubItems(1) = dbc_intSocios.BoundText
                mobjLista.SubItems(2) = dbc_intSocios.Text
                mobjLista.SubItems(3) = txt_strCNPJCPF.Text
                mobjLista.SubItems(4) = txt_strCotas.Text
                mobjLista.SubItems(5) = txt_dtmSocioInicio.Text
                mobjLista.SubItems(6) = txt_dtmSocioFim.Text
                LimpaTabSocios False
                dbc_intSocios.SetFocus
            End If
        End With
        TotalizacaoCotas
    ElseIf bytGuia = 5 Then
        blnOkAtividade = False
        If blnDadoAtividades = False Then Exit Function
        
        With lvw_Atividades
            If mblnAlterandoListaAtividade Then
                .SelectedItem.SubItems(1) = dbc_intAtividadePrincipal.BoundText
                .SelectedItem.SubItems(2) = Abs(CInt(CBool(chk_blnPrincipal.Value)))
                .SelectedItem.SubItems(3) = IIf(Val(chk_blnPrincipal.Value) = 1, "Principal", "Secundária")
                .SelectedItem.SubItems(4) = dbc_intAtividadePrincipal.Text
                .SelectedItem.SubItems(5) = txt_dtmAtividadeInicio.Text
                .SelectedItem.SubItems(6) = txt_dtmAtividadeFim.Text
                mblnAlterandoListaAtividade = False
                LimpaTabAtividade False
                dbc_intAtividadePrincipal.SetFocus
            Else
                Set mobjLista = .ListItems.Add(, , "")
                mobjLista.SubItems(1) = dbc_intAtividadePrincipal.BoundText
                mobjLista.SubItems(2) = Val(chk_blnPrincipal.Value)
                mobjLista.SubItems(3) = IIf(Val(chk_blnPrincipal.Value) = 1, "Principal", "Secundária")
                mobjLista.SubItems(4) = dbc_intAtividadePrincipal.Text
                mobjLista.SubItems(5) = txt_dtmAtividadeInicio.Text
                mobjLista.SubItems(6) = txt_dtmAtividadeFim.Text
                LimpaTabAtividade False
                dbc_intAtividadePrincipal.SetFocus
            End If
        End With
        blnOkAtividade = True
    ElseIf bytGuia = 6 Then
        If blnDadoISS = False Then Exit Function
        
        With lvw_ISS
            If mblnAlterandoListaISS Then
                .SelectedItem.SubItems(1) = dbc_intTipoISS.BoundText
                .SelectedItem.SubItems(2) = Trim(dbc_intTipoISS.Text)
                .SelectedItem.SubItems(3) = dbc_intListaServico.BoundText
                .SelectedItem.SubItems(4) = Trim(dbc_intListaServico.Text)
                .SelectedItem.SubItems(5) = gstrDataFormatada(txt_dtmissinicio.Text)
                .SelectedItem.SubItems(6) = gstrDataFormatada(txt_dtmissfim.Text)
                .SelectedItem.SubItems(7) = IIf(Trim(txt_intQuantidadeIss.Text) = "", "1", txt_intQuantidadeIss)
                mblnAlterandoListaISS = False
                LimpaISS False
                dbc_intTipoISS.SetFocus
            Else
                Set mobjLista = .ListItems.Add(, , "")
                mobjLista.SubItems(1) = dbc_intTipoISS.BoundText
                mobjLista.SubItems(2) = Trim(dbc_intTipoISS.Text)
                mobjLista.SubItems(3) = dbc_intListaServico.BoundText
                mobjLista.SubItems(4) = Trim(dbc_intListaServico.Text)
                mobjLista.SubItems(5) = gstrDataFormatada(txt_dtmissinicio.Text)
                mobjLista.SubItems(6) = gstrDataFormatada(txt_dtmissfim.Text)
                mobjLista.SubItems(7) = IIf(Trim(txt_intQuantidadeIss.Text) = "", "1", txt_intQuantidadeIss)
                LimpaISS False
                dbc_intTipoISS.SetFocus
            End If
        End With
    End If
End Function

Private Function blnDadosItens() As Boolean

    blnDadosItens = False
    
    If dbc_intFeira.MatchedWithList = False Then
        ExibeMensagem "O campo feira deve ser preenchido corretamente."
        dbc_intFeira.SetFocus
        Exit Function
    ElseIf dbc_intTipoFeira.MatchedWithList = False Then
        ExibeMensagem "O campo tipo de feira deve ser preenchido corretamente."
        dbc_intTipoFeira.SetFocus
        Exit Function
    End If
    
    blnDadosItens = True
    
End Function

Private Function StrSalvaItem(intPKIdEconomico As Long) As Boolean
    Dim strSql  As String
    Dim intInd  As Integer
    
    StrSalvaItem = False
    'Set gobjBanco = New clsBanco
    
    strSql = ""
    If lvw_Itens.ListItems.Count > 0 Then
        strSql = IIf(bytDBType = Oracle, "Begin", "")
    End If
    
    If blnAlterando Then
        strSql = strSql & " Delete from " & gstrEconomicoFeira & " Where intEconomico = " & intPKIdEconomico
        If lvw_Itens.ListItems.Count > 0 Then
            strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        End If
    End If
    
    If lvw_Itens.ListItems.Count > 0 Then
        With lvw_Itens
            For intInd = 1 To .ListItems.Count
                strSql = strSql & " INSERT INTO "
                strSql = strSql & gstrEconomicoFeira & " ("
                strSql = strSql & "INTECONOMICO, "
                strSql = strSql & "INTFEIRA, "
                strSql = strSql & "INTTIPOFEIRA, "
                strSql = strSql & "DBLAREA, "
                strSql = strSql & "STRNRBOX, "
                strSql = strSql & "dtmDtAtualizacao, "
                strSql = strSql & "lngCodUsr) "
                strSql = strSql & "Values("
                strSql = strSql & intPKIdEconomico & ", "
                strSql = strSql & .ListItems(intInd).Text & ", "
                strSql = strSql & .ListItems(intInd).SubItems(2) & ", "
                strSql = strSql & IIf(.ListItems(intInd).SubItems(4) <> "", gstrConvVrParaSql(.ListItems(intInd).SubItems(4)), "Null") & ", "
                strSql = strSql & IIf(.ListItems(intInd).SubItems(5) <> "", .ListItems(intInd).SubItems(5), "Null") & ", "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSql = strSql & glngCodUsr & " "
                strSql = strSql & ")" & IIf(bytDBType = Oracle, ";", "")
            Next
        End With
    End If
    If lvw_Itens.ListItems.Count > 0 Then
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    End If
    
    If strSql = "" Then
       StrSalvaItem = True
       Exit Function
    End If
    
    If gobjBanco.Execute(strSql) = False Then
        ExibeMensagem "Ocorreu um erro ao gravar as feiras. Os dados não foram gravados."
        StrSalvaItem = False
        Exit Function
    Else
        StrSalvaItem = True
    End If
    
End Function

Private Function strQueryFeira() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "Select Pkid, strDescricao From " & gstrFeira & " Order by strDescricao "
    strQueryFeira = strSql
End Function

Private Function strQueryHorarioFuncionamento() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "Select Pkid, strDescricao From " & "tblHorarioFuncionamento" & " Order by strDescricao "
    strQueryHorarioFuncionamento = strSql
End Function

Private Function strQueryTipoFeira() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "Select Pkid, strDescricao From " & gstrTipoFeira & " Order by strDescricao "
    strQueryTipoFeira = strSql
End Function

Sub LimpaTabFeira()
    dbc_intFeira.Text = ""
    dbc_intTipoFeira.Text = ""
    txt_areaFeira.Text = ""
    txt_strnrbox.Text = ""
End Sub

Private Function strQueryRelatorio() As String
    Dim strSql As String
    
    strQueryRelatorio = strSql
End Function

Private Function blnDadosItensAtiv() As Boolean
    blnDadosItensAtiv = False
        
    If Not dbc_intTipoTributo.MatchedWithList And Not dbc_intTributo.MatchedWithList Then Exit Function
        
    If dbc_intAtividades.MatchedWithList = False Then
        ExibeMensagem "O campo Atividade deve ser preenchido corretamente."
        dbc_intAtividades.SetFocus
        Exit Function
    ElseIf dbc_intTipoTributo.MatchedWithList = False Then
        ExibeMensagem "O campo Tipo de Tributo deve ser preenchido corretamente."
        dbc_intTipoTributo.SetFocus
        Exit Function
    ElseIf dbc_intTributo.MatchedWithList = False Then
        ExibeMensagem "O campo Tributo deve ser preenchido corretamente."
        dbc_intTributo.SetFocus
        Exit Function
    End If
    blnDadosItensAtiv = True
End Function

Sub LimpaTabTributo()
    Set dbc_intAtividades.RowSource = Nothing
    dbc_intAtividades.Text = ""
    Set dbc_intTipoTributo.RowSource = Nothing
    dbc_intTipoTributo.Text = ""
    Set dbc_intTributo.RowSource = Nothing
    dbc_intTributo.Text = ""
    txt_intQtd.Text = "1"
End Sub

Private Function strQueryTipoTributo() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "Select Pkid, Ltrim(Rtrim(strDescricao)) strDescricao From "
    strSql = strSql & gstrTributoTipo & " "
    strSql = strSql & "Where "
    strSql = strSql & "bytTipo in (" & TRIBUTO_TIPO_OUTROS & ", " & TRIBUTO_TIPO_HORARIO_ESPECIAL & ") "
    strSql = strSql & " Order by strDescricao"
    strQueryTipoTributo = strSql
End Function

Private Sub LimpaTabISS()
    Set dbc_intAtividades.RowSource = Nothing
    Set dbc_intTipoTributo.RowSource = Nothing
    Set dbc_intTributo.RowSource = Nothing
    dbc_intAtividades.Text = ""
    dbc_intTipoTributo.Text = ""
    dbc_intTributo.Text = ""
    mblnAlterandoListaTributos = False
    lvw_ItensAtivTrib.ListItems.Clear
End Sub

Private Function SalvaItensAtividadesTributo(intPKIdEconomico As Long, blnAlterando As Boolean) As Boolean
    Dim strSql  As String
    Dim strSqlAux  As String
    Dim intInd  As Integer
    Dim adoResultado As ADODB.Recordset
    
    SalvaItensAtividadesTributo = False
    
    Set gobjBanco = New clsBanco
    
    strSql = ""
    If lvw_ItensAtivTrib.ListItems.Count > 0 Then
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
    End If
    
    If blnAlterando Then
        strSql = strSql & "Delete "
        strSql = strSql & "From "
        strSql = strSql & gstrAtivEmpresaTributo & " "
        strSql = strSql & "Where "
        strSql = strSql & "Pkid in(Select AET.pkid From "
        strSql = strSql & gstrAtivEmpresaTributo & " AET, "
        strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
        strSql = strSql & gstrEconomico & " E "
        strSql = strSql & "Where E.Pkid = AE.intEconomico And AE.Pkid = AET.INTATIVIDADEDAEMPRESA And "
        strSql = strSql & "E.Pkid = " & intPKIdEconomico & ")"
        If lvw_ItensAtivTrib.ListItems.Count > 0 Then
            strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        End If
    End If
    
    If lvw_ItensAtivTrib.ListItems.Count > 0 Then
        With lvw_ItensAtivTrib
            For intInd = 1 To .ListItems.Count
                strSql = strSql & " INSERT INTO "
                strSql = strSql & gstrAtivEmpresaTributo & " ("
                strSql = strSql & "intatividadedaempresa, "
                strSql = strSql & "inttributo, "
                strSql = strSql & "dtmDtAtualizacao, "
                strSql = strSql & "lngCodUsr, "
                strSql = strSql & "intQtd) "
                strSql = strSql & "Values("
                strSqlAux = "" 'Vamos achar qual o ID da TblatividadeDaEmpresa
                strSqlAux = strSqlAux & "Select "
                strSqlAux = strSqlAux & "A.Pkid "
                strSqlAux = strSqlAux & "From "
                strSqlAux = strSqlAux & gstrAtividadeDaEmpresa & " A, "
                strSqlAux = strSqlAux & gstrEconomico & " B "
                strSqlAux = strSqlAux & "Where "
                strSqlAux = strSqlAux & "B.Pkid = A.Inteconomico AND "
                strSqlAux = strSqlAux & "intatividade = " & .ListItems(intInd).Text & " AND "
                strSqlAux = strSqlAux & "B.Pkid = " & intPKIdEconomico
                If gobjBanco.CriaADO(strSqlAux, 5, adoResultado) Then
                    strSql = strSql & Val(gstrENulo(adoResultado!Pkid)) & ", "
                End If
                strSql = strSql & .ListItems(intInd).SubItems(2) & ", "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSql = strSql & glngCodUsr & ", "
                strSql = strSql & IIf(Val(gstrENulo(.ListItems(intInd).SubItems(4))) = 0, 1, gstrENulo(.ListItems(intInd).SubItems(4))) & " "
                strSql = strSql & ")" & IIf(bytDBType = Oracle, ";", "")
            Next
        End With
    End If
    
    If lvw_ItensAtivTrib.ListItems.Count > 0 Then
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    End If
    If Trim(strSql) <> "" Then
        If gobjBanco.Execute(strSql) = False Then
            ExibeMensagem "Ocorreu um erro ao gravar os tributos. Os dados não foram gravados."
            SalvaItensAtividadesTributo = False
            Exit Function
        Else
            SalvaItensAtividadesTributo = True
        End If
    Else
        SalvaItensAtividadesTributo = True
    End If
    
End Function

Private Sub PreencheListaAtividadeTributo(intEconomico As Long)
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "AEC.Pkid AS IDAtividadeEC, "
    strSql = strSql & "AEC.Strdescricao AS strAtividade, "
    strSql = strSql & "AET.intQtd AS intQtd, "
    strSql = strSql & "T.Pkid AS IDtributo, "
    strSql = strSql & "T.strDescricao strTributo "
    strSql = strSql & "From "
    strSql = strSql & gstrEconomico & " E, "
    strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
    strSql = strSql & gstrAtivEmpresaTributo & " AET, "
    strSql = strSql & gstrTributo & " T, "
    strSql = strSql & gstrAtividadeEC & " AEC "
    strSql = strSql & "Where "
    strSql = strSql & "E.PKID = AE.Inteconomico AND "
    strSql = strSql & "AEC.Pkid = AE.INTATIVIDADE AND "
    strSql = strSql & "AE.Pkid = AET.INTATIVIDADEDAEMPRESA AND "
    strSql = strSql & "T.Pkid = AET.INTTRIBUTO AND "
    strSql = strSql & "E.Pkid = " & intEconomico
    
    lvw_ItensAtivTrib.ListItems.Clear
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                Do While Not .EOF
                    Set mobjLista = lvw_ItensAtivTrib.ListItems.Add(, , gstrENulo(!IDAtividadeEC))
                    mobjLista.SubItems(1) = gstrENulo(!strAtividade)
                    mobjLista.SubItems(2) = gstrENulo(!IDtributo)
                    mobjLista.SubItems(3) = gstrENulo(!strTributo)
                    mobjLista.SubItems(4) = gstrENulo(!intQtd)
                    .MoveNext
                Loop
            End If
        End With
    End If
End Sub

Private Function blnInscricaoImobiliaria() As Boolean
    Dim strSql As String
    Dim adoResultado    As ADODB.Recordset
    
    blnInscricaoImobiliaria = True
    If Trim(mskstrInscricaoImobiliaria.Text) <> "" Then
        strSql = ""
        strSql = strSql & "Select " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao "
        strSql = strSql & "FROM " & gstrImobiliario & " Where strInscricao = '" & String(gintLenInscricao - Len(mskstrInscricaoImobiliaria.Text), "0") & mskstrInscricaoImobiliaria.Text & "'"
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                If .EOF Then
                   ExibeMensagem "Inscrição Imobiliária inválida."
                   mskstrInscricaoImobiliaria.SetFocus
                    blnInscricaoImobiliaria = False
                End If
            End With
        End If
    End If
End Function
    
Private Function blnDadosItenPublicidade() As Boolean

    blnDadosItenPublicidade = False
    
    If Trim(txt_dtmPublicidadeInicio) <> "" Then
        If Not gblnDataValida(txt_dtmPublicidadeInicio) Then
            ExibeMensagem "Data de início da Publicidade inválida."
            txt_dtmPublicidadeInicio.SetFocus
            Exit Function
        End If
    Else
        ExibeMensagem "Data de início da Publicidade é obrigatória."
        txt_dtmPublicidadeInicio.SetFocus
        Exit Function
    End If
    
    If Len(txt_dtmPublicidadeFim) > 0 And Len(txt_dtmPublicidadeInicio) > 0 Then
        If CDate(txt_dtmPublicidadeFim) < CDate(txt_dtmPublicidadeInicio) Then
            ExibeMensagem "A data de término não pode ser menor que a data de início."
            txt_dtmPublicidadeFim.SetFocus
            Exit Function
        End If
    End If
    
    If Not dbc_intTributoPublicidade.MatchedWithList Then
        ExibeMensagem "É necessário informar um Tipo de Publicidade válido."
        dbc_intTributoPublicidade.SetFocus
        Exit Function
    ElseIf txt_intQuantidade.Text = "" Then
        ExibeMensagem "É necessário informar uma Quantidade."
        txt_intQuantidade.SetFocus
        Exit Function
    ElseIf txt_dblArea.Text = "" Then
        ExibeMensagem "É necessário informar uma Área."
        txt_dblArea.SetFocus
        Exit Function
    End If
    
    blnDadosItenPublicidade = True
End Function

Private Function strQueryEconomico() As String
    Dim strSql As String

    strSql = strSql & "SELECT EC.*, "
    strSql = strSql & gstrRIGHT("EC.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricaoCadastral, "
    strSql = strSql & gstrRIGHT("EC.strInscricaoImobiliaria", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricaoImobiliaria "
    strSql = strSql & "FROM " & gstrEconomico & " EC "
    strSql = strSql & "WHERE Pkid = " & txtPKId.Text
    
    strQueryEconomico = strSql
    
End Function

Private Sub PreencheProcessos()
Dim adoProcesso As ADODB.Recordset
Dim strSql As String

    'Rafael 30/09/2004
    'Essa função foi desenvolvida, pois quando trazia os digitos dos processos,
    'se fosse nulo, ele colocava zero (isso usando LeDaTabelaParaObj)
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "strcodprocabertura, intexerprocabertura, bitdigprocabertura, "
    strSql = strSql & "strcodprocencerramento, intexerprocencerramento, bitdigprocencerramento, "
    strSql = strSql & "strcodprocesso, intexerprocesso, bitdigprocesso "
    strSql = strSql & "FROM " & gstrEconomico & " "
    strSql = strSql & "WHERE pkID = " & txtPKId.Text
    
    Set adoProcesso = New ADODB.Recordset
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoProcesso) Then
       With adoProcesso
         txtstrCodProcAbertura.Text = gstrENulo(!strcodprocabertura)
         txtintExerProcAbertura.Text = gstrENulo(!intexerprocabertura)
         txtbitDigProcAbertura.Text = gstrENulo(!bitdigprocabertura)
        
         txtstrCodProcEncerramento.Text = gstrENulo(!strcodprocencerramento)
         txtintExerProcEncerramento.Text = gstrENulo(!intexerprocencerramento)
         txtbitDigProcEncerramento.Text = gstrENulo(!bitdigprocencerramento)
         
         txtstrcodprocesso.Text = gstrENulo(!strCodProcesso)
         txtintexerprocesso.Text = gstrENulo(!intExerProcesso)
         txtbitdigprocesso.Text = gstrENulo(!bitDigProcesso)
         
       End With
    End If
    Set adoProcesso = Nothing
    
End Sub

Private Sub LimpaProcessos()
    txtstrCodProcAbertura.Text = ""
    txtintExerProcAbertura.Text = ""
    txtbitDigProcAbertura.Text = ""
   
    txtstrCodProcEncerramento.Text = ""
    txtintExerProcEncerramento.Text = ""
    txtbitDigProcEncerramento.Text = ""
End Sub

Private Sub PreencheComboHistorico()
    cbo_intTipo.AddItem "Razão Social"
    cbo_intTipo.ItemData(cbo_intTipo.NewIndex) = 1
    
    cbo_intTipo.AddItem "Endereço"
    cbo_intTipo.ItemData(cbo_intTipo.NewIndex) = 2
    
    cbo_intTipo.AddItem "Atividade"
    cbo_intTipo.ItemData(cbo_intTipo.NewIndex) = 3
    
    cbo_intTipo.AddItem "Sócios"
    cbo_intTipo.ItemData(cbo_intTipo.NewIndex) = 4
    
    cbo_intTipo.AddItem "Ocorrência"
    cbo_intTipo.ItemData(cbo_intTipo.NewIndex) = 5
    
    cbo_intTipo.AddItem "Publicidade"
    cbo_intTipo.ItemData(cbo_intTipo.NewIndex) = 6

    cbo_intTipo.AddItem "ISSQN"
    cbo_intTipo.ItemData(cbo_intTipo.NewIndex) = 7

End Sub

Private Sub PreencheGridHistoricoOcorrencias()
    Dim strSql As String
    
    strSql = strSql & "Select "
    strSql = strSql & "Case When HEV.Byttipohistorico = 1 Then 'Razão Social' "
    strSql = strSql & "When HEV.Byttipohistorico = 2 Then 'Endereço' "
    strSql = strSql & "When HEV.Byttipohistorico = 3 Then 'Atividade' "
    strSql = strSql & "When HEV.Byttipohistorico = 4 Then 'Sócio' "
    strSql = strSql & "When HEV.Byttipohistorico = 5 Then 'Ocorrêcia' "
    strSql = strSql & "When HEV.Byttipohistorico = 6 Then 'Publicidade' "
    strSql = strSql & "When HEV.Byttipohistorico = 7 Then 'ISSQN' End as strTipo, "
    strSql = strSql & "HEV.Dtmdtinicial, "
    strSql = strSql & "HEV.Dtmdtfinal, "
    strSql = strSql & "HEV.STRDESCRICAO as StrOcorrencia "
    strSql = strSql & "From "
    strSql = strSql & gstrEconomico & " EC, "
    strSql = strSql & gstrHistoricoEconVariavel & " HEV "
    strSql = strSql & "Where "
    strSql = strSql & "EC.Pkid = HEV.Inteconomico AND "
    strSql = strSql & "HEV.Byttipohistorico = " & gstrItemData(cbo_intTipo, False) & " AND "
    strSql = strSql & "EC.Pkid = " & Val(txtPKId)
    strSql = strSql & " Order By HEV.dtmdtInicial Desc, HEV.Pkid Desc "

    LeDaTabelaParaObj "", tdb_HistOcorrencias, strSql
End Sub
Private Function strQuerySocios() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT SO.PKId, LTrim(RTrim(CO.strNome)) as strNome "
    strSql = strSql & "FROM " & gstrContribuinte & " CO, " & gstrSocio & " SO "
    strSql = strSql & "WHERE SO.intContribuinte = CO.PKId "
    strSql = strSql & "ORDER BY CO.strNome"
    
    strQuerySocios = strSql
End Function

Private Function blnDadoSocios() As Boolean
   blnDadoSocios = False
    
    If Trim(txt_dtmSocioInicio) <> "" Then
        If Not gblnDataValida(txt_dtmSocioInicio) Then
            ExibeMensagem "Data de início do Sócio inválida."
            txt_dtmSocioInicio.SetFocus
            Exit Function
        End If
    Else
        ExibeMensagem "Data de início do Sócio é obrigatória."
        txt_dtmSocioInicio.SetFocus
        Exit Function
    End If
    
    If Len(txt_dtmSocioFim) > 0 And Len(txt_dtmSocioInicio) > 0 Then
        If CDate(txt_dtmSocioFim) < CDate(txt_dtmSocioInicio) Then
            ExibeMensagem "A data de término não pode ser menor que a data de início."
            txt_dtmSocioFim.SetFocus
            Exit Function
        End If
    End If
    
    If Not dbc_intSocios.MatchedWithList Then
        ExibeMensagem "É necessário informar um sócio válido."
        dbc_intSocios.SetFocus
        Exit Function
    End If
    
    blnDadoSocios = True

End Function

Private Sub PreencheGrdSocios(lngPkid As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "SELECT "
    strSql = strSql & "SE.Pkid, "
    strSql = strSql & "SE.IntSocio, SE.dtmSocioInicio, SE.dtmSocioFim, "
    strSql = strSql & "CO.strNome, "
    strSql = strSql & "CO.StrCnpjCpf, "
    strSql = strSql & gstrISNULL("SE.intNumeroDeCotas", "0") & " Cotas "
    strSql = strSql & "FROM "
    strSql = strSql & gstrSocioEconomico & " SE, "
    strSql = strSql & gstrSocio & " SO, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & "WHERE "
    strSql = strSql & "SE.intSocio = SO.PKId AND "
    strSql = strSql & "SO.intContribuinte = CO.PKID AND "
    strSql = strSql & "SE.intCodEconomico = " & Val(txtPKId)
    
    Set gobjBanco = New clsBanco
    lvw_Socios.ListItems.Clear
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                If .EOF = False Then
                    Do While Not .EOF
                        Set mobjLista = lvw_Socios.ListItems.Add(, , gstrENulo(!Pkid))
                        mobjLista.SubItems(1) = gstrENulo(!intSocio)
                        mobjLista.SubItems(2) = gstrENulo(!strNome)
                        mobjLista.SubItems(3) = gstrCGCCPFFormatado(gstrENulo(!StrCnpjCpf))
                        mobjLista.SubItems(4) = gstrConvVrDoSql(gstrENulo(!cotas), 2)
                        mobjLista.SubItems(5) = gstrENulo(!dtmSocioInicio)
                        mobjLista.SubItems(6) = gstrENulo(!dtmSocioFim)
                        .MoveNext
                    Loop
                End If
            End With
        End If
    End If
    TotalizacaoCotas
End Sub

Private Function gstrPreencherCNPJCPF(lngPkid As Long) As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "Select "
    strSql = strSql & "CO.Strcnpjcpf "
    strSql = strSql & "From "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrSocio & " SO "
    strSql = strSql & "Where "
    strSql = strSql & "CO.Pkid = SO.Intcontribuinte AND "
    strSql = strSql & "SO.Pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            gstrPreencherCNPJCPF = gstrCGCCPFFormatado(gstrENulo(adoResultado!StrCnpjCpf))
        Else
            gstrPreencherCNPJCPF = ""
        End If
    End If
End Function

Private Sub LimpaTabSocios(Optional blnLimpaGrid As Boolean)
    dbc_intSocios.ListField = ""
    dbc_intSocios.Text = ""
    txt_strCNPJCPF.Text = ""
    txt_strCotas.Text = ""
    txt_dtmSocioInicio.Text = ""
    txt_dtmSocioFim.Text = ""
    mblnAlterandoListaSocios = False
    If blnLimpaGrid Then
        lvw_Socios.ListItems.Clear
    End If
End Sub

Private Sub TotalizacaoCotas()
    Dim intFor      As Integer
    Dim dblTotal    As Double
    
    dblTotal = 0
    
    With lvw_Socios
        For intFor = 1 To .ListItems.Count
            dblTotal = dblTotal + Val(.ListItems(intFor).SubItems(4))
        Next
    End With
    
    txt_TotalDeCotas = gstrConvVrDoSql(dblTotal)
        
End Sub

Private Function strQueryOcorrenciaProcesso() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Select Pkid, Ltrim(Rtrim(strdescricao)) as strdescricao From " & gstrOcorrenciaDoEconomico & " Order By strdescricao"
    
    strQueryOcorrenciaProcesso = strSql
    
End Function

Private Function strQueryAtividade() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT AEC.PKId, " & gstrCONVERT(CDT_NVARCHAR, "AEC.intCodigo") & strCONCAT & " ' - ' " & strCONCAT & " Ltrim(Rtrim(AEC.strDescricao)) " & strCONCAT & " ' / ' " & strCONCAT & " Ltrim(Rtrim(SA.STRNOMEDOSUBGRUPO)) AS strDescricao "
    strSql = strSql & "FROM " & gstrAtividadeEC & " AEC, "
    strSql = strSql & gstrSubGrupoDeAtividade & " SA "
    strSql = strSql & "Where Sa.Pkid = AEC.intSubGrupo "
    strSql = strSql & "ORDER BY AEC.strDescricao"
    
    strQueryAtividade = strSql
End Function

Private Sub PreencheGrdAtividades(lngPkid As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = "SELECT "
    strSql = strSql & "AE.Pkid PkidAtivEmpresa, "
    strSql = strSql & "AEC.Pkid PkidAtividade, "
    strSql = strSql & "AE.BLNPRINCIPAL, AE.dtmAtividadeInicio, AE.dtmAtividadeFim, "
    strSql = strSql & "Case When AE.BLNPRINCIPAL = 1 Then 'Principal' Else 'Secundária' end Status, "
    strSql = strSql & gstrCONVERT(CDT_NVARCHAR, "AEC.intCodigo") & strCONCAT & " ' - ' " & strCONCAT & " Ltrim(Rtrim(AEC.strDescricao)) " & strCONCAT & " ' / ' " & strCONCAT & " Ltrim(Rtrim(SA.STRNOMEDOSUBGRUPO)) AS strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
    strSql = strSql & gstrSubGrupoDeAtividade & " SA, "
    strSql = strSql & gstrAtividadeEC & " AEC "
    strSql = strSql & "Where "
    strSql = strSql & "AE.Intatividade = AEC.Pkid AND "
    strSql = strSql & "Sa.Pkid = AEC.intSubGrupo AND "
    strSql = strSql & "AE.Inteconomico = " & lngPkid
    strSql = strSql & " ORDER BY "
    strSql = strSql & "AEC.intCodigo "
    
    Set gobjBanco = New clsBanco
    lvw_Atividades.ListItems.Clear
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                If .EOF = False Then
                    Do While Not .EOF
                        Set mobjLista = lvw_Atividades.ListItems.Add(, , gstrENulo(!PkidAtivEmpresa))
                        mobjLista.SubItems(1) = gstrENulo(!PkidAtividade)
                        mobjLista.SubItems(2) = gstrENulo(!blnPrincipal)
                        mobjLista.SubItems(3) = gstrENulo(!Status)
                        mobjLista.SubItems(4) = gstrENulo(!strDescricao)
                        mobjLista.SubItems(5) = gstrDataFormatada(gstrENulo(!dtmAtividadeInicio))
                        mobjLista.SubItems(6) = gstrDataFormatada(gstrENulo(!dtmAtividadeFim))
                        .MoveNext
                    Loop
                End If
            End With
        End If
    End If
    
End Sub

Private Function blnDadoAtividades() As Boolean
    Dim intFor As Integer
    
    blnDadoAtividades = False
    
    If Trim(txt_dtmAtividadeInicio) <> "" Then
        If Not gblnDataValida(txt_dtmAtividadeInicio) Then
            ExibeMensagem "Data de início da Atividade inválida."
            txt_dtmAtividadeInicio.SetFocus
            Exit Function
        End If
    Else
        ExibeMensagem "Data de início da Atividade é obrigatória."
        txt_dtmAtividadeInicio.SetFocus
        Exit Function
    End If
    
    If Len(txt_dtmAtividadeFim) > 0 And Len(txt_dtmAtividadeInicio) > 0 Then
        If CDate(txt_dtmAtividadeFim) < CDate(txt_dtmAtividadeInicio) Then
            ExibeMensagem "A data de término não pode ser menor que a data de início."
            txt_dtmAtividadeInicio.SetFocus
            Exit Function
        End If
    End If
    
    If Not dbc_intAtividadePrincipal.MatchedWithList Then
        ExibeMensagem "É necessário informar uma atividade."
        dbc_intAtividadePrincipal.SetFocus
        Exit Function
    End If
    If chk_blnPrincipal And Not mblnAlterandoListaAtividade Then
        For intFor = 1 To lvw_Atividades.ListItems.Count
            If Abs(CInt(CBool(lvw_Atividades.ListItems(intFor).ListSubItems(2)))) = 1 Then
                ExibeMensagem "Já existe uma atividade preenchidada com status de principal."
                Exit Function
            End If
        Next
    End If
    
    With lvw_Atividades
        If .ListItems.Count > 0 Then
            For intFor = 1 To .ListItems.Count
                If .ListItems(intFor).SubItems(1) = dbc_intAtividadePrincipal.BoundText Then 'And .ListItems(.SelectedItem.Index).SubItems(1) <> dbc_intAtividadePrincipal.BoundText Then
                    'ExibeMensagem "Atividade não pode ser incluída, pois a mesma já se encontra cadastrada."
                    LimpaTabAtividade False
                    dbc_intAtividadePrincipal.SetFocus
                    mblnAlterandoListaAtividade = False
                    blnOkAtividade = True
                    Exit Function
                End If
            Next
        End If
    End With

    With lvw_Atividades
       If .ListItems.Count > 0 And mblnAlterandoListaAtividade Then
            For intFor = 1 To lvw_ItensAtivTrib.ListItems.Count
                If lvw_ItensAtivTrib.ListItems(intFor).Text = .ListItems(.SelectedItem.Index).SubItems(1) Then
                    ExibeMensagem "Atividade não pode ser alterada, pois a mesma se encontra relacionada nos tributos."
                    LimpaTabAtividade False
                    dbc_intAtividadePrincipal.SetFocus
                    mblnAlterandoListaAtividade = False
                    Exit Function
                End If
            Next
        End If
    End With
    
    blnDadoAtividades = True
End Function

Sub LimpaTabAtividade(Optional blnLimpaGrid As Boolean)
    chk_blnPrincipal.Value = 0
    dbc_intAtividadePrincipal.Text = ""
    txt_dtmAtividadeInicio.Text = ""
    txt_dtmAtividadeFim.Text = ""
    mblnAlterandoListaAtividade = False
    If blnLimpaGrid Then
        lvw_Atividades.ListItems.Clear
    End If
End Sub


Private Sub PreencheComboTributos()
    Dim strSql      As String
    Dim strSql1     As String
    Dim intFor      As Integer
    
    
    With lvw_Atividades
        If .ListItems.Count > 0 Then
            strSql = ""
            strSql = strSql & "SELECT A.Pkid, Rtrim(Ltrim(A.STRDESCRICAO)) as STRDESCRICAO "
            strSql = strSql & "FROM "
            strSql = strSql & gstrAtividadeEC & " A "
            strSql = strSql & "Where "
        
            For intFor = 1 To .ListItems.Count
                strSql1 = strSql1 & .ListItems(intFor).SubItems(1) & ","
            Next
            
            strSql1 = Mid(strSql1, 1, Len(strSql1) - 1)
            
            strSql = strSql & "A.Pkid in(" & strSql1 & ")"
            
            dbc_intAtividades.Tag = strSql & ";strDescricao"
            LeDaTabelaParaObj "", dbc_intAtividades, strSql
            
        End If
    End With
End Sub

Sub LimpaISS(Optional blnLimpaGrid As Boolean)

    Set dbc_intTipoISS.RowSource = Nothing
    dbc_intTipoISS.Text = ""
    Set dbc_intListaServico.RowSource = Nothing
    dbc_intListaServico.Text = ""
    
    txt_dtmissinicio = ""
    txt_dtmissfim = ""
    txt_intQuantidadeIss = "1"
    
    If blnLimpaGrid Then
        lvw_ISS.ListItems.Clear
    End If
    mblnAlterandoListaISS = False
End Sub

Private Function blnDadoISS() As Boolean
    Dim intFor As Integer
    
    blnDadoISS = False
    
    If Not dbc_intTipoISS.MatchedWithList Then
        ExibeMensagem "É necessário informar um Tipo de ISS."
        dbc_intTipoISS.SetFocus
        Exit Function
    End If

    If Not dbc_intListaServico.MatchedWithList Then
        ExibeMensagem "É necessário informar o campo de Lista de Serviço."
        dbc_intListaServico.SetFocus
        Exit Function
    End If

    
    If Trim(txt_dtmissinicio) <> "" Then
        If Not gblnDataValida(txt_dtmissinicio) Then
            ExibeMensagem "Data de início de ISSQN inválida."
            txt_dtmissinicio.SetFocus
            Exit Function
        End If
    Else
        ExibeMensagem "Data de início de ISSQN é obrigatória."
        txt_dtmissinicio.SetFocus
        Exit Function
    End If
    
    If Len(txt_dtmissfim) > 0 And Len(txt_dtmissinicio) > 0 Then
        If CDate(txt_dtmissfim) < CDate(txt_dtmissinicio) Then
            ExibeMensagem "A data de término  de ISSQN não pode ser menor que a data de início."
            txt_dtmissinicio.SetFocus
            Exit Function
        End If
    End If
   
    With lvw_ISS
        If .ListItems.Count > 0 Then
            If mblnAlterandoListaISS Then
                For intFor = 1 To .ListItems.Count
                    If .ListItems(intFor).SubItems(1) = dbc_intTipoISS.BoundText And _
                        .ListItems(intFor).SubItems(3) = dbc_intListaServico.BoundText And .ListItems(.SelectedItem.Index).SubItems(1) <> dbc_intTipoISS.BoundText Then
                        ExibeMensagem "Não pode ser inseridos Tipo de Iss e Lista de Serviço, pois ja se encontram cadastrados."
                        dbc_intTipoISS.SetFocus
                        mblnAlterandoListaISS = False
                        Exit Function
                    End If
                Next
            Else
                For intFor = 1 To .ListItems.Count
                    If .ListItems(intFor).SubItems(1) = dbc_intTipoISS.BoundText And Not mblnAlterandoListaISS _
                        And .ListItems(intFor).SubItems(3) = dbc_intListaServico.BoundText Then
                        ExibeMensagem "Não pode ser inseridos Tipo de Iss e Lista de Serviço, pois ja se encontram cadastrados."
                        dbc_intTipoISS.SetFocus
                        mblnAlterandoListaISS = False
                        Exit Function
                    End If
                Next
            
            End If
        End If
    End With
   
    blnDadoISS = True
End Function

Private Function strQueryTipoIss() As String
    Dim strSql As String
    
    strSql = "Select Pkid, Rtrim(Ltrim(strDescricao)) strDescricao From " & gstrTipoIss & " Order By strDescricao "
    strQueryTipoIss = strSql
    
End Function

Private Function blnSalvaISS(intPKIdEconomico As Long) As Boolean
    Dim strSql              As String
    Dim strSql1             As String
    Dim strSql2             As String
    Dim intFor              As Integer
    Dim strPkidHistorico    As String
    
    blnSalvaISS = False
    
    strSql = ""
    strSql2 = ""
    strPkidHistorico = ""
    
    If lvw_ISS.ListItems.Count <= 0 Then
        strSql = ""
        strSql = strSql & "DELETE FROM " & gstrIssEmpresa
        strSql = strSql & " WHERE intEconomico = " & intPKIdEconomico
    Else
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
        
        For intFor = 1 To lvw_ISS.ListItems.Count
            With lvw_ISS
                If .ListItems(intFor).Text <> "" Then
                
                    strPkidHistorico = strPkidHistorico & .ListItems(intFor).Text & ","
                    
                    strSql = strSql & "UPDATE " & gstrIssEmpresa
                    strSql = strSql & " SET inttipoiss = " & .ListItems(intFor).SubItems(1) & ", "
                    strSql = strSql & " intlistaservico = " & .ListItems(intFor).SubItems(3) & ", "
                    
                    If .ListItems(intFor).Selected And dbc_intTipoISS.MatchedWithList Then
                        strSql = strSql & " dtmissinicio = " & gstrConvDtParaSql(txt_dtmissinicio) & ", "
                        strSql = strSql & " dtmissfim = " & gstrConvDtParaSql(txt_dtmissfim) & ", "
                    Else
                        strSql = strSql & " dtmissinicio = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(5)) & ", "
                        strSql = strSql & " dtmissfim = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(6)) & ", "
                    End If
                    
                    strSql = strSql & " intquantidadeiss = " & .ListItems(intFor).SubItems(7) & ", "
                    strSql = strSql & " dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSql = strSql & " lngCodUsr = " & glngCodUsr
                    strSql = strSql & " WHERE Pkid = " & .ListItems(intFor).Text
                    strSql = strSql & IIf(bytDBType = Oracle, ";", "")
                Else
                    strSql2 = strSql2 & "INSERT INTO " & gstrIssEmpresa
                    strSql2 = strSql2 & " (inteconomico,"
                    strSql2 = strSql2 & " inttipoiss,"
                    strSql2 = strSql2 & " intlistaservico,"
                    strSql2 = strSql2 & " dtmissinicio,"
                    strSql2 = strSql2 & " dtmissfim,"
                    strSql2 = strSql2 & " intquantidadeiss,"
                    strSql2 = strSql2 & " dtmDtAtualizacao,"
                    strSql2 = strSql2 & " lngCodUsr)"
                    strSql2 = strSql2 & " VALUES( "
                    strSql2 = strSql2 & Val(intPKIdEconomico) & ", "
                    strSql2 = strSql2 & .ListItems(intFor).SubItems(1) & ", "
                    strSql2 = strSql2 & .ListItems(intFor).SubItems(3) & ", "
                    strSql2 = strSql2 & gstrConvDtParaSql(.ListItems(intFor).SubItems(5)) & ", "
                    strSql2 = strSql2 & gstrConvDtParaSql(.ListItems(intFor).SubItems(6)) & ", "
                    strSql2 = strSql2 & .ListItems(intFor).SubItems(7) & ", "
                    strSql2 = strSql2 & strGETDATE & ", "
                    strSql2 = strSql2 & glngCodUsr
                    strSql2 = strSql2 & ")"
                    strSql2 = strSql2 & IIf(bytDBType = Oracle, ";", "")
                End If
            End With
        Next
        
        If strPkidHistorico <> "" Then
            strPkidHistorico = Mid(strPkidHistorico, 1, Len(strPkidHistorico) - 1)
            strSql1 = ""
            strSql1 = strSql1 & "DELETE FROM " & gstrIssEmpresa
            strSql1 = strSql1 & " WHERE Pkid NOT in(" & strPkidHistorico & ")and"
            strSql1 = strSql1 & " intEconomico = " & intPKIdEconomico
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", " ")
            strSql = strSql & " " & strSql1
        Else
            strSql1 = strSql1 & "DELETE FROM " & gstrIssEmpresa
            strSql1 = strSql1 & " WHERE intEconomico = " & intPKIdEconomico
            strSql1 = strSql1 & IIf(bytDBType = Oracle, ";", "")
            strSql = strSql & " " & strSql1
        End If
        strSql = strSql & " " & strSql2
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSql) = False Then
        ExibeMensagem "Ocorreu um erro ao gravar os ISSQN. Os dados não foram gravados."
        blnSalvaISS = False
        Exit Function
    Else
        blnSalvaISS = True
    End If
    
End Function

Private Sub PreencheGrdISS(lngPkid As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
        
    lvw_ISS.ListItems.Clear
    
    strSql = "Select "
    strSql = strSql & "IE.Pkid intISSEmpresa, "
    strSql = strSql & "TI.Pkid intTipoISS, "
    strSql = strSql & "Rtrim(Ltrim(TI.Strdescricao)) strTipoISS, "
    strSql = strSql & "LS.Pkid intListaServico, "
    
    If bytDBType = EDatabases.SQLServer Then
       strSql = strSql & "REPLICATE('0',5 - " & strLen & "(LS.strCodigo))" & strCONCAT & "RTRIM(LTRIM(strCodigo)) " & _
                         strCONCAT & "' - '" & strCONCAT & _
                         " RTRIM(LTRIM(LS.strDescricao)) strListaServico, "
    Else
       strSql = strSql & "RTRIM(LTRIM( " & gstrCONVERT(CDT_VARCHAR, "LS.strCodigo,'00000'") & ")) " & _
                         strCONCAT & "' - '" & strCONCAT & _
                         " RTRIM(LTRIM(LS.strDescricao)) strListaServico, "
    End If
    
    strSql = strSql & "IE.Dtmissinicio, "
    strSql = strSql & "IE.Dtmissfim, "
    strSql = strSql & "IE.Intquantidadeiss "
    strSql = strSql & "From "
    strSql = strSql & gstrEconomico & " EC, "
    strSql = strSql & gstrIssEmpresa & " IE, "
    strSql = strSql & gstrTipoIss & " TI, "
    strSql = strSql & gstrListaServico & " LS "
    strSql = strSql & "Where "
    strSql = strSql & "Ec.Pkid = IE.Inteconomico AND "
    strSql = strSql & "TI.Pkid = IE.Inttipoiss AND "
    strSql = strSql & "LS.Pkid = IE.Intlistaservico AND "
    strSql = strSql & "EC.Pkid = " & Val(txtPKId)
       
    Set gobjBanco = New clsBanco
    lvw_Socios.ListItems.Clear
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                If .EOF = False Then
                    Do While Not .EOF
                        Set mobjLista = lvw_ISS.ListItems.Add(, , gstrENulo(!intISSEmpresa))
                        mobjLista.SubItems(1) = gstrENulo(!intTipoISS)
                        mobjLista.SubItems(2) = gstrENulo(!STRTIPOISS)
                        mobjLista.SubItems(3) = gstrENulo(!intListaServico)
                        mobjLista.SubItems(4) = gstrENulo(!strListaServico)
                        mobjLista.SubItems(5) = gstrDataFormatada(gstrENulo(!Dtmissinicio))
                        mobjLista.SubItems(6) = gstrDataFormatada(gstrENulo(!Dtmissfim))
                        mobjLista.SubItems(7) = gstrENulo(!Intquantidadeiss)
                        .MoveNext
                    Loop
                End If
            End With
        End If
    End If
End Sub

Private Function blnEncerraEmpresa(lngPkid As Long, strDtEncerramento) As Boolean
    Dim strSql As String
    
    blnEncerraEmpresa = False
    
    strSql = IIf(bytDBType = Oracle, "Begin ", "")
    
    strSql = strSql & "Update " & gstrHistoricoEconVariavel & " SET dtmdtFinal = " & gstrConvDtParaSql(strDtEncerramento) & " "
    strSql = strSql & "where Inteconomico = " & lngPkid & " AND Dtmdtfinal IS NULL"
    strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
    
    strSql = strSql & "Update " & gstrAtividadeDaEmpresa & " SET dtmatividadefim = " & gstrConvDtParaSql(strDtEncerramento) & " "
    strSql = strSql & "where Inteconomico = " & lngPkid & " AND dtmatividadefim IS NULL"
    strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
    
    strSql = strSql & "Update " & gstrSocioEconomico & " SET dtmsociofim = " & gstrConvDtParaSql(strDtEncerramento) & " "
    strSql = strSql & "where intcodeconomico = " & lngPkid & " AND dtmsociofim IS NULL"
    strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
    
    strSql = strSql & "Update " & gstrHistoricoPublicidades & " SET dtmpublicidadefim = " & gstrConvDtParaSql(strDtEncerramento) & " "
    strSql = strSql & "where inteconomico = " & lngPkid & " AND dtmpublicidadefim IS NULL"
    strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
    
    strSql = strSql & "Update " & gstrIssEmpresa & " SET dtmissfim = " & gstrConvDtParaSql(strDtEncerramento) & " "
    strSql = strSql & "where Inteconomico = " & lngPkid & " AND dtmissfim IS NULL"
    strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
    
    strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    
    If Not gobjBanco.Execute(strSql) Then
       Exit Function
    End If
    
    blnEncerraEmpresa = True
End Function

