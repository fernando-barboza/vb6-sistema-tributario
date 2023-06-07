VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmContNotasFiscais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Notas Fiscais"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   HelpContextID   =   15
   Icon            =   "ContNotasFiscais.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8520
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   75
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Controle de Notas Fiscais"
      TabPicture(0)   =   "ContNotasFiscais.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdd_DetNotasFiscais"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Cadastro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Dados"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame fra_Dados 
         Caption         =   "Dados das Notas Fiscais"
         Height          =   1335
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   8055
         Begin VB.TextBox txt_intNumFimTalao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6240
            TabIndex        =   7
            Top             =   960
            Width           =   1605
         End
         Begin VB.TextBox txt_intNumInicTalao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   960
            Width           =   1605
         End
         Begin VB.TextBox txt_strSerieNotaFiscal 
            Height          =   285
            Left            =   6840
            MaxLength       =   2
            TabIndex        =   5
            Top             =   600
            Width           =   1005
         End
         Begin VB.TextBox txt_strNumNotaFiscal 
            Height          =   285
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   4
            Top             =   600
            Width           =   2445
         End
         Begin VB.TextBox txt_dtmDtsolicitacao 
            Height          =   285
            Left            =   6840
            TabIndex        =   3
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txt_dtmDtExercicio 
            Height          =   285
            Left            =   4320
            TabIndex        =   2
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txt_strProtocolo 
            Height          =   285
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   1
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lblfimtalao 
            AutoSize        =   -1  'True
            Caption         =   "Número Final do Talão"
            Height          =   195
            Left            =   4560
            TabIndex        =   25
            Top             =   1020
            Width           =   1605
         End
         Begin VB.Label lblinictalao 
            AutoSize        =   -1  'True
            Caption         =   "Número Inicial do Talão"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1020
            Width           =   1680
         End
         Begin VB.Label lblserienotafiscal 
            AutoSize        =   -1  'True
            Caption         =   "Número de Série da Nota Fiscal"
            Height          =   195
            Left            =   4440
            TabIndex        =   23
            Top             =   660
            Width           =   2250
         End
         Begin VB.Label lblnumnotafiscal 
            AutoSize        =   -1  'True
            Caption         =   "Número da Nota Fiscal"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   660
            Width           =   1620
         End
         Begin VB.Label lbldatasolicitacao 
            AutoSize        =   -1  'True
            Caption         =   "Data da Solicitação"
            Height          =   195
            Left            =   5400
            TabIndex        =   21
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label lblExercicio 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   3870
            TabIndex        =   20
            Top             =   300
            Width           =   345
         End
         Begin VB.Label lblNumProcesso 
            AutoSize        =   -1  'True
            Caption         =   "Número do Processo"
            Height          =   195
            Left            =   360
            TabIndex        =   19
            Top             =   300
            Width           =   1485
         End
      End
      Begin VB.Frame fra_Cadastro 
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   320
         Width           =   8055
         Begin VB.TextBox txtPKId 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Txt_PKIdContribuinte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txt_strAtividadeBasica 
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   960
            Width           =   2205
         End
         Begin VB.TextBox txt_strCNPJCPF 
            Height          =   285
            Left            =   1335
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   960
            Width           =   1605
         End
         Begin MSMask.MaskEdBox msk_strInscricaoCadastral 
            Height          =   285
            Left            =   5640
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo dbcintContribuinte 
            Height          =   315
            Left            =   2520
            TabIndex        =   0
            Top             =   600
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Atividade Principal"
            Height          =   195
            Left            =   4200
            TabIndex        =   17
            Top             =   1020
            Width           =   1305
         End
         Begin VB.Label lbl_CNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ / CPF"
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   1020
            Width           =   870
         End
         Begin VB.Label lbl_PKId 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   735
            TabIndex        =   15
            Top             =   300
            Width           =   495
         End
         Begin VB.Label lblintContribuinte 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   795
            TabIndex        =   14
            Top             =   675
            Width           =   435
         End
         Begin VB.Label lblstrInscricaoCadastral 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   4170
            TabIndex        =   13
            Top             =   330
            Width           =   1350
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdd_DetNotasFiscais 
         Height          =   1485
         Left            =   120
         TabIndex        =   27
         Top             =   3120
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2619
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nº Processo"
         Columns(1).DataField=   "strProtocolo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Exercício"
         Columns(2).DataField=   "dtmDtExercicio"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Dt Solicitação"
         Columns(3).DataField=   "dtmDtSolicitacao"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nº Nota Fiscal"
         Columns(4).DataField=   "strNumNotaFiscal"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Série NF"
         Columns(5).DataField=   "strSerieNotaFiscal"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Nº Inicio Talão"
         Columns(6).DataField=   "intNumInicTalao"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Nº Final Talão"
         Columns(7).DataField=   "intNumFimTalao"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2566"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1349"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1270"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=1879"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1799"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2540"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2461"
         Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=1323"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=1244"
         Splits(0)._ColumnProps(30)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=1984"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=1905"
         Splits(0)._ColumnProps(35)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(37)=   "Column(7).Width=1905"
         Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1826"
         Splits(0)._ColumnProps(40)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
   Begin TrueOleDBGrid70.TDBGrid tdb_ContNotasFiscais 
      Height          =   1605
      Left            =   120
      TabIndex        =   26
      Top             =   4830
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2831
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "PKId"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nome"
      Columns(1).DataField=   "strNome"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Inscrição Cadastral"
      Columns(2).DataField=   "strInscricaoCadastral"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2540"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2461"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=8202"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=8123"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3307"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3228"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmContNotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim mblnAlterando                   As Boolean
Dim mblnAlterandoH                  As Boolean
Dim mobjAux                         As Object
Dim mblnClickOk                     As Boolean
Dim mblnClickCaracteristica         As Boolean
Dim oList                           As Object
    
Dim X                               As New XArrayDB 'Grid detalhes de Notas Fiscais
    
Dim xDet                            As XArrayDB
    
Dim mblnSelecionou                  As Boolean
Dim mblnPrimeiraVez                 As Boolean

Private Sub dbcintContribuinte_Click(Area As Integer)
    DropDownDataCombo dbcintContribuinte, Me, Area
    If Area = 2 Then
        ExibeDadosContribuinte
    End If
    If Area = 0 Then
        ExibeDadosNovos
    End If
End Sub

Private Sub dbcintContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContribuinte_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            dbcintContribuinte_Click 2
    End Select
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 648
    VirificaGradeListView Me
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

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    TrocaCorObjeto txt_PKIdContribuinte, True
    TrocaCorObjeto txtPKId, True
    
    
    dbcintContribuinte.Tag = "SELECT PKId, strNome FROM " & gstrContribuinte & " ORDER BY strNome;strNome"
        
    VerificaListaAutomatica gstrEconomico, tdb_ContNotasFiscais, strQuery
    
    VerificaObjParaAplicar mobjAux
End Sub

Function strQuery() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT EC.PKId, CO.strNome, EC.strInscricaoCadastral "
    strSQL = strSQL & " FROM " & gstrEconomico & " EC,"
    strSQL = strSQL & gstrContribuinte & " CO "
    strSQL = strSQL & " WHERE EC.intContribuinte = CO.PKId "
    strSQL = strSQL & " ORDER BY EC.PKId"
strQuery = strSQL
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub msk_strInscricaoCadastral_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", msk_strInscricaoCadastral
End Sub

Private Sub tdd_DetNotasFiscais_KeyPress(KeyAscii As Integer)
    On Error GoTo Err_Handle
    Select Case tdd_DetNotasFiscais.col
        Case 1
            CaracterValido KeyAscii, "A", tdd_DetNotasFiscais
     End Select
    Exit Sub
Err_Handle:
End Sub

Private Sub tdd_DetNotasFiscais_Click()
    mblnClickCaracteristica = True
End Sub

Private Sub tdb_ContNotasFiscais_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_ContNotasFiscais_FilterChange()
    gblnFilraCampos tdb_ContNotasFiscais
End Sub

Private Sub tdb_ContNotasFiscais_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown
        mblnClickOk = True
    Case Else
        mblnClickOk = False
    End Select
End Sub

Private Sub tdb_ContNotasFiscais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_ContNotasFiscais_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_ContNotasFiscais
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                If mblnClickOk Then
                    mblnClickOk = False
                    mblnAlterando = True
                    mblnSelecionou = True
                    
                    txtPKId = .Columns("PKID").Value
                    LeDaTabelaParaObj gstrEconomico, Me
                    gCorLinhaSelecionada tdb_ContNotasFiscais
                    
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                    
                    If mobjAux Is Nothing Then
                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                    End If
                        
                    dbcintContribuinte_Click 2
                    
                    LeDaTabelaParaObj gstrContNotasFiscais, tdd_DetNotasFiscais, strQueryDetNotasFiscais
                    
                    mblnAlterandoH = False
                    
                    tab_3dPasta.Tab = 0
                End If
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark      As Variant
    Dim intPKIdNotasFiscais As Integer

    Select Case UCase(strModoOperacao)
        Case "NOVO"
            LimpaObjeto Me, mblnAlterando
            Call NovoContNotasFiscais

        Case "SALVAR"
            If blnDadosOk Then
                Screen.MousePointer = vbHourglass
                If ToolBarGeral(strModoOperacao, gstrContNotasFiscais, mblnAlterando, tdb_ContNotasFiscais, Me, mobjAux, strQuery, , , , False) Then
'                    If mblnAlterando Then
                        intPKIdNotasFiscais = txtPKId
'                    Else: intPKIdNotasFiscais = glngPegaUltimaChave(gstrContNotasFiscais, "PKId")
'                    End If
                        If blnGravaNotasFiscais(mblnAlterando, intPKIdNotasFiscais) Then
                        
                
                        
                        End If
                End If
                tab_3dPasta.Tab = 0
                Screen.MousePointer = vbDefault
            End If
            
        Case "IMPRIMIR"
            ImprimeRelatorio rptControleDeNotasFiscais, strQuerryRelatorio
            
        Case "DELETAR"
            If blnDeletaNotasFiscais Then
                'Limpar tdd_detalhes
            End If
        Case gstrPreencherLista
            PreencherListaDeOpcoes dbcintContribuinte
        
        Case "FECHAR"
            Unload Me
    End Select
Screen.MousePointer = vbDefault
End Sub

Private Function blnDeletaNotasFiscais() As Boolean
    Dim strSQL As String

    If MsgBox("Confirma exclusão do registro de '" & dbcintContribuinte.Text & "' ?", vbQuestion + vbYesNo) = vbYes Then
        DeletaNotasFiscais Val(txtPKId)
        
        strSQL = ""
        strSQL = strSQL & "Delete From " & gstrContNotasFiscais & " "
        strSQL = strSQL & "Where PKId = " & Val(txtPKId)
        
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSQL
        
        VerificaListaAutomatica gstrContNotasFiscais, tdb_ContNotasFiscais, strQuery
    End If
    blnDeletaNotasFiscais = True
End Function

Sub ExibeDadosContribuinte()

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    txt_PKIdContribuinte = ""
    
    If dbcintContribuinte.BoundText = "" Then Exit Sub
    
    txt_PKIdContribuinte = dbcintContribuinte.BoundText
    
    strSQL = ""
    strSQL = strSQL & "SELECT CO.strNome, EC.strInscricaoCadastral, CO.strCNPJCPF, AB.strDescricao, "
    strSQL = strSQL & "NF.strProtocolo, NF.dtmDtExercicio, NF.dtmDtSolicitacao, NF.strNumNotaFiscal, "
    strSQL = strSQL & "NF.strSerieNotaFiscal, NF.intNumInicTalao, NF.intNumFimTalao "
    strSQL = strSQL & "FROM " & gstrEconomico & " EC "
'    strSql = strSql & "INNER JOIN " & gstrContribuinte & " CO "
    strSQL = strSQL & ", " & gstrContribuinte & " CO "
'    strSql = strSql & "ON EC.intContribuinte= CO.PKId "
'    strSql = strSql & "INNER JOIN " & gstrContNotasFiscais & " NF "
    strSQL = strSQL & ", " & gstrContNotasFiscais & " NF "
'    strSql = strSql & "ON EC.PKId = NF.intContribuinte "
'    strSql = strSql & "INNER JOIN " & gstrAtividadeBasica & " AB "
    strSQL = strSQL & ", " & gstrAtividadeBasica & " AB "
'    strSql = strSql & "ON EC.intAtividadeBasica = AB.PKId "
    strSQL = strSQL & "WHERE CO.PKId = " & dbcintContribuinte.BoundText
    
    strSQL = strSQL & " AND EC.intContribuinte= CO.PKId "
    strSQL = strSQL & " AND EC.PKId = NF.intContribuinte "
    strSQL = strSQL & " AND EC.intAtividadeBasica = AB.PKId "
    
    strSQL = strSQL
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 4, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                txt_strCNPJCPF = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!strCNPJCPF))
                msk_strInscricaoCadastral = gstrVerificaCampoNulo(!strInscricaoCadastral)
                txt_strAtividadeBasica = gstrVerificaCampoNulo(!strDescricao)
                txt_strProtocolo = gstrVerificaCampoNulo(!strProtocolo)
                txt_dtmDtExercicio = gstrDataFormatada(gstrVerificaCampoNulo(!dtmDtExercicio))
                txt_dtmDtsolicitacao = gstrDataFormatada(gstrVerificaCampoNulo(!dtmDtSolicitacao))
                txt_strNumNotaFiscal = gstrVerificaCampoNulo(!strNumNotaFiscal)
                txt_strSerieNotaFiscal = gstrVerificaCampoNulo(!strSerieNotaFiscal)
                txt_intNumInicTalao = gstrVerificaCampoNulo(!intNumInicTalao)
                txt_intNumFimTalao = gstrVerificaCampoNulo(!intNumFimTalao)
             Else
                txt_strCNPJCPF = ""
                msk_strInscricaoCadastral = ""
                txt_strAtividadeBasica = ""
                txt_strProtocolo = ""
                txt_dtmDtExercicio = ""
                txt_dtmDtsolicitacao = ""
                txt_strNumNotaFiscal = ""
                txt_strSerieNotaFiscal = ""
                txt_intNumInicTalao = ""
                txt_intNumFimTalao = ""
            End If
        End With
    End If
    
End Sub

Sub ExibeDadosNovos()
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    If dbcintContribuinte.BoundText = "" Then Exit Sub
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strCNPJCPF, strNome "
    strSQL = strSQL & "FROM " & gstrContribuinte & " "
    strSQL = strSQL & "WHERE PKId = " & dbcintContribuinte.BoundText
    strSQL = strSQL
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 4, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                txt_strCNPJCPF = gstrCGCCPFFormatado(gstrVerificaCampoNulo(!strCNPJCPF))
             Else
                txt_strCNPJCPF = ""
            End If
        End With
    End If
    
End Sub


Private Function blnDadosOk() As Boolean
    If dbcintContribuinte.BoundText = "" Then
        ExibeMensagem "O Nome do Contribuinte tem que ser selecionado."
        Exit Function
    End If
    If Trim(txt_intNumInicTalao) <> "" Or Trim(txt_intNumFimTalao) <> "" Then
        If Trim(txt_intNumInicTalao) = "" Then
            ExibeMensagem "O número inicial do talão deve ser informado."
            txt_intNumInicTalao.SetFocus
            Exit Function
        ElseIf Trim(txt_intNumFimTalao) = "" Then
            ExibeMensagem "O número final do talão deve ser informado."
            txt_intNumFimTalao.SetFocus
            Exit Function
        ElseIf Val(txt_intNumInicTalao) > Val(txt_intNumFimTalao) Then
            ExibeMensagem "O número final do talão não pode ser menor que o número inicial."
            txt_intNumInicTalao.SetFocus
            Exit Function
        End If
    End If
    blnDadosOk = True
End Function

Private Sub NovoContNotasFiscais()
    
    On Error GoTo err_NovoContNotasFiscais
    
    txtPKId = glngPegaProximaChave(gstrContNotasFiscais, "PKId")
    
    txt_PKIdContribuinte = ""
    txt_strCNPJCPF = ""
    msk_strInscricaoCadastral = ""
    txt_strAtividadeBasica = ""
    txt_strProtocolo = ""
    txt_dtmDtExercicio = ""
    txt_dtmDtsolicitacao = ""
    txt_strNumNotaFiscal = ""
    txt_strSerieNotaFiscal = ""
    txt_intNumInicTalao = ""
    txt_intNumFimTalao = ""
    
    tdd_DetNotasFiscais.DataSource = Nothing
    
    dbcintContribuinte.SetFocus
    
    mblnAlterando = False
    mblnAlterandoH = True
    tab_3dPasta.Tab = 0
err_NovoContNotasFiscais:
End Sub

Function blnGravaNotasFiscais(blnAlterando As Boolean, intContribuinte As Integer) As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL           As String
    Dim i                As Integer

    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    If IsNull(tdd_DetNotasFiscais.Columns(0).Value) Then
         mblnAlterandoH = True
    End If
    
    Select Case mblnAlterandoH
        Case True
            strSQL = ""
            strSQL = strSQL & "INSERT INTO " & gstrContNotasFiscais & " "
            strSQL = strSQL & "(intcontribuinte, strProtocolo, dtmDtExercicio, dtmDtSolicitacao, "
            strSQL = strSQL & "strNumNotaFiscal, strSerieNotaFiscal, intNumInicTalao, "
            strSQL = strSQL & "intNumFimTalao "
            strSQL = strSQL & ") VALUES ("
            strSQL = strSQL & "'" & Val(txtPKId) & "', "
            strSQL = strSQL & "'" & Trim(txt_strProtocolo) & "', "
            strSQL = strSQL & gstrConvDtParaSql(txt_dtmDtExercicio) & ", "
            strSQL = strSQL & gstrConvDtParaSql(txt_dtmDtsolicitacao) & ", "
            strSQL = strSQL & "'" & Trim(txt_strNumNotaFiscal) & "',"
            strSQL = strSQL & "'" & Trim(txt_strSerieNotaFiscal) & "',"
            strSQL = strSQL & Val(txt_intNumInicTalao) & ", "
            strSQL = strSQL & Val(txt_intNumFimTalao) & ""
            strSQL = strSQL & ") "
            
        Case False
            strSQL = ""
            strSQL = strSQL & "UPDATE " & gstrContNotasFiscais & " SET "
            strSQL = strSQL & "strProtocolo = '" & Trim(txt_strProtocolo) & "', "
            strSQL = strSQL & "dtmDtExercicio = " & gstrConvDtParaSql(txt_dtmDtExercicio) & ", "
            strSQL = strSQL & "dtmDtSolicitacao = " & gstrConvDtParaSql(txt_dtmDtsolicitacao) & ", "
            strSQL = strSQL & "strNumNotaFiscal = '" & Trim(txt_strNumNotaFiscal) & "', "
            strSQL = strSQL & "strSerieNotaFiscal= '" & Trim(txt_strSerieNotaFiscal) & "', "
            strSQL = strSQL & "intNumInicTalao = " & Val(txt_intNumInicTalao) & ", "
            strSQL = strSQL & "intNumFimTalao = " & Val(txt_intNumFimTalao) & ", "
'            strSql = strSql & "dtmDtAtualizacao = GETDATE(), "
            strSQL = strSQL & "dtmDtAtualizacao = " & strGETDATE & ", "
            strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
            strSQL = strSQL & "WHERE PKId = " & tdd_DetNotasFiscais.Columns(0).Value
            
        End Select
        
            If Not gobjBanco.Execute(strSQL, False) Then
                gobjBanco.ExecutaRollbackTrans
                ExibeMensagem "Ocorreu um erro ao gravar os detalhes da Nota Fiscal. Os dados não foram gravados."
                Exit Function
            End If
            gobjBanco.ExecutaCommitTrans
            Set gobjBanco = Nothing
            blnGravaNotasFiscais = True
            LeDaTabelaParaObj gstrContNotasFiscais, tdd_DetNotasFiscais, strQueryDetNotasFiscais
    
End Function

Sub DeletaNotasFiscais(intContribuinte As Integer)
    Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "DELETE FROM " & gstrContNotasFiscais & " "
    strSQL = strSQL & "WHERE intContribuinte = " & intContribuinte
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSQL
End Sub

Private Function strQueryDetNotasFiscais() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT NF.PKId, NF.intContribuinte, "
    strSQL = strSQL & " NF.strProtocolo, NF.dtmDtExercicio, NF.dtmDtSolicitacao, "
    strSQL = strSQL & " NF.strNumNotaFiscal, NF.strSerieNotaFiscal, NF.intNumInicTalao, "
    strSQL = strSQL & " NF.intNumFimTalao "
    strSQL = strSQL & " FROM " & gstrContNotasFiscais & " NF "
    strSQL = strSQL & " WHERE NF.intContribuinte = " & Val(txtPKId)
strQueryDetNotasFiscais = strSQL
End Function

Private Sub tdd_DetNotasFiscais_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim intResultado As Integer
    With tdd_DetNotasFiscais
        If Not .EOF And Not .BOF Then
            If mblnClickCaracteristica Then
                    mblnAlterando = True
                    mblnSelecionou = True
                    
                    intResultado = .Columns("PKID").Value
                    LeDaTabelaParaObj gstrContNotasFiscais, Me
                    gCorLinhaSelecionada tdd_DetNotasFiscais
                    
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                    
                    If mobjAux Is Nothing Then
                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                    End If
                        
                    ExibeDadosDetalhes intResultado
                    
                    
                    tab_3dPasta.Tab = 0
            End If
        End If
    End With
End Sub

Sub ExibeDadosDetalhes(intResultado As Integer)
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & " SELECT intContribuinte, strProtocolo, dtmDtExercicio, "
    strSQL = strSQL & " dtmDtSolicitacao, strNumNotaFiscal, strSerieNotaFiscal,"
    strSQL = strSQL & " intNumInicTalao, intNumFimTalao, PKId "
    strSQL = strSQL & "FROM " & gstrContNotasFiscais & " "
    strSQL = strSQL & "WHERE PKId = " & intResultado
    strSQL = strSQL
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 4, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                txt_strProtocolo = gstrVerificaCampoNulo(!strProtocolo)
                txt_dtmDtExercicio = gstrDataFormatada(gstrVerificaCampoNulo(!dtmDtExercicio))
                txt_dtmDtsolicitacao = gstrDataFormatada(gstrVerificaCampoNulo(!dtmDtSolicitacao))
                txt_strNumNotaFiscal = gstrVerificaCampoNulo(!strNumNotaFiscal)
                txt_strSerieNotaFiscal = gstrVerificaCampoNulo(!strSerieNotaFiscal)
                txt_intNumInicTalao = gstrVerificaCampoNulo(!intNumInicTalao)
                txt_intNumFimTalao = gstrVerificaCampoNulo(!intNumFimTalao)
             Else
                txt_strProtocolo = ""
                txt_dtmDtExercicio = ""
                txt_dtmDtsolicitacao = ""
                txt_strNumNotaFiscal = ""
                txt_strSerieNotaFiscal = ""
                txt_intNumInicTalao = ""
                txt_intNumFimTalao = ""
            End If
        End With
    End If
    
End Sub

Private Sub txt_dtmDtExercicio_GotFocus()
    MarcaCampo txt_dtmDtExercicio
End Sub

Private Sub txt_dtmDtExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDtExercicio
End Sub

Private Sub txt_dtmDtSolicitacao_GotFocus()
    MarcaCampo txt_dtmDtsolicitacao
End Sub

Private Sub txt_dtmDtSolicitacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDtsolicitacao
End Sub

Private Sub txt_intNumFimTalao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intNumFimTalao
End Sub

Private Sub txt_intNumInicTalao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intNumInicTalao
End Sub

Private Sub txt_strAtividadeBasica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strAtividadeBasica
End Sub

Private Sub txt_strCNPJCPF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strCNPJCPF
End Sub

Private Sub txt_strNumNotaFiscal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strNumNotaFiscal
End Sub

Private Sub txt_strProtocolo_KeyPress(KeyAscii As Integer)
        CaracterValido KeyAscii, "A", txt_strProtocolo
End Sub

Private Sub txt_strSerieNotaFiscal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strSerieNotaFiscal
End Sub

Function strQuerryRelatorio() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " EC.PKId, CO.strNome, EC.strInscricaoCadastral, "
    strSQL = strSQL & " NF.PKId, NF.intContribuinte, "
    strSQL = strSQL & " NF.strProtocolo, NF.dtmDtExercicio, NF.dtmDtSolicitacao, "
    strSQL = strSQL & " NF.strNumNotaFiscal, NF.strSerieNotaFiscal, NF.intNumInicTalao, "
    strSQL = strSQL & " NF.intNumFimTalao "
    
    strSQL = strSQL & " FROM " & gstrEconomico & " EC,"
    strSQL = strSQL & gstrContribuinte & " CO, "
    strSQL = strSQL & gstrContNotasFiscais & " NF "
    
    strSQL = strSQL & " WHERE EC.intContribuinte = CO.PKId "
    strSQL = strSQL & " AND NF.intContribuinte = EC.PKId "
    strSQL = strSQL & " ORDER BY CO.strNome "
strQuerryRelatorio = strSQL
End Function



