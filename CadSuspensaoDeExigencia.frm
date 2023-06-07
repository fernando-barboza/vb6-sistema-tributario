VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadSuspensaoDeExigencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suspensão de Exigências"
   ClientHeight    =   6810
   ClientLeft      =   3810
   ClientTop       =   2640
   ClientWidth     =   8655
   Icon            =   "CadSuspensaoDeExigencia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8655
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6660
      Left            =   90
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   60
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   11748
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Suspensão de Exigências"
      TabPicture(0)   =   "CadSuspensaoDeExigencia.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintintProtocolizacaoProcesso"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrsumula"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintOcorrencia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblstrCatalogoAssunto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_strContribuinte"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintOcorrencia"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbcintProtocolizacaoDoProcesso"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tdb_SuspensaoDeExigencia"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_Processo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_Inscricao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtintContribuinte"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fra_Parcela"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPKId"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtintParcelaReceita"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_strContribuinte"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_strCatalogoAssunto"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_strSumula"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.TextBox txt_strSumula 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   525
         Left            =   2235
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1470
         Width           =   5925
      End
      Begin VB.TextBox txt_strCatalogoAssunto 
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
         Left            =   2235
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   810
         Width           =   5925
      End
      Begin VB.TextBox txt_strContribuinte 
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
         Left            =   2235
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1140
         Width           =   5925
      End
      Begin VB.TextBox txtintParcelaReceita 
         Height          =   300
         Left            =   6375
         TabIndex        =   36
         Top             =   465
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtPKId 
         Height          =   300
         Left            =   4680
         TabIndex        =   35
         Top             =   465
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame fra_Parcela 
         Caption         =   "Parcela"
         Height          =   1320
         Left            =   135
         TabIndex        =   29
         Top             =   3030
         Width           =   8175
         Begin VB.TextBox txt_dtmDataDoVencimento 
            Height          =   285
            Left            =   6990
            MaxLength       =   12
            TabIndex        =   14
            Top             =   930
            Width           =   1035
         End
         Begin VB.TextBox txt_intNumeroParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2100
            MaxLength       =   2
            TabIndex        =   12
            Top             =   930
            Width           =   345
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   4500
            MaxLength       =   4
            TabIndex        =   13
            Top             =   930
            Width           =   525
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoCadastral 
            Height          =   315
            Left            =   2100
            TabIndex        =   10
            Top             =   210
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
            Height          =   315
            Left            =   2100
            TabIndex        =   11
            Top             =   570
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lbl_intExercicio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   3660
            TabIndex        =   34
            Top             =   1005
            Width           =   675
         End
         Begin VB.Label lbl_dtmDataDoVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Data de vencimento"
            Height          =   195
            Left            =   5400
            TabIndex        =   33
            Top             =   1005
            Width           =   1440
         End
         Begin VB.Label lbl_intParcela 
            AutoSize        =   -1  'True
            Caption         =   "N° da parcela receita"
            Height          =   195
            Left            =   450
            TabIndex        =   32
            Top             =   1005
            Width           =   1500
         End
         Begin VB.Label lbl_strInscricaoCadastral 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   600
            TabIndex        =   31
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label lbl_intComposicaoDaReceita 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Composição da Receita"
            Height          =   195
            Left            =   255
            TabIndex        =   30
            Top             =   660
            Width           =   1695
         End
      End
      Begin VB.TextBox txtintContribuinte 
         Height          =   300
         Left            =   5520
         TabIndex        =   28
         Top             =   465
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame fra_Inscricao 
         Height          =   630
         Left            =   135
         TabIndex        =   21
         Top             =   2355
         Width           =   8175
         Begin VB.OptionButton optbytTipoDeInscricaoCadastral 
            Caption         =   "Receitas Diversas"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   4
            Left            =   6480
            TabIndex        =   9
            Top             =   270
            Width           =   1605
         End
         Begin VB.OptionButton optbytTipoDeInscricaoCadastral 
            Caption         =   "Contribuição de Melhorias"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   3
            Left            =   4290
            TabIndex        =   8
            Top             =   270
            Width           =   2205
         End
         Begin VB.OptionButton optbytTipoDeInscricaoCadastral 
            Caption         =   "Econômico"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   2
            Left            =   3150
            TabIndex        =   7
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton optbytTipoDeInscricaoCadastral 
            Caption         =   "Imobiliário Rural"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton optbytTipoDeInscricaoCadastral 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   270
            Value           =   -1  'True
            Width           =   1605
         End
      End
      Begin VB.Frame fra_Processo 
         Caption         =   "Processo"
         Height          =   1185
         Left            =   135
         TabIndex        =   20
         Top             =   4395
         Width           =   8175
         Begin VB.CheckBox chkbytProcessoFinalizado 
            Caption         =   "Finalizado"
            Height          =   195
            Left            =   3690
            TabIndex        =   16
            Top             =   270
            Width           =   1005
         End
         Begin VB.TextBox txtstrTextoResultado 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   525
            Left            =   2100
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   555
            Width           =   5925
         End
         Begin VB.TextBox txtdtmProcesso 
            Height          =   285
            Left            =   2100
            MaxLength       =   12
            TabIndex        =   15
            Top             =   225
            Width           =   1035
         End
         Begin VB.Label lblstrTextoResultado 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Resultado"
            Height          =   195
            Left            =   1230
            TabIndex        =   27
            Top             =   705
            Width           =   720
         End
         Begin VB.Label lbldtmProcesso 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   1605
            TabIndex        =   26
            Top             =   270
            Width           =   345
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_SuspensaoDeExigencia 
         Height          =   885
         Left            =   135
         TabIndex        =   18
         Top             =   5640
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1561
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
         Columns(1).Caption=   "Contribuinte"
         Columns(1).DataField=   "strContribuinte"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nº da parcela"
         Columns(2).DataField=   "intNumeroParcela"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Data de Vencimento"
         Columns(3).DataField=   "dtmDataVencimento"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Processo Finalizado"
         Columns(4).DataField=   "strProcessoFinalizado"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=6376"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6297"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1931"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1852"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2805"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2725"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=2752"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2672"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=7,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintProtocolizacaoDoProcesso 
         Height          =   315
         Left            =   2235
         TabIndex        =   0
         Top             =   450
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintOcorrencia 
         Height          =   315
         Left            =   2235
         TabIndex        =   4
         Top             =   2040
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lbl_strContribuinte 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contribuinte"
         Height          =   195
         Left            =   1260
         TabIndex        =   37
         Top             =   1185
         Width           =   840
      End
      Begin VB.Label lblstrCatalogoAssunto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Assunto"
         Height          =   195
         Left            =   945
         TabIndex        =   25
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblintOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   1320
         TabIndex        =   24
         Top             =   2085
         Width           =   780
      End
      Begin VB.Label lblstrsumula 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Súmula"
         Height          =   195
         Left            =   1575
         TabIndex        =   23
         Top             =   1635
         Width           =   525
      End
      Begin VB.Label lblintintProtocolizacaoProcesso 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo"
         Height          =   195
         Left            =   1425
         TabIndex        =   22
         Top             =   510
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmCadSuspensaoDeExigencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnPrimeiraVezSuspensao         As Boolean
Dim blnAlterandoSuspensao           As Boolean

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoCadastral_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, Area
End Sub

Private Sub dbc_strInscricaoCadastral_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrencia_Click(Area As Integer)
    DropDownDataCombo dbcintOcorrencia, Me, Area
End Sub

Private Sub dbcintOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintProtocolizacaoDoProcesso_Click(Area As Integer)
    DropDownDataCombo dbcintProtocolizacaoDoProcesso, Me, Area
    If Area = 2 Then
        blnAlterandoSuspensao = False
        dbcintProtocolizacaoDoProcesso.IntegralHeight = False
        LimpaObjeto Me
        dbcintProtocolizacaoDoProcesso.IntegralHeight = True
        LimpaControlesDoFormulario
        BuscaDadosProtocolo
        optbytTipoDeInscricaoCadastral(0).Value = True
    End If
End Sub


Private Sub BuscaDadosProtocolo()
    Dim strSQL  As String
    Dim ADOTemp As ADODB.Recordset
    
    strSQL = ""

    strSQL = strSQL & " SELECT ASS.strDescricao, CON.strNome AS strContribuinte, PP.strSumula  "

    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrProtocolizacaoProcesso & " PP, "
    strSQL = strSQL & gstrCatalogoAssunto & " ASS, "
    strSQL = strSQL & gstrContribuinte & " CON "

    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PP.PKId = " & dbcintProtocolizacaoDoProcesso.BoundText
    strSQL = strSQL & " AND ASS.PKId = PP.intCodAssunto"
    strSQL = strSQL & " AND CON.PKId = PP.intCodContribuinte"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, ADOTemp) Then
        If Not (ADOTemp.BOF And ADOTemp.EOF) Then
            txt_strCatalogoAssunto.Text = gstrENulo(ADOTemp!strDescricao)
            txt_strContribuinte.Text = gstrENulo(ADOTemp!strContribuinte)
            txt_strSumula.Text = gstrENulo(ADOTemp!strSumula)
        End If
    End If
    Set gobjBanco = Nothing

End Sub

Private Sub dbcintProtocolizacaoDoProcesso_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintProtocolizacaoDoProcesso, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 656
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrAplicar, gstrDeletar, gstrImprimir
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrFechar
End Sub

Private Sub Form_Load()
    
    LeDaTabelaParaObj "", dbcintProtocolizacaoDoProcesso, strQueryProtocolo
    LeDaTabelaParaObj "", dbcintOcorrencia, strQueryOcorrencia
    LeDaTabelaParaObj "", tdb_SuspensaoDeExigencia, strQuerySuspensaoDeExigencia
    
    TrocaCorObjeto txt_strCatalogoAssunto, True
    TrocaCorObjeto txt_strSumula, True
    TrocaCorObjeto txt_strContribuinte, True
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrSalvar
End Sub

Private Function strQuerySuspensaoDeExigencia() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL  As String

    strSQL = ""

'    strSql = strSql & " SELECT SUS.PKId, CON.strNome AS strContribuinte, PAR.intNumeroParcela, PAR.dtmDataVencimento, CASE SUS.bytProcessoFinalizado WHEN 1 THEN 'Sim' ELSE 'Não' END AS strProcessoFinalizado  "
    strSQL = strSQL & " SELECT SUS.PKId, CON.strNome AS strContribuinte, PAR.intNumeroParcela, PAR.dtmDataVencimento, " & gstrCASEWHEN("SUS.bytProcessoFinalizado", "1, 'Sim'", "'Não'") & " AS strProcessoFinalizado  "

    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrSuspensaoDeExigencia & " SUS, "
    strSQL = strSQL & gstrParcelaReceita & " PAR, "
    strSQL = strSQL & gstrContribuinte & " CON "

    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PAR.PKId = SUS.intParcelaReceita "
    strSQL = strSQL & " AND CON.PKId = SUS.intContribuinte "
    strSQL = strSQL & " ORDER BY strContribuinte "
    strQuerySuspensaoDeExigencia = strSQL
    
End Function


Private Function strQueryProtocolo() As String
    Dim strSQL  As String
    
    strSQL = ""

    strSQL = strSQL & "SELECT DISTINCT PP.PKID, PP.strCodigo " & strCONCAT & "'/'" & strCONCAT & " PP.INTEXERCICIO " & strCONCAT & "'-'" & strCONCAT & " PP.BITDIGITO as strCodigo "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrProtocolizacaoProcesso & " PP "
    'strSQL = strSQL & " WHERE "
    
    'strSQL = strSQL & " AND PP.bitRevisaoCalculo = 1 "
    strSQL = strSQL & " ORDER BY strCodigo"
    strQueryProtocolo = strSQL
    
End Function

Private Function strQueryOcorrencia() As String
    Dim strSQL  As String

    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrOcorrencia
    strSQL = strSQL & " WHERE intUtilizacaoDaOcorrencia = 7 " ' 7 SUSPENSÃO DE EXIGÊNCIA
    strSQL = strSQL & " ORDER BY strDescricao"
    strQueryOcorrencia = strSQL
    
End Function

Private Function blnValidaDados() As Boolean
    
    If Not dbcintProtocolizacaoDoProcesso.MatchedWithList Then
        ExibeMensagem " O protocolo tem que ser selecionado."
        dbcintProtocolizacaoDoProcesso.SetFocus
        Exit Function
    End If
    
    If Not dbcintOcorrencia.MatchedWithList Then
        ExibeMensagem "A Ocorrência tem que ser selecionada."
        dbcintOcorrencia.SetFocus
        Exit Function
    End If
    
    If Not dbc_strInscricaoCadastral.MatchedWithList Then
        ExibeMensagem "A Inscrição Cadastral tem que ser selecionada."
        dbc_strInscricaoCadastral.SetFocus
        Exit Function
    End If
    
    If Not dbc_intComposicaoDaReceita.MatchedWithList Then
        ExibeMensagem "A Composição da Receita tem que ser selecionada."
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intNumeroParcela.Text) = "" Then
        ExibeMensagem " O Nº da Parcela Receita tem que ser digitado."
        txt_intNumeroParcela.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intExercicio.Text) = "" Then
        ExibeMensagem "O Exercício tem que ser digitado."
        txt_intExercicio.SetFocus
        Exit Function
    End If
    
    If Trim(txt_dtmDataDoVencimento.Text) = "" Then
        ExibeMensagem "A Data do Vencimento tem que ser digitada."
        txt_dtmDataDoVencimento.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_dtmDataDoVencimento.Text, True) Then
            txt_dtmDataDoVencimento.SetFocus
            Exit Function
    End If
    
    If Trim(txtdtmProcesso.Text) = "" Then
        ExibeMensagem "A Data do Processo tem que ser digitada."
        txtdtmProcesso.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txtdtmProcesso.Text, True) Then
            txtdtmProcesso.SetFocus
            Exit Function
    End If
    
    If Trim(txtstrTextoResultado.Text) = "" Then
        ExibeMensagem "O Resultado do Processo tem que ser digitado."
        txtstrTextoResultado.SetFocus
        Exit Function
    End If
    
    If Not EncontraParcela Then
        ExibeMensagem "Não foi encontrada nenhuma parcela com os dados informados "
        dbc_strInscricaoCadastral.SetFocus
        Exit Function
    End If
    
    blnValidaDados = True
End Function

Private Function EncontraParcela() As Boolean
    Dim strSQL  As String
    Dim ADOTemp As ADODB.Recordset
    
    strSQL = ""

    strSQL = strSQL & " SELECT PAR.PKId,  LAN.intContribuinte"

    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParcelaReceita & " PAR, "
    strSQL = strSQL & gstrLancamentoCalculo & " LAN "
    

    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PAR.intNumeroParcela = " & txt_intNumeroParcela.Text
    strSQL = strSQL & " AND PAR.dtmDataVencimento = " & gstrConvDtParaSql(txt_dtmDataDoVencimento.Text)
    strSQL = strSQL & " AND PAR.intComposicaoDaReceita = " & dbc_intComposicaoDaReceita.BoundText
    strSQL = strSQL & " AND LAN.PKId = PAR.intLancamentoCalculo "
    strSQL = strSQL & " AND LAN.intExercicio = " & txt_intExercicio.Text
    strSQL = strSQL & " AND LAN.strInscricaoCadastral = '" & dbc_strInscricaoCadastral.BoundText & "'"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, ADOTemp) Then
        If Not (ADOTemp.BOF And ADOTemp.EOF) Then
            txtintParcelaReceita.Text = gstrENulo(ADOTemp!Pkid)
            txtintContribuinte.Text = gstrENulo(ADOTemp!intContribuinte)
            EncontraParcela = True
        End If
    End If
    Set gobjBanco = Nothing
    
End Function


Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSQL As String
    Dim intPKIdParcelaReceita As Integer
    
    If UCase(strModoOperacao) = gstrPreencherLista Then
        Dim intAuxIndice As Integer
        If dbcintProtocolizacaoDoProcesso.MatchedWithList Then
            For intAuxIndice = 0 To optbytTipoDeInscricaoCadastral.Count - 1
                If optbytTipoDeInscricaoCadastral(intAuxIndice).Value = True Then
                    Exit For
                End If
            Next
            Select Case intAuxIndice
                Case 4
                    strSQL = ""
                    strSQL = "SELECT DISTINCT REC.intContribuinte, CON.strNome FROM " & gstrReceitaDiversa & " REC, " & gstrContribuinte & " CON, " & gstrProtocolizacaoProcesso & " PRO WHERE PRO.PKID = " & dbcintProtocolizacaoDoProcesso.BoundText & " AND REC.intContribuinte = PRO.intCodContribuinte AND CON.PKId = REC.intContribuinte ORDER BY CON.strNome ;CON.strNome"
                Case Else
                    strSQL = strQueryInscricao(intAuxIndice) & ";A.strInscricaoAnterior"
            End Select
            dbc_strInscricaoCadastral.Tag = strSQL
            PreencherListaDeOpcoes dbc_strInscricaoCadastral
        End If
        Exit Sub
    End If
    

    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        LimpaControlesDoFormulario
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    intPKIdParcelaReceita = Val(txtintParcelaReceita.Text)
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If blnValidaDados Then
            If ToolBarGeral(strModoOperacao, gstrSuspensaoDeExigencia, blnAlterandoSuspensao, tdb_SuspensaoDeExigencia, Me, , strQuerySuspensaoDeExigencia, , , , True) Then
                blnPrimeiraVezSuspensao = False
                LimpaControlesDoFormulario
                If Not blnAlterandoSuspensao Then
                    AtualizaParcelaReceita intPKIdParcelaReceita, 1
                End If
            End If
        End If
    Else
        If ToolBarGeral(strModoOperacao, gstrSuspensaoDeExigencia, blnAlterandoSuspensao, tdb_SuspensaoDeExigencia, Me, , strQuerySuspensaoDeExigencia, , , , True) Then
            If UCase(strModoOperacao) = UCase(gstrDeletar) Then
                blnPrimeiraVezSuspensao = False
                LimpaControlesDoFormulario
                AtualizaParcelaReceita intPKIdParcelaReceita, 0
            End If
        End If
    End If
    
End Sub

Private Sub AtualizaParcelaReceita(intPKID As Integer, intValor As Integer)
    Dim strSQL  As String
    Dim ADOTemp As ADODB.Recordset
    
    strSQL = ""

    strSQL = strSQL & " UPDATE " & gstrParcelaReceita
    strSQL = strSQL & " SET  bytSuspensaoDeExigencia = " & intValor
    strSQL = strSQL & " WHERE PKId = " & intPKID
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSQL
    Set gobjBanco = Nothing
    
End Sub

Sub LimpaControlesDoFormulario()
    txt_strCatalogoAssunto.Text = ""
    txt_strContribuinte.Text = ""
    txt_strSumula.Text = ""
    dbc_strInscricaoCadastral.BoundText = ""
    Set dbc_strInscricaoCadastral.RowSource = Nothing
    dbc_intComposicaoDaReceita.BoundText = ""
    Set dbc_intComposicaoDaReceita.RowSource = Nothing
    txt_intNumeroParcela.Text = ""
    txt_intExercicio.Text = ""
    txt_dtmDataDoVencimento.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnPrimeiraVezSuspensao = False
    blnAlterandoSuspensao = False
End Sub

Private Sub optbytTipoDeInscricaoCadastral_Click(Index As Integer)
Dim strSQL As String
Dim intIndice As Integer
    
    optbytTipoDeInscricaoCadastral(Index).CausesValidation = True

    For intIndice = 0 To 4
        If intIndice <> Index Then
            optbytTipoDeInscricaoCadastral(intIndice).CausesValidation = False
        End If
    Next
    
    Set dbc_strInscricaoCadastral.RowSource = Nothing
    dbc_strInscricaoCadastral.Text = ""
    
    If Index = 4 Then
        lbl_strInscricaoCadastral.Caption = "Contribuinte"
    Else
        lbl_strInscricaoCadastral.Caption = "Inscrição Cadastral"
    End If
    
    If dbcintProtocolizacaoDoProcesso.MatchedWithList Then
        dbc_intComposicaoDaReceita.BoundText = ""
        LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoDaReceita, strQuerryComposicao(Index)
    End If
    
End Sub

Private Function strQuerryComposicao(Index As Integer) As String
    Dim strSQL As String
    Dim Utilizacao As Integer
    
    Utilizacao = 0
    If Index = 0 Or Index = 1 Or Index = 3 Then
        Utilizacao = 1
    ElseIf Index = 2 Then
        Utilizacao = 2
    ElseIf Index = 4 Then
        Utilizacao = 4
    End If
    strSQL = ""
    strSQL = strSQL & " SELECT COM.PKId, COM.strDescricao "
    
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita & " COM, "
    strSQL = strSQL & gstrLancamentoCalculo & " LAN, "
    strSQL = strSQL & gstrProtocolizacaoProcesso & " PRO "
    
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " COM.intUtilizacao = " & Utilizacao
    strSQL = strSQL & " AND COM.PKId = LAN.intComposicaoReceita "
    strSQL = strSQL & " AND LAN.intContribuinte = PRO.intCodContribuinte "
    strSQL = strSQL & " AND PRO.PKId = " & dbcintProtocolizacaoDoProcesso.BoundText
    
    strSQL = strSQL & " ORDER BY COM.strDescricao "
    strQuerryComposicao = strSQL
End Function

Private Function strQueryInscricao(Index As Integer) As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSQL As String
    strSQL = ""
    If Index = 0 Or Index = 1 Then
'        strSQL = strSQL & " SELECT A.strInscricaoAnterior, LTRIM(RTRIM(A.strInscricaoAnterior)) + ' - ' +  LTRIM(RTRIM(C.strNome)) AS Descricao "
        strSQL = strSQL & " SELECT A.strInscricaoAnterior, LTRIM(RTRIM(A.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(C.strNome)) AS Descricao "
    ElseIf Index = 2 Then
'        strSQL = strSQL & " SELECT A.strInscricaoCadastral, LTRIM(RTRIM(A.strInscricaoCadastral)) + ' - ' +  LTRIM(RTRIM(C.strNome)) AS Descricao "
        strSQL = strSQL & " SELECT A.strInscricaoCadastral, LTRIM(RTRIM(A.strInscricaoCadastral)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(C.strNome)) AS Descricao "
    ElseIf Index = 3 Then
'        strSQL = strSQL & " SELECT A.strInscricaoAnterior, LTRIM(RTRIM(A.strInscricaoAnterior)) + ' - ' +  LTRIM(RTRIM(D.strNome)) AS Descricao "
        strSQL = strSQL & " SELECT A.strInscricaoAnterior, LTRIM(RTRIM(A.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(D.strNome)) AS Descricao "
    End If
    
    strSQL = strSQL & " FROM "
    
    If Index = 0 Then
        strSQL = strSQL & gstrImobiliario & " A, "
        strSQL = strSQL & gstrProtocolizacaoProcesso & " B, "
        strSQL = strSQL & gstrContribuinte & " C "
    ElseIf Index = 1 Then
        strSQL = strSQL & gstrImobiliarioRural & " A, "
        strSQL = strSQL & gstrProtocolizacaoProcesso & " B, "
        strSQL = strSQL & gstrContribuinte & " C "
    ElseIf Index = 2 Then
        strSQL = strSQL & gstrEconomico & " A, "
        strSQL = strSQL & gstrProtocolizacaoProcesso & " B, "
        strSQL = strSQL & gstrContribuinte & " C "
    ElseIf Index = 3 Then
        strSQL = strSQL & gstrImobiliario & " A, "
        strSQL = strSQL & gstrContribuicaoMelhoria & " B, "
        strSQL = strSQL & gstrProtocolizacaoProcesso & " C, "
        strSQL = strSQL & gstrContribuinte & " D "
    End If
    strSQL = strSQL & " WHERE "
    If Index = 0 Or Index = 1 Then
        strSQL = strSQL & " A.intContribuinte = B.intCodContribuinte "
        strSQL = strSQL & " AND B.PKId = " & dbcintProtocolizacaoDoProcesso.BoundText
        strSQL = strSQL & " AND C.PKId = B.intCodContribuinte "
        strSQL = strSQL & " ORDER BY Descricao "
    ElseIf Index = 2 Then
        strSQL = strSQL & " A.intContribuinte = B.intCodContribuinte "
        strSQL = strSQL & " AND B.PKId = " & dbcintProtocolizacaoDoProcesso.BoundText
        strSQL = strSQL & " AND C.PKId = B.intCodContribuinte "
        strSQL = strSQL & " ORDER BY Descricao "
    ElseIf Index = 3 Then
        strSQL = strSQL & " B.intImobiliario = A.PKId "
        strSQL = strSQL & " AND A.intContribuinte = C.intCodContribuinte "
        strSQL = strSQL & " AND C.PKId = " & dbcintProtocolizacaoDoProcesso.BoundText
        strSQL = strSQL & " AND D.PKId = C.intCodContribuinte "
        strSQL = strSQL & " ORDER BY Descricao "
    End If
    
strQueryInscricao = strSQL
End Function

Private Sub tdb_SuspensaoDeExigencia_Click()
    blnPrimeiraVezSuspensao = True
End Sub

Private Sub tdb_SuspensaoDeExigencia_FilterChange()
    blnPrimeiraVezSuspensao = False
    gblnFilraCampos tdb_SuspensaoDeExigencia
End Sub

Private Sub tdb_SuspensaoDeExigencia_KeyPress(KeyAscii As Integer)
    Select Case tdb_SuspensaoDeExigencia.Col
        Case 2
            CaracterValido KeyAscii, "N", tdb_SuspensaoDeExigencia
        Case 3
            CaracterValido KeyAscii, "D", tdb_SuspensaoDeExigencia
        Case Else
            CaracterValido KeyAscii, "A", tdb_SuspensaoDeExigencia
    End Select
End Sub

Private Sub tdb_SuspensaoDeExigencia_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim intIndice As Integer
    
    If blnPrimeiraVezSuspensao Then
        With tdb_SuspensaoDeExigencia
            If Not .EOF And Not .BOF Then
                blnAlterandoSuspensao = True
                txtPKId.Text = tdb_SuspensaoDeExigencia.Columns("PKId").Value
                LeDaTabelaParaObj gstrSuspensaoDeExigencia, Me
                For intIndice = 0 To 4
                    If optbytTipoDeInscricaoCadastral(intIndice).Value Then
                        Exit For
                    End If
                Next
                BuscaDadosProtocolo
                optbytTipoDeInscricaoCadastral_Click (intIndice)
                PreencheDadosParcela
                gCorLinhaSelecionada tdb_SuspensaoDeExigencia
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
            End If
        End With
    End If
    
End Sub

Private Sub PreencheDadosParcela()
    Dim strSQL  As String
    Dim ADOTemp As ADODB.Recordset
    
    strSQL = ""

    strSQL = strSQL & " SELECT LAN.strInscricaoCadastral, LAN.intComposicaoReceita, "
    strSQL = strSQL & " PAR.intNumeroParcela, LAN.intExercicio, PAR.dtmDataVencimento "

    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParcelaReceita & " PAR, "
    strSQL = strSQL & gstrLancamentoCalculo & " LAN "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PAR.PKId = " & txtintParcelaReceita.Text
    strSQL = strSQL & " AND LAN.PKId = PAR.intLancamentoCalculo "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, ADOTemp) Then
        If Not (ADOTemp.BOF And ADOTemp.EOF) Then
            dbc_strInscricaoCadastral.BoundText = gstrENulo(ADOTemp!strInscricaoCadastral)
            dbc_intComposicaoDaReceita.BoundText = gstrENulo(ADOTemp!intComposicaoReceita)
            txt_intNumeroParcela.Text = gstrENulo(ADOTemp!intNumeroParcela)
            txt_intExercicio.Text = gstrENulo(ADOTemp!intExercicio)
            txt_dtmDataDoVencimento.Text = gstrENulo(ADOTemp!dtmDataVencimento)
        End If
    End If
    Set gobjBanco = Nothing
    
End Sub


Private Sub txt_dtmDataDoVencimento_GotFocus()
    MarcaCampo txt_dtmDataDoVencimento
End Sub

Private Sub txt_dtmDataDoVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataDoVencimento
End Sub

Private Sub txt_dtmDataDoVencimento_LostFocus()
    txt_dtmDataDoVencimento.Text = gstrDataFormatada(txt_dtmDataDoVencimento.Text)
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txtdtmProcesso_GotFocus()
    MarcaCampo txtdtmProcesso
End Sub

Private Sub txtdtmProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmProcesso
End Sub

Private Sub txtdtmProcesso_LostFocus()
    txtdtmProcesso.Text = gstrDataFormatada(txtdtmProcesso.Text)
End Sub

Private Sub txt_intNumeroParcela_GotFocus()
    MarcaCampo txt_intNumeroParcela
End Sub

Private Sub txt_intNumeroParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intNumeroParcela
End Sub

Private Sub txtstrTextoResultado_GotFocus()
    MarcaCampo txtstrTextoResultado
End Sub

Private Sub txtstrTextoResultado_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrTextoResultado
End Sub



