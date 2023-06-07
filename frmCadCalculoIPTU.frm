VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadCalculoIPTU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tributos"
   ClientHeight    =   7860
   ClientLeft      =   1500
   ClientTop       =   2070
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   9345
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   7755
      Left            =   75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   13679
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tributos"
      TabPicture(0)   =   "frmCadCalculoIPTU.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_inicial"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Final"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "prgStatus"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "mskstrInscricaoAtual"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "mskstrInscricaoFinal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "mskstrInscricaoInicial"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_Identificação"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_Emissao"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_Exercicios"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_ComposicaoDaReceita"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chk_Simulado"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "ISS Construção"
      TabPicture(1)   =   "frmCadCalculoIPTU.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Observacao"
      Tab(1).Control(1)=   "fra_Predios"
      Tab(1).Control(2)=   "txtbitDigitoProcesso"
      Tab(1).Control(3)=   "txtintExercicioProcesso"
      Tab(1).Control(4)=   "txtstrCodigoProcesso"
      Tab(1).Control(5)=   "cmd_ISSInscricao(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt_strComposicaoISS"
      Tab(1).Control(7)=   "dbcstrInscricao"
      Tab(1).Control(8)=   "dbc_intExercicioISS"
      Tab(1).Control(9)=   "tab_3dEnderecos"
      Tab(1).Control(10)=   "dbc_strEmissaoIss"
      Tab(1).Control(11)=   "dbcintContribuinte"
      Tab(1).Control(12)=   "lbl_Contribuinte"
      Tab(1).Control(13)=   "lbl_Emissao(1)"
      Tab(1).Control(14)=   "lblstrDescricao"
      Tab(1).Control(15)=   "lbl_intExercicioISS"
      Tab(1).Control(16)=   "lbl_InscricaoISS"
      Tab(1).Control(17)=   "lbl_ComposicaoISS"
      Tab(1).ControlCount=   18
      Begin VB.Frame fra_Observacao 
         Caption         =   "Observação"
         Height          =   1455
         Left            =   -74895
         TabIndex        =   99
         Top             =   2520
         Width           =   8970
         Begin VB.Frame fra_Impressoes 
            Caption         =   " Opções de impressão"
            Height          =   555
            Left            =   6930
            TabIndex        =   103
            Top             =   150
            Width           =   1920
            Begin VB.CheckBox chk_Carne 
               Caption         =   "Carnê"
               Height          =   195
               Left            =   600
               TabIndex        =   104
               Top             =   270
               Value           =   1  'Checked
               Width           =   975
            End
         End
         Begin VB.TextBox txtstrObservacoes 
            Height          =   1095
            Left            =   90
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   240
            Width           =   6750
         End
      End
      Begin VB.Frame fra_Predios 
         Caption         =   "Prédios"
         Height          =   3630
         Left            =   -74895
         TabIndex        =   79
         Top             =   3990
         Width           =   8970
         Begin VB.TextBox txtstrCategoriaConstrucao 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   2880
            MaxLength       =   21
            TabIndex        =   51
            Top             =   195
            Width           =   2715
         End
         Begin VB.TextBox txtstrPadrao 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   300
            Left            =   6285
            MaxLength       =   21
            TabIndex        =   52
            Top             =   195
            Width           =   2580
         End
         Begin VB.CheckBox chkintDemolicao 
            Caption         =   "Demolição"
            Height          =   240
            Left            =   7770
            TabIndex        =   57
            Top             =   975
            Width           =   1065
         End
         Begin VB.TextBox txtdblIssPagar 
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
            Left            =   7770
            MaxLength       =   21
            TabIndex        =   63
            Top             =   1650
            Width           =   1095
         End
         Begin VB.TextBox txtdblIssAbatimento 
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
            Left            =   4560
            MaxLength       =   21
            TabIndex        =   62
            Top             =   1650
            Width           =   1095
         End
         Begin VB.TextBox txtdblAliquota 
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
            Left            =   7770
            MaxLength       =   21
            TabIndex        =   60
            Top             =   1290
            Width           =   1095
         End
         Begin VB.TextBox txtdblIssDevido 
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
            Left            =   1275
            MaxLength       =   21
            TabIndex        =   61
            Top             =   1650
            Width           =   1095
         End
         Begin VB.TextBox txtdblVlrServico 
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
            Left            =   4560
            MaxLength       =   21
            TabIndex        =   59
            Top             =   1290
            Width           =   1095
         End
         Begin VB.TextBox txtdblVlrM2Servico 
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
            Left            =   1275
            MaxLength       =   21
            TabIndex        =   58
            Top             =   1290
            Width           =   1095
         End
         Begin VB.TextBox txtdtmConstrucao 
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
            Left            =   3375
            MaxLength       =   10
            TabIndex        =   54
            Top             =   570
            Width           =   1080
         End
         Begin VB.TextBox txtdblArea 
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
            Left            =   930
            MaxLength       =   21
            TabIndex        =   53
            Top             =   570
            Width           =   1095
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Predios 
            Height          =   1530
            Left            =   90
            TabIndex        =   93
            Top             =   2010
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   2699
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Pkid"
            Columns(0).DataField=   "Pkid"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nº Edificação"
            Columns(1).DataField=   "intNEdificacao"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Área"
            Columns(2).DataField=   "intMedidaDaArea"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Dt. Construção"
            Columns(3).DataField=   "Dtmultimareforma"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "strCategoriaConstrucao"
            Columns(4).DataField=   "strCategoriaConstrucao"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "strPadrao"
            Columns(5).DataField=   "strPadrao"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "intIssConstrucaoTipo"
            Columns(6).DataField=   "intIssConstrucaoTipo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Tipo"
            Columns(7).DataField=   "strIssConstrucaoTipo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "intIssConstrucaoPadrao"
            Columns(8).DataField=   "intIssConstrucaoPadrao"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Padrão"
            Columns(9).DataField=   "strIssConstrucaoPadrao"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "bitDemolicao"
            Columns(10).DataField=   "bitDemolicao"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "ValorM2Servico"
            Columns(11).DataField=   "ValorM2Servico"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "ValorServico"
            Columns(12).DataField=   "ValorServico"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Aliquota"
            Columns(13).DataField=   "Aliquota"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "IssDevido"
            Columns(14).DataField=   "IssDevido"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "IssAbatimento"
            Columns(15).DataField=   "IssAbatimento"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "IssAPagar"
            Columns(16).DataField=   "IssAPagar"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "PkidArray"
            Columns(17).DataField=   "PkidArray"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(18)._VlistStyle=   0
            Columns(18)._MaxComboItems=   5
            Columns(18).Caption=   "PorcDemolicao"
            Columns(18).DataField=   "PorcDemolicao"
            Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   19
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=19"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1879"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1799"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=2461"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2381"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=2117"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2037"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(29)=   "Column(5).Width=3863"
            Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=3784"
            Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(33)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(35)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(39)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(41)=   "Column(7).Width=3863"
            Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=3784"
            Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(46)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(49)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(50)=   "Column(8).Visible=0"
            Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(52)=   "Column(9).Width=4471"
            Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=4392"
            Splits(0)._ColumnProps(55)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(56)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(57)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(58)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(59)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(60)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(61)=   "Column(10).Visible=0"
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
            Splits(0)._ColumnProps(75)=   "Column(13).Width=2725"
            Splits(0)._ColumnProps(76)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(77)=   "Column(13)._WidthInPix=2646"
            Splits(0)._ColumnProps(78)=   "Column(13)._EditAlways=0"
            Splits(0)._ColumnProps(79)=   "Column(13).Visible=0"
            Splits(0)._ColumnProps(80)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(81)=   "Column(14).Width=2725"
            Splits(0)._ColumnProps(82)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(83)=   "Column(14)._WidthInPix=2646"
            Splits(0)._ColumnProps(84)=   "Column(14)._EditAlways=0"
            Splits(0)._ColumnProps(85)=   "Column(14).Visible=0"
            Splits(0)._ColumnProps(86)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(87)=   "Column(15).Width=2725"
            Splits(0)._ColumnProps(88)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(89)=   "Column(15)._WidthInPix=2646"
            Splits(0)._ColumnProps(90)=   "Column(15)._EditAlways=0"
            Splits(0)._ColumnProps(91)=   "Column(15).Visible=0"
            Splits(0)._ColumnProps(92)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(93)=   "Column(16).Width=2725"
            Splits(0)._ColumnProps(94)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(95)=   "Column(16)._WidthInPix=2646"
            Splits(0)._ColumnProps(96)=   "Column(16)._EditAlways=0"
            Splits(0)._ColumnProps(97)=   "Column(16).Visible=0"
            Splits(0)._ColumnProps(98)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(99)=   "Column(17).Width=2725"
            Splits(0)._ColumnProps(100)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(101)=   "Column(17)._WidthInPix=2646"
            Splits(0)._ColumnProps(102)=   "Column(17)._EditAlways=0"
            Splits(0)._ColumnProps(103)=   "Column(17).Visible=0"
            Splits(0)._ColumnProps(104)=   "Column(17).Order=18"
            Splits(0)._ColumnProps(105)=   "Column(18).Width=2725"
            Splits(0)._ColumnProps(106)=   "Column(18).DividerColor=0"
            Splits(0)._ColumnProps(107)=   "Column(18)._WidthInPix=2646"
            Splits(0)._ColumnProps(108)=   "Column(18)._EditAlways=0"
            Splits(0)._ColumnProps(109)=   "Column(18).Visible=0"
            Splits(0)._ColumnProps(110)=   "Column(18).Order=19"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DataMode        =   4
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=78,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=82,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=79,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=80,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=81,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=86,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=83,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=84,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=85,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=90,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=87,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=88,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=89,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=94,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=91,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=92,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=93,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=98,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=95,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=96,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=97,.parent=17"
            _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=102,.parent=13"
            _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=99,.parent=14"
            _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=100,.parent=15"
            _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=101,.parent=17"
            _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=106,.parent=13"
            _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=103,.parent=14"
            _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=104,.parent=15"
            _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=105,.parent=17"
            _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=110,.parent=13"
            _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=107,.parent=14"
            _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=108,.parent=15"
            _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=109,.parent=17"
            _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=54,.parent=13"
            _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=51,.parent=14"
            _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=52,.parent=15"
            _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=53,.parent=17"
            _StyleDefs(112) =   "Named:id=33:Normal"
            _StyleDefs(113) =   ":id=33,.parent=0"
            _StyleDefs(114) =   "Named:id=34:Heading"
            _StyleDefs(115) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(116) =   ":id=34,.wraptext=-1"
            _StyleDefs(117) =   "Named:id=35:Footing"
            _StyleDefs(118) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(119) =   "Named:id=36:Selected"
            _StyleDefs(120) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(121) =   "Named:id=37:Caption"
            _StyleDefs(122) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(123) =   "Named:id=38:HighlightRow"
            _StyleDefs(124) =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
            _StyleDefs(125) =   "Named:id=39:EvenRow"
            _StyleDefs(126) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(127) =   "Named:id=40:OddRow"
            _StyleDefs(128) =   ":id=40,.parent=33"
            _StyleDefs(129) =   "Named:id=41:RecordSelector"
            _StyleDefs(130) =   ":id=41,.parent=34"
            _StyleDefs(131) =   "Named:id=42:FilterBar"
            _StyleDefs(132) =   ":id=42,.parent=33"
         End
         Begin MSDataListLib.DataCombo dbcintIssConstrucaoTipo 
            Height          =   315
            Left            =   5430
            TabIndex        =   55
            Top             =   555
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintAcabamento 
            Height          =   315
            Left            =   1275
            TabIndex        =   56
            Top             =   915
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintPredios 
            Height          =   315
            Left            =   930
            TabIndex        =   50
            Top             =   195
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lblstrPadrao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Padrão"
            Height          =   195
            Left            =   5715
            TabIndex        =   82
            Top             =   270
            Width           =   510
         End
         Begin VB.Label lblintCategoriaConstrucao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Categoria"
            Height          =   195
            Left            =   2130
            TabIndex        =   81
            Top             =   270
            Width           =   675
         End
         Begin VB.Label lblPredios 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Edificação"
            Height          =   195
            Left            =   135
            TabIndex        =   80
            Top             =   270
            Width           =   750
         End
         Begin VB.Label lbl_dblIssPagar 
            AutoSize        =   -1  'True
            Caption         =   "ISS A Pagar"
            Height          =   195
            Left            =   6840
            TabIndex        =   92
            Top             =   1710
            Width           =   870
         End
         Begin VB.Label lbl_dblIssAbatimento 
            AutoSize        =   -1  'True
            Caption         =   "ISS Abatimento"
            Height          =   195
            Left            =   3360
            TabIndex        =   91
            Top             =   1710
            Width           =   1095
         End
         Begin VB.Label lbl_dblAliquota 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota"
            Height          =   195
            Left            =   7110
            TabIndex        =   89
            Top             =   1350
            Width           =   600
         End
         Begin VB.Label lbl_dblIssDevido 
            AutoSize        =   -1  'True
            Caption         =   "ISS Devido"
            Height          =   195
            Left            =   165
            TabIndex        =   90
            Top             =   1710
            Width           =   810
         End
         Begin VB.Label lbl_VlrServico 
            AutoSize        =   -1  'True
            Caption         =   "Vlr. Serviço"
            Height          =   195
            Left            =   3630
            TabIndex        =   88
            Top             =   1350
            Width           =   810
         End
         Begin VB.Label lbl_dblVlrM2Servico 
            AutoSize        =   -1  'True
            Caption         =   "Vlr. m2 Serviço"
            Height          =   195
            Left            =   150
            TabIndex        =   87
            Top             =   1350
            Width           =   1065
         End
         Begin VB.Label lbl_intAcabamento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Acabamento"
            Height          =   195
            Left            =   135
            TabIndex        =   86
            Top             =   990
            Width           =   900
         End
         Begin VB.Label lbl_intIssConstrucaoTipo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Construção"
            Height          =   195
            Left            =   4560
            TabIndex        =   85
            Top             =   630
            Width           =   810
         End
         Begin VB.Label lbl_DtConstrucao 
            AutoSize        =   -1  'True
            Caption         =   "Data Construção"
            Height          =   195
            Left            =   2115
            TabIndex        =   84
            Top             =   630
            Width           =   1200
         End
         Begin VB.Label lbl_dblArea 
            AutoSize        =   -1  'True
            Caption         =   "Área"
            Height          =   195
            Left            =   150
            TabIndex        =   83
            Top             =   630
            Width           =   330
         End
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
         Left            =   -72615
         MaxLength       =   2
         TabIndex        =   30
         Top             =   765
         Width           =   285
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
         Left            =   -73125
         MaxLength       =   4
         TabIndex        =   29
         Top             =   765
         Width           =   465
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
         Left            =   -73995
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   765
         Width           =   825
      End
      Begin VB.CommandButton cmd_ISSInscricao 
         Height          =   300
         Index           =   2
         Left            =   -68055
         Picture         =   "frmCadCalculoIPTU.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativar Cadastro Imobiliário"
         Top             =   405
         Width           =   360
      End
      Begin VB.TextBox txt_strComposicaoISS 
         Height          =   285
         Left            =   -73980
         MaxLength       =   50
         TabIndex        =   21
         Top             =   420
         Width           =   2925
      End
      Begin VB.CheckBox chk_Simulado 
         Caption         =   "Simulado"
         Height          =   285
         Left            =   7680
         TabIndex        =   13
         Top             =   1095
         Width           =   1035
      End
      Begin VB.Frame fra_ComposicaoDaReceita 
         Caption         =   "Composição da Receita"
         Height          =   780
         Left            =   585
         TabIndex        =   1
         Top             =   975
         Width           =   5460
         Begin VB.CommandButton cmd_Composicao 
            Height          =   300
            Left            =   4965
            Picture         =   "frmCadCalculoIPTU.frx":0156
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro de Composição da Receita"
            Top             =   315
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbc_intComposicao 
            Height          =   315
            Left            =   1140
            TabIndex        =   3
            Top             =   315
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_Composicao 
            AutoSize        =   -1  'True
            Caption         =   "Composição"
            Height          =   195
            Left            =   165
            TabIndex        =   2
            Top             =   360
            Width           =   870
         End
      End
      Begin VB.Frame fra_Exercicios 
         Height          =   690
         Left            =   6720
         TabIndex        =   17
         Top             =   3150
         Width           =   1905
         Begin MSDataListLib.DataCombo dbc_intExercicioInicial 
            Height          =   315
            Left            =   840
            TabIndex        =   19
            Top             =   255
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intExercicioFinal 
            Height          =   315
            Left            =   3480
            TabIndex        =   94
            Top             =   255
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_ExercicioFinal 
            AutoSize        =   -1  'True
            Caption         =   "Exercício Final"
            Height          =   195
            Left            =   2370
            TabIndex        =   95
            Top             =   330
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lbl_ExercicioInicial 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   105
            TabIndex        =   18
            Top             =   330
            Width           =   675
         End
      End
      Begin VB.Frame fra_Emissao 
         Caption         =   "Emissão"
         Height          =   690
         Left            =   6735
         TabIndex        =   14
         Top             =   2400
         Width           =   1875
         Begin MSDataListLib.DataCombo dbc_strEmissao 
            Height          =   315
            Left            =   765
            TabIndex        =   16
            Top             =   255
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_Emissao 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   15
            Top             =   345
            Width           =   585
         End
      End
      Begin VB.Frame fra_Identificação 
         Caption         =   "Identificação"
         Height          =   1635
         Left            =   555
         TabIndex        =   5
         Top             =   2400
         Width           =   5460
         Begin VB.CheckBox chk_Critica 
            Caption         =   "Visualizar resumo por Inscrição"
            Height          =   285
            Left            =   1995
            TabIndex        =   102
            Top             =   1290
            Width           =   2625
         End
         Begin VB.CommandButton cmd_Inscricao 
            Height          =   300
            Index           =   0
            Left            =   4965
            Picture         =   "frmCadCalculoIPTU.frx":0274
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro"
            Top             =   300
            Width           =   360
         End
         Begin VB.CommandButton cmd_Inscricao 
            Height          =   300
            Index           =   1
            Left            =   4965
            Picture         =   "frmCadCalculoIPTU.frx":0392
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Tag             =   "590"
            ToolTipText     =   "Ativa Cadastro"
            Top             =   705
            Width           =   360
         End
         Begin VB.CheckBox chk_SelecionarTodos 
            Caption         =   "Selecionar Todas as Inscrições"
            Height          =   225
            Left            =   1995
            TabIndex        =   12
            Top             =   1065
            Width           =   3240
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoInicial 
            Height          =   315
            Left            =   2010
            TabIndex        =   7
            Top             =   300
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoFinal 
            Height          =   315
            Left            =   2010
            TabIndex        =   10
            Top             =   690
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_strInscricaoCadastralInicial 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Inicial"
            Height          =   195
            Left            =   75
            TabIndex        =   6
            Top             =   405
            Width           =   1800
         End
         Begin VB.Label lbl_strInscricaoCadastralFinal 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Final"
            Height          =   195
            Left            =   75
            TabIndex        =   9
            Top             =   765
            Width           =   1725
         End
      End
      Begin MSMask.MaskEdBox mskstrInscricaoInicial 
         Height          =   300
         Left            =   2340
         TabIndex        =   96
         Top             =   30
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskstrInscricaoFinal 
         Height          =   300
         Left            =   3645
         TabIndex        =   97
         Top             =   30
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskstrInscricaoAtual 
         Height          =   300
         Left            =   4860
         TabIndex        =   98
         Top             =   30
         Visible         =   0   'False
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin MSDataListLib.DataCombo dbcstrInscricao 
         Height          =   315
         Left            =   -70200
         TabIndex        =   23
         Top             =   405
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intExercicioISS 
         Height          =   315
         Left            =   -66810
         TabIndex        =   26
         Top             =   405
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin TabDlg.SSTab tab_3dEnderecos 
         Height          =   1350
         Left            =   -74895
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1140
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   2381
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Endereço"
         TabPicture(0)   =   "frmCadCalculoIPTU.frx":04B0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra_EndImobiliario"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Endereço de Notificação"
         TabPicture(1)   =   "frmCadCalculoIPTU.frx":04CC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame1 
            Height          =   930
            Left            =   -74850
            TabIndex        =   64
            Top             =   315
            Width           =   8655
            Begin VB.TextBox txtstrMunicipioC 
               Height          =   300
               Left            =   4185
               MaxLength       =   50
               TabIndex        =   74
               Top             =   540
               Width           =   2235
            End
            Begin VB.TextBox txtstrUFC 
               Height          =   300
               Left            =   6765
               MaxLength       =   2
               TabIndex        =   76
               Top             =   540
               Width           =   375
            End
            Begin VB.TextBox txtstrNumeroC 
               Height          =   300
               Left            =   5475
               MaxLength       =   10
               TabIndex        =   68
               Top             =   180
               Width           =   825
            End
            Begin VB.TextBox txtintCepC 
               Height          =   300
               Left            =   7560
               MaxLength       =   9
               TabIndex        =   78
               Top             =   525
               Width           =   1005
            End
            Begin VB.TextBox txtstrComplementoC 
               Height          =   300
               Left            =   6960
               MaxLength       =   10
               TabIndex        =   70
               Top             =   180
               Width           =   1590
            End
            Begin VB.TextBox txtstrLogradouroC 
               Height          =   300
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   66
               Top             =   180
               Width           =   4065
            End
            Begin VB.TextBox txtstrBairroC 
               Height          =   300
               Left            =   675
               MaxLength       =   50
               TabIndex        =   72
               Top             =   540
               Width           =   2670
            End
            Begin VB.Label lbl_MunicipioC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   3435
               TabIndex        =   73
               Top             =   615
               Width           =   705
            End
            Begin VB.Label lbl_UFC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   6495
               TabIndex        =   75
               Top             =   630
               Width           =   210
            End
            Begin VB.Label lbl_CepC 
               AutoSize        =   -1  'True
               Caption         =   "CEP"
               Height          =   195
               Left            =   7200
               TabIndex        =   77
               Top             =   615
               Width           =   315
            End
            Begin VB.Label lbl_ComplementoC 
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6435
               TabIndex        =   69
               Top             =   270
               Width           =   480
            End
            Begin VB.Label lbl_NumeroC 
               AutoSize        =   -1  'True
               Caption         =   "N°"
               Height          =   195
               Left            =   5250
               TabIndex        =   67
               Top             =   270
               Width           =   180
            End
            Begin VB.Label lbl_BairroC 
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   210
               TabIndex        =   71
               Top             =   630
               Width           =   405
            End
            Begin VB.Label lbl_LogradouroC 
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   195
               TabIndex        =   65
               Top             =   270
               Width           =   810
            End
         End
         Begin VB.Frame fra_EndImobiliario 
            Height          =   930
            Left            =   150
            TabIndex        =   34
            Top             =   315
            Width           =   8655
            Begin VB.TextBox txtstrBairro 
               Height          =   300
               Left            =   675
               MaxLength       =   50
               TabIndex        =   42
               Top             =   540
               Width           =   2670
            End
            Begin VB.TextBox txtstrLogradouro 
               Height          =   300
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   36
               Top             =   180
               Width           =   4065
            End
            Begin VB.TextBox txtstrComplemento 
               Height          =   300
               Left            =   6960
               MaxLength       =   10
               TabIndex        =   40
               Top             =   180
               Width           =   1590
            End
            Begin VB.TextBox txtintCep 
               Height          =   300
               Left            =   7560
               MaxLength       =   9
               TabIndex        =   48
               Top             =   525
               Width           =   1005
            End
            Begin VB.TextBox txtstrNumero 
               Height          =   300
               Left            =   5475
               MaxLength       =   10
               TabIndex        =   38
               Top             =   180
               Width           =   825
            End
            Begin VB.TextBox txtstrMunicipio 
               Height          =   300
               Left            =   4185
               MaxLength       =   50
               TabIndex        =   44
               Top             =   540
               Width           =   2235
            End
            Begin VB.TextBox txtstrUf 
               Height          =   300
               Left            =   6765
               MaxLength       =   2
               TabIndex        =   46
               Top             =   540
               Width           =   375
            End
            Begin VB.Label lblintLogradouro 
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   195
               TabIndex        =   35
               Top             =   270
               Width           =   810
            End
            Begin VB.Label lblintBairro 
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   210
               TabIndex        =   41
               Top             =   630
               Width           =   405
            End
            Begin VB.Label lblintNumero 
               AutoSize        =   -1  'True
               Caption         =   "N°"
               Height          =   195
               Left            =   5250
               TabIndex        =   37
               Top             =   270
               Width           =   180
            End
            Begin VB.Label lblstrComplemento 
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6435
               TabIndex        =   39
               Top             =   270
               Width           =   480
            End
            Begin VB.Label lblintCep 
               AutoSize        =   -1  'True
               Caption         =   "CEP"
               Height          =   195
               Left            =   7200
               TabIndex        =   47
               Top             =   615
               Width           =   315
            End
            Begin VB.Label lblstrMunicipio 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   3435
               TabIndex        =   43
               Top             =   615
               Width           =   705
            End
            Begin VB.Label lblstrUf 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   6495
               TabIndex        =   45
               Top             =   630
               Width           =   210
            End
         End
      End
      Begin MSDataListLib.DataCombo dbc_strEmissaoIss 
         Height          =   315
         Left            =   -66810
         TabIndex        =   32
         Top             =   765
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintContribuinte 
         Height          =   315
         Left            =   -71220
         TabIndex        =   101
         Top             =   765
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSComctlLib.ProgressBar prgStatus 
         Height          =   300
         Left            =   720
         TabIndex        =   105
         Top             =   4500
         Visible         =   0   'False
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lbl_Final 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   6420
         TabIndex        =   107
         Top             =   4860
         Width           =   1995
      End
      Begin VB.Label lbl_inicial 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   720
         TabIndex        =   106
         Top             =   4830
         Width           =   885
      End
      Begin VB.Label lbl_Contribuinte 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Contribuinte"
         Height          =   195
         Left            =   -72090
         TabIndex        =   100
         Top             =   855
         Width           =   840
      End
      Begin VB.Label lbl_Emissao 
         AutoSize        =   -1  'True
         Caption         =   "Emissão"
         Height          =   195
         Index           =   1
         Left            =   -67425
         TabIndex        =   31
         Top             =   855
         Width           =   585
      End
      Begin VB.Label lblstrDescricao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Processo"
         Height          =   195
         Left            =   -74700
         TabIndex        =   27
         Top             =   825
         Width           =   660
      End
      Begin VB.Label lbl_intExercicioISS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   -67515
         TabIndex        =   25
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lbl_InscricaoISS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição"
         Height          =   195
         Left            =   -70875
         TabIndex        =   22
         Top             =   480
         Width           =   645
      End
      Begin VB.Label lbl_ComposicaoISS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Composição"
         Height          =   195
         Left            =   -74910
         TabIndex        =   20
         Top             =   480
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmCadCalculoIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnModoConsultaEmissao                    As Boolean
Dim bytTipoComposicao                         As Byte

Dim vetRelLanctoDevolver                      As XArrayDB

Dim vetPredios                                As XArrayDB

Dim PkidArray                                 As Integer

Private Const BYT_CALCULOIMPOSTO_OK           As Byte = 0
Private Const BYT_CALCULOIMPOSTO_ERRO_RECEITA As Byte = 1
Private Const BYT_CALCULOIMPOSTO_ERRO_GERAL   As Byte = 2

Private Const PREDIO_PKID                     As Byte = 0
Private Const PREDIO_NEDIFICACAO              As Byte = 1
Private Const PREDIO_MEDIDAAREA               As Byte = 2
Private Const PREDIO_DATACONSTRUCAO           As Byte = 3
Private Const PREDIO_STRCATEGORIACONSTRUCAO   As Byte = 4
Private Const PREDIO_PADRAO                   As Byte = 5
Private Const PREDIO_INTISSCONSTRUCAOTIPO     As Byte = 6
Private Const PREDIO_STRISSCONSTRUCAOTIPO     As Byte = 7
Private Const PREDIO_INTISSCONSTRUCAOPADRAO   As Byte = 8
Private Const PREDIO_STRISSCONSTRUCAOPADRAO   As Byte = 9
Private Const PREDIO_DEMOLICAO                As Byte = 10
Private Const PREDIO_VALORM2SERVICO           As Byte = 11
Private Const PREDIO_VALORSERVICO             As Byte = 12
Private Const PREDIO_ALIQUOTA                 As Byte = 13
Private Const PREDIO_ISSDEVIDO                As Byte = 14
Private Const PREDIO_ISSABATIMENTO            As Byte = 15
Private Const PREDIO_ISSAPAGAR                As Byte = 16
Private Const PREDIO_PKIDARRAY                As Byte = 17
Private Const PREDIO_PORCDEMOLICAO            As Byte = 18
Private Const PREDIO_MEDIDAAREAORIG           As Byte = 19

Private Sub chkintDemolicao_Click()
    CarregaValoresDoIssConstrucao
End Sub

Private Sub cmd_ISSInscricao_Click(Index As Integer)
    'CarregaForm frmCadImobiliario, dbcstrInscricao
    ChamaFormCadastro frmCadImobiliario, dbcstrInscricao
End Sub

Private Sub dbc_strEmissaoIss_Click(Area As Integer)
    DropDownDataCombo dbc_strEmissaoIss, Me, Area
End Sub

Private Sub dbc_strEmissaoIss_GotFocus()
    MarcaCampo dbc_strEmissaoIss
End Sub

Private Sub dbc_strEmissaoIss_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strEmissaoIss, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strEmissaoIss_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_strEmissaoIss
End Sub

Private Sub dbcintAcabamento_Change()
    If dbcintAcabamento.MatchedWithList Then
        CarregaValoresDoIssConstrucao
    End If
End Sub

Private Sub dbcintAcabamento_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintAcabamento, Me, Area
End Sub

Private Sub dbcintIssConstrucaoTipo_Change()
    If dbcintIssConstrucaoTipo.MatchedWithList Then
        dbcintAcabamento.Tag = strQueryIssConstrucaoPadrao & ";strDescricao"
        PreencherListaDeOpcoes dbcintAcabamento
        CarregaValoresDoIssConstrucao
    End If
End Sub

Private Sub dbcintIssConstrucaoTipo_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcintIssConstrucaoTipo, Me, Area
End Sub

Private Sub dbcintPredios_Change()
Dim strSQL       As String
Dim adoRec       As ADODB.Recordset
Dim adoPontuacao As ADODB.Recordset
    
    If dbcintPredios.MatchedWithList Then
        
        dbcintIssConstrucaoTipo.BoundText = Space$(0)
        dbcintAcabamento.BoundText = Space$(0)
        chkintDemolicao.Value = vbUnchecked
        chkintDemolicao.Tag = ""
        txtdblVlrM2Servico = Space$(0)
        txtdblVlrServico = Space$(0)
        txtdblAliquota = Space$(0)
        txtdblIssDevido = Space$(0)
        txtdblIssAbatimento = Space$(0)
        txtdblIssPagar = Space$(0)
        
        strSQL = ""
        strSQL = strSQL & "SELECT AI.Pkid, AI.intMedidaDaArea, AI.dtmUltimaReforma, AI.intCategoriaConstrucao, CC.strDescricao, "
        strSQL = strSQL & "SUM(TV.DBLVALOR) Pontos "
        strSQL = strSQL & "FROM " & gstrAreaImobiliario & " AI, " & gstrTipoDeArea & " TA, " & gstrCaracteristicaDoImovel & " CI, " & gstrDetalheDaCaracteristica & " DC, "
        strSQL = strSQL & gstrTabelaDeValor & " TV, " & gstrCategoriaConstrucao & " CC "
        strSQL = strSQL & "WHERE AI.Pkid = " & dbcintPredios.BoundText & _
                          " AND TA.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " AI.intTipoDeArea " & _
                          " AND CI.INTCODIGOIMOBILIARIO = AI.intImobiliario " & _
                          " AND DC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CI.INTCODIGODETALHEDACARACTERISTI " & _
                          " AND TV.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " DC.INTTABELADEVALORES " & _
                          " AND CC.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " AI.INTCATEGORIACONSTRUCAO " & _
                          " AND CI.intArea = AI.Pkid "
        strSQL = strSQL & "GROUP BY AI.Pkid, AI.intMedidaDaArea, AI.dtmUltimaReforma, AI.intCategoriaConstrucao, CC.strDescricao"
    
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoRec) Then
            If Not adoRec.EOF Then
            
                strSQL = ""
                strSQL = strSQL & "SELECT FP.strDescricao FROM " & gstrFaixaPontosPredio & " FP WHERE FP.Intcategoriaconstrucao = " & adoRec("intCategoriaConstrucao").Value & " AND (dblPontoinicial <= " & adoRec("Pontos").Value & " AND dblPontoFinal >= " & adoRec("Pontos").Value & ")"
            
                If gobjBanco.CriaADO(strSQL, 5, adoPontuacao) Then
        
                    txtdblArea.Text = gstrConvVrDoSql(gstrENulo(adoRec("intMedidaDaArea").Value), , , True)
                    txtdblArea.Tag = gstrConvVrDoSql(gstrENulo(adoRec("intMedidaDaArea").Value), , , True)
                    txtdtmConstrucao.Text = gstrENulo(adoRec("dtmUltimaReforma").Value)
                    txtstrCategoriaConstrucao.Text = gstrENulo(adoRec("strDescricao").Value)
                    txtstrPadrao.Text = gstrENulo(adoPontuacao("strDescricao").Value)
        
                End If
            End If
        End If
    End If
    
End Sub

Private Sub dbcstrInscricao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcstrInscricao, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 664
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir, gstrSalvar, gstrDeletar
    HabilitaDesabilitaBotao1 tab_3dPasta.Tab = 1, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub

Private Sub Form_Load()

    dbc_intComposicao.Tag = strQueryComposicao & ";strDescricao"
    dbcintIssConstrucaoTipo.Tag = strQueryIssConstrucaoTipo & ";strDescricao"
    dbcintContribuinte.Tag = strQueryDataComboContribuinte & ";strnome"

    Set vetPredios = New XArrayDB
    vetPredios.Clear
    vetPredios.ReDim 0, 0, 0, 19
    PkidArray = 0
    
    blnModoConsultaEmissao = False
    
    tab_3dPasta.TabEnabled(1) = False
    
    TrocaCorObjeto txtstrCategoriaConstrucao, True
    TrocaCorObjeto txtstrPadrao, True
    TrocaCorObjeto txt_strComposicaoISS, True
    TrocaCorObjeto txtdblVlrM2Servico, True
    TrocaCorObjeto txtdblVlrServico, True
    TrocaCorObjeto txtdblAliquota, True
    TrocaCorObjeto txtdblIssDevido, True
    TrocaCorObjeto txtdblIssPagar, True
    
    TrocaCorObjeto txtstrLogradouro, True
    TrocaCorObjeto txtstrNumero, True
    TrocaCorObjeto txtstrComplemento, True
    TrocaCorObjeto txtintCep, True
    TrocaCorObjeto txtstrBairro, True
    TrocaCorObjeto txtstrMunicipio, True
    TrocaCorObjeto txtstrUf, True
    'TrocaCorObjeto txtstrLogradouroC, True
    'TrocaCorObjeto txtstrNumeroC, True
    'TrocaCorObjeto txtstrComplementoC, True
    'TrocaCorObjeto txtstrBairroC, True
    'TrocaCorObjeto txtstrMunicipioC, True
    'TrocaCorObjeto txtstrUFC, True
    'TrocaCorObjeto txtintCepC, True
    TrocaCorObjeto dbcintContribuinte, True
End Sub

Private Sub dbc_intComposicao_Change()

    LimpaDataCombo dbc_strEmissao
    LimpaDataCombo dbc_intExercicioInicial
    LimpaDataCombo dbc_intExercicioFinal
    LimpaDataCombo dbc_strInscricaoInicial
    LimpaDataCombo dbc_strInscricaoFinal
    LimpaDataCombo dbc_strEmissaoIss
    LimpaDataCombo dbc_intExercicioISS
    LimpaDataCombo dbcstrInscricao

    If dbc_intComposicao.MatchedWithList Then
        
        DefineComposicao dbc_intComposicao.BoundText
                
        TrocaCorObjeto dbc_strEmissao, bytTipoComposicao = TYP_ISS_CONSTRUCAO
        TrocaCorObjeto dbc_strInscricaoInicial, bytTipoComposicao = TYP_ISS_CONSTRUCAO
        TrocaCorObjeto dbc_strInscricaoFinal, bytTipoComposicao = TYP_ISS_CONSTRUCAO
        TrocaCorObjeto dbc_intExercicioInicial, bytTipoComposicao = TYP_ISS_CONSTRUCAO
        'TrocaCorObjeto chk_SelecionarTodos, bytTipoComposicao = TYP_ISS_CONSTRUCAO
        chk_SelecionarTodos.Enabled = Not bytTipoComposicao = TYP_ISS_CONSTRUCAO
        TrocaCorObjeto cmd_Inscricao(0), bytTipoComposicao = TYP_ISS_CONSTRUCAO
        TrocaCorObjeto cmd_Inscricao(1), bytTipoComposicao = TYP_ISS_CONSTRUCAO
        
        tab_3dPasta.TabEnabled(1) = bytTipoComposicao = TYP_ISS_CONSTRUCAO
        
        PreencheEmissao
        
        'Caso seja do tipo ISS Contrucao vamos preparar a outra guia
        If bytTipoComposicao = TYP_ISS_CONSTRUCAO Then
            
            tab_3dPasta.Tab = 1
            dbcstrInscricao.SetFocus
            
            txt_strComposicaoISS = dbc_intComposicao.Text
            
            PreencheExercicio dbc_intExercicioISS
            
            dbcstrInscricao.Tag = strQueryInscricao & ";strInscricao"
            
        Else
            
            PreencheExercicio dbc_intExercicioInicial
            
            dbc_strInscricaoInicial.Tag = strQueryInscricao & IIf(bytTipoComposicao = TYP_IMOBILIARIA Or bytTipoComposicao = TYP_OUTROS, ";strInscricao", ";strInscricaoCadastral")
            dbc_strInscricaoFinal.Tag = strQueryInscricao & IIf(bytTipoComposicao = TYP_IMOBILIARIA Or bytTipoComposicao = TYP_OUTROS, ";strInscricao", ";strInscricaoCadastral")
            
        End If
    
    End If
    
End Sub

Private Sub dbc_intComposicao_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbc_intComposicao, Me, Area
End Sub

Private Sub dbc_intComposicao_GotFocus()
    MarcaCampo dbc_intComposicao
End Sub

Private Sub dbc_intComposicao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicao
End Sub

Private Sub dbc_strInscricaoInicial_GotFocus()
    MarcaCampo dbc_strInscricaoInicial
End Sub

Private Sub dbc_strInscricaoInicial_Change()
    mskstrInscricaoInicial = dbc_strInscricaoInicial.Text
End Sub

Private Sub dbc_strInscricaoInicial_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbc_strInscricaoInicial, Me, Area
End Sub

Private Sub dbc_strInscricaoFinal_GotFocus()
    MarcaCampo dbc_strInscricaoFinal
End Sub

Private Sub dbc_strInscricaoFinal_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbc_strInscricaoFinal, Me, Area
End Sub

Private Sub dbc_strInscricaoFinal_Change()
    mskstrInscricaoFinal = dbc_strInscricaoFinal.Text
End Sub

Private Sub dbcstrInscricao_GotFocus()
    MarcaCampo dbcstrInscricao
End Sub

Private Sub dbcstrInscricao_Change()
    If dbcstrInscricao.MatchedWithList Then
        LimpaDadosISSConstrucao True
        PreencheDadosCadastro
    End If
End Sub

Private Sub dbcstrInscricao_Click(Area As Integer)
    If Area = 0 Then DropDownDataCombo dbcstrInscricao, Me, Area
End Sub

Private Sub chk_SelecionarTodos_Click()
    TrocaCorObjeto dbc_strInscricaoInicial, chk_SelecionarTodos.Value
    TrocaCorObjeto dbc_strInscricaoFinal, chk_SelecionarTodos.Value
End Sub

Private Sub cmd_Composicao_Click()
    ChamaFormCadastro frmCadComposicaoDaReceita, dbc_intComposicao
End Sub

Private Sub cmd_Inscricao_Click(Index As Integer)
    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "É necessário selecionar alguma Composição da Receita."
        Exit Sub
    End If
    
'    If Index = 2 Then
'        ChamaFormCadastro frmCadImobiliario, dbcstrInscricao
'        'CarregaForm frmCadImobiliario, dbcstrInscricao
'        Exit Sub
'    End If

    If bytTipoComposicao = TYP_IMOBILIARIA Or bytTipoComposicao = TYP_OUTROS Then
        If Index = 0 Then
            ChamaFormCadastro frmCadImobiliario, dbc_strInscricaoInicial
        Else
            ChamaFormCadastro frmCadImobiliario, dbc_strInscricaoFinal
        End If
    Else
        If Index = 0 Then
            ChamaFormCadastro frmCadEconomico, dbc_strInscricaoInicial
        Else
            ChamaFormCadastro frmCadEconomico, dbc_strInscricaoFinal
        End If
    End If
End Sub

Private Sub dbc_strEmissao_Click(Area As Integer)
    DropDownDataCombo dbc_strEmissao, Me, Area
End Sub

Private Sub dbc_strEmissao_GotFocus()
    MarcaCampo dbc_strEmissao
End Sub

Private Sub dbc_strEmissao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strEmissao, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strEmissao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_strEmissao
End Sub

Private Sub dbc_strEmissao_Change()
    TrocaCorObjeto dbc_strInscricaoInicial, dbc_strEmissao.Text <> ""
    TrocaCorObjeto dbc_strInscricaoFinal, dbc_strEmissao.Text <> ""
    chk_SelecionarTodos.Enabled = Not dbc_strEmissao.Text <> ""
    blnModoConsultaEmissao = dbc_strEmissao.Text <> ""
End Sub

Private Sub dbc_intExercicioInicial_Click(Area As Integer)
    DropDownDataCombo dbc_intExercicioInicial, Me, Area
End Sub

Private Sub dbc_intExercicioInicial_GotFocus()
    MarcaCampo dbc_intExercicioInicial
End Sub

Private Sub dbc_intExercicioInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intExercicioInicial, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intExercicioInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intExercicioInicial
End Sub

Private Sub dbc_intExercicioFinal_Click(Area As Integer)
    DropDownDataCombo dbc_intExercicioFinal, Me, Area
End Sub

Private Sub dbc_intExercicioFinal_GotFocus()
    MarcaCampo dbc_intExercicioFinal
End Sub

Private Sub dbc_intExercicioFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intExercicioFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intExercicioFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intExercicioFinal
End Sub

Private Sub dbc_intExercicioISS_Click(Area As Integer)
    DropDownDataCombo dbc_intExercicioISS, Me, Area
End Sub

Private Sub dbc_intExercicioISS_GotFocus()
    MarcaCampo dbc_intExercicioISS
End Sub

Private Sub dbc_intExercicioISS_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intExercicioISS, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intExercicioISS_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intExercicioISS
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    
    HabilitaDesabilitaBotao1 PreviousTab = 0, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    
End Sub

Private Sub txtdblArea_GotFocus()
    MarcaCampo txtdblArea
End Sub

Private Sub txtdblArea_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblArea
End Sub

Private Sub txtdblArea_LostFocus()
    txtdblArea = gstrConvVrDoSql(txtdblArea)
    If Len(txtdblArea.Text) > 0 And Len(txtdblVlrM2Servico.Text) > 0 Then
        txtdblIssDevido.Text = gstrConvVrDoSql((txtdblArea.Text * txtdblVlrM2Servico.Text * (txtdblAliquota.Text / 100)), 2)
        txtdblVlrServico.Text = gstrConvVrDoSql(txtdblVlrM2Servico.Text * txtdblArea.Text, 2)
        If Len(txtdblIssAbatimento.Text) > 0 Then
             txtdblIssPagar.Text = gstrConvVrDoSql(txtdblIssDevido.Text - txtdblIssAbatimento.Text, 2)
        Else
            txtdblIssPagar.Text = gstrConvVrDoSql(txtdblIssDevido.Text, 2)
        End If
    End If
End Sub

Private Sub tdb_Predios_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Predios
        If Not .EOF And Not .BOF Then
            gCorLinhaSelecionada tdb_Predios
                        
            dbcintPredios.Text = .Columns("intNEdificacao").Value
            txtstrCategoriaConstrucao.Text = .Columns("strCategoriaConstrucao").Value
            txtstrPadrao.Text = .Columns("strPadrao").Value
            txtdblArea.Text = .Columns("intMedidaDaArea").Value
            txtdblArea.Tag = ""
            txtdtmConstrucao.Text = .Columns("dtmUltimaReforma").Value
            LeDaTabelaParaObj gstrIssConstrucaoTipo, dbcintIssConstrucaoTipo, strQueryIssConstrucaoTipo
            dbcintIssConstrucaoTipo.BoundText = .Columns("intIssConstrucaoTipo").Value
            'LeDaTabelaParaObj gstrIssConstrucaoPadrao, dbcintAcabamento, strQueryIssConstrucaoPadrao
            dbcintAcabamento.BoundText = .Columns("intIssConstrucaoPadrao").Value
            chkintDemolicao.Value = Val(.Columns("bitDemolicao").Value)
            chkintDemolicao.Tag = Val(.Columns("PorcDemolicao").Value)
            txtdblVlrM2Servico.Text = .Columns("ValorM2Servico").Value
            txtdblVlrServico.Text = .Columns("ValorServico").Value
            txtdblAliquota.Text = .Columns("Aliquota").Value
            txtdblIssDevido.Text = .Columns("IssDevido").Value
            txtdblIssAbatimento.Text = .Columns("IssAbatimento").Value
            txtdblIssPagar.Text = .Columns("IssAPagar").Value
            
        End If
    End With

End Sub

Private Sub txtbitDigitoProcesso_GotFocus()
    MarcaCampo txtbitDigitoProcesso
End Sub

Private Sub txtbitDigitoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigitoProcesso
End Sub

Private Sub txtdblIssAbatimento_GotFocus()
    MarcaCampo txtdblIssAbatimento
End Sub

Private Sub txtdblIssAbatimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblIssAbatimento
End Sub

Private Sub txtdblIssAbatimento_LostFocus()
    txtdblIssAbatimento = gstrConvVrDoSql(txtdblIssAbatimento)
    If Len(txtdblIssAbatimento.Text) > 0 Then
        txtdblIssPagar.Text = gstrConvVrDoSql(txtdblIssDevido.Text - txtdblIssAbatimento.Text, 2)
    Else
        txtdblIssPagar.Text = gstrConvVrDoSql(txtdblIssDevido.Text, 2)
    End If
End Sub

Private Sub txtdtmConstrucao_GotFocus()
    MarcaCampo txtdtmConstrucao
End Sub

Private Sub txtdtmConstrucao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmConstrucao
End Sub

Private Sub txtdtmConstrucao_LostFocus()
    txtdtmConstrucao = gstrDataFormatada(txtdtmConstrucao)
    If Len(txtdtmConstrucao.Text) > 0 Then CarregaValoresDoIssConstrucao
End Sub

Private Sub txtstrCodigoProcesso_GotFocus()
    MarcaCampo txtstrCodigoProcesso
End Sub

Private Sub txtstrCodigoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigoProcesso
End Sub

Private Sub txtintExercicioProcesso_GotFocus()
    MarcaCampo txtintExercicioProcesso
End Sub

Private Sub txtintExercicioProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicioProcesso
End Sub

Private Sub txtstrNumero_GotFocus()
    MarcaCampo txtstrNumero
End Sub

Private Sub txtstrNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumero
End Sub

Private Sub txtstrComplemento_GotFocus()
    MarcaCampo txtstrComplemento
End Sub

Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplemento
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub

Private Sub txtintCEP_LostFocus()
    If txtintCep.Text = "" Then
        Exit Sub
    End If
    
    txtintCep = gstrCEPFormatado(txtintCep)
    CepLogradouro txtintCep, txtstrLogradouro, txtstrBairro, , txtstrUf, , , , False, False, False, False, False, False
    
End Sub

Private Sub txtstrNumeroC_GotFocus()
    MarcaCampo txtstrNumeroC
End Sub

Private Sub txtstrNumeroC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumeroC
End Sub

Private Sub txtstrComplementoC_GotFocus()
    MarcaCampo txtstrComplementoC
End Sub

Private Sub txtstrComplementoC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrComplementoC
End Sub

Private Sub txtintCepC_GotFocus()
    MarcaCampo txtintCEPC
End Sub

Private Sub txtintCepC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCEPC
End Sub

Private Sub txtintCepC_LostFocus()
    If txtintCEPC.Text = "" Then
        Exit Sub
    End If
    
    txtstrNumeroC.Text = ""
    txtstrComplementoC.Text = ""
    
    txtintCEPC = gstrCEPFormatado(txtintCEPC)
    CepLogradouro txtintCEPC, txtstrLogradouroC, txtstrBairroC, txtstrMunicipioC, txtstrUFC, , , , False, False, False, False, False, False
    
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case UCase(strModoOperacao)
        
        Case Is = UCase(gstrPreencherLista)
                PreencherListaDeOpcoes Me.ActiveControl
                
        Case Is = UCase(gstrIncluirItem)
            
            If blnDadosIssOk Then
                IncluiValoresNoGrid
            End If
            
        Case Is = UCase(gstrExcluirItem)
            
            If Val(tdb_Predios.Columns("PkidArray").Value & Space$(0)) > 0 Then
                ExcluiValoresDoGrid
            End If
            
        Case Is = UCase(gstrNovo)
                
            If tab_3dPasta.Tab = 0 Then
            
                Limpa_Controles Me, True, True, False, True, False
            
                Set vetPredios = New XArrayDB
                vetPredios.Clear
                vetPredios.ReDim 0, 0, 0, 19
                PkidArray = 0
                
                Set tdb_Predios.Array = vetPredios
                tdb_Predios.ReBind
                tdb_Predios.Refresh
                        
                TrocaCorObjeto dbc_strEmissao, False
                TrocaCorObjeto dbc_strInscricaoInicial, False
                TrocaCorObjeto dbc_strInscricaoFinal, False
                TrocaCorObjeto dbc_intExercicioInicial, False
                'TrocaCorObjeto chk_SelecionarTodos, False
                chk_SelecionarTodos.Enabled = True
                TrocaCorObjeto cmd_Inscricao(0), False
                TrocaCorObjeto cmd_Inscricao(1), False
        
                tab_3dPasta.TabEnabled(1) = False
                        
                dbc_intComposicao.SetFocus
            
            Else
                
                If Not dbcintPredios.MatchedWithList And Len(txtdblArea) = 0 And Len(txtdtmConstrucao) = 0 And Not dbcintIssConstrucaoTipo.MatchedWithList And Not dbcintAcabamento.MatchedWithList Then
                    dbcstrInscricao.BoundText = Space$(0)
                    dbc_intExercicioISS.BoundText = Space$(0)
                    Set dbcintContribuinte.RowSource = Nothing
                    dbcintContribuinte.Text = ""
                    LimpaDadosISSConstrucao True
                    tab_3dEnderecos.Tab = 0
                    dbcstrInscricao.SetFocus
                Else
                    LimpaDadosISSConstrucao False
                    
                    dbcintPredios.SetFocus
                End If
                
            End If
        
        Case Is = UCase(gstrCalcularReajuste)
            'If VerificaCalculoEmAndamento Then
                If blnDadosOK Then
                    If bytTipoComposicao = TYP_IMOBILIARIA Or bytTipoComposicao = TYP_OUTROS Then
                        RealizaCalculoImobiliario
                    ElseIf bytTipoComposicao = TYP_ECONOMICA Then
                        RealizaCalculoEconomico
                    Else
                        If blnEndNotificacaoOK Then
                            RealizaCalculoIssConstrucao
                        End If
                    End If
                End If
                LiberaCalculo
            'End If
        
    End Select

End Sub

Private Function blnEndNotificacaoOK() As Boolean
    blnEndNotificacaoOK = False
    
    If Trim(txtstrLogradouroC.Text) = "" Then
        ExibeMensagem "O logradouro de notificação deve ser informado."
        txtstrLogradouroC.SetFocus
        Exit Function
    ElseIf Trim(txtstrNumeroC.Text) = "" Then
        ExibeMensagem "O número de notificação deve ser informado."
        txtstrNumeroC.SetFocus
        Exit Function
    ElseIf Trim(txtintCEPC.Text) = "" Then
        ExibeMensagem "O cep de notificação deve ser informado."
        txtintCEPC.SetFocus
        Exit Function
    End If
    
    blnEndNotificacaoOK = True
End Function

Private Sub RealizaCalculoImobiliario()
    Dim adoResultado             As ADODB.Recordset
    Dim adoAux                   As ADODB.Recordset
    Dim adoParameters            As ADODB.Parameters
    Dim adoteste                As ADODB.Recordset
    Dim strMensagem              As String
    Dim strSQL                   As String
    Dim intFor                   As Integer
    Dim dblValorTerrenoExcedente As Double
    Dim lngPkidLancamentoAlfa    As Long
    Dim bytResultadoCalculo      As Byte
    Dim strCriticas              As String 'Hugo 17/01/2005
    
    Screen.MousePointer = vbHourglass
    
    prgStatus.Visible = True
    
    Set vetRelLanctoDevolver = New XArrayDB
    
    'Vamos retornar o intervalo de inscricoes para realizar o loop do calculo
    strSQL = "SELECT Pkid, strInscricao FROM " & gstrImobiliario & " " & strREADPAST
    If dbc_strEmissao.MatchedWithList Then
        strSQL = strSQL & " WHERE UPPER(strEmissao) = '" & String(gintLenEmissao - Len(Trim(dbc_strEmissao.Text)), "0") & UCase(dbc_strEmissao.Text) & "' And "
    Else
        If chk_SelecionarTodos.Value = vbUnchecked Then
            strSQL = strSQL & " WHERE strInscricao BETWEEN '" & String(gintLenInscricao - Len(Trim(dbc_strInscricaoInicial.Text)), "0") & UCase(dbc_strInscricaoInicial.Text) & "' AND '" & String(gintLenInscricao - Len(Trim(dbc_strInscricaoFinal.Text)), "0") & UCase(dbc_strInscricaoFinal.Text) & "' And "
        End If
    End If
    strSQL = strSQL & " Dtmdtcancelamento is Null "
    strSQL = strSQL & " ORDER BY strInscricao"

    Set gobjBanco = New clsBanco
    vetRelLanctoDevolver.ReDim 0, 0, 0, 2
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
        
            If MsgBox("Deseja excluir os registros de críticas anteriores ?", vbYesNo, "Mensagem ao Usuário") = vbYes Then
                gobjBanco.Execute "DELETE FROM " & gstrCriticaIptu & " WHERE lngCodUsr = " & glngCodUsr
            End If
            
            prgStatus.Max = adoResultado.RecordCount
            lbl_Final.Caption = adoResultado.RecordCount
            'Vamos calcular o valor inscricao por inscricao
            For intFor = 0 To adoResultado.RecordCount - 1
                lngPkidLancamentoAlfa = InscricaoJaCadastrada(adoResultado("strInscricao").Value, dbc_intExercicioInicial.Text, dbc_intComposicao.BoundText)
                
                gobjBanco.ExecutaBeginTrans
                
                'Vamos obter os parametros necessario para calcular a inscricao
                strSQL = "SELECT IM.Pkid, IM.Dblarea, " & _
                                "(SELECT CI.intCodigoDetalheDaCaracteristi FROM " & gstrCaracteristicaDoImovel & " CI WHERE CI.intCodigoCaracteristicaGeral = (SELECT DISTINCT DC.INTCARACTERISTICA FROM " & gstrDetalheDaCaracteristica & " DC WHERE intReferenciaTributo = " & FATOR_TOPOGRAFIA & ") AND  CI.intCodigoImobiliario = IM.PKID) FatorTopografia, " & _
                                "(SELECT CI.intCodigoDetalheDaCaracteristi FROM " & gstrCaracteristicaDoImovel & " CI WHERE CI.intCodigoCaracteristicaGeral = (SELECT DISTINCT DC.INTCARACTERISTICA FROM " & gstrDetalheDaCaracteristica & " DC WHERE intReferenciaTributo = " & FATOR_PEDOLOGIA & ") AND  CI.intCodigoImobiliario = IM.PKID) FatorPedologia, " & _
                                "(SELECT CI.intCodigoDetalheDaCaracteristi FROM " & gstrCaracteristicaDoImovel & " CI WHERE CI.intCodigoCaracteristicaGeral = (SELECT DISTINCT DC.INTCARACTERISTICA FROM " & gstrDetalheDaCaracteristica & " DC WHERE intReferenciaTributo = " & FATOR_SITUACAO & ") AND  CI.intCodigoImobiliario = IM.PKID) FatorSituacao, " & _
                                "(SELECT CI.intCodigoDetalheDaCaracteristi FROM " & gstrCaracteristicaDoImovel & " CI WHERE CI.intCodigoCaracteristicaGeral = (SELECT DISTINCT DC.INTCARACTERISTICA FROM " & gstrDetalheDaCaracteristica & " DC WHERE intReferenciaTributo = " & FATOR_ZONEAMENTO & ") AND  CI.intCodigoImobiliario = IM.PKID) FatorZoneamento, " & _
                                "(SELECT CI.intCodigoDetalheDaCaracteristi FROM " & gstrCaracteristicaDoImovel & " CI WHERE CI.intCodigoCaracteristicaGeral = (SELECT DISTINCT DC.INTCARACTERISTICA FROM " & gstrDetalheDaCaracteristica & " DC WHERE intReferenciaTributo = " & FATOR_DESVIO_FERROVIARIO & ") AND  CI.intCodigoImobiliario = IM.PKID) FatorDesvio, " & _
                                "(SELECT CI.intCodigoDetalheDaCaracteristi FROM " & gstrCaracteristicaDoImovel & " CI WHERE CI.intCodigoCaracteristicaGeral = (SELECT DISTINCT DC.INTCARACTERISTICA FROM " & gstrDetalheDaCaracteristica & " DC WHERE intReferenciaTributo = " & FATOR_CORREGO & ") AND  CI.intCodigoImobiliario = IM.PKID) FatorCorrego, "
                If bytDBType = Oracle Then
                    strSQL = strSQL & "(SELECT SUM(" & gstrCONVERT(cdt_numeric, "REPLACE(TI.Strmedidadatestada, '.', ',')") & ") FROM " & gstrTipoDeTestada & " TT, " & gstrTestadaImobiliario & " TI WHERE TI.INTIMOBILIARIO = IM.PKID AND TI.intTipoDeTestada = TT.Pkid AND TT.bytPrincipal = 0) Testadas, " & _
                                "SUM(" & gstrCONVERT(cdt_numeric, "REPLACE(TI2.Strmedidadatestada, '.', ',')") & ")TestadaPrincipal, "
                Else
                    strSQL = strSQL & "(SELECT SUM(CONVERT(Decimal(18,2), TI.Strmedidadatestada) ) FROM " & gstrTipoDeTestada & " TT, " & gstrTestadaImobiliario & " TI WHERE TI.INTIMOBILIARIO = IM.PKID AND TI.intTipoDeTestada = TT.Pkid AND TT.bytPrincipal = 0) Testadas, " & _
                                "SUM(Convert(Decimal(18,2),TI2.Strmedidadatestada) )TestadaPrincipal, "
                End If
                strSQL = strSQL & "VMT.DBLVALOR ValorFaceDeQuadra " & _
                         "FROM " & gstrImobiliario & " IM " & strREADPAST & " , " & gstrTipoDeTestada & " TT2, " & gstrTestadaImobiliario & " TI2, " & gstrHistoricoFaceDeQuadra & " HFQ, " & gstrValorMetroTerreno & " VMT " & _
                         "WHERE TI2.INTIMOBILIARIO = IM.PKID AND TI2.intTipoDeTestada = TT2.Pkid AND TT2.bytPrincipal = 1 AND " & _
                                "TI2.INTFACEDEQUADRA = HFQ.INTFACEDEQUADRA AND HFQ.INTVALORMETROTERRENO = VMT.PKID AND " & _
                                "HFQ.intExercicio = " & dbc_intExercicioInicial.Text & " AND " & _
                                "VMT.intExercicio = HFQ.intExercicio AND " & _
                                "IM.Pkid = " & adoResultado("Pkid").Value & " " & _
                         "GROUP BY IM.STRINSCRICAO, IM.Dblarea ,IM.PKID, VMT.DBLVALOR "
                         
                If gobjBanco.CriaADO(strSQL, 20, adoAux) Then
                    With adoAux
                    
                    If Not adoAux.EOF Then
                    
                        'Vamos verificar se existe algum fator nao definido e criticar o mesmo
                        '***AGORA ESTA SENDO VERIFICADO NA PROCEDURE, POIS NEM TODAS PREFEITURAS UTILIZAM
                        'If Not IsNull(!FatorTopografia) And Not IsNull(!FatorPedologia) And Not IsNull(!FatorSituacao) And Not IsNull(!FatorZoneamento) And Not IsNull(!FatorDesvio) And Not IsNull(!FatorCorrego) Then
                            If bytTipoComposicao = TYP_OUTROS Then
                                strSQL = gstrStoredProcedure("sp_CalculoImobiliarioOutros", gstrENulo(!Pkid) & ", " & gstrENulo(dbc_intComposicao.BoundText) & ", " & gstrENulo(dbc_intExercicioInicial.BoundText) & ", " & gstrConvVrParaSql(!dblArea) & ", " & gstrConvVrParaSql(!Testadas) & ", " & gstrConvVrParaSql(!TestadaPrincipal) & ", " & gstrENulo(!FatorTopografia, , True) & ", " & gstrENulo(!FatorPedologia, , True) & ", " & gstrENulo(!FatorSituacao, , True) & ", " & gstrENulo(!FatorZoneamento, , True) & ", " & gstrENulo(!FatorDesvio, , True) & ", " & gstrENulo(!FatorCorrego, , True) & ", " & gstrConvVrParaSql(!valorfacedequadra) & ", " & chk_Simulado.Value & ", " & glngCodUsr & ", " & lngPkidLancamentoAlfa)
                            Else
                                strSQL = gstrStoredProcedure("sp_CalculoImobiliario", gstrENulo(!Pkid) & ", " & gstrENulo(dbc_intComposicao.BoundText) & ", " & gstrENulo(dbc_intExercicioInicial.BoundText) & ", " & gstrConvVrParaSql(!dblArea) & ", " & gstrConvVrParaSql(!Testadas) & ", " & gstrConvVrParaSql(!TestadaPrincipal) & ", " & gstrENulo(!FatorTopografia, , True) & ", " & gstrENulo(!FatorPedologia, , True) & ", " & gstrENulo(!FatorSituacao, , True) & ", " & gstrENulo(!FatorZoneamento, , True) & ", " & gstrENulo(!FatorDesvio, , True) & ", " & gstrENulo(!FatorCorrego, , True) & ", " & gstrConvVrParaSql(!valorfacedequadra) & ", " & chk_Simulado.Value & ", " & glngCodUsr & ", " & lngPkidLancamentoAlfa)
                            End If

TentarNovamente:
                            
                            If gobjBanco.ExecuteStoredProcedure(strSQL, 40, , adoParameters, False) Then
                            'If gobjBanco.CriaADO(strSql, 10, adoParameters) Then
                                If Not (adoParameters Is Nothing) Then
                                    If chk_Critica.Value Then
                                        MsgBox "Estes são os valores que a  procedure retorna: " & Chr(13) & _
                                        " Area terreno: " & gstrConvVrDoSql(adoParameters("V_dblRetAreaTerreno").Value, 2) & Chr(13) & _
                                        " Valor venal terreno: " & gstrConvVrDoSql(adoParameters("V_dblRetValorVenalTerreno").Value, 2) & Chr(13) & _
                                        " Area terreno excedente: " & gstrConvVrDoSql(adoParameters("V_dblRetAreaTerrenoExcedente").Value, 2) & Chr(13) & _
                                        " Valor terreno excedente: " & gstrConvVrDoSql(adoParameters("V_dblRetValorTerrenoExcedente").Value, 2) & Chr(13) & _
                                        " Area total predio: " & gstrConvVrDoSql(adoParameters("V_dblRetAreaTotalPredio").Value, 2) & Chr(13) & _
                                        " Valor total predio: " & gstrConvVrDoSql(adoParameters("V_dblRetValorTotalPredio").Value, 2)
                                    End If
                                    
                                    'Caso nao tenha retornado o Pkid da LancamentoAlfa, nao foi concluida a procedure, entao nao passaremos pela gravacao em LancamentoValor e Lancamento Receita
                                    If IsNull(adoParameters("V_lngRetPkidLancamentoAlfa").Value) Then
                                        gobjBanco.ExecutaRollbackTrans
                                        CriaCriticaDeIptu Trim(dbc_intComposicao.Text), adoResultado("strInscricao").Value, Trim(dbc_intExercicioInicial.Text), "A rotina de calculo não foi concluída.", "Não foi criado registro em Lançamento Alfa na procedure (sp_CalculoImobiliario)", Trim(dbc_strEmissao.Text)
                                        GoTo Proxima_Inscricao
                                    End If
                                    
                                    'Caso seja uma simulacao vamos finalizar este calculo e ir para a proxima inscricao
                                    If chk_Simulado.Value Then
                                        gobjBanco.ExecutaRollbackTrans
                                        GoTo Proxima_Inscricao
                                    End If
                                    
                                    lngPkidLancamentoAlfa = adoParameters("V_lngRetPkidLancamentoAlfa").Value
                                    
                                    'Vamos alimentar array do relatorio de contribuintes com valor a devolver
                                    If Len(vetRelLanctoDevolver(0, 0)) <> 0 Then
                                        vetRelLanctoDevolver.ReDim 0, vetRelLanctoDevolver.UpperBound(1) + 1, 0, 2
                                    End If
                                    vetRelLanctoDevolver(vetRelLanctoDevolver.UpperBound(1), 0) = adoParameters("V_lngRetPkidLancamentoAlfa").Value
                                    vetRelLanctoDevolver(vetRelLanctoDevolver.UpperBound(1), 2) = adoParameters("V_dblRetValorTotalCancelamento").Value
                                    
                                End If
                                
                            Else
                                'Vamos verificar se é problema de Unique Key de Numero Aviso
                                If InStr(1, UCase(gstrErrorInStoredProcedure), "UK_TBLLANCAMENTOALFA_NUMAVISO") > 0 Then
                                    GoTo TentarNovamente
                                End If
                                
                                gobjBanco.ExecutaRollbackTrans
                                
                                CriaCriticaDeIptu Trim(dbc_intComposicao.Text), adoResultado("strInscricao").Value, Trim(dbc_intExercicioInicial.Text), gstrErrorInStoredProcedure, "Erro ao executar Stored Procedure (sp_CalculoImobiliario)", Trim(dbc_strEmissao.Text)
                                
                                gstrErrorInStoredProcedure = ""

                                GoTo Proxima_Inscricao
                                
                            End If
                            
                        'Else
                        '
                        '    strMensagem = "O(s) fator(es) de"
                        '    If IsNull(!FatorTopografia) Then strMensagem = strMensagem & " Topografia, "
                        '    If IsNull(!FatorPedologia) Then strMensagem = strMensagem & " Pedologia, "
                        '    If IsNull(!FatorSituacao) Then strMensagem = strMensagem & " Situação, "
                        '    If IsNull(!FatorZoneamento) Then strMensagem = strMensagem & " Zoneamento, "
                        '    If IsNull(!FatorDesvio) Then strMensagem = strMensagem & " Desvio Ferroviario, "
                        '    If IsNull(!FatorCorrego) Then strMensagem = strMensagem & " Córrego, "
                        '    strMensagem = strMensagem & "não foi(ram) definido(s) para a inscrição " & gstrENulo(adoResultado("strInscricao").Value) & "."
                        '
                        '    'ExibeMensagem strMensagem
                        '
                        '    gobjBanco.ExecutaRollbackTrans
                        '
                        '    CriaCriticaDeIptu Trim(dbc_intComposicao.Text), adoResultado("strInscricao").Value, Trim(dbc_intExercicioInicial.Text), strMensagem, , Trim(dbc_strEmissao.Text)
                        '
                        '    GoTo Proxima_Inscricao
                        '
                        'End If
                    
                    Else
                        
                        gobjBanco.ExecutaRollbackTrans
                        
                        CriaCriticaDeIptu Trim(dbc_intComposicao.Text), adoResultado("strInscricao").Value, Trim(dbc_intExercicioInicial.Text), "Não foi(ram) encontrado(s) algum(ns) dado(s) de referência para a inscrição", "Não foi definida uma Testada Principal, ou o Exercício da mesma não correspondente ao selecionado", Trim(dbc_strEmissao.Text)
                        
                        GoTo Proxima_Inscricao
    
                    End If
                    
                    End With
                
                Else
                    
                    gobjBanco.ExecutaRollbackTrans
                    
                    CriaCriticaDeIptu Trim(dbc_intComposicao.Text), adoResultado("strInscricao").Value, Trim(dbc_intExercicioInicial.Text), "Não foi possível obter os parâmetros necessários para calculo.", "Não foi encontrada referencia em uma das seguintes tabelas: " & gstrTipoDeTestada & ", " & gstrTestadaImobiliario & ", " & gstrHistoricoFaceDeQuadra & ", " & gstrValorMetroTerreno, Trim(dbc_strEmissao.Text)
                                    
                    GoTo Proxima_Inscricao
                    
                End If
                
                bytResultadoCalculo = CalculaImpostos(adoResultado("Pkid").Value, lngPkidLancamentoAlfa, adoResultado("strInscricao").Value)
                
                If bytResultadoCalculo = BYT_CALCULOIMPOSTO_ERRO_RECEITA Then
                    gobjBanco.ExecutaRollbackTrans
                    CriaCriticaDeIptu Trim(dbc_intComposicao.Text), adoResultado("strInscricao").Value, Trim(dbc_intExercicioInicial.Text), gstrErrorInStoredProcedure, "Erro ao executar a função (CalculaImpostos)", Trim(dbc_strEmissao.Text)
                    gstrErrorInStoredProcedure = ""
                    GoTo Proxima_Inscricao
                ElseIf bytResultadoCalculo = BYT_CALCULOIMPOSTO_ERRO_GERAL Then
                    gobjBanco.ExecutaRollbackTrans
                    CriaCriticaDeIptu Trim(dbc_intComposicao.Text), adoResultado("strInscricao").Value, Trim(dbc_intExercicioInicial.Text), gstrErrorInStoredProcedure, "Erro ao executar a função (CalculaImpostos)", Trim(dbc_strEmissao.Text)
                    gstrErrorInStoredProcedure = ""
                    Exit Sub
                End If
                
                gobjBanco.ExecutaCommitTrans
                
Proxima_Inscricao:
DoEvents
                prgStatus.Value = adoResultado.AbsolutePosition
                lbl_Inicial.Caption = adoResultado.AbsolutePosition
                adoResultado.MoveNext
                
            Next
            
            'Vamos imprimir o relatorio de lancamentos
            If vetRelLanctoDevolver.Count(1) > 0 Then
                If Not IsNull(vetRelLanctoDevolver(0, 0)) Then
                    'ImprimeRelatorioPorArray rptLancamentosCompCancel, , "Lançamentos com Compensação / Cancelamento", , vetRelLanctoDevolver, True
                End If
            End If
        Else
            ExibeMensagem "Não foi(ram) encontrado(s) dado(s) de referência para informções passadas."
        End If
    End If
    
    Screen.MousePointer = vbDefault
    prgStatus.Visible = False
    
End Sub

Private Sub RealizaCalculoEconomico()
    Dim adoResultado                As ADODB.Recordset
    Dim adoAux                      As ADODB.Recordset
    Dim strReceitasSujeitas         As String
    Dim strSQL                      As String
    Dim intFor                      As Integer
    Dim lngPkidLancamentoAlfa       As Long
    Dim bytResultadoCalculo         As Byte
    
    Dim strCriticasInscrSemReceita  As String
    
    Screen.MousePointer = vbHourglass
    
    prgStatus.Visible = True
    
    strCriticasInscrSemReceita = Space$(0)
    
    'Vamos retornar o intervalo de inscricoes para realizar o loop do calculo
    strSQL = "SELECT E.Pkid, E.strInscricaoCadastral FROM " & gstrEconomico & " E " & strREADPAST & ", " & gstrAtividadeDaEmpresa & " AE, " & gstrAtivEmpresaTributo & " AET "
        
    If dbc_strEmissao.MatchedWithList Then
        strSQL = strSQL & " WHERE strEmissao = '" & String(gintLenEmissao - Len(Trim(dbc_strEmissao.Text)), "0") & UCase(dbc_strEmissao.Text) & "'"
    Else
        strSQL = strSQL & " WHERE strInscricaoCadastral BETWEEN '" & String(gintLenInscricao - Len(Trim(dbc_strInscricaoInicial.Text)), "0") & UCase(dbc_strInscricaoInicial.Text) & "' AND '" & String(gintLenInscricao - Len(Trim(dbc_strInscricaoFinal.Text)), "0") & UCase(dbc_strInscricaoFinal.Text) & "'"
    End If
    
    strSQL = strSQL & " AND E.Pkid = AE.Inteconomico AND AE.Pkid = AET.Intatividadedaempresa"
    
    strSQL = strSQL & " AND Dtmdataencerramento Is Null Group By E.Pkid, E.strInscricaoCadastral "
    
    strSQL = strSQL & " ORDER BY strInscricaoCadastral"

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            prgStatus.Max = adoResultado.RecordCount
            lbl_Final.Caption = adoResultado.RecordCount
            'Vamos calcular o valor inscricao por inscricao
            For intFor = 0 To adoResultado.RecordCount - 1
                
                lngPkidLancamentoAlfa = InscricaoJaCadastrada(adoResultado("strInscricaoCadastral").Value, dbc_intExercicioInicial.Text, dbc_intComposicao.BoundText)
                
                gobjBanco.ExecutaBeginTrans
                
                'Vamos obter as receitas a que o economico esta sujeito
'                strSql = "SELECT TT.intReceita, ET.intTributo " & _
'                         "FROM " & gstrEconomico & " EC, " & gstrAtividadeDaEmpresa & " AE, tblAtivEmpresaTributo ET, " & gstrTributo & " TR, " & _
'                                 gstrTributoTipo & " TT " & _
'                         "WHERE EC.Pkid = " & adoResultado("Pkid").Value & " AND " & _
'                               "AE.intEconomico = EC.Pkid AND AE.blnPrincipal = 1 AND " & _
'                               "ET.intAtividadeDaEmpresa = AE.Pkid AND TR.Pkid = ET.intTributo AND " & _
'                               "TT.Pkid = TR.intTributoTipo " & _
'                         "GROUP BY TT.intReceita, ET.intTributo " & _
'                         " UNION ALL " & _
'                         "SELECT TT.intReceita, HP.intTributo " & _
'                         "FROM " & gstrEconomico & " EC, " & gstrHistoricoPublicidades & " HP, " & gstrTributo & " TR, " & gstrTributoTipo & " TT " & _
'                         "WHERE EC.Pkid = " & adoResultado("Pkid").Value & " AND " & _
'                               "HP.intEconomico = EC.Pkid AND " & _
'                               "TR.Pkid = HP.intTributo AND " & _
'                               "TT.Pkid = TR.intTributoTipo " & _
'                               "GROUP BY TT.intReceita, HP.intTributo"
                                               
                If bytDBType = Oracle Then
                    strSQL = "Select "
                    strSQL = strSQL & "intReceita, "
                    strSQL = strSQL & "intTributo "
                    strSQL = strSQL & "From "
                    strSQL = strSQL & "(SELECT "
                    strSQL = strSQL & "TT.intReceita, "
                    strSQL = strSQL & "ET.intTributo, "
                    strSQL = strSQL & "TE.Dblvalor "
                    strSQL = strSQL & "FROM "
                    strSQL = strSQL & "tblEconomico EC " & strREADPAST & " , "
                    strSQL = strSQL & "tblAtividadeDaEmpresa AE, "
                    strSQL = strSQL & "tblAtivEmpresaTributo ET, "
                    strSQL = strSQL & "tblTributo TR, "
                    strSQL = strSQL & "tblTributoTipo TT, "
                    strSQL = strSQL & "tbltributoexercicio TE "
                    strSQL = strSQL & "WHERE "
                    strSQL = strSQL & "EC.Pkid = " & adoResultado("Pkid").Value & " AND "
                    strSQL = strSQL & "AE.intEconomico = EC.Pkid AND "
                    strSQL = strSQL & "ET.intAtividadeDaEmpresa = AE.Pkid AND "
                    strSQL = strSQL & "TR.Pkid = ET.intTributo AND "
                    strSQL = strSQL & "TT.Pkid = TR.intTributoTipo and "
                    strSQL = strSQL & "TR.Pkid = TE.Inttributo AND "
                    strSQL = strSQL & "TE.Intexercicio = " & dbc_intExercicioInicial.Text
                    strSQL = strSQL & " GROUP BY "
                    strSQL = strSQL & "TT.intReceita, ET.intTributo ,TE.Dblvalor "
                    strSQL = strSQL & "Order by "
                    strSQL = strSQL & "TE.Dblvalor Desc "
                    strSQL = strSQL & ") A "
                    strSQL = gstrTOPnOracle(strSQL, 1)
                Else
                    strSQL = "SELECT TT.intReceita, ET.intTributo " & _
                        "FROM " & gstrEconomico & " EC " & strREADPAST & " , " & gstrAtividadeDaEmpresa & " AE, tblAtivEmpresaTributo ET, " & gstrTributo & " TR, " & _
                                gstrTributoTipo & " TT " & _
                        "WHERE EC.Pkid = " & adoResultado("Pkid").Value & " AND " & _
                              "AE.intEconomico = EC.Pkid AND AE.blnPrincipal = 1 AND " & _
                              "ET.intAtividadeDaEmpresa = AE.Pkid AND TR.Pkid = ET.intTributo AND " & _
                              "TT.Pkid = TR.intTributoTipo " & _
                        "GROUP BY TT.intReceita, ET.intTributo "
                End If
                
                 strSQL = strSQL & " UNION ALL " & _
                 "SELECT TT.intReceita, HP.intTributo " & _
                 "FROM " & gstrEconomico & " EC " & strREADPAST & " , " & gstrHistoricoPublicidades & " HP, " & gstrTributo & " TR, " & gstrTributoTipo & " TT " & _
                 "WHERE EC.Pkid = " & adoResultado("Pkid").Value & " AND " & _
                       "HP.intEconomico = EC.Pkid AND " & _
                       "TR.Pkid = HP.intTributo AND " & _
                       "TT.Pkid = TR.intTributoTipo " & _
                       "GROUP BY TT.intReceita, HP.intTributo"
                               
                strReceitasSujeitas = ""
                
                If gobjBanco.CriaADO(strSQL, 5, adoAux) Then
                    
                    If adoAux.EOF Then
                        strCriticasInscrSemReceita = strCriticasInscrSemReceita & IIf(Len(strCriticasInscrSemReceita) > 0, "," & Val(adoResultado("strInscricaoCadastral").Value), Val(adoResultado("strInscricaoCadastral").Value))
                        gobjBanco.ExecutaRollbackTrans
                        GoTo Proxima_Inscricao
                    End If
                    
                    Do While Not adoAux.EOF
                        strReceitasSujeitas = strReceitasSujeitas & "|" & adoAux("intReceita").Value & ";" & adoAux("intTributo").Value
                        adoAux.MoveNext
                    Loop
                    
                    strReceitasSujeitas = strReceitasSujeitas & "|"
                    
                Else
                    gobjBanco.ExecutaRollbackTrans
                    GoTo Proxima_Inscricao
                End If
                
                bytResultadoCalculo = CalculaImpostos(adoResultado("Pkid").Value, lngPkidLancamentoAlfa, adoResultado("strInscricaoCadastral").Value, strReceitasSujeitas, False)
                
                If bytResultadoCalculo = BYT_CALCULOIMPOSTO_ERRO_RECEITA Then
                    gobjBanco.ExecutaRollbackTrans
                    GoTo Proxima_Inscricao
                ElseIf bytResultadoCalculo = BYT_CALCULOIMPOSTO_ERRO_GERAL Then
                    gobjBanco.ExecutaRollbackTrans
                    ExibeMensagem "O calculo foi realizado com ocorrência de críticas."
                    Exit Sub
                End If
                
                If chk_Simulado Then
                    gobjBanco.ExecutaRollbackTrans
                Else
                    gobjBanco.ExecutaCommitTrans
                End If
                
Proxima_Inscricao:
                DoEvents
                prgStatus.Value = adoResultado.AbsolutePosition
                lbl_Inicial.Caption = adoResultado.AbsolutePosition
                adoResultado.MoveNext
                
            Next
            
            If Len(strCriticasInscrSemReceita) > 0 Then
                ExibeMensagem "Não foi possível encontrar Receita(s) no Tributo da(s) Inscrição(ões) " & strCriticasInscrSemReceita
            End If
            
            ExibeMensagem "O calculo foi realizado com sucesso."
            
        Else
            ExibeMensagem "Não foi possível encontrar registros com os parâmetros passados."
        End If
    End If
    
    Screen.MousePointer = vbDefault
    prgStatus.Visible = False
    
End Sub

Private Sub RealizaCalculoIssConstrucao()
Dim adoResultado             As ADODB.Recordset

Dim strSQL                   As String

Dim lngPkidLancamentoAlfa    As Long

Dim bytResultadoCalculo      As Byte
    
    Screen.MousePointer = vbHourglass
    
    'Vamos retornar o intervalo de inscricoes para realizar o loop do calculo
    strSQL = "SELECT Pkid, strInscricao FROM " & gstrImobiliario & " " & strREADPAST
    strSQL = strSQL & " WHERE strInscricao = '" & String(gintLenInscricao - Len(Trim(dbcstrInscricao.Text)), "0") & UCase(dbcstrInscricao.Text) & "'"
    strSQL = strSQL & " ORDER BY strInscricao"

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            
        lngPkidLancamentoAlfa = InscricaoJaCadastrada(adoResultado("strInscricao").Value, dbc_intExercicioISS.Text, dbc_intComposicao.BoundText)
        
        gobjBanco.ExecutaBeginTrans
        
        bytResultadoCalculo = CalculaImpostos(adoResultado("Pkid").Value, lngPkidLancamentoAlfa, adoResultado("strInscricao").Value)
        
        If bytResultadoCalculo = BYT_CALCULOIMPOSTO_ERRO_RECEITA Then
            gobjBanco.ExecutaRollbackTrans
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf bytResultadoCalculo = BYT_CALCULOIMPOSTO_ERRO_GERAL Then
            gobjBanco.ExecutaRollbackTrans
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        gobjBanco.ExecutaCommitTrans
        
    End If
        
    ExibeMensagem "Calculo realizado com sucesso."
    
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub PreencheEmissao()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset

    strSQL = "SELECT DISTINCT "
    strSQL = strSQL & gstrCONVERT(cdt_numeric, "strEmissao")
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParametroIPTU & " " & strREADPAST
    strSQL = strSQL & " WHERE intComposicaoDaReceita = " & dbc_intComposicao.BoundText
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(cdt_numeric, "strEmissao")


    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dbc_strEmissao.ListField = adoResultado.Fields(0).Name
            Set dbc_strEmissao.RowSource = adoResultado
            dbc_strEmissaoIss.ListField = adoResultado.Fields(0).Name
            Set dbc_strEmissaoIss.RowSource = adoResultado
        End If
    End If

End Sub

Private Sub PreencheExercicio(dbcExercicio As DataCombo)
Dim strSQL As String
Dim adoResultado As ADODB.Recordset

    strSQL = "SELECT DISTINCT intExercicio"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrParametroIPTU & " " & strREADPAST
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " intComposicaoDaReceita = " & dbc_intComposicao.BoundText
    strSQL = strSQL & IIf(dbc_strEmissao.MatchedWithList, " AND UPPER(strEmissao) = " & String(gintLenEmissao - Len(Trim(dbc_strEmissao.Text)), "0") & UCase(dbc_strEmissao.Text), "")
    strSQL = strSQL & " ORDER BY intExercicio"

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            dbcExercicio.ListField = adoResultado.Fields(0).Name
            Set dbcExercicio.RowSource = adoResultado
            'dbc_intExercicioFinal.ListField = adoResultado.Fields(0).Name
            'Set dbc_intExercicioFinal.RowSource = adoResultado
        End If
    End If

End Sub

Private Sub DefineComposicao(PkidComposicao As Long)
Dim strSQL       As String
Dim adoResultado As ADODB.Recordset

    strSQL = "SELECT intUtilizacao"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " Pkid = " & PkidComposicao

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            bytTipoComposicao = adoResultado("intUtilizacao").Value
        End If
    End If

End Sub

Private Sub LimpaDataCombo(dbcAux As DataCombo)

    dbcAux.Tag = ""
    dbcAux.Text = ""
    dbcAux.ListField = ""
    
End Sub

Private Function blnDadosOK() As Boolean

    blnDadosOK = False

    If Not dbc_intComposicao.MatchedWithList Then
        ExibeMensagem "Selecione uma Composição da Receita válida."
        dbc_intComposicao.SetFocus
        Exit Function
    End If
    
    If bytTipoComposicao = TYP_ISS_CONSTRUCAO Then
                    
        If vetPredios.UpperBound(1) = -1 Then
            ExibeMensagem "É preciso inserir algum prédio."
            dbcintPredios.SetFocus
            Exit Function
        ElseIf Val(vetPredios(0, PREDIO_MEDIDAAREA)) = 0 Then
            ExibeMensagem "É preciso inserir algum prédio."
            dbcintPredios.SetFocus
            Exit Function
        End If
        
        If Not dbcstrInscricao.MatchedWithList Then
            ExibeMensagem "Selecione uma Inscrição Inicial válida."
            dbcstrInscricao.SetFocus
            Exit Function
        End If
    
        If Not dbc_strEmissaoIss.MatchedWithList Then
            ExibeMensagem "Selecione uma Emissão válida."
            dbc_strEmissaoIss.SetFocus
            Exit Function
        End If
    
        If Not dbc_intExercicioISS.MatchedWithList Then
            ExibeMensagem "Selecione um Exercício válido."
            dbc_intExercicioISS.SetFocus
            Exit Function
        End If
        
        If Not VerificaProcesso Then Exit Function
        
    Else
        
        If blnModoConsultaEmissao Then
            If Not dbc_strEmissao.MatchedWithList Then
                ExibeMensagem "Selecione uma Emissão válida."
                dbc_strEmissao.SetFocus
                Exit Function
            End If
        Else
            If chk_SelecionarTodos.Value = 0 Then
                If Not dbc_strInscricaoInicial.MatchedWithList Then
                    ExibeMensagem "Selecione uma Inscrição Inicial válida."
                    dbc_strInscricaoInicial.SetFocus
                    Exit Function
                ElseIf Not dbc_strInscricaoFinal.MatchedWithList Then
                    ExibeMensagem "Selecione uma Inscrição Final válida."
                    dbc_strInscricaoFinal.SetFocus
                    Exit Function
                End If
            End If
        End If
        
        If dbc_strInscricaoInicial.Text > dbc_strInscricaoFinal.Text Then
            ExibeMensagem "A Inscrição inicial não pode ser superior à final."
            dbc_strInscricaoInicial.SetFocus
            Exit Function
        End If
        
        If Not dbc_intExercicioInicial.MatchedWithList Then
            ExibeMensagem "Selecione um Exercício válido."
            dbc_intExercicioInicial.SetFocus
            Exit Function
        End If
    
    End If
    
    blnDadosOK = True

End Function

Private Function blnDadosIssOk() As Boolean

    blnDadosIssOk = False
  
    If Not dbcstrInscricao.MatchedWithList Then
        ExibeMensagem "Selecione uma inscrição válida."
        dbcstrInscricao.SetFocus
        Exit Function
    End If
  
    If Val(txtdblArea.Text) = 0 Then
        ExibeMensagem "A área do prédio deve ser informada."
        txtdblArea.SetFocus
        Exit Function
    End If

    If Len(txtdtmConstrucao.Text) = 0 Then
        ExibeMensagem "A data de construção do prédio deve ser informada."
        txtdtmConstrucao.SetFocus
        Exit Function
    End If

    If Not dbcintIssConstrucaoTipo.MatchedWithList Then
        ExibeMensagem "Selecione um tipo de construção válida."
        dbcintIssConstrucaoTipo.SetFocus
        Exit Function
    End If
    
    If Not dbcintAcabamento.MatchedWithList Then
        ExibeMensagem "Selecione um tipo de acabamento válido."
        dbcintAcabamento.SetFocus
        Exit Function
    End If
    
    If Val(txtdblVlrM2Servico.Text) = 0 Then
        ExibeMensagem "O valor do m2 de serviço deve ser informado."
        Exit Function
    End If
    
    If Val(txtdblVlrServico.Text) = 0 Then
        ExibeMensagem "O valor do serviço deve ser informado."
        Exit Function
    End If
    
    If Val(txtdblAliquota.Text) = 0 Then
        ExibeMensagem "O valor da alíquota deve ser informado."
        Exit Function
    End If
    
    If Val(txtdblIssDevido.Text) = 0 Then
        ExibeMensagem "O valor do Iss devido deve ser informado."
        Exit Function
    End If
    
    If Val(txtdblIssPagar.Text) = 0 Then
        ExibeMensagem "O valor do Iss a pagar deve ser informado."
        Exit Function
    End If
    
    blnDadosIssOk = True

End Function

Private Function strQueryComposicao() As String
Dim strSQL As String

    strSQL = "SELECT Pkid,"
    strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao Descricao"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " intUtilizacao in (1,2,6,7) "
    strSQL = strSQL & " ORDER BY intCodigo"

    strQueryComposicao = strSQL

End Function

Private Function strQueryInscricao() As String
Dim strSQL As String
    
    If bytTipoComposicao = TYP_ECONOMICA Then
        strSQL = "SELECT Pkid, " & gstrRIGHT("strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strInscricao "
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrEconomico
        strSQL = strSQL & " ORDER BY strInscricaoCadastral"
    Else
        strSQL = "SELECT Pkid, " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(bytTipoComposicao)) & " strInscricao "
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrImobiliario
        strSQL = strSQL & " ORDER BY strInscricao"
    End If
    
    strQueryInscricao = strSQL

End Function

Private Function RetornaPkidParametroIPTU(strTabela As String, lngPKId As Long, Optional strEmissao As String) As Long
Dim adoAux As ADODB.Recordset
Dim strSQL As String

    'Vamos obter o Pkid na tabela de Parametros de IPTU, os demais serao calculados na Procedure
    strSQL = "SELECT Pkid " & _
             "FROM " & gstrParametroIPTU & " " & strREADPAST & _
             "WHERE intComposicaoDaReceita = " & gstrENulo(dbc_intComposicao.BoundText) & " AND intExercicio = " & IIf(bytTipoComposicao = TYP_ISS_CONSTRUCAO, dbc_intExercicioISS.Text, dbc_intExercicioInicial.Text)
                 
    If Len(strEmissao) > 0 Then
        strSQL = strSQL & " AND strEmissao = '" & String(gintLenEmissao - Len(Trim(strEmissao)), "0") & UCase(strEmissao) & "'"
    Else
        strSQL = strSQL & " AND strEmissao = (SELECT strEmissao FROM " & strTabela & " WHERE Pkid = " & lngPKId & ") "
    End If
    
    strSQL = strSQL & "ORDER BY strEmissao DESC"
                    
    If gobjBanco.CriaADO(strSQL, 5, adoAux) Then
        If Not adoAux.EOF Then
            RetornaPkidParametroIPTU = adoAux("Pkid").Value
        Else
            RetornaPkidParametroIPTU = 0
        End If
    Else
        RetornaPkidParametroIPTU = 0
    End If
    
End Function

Private Function CalculaImpostos(lngPkidPrincipal As Long, lngPkidLancamentoAlfa As Long, strIncricaoCadastral As String, Optional strReceitasSujeitas As String, Optional blnExibeMensagens As Boolean = True) As Byte
Dim adoReceita                      As ADODB.Recordset
Dim adoPlanosPagto                  As ADODB.Recordset
Dim adoFormula                      As ADODB.Recordset
Dim adoParameters                   As ADODB.Parameters
Dim adoParcelas                     As ADODB.Recordset
Dim adoLancamentoAlfa               As ADODB.Recordset
Dim adoLancamentoValor              As ADODB.Recordset

Dim aImpostos()                     As String
Dim aImpostosAux()                  As String

Dim strSQL                          As String

Dim intForReceita                   As Integer
Dim intForFormula                   As Integer
Dim intForPlanosPagto               As Integer
Dim intForParcelas                  As Integer
Dim intQtdeParcelas                 As Integer

Dim intPrimeiraParcelaDoPlano       As Long    'Utilizada para armazenar a 1ª parcela no caso de haver valor nao parcelado

Dim intMoedaAtual                   As Integer

Dim intNumeroSequencial             As Integer 'Utilizado para a procedure identificar se é a 1ª vez de gravacao da inscricao
Dim intFor                          As Integer

Dim dblValorImposto                 As Double
Dim dblValorImpostoCompensacao      As Double
Dim dblValorImpostoDesconto         As Double
Dim dblValorImpostoNaoParcelado     As Double
Dim dblValorImpostoNaoParceladoDesc As Double
Dim dblValorPorParcela              As Double
Dim dblValorDiferencaParcela        As Double
Dim dblPorcentagemDaReceita         As Double
    
Dim blnIssVariavel                  As Boolean 'Variaveis utilizadas
Dim dtmAbertura                     As Date    'no caso de iss variavel

    intNumeroSequencial = 0
    
    'Vamos obter a moeda atual para gerar as parcelas
    strSQL = "SELECT EM.intMoeda FROM " & gstrEmpresa & " EM "

    If gobjBanco.CriaADO(strSQL, 5, adoParcelas) Then
        If Not adoParcelas.EOF Then
            If Not IsNull(adoParcelas("intMoeda").Value) Then
                intMoedaAtual = adoParcelas("intMoeda").Value
            Else
                gstrErrorInStoredProcedure = "Não foi encontrada Moeda Atual na tabela de Parâmetros."
                ExibeMensagem gstrErrorInStoredProcedure
                CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
                Exit Function
            End If
        Else
            gstrErrorInStoredProcedure = "Não foi encontrada Moeda Atual na tabela de Parâmetros."
            ExibeMensagem gstrErrorInStoredProcedure
            CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
            Exit Function
        End If
    Else
        gstrErrorInStoredProcedure = "Não foi possível encontrar Moeda Atual na tabela de Parâmetros."
        ExibeMensagem gstrErrorInStoredProcedure
        CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
        Exit Function
    End If
    adoParcelas.Close: Set adoParcelas = Nothing
    
    'Vamos calcular os impostos de cada Receita da Composicao
    'Vamos obter todas as receitas a serem calculadas nesta composicao de receita
    strSQL = "SELECT RC.strDescricao, VC.INTRECEITA, RC.bytTipo FROM " & gstrValorCompoRec & " VC, " & gstrReceita & " RC  WHERE RC.Pkid = VC.intReceita And VC.intComposicaoDaReceita = " & dbc_intComposicao.BoundText & " order by VC.INTRECEITA desc"

    If gobjBanco.CriaADO(strSQL, 5, adoReceita) Then
                 
        ReDim aImpostos(3, 0)
                
        If Not adoReceita.EOF Then
                    
            For intForReceita = 0 To adoReceita.RecordCount - 1
                    
                'Vamos obter as formulas de calculo de cada receita
                strSQL = "SELECT RE.blnECalculada, RE.dblValor, RE.blnParcelar, FC.strNome FROM " & gstrReceitasExercicio & " RE, " & gstrFormulaDeCalculo & " FC WHERE RE.intReceita = " & adoReceita("intReceita").Value & " AND RE.intExercicio = " & IIf(bytTipoComposicao = TYP_ISS_CONSTRUCAO, dbc_intExercicioISS.Text, dbc_intExercicioInicial.Text) & " AND RE.intFormulaDeCalculo " & strOUTJSQLServer & "= FC.pkid " & strOUTJOracle
                'strSql = "SELECT strNome FROM " & gstrFormulaDeCalculo & " WHERE intReceita = " & adoReceita("intReceita").Value
                
                If gobjBanco.CriaADO(strSQL, 5, adoFormula) Then
                            
                    If Not adoFormula.EOF Then
                    
                        For intForFormula = 0 To adoFormula.RecordCount - 1
                            
                            'Vamos verificar se é calculada ou valor fixo
                            If adoFormula("blnECalculada").Value <> 0 Then
                            
                                'Vamos executar a procedure de cada formula (Prevendo como parametros: PkidLancamentoAlfa, Composicao, PkidImobiliario/Economico, Receita, Exercicio, Simulacao, CodUsuario, Numero sequencial, Receitas Sujeitas, e Emissao)
                                strSQL = gstrStoredProcedure(adoFormula("strNome").Value, lngPkidLancamentoAlfa & ", " & dbc_intComposicao.BoundText & ", " & lngPkidPrincipal & ", " & adoReceita("intReceita").Value & ", " & IIf(bytTipoComposicao = TYP_ISS_CONSTRUCAO, dbc_intExercicioISS.Text, dbc_intExercicioInicial.Text) & ", " & chk_Simulado & ", " & glngCodUsr & ", " & intNumeroSequencial & ", '" & strReceitasSujeitas & "', '" & IIf(bytTipoComposicao = TYP_ISS_CONSTRUCAO, String(gintLenEmissao - Len(Trim(dbc_strEmissaoIss.Text)), "0") & dbc_strEmissaoIss.Text, String(gintLenEmissao - Len(Trim(dbc_strEmissao.Text)), "0") & dbc_strEmissao.Text) & "'")
                                
                                'Variavel que controla o numero de passagens da mesma Inscricao
                                intNumeroSequencial = intNumeroSequencial + 1

                                If gobjBanco.ExecuteStoredProcedure(strSQL, 10, , adoParameters) Then
                                
                                    'Valor do imposto obtido na formula de calculo
                                    If Not (adoParameters Is Nothing) Then
                                        
                                        '***ESTA PARTE ESTA FORA DO PADRAO DE CALCULO, NO CASO DE ISS CONSTRUCAO***
                                        'Caso seja composicao ISS Construcao, o calculo ja esta sendo feito na tela
                                        If bytTipoComposicao = TYP_ISS_CONSTRUCAO Then
                                            
                                            'Vamos gravar dados na tabela 'TBLLANCTOISSCONSTRUCAO'
                                            If chk_Simulado.Value = False Then
                                                
                                                gobjBanco.Execute "INSERT INTO " & gstrLanctoIssConstrucao & "(intLancamentoAlfa, strCodigoProcesso , intExercicioProcesso, bitDigitoProcesso, dtmLancamento, strObservacoes, dtmDtAtualizacao, lngCodUsr) VALUES (" & _
                                                                                  adoParameters("V_lngPkidLancamentoAlfa").Value & ",'" & txtstrCodigoProcesso & "'," & gstrENulo(txtintExercicioProcesso, , True) & "," & gstrENulo(txtbitDigitoProcesso, , True) & "," & gstrConvDtParaSql(gstrDataDoSistema) & ",'" & Trim(txtstrObservacoes) & "'," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                                                                                                                                  
                                                'Vamos gravar dados na tabela 'TBLLANCTOISSCONSTRUCAOPREDIOS'
                                                For intFor = 0 To vetPredios.UpperBound(1)
                                                    
                                                    gobjBanco.Execute "INSERT INTO " & gstrLanctoIssConstrucaoPredios & "(intLanctoIssConstrucao, dblAreaConstruida, dblAreaLancada, dtmDataConstrucao, strTipoConstrucao, strTipoAcabamento, dblPorcDemolicao, dblValorM2, dblValorServico, dblAliquotaIss, dblValorLancto, dblValorAbatido, dtmDtAtualizacao, lngCodUsr) VALUES (" & _
                                                                                       glngRetornaPkidTabelaPai("seqtblLanctoIssConstrucao", gstrLanctoIssConstrucao) & "," & gstrConvVrParaSql(vetPredios(intFor, PREDIO_MEDIDAAREAORIG)) & "," & gstrConvVrParaSql(vetPredios(intFor, PREDIO_MEDIDAAREA)) & "," & gstrConvDtParaSql(vetPredios(intFor, PREDIO_DATACONSTRUCAO)) & ",'" & vetPredios(intFor, PREDIO_STRISSCONSTRUCAOTIPO) & "','" & vetPredios(intFor, PREDIO_STRISSCONSTRUCAOPADRAO) & "'," & gstrConvVrParaSql(vetPredios(intFor, PREDIO_PORCDEMOLICAO)) & "," & gstrConvVrParaSql(vetPredios(intFor, PREDIO_VALORM2SERVICO)) & "," & gstrConvVrParaSql(vetPredios(intFor, PREDIO_VALORSERVICO)) & "," & gstrConvVrParaSql(vetPredios(intFor, PREDIO_ALIQUOTA)) & "," & gstrConvVrParaSql(vetPredios(intFor, PREDIO_ISSDEVIDO)) & "," & gstrConvVrParaSql(vetPredios(intFor, PREDIO_ISSABATIMENTO)) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                                                
                                                    'Vamos somar o Valor Total do imposto
                                                    dblValorImposto = dblValorImposto + gstrConvVrDoSql(vetPredios(intFor, PREDIO_ISSAPAGAR), 2)

                                                Next
                                            
                                                
                                                'Vamos criar um array para termos os valores individuais por receita
                                                If aImpostos(1, 0) <> "" Then ReDim Preserve aImpostos(3, UBound(aImpostos, 2) + 1)
                                                aImpostos(0, UBound(aImpostos, 2)) = CCur(aImpostos(0, UBound(aImpostos, 2)) & 0) + gstrConvVrDoSql(dblValorImposto)
                                                aImpostos(1, UBound(aImpostos, 2)) = adoReceita("intReceita").Value
                                                aImpostos(2, UBound(aImpostos, 2)) = adoFormula("blnParcelar").Value
                                                aImpostos(3, UBound(aImpostos, 2)) = adoReceita("bytTipo").Value
                                                
                                                'Vamos atualizar valores de variaveis que sao definidas nas formulas (no caso de Economico, pois no imobiliario ja sao definidas no calculoimobiliario)
                                                lngPkidLancamentoAlfa = adoParameters("V_lngPkidLancamentoAlfa").Value
                                                
                                                'Alteração do End. de Notificação
                                                strSQL = "UPDATE "
                                                strSQL = strSQL & gstrLancamentoAlfa & " "
                                                strSQL = strSQL & "SET strLogradouroC = '" & Trim(txtstrLogradouroC.Text) & "', "
                                                strSQL = strSQL & "strNumeroC = '" & Trim(txtstrNumeroC.Text) & "', "
                                                strSQL = strSQL & "strComplementoC = '" & Trim(txtstrComplementoC.Text) & "', "
                                                strSQL = strSQL & "strBairroC = '" & Trim(txtstrBairroC.Text) & "', "
                                                strSQL = strSQL & "strMunicipioC = '" & Trim(txtstrMunicipioC.Text) & "', "
                                                strSQL = strSQL & "strUFC = '" & Trim(txtstrUFC.Text) & "', "
                                                strSQL = strSQL & "intCepC = " & Replace(gstrConvVrParaSql(Trim(txtintCEPC.Text)), "-", "") & " "
                                                strSQL = strSQL & "WHERE pkID = " & lngPkidLancamentoAlfa
                                                
                                                gobjBanco.Execute strSQL
                                                
                                            End If
                                            
                                        Else
                                            'Vamos somar o Valor Total do imposto
                                            If Not blnIssVariavel Then blnIssVariavel = adoParameters("V_blnIssVariavel").Value
                                            If blnIssVariavel Then dtmAbertura = adoParameters("V_dtmAbertura").Value
                                            
                                            If CDbl(gstrConvVrDoSql(adoParameters("V_dblValorImposto").Value, 2)) > 0 Or blnIssVariavel Then
                                                dblValorImposto = dblValorImposto + gstrConvVrDoSql(adoParameters("V_dblValorImposto").Value, 2)
                                                
                                                'Vamos criar um array para termos os valores individuais por receita
                                                If aImpostos(1, 0) <> "" Then ReDim Preserve aImpostos(3, UBound(aImpostos, 2) + 1)
                                                aImpostos(0, UBound(aImpostos, 2)) = CCur(aImpostos(0, UBound(aImpostos, 2)) & 0) + gstrConvVrDoSql(adoParameters("V_dblValorImposto").Value)
                                                aImpostos(1, UBound(aImpostos, 2)) = adoReceita("intReceita").Value
                                                aImpostos(2, UBound(aImpostos, 2)) = adoFormula("blnParcelar").Value
                                                aImpostos(3, UBound(aImpostos, 2)) = adoReceita("bytTipo").Value
                                                
                                                'Vamos atualizar valores de variaveis que sao definidas nas formulas (no caso de Economico, pois no imobiliario ja sao definidas no calculoimobiliario)
                                                lngPkidLancamentoAlfa = Val(gstrENulo(adoParameters("V_lngPkidLancamentoAlfa").Value))
                                                strReceitasSujeitas = Space$(0) & adoParameters("V_strRetReceitasSujeitas").Value
                                            
                                                'Vamos forcar no caso de ja existir uma receita do tipo iss variavel nao alterar
                                            End If
                                        End If
                                        
                                    End If
                                    
                                Else
                                    CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_RECEITA
                                    Exit Function
                                End If
                            
                            Else
                                
                                'Vamos somar o Valor Total do imposto
                                dblValorImposto = dblValorImposto + gstrConvVrDoSql(adoFormula("dblValor").Value, 2)
                                
                                'Vamos criar um array para termos os valores individuais por receita
                                If aImpostos(1, 0) <> "" Then ReDim Preserve aImpostos(3, UBound(aImpostos, 2) + 1)
                                aImpostos(0, UBound(aImpostos, 2)) = CCur(aImpostos(0, UBound(aImpostos, 2)) & 0) + gstrConvVrDoSql(adoFormula("dblValor").Value)
                                aImpostos(1, UBound(aImpostos, 2)) = adoReceita("intReceita").Value
                                aImpostos(2, UBound(aImpostos, 2)) = adoFormula("blnParcelar").Value
                                aImpostos(3, UBound(aImpostos, 2)) = adoReceita("bytTipo").Value
                                
                            End If
                            
                            adoFormula.MoveNext
                            
                        Next
                        
                        adoFormula.Close: Set adoFormula = Nothing
                            
                    'Else
                    '    ExibeMensagem "Não foi(ram) encontrada(s) Fórmulas de Calculo para a Receita " & adoReceita("strDescricao").Value & "."
                    '    CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
                    '    Exit Function
                    End If
                            
                Else
                    gstrErrorInStoredProcedure = "Não foi possível encontrar Fórmulas de Calculo para a Receita " & adoReceita("strDescricao").Value & "."
                    ExibeMensagem gstrErrorInStoredProcedure
                    CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
                    Exit Function
                End If
                    
                adoReceita.MoveNext
                    
            Next
                    
            adoReceita.Close: Set adoReceita = Nothing
                
        Else
            gstrErrorInStoredProcedure = "Não foi(ram) encontrada(s) Receitas para esta Composição de Receita."
            ExibeMensagem gstrErrorInStoredProcedure
            CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
            Exit Function
        End If
                
    Else
        gstrErrorInStoredProcedure = "Não foi possível encontrar Receitas para esta Composição de Receita."
        ExibeMensagem gstrErrorInStoredProcedure
        CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
        Exit Function
    End If
    
    'Vamos verificar se existe algum valor de imposto - exceto Iss Variavel
    If Not blnIssVariavel Then
        If dblValorImposto = 0 Then
            GoTo NaoGerarParcelas
        End If
    End If
    
    'Caso nao tenha sido gerado alfa vamos sair da sub
    If lngPkidLancamentoAlfa = 0 Then Exit Function
    
    'Vamos aplicar o desconto de Valor de Compensacao
    If gobjBanco.CriaADO("SELECT dblValorCompensacao FROM " & gstrLancamentoAlfa & " " & strREADPAST & " WHERE Pkid = " & lngPkidLancamentoAlfa, 5, adoLancamentoAlfa) Then
                    
        If Not adoLancamentoAlfa.EOF Then
        
            If adoLancamentoAlfa("dblValorCompensacao") > 0 Then
                
                'Vamos obter a proporcao da compensacao para cada receita
                dblPorcentagemDaReceita = adoLancamentoAlfa("dblValorCompensacao") / dblValorImposto * 100
                
                'Vamos aplicar a compensacao nas receitas
                For intForReceita = 0 To UBound(aImpostos, 2)
                                        
                    'Nao vamos gravar valores zerados
                    If dblPorcentagemDaReceita <> 0 Then
                                
                        aImpostos(0, intForReceita) = gstrConvVrDoSql(CCur(aImpostos(0, intForReceita) & 0) - (aImpostos(0, intForReceita) * dblPorcentagemDaReceita) / 100)
                        dblValorImpostoCompensacao = dblValorImpostoCompensacao + aImpostos(0, intForReceita)
                        
                    End If
                
                Next
                
                dblValorImposto = dblValorImposto - gstrConvVrDoSql(adoLancamentoAlfa("dblValorCompensacao"), 2)
                                
                'Caso valor calculado for <= 0 nao vamos gerar parcelas
                If dblValorImposto <= 0 Then
                    'Vamos alimentar o array do relatorio de contribuinte  com valor a devolver
                    vetRelLanctoDevolver(vetRelLanctoDevolver.UpperBound(1), 1) = dblValorImposto
                    GoTo NaoGerarParcelas
                Else
                    'Vamos excluir este registro do array pois so interessam registros com valores a devolver
                    vetRelLanctoDevolver.DeleteRows vetRelLanctoDevolver.UpperBound(1)
                End If
                
                'Vamos verificar se ha diferenca entre o valor do imposto com o valor da soma das receita
                If dblValorImposto <> dblValorImpostoCompensacao Then
                
                    'Vamos aplicar a diferenca na primeira receita com valor <> de zero
                    For intForReceita = 0 To UBound(aImpostos, 2)
                        If aImpostos(0, intForReceita) <> 0 Then
                            aImpostos(0, intForReceita) = CCur(aImpostos(0, intForReceita) & 0) - (dblValorImpostoCompensacao - dblValorImposto)
                            Exit For
                        End If
                    Next
                
                End If
                
            End If
        
        End If
    
    End If
    
    'Vamos dividir os valores do imposto por receitas parceladas e nao parceladas
'    dblValorImposto = 0
'    dblValorImpostoNaoParcelado = 0
'    For intForReceita = 0 To UBound(aImpostos, 2)
'        If Not CStr(aImpostos(2, intForReceita)) = "" Then
'            If aImpostos(2, intForReceita) = 0 Then
'                dblValorImposto = dblValorImposto + aImpostos(0, intForReceita)
'            Else
'                dblValorImpostoNaoParcelado = dblValorImpostoNaoParcelado + aImpostos(0, intForReceita)
'            End If
'        End If
'    Next
    
    'Vamos obter os planos de pagamento
    strSQL = "SELECT Pkid, dblDesconto, strTitulo, intIndexadorEconomico, intDesctoTaxas FROM " & gstrParametroIPTUPagto & " " & strREADPAST & " WHERE intParametroIptu = " & RetornaPkidParametroIPTU(IIf(bytTipoComposicao = TYP_ECONOMICA, gstrEconomico, gstrImobiliario), lngPkidPrincipal, IIf(bytTipoComposicao = TYP_ISS_CONSTRUCAO, String(gintLenEmissao - Len(Trim(dbc_strEmissaoIss.Text)), "0") & dbc_strEmissaoIss.Text, ""))
    
    If gobjBanco.CriaADO(strSQL, 5, adoPlanosPagto) Then
                    
        If Not adoPlanosPagto.EOF Then
        
            For intForPlanosPagto = 0 To adoPlanosPagto.RecordCount - 1
                    
                intPrimeiraParcelaDoPlano = 0
                
                'Vamos dividir os valores do imposto por receitas parceladas e nao parceladas
                dblValorImposto = 0
                dblValorImpostoNaoParcelado = 0
                
                'Vamos utilizar um array auxiliar para aplicar descontos por receitas, no final do for retornaremos o original
                aImpostosAux = aImpostos
                
                For intForReceita = 0 To UBound(aImpostos, 2)
                    
                    'Caso nao exista receita nao gravaremos este lancamento
                    If aImpostos(0, 0) = "" Then
                        CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_RECEITA
                        gstrErrorInStoredProcedure = "Não existem receitas, pois o valor do lançamento é Zero."
                        Exit Function
                    End If
                    
                    'Vamos verificar o tipo de receita e aplicar o desconto
                    If aImpostos(3, intForReceita) = 2 Then 'Tipo Imposto
                        aImpostos(0, intForReceita) = aImpostos(0, intForReceita) - (aImpostos(0, intForReceita) * (IIf(IsNull(adoPlanosPagto("dblDesconto")), 0, adoPlanosPagto("dblDesconto")) / 100))
                    ElseIf aImpostos(3, intForReceita) = 3 Then 'Tipo Taxa
                        aImpostos(0, intForReceita) = aImpostos(0, intForReceita) - (aImpostos(0, intForReceita) * (IIf(IsNull(adoPlanosPagto("intDesctoTaxas")), 0, adoPlanosPagto("intDesctoTaxas")) / 100))
                    End If
                    
                    If Not CStr(aImpostos(2, intForReceita)) = "" Then
                        If aImpostos(2, intForReceita) = 0 Then
                            dblValorImposto = dblValorImposto + aImpostos(0, intForReceita)
                        Else
                            dblValorImpostoNaoParcelado = dblValorImpostoNaoParcelado + aImpostos(0, intForReceita)
                        End If
                    End If
                Next
                
                '***Os descontos ja estao sendo aplicados nas receitas***
                'Vamos aplicar o desconto do plano corrente no total de impostos calculado
                'dblValorImpostoDesconto = dblValorImposto - (dblValorImposto * (IIf(IsNull(adoPlanosPagto("dblDesconto")), 0, adoPlanosPagto("dblDesconto")) / 100))
                dblValorImpostoDesconto = dblValorImposto
                'Vamos aplicar o desconto do plano corrente no total de impostos nao parcelados
                'dblValorImpostoNaoParceladoDesc = dblValorImpostoNaoParcelado - (dblValorImpostoNaoParcelado * (IIf(IsNull(adoPlanosPagto("dblDesconto")), 0, adoPlanosPagto("dblDesconto")) / 100))
                dblValorImpostoNaoParceladoDesc = dblValorImpostoNaoParcelado
                
                'Vamos obter as parcelas do plano corrente
                strSQL = "SELECT FP.intParcela, FP.dtmDtVencimento, PP.bytParcelado FROM " & gstrFormaPagtoVencimentos & " FP," & gstrParametroIPTUPagto & " PP WHERE FP.intFormaPagto = PP.Pkid AND PP.Pkid = " & adoPlanosPagto("Pkid").Value & " Order By FP.intParcela "
                    
                If gobjBanco.CriaADO(strSQL, 5, adoParcelas) Then
                            
                    If Not adoParcelas.EOF Then
                        
                        'Vamos verificar se a qtde de parcelas esta dentro do valor minimo por parcela
                        intQtdeParcelas = VerificaValorMinimoPorParcela(dblValorImpostoDesconto, adoParcelas.RecordCount)
                        
                        'Vamos verificar se existe valor para parcelar
                        If dblValorImposto > 0 Or blnIssVariavel Then
                        
                            'Caso seja ISS Variavel vamos parcelar com o valor sem dividir por parcela
                            If blnIssVariavel Then
                                'Vamos atribuir o valor por parcela
                                dblValorPorParcela = gstrConvVrDoSql(dblValorImpostoDesconto, 2)
                            Else
                                'Vamos calcular o valor por parcela
                                dblValorPorParcela = gstrConvVrDoSql(dblValorImpostoDesconto / intQtdeParcelas, 2)
                                
                                'Vamos verificar se a diferenca do valor total das receitas com o valor total da parcela ja atualizado
                                dblValorDiferencaParcela = gstrConvVrDoSql((dblValorPorParcela * intQtdeParcelas) - dblValorImpostoDesconto, 2)
                            End If
                        
                            For intForParcelas = 0 To intQtdeParcelas - 1
                                
                                'Caso seja ISS Variavel vamos verificar a data da abertura definir as parcelas
                                If blnIssVariavel Then
                                    'Se a abertura for no mesmo exercicio, so vamos gerar parcelas com mês superior à abertura
                                    If Year(dtmAbertura) = Year(gstrDataDoSistema) Then
                                        If CDate("01/" & Month(dtmAbertura) & "/" & Year(dtmAbertura)) >= CDate("01/" & Month(adoParcelas("dtmDtVencimento").Value) & "/" & Year(adoParcelas("dtmDtVencimento").Value)) Then
                                            GoTo ProximaParcela
                                        End If
                                    End If
                                End If
                                                
                                'Vamos armazenar a 1ª parcela do plano para utilizar no caso de valor nao parcelado
                                If intPrimeiraParcelaDoPlano = 0 Then
                                    intPrimeiraParcelaDoPlano = adoParcelas("intParcela").Value
                                End If
                                
                                'Vamos gravar as parcelas na tabela tblLancamentoValor
                                strSQL = "INSERT INTO " & gstrLancamentoValor & " " & _
                                         "(intLancamentoAlfa, intParcela, dtmDtVencimento, dblValor, intMoeda, bitParcelaValida, dtmDtAtualizacao, lngCodUsr)" & _
                                         " VALUES " & _
                                         "(" & lngPkidLancamentoAlfa & "," & adoParcelas("intParcela").Value & "," & gstrConvDtParaSql(adoParcelas("dtmDtVencimento").Value) & "," & gstrConvVrParaSql(dblValorPorParcela - dblValorDiferencaParcela) & "," & intMoedaAtual & "," & adoParcelas("bytParcelado").Value & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                                  
                                gobjBanco.Execute strSQL
                                     
                                dblValorDiferencaParcela = 0
                                
                                'Vamos varrer o array com Receitas e Valores, para gravar as Receitas com a proporcao de valor de cada receita
                                For intForFormula = 0 To UBound(aImpostos, 2)
                                            
                                    'Vamos verificar se a receita é parcelada
                                    If aImpostos(2, intForFormula) = 0 Then
                                        
                                        'Vamos obter a proporcao do valor para a receita
                                        If CDbl(gstrConvVrDoSql(aImpostos(0, intForFormula), , , True)) <> 0 And dblValorImposto <> 0 Then
                                            dblPorcentagemDaReceita = aImpostos(0, intForFormula) / dblValorImposto
                                        End If
                                    
                                        'Nao vamos gravar valores zerados
                                        If dblPorcentagemDaReceita <> 0 Or blnIssVariavel Then
                                        
                                            'Vamos gravar as receitas na tabela tblLancamentoReceita
                                            strSQL = "INSERT INTO " & gstrLancamentoReceita & " " & _
                                                     "(intLancamentoValor, intReceita, dblValor, dtmDtAtualizacao, lngCodUsr)" & _
                                                     "  " & _
                                                     "(SELECT " & glngRetornaPkidTabelaPai("seqtbllancamentovalor", gstrLancamentoValor) & ", " & aImpostos(1, intForFormula) & "," & gstrConvVrParaSql(dblValorPorParcela * dblPorcentagemDaReceita) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & IIf(bytDBType = EDatabases.Oracle, " FROM Dual)", ")")
                                                        
                                            gobjBanco.Execute strSQL
                                            
                                        End If
                                    
                                    End If
                                    
                                Next
                                        
ProximaParcela:
                                
                                adoParcelas.MoveNext
    
                            Next
                            
                            'Caso exista valores nao parcelados vamos somar na 1ª parcela
                            If dblValorImpostoNaoParcelado > 0 Then
                            
                                'Vamos atribuir o valor da parcela
                                dblValorPorParcela = gstrConvVrDoSql(dblValorImpostoNaoParceladoDesc, 2)
                                
                                'Vamos obter a 1ª parcela
                                If gobjBanco.CriaADO("SELECT Pkid FROM " & gstrLancamentoValor & " " & strREADPAST & " WHERE intLancamentoAlfa = " & lngPkidLancamentoAlfa & " And intParcela = " & intPrimeiraParcelaDoPlano, 5, adoLancamentoValor) Then
                    
                                    If Not adoLancamentoValor.EOF Then
                                    
                                        'Vamos gravar as parcelas na tabela tblLancamentoValor
                                        strSQL = "UPDATE " & gstrLancamentoValor
                                        strSQL = strSQL & " SET"
                                        strSQL = strSQL & " dblValor = dblValor + " & gstrConvVrParaSql(dblValorPorParcela)
                                        strSQL = strSQL & " WHERE pkid = " & adoLancamentoValor("Pkid").Value
                                              
                                        gobjBanco.Execute strSQL
                                            
                                        'Vamos varrer o array com Receitas e Valores, para gravar as Receitas com a proporcao de valor de cada receita
                                        For intForFormula = 0 To UBound(aImpostos, 2)
                                                    
                                            If aImpostos(2, intForFormula) = 1 Then
                                                
                                                'Vamos obter a proporcao do valor para a receita
                                                If CDbl(gstrConvVrDoSql(aImpostos(0, intForFormula), , , True)) <> 0 And dblValorImpostoNaoParcelado <> 0 Then
                                                    dblPorcentagemDaReceita = aImpostos(0, intForFormula) / dblValorImpostoNaoParcelado
                                                End If
                                                
                                                'Nao vamos gravar valores zerados
                                                If dblPorcentagemDaReceita <> 0 Or blnIssVariavel Then
                                                    
                                                    'Vamos gravar as receitas na tabela tblLancamentoReceita
                                                    strSQL = "INSERT INTO " & gstrLancamentoReceita & " " & _
                                                             "(intLancamentoValor, intReceita, dblValor, dtmDtAtualizacao, lngCodUsr)" & _
                                                             "  " & _
                                                             "(SELECT " & adoLancamentoValor("Pkid").Value & ", " & aImpostos(1, intForFormula) & "," & gstrConvVrParaSql(dblValorPorParcela * dblPorcentagemDaReceita) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & IIf(bytDBType = EDatabases.Oracle, " FROM Dual)", ")")
                                                                
                                                    gobjBanco.Execute strSQL
                                                    
                                                End If
                                                
                                            End If
                                            
                                        Next
                                        
                                        adoLancamentoValor.Close: Set adoLancamentoValor = Nothing
                                        
                                    Else
                                        gstrErrorInStoredProcedure = "Não foi possível localizar a 1ª parcela."
                                        ExibeMensagem gstrErrorInStoredProcedure
                                        CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
                                    End If
                                Else
                                    gstrErrorInStoredProcedure = "Não foi possível localizar a 1ª parcela."
                                    ExibeMensagem gstrErrorInStoredProcedure
                                    CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
                                End If
                                
                            End If
                        
                        Else 'Caso só exista valor nao parcelado
                            
                            If dblValorImpostoNaoParcelado > 0 Then
                            
                                'Vamos atribuir o valor da parcela
                                dblValorPorParcela = gstrConvVrDoSql(dblValorImpostoNaoParceladoDesc, 2)
                                    
                                'Vamos gravar as parcelas na tabela tblLancamentoValor
                                strSQL = "INSERT INTO " & gstrLancamentoValor & " " & _
                                             "(intLancamentoAlfa, intParcela, dtmDtVencimento, dblValor, intMoeda, bitParcelaValida, dtmDtAtualizacao, lngCodUsr)" & _
                                             " VALUES " & _
                                             "(" & lngPkidLancamentoAlfa & "," & adoParcelas("intParcela").Value & "," & gstrConvDtParaSql(adoParcelas("dtmDtVencimento").Value) & "," & gstrConvVrParaSql(dblValorPorParcela) & "," & intMoedaAtual & "," & adoParcelas("bytParcelado").Value & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & ")"
                                      
                                gobjBanco.Execute strSQL
                                    
                                'Vamos varrer o array com Receitas e Valores, para gravar as Receitas com a proporcao de valor de cada receita
                                For intForFormula = 0 To UBound(aImpostos, 2)
                                            
                                    If aImpostos(2, intForFormula) = 1 Then
                                        
                                        'Vamos obter a proporcao do valor para a receita
                                        If CDbl(gstrConvVrDoSql(aImpostos(0, intForFormula), , , True)) <> 0 And dblValorImpostoNaoParcelado <> 0 Then
                                            dblPorcentagemDaReceita = aImpostos(0, intForFormula) / dblValorImpostoNaoParcelado
                                        End If
                                        
                                        'Nao vamos gravar valores zerados
                                        If dblPorcentagemDaReceita <> 0 Or blnIssVariavel Then
                                        
                                            'Vamos gravar as receitas na tabela tblLancamentoReceita
                                            strSQL = "INSERT INTO " & gstrLancamentoReceita & " " & _
                                                     "(intLancamentoValor, intReceita, dblValor, dtmDtAtualizacao, lngCodUsr)" & _
                                                     "  " & _
                                                     "(SELECT " & glngRetornaPkidTabelaPai("seqtbllancamentovalor", gstrLancamentoValor) & ", " & aImpostos(1, intForFormula) & "," & gstrConvVrParaSql(dblValorPorParcela * dblPorcentagemDaReceita) & "," & gstrConvDtParaSql(gstrDataDoSistema) & "," & glngCodUsr & IIf(bytDBType = EDatabases.Oracle, " FROM Dual)", ")")
                                                        
                                            gobjBanco.Execute strSQL
                                            
                                        End If
                                        
                                    End If
                                    
                                Next
                            
                            End If
                            
                        End If
                        
                    Else
                        gstrErrorInStoredProcedure = "Não foi(ram) encontrada(s) parcelas para o Plano de Pagamento " & adoPlanosPagto("strTitulo").Value & "."
                        ExibeMensagem gstrErrorInStoredProcedure
                        CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
                        Exit Function
                    End If
                            
                Else
                    gstrErrorInStoredProcedure = "Não foi possível encontrar parcelas para o Plano de Pagamento " & adoPlanosPagto("strTitulo").Value & "."
                    ExibeMensagem gstrErrorInStoredProcedure
                    CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
                    Exit Function
                End If
                
                adoPlanosPagto.MoveNext
                    
                'Vamos redefinir originalmente o array auxiliar para aplicar descontos por receitas
                aImpostos = aImpostosAux
        
            Next
                
        Else
            gstrErrorInStoredProcedure = "Não foi(ram) encontrado(s) valores da tabela de Parâmetros de IPTU Pagamento."
            ExibeMensagem gstrErrorInStoredProcedure
            CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
            Exit Function
        End If
                
    Else
        gstrErrorInStoredProcedure = "Não foi possível retornar valores da tabela de Parâmetros de IPTU Pagamento."
        ExibeMensagem gstrErrorInStoredProcedure
        CalculaImpostos = BYT_CALCULOIMPOSTO_ERRO_GERAL
        Exit Function
    End If
    
NaoGerarParcelas:

    CalculaImpostos = BYT_CALCULOIMPOSTO_OK
    
    If bytTipoComposicao = TYP_ISS_CONSTRUCAO And chk_Carne.Value = 1 Then
       ImprimeRelatorio rptCapaCarneISSConstrucao, strQueryCarneISSConstrucao(lngPkidLancamentoAlfa), "Carne de ISS Construção."
    End If
    
End Function

Private Function InscricaoJaCadastrada(strInscricao As String, intExercicio As Integer, intComposicaoDaReceita As Long) As Long
Dim strSQL              As String
Dim adoResultado        As ADODB.Recordset

    strSQL = "SELECT LA.Pkid "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoAlfa & " LA " & strREADPAST
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & UCase(strInscricao) & "' AND LA.intExercicio = " & intExercicio & " AND intComposicaoDaReceita = " & intComposicaoDaReceita & " AND dtmDtCancelamento IS NULL"

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSQL, 15, adoResultado) Then
        If Not adoResultado.EOF Then
            InscricaoJaCadastrada = adoResultado("Pkid")
        Else
            InscricaoJaCadastrada = 0
        End If
    End If

End Function

Private Function VerificaCalculoEmAndamento() As Boolean
Dim strSQL              As String
Dim adoResultado        As ADODB.Recordset
Dim strSqlIncluiUsuario As String

    strSqlIncluiUsuario = "UPDATE " & gstrParametrosTributario
    strSqlIncluiUsuario = strSqlIncluiUsuario & " SET"
    strSqlIncluiUsuario = strSqlIncluiUsuario & " intUsuario = " & glngCodUsr
    strSqlIncluiUsuario = strSqlIncluiUsuario & ", dtmDtAtualizacao = " & strGETDATE
    strSqlIncluiUsuario = strSqlIncluiUsuario & ", lngCodUsr = " & glngCodUsr

    strSQL = "SELECT US.strNome"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrUsuarios & " US, "
    strSQL = strSQL & gstrParametrosTributario & " PA"
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PA.intUsuario = US.Pkid "

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            ExibeMensagem "Cálculo em Andamento pelo usuário(a) " & adoResultado!STRNOME
            VerificaCalculoEmAndamento = False
        Else
            If gobjBanco.Execute(strSqlIncluiUsuario) Then
                VerificaCalculoEmAndamento = True
            End If
        End If
    End If

End Function

Private Sub LiberaCalculo()
Dim strSQL As String

    strSQL = "UPDATE " & gstrParametrosTributario
    strSQL = strSQL & " SET"
    strSQL = strSQL & " intUsuario = ''"
    strSQL = strSQL & ", dtmDtAtualizacao = " & strGETDATE
    strSQL = strSQL & ", lngCodUsr = " & glngCodUsr

    Set gobjBanco = New clsBanco

    gobjBanco.Execute strSQL

End Sub

Private Sub PreencheDadosCadastro()
Dim adoResultado As New ADODB.Recordset
Dim strSQL       As String
    
    Set gobjBanco = New clsBanco
    
    'Vamos achar o endereço do Imóvel
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strQueryLogradouroImovel(dbcstrInscricao.Text), 5, adoResultado) Then
        If Not adoResultado.EOF Then
            PreencherListaDeOpcoes dbcintContribuinte, gstrENulo(adoResultado!intContribuinte)
            txtstrLogradouro = gstrENulo(adoResultado!strLogradouro)
            txtstrNumero = gstrENulo(adoResultado!INTNUMERO)
            txtstrComplemento = gstrENulo(adoResultado!STRCOMPLEMENTO)
            txtintCep = Format(gstrENulo(adoResultado!INTCEP), "00000-000")
            txtstrBairro = gstrENulo(adoResultado!strBairro) & " / " & gstrENulo(adoResultado!strEstado)
            txtstrMunicipio = gstrENulo(adoResultado!STRMUNICIPIO)
            txtstrUf = gstrENulo(adoResultado!strEstado)
            If Trim(gstrENulo(adoResultado!strlogradouroc)) = "" And Trim(gstrENulo(adoResultado!strBairroC)) = "" And Trim(gstrENulo(adoResultado!strMunicipioC)) = "" Then
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strQueryLogradouroContribuinte(dbcstrInscricao.Text), 5, adoResultado) Then
                    If Not adoResultado.EOF Then
                        txtstrLogradouroC = gstrENulo(adoResultado!strlogradouroc)
                        txtstrNumeroC = gstrENulo(adoResultado!intNumeroC)
                        txtstrComplementoC = gstrENulo(adoResultado!strComplementoC)
                        txtstrBairroC = gstrENulo(adoResultado!strBairroC)
                        txtstrMunicipioC = gstrENulo(adoResultado!strMunicipioC)
                        txtstrUFC = gstrENulo(adoResultado!strestadoc)
                        txtintCEPC = Format(gstrENulo(adoResultado!intcepc), "00000-000")
                    End If
                End If
            Else
                txtstrLogradouroC = gstrENulo(adoResultado!strlogradouroc)
                txtstrNumeroC = gstrENulo(adoResultado!intNumeroC)
                txtstrComplementoC = gstrENulo(adoResultado!strComplementoC)
                txtstrBairroC = gstrENulo(adoResultado!strBairroC)
                txtstrMunicipioC = gstrENulo(adoResultado!strMunicipioC)
                txtstrUFC = gstrENulo(adoResultado!strestadoc)
                txtintCEPC = Format(gstrENulo(adoResultado!intcepc), "00000-000")
            End If
            'Vamos carregar os predios do imobiliario selecionado
            LeDaTabelaParaObj gstrAreaImobiliario, dbcintPredios, strQueryPredios
            
            'MontaArray
            
        Else
            ExibeMensagem "Não foi(ram) encontrado(s) registro(s) com esta Inscrição no Cadastro Imobiliário."
            Exit Sub
        End If
    End If
    
End Sub

Private Function strQueryLogradouroImovel(strInscricao As String) As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "BA.strDescricao AS strBairro, "
    strSQL = strSQL & " RTRIM(LTRIM(L.strDescricao)) "
    strSQL = strSQL & strCONCAT & gstrISNULL("TL.strSigla", "''", "', '")
    strSQL = strSQL & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''")
    strSQL = strSQL & strCONCAT & gstrISNULL("U.strDescricao", "' '", "', '")
    strSQL = strSQL & strCONCAT & gstrISNULL("U.strDescricao", "''") & ")) AS strLogradouro, "
    strSQL = strSQL & "IM.intContribuinte, "
    strSQL = strSQL & "IM.intNumero, "
    strSQL = strSQL & "L.Intcep, "
    strSQL = strSQL & "IM.strComplemento, "
    strSQL = strSQL & "IM.strLote, "
    strSQL = strSQL & "IM.strQuadra, "
    strSQL = strSQL & "(SELECT MU.strDescricao FROM tblMunicipio  MU WHERE BA.intMunicipio = MU.PKId ) AS strMunicipio, "
    strSQL = strSQL & "(SELECT UF.strSigla FROM " & gstrUF & " UF WHERE UF.PKId = "
    strSQL = strSQL & "(SELECT MU.intUF FROM tblMunicipio  MU WHERE BA.intMunicipio = MU.PKId )) AS strEstado, "
    strSQL = strSQL & "IM.strLogradouroC, "
    strSQL = strSQL & "IM.intNumeroC, "
    strSQL = strSQL & "IM.strBairroC, "
    strSQL = strSQL & "IM.IntcepC, "
    strSQL = strSQL & "IM.strComplementoC, "
    strSQL = strSQL & "(SELECT MU.strDescricao FROM " & gstrCidade & " MU WHERE MU.PKId = IM.intMunicipioC) strMunicipioC, "
    strSQL = strSQL & "(SELECT UF.strSigla FROM " & gstrUF & " UF WHERE UF.PKId = IM.intUFC) strEstadoC "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrImobiliario & " IM, "
    strSQL = strSQL & gstrBairro & " BA, "
    strSQL = strSQL & gstrLogradouro & " L, "
    strSQL = strSQL & gstrTituloLogradouro & " U, "
    strSQL = strSQL & gstrTipoLogradouro & " TL "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " L.pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " IM.Intlogradouro "
    strSQL = strSQL & " AND L.intBairro = BA.PKId"
    strSQL = strSQL & " AND L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle
    'strSql = strSql & " AND L.DtmdtExclusao is null "
    strSQL = strSQL & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
    strSQL = strSQL & " AND IM.Strinscricao = " & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & UCase(strInscricao)
        
    strQueryLogradouroImovel = strSQL
    
End Function

'Private Sub MontaArray()
'Dim varAux       As Variant
'Dim strSql       As String
'Dim adoRec       As ADODB.Recordset
'Dim adoPontuacao As ADODB.Recordset
'
'    strSql = ""
'    strSql = strSql & "SELECT  AI.Pkid, AI.intNEdificacao, AI.intMedidaDaArea, AI.dtmUltimaReforma, AI.intCategoriaConstrucao, CC.strDescricao strCategoriaConstrucao, "
'    strSql = strSql & "SUM(TV.DBLVALOR) Pontos "
'    strSql = strSql & "FROM " & gstrAreaImobiliario & " AI, " & gstrCategoriaConstrucao & " CC, " & gstrTipoDeArea & " TA, " & gstrCaracteristicaDoImovel & " CI, " & gstrDetalheDaCaracteristica & " DC, "
'    strSql = strSql & gstrTabelaDeValor & " TV "
'    strSql = strSql & "WHERE AI.intImobiliario = " & dbcstrInscricao.BoundText & _
'                      " AND TA.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " AI.intTipoDeArea " & _
'                      " AND CI.INTCODIGOIMOBILIARIO = AI.intImobiliario " & _
'                      " AND CC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " AI.intCategoriaConstrucao " & _
'                      " AND DC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CI.INTCODIGODETALHEDACARACTERISTI " & _
'                      " AND TV.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " DC.INTTABELADEVALORES " & _
'                      " AND CI.intArea = AI.Pkid "
'    strSql = strSql & "GROUP BY AI.Pkid, AI.intNEdificacao, AI.intMedidaDaArea, AI.dtmUltimaReforma, AI.intCategoriaConstrucao, CC.strDescricao"
'
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoRec) Then
'        If Not adoRec.EOF Then
'
'            strSql = ""
'            strSql = strSql & "SELECT FP.strDescricao FROM " & gstrFaixaPontosPredio & " FP WHERE FP.Intcategoriaconstrucao = " & adoRec("intCategoriaConstrucao").Value & " AND (dblPontoinicial <= " & adoRec("Pontos").Value & " AND dblPontoFinal >= " & adoRec("Pontos").Value & ")"
'
'            If gobjBanco.CriaADO(strSql, 5, adoPontuacao) Then
'
'                Set vetPredios = New XArrayDB
'                vetPredios.Clear
'                With adoRec
'                    vetPredios.ReDim 0, .RecordCount - 1, 0, 18
'                    Do While Not .EOF
'                        varAux = adoRec("Pkid").Value
'                        vetPredios(.AbsolutePosition - 1, PREDIO_PKID) = varAux
'                        varAux = adoRec("intNEdificacao").Value
'                        vetPredios(.AbsolutePosition - 1, PREDIO_NEDIFICACAO) = varAux
'                        varAux = adoRec("intMedidaDaArea").Value
'                        vetPredios(.AbsolutePosition - 1, PREDIO_MEDIDAAREA) = varAux
'                        varAux = adoRec("dtmUltimaReforma").Value
'                        vetPredios(.AbsolutePosition - 1, PREDIO_DATACONSTRUCAO) = varAux
'                        varAux = adoRec("intCategoriaConstrucao").Value
'                        vetPredios(.AbsolutePosition - 1, PREDIO_INTCATEGORIACONSTRUCAO) = varAux
'                        varAux = adoRec("strCategoriaConstrucao").Value
'                        vetPredios(.AbsolutePosition - 1, PREDIO_STRCATEGORIACONSTRUCAO) = varAux
'                        varAux = adoPontuacao("strDescricao").Value
'                        vetPredios(.AbsolutePosition - 1, PREDIO_PADRAO) = varAux
'                        vetPredios(.AbsolutePosition - 1, PREDIO_DEMOLICAO) = 0
'                        .MoveNext
'                    Loop
'    '                Else
'    '                    vetPredios.ReDim 0, 0, 0, 18
'    '                    vetPredios(0, 0) = ""
'    '                    vetPredios(0, 1) = ""
'    '                    vetPredios(0, 2) = ""
'    '                    vetPredios(0, 3) = ""
'    '                    vetPredios(0, 4) = ""
'    '                    vetPredios(0, 5) = ""
'    '                    vetPredios(0, 6) = ""
'    '                End If
'                End With
'
'                Set tdb_Predios.Array = vetPredios
'                tdb_Predios.Rebind
'                tdb_Predios.Refresh
'
'            End If
'        End If
'    End If
'End Sub

Private Sub LimpaDadosISSConstrucao(blnTodos As Boolean)
        
    If blnTodos Then
    
        txtstrCodigoProcesso = Space$(0)
        txtintExercicioProcesso = Space$(0)
        txtbitDigitoProcesso = Space$(0)
    
        txtstrLogradouro = Space$(0)
        txtstrNumero = Space$(0)
        txtstrComplemento = Space$(0)
        txtintCep = Space$(0)
        txtstrBairro = Space$(0)
        txtstrMunicipio = Space$(0)
        txtstrUf = Space$(0)
        txtstrLogradouroC = Space$(0)
        txtstrNumeroC = Space$(0)
        txtstrComplementoC = Space$(0)
        txtstrBairroC = Space$(0)
        txtstrMunicipioC = Space$(0)
        txtstrUFC = Space$(0)
        txtintCEPC = Space$(0)
        txtstrObservacoes = Space$(0)
    
        Set vetPredios = New XArrayDB
        vetPredios.Clear
        vetPredios.ReDim 0, 0, 0, 19
        PkidArray = 0
        
        Set tdb_Predios.Array = vetPredios
        tdb_Predios.ReBind
        tdb_Predios.Refresh
        
        Set dbcintPredios.RowSource = Nothing
    
    End If
    
    dbcintPredios.BoundText = Space$(0)
    txtstrCategoriaConstrucao.Text = Space$(0)
    txtstrPadrao.Text = Space$(0)

    txtdblArea = Space$(0)
    txtdblArea.Tag = Space$(0)
    txtdtmConstrucao = Space$(0)
    dbcintIssConstrucaoTipo.BoundText = Space$(0)
    dbcintAcabamento.BoundText = Space$(0)
    chkintDemolicao.Value = vbUnchecked
    chkintDemolicao.Tag = ""
    txtdblVlrM2Servico = Space$(0)
    txtdblVlrServico = Space$(0)
    txtdblAliquota = Space$(0)
    txtdblIssDevido = Space$(0)
    txtdblIssAbatimento = Space$(0)
    txtdblIssPagar = Space$(0)

End Sub

Private Function strQueryIssConstrucaoTipo() As String
Dim strSQL As String

    strSQL = "SELECT PKid, strDescricao FROM " & gstrIssConstrucaoTipo
    strSQL = strSQL & " ORDER BY strDescricao"
    
    strQueryIssConstrucaoTipo = strSQL
    
End Function

Private Function strQueryIssConstrucaoPadrao() As String
Dim strSQL As String

    strSQL = " SELECT CP.PKid, CP.strDescricao " & _
             " FROM " & gstrIssConstrucaoPadrao & " CP "
    
    If dbcintIssConstrucaoTipo.MatchedWithList Then
        strSQL = strSQL & " WHERE CP.intIssConstrucaoTipo = " & dbcintIssConstrucaoTipo.BoundText
    End If
             
    strSQL = strSQL & " ORDER BY strDescricao"
    
    strQueryIssConstrucaoPadrao = strSQL
    
End Function

Private Function strQueryPredios() As String
Dim strSQL As String

    strSQL = "SELECT  AI.Pkid, AI.intNEdificacao "
    strSQL = strSQL & "FROM " & gstrAreaImobiliario & " AI "
    strSQL = strSQL & "WHERE AI.intImobiliario = " & dbcstrInscricao.BoundText
    
    strQueryPredios = strSQL
    
End Function

Private Function strQueryCategoriaConstrucao() As String
Dim strSQL As String

    strSQL = "SELECT PKid, strDescricao FROM " & gstrCategoriaConstrucao
    strSQL = strSQL & " ORDER BY strDescricao"
    
    strQueryCategoriaConstrucao = strSQL
    
End Function

Private Sub IncluiValoresNoGrid()
Dim intFor           As Integer
Dim blnAlterarPredio As Boolean

    blnAlterarPredio = False
    
    'Forcada a realizacao do evento LostFocus do valor para atualizacao de valores
    txtdblArea_LostFocus
    
    For intFor = 0 To vetPredios.UpperBound(1)
    
        If Val(dbcintPredios.BoundText) > 0 And Val(vetPredios(intFor, PREDIO_PKID) & Space$(0)) = Val(dbcintPredios.BoundText) Then
            
            If MsgBox("Este prédio já se encontra selecionado. Deseja alterar o existente", vbYesNo, "Mensagem ao Usuário") = vbYes Then
                blnAlterarPredio = True
                Exit For
            Else
                Exit Sub
            End If
            
        End If
        
    Next
        
    'Nao vamos deixar inserir predio com valor negativo
    If Val(txtdblIssPagar.Text) < 0 Then
        ExibeMensagem "Não é permitido inserir um prédio com valor negativo."
        Exit Sub
    End If
    
    If vetPredios.UpperBound(1) = -1 Then
        vetPredios.ReDim 0, 0, 0, 19
    End If
    
    If Len(vetPredios(0, PREDIO_NEDIFICACAO)) <> 0 Then
        vetPredios.ReDim 0, vetPredios.UpperBound(1) + 1, 0, 19
    End If
            
    If blnAlterarPredio Then vetPredios.DeleteRows intFor
    
    'Ponteiro para localizacao de predios no array no momento de excluir do array
    PkidArray = PkidArray + 1
    
    vetPredios(vetPredios.UpperBound(1), PREDIO_PKID) = dbcintPredios.BoundText
    vetPredios(vetPredios.UpperBound(1), PREDIO_NEDIFICACAO) = IIf(dbcintPredios.MatchedWithList, dbcintPredios.Text, "X")
    vetPredios(vetPredios.UpperBound(1), PREDIO_STRCATEGORIACONSTRUCAO) = txtstrCategoriaConstrucao.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_PADRAO) = txtstrPadrao.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_MEDIDAAREA) = gstrConvVrDoSql(txtdblArea.Text, 2)
    vetPredios(vetPredios.UpperBound(1), PREDIO_MEDIDAAREAORIG) = gstrConvVrDoSql(txtdblArea.Tag, 2)
    vetPredios(vetPredios.UpperBound(1), PREDIO_DATACONSTRUCAO) = txtdtmConstrucao.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_INTISSCONSTRUCAOTIPO) = dbcintIssConstrucaoTipo.BoundText
    vetPredios(vetPredios.UpperBound(1), PREDIO_STRISSCONSTRUCAOTIPO) = dbcintIssConstrucaoTipo.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_INTISSCONSTRUCAOPADRAO) = dbcintAcabamento.BoundText
    vetPredios(vetPredios.UpperBound(1), PREDIO_STRISSCONSTRUCAOPADRAO) = dbcintAcabamento.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_DEMOLICAO) = chkintDemolicao.Value
    vetPredios(vetPredios.UpperBound(1), PREDIO_VALORM2SERVICO) = txtdblVlrM2Servico.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_VALORSERVICO) = txtdblVlrServico.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_ALIQUOTA) = txtdblAliquota.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_ISSDEVIDO) = txtdblIssDevido.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_ISSABATIMENTO) = gstrConvVrDoSql(txtdblIssAbatimento.Text, 2)
    vetPredios(vetPredios.UpperBound(1), PREDIO_ISSAPAGAR) = txtdblIssPagar.Text
    vetPredios(vetPredios.UpperBound(1), PREDIO_PKIDARRAY) = PkidArray
    vetPredios(vetPredios.UpperBound(1), PREDIO_PORCDEMOLICAO) = chkintDemolicao.Tag
    
    Set tdb_Predios.Array = vetPredios
    tdb_Predios.ReBind
    tdb_Predios.Refresh
    
End Sub

Private Sub ExcluiValoresDoGrid()
Dim intFor As Integer
        
    For intFor = 0 To vetPredios.UpperBound(1)
    
        If Val(tdb_Predios.Columns("PkidArray").Value) = Val(vetPredios(intFor, PREDIO_PKIDARRAY)) Then
            
            vetPredios.DeleteRows intFor
            Exit For
        
        End If
        
    Next

    Set tdb_Predios.Array = vetPredios
    tdb_Predios.ReBind
    tdb_Predios.Refresh
        
    LimpaDadosISSConstrucao False
    
End Sub

Private Sub CarregaValoresDoIssConstrucao()
Dim strSQL       As String
Dim adoRec       As ADODB.Recordset
Dim dblIndexador As Double

    If Not dbcintIssConstrucaoTipo.MatchedWithList Or Not dbcintAcabamento.MatchedWithList Or Len(txtdtmConstrucao.Text) = 0 Then
        txtdblVlrM2Servico.Text = Space$(0)
        txtdblIssDevido.Text = Space$(0)
        txtdblIssPagar.Text = Space$(0)
        chkintDemolicao.Tag = ""
    Else
        strSQL = ""
        strSQL = strSQL & "SELECT  IV.dblValorM2, IT.dblPorcDemolicao, IT.dblAliquotaIss, "
        strSQL = strSQL & "(SELECT FA.dblvalor FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = IE.INTINDEXADORECONOMICO and FA.dtmdata = " & gstrConvDtParaSql(gstrDataDoSistema, True) & ") DataAtual, "
        strSQL = strSQL & "(SELECT FA.dblvalor FROM " & gstrFormaAtualizacaoValor & " FA WHERE FA.INTINDEXADORECONOMICO = IE.INTINDEXADORECONOMICO and FA.dtmdata = " & gstrConvDtParaSql("01/" & Month(gstrDataDoSistema) & "/" & Year(gstrDataDoSistema), True) & " ) DataMes "
        strSQL = strSQL & "FROM " & gstrIssConstrucaoTipo & " IT, " & gstrIssConstrucaoVlrM2 & " IV, " & gstrTipoPadraoExercicio & " TP, " & gstrIssConstrucaoExercicio & " IE "
        strSQL = strSQL & "WHERE TP.intIssConstrucaoTipo = " & dbcintIssConstrucaoTipo.BoundText & _
                          " AND TP.intIssConstrucaoPadrao = " & dbcintAcabamento.BoundText & _
                          " AND IV.PKID = TP.intIssConstrucaoValorM2 " & _
                          " AND IE.PKID = IV.intIssConstrucaoExercicio " & _
                          " AND IE.intExercicio = " & Year(txtdtmConstrucao.Text) & _
                          " AND IT.PKID = TP.intIssConstrucaoTipo "
    
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoRec) Then
            If Not adoRec.EOF Then
            
                If Not IsNull(adoRec("DataAtual")) Then
                    dblIndexador = adoRec("DataAtual")
                ElseIf Not IsNull(adoRec("DataMes")) Then
                    dblIndexador = adoRec("DataMes")
                Else
                    txtdblVlrM2Servico.Text = Space$(0)
                    txtdblIssDevido.Text = Space$(0)
                    txtdblIssPagar.Text = Space$(0)
                    chkintDemolicao.Tag = ""
                
                    ExibeMensagem "Não foi encontrado indexador econômico."
                End If
                
                If chkintDemolicao.Value = vbChecked Then
                    txtdblVlrM2Servico.Text = TruncaValores(gstrConvVrDoSql((adoRec("dblValorM2").Value * dblIndexador) - ((adoRec("dblValorM2").Value * dblIndexador) * (adoRec("dblPorcDemolicao").Value / 100)), 4, , True), 2)
                    chkintDemolicao.Tag = gstrConvVrDoSql(adoRec("dblPorcDemolicao").Value, 4)
                Else
                    txtdblVlrM2Servico.Text = TruncaValores(gstrConvVrDoSql(adoRec("dblValorM2").Value * dblIndexador, 4, , True), 2)
                    chkintDemolicao.Tag = ""
                End If
                
                txtdblAliquota.Text = gstrConvVrDoSql(adoRec("dblAliquotaIss").Value, 2)
                txtdblVlrServico.Text = gstrConvVrDoSql(txtdblVlrM2Servico.Text * gstrConvVrDoSql(txtdblArea.Text, , , True), 2)
                txtdblIssDevido.Text = gstrConvVrDoSql((gstrConvVrDoSql(txtdblArea.Text, , , True) * gstrConvVrDoSql(txtdblVlrM2Servico.Text, , , True) * gstrConvVrDoSql((txtdblAliquota.Text / 100), , , True)), 2)
                
                If Len(txtdblIssAbatimento.Text) > 0 Then
                    txtdblIssPagar.Text = gstrConvVrDoSql(txtdblIssDevido.Text - txtdblIssAbatimento.Text, 2)
                Else
                    txtdblIssPagar.Text = gstrConvVrDoSql(txtdblIssDevido.Text, 2)
                End If
                
            Else
                txtdblVlrM2Servico.Text = Space$(0)
                txtdblIssDevido.Text = Space$(0)
                txtdblIssPagar.Text = Space$(0)
                chkintDemolicao.Tag = ""
                
                ExibeMensagem "Ano de construção sem valor cadastrado."
            End If
        End If
    End If
    
End Sub

Private Function VerificaProcesso() As Boolean
    
    'Caso esteja em branco vamos sair da rotina de verificacao
    If Trim(txtstrCodigoProcesso.Text) = "" And Trim(txtintExercicioProcesso.Text) = "" And Trim(txtbitDigitoProcesso.Text) = "" Then
         VerificaProcesso = True
         Exit Function
    End If
    
    'caso seja informado algum campo, vamos verificar se é valido
    If Trim(txtstrCodigoProcesso.Text) = "" Then
        ExibeMensagem "O Código do Processo deve ser informado."
        txtstrCodigoProcesso.SetFocus
        Exit Function
    ElseIf Trim(txtintExercicioProcesso.Text) = "" Then
        ExibeMensagem "O Exercício do Processo deve ser informado."
        txtintExercicioProcesso.SetFocus
        Exit Function
    ElseIf Trim(txtbitDigitoProcesso.Text) = "" Then
         ExibeMensagem "O Dígito do Processo deve ser informado."
         txtbitDigitoProcesso.SetFocus
         Exit Function
    End If
  
    If gblnExisteCodigo(2, gstrProtocolizacaoProcesso, "strCodigo", "'" & Trim(txtstrCodigoProcesso.Text) & "'", _
       "intExercicio", Trim(txtintExercicioProcesso.Text), "bitDigito", Trim(txtbitDigitoProcesso.Text)) = False Then
       ExibeMensagem "O Processo " & Trim(txtstrCodigoProcesso.Text) & "/" & _
                     Trim(txtintExercicioProcesso.Text) & "-" & Trim(txtbitDigitoProcesso.Text) & " não existe."
       txtstrCodigoProcesso.SetFocus
       Exit Function
    End If
  
    VerificaProcesso = True
    
End Function

Private Sub txtstrObservacoes_GotFocus()
    MarcaCampo txtstrObservacoes
End Sub

Private Sub txtstrObservacoes_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrObservacoes
End Sub

Private Function TruncaValores(strValor As String, bytCasasDecimais As Byte) As Double
Dim bytPos   As Byte

    bytPos = (Len(strValor) - InStr(strValor, ",")) - bytCasasDecimais
    
    TruncaValores = Mid(strValor, 1, Len(strValor) - bytPos)
    
End Function

Private Function strQueryDataComboContribuinte()
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strNome "
    strSQL = strSQL & "FROM " & gstrContribuinte & " "
    strSQL = strSQL & "ORDER BY strNome"
    
    strQueryDataComboContribuinte = strSQL
End Function

Private Function strQueryLogradouroContribuinte(strInscricao As String) As String
    Dim strSQL As String
    
    strSQL = strSQL & "Select "
    strSQL = strSQL & "CO.strLogradouroC, "
    strSQL = strSQL & "CO.intNumeroC, "
    strSQL = strSQL & "CO.strBairroC, "
    strSQL = strSQL & "CO.IntcepC, "
    strSQL = strSQL & "CO.strComplementoC, "
    strSQL = strSQL & "(SELECT MU.strDescricao FROM " & gstrCidade & " MU WHERE MU.PKId = IM.intMunicipioC) strMunicipioC, "
    strSQL = strSQL & "(SELECT UF.strSigla FROM " & gstrUF & " UF WHERE UF.PKId = IM.intUFC) strEstadoC "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrImobiliario & " IM, "
    strSQL = strSQL & gstrContribuinte & " CO, "
    strSQL = strSQL & gstrBairro & " BA, "
    strSQL = strSQL & gstrLogradouro & " L, "
    strSQL = strSQL & gstrTituloLogradouro & " U, "
    strSQL = strSQL & gstrTipoLogradouro & " TL "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " Co.pkid = IM.intContribuinte "
    strSQL = strSQL & " AND L.pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " IM.Intlogradouro "
    strSQL = strSQL & " AND L.intBairro = BA.PKId"
    strSQL = strSQL & " AND L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle
    'strSql = strSql & " AND L.DtmdtExclusao is null "
    strSQL = strSQL & " AND L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle
    strSQL = strSQL & " AND IM.Strinscricao = " & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & UCase(strInscricao)
    
    strQueryLogradouroContribuinte = strSQL
End Function

Private Sub CriaCriticaDeIptu(strComposicao As String, strInscricao As String, intExercicio As Integer, strOcorrencia As String, Optional STRCOMPLEMENTO As String, Optional strEmissao As String)
Dim strSQL As String
                        
    strSQL = ""
    strSQL = strSQL & "INSERT INTO " & gstrCriticaIptu & " (strcomposicao, strinscricao, intexercicio, strocoreencia, strcomplemento, strEmissao, dtmdtatualizacao, lngcodusr)"
    strSQL = strSQL & "Values  ('"
    strSQL = strSQL & strComposicao & "', '"
    strSQL = strSQL & strInscricao & "', "
    strSQL = strSQL & intExercicio & ", '"
    strSQL = strSQL & strOcorrencia & "', '"
    strSQL = strSQL & STRCOMPLEMENTO & "', '"
    strSQL = strSQL & strEmissao & "', "
    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema(True)) & ", "
    strSQL = strSQL & glngCodUsr & ")"
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    gobjBanco.Execute strSQL, False
    gobjBanco.ExecutaCommitTrans

End Sub

Private Function VerificaValorMinimoPorParcela(dblValorTotal As Double, intQtdeParcelas As Integer) As Integer
Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
    
    VerificaValorMinimoPorParcela = intQtdeParcelas
    
    strSQL = "Select " & gstrISNULL("dblParcelaMinima", "0") & " as dblParcelaMinima From " & gstrComposicaoDaReceita & " Where pkid = " & dbc_intComposicao.BoundText

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            
            If Val(adoResultado!dblParcelaMinima) > 0 Then
            
                If (dblValorTotal / intQtdeParcelas) < adoResultado!dblParcelaMinima Then
                    VerificaValorMinimoPorParcela = Int(dblValorTotal / adoResultado!dblParcelaMinima)
                End If
            
            End If
            
        End If
    End If
    
End Function

