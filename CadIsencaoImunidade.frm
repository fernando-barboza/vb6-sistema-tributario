VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadIsencaoImunidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Isenções e Imunidades"
   ClientHeight    =   7890
   ClientLeft      =   2070
   ClientTop       =   2055
   ClientWidth     =   8565
   HelpContextID   =   9
   Icon            =   "CadIsencaoImunidade.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   8565
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5850
      Left            =   60
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   45
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10319
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Isenções e Imunidades"
      TabPicture(0)   =   "CadIsencaoImunidade.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrInscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintComposicaoDaReceita"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintContribuinte"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintIdentificacao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintComposicaoDaReceita"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_Inscricao"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_Isencao"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_strContribuinte"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_HistoricoFaceDeQuadra"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_strPromissario"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Período"
      TabPicture(1)   =   "CadIsencaoImunidade.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvw_Periodo"
      Tab(1).Control(1)=   "fra_Periodo"
      Tab(1).Control(2)=   "txtPKId"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox txt_strPromissario 
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
         Left            =   2605
         MaxLength       =   100
         TabIndex        =   42
         Top             =   1755
         Width           =   4935
      End
      Begin VB.TextBox txtPKId 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71730
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Frame fra_Periodo 
         Caption         =   "Período"
         Height          =   2565
         Left            =   -74160
         TabIndex        =   36
         Top             =   420
         Width           =   6525
         Begin VB.TextBox txt_strCodigoProcesso 
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
            Left            =   1440
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   675
            Width           =   960
         End
         Begin VB.TextBox txt_intExercicioProcesso 
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
            Left            =   2430
            MaxLength       =   4
            TabIndex        =   18
            Top             =   675
            Width           =   465
         End
         Begin VB.TextBox txt_bitDigitoProcesso 
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
            Left            =   2895
            MaxLength       =   1
            TabIndex        =   19
            Top             =   675
            Width           =   285
         End
         Begin VB.OptionButton opt_bytPosicao 
            Caption         =   "Em Andamento"
            Height          =   195
            Index           =   2
            Left            =   405
            TabIndex        =   21
            Top             =   1125
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.OptionButton opt_bytPosicao 
            Caption         =   "Indeferido"
            Height          =   195
            Index           =   1
            Left            =   4950
            TabIndex        =   23
            Top             =   1125
            Width           =   1425
         End
         Begin VB.OptionButton opt_bytPosicao 
            Caption         =   "Deferido"
            Height          =   195
            Index           =   0
            Left            =   2835
            TabIndex        =   22
            Top             =   1125
            Width           =   1425
         End
         Begin VB.CheckBox chk_bytCancelado 
            Caption         =   "Cancelado"
            Height          =   195
            Left            =   3455
            TabIndex        =   20
            Top             =   735
            Width           =   1185
         End
         Begin VB.TextBox txt_strobservacao 
            Height          =   885
            Left            =   90
            MaxLength       =   400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   1560
            Width           =   6345
         End
         Begin VB.TextBox txt_dtmData 
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
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   14
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txt_dtmFinal 
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
            Left            =   5385
            MaxLength       =   10
            TabIndex        =   16
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txt_dtmInicial 
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
            Left            =   3450
            MaxLength       =   10
            TabIndex        =   15
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Processo"
            Height          =   195
            Left            =   135
            TabIndex        =   41
            Top             =   675
            Width           =   660
         End
         Begin VB.Label lbl_Data 
            AutoSize        =   -1  'True
            Caption         =   "Data da Inclusão"
            Height          =   195
            Left            =   135
            TabIndex        =   39
            Top             =   345
            Width           =   1215
         End
         Begin VB.Label lbldtmFinal 
            AutoSize        =   -1  'True
            Caption         =   "Data Final"
            Height          =   195
            Left            =   4575
            TabIndex        =   38
            Top             =   345
            Width           =   720
         End
         Begin VB.Label lbldtmInicial 
            AutoSize        =   -1  'True
            Caption         =   "Data Inicial"
            Height          =   195
            Left            =   2565
            TabIndex        =   37
            Top             =   345
            Width           =   795
         End
      End
      Begin VB.Frame fra_HistoricoFaceDeQuadra 
         Caption         =   "Receitas"
         Height          =   2205
         Left            =   900
         TabIndex        =   32
         Top             =   3435
         Width           =   6750
         Begin VB.TextBox txt_dblAliquota 
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
            Left            =   5100
            MaxLength       =   20
            TabIndex        =   11
            Top             =   240
            Width           =   1395
         End
         Begin MSComctlLib.ListView lvw_Receita 
            Height          =   1485
            Left            =   240
            TabIndex        =   12
            Top             =   630
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   2619
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Pkid"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Receita"
               Object.Width           =   8731
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Alíquota"
               Object.Width           =   2205
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "IntReceita"
               Object.Width           =   0
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbc_intReceita 
            Height          =   315
            HelpContextID   =   1
            Left            =   900
            TabIndex        =   10
            Top             =   240
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Alíquota"
            Height          =   195
            Left            =   4425
            TabIndex        =   35
            Top             =   330
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Receita"
            Height          =   195
            Left            =   285
            TabIndex        =   34
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.TextBox txt_strContribuinte 
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
         Left            =   2610
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1395
         Width           =   4935
      End
      Begin VB.Frame fra_Isencao 
         Caption         =   "Definição"
         Height          =   930
         Left            =   885
         TabIndex        =   27
         Top             =   2475
         Width           =   6765
         Begin VB.OptionButton optbitDefinicao 
            Caption         =   "Imunidade"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   0
            Left            =   615
            TabIndex        =   5
            Top             =   210
            Width           =   1185
         End
         Begin VB.OptionButton optbitDefinicao 
            Caption         =   "Isenção"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   1
            Left            =   3060
            TabIndex        =   6
            Top             =   210
            Width           =   945
         End
         Begin VB.OptionButton optbitDefinicao 
            Caption         =   "Não Incidência"
            CausesValidation=   0   'False
            Height          =   225
            Index           =   2
            Left            =   5190
            TabIndex        =   7
            Top             =   210
            Width           =   1425
         End
         Begin VB.CommandButton cmd_TipoIsencaoImunidade 
            Height          =   315
            Left            =   6120
            Picture         =   "CadIsencaoImunidade.frx":107A
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Tag             =   "585"
            ToolTipText     =   "Ativa Tipo Isenção Imunidade"
            Top             =   465
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dbcintTipoIsencaoImunidade 
            Height          =   315
            Left            =   2460
            TabIndex        =   8
            Top             =   465
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Text            =   ""
         End
         Begin VB.Label lblintTipoIsencaoImunidade 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo Isenção Imunidade"
            Height          =   195
            Left            =   660
            TabIndex        =   33
            Top             =   555
            Width           =   1710
         End
      End
      Begin VB.Frame fra_Inscricao 
         Height          =   585
         Left            =   120
         TabIndex        =   26
         Top             =   330
         Width           =   8175
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Econômico"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   1
            Left            =   4905
            TabIndex        =   1
            Top             =   270
            Width           =   1575
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   0
            Left            =   2055
            TabIndex        =   0
            Top             =   270
            Value           =   -1  'True
            Width           =   1605
         End
      End
      Begin MSDataListLib.DataCombo dbcintComposicaoDaReceita 
         Height          =   315
         Left            =   2610
         TabIndex        =   4
         Top             =   2115
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintIdentificacao 
         Height          =   315
         HelpContextID   =   1
         Left            =   2610
         TabIndex        =   2
         Top             =   1020
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSComctlLib.ListView lvw_Periodo 
         Height          =   1695
         Left            =   -74940
         TabIndex        =   25
         Top             =   3330
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   2990
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pkid"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Data"
            Object.Width           =   2170
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Data Inicial"
            Object.Width           =   2170
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Data Final"
            Object.Width           =   2170
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Posição"
            Object.Width           =   4099
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Processo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cancelamento"
            Object.Width           =   3836
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Observação"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "BytCancelamento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "bytPosicao"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Promissário"
         Height          =   195
         Left            =   1620
         TabIndex        =   43
         Top             =   1845
         Width           =   795
      End
      Begin VB.Label lblintContribuinte 
         AutoSize        =   -1  'True
         Caption         =   "Contribuinte"
         Height          =   195
         Left            =   1635
         TabIndex        =   31
         Top             =   1485
         Width           =   840
      End
      Begin VB.Label lblintComposicaoDaReceita 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   780
         TabIndex        =   30
         Top             =   2190
         Width           =   1695
      End
      Begin VB.Label lblstrInscricao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   1125
         TabIndex        =   29
         Top             =   1140
         Width           =   1350
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Receita 
      Height          =   1845
      Left            =   30
      TabIndex        =   13
      Top             =   5985
      Width           =   8415
      _ExtentX        =   14843
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
      Columns(1).Caption=   "Inscrição / Contribuinte "
      Columns(1).DataField=   "strinscricao"
      Columns(1).NumberFormat=   "FormatText Event"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Contribuinte"
      Columns(2).DataField=   "StrNome"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Promissário"
      Columns(3).DataField=   "strPromissario"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Composição da Receita"
      Columns(4).DataField=   "StrComposicao"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "intComposicaoDaReceita"
      Columns(5).DataField=   "intComposicaoDaReceita"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Utilizacao"
      Columns(6).DataField=   "intUtilizacao"
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
      Splits(0)._ColumnProps(8)=   "Column(1).Width=3069"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2990"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=5662"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=5583"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=5371"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=5292"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=5821"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=5741"
      Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(33)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(34)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(41)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=164,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadIsencaoImunidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intPkid              As Long
Public strFormulario        As String
Public mblnPrimeiraVez      As Boolean
Dim mblnClickOk             As Boolean
Dim mblnAlterando           As Boolean
Dim mobjAux                 As Object
Dim mblnSelecionou          As Boolean
Dim mobjLista               As Variant
Dim blnAlteraReceita        As Boolean
Dim blnAlteraPeriodo        As Boolean

Private Sub cmd_TipoIsencaoImunidade_Click()
    CarregaForm frmTipoIsencaoImunidade, dbcintTipoIsencaoImunidade
End Sub

Private Sub dbcintComposicaoDaReceita_Click(Area As Integer)
'    DropDownDataCombo dbcintComposicaoDaReceita, Me, Area
    
'    If Area = 2 Then
'        dbc_intReceita.Tag = strQueryReceita & ";strDescricao"
'        PreencherListaDeOpcoes dbc_intReceita
'    Else
'        Set dbc_intReceita.RowSource = Nothing
'        dbc_intReceita.Text = ""
'    End If
    
End Sub

Private Sub dbcintComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintComposicaoDaReceita, Me, , KeyCode, Shift
    
    If dbcintComposicaoDaReceita.MatchedWithList Then
        dbc_intReceita.Tag = strQueryReceita & ";strDescricao"
        Set dbc_intReceita.RowSource = Nothing
        dbc_intReceita.Text = ""
    End If
End Sub

Private Sub dbcintComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintComposicaoDaReceita, True
End Sub

Private Sub dbcintComposicaoDaReceita_LostFocus()
'    DropDownDataCombo dbcintComposicaoDaReceita, Me
'    If dbcintComposicaoDaReceita.MatchedWithList Then
'        dbc_intReceita.Tag = strQueryReceita & ";strDescricao"
'        LeDaTabelaParaObj "", dbc_intReceita, strQueryReceita
'        Set dbc_intReceita.RowSource = Nothing
'        dbc_intReceita.Text = ""
'    End If
End Sub

Private Sub dbcintIdentificacao_Change()
    If dbcintIdentificacao.MatchedWithList Then
        PreencheContribuinte
        If optbitTipoDeInscricao(0).Value = True Then
           PreenchePromissario
        End If
    End If
End Sub

Private Sub dbcintIdentificacao_Click(Area As Integer)
    DropDownDataCombo dbcintIdentificacao, Me, Area
    If dbcintIdentificacao.MatchedWithList Then
        PreencheContribuinte
        If optbitTipoDeInscricao(0).Value = True Then
           PreenchePromissario
        End If
    End If
End Sub

Private Sub dbcintIdentificacao_GotFocus()
    MarcaCampo dbcintIdentificacao
End Sub


Private Sub dbcintIdentificacao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintIdentificacao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintIdentificacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintIdentificacao, True
End Sub

Private Sub dbcintIdentificacao_LostFocus()
    If Trim(dbcintIdentificacao.Text) <> "" Then
       VerificaInscricao
    End If
End Sub

Private Sub dbcintTipoIsencaoImunidade_GotFocus()
    Dim adoResultado As ADODB.Recordset
    Dim strSql As String
    
    If dbcintTipoIsencaoImunidade.MatchedWithList Then
       strSql = ""
       strSql = strSql & "Select Im.Inttipo "
       strSql = strSql & "from tbltipoisencaoimunidade IM "
       strSql = strSql & "Where IM.STRDESCRICAO like '" & dbcintTipoIsencaoImunidade.Text & "'"
       
       Set gobjBanco = New clsBanco
       If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
             If adoResultado.EOF = False Then
                With adoResultado
                   If (!intTipo = 0) Then
                       optbitDefinicao(0).Value = True
                   ElseIf (!intTipo = 1) Then
                       optbitDefinicao(1).Value = True
                   ElseIf (!intTipo = 2) Then
                       optbitDefinicao(2).Value = True
                   End If
                End With
              Else
                ExibeMensagem "Os dados referentes a opcão escolhida não foram encontrados!", vbExclamation
              End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 629
    VirificaGradeListView Me
    
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    
    If mblnSelecionou Then
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    
    If mblnAlterando Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar
    End If
    
    If mobjAux Is Nothing Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
    If tab_3dPasta.Tab = 1 Then
       txt_dtmInicial.SetFocus
    End If
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()

    mblnAlterando = False
    
    VerificaObjParaAplicar mobjAux
    optbitTipoDeInscricao_Click 0
    optbitDefinicao(0).Value = True
    TrocaCorObjeto txt_strContribuinte, True
    TrocaCorObjeto txt_strPromissario, True
    dbcintComposicaoDaReceita.Tag = strQuerryComposicao(0) & ";strDescricao"
    dbcintTipoIsencaoImunidade.Tag = strQueryTipoIsencaoImunidade & ";strDescricao"
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    tab_3dPasta.TabEnabled(1) = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
    If UCase(strFormulario) = "FRMCADIMOBILIARIO" Then
        frmCadImobiliario.PreencheCkIsencao
    End If
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
End Sub

Private Sub lvw_Periodo_Click()
Dim bytPos As Byte
    If lvw_Periodo.ListItems.Count >= 1 Then
        blnAlteraPeriodo = True
    End If
    If lvw_Periodo.ListItems.Count >= 1 Then
        With lvw_Periodo
            txt_dtmData.Text = .SelectedItem.SubItems(1)
            txt_dtmInicial.Text = .SelectedItem.SubItems(2)
            txt_dtmFinal.Text = .SelectedItem.SubItems(3)
            'PROCESSO
            If Trim(.SelectedItem.SubItems(5)) <> "" Then
               bytPos = InStr(1, .SelectedItem.SubItems(5), "-")
               txt_bitDigitoProcesso.Text = Mid(.SelectedItem.SubItems(5), bytPos + 1, 2)
               txt_intExercicioProcesso.Text = Mid(.SelectedItem.SubItems(5), bytPos - 4, 4)
               txt_strCodigoProcesso.Text = Mid(.SelectedItem.SubItems(5), 1, (bytPos - 6))
            Else
               txt_bitDigitoProcesso.Text = ""
               txt_intExercicioProcesso.Text = ""
               txt_strCodigoProcesso.Text = ""
            End If
            
            txt_strobservacao.Text = .SelectedItem.SubItems(7)
            chk_bytCancelado.Value = .SelectedItem.SubItems(8)
            opt_bytPosicao(.SelectedItem.SubItems(9)).Value = True
        End With
    End If
End Sub

Private Sub lvw_Receita_Click()
    If lvw_Receita.ListItems.Count >= 1 Then
        blnAlteraReceita = True
    End If
End Sub

'Private Sub optbitDefinicao_LostFocus(Index As Integer)
'    dbcintTipoIsencaoImunidade.ListField = ""
'    dbcintTipoIsencaoImunidade.Text = ""
'    dbcintTipoIsencaoImunidade.Tag = strQueryTipoIsencaoImunidade & ";strDescricao"
'End Sub

 Private Sub optbitDefinicao_Click(Index As Integer)
    dbcintTipoIsencaoImunidade.ListField = ""
    dbcintTipoIsencaoImunidade.Text = ""
    dbcintTipoIsencaoImunidade.Tag = strQueryTipoIsencaoImunidade & ";strDescricao"
End Sub

Private Sub optbitTipoDeInscricao_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", optbitTipoDeInscricao, True
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    Select Case tab_3dPasta.Tab
        Case 0
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        Case 1
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
            chk_bytCancelado.Value = 0
            opt_bytPosicao(0).Value = True
            LimpaControlesPeriodo
    End Select
    
    If mblnAlterando Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    End If
    
End Sub

Private Sub tdb_Receita_Click()
    mblnPrimeiraVez = True
End Sub

Sub tdb_Receita_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Receita_FilterChange()
    gblnFilraCampos tdb_Receita
    mblnPrimeiraVez = False
    mblnClickOk = False
End Sub

Private Sub tdb_Receita_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
Dim strInscricao As String

    Select Case ColIndex
        Case 1
            strInscricao = Value
            Value = gstrFormataInscricao(strInscricao, tdb_Receita.Columns("intUtilizacao").Value)
    End Select
End Sub

Private Sub tdb_Receita_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Receita, ColIndex
End Sub

Private Sub tdb_Receita_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown
            mblnClickOk = True
        Case Else
            mblnClickOk = False
    End Select
    
End Sub

Private Sub tdb_Receita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", tdb_Receita, True
End Sub

Private Sub tdb_Receita_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Receita_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strSql As String
    With tdb_Receita
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
               If mblnClickOk Then
                   tab_3dPasta.TabEnabled(1) = True
                   mblnAlterando = True
                   mblnClickOk = False
                   txtPKId.Text = .Columns("PKID").Value
                   LeDaTabelaParaObj gstrIsencaoImunidade, Me
                   PreencheContribuinte
                   If optbitTipoDeInscricao(0).Value = True Then
                      PreenchePromissario
                   End If
                   dbcintComposicaoDaReceita.BoundText = .Columns("intComposicaoDaReceita").Value
                   gCorLinhaSelecionada tdb_Receita
                   PreencheReceitas CLng(txtPKId.Text)
                   PreenchePeriodo CLng(txtPKId.Text)
                   LimpaControlesPeriodo
                   HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                   If mobjAux Is Nothing Then
                       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                   Else
                       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                   End If
                   'DesabilitaIsencao
                   abilitaIsencao
                   HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
                   mblnSelecionou = True
               End If
            End If
        End If
    End With
End Sub

Private Sub PreenchePromissario()
Dim strSql As String
Dim adoResultado As ADODB.Recordset

  'Antes da rotina ser chamada, é verificado se a inscrição é imobiliária
  strSql = ""
  strSql = strSql & "SELECT CT.strNome "
  strSql = strSql & "FROM "
  
  strSql = strSql & gstrContribuinte & " CT, "
  strSql = strSql & gstrImobiliario & " IM "
  strSql = strSql & "WHERE "
  strSql = strSql & "CT.pkID = IM.intPromissario AND "
  strSql = strSql & "IM.pkID = " & dbcintIdentificacao.BoundText
  
  Set gobjBanco = New clsBanco
  If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
     If Not adoResultado.EOF Then
        txt_strPromissario.Text = gstrENulo(adoResultado!strNome)
     Else
        txt_strPromissario.Text = ""
     End If
  End If

End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSql          As String
    Dim strInscricao    As String
    Dim intIndice       As Integer
    Dim blnAlterando    As Boolean
    Dim lngPkidIsencao  As Long
    
    lngPkidIsencao = Val(txtPKId)
    
    
    If UCase(strModoOperacao) = UCase(gstrIncluirItem) Then
        If tab_3dPasta.Tab = 0 Then
            If Not mblnAlterando Then
                IncluirItemNoGrid
            End If
        ElseIf tab_3dPasta.Tab = 1 Then             'Incluir Item
            If VerificaProcesso = True Then
               IncluirItemNoGrid
            End If
        End If
    
    ElseIf UCase(strModoOperacao) = UCase(gstrExcluirItem) Then
        If tab_3dPasta.Tab = 0 Then
            If Not mblnAlterando Then
                ExcluirItemNoGrid
            End If
        ElseIf tab_3dPasta.Tab = 1 Then             'Excluir item
            If mblnAlterando Then
                If lvw_Periodo.ListItems.Count > 1 Then
                    If gblnExclusaoGravacaoOk("SALVAR", "Deseja realmente excluir o período ") Then
                        ExcluirItemNoGrid
                        LimpaPeriodo
                        GravaPeriodo CLng(txtPKId)
                        PreenchePeriodo CLng(txtPKId)
                        txt_dtmData.SetFocus
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                    ExibeMensagem "Não pode ser excluídos todos os períodos."
                    Exit Sub
                End If
            End If
            ExcluirItemNoGrid
        End If
    End If
    
    If UCase(strModoOperacao) = gstrPreencherLista Or UCase(strModoOperacao) = gstrLocalizar Or UCase(strModoOperacao) = gstrRefresh Then
        mblnPrimeiraVez = True
        mblnClickOk = True
        
        For intIndice = 0 To 1
            If optbitTipoDeInscricao(intIndice).Value Then
                strSql = strQuery(intIndice)
                Exit For
            End If
        Next
'        If mblnPrimeiraVez = False Then
            If Me.ActiveControl.Name = "dbcintIdentificacao" Or Me.ActiveControl.Name = "optbitTipoDeInscricao" Then
                If UCase(strModoOperacao) = gstrPreencherLista Then
'                    dbcintComposicaoDaReceita.Text = ""
'                    dbcintComposicaoDaReceita.ListField = ""
'                    dbc_intReceita.Text = ""
'                    dbc_intReceita.ListField = ""
'                    dbcintTipoIsencaoImunidade.Text = ""
'                    dbcintTipoIsencaoImunidade.ListField = ""
'                    optbitDefinicao(0).Value = True
'                    dbc_intReceita.Tag = ""
'                    dbcintTipoIsencaoImunidade.Tag = ""
'                    dbcintComposicaoDaReceita.Tag = ""
'                    lvw_Receita.ListItems.Clear
                    Select Case optbitTipoDeInscricao(0).Value
                        Case True
                            dbcintIdentificacao.Tag = strQueryImobiliario(intIndice) & "Order by strinscricao;strinscricao"
                        Case Else
                            dbcintIdentificacao.Tag = strQueryImobiliario(intIndice) & "Order by strInscricaoCadastral;Strinscricaocadastral"
                    End Select
'                    If Trim(Me.ActiveControl.Text) <> "" Then
'                         dbcintIdentificacao.Tag = gstrQueryIdentificacao(intIndice, True) & "; strintinscricao"
'                         LeDaTabelaParaObj "", dbcintIdentificacao, gstrQueryIdentificacao(intIndice)
'                    Else
                         ToolBarGeral strModoOperacao, gstrIsencaoImunidade, mblnAlterando, tdb_Receita, Me, mobjAux, strSql
'                    End If
                    Exit Sub
                End If
            ElseIf Me.ActiveControl.Name = "dbcintTipoIsencaoImunidade" Then
'                    dbc_intReceita.Tag = ""
'                    dbcintIdentificacao.Tag = ""
'                    dbcintComposicaoDaReceita.Tag = ""
                ToolBarGeral strModoOperacao, gstrIsencaoImunidade, mblnAlterando, tdb_Receita, Me, mobjAux, strSql
            ElseIf Me.ActiveControl.Name = "dbcintComposicaoDaReceita" Then
'                    dbc_intReceita.Tag = ""
'                    dbcintTipoIsencaoImunidade.Tag = ""
'                    dbcintIdentificacao.Tag = ""
                If Me.ActiveControl.Text = "" Then
                    dbcintComposicaoDaReceita.Tag = strQuerryComposicao(intIndice) & ";strDescricao"
                    LeDaTabelaParaObj "", dbcintComposicaoDaReceita, strQuerryComposicao(intIndice)
                Else
                    dbcintComposicaoDaReceita.Tag = strQuerryComposicao(intIndice, True) & ";strDescricao"
                    LeDaTabelaParaObj "", dbcintComposicaoDaReceita, strQuerryComposicao(intIndice, True)
                End If
            
            ElseIf Me.ActiveControl.Name = "dbc_intReceita" Then
                If dbcintComposicaoDaReceita.MatchedWithList Then
                     dbc_intReceita.Tag = strQueryReceita & ";strDescricao"
                     PreencherListaDeOpcoes dbc_intReceita
                End If
            End If
'         End If
            
        If UCase(strModoOperacao) = gstrLocalizar Then
            dbcintComposicaoDaReceita.Tag = strQuerryComposicao(0) & ";strDescricao"
'            dbcintComposicaoDaReceita.Text = ""
'            dbcintComposicaoDaReceita.Tag = ""
'            dbc_intReceita.Text = ""
'            dbc_intReceita.ListField = ""
'            dbcintTipoIsencaoImunidade.Text = ""
'            dbcintTipoIsencaoImunidade.ListField = ""
'             lvw_Receita.ListItems.Clear
'             lvw_Periodo.ListItems.Clear
'            optbitDefinicao(0).Value = True
'            dbcintTipoIsencaoImunidade.Tag = ""
'            dbc_intReceita.Tag = ""
'            dbcintComposicaoDaReceita.Tag = ""
            If intIndice = 0 Then
                LeDaTabelaParaObj gstrImobiliario, tdb_Receita, strSql
            ElseIf intIndice = 1 Then
                LeDaTabelaParaObj gstrEconomico, tdb_Receita, strSql
            End If
        End If
               'ToolBarGeral strModoOperacao, gstrIsencaoImunidade, mblnAlterando, tdb_Receita, Me, mobjAux, strSql
        Exit Sub
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        If tab_3dPasta.Tab = 0 Then
            ToolBarGeral strModoOperacao, gstrIsencaoImunidade, mblnAlterando, tdb_Receita, Me, mobjAux, strSql
            LimpaCombos
            abilitaIsencao
            dbcintIdentificacao.SetFocus
            lvw_Periodo.ListItems.Clear
            lvw_Receita.ListItems.Clear
            dbc_intReceita.Text = ""
            dbcintComposicaoDaReceita.Text = ""
            dbcintComposicaoDaReceita.ListField = ""
            LimpaControlesPeriodo
            optbitDefinicao(0).Value = True
            optbitDefinicao(0).SetFocus
            optbitTipoDeInscricao(0).Value = True
            optbitTipoDeInscricao(0).SetFocus
            blnAlteraPeriodo = False
            blnAlteraReceita = False
            mblnPrimeiraVez = False
            mblnClickOk = False
            tab_3dPasta.Tab = 0
            tab_3dPasta.TabEnabled(1) = False
            Exit Sub
        Else            'NOVO ÍTEM
            NovoPeriodo
            txt_dtmInicial.SetFocus
        End If
            
    End If
    
    If UCase(strModoOperacao) <> "NOVO" Then
        For intIndice = 0 To 1
            If optbitTipoDeInscricao(intIndice).Value Then
                strSql = strQuery(intIndice)
                Exit For
            End If
        Next
    End If
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
       mblnPrimeiraVez = False
    End If
    
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        ImprimeRelatorio rptIsencaoImunidade, strQueryRelatorio
    End If
    
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Or UCase(strModoOperacao) = UCase(gstrDeletar) Then
    
        blnAlterando = mblnAlterando
        
        If UCase(strModoOperacao) = UCase(gstrSalvar) Then
            If blnDadosOk(strModoOperacao) Then
                If Not mblnAlterando Then
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaBeginTrans
                    If ToolBarGeral(strModoOperacao, gstrIsencaoImunidade, mblnAlterando, tdb_Receita, Me, mobjAux, strSql) = True Then
                        If Not blnAlterando Then
                            lngPkidIsencao = glngPegaUltimaChave(gstrIsencaoImunidade, "PKId")
                        End If
                        If GravaReceitas(lngPkidIsencao) Then
                            gobjBanco.ExecutaCommitTrans
                            LimpaCombos
                            lvw_Periodo.ListItems.Clear
                            lvw_Receita.ListItems.Clear
                            dbcintIdentificacao.SetFocus
                            HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
                            dbc_intReceita.Tag = ""
                            LimpaPeriodo
                            mblnPrimeiraVez = False
                            tab_3dPasta.Tab = 0
                            DoEvents
                            For intIndice = 0 To 1
                                If optbitTipoDeInscricao(intIndice).Value Then
                                    LeDaTabelaParaObj gstrIsencaoImunidade, tdb_Receita, strQuery(intIndice, lngPkidIsencao)
                                    Exit For
                                End If
                            Next
                        Else
                            gobjBanco.ExecutaRollbackTrans
                        End If
                    Else
                        gobjBanco.ExecutaRollbackTrans
                    End If
                End If
            End If
        Else
            If ExcluirIsencao = True Then
                Limpa_Controles Me, True, True, True, True, True
                LimpaCombos
                abilitaIsencao
                LimpaPeriodo
                lvw_Periodo.ListItems.Clear
                lvw_Receita.ListItems.Clear
                dbcintIdentificacao.SetFocus
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
                dbc_intReceita.Tag = ""
                mblnPrimeiraVez = False
                optbitDefinicao.Item(0).Value = True
                optbitTipoDeInscricao.Item(0).Value = True
                tab_3dPasta.Tab = 0
                tab_3dPasta.TabEnabled(1) = False
           End If
        End If
    End If
End Sub

Private Function gstrQueryIdentificacao(bytsistema As Integer, Optional blntag As Boolean) As String
    Dim strSql As String
    
    strSql = "select IM.intIdentificacao, "
    strSql = strSql & strSUBSTRING & "(" & gstrRIGHT("A.strInscricao", gintRetornaTamanhoMascara(IIf(bytsistema = 0, TYP_IMOBILIARIA, TYP_ECONOMICA))) & ",1,20) strinscricao FROM tblIsencaoImunidade IM,  tblContribuinte CO, tblComposicaoDaReceita CR, "
    strSql = strSql & "(select a.pkid, a.strinscricao , a.intcontribuinte from tblImobiliario A union select a.pkid, a.strinscricao, a.intPromissario from tblImobiliario A) A "
    strSql = strSql & "where A.PKId = IM.intIdentificacao  AND CO.PKId = A.intContribuinte AND CR.pkid = IM.intComposicaoDaReceita "

    If blntag = False Then
        strSql = strSql & "AND A.strInscricao LIKE " & "'" & UCase(String(gintLenInscricao - gintRetornaTamanhoMascara(IIf(bytsistema = 0, TYP_IMOBILIARIA, TYP_ECONOMICA)), "0") & dbcintIdentificacao.Text) & "%' "
    End If

    strSql = strSql & "AND IM.bitTipoDeInscricao = " & bytsistema & " order by " & gstrCONVERT(cdt_numeric, "strinscricao")
    
    gstrQueryIdentificacao = strSql

End Function

    
Private Function VerificaProcesso() As Boolean
  If Not gblnDataValida(txt_dtmData.Text) Then
     ExibeMensagem "A data da inclusão informada não é válida."
     txt_dtmData.SetFocus
     Exit Function
  ElseIf Not gblnDataValida(txt_dtmInicial.Text) Then
       ExibeMensagem "A data inicial informada não é válida."
       txt_dtmInicial.SetFocus
       Exit Function
  ElseIf Not gblnDataValida(txt_dtmFinal.Text) Then
       ExibeMensagem "A data final informada não é válida."
       txt_dtmFinal.SetFocus
       Exit Function
  ElseIf Trim(txt_strCodigoProcesso.Text) = "" Then
     ExibeMensagem "O código do processo deve ser informado."
     txt_strCodigoProcesso.SetFocus
     Exit Function
  ElseIf Trim(txt_intExercicioProcesso.Text) = "" Then
         ExibeMensagem "O exercício do processo deve ser informado."
         txt_intExercicioProcesso.SetFocus
         Exit Function
      ElseIf Trim(txt_bitDigitoProcesso.Text) = "" Then
             ExibeMensagem "O dígito do processo deve ser informado."
             txt_bitDigitoProcesso.SetFocus
             Exit Function
      End If
  
  If gblnExisteCodigo(2, gstrProtocolizacaoProcesso, "strCodigo", "'" & Trim(txt_strCodigoProcesso.Text) & "'", _
     "intExercicio", Trim(txt_intExercicioProcesso.Text), "bitDigito", Trim(txt_bitDigitoProcesso.Text)) = False Then
     ExibeMensagem "O processo " & Trim(txt_strCodigoProcesso.Text) & "/" & _
                   Trim(txt_intExercicioProcesso.Text) & "-" & Trim(txt_bitDigitoProcesso.Text) & " não existe."
     txt_strCodigoProcesso.SetFocus
     Exit Function
  End If
  
  VerificaProcesso = True
End Function

Private Sub UltimoProcesso()
    Dim strSql As String
    Dim adoProcesso As ADODB.Recordset
    Dim intFor As Integer
        
    Dim strProcesso As String
    Dim strDigito As String
    Dim strExercicio As String
    Dim strCodigo As String
    Dim bytPos As Integer
    
    If lvw_Periodo.ListItems.Count >= 1 Then
        For intFor = 1 To lvw_Periodo.ListItems.Count
            With lvw_Periodo
                If Trim(.ListItems(intFor).SubItems(5)) <> "" Then
                    bytPos = InStr(1, .ListItems(intFor).SubItems(5), "-")
                    strProcesso = Trim(.ListItems(intFor).SubItems(5))
                    strDigito = Mid(strProcesso, bytPos + 1, 2)
                    strExercicio = Mid(strProcesso, bytPos - 4, 4)
                    strCodigo = Mid(strProcesso, 1, (bytPos - 6))
                    Exit For
                End If
            End With
        Next
    End If
    
    
    If strProcesso <> "" Then
        For intFor = 1 To lvw_Periodo.ListItems.Count
            With lvw_Periodo
                If Trim(.ListItems(intFor).SubItems(5)) > strProcesso Then
                    bytPos = InStr(1, .ListItems(intFor).SubItems(5), "-")
                    strProcesso = Trim(.ListItems(intFor).SubItems(5))
                    
                    strDigito = Mid(.ListItems(intFor).SubItems(5), bytPos + 1, 2)
                    strExercicio = Mid(.ListItems(intFor).SubItems(5), bytPos - 4, 4)
                    strCodigo = Mid(.ListItems(intFor).SubItems(5), 1, (bytPos - 6))
                End If
            End With
        Next
    End If
    
    txt_strCodigoProcesso.Text = strCodigo
    txt_intExercicioProcesso.Text = strExercicio
    txt_bitDigitoProcesso.Text = strDigito
    
  
End Sub



Private Function strQuery(intTipoDeInscricao As Integer, Optional lngPkid As Long = 0) As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT IM.Pkid, "
    Select Case intTipoDeInscricao
        Case 0
            strSql = strSql & gstrRIGHT("A.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
        Case 1
            strSql = strSql & gstrRIGHT("A.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
    End Select
    strSql = strSql & "CO.StrNome, "
    
    If intTipoDeInscricao = 0 Then
       strSql = strSql & "CO1.strNome strPromissario, "
    End If
    
    strSql = strSql & "CR.strDescricao as StrComposicao, "
    strSql = strSql & "CR.intUtilizacao, "
    strSql = strSql & "IM.intComposicaoDaReceita "
    strSql = strSql & "FROM "
    strSql = strSql & gstrIsencaoImunidade & " IM,"
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
    
    If intTipoDeInscricao = 0 Then
       strSql = strSql & gstrContribuinte & " CO1, "
    End If
        
    Select Case intTipoDeInscricao
        Case 0
            strSql = strSql & gstrImobiliario & " A "
        Case 1
            strSql = strSql & gstrEconomico & " A "
    End Select
    
    strSql = strSql & "WHERE "
    
    If lngPkid > 0 Then
        strSql = strSql & "IM.Pkid = " & lngPkid & " AND "
    End If
    strSql = strSql & "A.PKId = IM.intIdentificacao AND "
    strSql = strSql & "CO.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & " A.intContribuinte AND "
    strSql = strSql & "CR.pkid = IM.intComposicaoDaReceita AND "
    
    'If UCase(strModoOperacao) = gstrPreencherLista Then
       ' dbcintIdentificacao.TabIndex = ""
      ' If dbcintIdentificacao.Text = "" Then
     '      strSql = strSql & "CR.strdescricao = '" & dbcintComposicaoDaReceita.Text & "' AND "
     '  End If
    'End If
    
    If dbcintIdentificacao.Text <> "" Then
        If intTipoDeInscricao = 0 Then
           strSql = strSql & "A.strInscricao LIKE " & "'" & UCase(String(gintLenInscricao - gintRetornaTamanhoMascara(TYP_IMOBILIARIA), "0") & dbcintIdentificacao.Text) & "%' AND "
          'strSql = strSql & "A.strInscricao = '" & (String(gintLenInscricao - Len(Trim(dbcintIdentificacao.Text)), "0") & Trim(dbcintIdentificacao.Text)) & "' AND "
        Else
           strSql = strSql & "A.strInscricaoCadastral LIKE " & "'" & UCase(String(gintLenInscricao - gintRetornaTamanhoMascara(TYP_ECONOMICA), "0") & dbcintIdentificacao.Text) & "%' AND "
          'strSql = strSql & "A.strInscricaoCadastral = '" & (String(gintLenInscricao - Len(Trim(dbcintIdentificacao.Text)), "0") & Trim(dbcintIdentificacao.Text)) & "' AND "
        End If
    End If
    
    If intTipoDeInscricao = 0 Then
       strSql = strSql & "CO1.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & " A.intPromissario AND "
    End If
    
    strSql = strSql & "IM.bitTipoDeInscricao = " & intTipoDeInscricao & " "
    
    strSql = strSql & "ORDER BY " & gstrCONVERT(cdt_numeric, IIf(intTipoDeInscricao = 0, "strinscricao", "Strinscricaocadastral"))
    
    strQuery = strSql
    
End Function

Private Function strQueryRelatorio() As String
    'RESPONSAVEL LEANDRO    29/06/2004
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT II.pkid, " & _
       gstrRIGHT("IM.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " Inscricao, " & _
       "CO.StrNome Contribuinte, " & _
       "CR.strDescricao Composicao, " & _
       "CR.intUtilizacao, " & _
       "RE.strDescricao RECEITA "
    strSql = strSql & "FROM " & gstrIsencaoImunidade & " II, " & _
       gstrImobiliario & " IM , " & _
       gstrComposicaoDaReceita & " CR, " & _
       gstrContribuinte & " CO, " & _
       gstrIsencaoReceita & " IR, " & _
       gstrReceita & " RE "
    strSql = strSql & "WHERE  IM.PKId = II.intIdentificacao AND " & _
       "CR.pkid = II.intComposicaoDaReceita AND " & _
       "CO.PKId = IM.intContribuinte AND " & _
       "IR.INTISENCAOIMUNIDADE  = II.pkid AND " & _
       "RE.Pkid = IR.intReceita "
    
    strSql = strSql & " UNION ALL "
    
    strSql = strSql & "SELECT II.pkid, " & _
       gstrRIGHT("EC.strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " Inscricao, " & _
       "CO.StrNome Contribuinte, " & _
       "CR.strDescricao Composicao, " & _
       "CR.intUtilizacao, " & _
       "RE.strDescricao RECEITA "
    strSql = strSql & "FROM " & gstrIsencaoImunidade & " II, " & _
       gstrEconomico & " EC , " & _
       gstrComposicaoDaReceita & " CR, " & _
       gstrContribuinte & " CO, " & _
       gstrIsencaoReceita & " IR, " & _
       gstrReceita & " RE "
    strSql = strSql & "WHERE  EC.PKId = II.intIdentificacao AND " & _
       "CR.pkid = II.intComposicaoDaReceita AND " & _
       "CO.PKId = EC.intContribuinte AND " & _
       "IR.INTISENCAOIMUNIDADE  = II.pkid AND " & _
       "RE.Pkid = IR.intReceita "
    
    strQueryRelatorio = strSql
      
End Function

Private Function strQuerryComposicao(Index As Integer, Optional blntag As Boolean) As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & " SELECT PKId, Ltrim(Rtrim(strDescricao)) as strDescricao "
    strSql = strSql & " FROM " & gstrComposicaoDaReceita
    
    If Index = 0 Then
        strSql = strSql & " WHERE intUtilizacao = 1 "
    ElseIf Index = 1 Then
        strSql = strSql & " WHERE intUtilizacao = 2 "
    End If
    
    If blntag = True Then
        strSql = strSql & "AND strDescricao like '" & UCase(dbcintComposicaoDaReceita.Text) & "%'"
    End If
    
    strSql = strSql & " ORDER BY strDescricao "
    
    strQuerryComposicao = strSql
    
End Function

Private Sub optbitTipoDeInscricao_Click(Index As Integer)
    Dim strSql      As String
    Dim intIndice   As Integer
    
    optbitTipoDeInscricao(Index).CausesValidation = True
    
    For intIndice = 0 To 1
        If intIndice <> Index Then
            optbitTipoDeInscricao(intIndice).CausesValidation = False
        End If
    Next

    dbcintTipoIsencaoImunidade.Text = ""
    
    If optbitTipoDeInscricao(Index).Value Then
'        Set tdb_Receita.DataSource = Nothing
'        Set dbcintIdentificacao.RowSource = Nothing
'        Set dbc_intReceita.RowSource = Nothing
'
'        dbc_intReceita.Text = ""
'        dbcintIdentificacao.Text = ""
'        txt_strContribuinte.Text = ""
'        txt_strPromissario.Text = ""
'        lvw_Periodo.ListItems.Clear
'        lvw_Receita.ListItems.Clear
'        mblnPrimeiraVez = False
'        mblnAlterando = False
        
        LeDaTabelaParaObj gstrComposicaoDaReceita, dbcintComposicaoDaReceita, strQuerryComposicao(Index)
        dbcintComposicaoDaReceita.Tag = strQuerryComposicao(Index) & ";strDescricao"
        
        Select Case Index
            Case 0
                dbcintIdentificacao.Tag = strQueryImobiliario(Index) & "Order by strinscricao;strinscricao"
            Case 1
                dbcintIdentificacao.Tag = strQueryImobiliario(Index) & "Order by strInscricaoCadastral;Strinscricaocadastral"
        End Select
    End If

End Sub

Private Function blnDadosOk(strOperacao As String) As Boolean
    Dim i As Integer
    Dim blnMarcado As Boolean
    
    If strOperacao = "SALVAR" Then
        If Trim(dbcintIdentificacao.Text) <> "" Then
           If VerificaInscricao = False Then
              dbcintIdentificacao.SetFocus
              Exit Function
           End If
        Else
           ExibeMensagem "A inscrição deve ser preenchida."
           dbcintIdentificacao.SetFocus
           Exit Function
        End If
        
        Do While (i < 5 And Not blnMarcado)
            If optbitTipoDeInscricao(i).Value Then
                blnMarcado = True
            End If
            i = i + 1
        Loop
        If Not blnMarcado Then
            ExibeMensagem "Selecione um tipo de inscrição ."
            Exit Function
        End If
       
        If dbcintComposicaoDaReceita.Text = "" Then
            ExibeMensagem "Selecione uma " & lblintComposicaoDaReceita.Caption & " ."
            dbcintComposicaoDaReceita.SetFocus
            Exit Function
        End If
        blnMarcado = False
        i = 0
        Do While (i < 3 And Not blnMarcado)
            If optbitDefinicao(i).Value Then
                blnMarcado = True
            End If
            i = i + 1
        Loop
        If Not blnMarcado Then
            ExibeMensagem "Selecione um tipo de " & fra_Isencao.Caption & " ."
            Exit Function
        End If
        
        If Trim(dbcintTipoIsencaoImunidade.Text) = "" Then
           ExibeMensagem "O campo Tipo de Isenção Imunidade deve ser preenchido."
           dbcintTipoIsencaoImunidade.SetFocus
           Exit Function
        End If
        If VerificaRegistro = False Then
           Exit Function
        End If
    End If
    
    If lvw_Receita.ListItems.Count < 1 Then
        ExibeMensagem "É necessário 1 receita no mínimo para isenção."
        tab_3dPasta.Tab = 0
        Exit Function
    End If
    
'    If lvw_Periodo.ListItems.Count < 1 Then
'        ExibeMensagem "É necessário 1 período no mínimo para isenção."
'        tab_3dPasta.Tab = 1
'        Exit Function
'    End If
    
    blnDadosOk = True
End Function

Private Sub txt_dblAliquota_GotFocus()
    MarcaCampo txt_dblAliquota
End Sub

Private Sub txt_dblAliquota_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblAliquota
End Sub

Private Sub txt_dblAliquota_LostFocus()
    txt_dblAliquota = gstrConvVrDoSql(txt_dblAliquota, 2)
End Sub

Private Sub txt_dtmData_GotFocus()
    txt_dtmData = gstrDataFormatada(Date)
    MarcaCampo txt_dtmData
End Sub

Private Sub txt_dtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmData
End Sub

Private Sub txt_dtmData_LostFocus()
    txt_dtmData = gstrDataFormatada(txt_dtmData)
    If Trim(txt_dtmData.Text) <> "" Then
        If gblnDataValida(txt_dtmData, True) = False Then
            txt_dtmData.SetFocus
        End If
    End If
End Sub

Private Sub txt_dtmFinal_LostFocus()
    txt_dtmFinal.Text = gstrDataFormatada(txt_dtmFinal.Text)
    If Trim(txt_dtmFinal.Text) <> "" Then
        If gblnDataValida(txt_dtmFinal, True) = False Then
            txt_dtmFinal.SetFocus
        End If
    End If
End Sub

Private Sub txt_dtmInicial_LostFocus()
    txt_dtmInicial.Text = gstrDataFormatada(txt_dtmInicial.Text)
    If Trim(txt_dtmInicial.Text) <> "" Then
        If gblnDataValida(txt_dtmInicial, True) = False Then
            txt_dtmInicial.SetFocus
        End If
    End If

End Sub

Private Sub txt_dtmInicial_GotFocus()
    MarcaCampo txt_dtmInicial
End Sub

Private Sub txt_dtmInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmInicial, True
End Sub

Private Sub txt_dtmFinal_GotFocus()
    MarcaCampo txt_dtmFinal
End Sub

Private Sub txt_dtmFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmFinal, True
End Sub

Private Sub txt_strCodigoProcesso_GotFocus()
    If Trim(txt_dtmData.Text) <> "" And Trim(txt_dtmInicial.Text) <> "" And Trim(txt_dtmFinal.Text) <> "" Then
        UltimoProcesso
    End If
End Sub

Private Sub txt_strObservacao_GotFocus()
    MarcaCampo txt_strobservacao
End Sub

Private Sub txt_strobservacao_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        CaracterValido KeyAscii, "A", txt_strobservacao
    End If
End Sub

Private Function strQueryImobiliario(ID As Integer) As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "Select "
    
    Select Case ID
        Case 0
            strSql = strSql & "Pkid, " & gstrRIGHT("strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao "
        Case 1
            strSql = strSql & "Pkid, " & gstrRIGHT("strInscricaoCadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " Strinscricaocadastral "
    End Select
    
    strSql = strSql & "From "
    
    Select Case ID
        Case 0
            strSql = strSql & gstrImobiliario
        Case 1
            strSql = strSql & gstrEconomico
    End Select
    
    strSql = strSql & " "
    strQueryImobiliario = strSql

End Function

Private Sub PreencheContribuinte()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    Dim intCont As Byte
    Dim intIndice As Byte
    
    For intCont = 0 To 1
        If optbitTipoDeInscricao(intCont).Value Then
            intIndice = intCont
            Exit For
        End If
    Next
    
     
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "B.strNome "
    strSql = strSql & "FROM "
    
    Select Case intIndice
        Case 0
            strSql = strSql & gstrImobiliario & " A, "
        Case 1
            strSql = strSql & gstrEconomico & " A, "
    End Select
    
    strSql = strSql & gstrContribuinte & " B "
    strSql = strSql & "Where "
    strSql = strSql & "B.Pkid = A.Intcontribuinte AND "
    strSql = strSql & "A.pkid =" & dbcintIdentificacao.BoundText
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        txt_strContribuinte.Text = gstrENulo(adoResultado!strNome)
    End If
    
    

End Sub

Private Sub LimpaCombos()
    Set dbcintComposicaoDaReceita.DataSource = Nothing
    dbcintComposicaoDaReceita.ReFill
    Set dbcintIdentificacao.RowSource = Nothing
    Set dbcintTipoIsencaoImunidade.RowSource = Nothing
    Set dbc_intReceita.RowSource = Nothing
    dbc_intReceita.Text = ""
    dbcintIdentificacao.Text = ""
    txt_dblAliquota.Text = ""
    txt_strContribuinte.Text = ""
    txt_strPromissario.Text = ""
End Sub

Private Function strQueryTipoIsencaoImunidade() As String
Dim strSql  As String
Dim intCont As Integer
    
    strSql = "SELECT Pkid,"
    strSql = strSql & " strDescricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrTipoIsencaoImunidade
    For intCont = optbitDefinicao.LBound To optbitDefinicao.UBound
        If optbitDefinicao(intCont).Value = True Then
            strSql = strSql & " WHERE intTipo =" & optbitDefinicao(intCont).Index
            Exit For
        End If
    Next
    
    strSql = strSql & " ORDER BY strDescricao"

    strQueryTipoIsencaoImunidade = strSql

End Function

Private Function VerificaInscricao() As Boolean
Dim intCont As Byte
Dim intIndice As Byte
Dim strSql As String
Dim strNomeTabela As String
Dim adoResultado As ADODB.Recordset

  For intCont = 0 To 1
    If optbitTipoDeInscricao(intCont).Value Then
       intIndice = intCont
       Exit For
    End If
  Next
  
  strSql = "SELECT * FROM "
  
  Select Case intIndice
        Case 0
            strSql = strSql & gstrImobiliario & " A "
            strNomeTabela = gstrImobiliario
        Case 1
            strSql = strSql & gstrEconomico & " A "
            strNomeTabela = gstrEconomico
  End Select
 
  strSql = strSql & "WHERE A.strInscricao" & IIf(intIndice = 1, "Cadastral", Empty)
  strSql = strSql & " ='" & String(gintLenInscricao - Len(Trim(dbcintIdentificacao.Text)), "0") & Trim(dbcintIdentificacao.Text) & "'"
  
  If gblnExisteCodigo(1, strNomeTabela, "strInscricao" & IIf(intIndice = 1, "Cadastral", Empty), String(gintLenInscricao - Len(Trim(dbcintIdentificacao.Text)), "0") & Trim(dbcintIdentificacao.Text)) = False Then
     ExibeMensagem "A Inscrição " & Trim(dbcintIdentificacao.Text) & " não existe."
     Exit Function
  End If
  
  Set gobjBanco = New clsBanco
  If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
     If adoResultado.EOF Then
        Set adoResultado = Nothing
        ExibeMensagem "A Inscrição " & Trim(dbcintIdentificacao.Text) & " não existe."
        Exit Function
     End If
  End If
  
  VerificaInscricao = True
End Function

Private Function VerificaRegistro() As Boolean
    Dim adoResultado As ADODB.Recordset
    Dim strSql As String
    Dim bytIndice As Byte
    Dim bytDefinicao As Byte

    For bytIndice = 0 To 1
        If optbitTipoDeInscricao(bytIndice).Value = True Then
           Exit For
        End If
    Next
  
    For bytDefinicao = 0 To 2
        If optbitDefinicao(bytDefinicao).Value = True Then
           Exit For
        End If
    Next
  
  strSql = ""
  strSql = strSql & "SELECT * FROM " & gstrIsencaoImunidade
  strSql = strSql & " WHERE bitTipoDeInscricao=" & bytIndice
  strSql = strSql & " AND intIdentificacao=" & dbcintIdentificacao.BoundText
  strSql = strSql & " AND intComposicaoDaReceita=" & dbcintComposicaoDaReceita.BoundText
  strSql = strSql & " AND bitDefinicao=" & bytDefinicao
  strSql = strSql & " AND intTipoIsencaoImunidade=" & dbcintTipoIsencaoImunidade.BoundText

    
  Set gobjBanco = New clsBanco
  If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
     If Not adoResultado.EOF Then
        Set adoResultado = Nothing
        If Not mblnAlterando = True Then
           ExibeMensagem "O registro já se encontra cadastrado."
           Exit Function
        'Else
        '   ExibeMensagem "O registro já se encontra cadastrado ou o mesmo não sofreu alteração."
        End If
        
     End If
     Set adoResultado = Nothing
  End If
  
  VerificaRegistro = True
End Function


Private Function strQueryReceita() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT REC.PKId, Ltrim(Rtrim(REC.strDescricao)) as strDescricao"
    strSql = strSql & " FROM " & gstrReceita & " REC, "
    strSql = strSql & gstrValorCompoRec & " VCR"
    strSql = strSql & " WHERE REC.PKId = VCR.intReceita "
    strSql = strSql & " AND  VCR.intComposicaoDaReceita = " & dbcintComposicaoDaReceita.BoundText
    strSql = strSql & " ORDER BY REC.strDescricao"

    strQueryReceita = strSql
    
End Function

Private Sub PreencheReceitas(lngPkid As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "IR.Pkid, "
    strSql = strSql & "R.Strdescricao AS strReceita, "
    strSql = strSql & "IR.DBLALIQUOTA, "
    strSql = strSql & "R.PKID as IntReceita "
    strSql = strSql & "From "
    strSql = strSql & gstrIsencaoImunidade & " I, "
    strSql = strSql & gstrIsencaoReceita & " IR, "
    strSql = strSql & gstrReceita & " R "
    strSql = strSql & "Where "
    strSql = strSql & "I.Pkid = IR.Intisencaoimunidade AND "
    strSql = strSql & "R.Pkid = IR.Intreceita AND "
    strSql = strSql & "I.Pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    
    lvw_Receita.ListItems.Clear
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                If .EOF = False Then
                    Do While Not .EOF
                        Set mobjLista = lvw_Receita.ListItems.Add(, , gstrENulo(!Pkid))
                        mobjLista.SubItems(1) = gstrENulo(!strReceita)
                        mobjLista.SubItems(2) = gstrConvVrDoSql(gstrENulo(!dblAliquota), 2, , True)
                        mobjLista.SubItems(3) = gstrENulo(!intReceita)
                        .MoveNext
                    Loop
                End If
            End With
        End If
    End If
End Sub

Private Sub PreenchePeriodo(lngPkid As Long)
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "IP.Pkid, "
    strSql = strSql & "IP.DTMDATA, "
    strSql = strSql & "IP.DTMINICIAL, "
    strSql = strSql & "IP.DTMFINAL, "
    strSql = strSql & "IP.bytposicao, "
    strSql = strSql & "IP.strCodigoProcesso, "
    strSql = strSql & "IP.intExercicioProcesso, "
    strSql = strSql & "IP.bitDigitoProcesso, "
    strSql = strSql & "IP.bytcancelamento, "
    strSql = strSql & "IP.strobservacao "
    strSql = strSql & "From "
    strSql = strSql & gstrIsencaoImunidade & " I, "
    strSql = strSql & gstrIsencaoPeriodo & " IP "
    strSql = strSql & "Where "
    strSql = strSql & "I.Pkid = IP.Intisencaoimunidade AND "
    strSql = strSql & "I.Pkid = " & lngPkid
    strSql = strSql & " Order by IP.DTMINICIAL DESC"
    
    Set gobjBanco = New clsBanco
    
    lvw_Periodo.ListItems.Clear
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                If .EOF = False Then
                    Do While Not .EOF
                        Set mobjLista = lvw_Periodo.ListItems.Add(, , gstrENulo(!Pkid))
                        mobjLista.SubItems(1) = gstrDataFormatada(gstrENulo(!DTMDATA))
                        mobjLista.SubItems(2) = gstrDataFormatada(gstrENulo(!DTMINICIAL))
                        mobjLista.SubItems(3) = gstrDataFormatada(gstrENulo(!DTMFINAL))
                        If gstrENulo(!bytPosicao) = 0 Then
                            mobjLista.SubItems(4) = "Deferido"
                        ElseIf gstrENulo(!bytPosicao) = 1 Then
                            mobjLista.SubItems(4) = "Indeferido"
                        ElseIf gstrENulo(!bytPosicao) = 2 Then
                            mobjLista.SubItems(4) = "Em Andamento"
                        End If
                        'PROCESSO
                        If !strCodigoProcesso <> "" And !intExercicioProcesso <> "" & _
                           !bitDigitoProcesso <> "" Then
                           mobjLista.SubItems(5) = gstrENulo(!strCodigoProcesso) & "/" & gstrENulo(!intExercicioProcesso) & "-" & gstrENulo(!bitDigitoProcesso)
                        Else
                           mobjLista.SubItems(5) = ""
                        End If
                        mobjLista.SubItems(6) = IIf(CBool(gstrENulo(!Bytcancelamento)), "Cancelado", "Não cancelado")
                        mobjLista.SubItems(7) = gstrENulo(!strObservacao)
                        mobjLista.SubItems(8) = gstrENulo(!Bytcancelamento)
                        mobjLista.SubItems(9) = gstrENulo(!bytPosicao)
                        
                        .MoveNext
                    Loop
                End If
            End With
        End If
    End If
End Sub

Private Function GravaReceitas(lngPkid As Long) As Boolean
    Dim strSql              As String
    Dim intFor              As Integer
  
    GravaReceitas = False

    strSql = IIf(bytDBType = Oracle, "Begin ", " ")
        
    With lvw_Receita
        strSql = strSql & "DELETE " & gstrIsencaoReceita
        strSql = strSql & " WHERE intIsencaoImunidade = "
        strSql = strSql & Val(lngPkid)
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")
    End With
        
    For intFor = 1 To lvw_Receita.ListItems.Count
        With lvw_Receita
            strSql = strSql & "INSERT INTO " & gstrIsencaoReceita
            strSql = strSql & " (intIsencaoImunidade,"
            strSql = strSql & " intReceita,"
            strSql = strSql & " dblAliquota,"
            strSql = strSql & " dtmDtAtualizacao,"
            strSql = strSql & " lngCodUsr)"
            strSql = strSql & " VALUES( "
            strSql = strSql & Val(lngPkid) & ", "
            strSql = strSql & .ListItems(intFor).SubItems(3) & ", "
            strSql = strSql & gstrConvVrParaSql(.ListItems(intFor).SubItems(2)) & ", "
            strSql = strSql & strGETDATE & ", "
            strSql = strSql & glngCodUsr
            strSql = strSql & ")"
            strSql = strSql & IIf(bytDBType = Oracle, ";", "")
        End With
    Next
    strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSql) = False Then
        ExibeMensagem "Ocorreu um erro ao gravar as Receitas. Os dados não foram gravados."
        GravaReceitas = False
        Exit Function
    Else
        GravaReceitas = True
    End If
End Function

Private Function GravaPeriodo(lngPkid As Long) As Boolean
    Dim strSql              As String
    Dim strAux              As String
    Dim intFor              As Integer
    Dim bytPos              As Byte
    
    GravaPeriodo = False

    strSql = IIf(bytDBType = Oracle, "Begin ", " ")
    
    For intFor = 1 To lvw_Periodo.ListItems.Count
        With lvw_Periodo
            If Val(.ListItems(intFor).Text) <> Val("0") Then
                strAux = strAux & .ListItems(intFor).Text & ","
            End If
        End With
    Next
    
    If Trim(strAux) <> "" Then
        strSql = strSql & "Delete From " & gstrIsencaoPeriodo
        strSql = strSql & " Where  "
        strSql = strSql & "intisencaoimunidade = " & lngPkid
        strAux = Mid(strAux, 1, Len(strAux) - 1)
        strSql = strSql & " AND not Pkid in(" & strAux & ")"
        strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
    End If
        
    For intFor = 1 To lvw_Periodo.ListItems.Count
        With lvw_Periodo
            If Val(.ListItems(intFor).Text) <> Val("0") Then
                strSql = strSql & "Update " & gstrIsencaoPeriodo
                strSql = strSql & " Set "
                strSql = strSql & "intisencaoimunidade = " & lngPkid & ", "
                
                strSql = strSql & "dtmData = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(1)) & ", "
                'strSql = strSql & "dtminicial = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(2)) & ", "
                strSql = strSql & "dtminicial = " & IIf(Not gblnDataValida(.ListItems(intFor).SubItems(2)), "'" & .ListItems(intFor).SubItems(2) & "', ", gstrConvDtParaSql(.ListItems(intFor).SubItems(2)) & ", ")
                'strSql = strSql & "dtmfinal = " & gstrConvDtParaSql(.ListItems(intFor).SubItems(3)) & ", "
                strSql = strSql & "dtmfinal = " & IIf(Not gblnDataValida(.ListItems(intFor).SubItems(3)), "'" & .ListItems(intFor).SubItems(2) & "', ", gstrConvDtParaSql(.ListItems(intFor).SubItems(3)) & ", ")
                'PROCESSO
                If .ListItems(intFor).SubItems(5) <> "" Then
                   bytPos = InStr(1, .ListItems(intFor).SubItems(5), "-")
                   strSql = strSql & "bitDigitoProcesso = " & gstrENulo(Trim(Mid(.ListItems(intFor).SubItems(5), bytPos + 1, 2)), , True) & ", "
                   strSql = strSql & "intExercicioProcesso = " & gstrENulo(Trim(Mid(.ListItems(intFor).SubItems(5), bytPos - 4, 4)), , True) & ", "
                   strSql = strSql & "strCodigoProcesso = '" & Mid(.ListItems(intFor).SubItems(5), 1, (bytPos - 6)) & "', "
                Else
                   strSql = strSql & "bitDigitoProcesso = NULL, "
                   strSql = strSql & "intExercicioProcesso = NULL, "
                   strSql = strSql & "strCodigoProcesso = NULL, "
                End If
                
                strSql = strSql & "strObservacao = '" & Trim(.ListItems(intFor).SubItems(7)) & "', "
                strSql = strSql & "bytCancelamento = " & gstrENulo(.ListItems(intFor).SubItems(8), , True) & ", "
                strSql = strSql & "bytPosicao = " & gstrENulo(.ListItems(intFor).SubItems(9), , True) & ", "
                strSql = strSql & "dtmdtatualizacao = " & strGETDATE & ", "
                strSql = strSql & "lngcodusr = " & glngCodUsr
                strSql = strSql & " Where pkid = " & lvw_Periodo.ListItems(intFor).Text
                strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
            Else
                strSql = strSql & "INSERT INTO " & gstrIsencaoPeriodo & " "
                strSql = strSql & "(intIsencaoImunidade, "
                strSql = strSql & "dtmData, "
                strSql = strSql & "dtminicial, "
                strSql = strSql & "dtmfinal, "
                'PROCESSO
                strSql = strSql & "strCodigoProcesso, "
                strSql = strSql & "intExercicioProcesso, "
                strSql = strSql & "bitDigitoProcesso, "
                strSql = strSql & "strObservacao, "
                strSql = strSql & "bytCancelamento, "
                strSql = strSql & "bytPosicao, "
                strSql = strSql & "dtmDtAtualizacao, "
                strSql = strSql & "lngCodUsr) "
                strSql = strSql & "VALUES( "
                strSql = strSql & lngPkid & ", "
                strSql = strSql & gstrConvDtParaSql(.ListItems(intFor).SubItems(1)) & ", "
                strSql = strSql & gstrConvDtParaSql(.ListItems(intFor).SubItems(2)) & ", "
                strSql = strSql & gstrConvDtParaSql(.ListItems(intFor).SubItems(3)) & ", "
                'PROCESSO
                If .ListItems(intFor).SubItems(5) <> "" Then
                   bytPos = InStr(1, .ListItems(intFor).SubItems(5), "-")
                   strSql = strSql & "'" & Mid(.ListItems(intFor).SubItems(5), 1, (bytPos - 6)) & "', "
                   strSql = strSql & gstrENulo(Trim(Mid(.ListItems(intFor).SubItems(5), bytPos - 4, 4)), , True) & ", "
                   strSql = strSql & gstrENulo(Trim(Mid(.ListItems(intFor).SubItems(5), bytPos + 1, 2)), , True) & ", "
                Else
                   strSql = strSql & "bitDigitoProcesso = NULL, "
                   strSql = strSql & "intExercicioProcesso = NULL, "
                   strSql = strSql & "strCodigoProcesso = NULL, "
                End If
                
                strSql = strSql & "'" & Trim(.ListItems(intFor).SubItems(7)) & "', "
                strSql = strSql & gstrENulo(.ListItems(intFor).SubItems(8), , True) & ", "
                strSql = strSql & gstrENulo(.ListItems(intFor).SubItems(9), , True) & ", "
                strSql = strSql & strGETDATE & ", "
                strSql = strSql & glngCodUsr
                strSql = strSql & ")"
                strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
            End If
        End With
    Next
    strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    
    If gobjBanco.Execute(strSql) = False Then
        ExibeMensagem "Ocorreu um erro ao gravar os períodos. Os dados não foram gravados."
        GravaPeriodo = False
        Exit Function
    Else
        GravaPeriodo = True
    End If
End Function

Private Sub IncluirItemNoGrid()
    Dim intInd          As Integer
    Dim blnGravar       As Boolean
    Dim bytPosicao      As Byte
    Dim blnImprimeGuia  As Boolean
    
    blnImprimeGuia = blnAlteraPeriodo
    
    If tab_3dPasta.Tab = 0 Then
    
        If blnDadosItens = False Then Exit Sub
        
        With lvw_Receita
            If blnAlteraReceita Then
                For intInd = 1 To .ListItems.Count
                    If .SelectedItem.Index <> intInd Then
                        If Trim(dbc_intReceita.BoundText) = .ListItems(intInd).SubItems(3) Then
                            ExibeMensagem "Não é possível incluir receitas iguais."
                            dbc_intReceita.SetFocus
                            Exit Sub
                        End If
                    End If
                Next
                .SelectedItem.Text = ""
                .SelectedItem.SubItems(1) = dbc_intReceita.Text
                .SelectedItem.SubItems(2) = gstrConvVrDoSql(txt_dblAliquota.Text, 2)
                .SelectedItem.SubItems(3) = dbc_intReceita.BoundText
                blnAlteraReceita = False
            Else
                For intInd = 1 To .ListItems.Count
                    'If .SelectedItem.Index <> intInd Then
                        If Trim(dbc_intReceita.BoundText) = .ListItems(intInd).SubItems(3) Then
                            ExibeMensagem "Não é possível incluir receitas iguais."
                            dbc_intReceita.Text = ""
                            txt_dblAliquota.Text = ""
                            dbc_intReceita.SetFocus
                            Exit Sub
                        End If
                    'End If
                Next
                Set mobjLista = .ListItems.Add(, , "")
                mobjLista.SubItems(1) = dbc_intReceita.Text
                mobjLista.SubItems(2) = gstrConvVrDoSql(txt_dblAliquota.Text, 2)
                mobjLista.SubItems(3) = dbc_intReceita.BoundText
            End If
        End With
        LimpaReceita
    ElseIf tab_3dPasta.Tab = 1 Then
    
        If blnDadosItens = False Then Exit Sub
        
        If mblnAlterando Then
            If gblnExclusaoGravacaoOk("SALVAR", "Deseja realmente " & IIf(blnAlteraPeriodo, "alterar", "inserir novo") & " o período ") Then
                blnGravar = True
            Else
                Exit Sub
            End If
        End If
        
        With lvw_Periodo
            If blnAlteraPeriodo Then
                For intInd = 1 To .ListItems.Count
                    If .SelectedItem.Index <> intInd Then
                        If Trim(txt_dtmInicial.Text) = .ListItems(intInd).SubItems(2) And Trim(txt_dtmFinal.Text) = .ListItems(intInd).SubItems(3) Then
                            ExibeMensagem "Não é possível incluir períodos iguais."
                            Exit Sub
                        End If
                    End If
                Next
                .SelectedItem.SubItems(1) = gstrDataFormatada(txt_dtmData.Text)
                .SelectedItem.SubItems(2) = gstrDataFormatada(txt_dtmInicial.Text)
                .SelectedItem.SubItems(3) = gstrDataFormatada(txt_dtmFinal.Text)
                If opt_bytPosicao(0).Value = True Then
                    .SelectedItem.SubItems(4) = "Deferido"
                    bytPosicao = 0
                ElseIf opt_bytPosicao(1).Value = True Then
                    .SelectedItem.SubItems(4) = "Indeferido"
                    bytPosicao = 1
                ElseIf opt_bytPosicao(2).Value = True Then
                    .SelectedItem.SubItems(4) = "Em Andamento"
                    bytPosicao = 2
                End If
                'PROCESSO
                .SelectedItem.SubItems(5) = Trim(txt_strCodigoProcesso.Text) & "/" & Trim(txt_intExercicioProcesso.Text) & _
                                             "-" & Trim(txt_bitDigitoProcesso.Text)
                
                .SelectedItem.SubItems(6) = IIf(chk_bytCancelado.Value, "Cancelado", "Não cancelado")
                .SelectedItem.SubItems(7) = Trim(txt_strobservacao)
                .SelectedItem.SubItems(8) = Abs(chk_bytCancelado.Value)
                .SelectedItem.SubItems(9) = bytPosicao
                blnAlteraPeriodo = False
            Else
                For intInd = 1 To .ListItems.Count
                    If Trim(txt_dtmInicial.Text) = .ListItems(intInd).SubItems(2) And Trim(txt_dtmFinal.Text) = .ListItems(intInd).SubItems(3) Then
                        ExibeMensagem "Não é possível incluir períodos iguais."
                        Exit Sub
                    End If
                Next
                Set mobjLista = .ListItems.Add(, , "")
                mobjLista.SubItems(1) = gstrDataFormatada(txt_dtmData.Text)
                mobjLista.SubItems(2) = gstrDataFormatada(txt_dtmInicial.Text)
                mobjLista.SubItems(3) = gstrDataFormatada(txt_dtmFinal.Text)
                If opt_bytPosicao(0).Value = True Then
                    mobjLista.SubItems(4) = "Deferido"
                    bytPosicao = 0
                ElseIf opt_bytPosicao(1).Value = True Then
                    mobjLista.SubItems(4) = "Indeferido"
                    bytPosicao = 1
                ElseIf opt_bytPosicao(2).Value = True Then
                    mobjLista.SubItems(4) = "Em Andamento"
                    bytPosicao = 2
                End If
                'PROCESSO
                mobjLista.SubItems(5) = Trim(txt_strCodigoProcesso.Text) & "/" & Trim(txt_intExercicioProcesso.Text) & _
                                             "-" & Trim(txt_bitDigitoProcesso.Text)
                                             
                mobjLista.SubItems(6) = IIf(chk_bytCancelado.Value, "Cancelado", "Não cancelado")
                mobjLista.SubItems(7) = Trim(txt_strobservacao)
                mobjLista.SubItems(8) = Abs(chk_bytCancelado.Value)
                mobjLista.SubItems(9) = bytPosicao
            End If
        End With
        
        If blnGravar Then
            GravaPeriodo CLng(txtPKId)
            PreenchePeriodo CLng(txtPKId)
        End If
        
        If Not blnImprimeGuia Then
            ImprimeRelatorio rptRenovacaoIsencao, strQueryRelatorioRenovacao(CLng(txtPKId), txt_dtmInicial.Text), "Renovação da Isenção de I.P.T.U."
        End If
        
        NovoPeriodo
    End If
End Sub

Private Function blnDadosItens() As Boolean

    blnDadosItens = False
        
    If tab_3dPasta.Tab = 0 Then
        If dbc_intReceita.MatchedWithList = False Then
            ExibeMensagem "O campo receita deve ser preenchido corretamente."
            dbc_intReceita.SetFocus
            Exit Function
        ElseIf Val(gstrConvVrParaSql(gstrConvVrDoSql(txt_dblAliquota.Text, 2, , True))) < 1 Then
            ExibeMensagem "O campo alíquota deve ser maior que 0."
            txt_dblAliquota.SetFocus
            Exit Function
        ElseIf Val(gstrConvVrParaSql(gstrConvVrDoSql(txt_dblAliquota.Text, 2, , True))) > 100 Then
            ExibeMensagem "O campo alíquota deve ser menor que 100."
            txt_dblAliquota.SetFocus
            Exit Function
        End If
    ElseIf tab_3dPasta.Tab = 1 Then
        If Trim(txt_dtmData) = "" Then
            ExibeMensagem "O campo data deve ser preenchido corretamente."
            txt_dtmData.SetFocus
            Exit Function
        ElseIf Trim(txt_dtmInicial) = "" Then
            ExibeMensagem "O campo data inicial deve ser preenchido corretamente."
            txt_dtmInicial.SetFocus
            Exit Function
        ElseIf Trim(txt_dtmFinal) = "" Then
            ExibeMensagem "O campo data final deve ser preenchido corretamente."
            txt_dtmInicial.SetFocus
            Exit Function
        ElseIf CDate(txt_dtmInicial) > CDate(txt_dtmFinal) Then
            ExibeMensagem "A data inicial deve ser menor que a data final."
            txt_dtmInicial.SetFocus
            Exit Function
        End If
    End If
    
    blnDadosItens = True
    
End Function

Private Sub LimpaReceita()
    dbc_intReceita.Text = ""
    txt_dblAliquota.Text = ""
End Sub

Private Sub LimpaPeriodo()
    txt_strCodigoProcesso.Text = ""
    txt_intExercicioProcesso.Text = ""
    txt_bitDigitoProcesso.Text = ""
    txt_dtmData.Text = ""
    txt_dtmInicial.Text = ""
    txt_dtmFinal.Text = ""
    txt_strobservacao.Text = ""
    chk_bytCancelado.Value = False
    opt_bytPosicao(2).Value = True
End Sub

Private Sub ExcluirItemNoGrid()
    If tab_3dPasta.Tab = 0 Then
        With lvw_Receita
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
        End With
        blnAlteraReceita = False
    ElseIf tab_3dPasta.Tab = 1 Then
        With lvw_Periodo
            If .ListItems.Count > 0 Then
                .ListItems.Remove .SelectedItem.Index
            End If
        End With
        blnAlteraPeriodo = False
    End If
End Sub

Private Sub DesabilitaIsencao()
    TrocaCorObjeto dbcintIdentificacao, True
    TrocaCorObjeto dbcintComposicaoDaReceita, True
    TrocaCorObjeto dbcintTipoIsencaoImunidade, True
    TrocaCorObjeto dbc_intReceita, True
    TrocaCorObjeto txt_dblAliquota, True
    optbitTipoDeInscricao.Item(0).Enabled = False
    optbitTipoDeInscricao.Item(1).Enabled = False
    optbitDefinicao.Item(0).Enabled = False
    optbitDefinicao.Item(1).Enabled = False
    optbitDefinicao.Item(2).Enabled = False
    cmd_TipoIsencaoImunidade.Enabled = False
    
End Sub

Private Sub abilitaIsencao()
    TrocaCorObjeto dbcintIdentificacao, False
    TrocaCorObjeto dbcintComposicaoDaReceita, False
    TrocaCorObjeto dbcintTipoIsencaoImunidade, False
    TrocaCorObjeto dbc_intReceita, False
    TrocaCorObjeto txt_dblAliquota, False
    optbitTipoDeInscricao.Item(0).Enabled = True
    optbitTipoDeInscricao.Item(1).Enabled = True
    optbitDefinicao.Item(0).Enabled = True
    optbitDefinicao.Item(1).Enabled = True
    optbitDefinicao.Item(2).Enabled = True
    cmd_TipoIsencaoImunidade.Enabled = True
End Sub

Private Function ExcluirIsencao() As Boolean
    Dim strSql As String
    
    If MsgBox("Confirma exclusão do registro de '" & dbcintIdentificacao.Text & "' ?", vbQuestion + vbYesNo) = vbYes Then
        gobjBanco.ExecutaBeginTrans
        
        strSql = IIf(bytDBType = Oracle, "Begin ", "")
        
        strSql = strSql & "DELETE FROM " & gstrIsencaoReceita
        strSql = strSql & " WHERE INTISENCAOIMUNIDADE = " & Val(txtPKId)
        strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
        
        strSql = strSql & "DELETE FROM " & gstrIsencaoPeriodo
        strSql = strSql & " WHERE INTISENCAOIMUNIDADE = " & Val(txtPKId)
        strSql = strSql & IIf(bytDBType = Oracle, ";", " ")

        strSql = strSql & "Delete From " & gstrIsencaoImunidade & " "
        strSql = strSql & "Where PKId = " & txtPKId
        strSql = strSql & IIf(bytDBType = Oracle, ";", " ")
        
        strSql = strSql & IIf(bytDBType = Oracle, "End;", "")

        Set gobjBanco = New clsBanco
        
        If gobjBanco.Execute(strSql) Then
            gobjBanco.ExecutaCommitTrans
            ExcluirIsencao = True
        Else
            gobjBanco.ExecutaRollbackTrans
            ExcluirIsencao = False
            ExibeMensagem "Ocorreu um erro ao excluir o registro."
        End If
    End If
        
End Function

Private Sub LimpaControlesPeriodo()
    txt_strobservacao.Text = ""
    txt_dtmData.Text = ""
    txt_dtmInicial.Text = ""
    txt_dtmFinal.Text = ""
    txt_strCodigoProcesso.Text = ""
    txt_intExercicioProcesso.Text = ""
    txt_bitDigitoProcesso.Text = ""
    opt_bytPosicao(2) = vbChecked
End Sub

Private Function strQueryRelatorioRenovacao(intIsencaoImunidade As Long, dtmDataInicial As Date)
Dim strSql As String
    
    strSql = ""
    strSql = "SELECT CO.Pkid, " & Year(dtmDataInicial) & " intExercicio, IM.strInscricao, COP.strNome strPromissario, CO.strNome strProprietario, CO.strIdentidade, "
    strSql = strSql & " Ltrim(rtrim(" & gstrISNULL("TP.strSigla", "''") & ")) " & strCONCAT & " ' ' " & strCONCAT & " ltrim(rtrim(" & gstrISNULL("TT.strSigla", "''") & ")) " & strCONCAT & " ' ' " & strCONCAT & " ltrim(rtrim(LO.strDescricao)) AS strLogradouro, IM.intNumero, "
    strSql = strSql & " BA.strDescricao strBairro, IM.intCep "
    strSql = strSql & " FROM " & gstrIsencaoImunidade & " II, "
    strSql = strSql & gstrImobiliario & " IM, "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrContribuinte & " COP, "
    strSql = strSql & gstrLogradouro & " LO, "
    strSql = strSql & gstrBairro & " BA, "
    strSql = strSql & gstrTipoLogradouro & " TP, "
    strSql = strSql & gstrTituloLogradouro & " TT "
    strSql = strSql & " WHERE II.Pkid  = " & intIsencaoImunidade
    strSql = strSql & " AND II.intIdentificacao  = IM.Pkid "
    strSql = strSql & " AND IM.intContribuinte " & strOUTJSQLServer & "= CO.Pkid " & strOUTJOracle
    strSql = strSql & " AND IM.intLogradouro = LO.PKId "
    strSql = strSql & " AND LO.intTipoLogradouro    " & strOUTJSQLServer & "= TP.PKId " & strOUTJOracle
    strSql = strSql & " AND LO.intTituloLogradouro  " & strOUTJSQLServer & "= TT.PKId " & strOUTJOracle
    strSql = strSql & " AND IM.intBairro  " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle
    strSql = strSql & " AND IM.intPromissario " & strOUTJSQLServer & "= COP.Pkid " & strOUTJOracle
    
    strQueryRelatorioRenovacao = strSql
    
End Function

Private Sub NovoPeriodo()
    LimpaPeriodo
    UltimoProcesso
    txt_dtmData = gstrDataFormatada(Date)
    blnAlteraPeriodo = False
End Sub

