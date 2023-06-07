VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmcadISSQNEstimadoold 
   Caption         =   "formulario com objetos antigos"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   Icon            =   "CadISSQNEstimadoold.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   7200
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   12700
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Arrecadação de ISSQN Variável"
      TabPicture(0)   =   "CadISSQNEstimadoold.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Inscricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Endereco"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Lançamento"
      TabPicture(1)   =   "CadISSQNEstimadoold.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "tdb_Lista"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Lançamento"
         Height          =   2985
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   8475
         Begin VB.TextBox txtDtmLancamento 
            Height          =   285
            Left            =   840
            MaxLength       =   10
            TabIndex        =   42
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtQtdParcelas 
            Height          =   285
            Left            =   2760
            MaxLength       =   3
            TabIndex        =   41
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtdtmParcela 
            Height          =   285
            Left            =   5040
            MaxLength       =   10
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtintExercicio 
            Height          =   285
            Left            =   7200
            MaxLength       =   4
            TabIndex        =   39
            Top             =   360
            Width           =   735
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   840
            TabIndex        =   43
            Top             =   1200
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Composição da Receita"
            Height          =   195
            Left            =   1920
            TabIndex        =   48
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label lblDtmLancamento 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   360
            TabIndex        =   47
            Top             =   405
            Width           =   345
         End
         Begin VB.Label lblQtdParcelas 
            AutoSize        =   -1  'True
            Caption         =   "Parcelas"
            Height          =   195
            Left            =   2040
            TabIndex        =   46
            Top             =   405
            Width           =   615
         End
         Begin VB.Label lbldtmParcela 
            AutoSize        =   -1  'True
            Caption         =   "Data primeira parcela"
            Height          =   195
            Left            =   3480
            TabIndex        =   45
            Top             =   405
            Width           =   1500
         End
         Begin VB.Label lblintExericicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   6480
            TabIndex        =   44
            Top             =   405
            Width           =   675
         End
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   " Endereço do estabelecimento "
         Height          =   1305
         Left            =   150
         TabIndex        =   23
         Top             =   2850
         Width           =   8475
         Begin VB.TextBox txt_Bairro 
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   540
            Width           =   3705
         End
         Begin VB.TextBox txt_Logradouro 
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   3705
         End
         Begin VB.TextBox txt_Cep 
            Height          =   285
            Left            =   7125
            MaxLength       =   9
            TabIndex        =   28
            Top             =   840
            Width           =   1155
         End
         Begin VB.TextBox txt_UF 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txt_Complemento 
            Height          =   285
            Left            =   7125
            MaxLength       =   20
            TabIndex        =   26
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox txt_Numero 
            Height          =   285
            Left            =   5700
            MaxLength       =   6
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txt_Municipio 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   840
            Width           =   3705
         End
         Begin VB.Label lblintLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   630
            TabIndex        =   37
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lblintBairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   1035
            TabIndex        =   36
            Top             =   585
            Width           =   405
         End
         Begin VB.Label lblstrMunicipio 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   735
            TabIndex        =   35
            Top             =   885
            Width           =   705
         End
         Begin VB.Label lblintCep 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   6815
            TabIndex        =   34
            Top             =   885
            Width           =   285
         End
         Begin VB.Label lblstrUF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   5400
            TabIndex        =   33
            Top             =   885
            Width           =   210
         End
         Begin VB.Label lblintNumero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5430
            TabIndex        =   32
            Top             =   285
            Width           =   180
         End
         Begin VB.Label lbl_Complemento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6620
            TabIndex        =   31
            Top             =   285
            Width           =   480
         End
      End
      Begin VB.Frame fra_Inscricao 
         Height          =   2385
         Left            =   150
         TabIndex        =   1
         Top             =   360
         Width           =   8475
         Begin VB.TextBox txt_Nome 
            Height          =   285
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   600
            Width           =   4730
         End
         Begin VB.TextBox txt_InscricaoEstadual 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4695
            MaxLength       =   100
            TabIndex        =   11
            Top             =   1200
            Width           =   1605
         End
         Begin VB.Frame fra_NaturezaJuridica 
            Enabled         =   0   'False
            Height          =   1620
            Left            =   6435
            TabIndex        =   4
            Top             =   480
            Width           =   1935
            Begin VB.CheckBox chkblnMicroEmpresa 
               Caption         =   "Micro-Empresa"
               Height          =   195
               Left            =   150
               TabIndex        =   9
               Top             =   1200
               Width           =   1455
            End
            Begin VB.OptionButton opt_NaturezaJuridica 
               Caption         =   "Outros"
               Height          =   195
               Index           =   3
               Left            =   1110
               TabIndex        =   8
               Top             =   600
               Width           =   795
            End
            Begin VB.OptionButton opt_NaturezaJuridica 
               Caption         =   "SC"
               Height          =   195
               Index           =   2
               Left            =   1110
               TabIndex        =   7
               Top             =   270
               Width           =   585
            End
            Begin VB.OptionButton opt_NaturezaJuridica 
               Caption         =   "Física"
               Height          =   195
               Index           =   0
               Left            =   150
               TabIndex        =   6
               Top             =   270
               Width           =   915
            End
            Begin VB.OptionButton opt_NaturezaJuridica 
               Caption         =   "Jurídica"
               Height          =   195
               Index           =   1
               Left            =   150
               TabIndex        =   5
               Top             =   600
               Width           =   1035
            End
            Begin VB.Label lbl_Natureza 
               AutoSize        =   -1  'True
               Caption         =   " Natureza Jurídica "
               Height          =   195
               Left            =   150
               TabIndex        =   10
               Top             =   0
               Width           =   1350
            End
         End
         Begin VB.TextBox txt_CNPJCPF 
            Height          =   285
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1200
            Width           =   1605
         End
         Begin VB.TextBox txt_NomeFantasia 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1575
            MaxLength       =   100
            TabIndex        =   2
            Top             =   900
            Width           =   4730
         End
         Begin MSDataListLib.DataList dlsAtividade 
            Height          =   450
            Left            =   1575
            TabIndex        =   13
            Top             =   1800
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   794
            _Version        =   393216
         End
         Begin MSMask.MaskEdBox mskInscricaoImobiliaria 
            Height          =   285
            Left            =   1575
            TabIndex        =   14
            Top             =   1500
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo dbcintEconomico 
            Height          =   315
            Left            =   1575
            TabIndex        =   15
            Top             =   240
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblstrInscricaoImobiliaria 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Imobiliária"
            Height          =   195
            Left            =   60
            TabIndex        =   22
            Top             =   1545
            Width           =   1380
         End
         Begin VB.Label lbl_CNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ / CPF"
            Height          =   195
            Left            =   570
            TabIndex        =   21
            Top             =   1245
            Width           =   870
         End
         Begin VB.Label lblintAtividade 
            AutoSize        =   -1  'True
            Caption         =   "Atividade Empresa"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1845
            Width           =   1320
         End
         Begin VB.Label lblstrInscricaoEstadual 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Estadual"
            Height          =   195
            Left            =   3255
            TabIndex        =   19
            Top             =   1245
            Width           =   1305
         End
         Begin VB.Label lblstrNomeFantasia 
            AutoSize        =   -1  'True
            Caption         =   "Nome Fantasia"
            Height          =   195
            Left            =   375
            TabIndex        =   18
            Top             =   945
            Width           =   1065
         End
         Begin VB.Label lblintContribuinte 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social"
            Height          =   195
            Left            =   495
            TabIndex        =   17
            Top             =   645
            Width           =   945
         End
         Begin VB.Label lblstrInscricaoCadastral 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   300
            Width           =   1350
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   49
         Top             =   3600
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   3201
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
         Columns(1).Caption=   "Mês/Ano"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor Declarado"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Valor Calculado"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Data Pagamento"
         Columns(4).DataField=   ""
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3678"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3598"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=3254"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3678"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3598"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=3678"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=3598"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=3678"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=3598"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
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
   End
End
Attribute VB_Name = "frmcadISSQNEstimadoold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtDtmLancamento_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "D", txtdtmLancamento
End Sub

Private Sub txtdtmParcela_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "D", txtdtmParcela
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txtIntexercicio
End Sub

Private Sub txtQtdParcelas_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txtQtdParcelas
End Sub

