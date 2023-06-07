VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadISSQNVariavel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Declaração de ISSQN Variável"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   Icon            =   "CadISSQNVariavel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6840
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   12065
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Controle de Declaração de ISSQN Variável"
      TabPicture(0)   =   "CadISSQNVariavel.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_Lista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Inscricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Endereco"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Emissão de Guias de Arrecadação"
      TabPicture(1)   =   "CadISSQNVariavel.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_EmissaoDeGuias"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5475
         Left            =   -74610
         TabIndex        =   49
         Top             =   780
         Width           =   8025
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1695
            Left            =   540
            TabIndex        =   52
            Top             =   3540
            Width           =   6945
            Begin VB.TextBox txt_Mensagem2 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   33
               Top             =   0
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   34
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbl_Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   53
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1695
            Left            =   540
            TabIndex        =   50
            Top             =   1710
            Width           =   6945
            Begin VB.TextBox txt_Mensagem1 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin VB.CheckBox chk_EmBranco1 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   30
               Top             =   0
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem1 
               Height          =   315
               Left            =   1080
               TabIndex        =   31
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbl_Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   51
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   2250
            MaxLength       =   4
            TabIndex        =   28
            Top             =   1215
            Width           =   525
         End
         Begin VB.TextBox txt_DataDeVencimento 
            Height          =   285
            Left            =   6450
            MaxLength       =   15
            TabIndex        =   29
            Top             =   1215
            Width           =   1035
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoFinal 
            Height          =   315
            Left            =   2250
            TabIndex        =   27
            Top             =   810
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoInicial 
            Height          =   315
            Left            =   2250
            TabIndex        =   26
            Top             =   420
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Final"
            Height          =   195
            Left            =   435
            TabIndex        =   57
            Top             =   900
            Width           =   1725
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Inicial"
            Height          =   195
            Left            =   360
            TabIndex        =   56
            Top             =   510
            Width           =   1800
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   1485
            TabIndex        =   55
            Top             =   1290
            Width           =   675
         End
         Begin VB.Label lbl_alabel 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vencimento"
            Height          =   195
            Left            =   4890
            TabIndex        =   54
            Top             =   1290
            Width           =   1455
         End
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   " Endereço do estabelecimento "
         Height          =   1305
         Left            =   150
         TabIndex        =   18
         Top             =   2970
         Width           =   8475
         Begin VB.TextBox txt_Bairro 
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   540
            Width           =   3705
         End
         Begin VB.TextBox txt_Logradouro 
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   3705
         End
         Begin VB.TextBox txt_Cep 
            Height          =   285
            Left            =   7125
            MaxLength       =   9
            TabIndex        =   23
            Top             =   840
            Width           =   1155
         End
         Begin VB.TextBox txt_UF 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5700
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txt_Complemento 
            Height          =   285
            Left            =   7125
            MaxLength       =   20
            TabIndex        =   21
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox txt_Numero 
            Height          =   285
            Left            =   5700
            MaxLength       =   6
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txt_Municipio 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   840
            Width           =   3705
         End
         Begin VB.Label lblintLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   630
            TabIndex        =   40
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lblintBairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   1035
            TabIndex        =   39
            Top             =   585
            Width           =   405
         End
         Begin VB.Label lblstrMunicipio 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   735
            TabIndex        =   38
            Top             =   885
            Width           =   705
         End
         Begin VB.Label lblintCep 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   6815
            TabIndex        =   37
            Top             =   885
            Width           =   285
         End
         Begin VB.Label lblstrUF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   5400
            TabIndex        =   36
            Top             =   885
            Width           =   210
         End
         Begin VB.Label lblintNumero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5430
            TabIndex        =   25
            Top             =   285
            Width           =   180
         End
         Begin VB.Label lbl_Complemento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6620
            TabIndex        =   24
            Top             =   285
            Width           =   480
         End
      End
      Begin VB.Frame fra_Inscricao 
         Height          =   2385
         Left            =   150
         TabIndex        =   1
         Top             =   480
         Width           =   8475
         Begin MSDataListLib.DataList dlsAtividade 
            Height          =   450
            Left            =   1575
            TabIndex        =   48
            Top             =   1800
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   794
            _Version        =   393216
         End
         Begin MSMask.MaskEdBox mskInscricaoImobiliaria 
            Height          =   285
            Left            =   1575
            TabIndex        =   47
            Top             =   1500
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.TextBox txt_Nome 
            Height          =   285
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   600
            Width           =   4730
         End
         Begin VB.TextBox txt_InscricaoEstadual 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4695
            MaxLength       =   100
            TabIndex        =   41
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
               TabIndex        =   43
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
               TabIndex        =   9
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
         Begin MSDataListLib.DataCombo dbcintEconomico 
            Height          =   315
            Left            =   1575
            TabIndex        =   10
            Top             =   240
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblstrInscricaoImobiliaria 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Imobiliária"
            Height          =   195
            Left            =   60
            TabIndex        =   17
            Top             =   1545
            Width           =   1380
         End
         Begin VB.Label lbl_CNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ / CPF"
            Height          =   195
            Left            =   570
            TabIndex        =   16
            Top             =   1245
            Width           =   870
         End
         Begin VB.Label lblintAtividade 
            AutoSize        =   -1  'True
            Caption         =   "Atividade Empresa"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   1845
            Width           =   1320
         End
         Begin VB.Label lblstrInscricaoEstadual 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Estadual"
            Height          =   195
            Left            =   3255
            TabIndex        =   14
            Top             =   1245
            Width           =   1305
         End
         Begin VB.Label lblstrNomeFantasia 
            AutoSize        =   -1  'True
            Caption         =   "Nome Fantasia"
            Height          =   195
            Left            =   375
            TabIndex        =   13
            Top             =   945
            Width           =   1065
         End
         Begin VB.Label lblintContribuinte 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social"
            Height          =   195
            Left            =   495
            TabIndex        =   12
            Top             =   645
            Width           =   945
         End
         Begin VB.Label lblstrInscricaoCadastral 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   90
            TabIndex        =   11
            Top             =   300
            Width           =   1350
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2295
         Left            =   150
         TabIndex        =   46
         Top             =   4380
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   4048
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
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
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
End
Attribute VB_Name = "frmCadISSQNVariavel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoResultado                As ADODB.Recordset

Private Function strQuery() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strInscricaoCadastral FROM "
    strSql = strSql & gstrEconomico
'    strSql = strSql & " ORDER BY CONVERT(NUMERIC,strInscricaoCadastral)"
    strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoCadastral")
strQuery = strSql
End Function

Private Function VerificaMascaraInscricao() As String
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
  
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
    VerificaMascaraInscricao = strMascara
    
End Function

Private Function strQueryAtividade(lngEconomico As Long) As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String

strSql = ""
strSql = strSql & " SELECT C.PKId, "
'strSql = strSql & " CASE blnPrincipal"
'strSql = strSql & " WHEN 0 then strDescricao"
'strSql = strSql & " ELSE strDescricao + ' - Principal' END As Descricao"
strSql = strSql & gstrCASEWHEN("blnPrincipal", _
                    "0, strDescricao", "strDescricao " & strCONCAT & " ' - Principal'") & " As Descricao"
strSql = strSql & " FROM "
strSql = strSql & gstrAtividadeEC & " A, "
strSql = strSql & gstrAtividadeDaEmpresa & " B, "
strSql = strSql & gstrEconomico & " C"
strSql = strSql & " WHERE "
strSql = strSql & " C.PKId = B.intEconomico"
strSql = strSql & " AND A.PKId = B.intAtividade"
strSql = strSql & " AND C.PKId = " & lngEconomico
strQueryAtividade = strSql
End Function

Private Sub CarregaDadosCadastrais()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
Dim adoRec As ADODB.Recordset

txt_Municipio = gstrCidadeEmpresa

strSql = ""
strSql = strSql & "SELECT CO.PKId, EC.strInscricaoImobiliaria, CO.strNome, CO.bytNaturezaJuridica, CO.strInscricaoEstadual, "
strSql = strSql & "CO.bytNaturezaJuridica, CO.strCNPJCPF , CO.intNumero, "
strSql = strSql & "CO.strNomeFantasia Fantasia, CO.strComplemento, CO.intCEP, CO.intLogradouro, "
strSql = strSql & "BA.strDescricao AS Bairro, "
strSql = strSql & "CI.strDescricao AS Municipio, UF.strSigla AS UF, EC.blnMicroEmpresa "
strSql = strSql & "FROM " & gstrContribuinte & " CO , " & gstrCidade & " CI, "
strSql = strSql & gstrEconomico & " EC,"
strSql = strSql & gstrBairro & " BA, " & gstrUF & " UF "
strSql = strSql & "Where CO.PKId = EC.intContribuinte "
'strSql = strSql & "AND CO.intMunicipio *= CI.PKId "
strSql = strSql & "AND CO.intMunicipio " & strOUTJSQLServer & "= CI.PKId " & strOUTJOracle
'strSql = strSql & "AND CO.intUF *= UF.PKId "
strSql = strSql & "AND CO.intUF " & strOUTJSQLServer & "= UF.PKId " & strOUTJOracle
'strSql = strSql & "AND CO.intBairro *= BA.PKId "
strSql = strSql & "AND CO.intBairro " & strOUTJSQLServer & "= BA.PKId " & strOUTJOracle
strSql = strSql & "AND EC.PKId = " & dbcintEconomico.BoundText

Set gobjBanco = New clsBanco
If gobjBanco.CriaADO(strSql, 4, adoRec) Then
    With adoRec
        If Not .EOF Then
            txt_Nome = gstrENulo(!STRNOME)
            txt_NomeFantasia = gstrENulo(!Fantasia)
            
            If Len(gstrENulo(!StrCnpjCpf)) = 11 Then
                txt_CNPJCPF = Format(!StrCnpjCpf, "@@@.@@@.@@@-@@")
            ElseIf Len(gstrENulo(!StrCnpjCpf)) = 14 Then
                txt_CNPJCPF = Format(!StrCnpjCpf, "@@.@@@.@@@/@@@@-@@")
            Else
                txt_CNPJCPF = ""
            End If
            txt_InscricaoEstadual = gstrENulo(!strInscricaoEstadual)
            mskInscricaoImobiliaria.Text = ""
            mskInscricaoImobiliaria.Mask = VerificaMascaraInscricao
            mskInscricaoImobiliaria = gstrENulo(!strInscricaoImobiliaria)
            opt_NaturezaJuridica(!bytNaturezaJuridica).Value = True
            chkblnMicroEmpresa.Value = Abs(!blnMicroEmpresa)
        End If
        strSql = strQueryLogradouro(!Pkid)
        
    End With
End If
Set gobjBanco = New clsBanco
If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    With adoRec
        txt_Logradouro = gstrENulo(!strTipoTituloLogradouro)
        txt_Numero = gstrENulo(!INTNUMERO)
        txt_Complemento = gstrENulo(!STRCOMPLEMENTO)
        txt_Bairro = gstrENulo(!STRBAIRRO)
        txt_Municipio = gstrENulo(!STRMUNICIPIO)
        txt_UF = gstrENulo(!STRUF)
        txt_Cep = gstrENulo(!INTCEP)
    End With
End If
strSql = strQueryAtividade(dbcintEconomico.BoundText)
LeDaTabelaParaObj "", dlsAtividade, strSql
End Sub

Private Function strQueryLogradouro(lngContribuinte As Long) As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
'            Foi mantida a forma antiga para o SQL Server pois não era possível o
'            deslocamento completo devido à incompatibilidade entres os bancos.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String

strSql = ""
'strSql = strSql & " SELECT A.strCNPJCPF, ISNULL(C.strDescricao + ' : ' ,'') + ISNULL(D.strDescricao + ' ' ,'')  + "
strSql = strSql & " SELECT A.strCNPJCPF, " & gstrISNULL("C.strDescricao", "''", "C.strDescricao" & strCONCAT & " ' : '") & strCONCAT & gstrISNULL("D.strDescricao " & strCONCAT & " ' '", "''") & strCONCAT
'strSql = strSql & " ISNULL(B.strDescricao ,'') AS strTipoTituloLogradouro, A.intNumero, A.strComplemento, "
strSql = strSql & gstrISNULL("B.strDescricao", "''") & " AS strTipoTituloLogradouro, A.intNumero, A.strComplemento, "
strSql = strSql & "E.strDescricao AS strBairro, F.strDescricao AS strMunicipio, G.strSigla AS strUF, A.intCEP FROM "
strSql = strSql & gstrContribuinte & " A "
If (bytDBType = EDatabases.SQLServer) Then
    strSql = strSql & "LEFT JOIN " & gstrLogradouro & " B ON A.intLogradouro = B.PKId "
    strSql = strSql & "LEFT JOIN " & gstrTipoLogradouro & " C ON B.intTipoLogradouro = C.PKId "
    strSql = strSql & "LEFT JOIN " & gstrTituloLogradouro & " D ON B.intTituloLogradouro = D.PKId "
    strSql = strSql & "LEFT JOIN " & gstrBairro & " E ON A.intBairro = E.PKId "
    strSql = strSql & "LEFT JOIN " & gstrCidade & " F ON A.intMunicipio = F.PKId "
    strSql = strSql & "LEFT JOIN " & gstrUF & " G ON A.intUF = G.PKId "

ElseIf (bytDBType = EDatabases.Oracle) Then
    strSql = strSql & ", " & gstrLogradouro & " B"
    strSql = strSql & ", " & gstrTipoLogradouro & " C"
    strSql = strSql & ", " & gstrTituloLogradouro & " D"
    strSql = strSql & ", " & gstrBairro & " E"
    strSql = strSql & ", " & gstrCidade & " F"
    strSql = strSql & ", " & gstrUF & " G "

End If

strSql = strSql & " WHERE A.PKId = " & lngContribuinte

If (bytDBType = EDatabases.Oracle) Then
    strSql = strSql & " AND A.intLogradouro = B.PKId " & strOUTJOracle
    strSql = strSql & " AND B.intTipoLogradouro = C.PKId " & strOUTJOracle
    strSql = strSql & " AND B.intTituloLogradouro = D.PKId " & strOUTJOracle
    strSql = strSql & " AND A.intBairro = E.PKId " & strOUTJOracle
    strSql = strSql & " AND A.intMunicipio = F.PKId " & strOUTJOracle
    strSql = strSql & " AND A.intUF = G.PKId "

End If

strSql = strSql & " ORDER BY A.PKId "

strQueryLogradouro = strSql
End Function

Private Sub dbc_intMensagem1_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intMensagem1, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intMensagem2_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intMensagem2, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoFinal_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoFinal, Me, Area
End Sub

Private Sub dbc_strInscricaoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoInicial_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoInicial, Me, Area
End Sub

Private Sub dbc_strInscricaoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoInicial, Me, , KeyCode, Shift
End Sub

Private Sub dbcintEconomico_Click(Area As Integer)
    DropDownDataCombo dbcintEconomico, Me, Area
    If Area = 2 Then
        If dbcintEconomico.MatchedWithList Then
            CarregaDadosCadastrais
        End If
    End If
End Sub

Private Sub dbcintEconomico_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintEconomico, Me, , KeyCode, Shift
End Sub

Private Sub Form_Load()
Dim strSql As String

TrocaCorObjeto txt_Nome, True
TrocaCorObjeto txt_NomeFantasia, True
TrocaCorObjeto txt_CNPJCPF, True
TrocaCorObjeto txt_InscricaoEstadual, True
TrocaCorObjeto mskInscricaoImobiliaria, True
'TrocaCorObjeto dlsAtividade, True
TrocaCorObjeto txt_Logradouro, True
TrocaCorObjeto txt_Complemento, True
TrocaCorObjeto txt_Bairro, True
TrocaCorObjeto txt_Municipio, True
TrocaCorObjeto txt_UF, True
TrocaCorObjeto txt_Numero, True
TrocaCorObjeto txt_Cep, True

strSql = strQuery
dbcintEconomico.Tag = strSql & ";strInscricaoCadastral"

'''GUIA
    dbc_strInscricaoInicial.Tag = strQueryInscricao & ";EC.strInscricaoCadastral"
    dbc_strInscricaoFinal.Tag = strQueryInscricao & ";EC.strInscricaoCadastral"
    dbc_intMensagem1.Tag = strQueryMensagem & ";strDescricao"
    dbc_intMensagem2.Tag = strQueryMensagem & ";strDescricao"
End Sub


Public Sub MantemForm(strModoOperacao As String)
    If UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
        Exit Sub
    End If
    If tab_3DPasta.Tab = 2 Then
    
        If UCase(strModoOperacao) = UCase(gstrImprimir) Then
            If blnDadosGuiaOK = True Then
                Set gfrmFormularioQueEstaImprimindoGuia = Me
                'ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, gstrQuerryRelatorioGuiaDeArrecadacao(dbc_strInscricaoInicial.Text, dbc_strInscricaoFinal.Text, Val(txt_intExercicio.Text), txt_DataDeVencimento.Text)
                'Olhar com Renato
                'ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, gstrQuerryRelatorioGuiaDeArrecadacao(dbc_strInscricaoInicial.Text, dbc_strInscricaoFinal.Text, txt_intExercicio.Text, dbc_intComposicaoDaReceita.BoundText, , txt_DataDeVencimento.Text)
            End If
        End If
        If UCase(strModoOperacao) = UCase(gstrNovo) Then
            LimpaObjetos
        End If
        If UCase(strModoOperacao) = UCase(gstrFechar) Then
            Unload Me
        End If
        
    Else
    
        'Configurar eventos mantem form.....estava vazio
        Exit Sub
        
    End If
End Sub






'''>>>>>>>>>>>>>>>>>>GUIA DE ARRECADAÇÃO





Private Function strQueryInscricao() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT EC.PKId, EC.strInscricaoCadastral "
    strSql = strSql & " FROM "
    strSql = strSql & gstrEconomico & " EC,"
    strSql = strSql & gstrTributoEmpresa & " EM"
    strSql = strSql & " WHERE EC.PKId = EM.intEconomico AND EC.dtmDataBaixa IS NULL "
'    strSql = strSql & " ORDER BY CONVERT(NUMERIC, EC.strInscricaoCadastral) "
    strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "EC.strInscricaoCadastral")
strQueryInscricao = strSql
End Function

Private Sub chk_EmBranco1_Click()
    If chk_EmBranco1.Value = 1 Then
        dbc_intMensagem1.BoundText = ""
        dbc_intMensagem1.Enabled = False
        TrocaCorObjeto dbc_intMensagem1, True
        txt_Mensagem1.Text = ""
        txt_Mensagem1.Enabled = False
        TrocaCorObjeto txt_Mensagem1, True
    Else
        dbc_intMensagem1.Enabled = True
        TrocaCorObjeto dbc_intMensagem1, False
        txt_Mensagem1.Enabled = True
        TrocaCorObjeto txt_Mensagem1, False
    End If
End Sub

Private Sub chk_EmBranco2_Click()
    If chk_EmBranco2.Value = 1 Then
        dbc_intMensagem2.BoundText = ""
        dbc_intMensagem2.Enabled = False
        TrocaCorObjeto dbc_intMensagem2, True
        txt_Mensagem2.Text = ""
        txt_Mensagem2.Enabled = False
        TrocaCorObjeto txt_Mensagem2, True
    Else
        dbc_intMensagem2.Enabled = True
        TrocaCorObjeto dbc_intMensagem2, False
        txt_Mensagem2.Enabled = True
        TrocaCorObjeto txt_Mensagem2, False
    End If
End Sub

Private Sub dbc_intMensagem1_Click(Area As Integer)
    DropDownDataCombo dbc_intMensagem1, Me, Area
    If Area = 2 Then
        LeDoComboParaTXT1
    End If
End Sub

Private Sub dbc_intMensagem2_Click(Area As Integer)
    DropDownDataCombo dbc_intMensagem2, Me, Area
    If Area = 2 Then
        LeDoComboParaTXT2
    End If
End Sub

Private Function LeDoComboParaTXT1()
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT strMensagem "
    strSql = strSql & " FROM " & gstrMensagem
    strSql = strSql & " WHERE PKId = " & Val(dbc_intMensagem1.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem1.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem1.Text = ""
        End If
    End If
End Function

Private Function LeDoComboParaTXT2()
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT strMensagem "
    strSql = strSql & " FROM " & gstrMensagem
    strSql = strSql & " WHERE PKId = " & Val(dbc_intMensagem2.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem2.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem2.Text = ""
        End If
    End If
End Function

Private Function strQueryMensagem() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSql As String

    strSql = ""
'    strSQL = strSQL & "SELECT PKId, ltrim(rtrim(PKId)) + ' - ' + ltrim(rtrim(strDescricao)) as Descricao "
    strSql = strSql & "SELECT PKId, ltrim(rtrim(PKId)) " & strCONCAT & " ' - ' " & strCONCAT & " ltrim(rtrim(strDescricao)) as Descricao "
    strSql = strSql & " FROM " & gstrMensagem
    strSql = strSql & " ORDER BY PKId "

strQueryMensagem = strSql
End Function

Private Function blnDadosGuiaOK() As Boolean
blnDadosGuiaOK = False

    If dbc_strInscricaoInicial.BoundText = "" Then
        ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para gerar a Guia de Arrecadação."
        dbc_strInscricaoInicial.SetFocus
        Exit Function
    End If
    
    If dbc_strInscricaoFinal.BoundText = "" Then
        ExibeMensagem "Selecione uma Inscrição Cadastral Final para gerar a Guia de Arrecadação."
        dbc_strInscricaoFinal.SetFocus
        Exit Function
    End If
    
    If txt_intExercicio.Text = "" Then
        ExibeMensagem "O Exercício deve ser Digitado."
        txt_intExercicio.SetFocus
        Exit Function
    End If
    
    If txt_DataDeVencimento.Text = "" Then
        ExibeMensagem "A data de vencimento deve ser digitada."
        txt_DataDeVencimento.SetFocus
        Exit Function
    ElseIf gblnDataValida(txt_DataDeVencimento.Text) = False Then
        ExibeMensagem "Data de vencimento inválida."
        txt_DataDeVencimento.SetFocus
        Exit Function
    End If
    
    If chk_EmBranco1.Value = 0 Then
        If txt_Mensagem1.Text = "" Then
            ExibeMensagem "A mensagem 1 tem que ser selecionada."
            Exit Function
        End If
    End If
    
    If chk_EmBranco2.Value = 0 Then
        If txt_Mensagem2.Text = "" Then
            ExibeMensagem "A mensagem 2 tem que ser selecionada."
            Exit Function
        End If
    End If
    
blnDadosGuiaOK = True
End Function

Private Sub LimpaObjetos()
    dbc_strInscricaoInicial.BoundText = ""
    dbc_strInscricaoFinal.BoundText = ""
    txt_intExercicio.Text = ""
    txt_DataDeVencimento.Text = ""
    dbc_intMensagem1.BoundText = ""
    dbc_intMensagem2.BoundText = ""
    txt_Mensagem1 = ""
    txt_Mensagem2 = ""
    dbc_strInscricaoInicial.SetFocus
End Sub

Private Sub dbc_intMensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem1
End Sub

Private Sub dbc_intMensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem2
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    Select Case tab_3DPasta.Tab
        Case 0
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
        Case 1
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
        Case 2
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrNovo
    End Select
End Sub


Private Sub txt_DataDeVencimento_GotFocus()
    MarcaCampo txt_DataDeVencimento
End Sub

Private Sub txt_DataDeVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataDeVencimento
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub dbc_strInscricaoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoInicial
End Sub

Private Sub dbc_strInscricaoFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoFinal
End Sub

