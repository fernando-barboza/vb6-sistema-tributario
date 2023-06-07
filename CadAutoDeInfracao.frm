VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadAutoDeInfracao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autos de Infração"
   ClientHeight    =   6750
   ClientLeft      =   885
   ClientTop       =   1950
   ClientWidth     =   8610
   Icon            =   "CadAutoDeInfracao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8610
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6660
      Left            =   60
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   30
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   11748
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Autos de Infração"
      TabPicture(0)   =   "CadAutoDeInfracao.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintOrdermServico"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPKId"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrInscricaoCadastral"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintComposicaoDaReceita"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_Fiscais"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblstrDescricao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblintOcorrencia"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbcintOcorrencia"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tdb_AutoDeInfracao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dbcintComposicaoDaReceita"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbcintOrdermServico"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPKId"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fra_Inscricao"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fra_Endereco"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dbl_Fiscais"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtstrDescricao"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_strInscricaoCadastral"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtintContribuinte"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtstrInscricaoCadastral"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Taxas "
      TabPicture(1)   =   "CadAutoDeInfracao.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdb_Taxas"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_Frame"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Emissão de Guias de Arrecadação "
      TabPicture(2)   =   "CadAutoDeInfracao.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_EmissaoDeGuias"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtstrInscricaoCadastral 
         Height          =   300
         Left            =   3570
         TabIndex        =   5
         Top             =   1050
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.TextBox txtintContribuinte 
         Height          =   300
         Left            =   2940
         TabIndex        =   4
         Top             =   1050
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txt_strInscricaoCadastral 
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
         Left            =   2250
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2745
         Width           =   5925
      End
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5445
         Left            =   -74730
         TabIndex        =   63
         Top             =   690
         Width           =   7935
         Begin VB.TextBox txt_DataDeVencimento 
            Height          =   285
            Left            =   6420
            MaxLength       =   15
            TabIndex        =   31
            Top             =   1215
            Width           =   1035
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   2220
            MaxLength       =   4
            TabIndex        =   30
            Top             =   1215
            Width           =   525
         End
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1695
            Left            =   510
            TabIndex        =   66
            Top             =   1710
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco1 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   32
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem1 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem1 
               Height          =   315
               Left            =   1080
               TabIndex        =   33
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
               TabIndex        =   67
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1695
            Left            =   510
            TabIndex        =   64
            Top             =   3540
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   35
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem2 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   36
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
               TabIndex        =   65
               Top             =   390
               Width           =   780
            End
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoFinal 
            Height          =   315
            Left            =   2220
            TabIndex        =   29
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
            Left            =   2220
            TabIndex        =   28
            Top             =   420
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_DataDeVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vencimento"
            Height          =   195
            Left            =   4860
            TabIndex        =   71
            Top             =   1290
            Width           =   1455
         End
         Begin VB.Label lbl_intExercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   1455
            TabIndex        =   70
            Top             =   1290
            Width           =   675
         End
         Begin VB.Label lblstrInscricaoInicial 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Inicial"
            Height          =   195
            Left            =   330
            TabIndex        =   69
            Top             =   510
            Width           =   1800
         End
         Begin VB.Label lblstrInscricaoFinal 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Final"
            Height          =   195
            Left            =   405
            TabIndex        =   68
            Top             =   900
            Width           =   1725
         End
      End
      Begin VB.TextBox txtstrDescricao 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   450
         Left            =   2250
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1395
         Width           =   5925
      End
      Begin MSDataListLib.DataList dbl_Fiscais 
         Height          =   450
         Left            =   2250
         TabIndex        =   8
         Top             =   2250
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   794
         _Version        =   393216
      End
      Begin VB.Frame fra_Frame 
         Height          =   1560
         Left            =   -74805
         TabIndex        =   53
         Top             =   630
         Width           =   7995
         Begin VB.TextBox txtdtmLancamento 
            Height          =   285
            Left            =   2010
            MaxLength       =   15
            TabIndex        =   23
            Top             =   675
            Width           =   1035
         End
         Begin VB.TextBox txtdtmVencimento 
            Height          =   285
            Left            =   6195
            MaxLength       =   15
            TabIndex        =   24
            Top             =   675
            Width           =   1005
         End
         Begin VB.TextBox txtintIntervaloParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6645
            MaxLength       =   15
            TabIndex        =   26
            Top             =   1035
            Width           =   555
         End
         Begin VB.TextBox txtdblValorParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2010
            MaxLength       =   15
            TabIndex        =   21
            Top             =   300
            Width           =   1320
         End
         Begin VB.TextBox txtintExercicio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6465
            MaxLength       =   4
            TabIndex        =   22
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtintNumeroDeParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2010
            MaxLength       =   3
            TabIndex        =   25
            Top             =   1035
            Width           =   1005
         End
         Begin VB.Label lbl_dias 
            AutoSize        =   -1  'True
            Caption         =   "dias."
            Height          =   195
            Left            =   7290
            TabIndex        =   60
            Top             =   1110
            Width           =   330
         End
         Begin VB.Label lblintIntervaloParcela 
            AutoSize        =   -1  'True
            Caption         =   "Intervalo entre Parcelas"
            Height          =   195
            Left            =   4815
            TabIndex        =   59
            Top             =   1110
            Width           =   1680
         End
         Begin VB.Label lbldblValorParcela 
            AutoSize        =   -1  'True
            Caption         =   "Valor Parcela"
            Height          =   195
            Left            =   930
            TabIndex        =   58
            Top             =   375
            Width           =   945
         End
         Begin VB.Label lblintExercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   5640
            TabIndex        =   57
            Top             =   375
            Width           =   675
         End
         Begin VB.Label lblintNumeroDeParcela 
            AutoSize        =   -1  'True
            Caption         =   "Número de Parcelas"
            Height          =   195
            Left            =   435
            TabIndex        =   56
            Top             =   1110
            Width           =   1440
         End
         Begin VB.Label lbldtmVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vencimento"
            Height          =   195
            Left            =   4590
            TabIndex        =   55
            Top             =   735
            Width           =   1455
         End
         Begin VB.Label lbldtmLancamento 
            AutoSize        =   -1  'True
            Caption         =   "Data de Lançamento"
            Height          =   195
            Left            =   375
            TabIndex        =   54
            Top             =   735
            Width           =   1500
         End
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   "Endereço"
         Height          =   1335
         Left            =   330
         TabIndex        =   44
         Top             =   3105
         Width           =   7845
         Begin VB.TextBox txt_UF 
            Height          =   285
            Left            =   5070
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   930
            Width           =   510
         End
         Begin VB.TextBox txt_Municipio 
            Height          =   285
            Left            =   5070
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   600
            Width           =   2625
         End
         Begin VB.TextBox txt_Cep 
            Height          =   285
            Left            =   6615
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   930
            Width           =   1080
         End
         Begin VB.TextBox txt_Complemento 
            Height          =   285
            Left            =   6825
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   270
            Width           =   870
         End
         Begin VB.TextBox txt_Numero 
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   270
            Width           =   795
         End
         Begin VB.TextBox txt_Logradouro 
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   270
            Width           =   4005
         End
         Begin VB.TextBox txt_Bairro 
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   600
            Width           =   3105
         End
         Begin VB.TextBox txt_Distrito 
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   930
            Width           =   3525
         End
         Begin VB.Label lbl_strDistrito 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Distrito"
            Height          =   195
            Left            =   450
            TabIndex        =   52
            Top             =   990
            Width           =   480
         End
         Begin VB.Label lbl_intMunicipio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   4290
            TabIndex        =   51
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lbl_intBairro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   525
            TabIndex        =   50
            Top             =   660
            Width           =   405
         End
         Begin VB.Label lbl_intLogradouro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lbl_intNumero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5145
            TabIndex        =   48
            Top             =   330
            Width           =   180
         End
         Begin VB.Label lbl_strComplemento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6270
            TabIndex        =   47
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl_intUF 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   4770
            TabIndex        =   46
            Top             =   1005
            Width           =   210
         End
         Begin VB.Label lbl_intCep 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   6240
            TabIndex        =   45
            Top             =   990
            Width           =   285
         End
      End
      Begin VB.Frame fra_Inscricao 
         Height          =   630
         Left            =   2250
         TabIndex        =   39
         Top             =   360
         Width           =   5925
         Begin VB.OptionButton optbytTipoDeInscricaoCadastral 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   0
            Top             =   270
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optbytTipoDeInscricaoCadastral 
            Caption         =   "Imobiliário Rural"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   1
            Left            =   2400
            TabIndex        =   1
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton optbytTipoDeInscricaoCadastral 
            Caption         =   "Econômico"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   2
            Left            =   4545
            TabIndex        =   2
            Top             =   270
            Width           =   1155
         End
      End
      Begin VB.TextBox txtPKId 
         Height          =   300
         Left            =   2250
         TabIndex        =   3
         Top             =   1050
         Width           =   555
      End
      Begin MSDataListLib.DataCombo dbcintOrdermServico 
         Height          =   315
         Left            =   2250
         TabIndex        =   7
         Top             =   1890
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintComposicaoDaReceita 
         Height          =   315
         Left            =   2250
         TabIndex        =   18
         Top             =   4545
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Taxas 
         Height          =   3810
         Left            =   -74790
         TabIndex        =   27
         Top             =   2565
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   6720
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   68
         Columns(1)._MaxComboItems=   20
         Columns(1).ValueItems(0)._DefaultItem=   0
         Columns(1).ValueItems(0).Value=   ""
         Columns(1).ValueItems(0).Value.vt=   8
         Columns(1).ValueItems(0).DisplayValue=   ""
         Columns(1).ValueItems(0).DisplayValue.vt=   8
         Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(1).ValueItems.Count=   1
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
         Columns(2).DropDown=   "tdd_Atividades"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   1
         Splits(0).MarqueeStyle=   5
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=529"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=450"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=12965"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=12885"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(2).AutoDropDown=1"
         Splits(0)._ColumnProps(20)=   "Column(2).DropDownList=1"
         Splits(0)._ColumnProps(21)=   "Column(2).AutoCompletion=1"
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
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   0
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
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_AutoDeInfracao 
         Height          =   1095
         Left            =   330
         TabIndex        =   20
         Top             =   5340
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   1931
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nº do Auto"
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Inscrição Cadastral"
         Columns(1).DataField=   "strInscricaoCadastral"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Proprietário"
         Columns(2).DataField=   "strNome"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3466"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3387"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=7752"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=7673"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintOcorrencia 
         Height          =   315
         Left            =   2250
         TabIndex        =   19
         Top             =   4905
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblintOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   1335
         TabIndex        =   72
         Top             =   4995
         Width           =   780
      End
      Begin VB.Label lblstrDescricao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1380
         TabIndex        =   62
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lbl_Fiscais 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fiscais"
         Height          =   195
         Left            =   1635
         TabIndex        =   61
         Top             =   2385
         Width           =   480
      End
      Begin VB.Label lblintComposicaoDaReceita 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   420
         TabIndex        =   43
         Top             =   4620
         Width           =   1695
      End
      Begin VB.Label lblstrInscricaoCadastral 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   765
         TabIndex        =   42
         Top             =   2790
         Width           =   1350
      End
      Begin VB.Label lblPKId 
         AutoSize        =   -1  'True
         Caption         =   "Nº do Auto"
         Height          =   195
         Left            =   1335
         TabIndex        =   41
         Top             =   1125
         Width           =   780
      End
      Begin VB.Label lblintOrdermServico 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nº da O.S."
         Height          =   195
         Left            =   1335
         TabIndex        =   40
         Top             =   1980
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmCadAutoDeInfracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xTaxa                     As XArrayDB
Dim blnPrimeiraVez            As Boolean
Dim blnRowColChange           As Boolean
Dim mobjAux                   As Object

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

Private Sub dbcintComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrencia_Click(Area As Integer)
    DropDownDataCombo dbcintOcorrencia, Me, Area
End Sub

Private Sub dbcintOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOrdermServico_Click(Area As Integer)
    Dim intIndice As Integer
    
    DropDownDataCombo dbcintOrdermServico, Me, Area
    
    If Area = 2 And dbcintOrdermServico.MatchedWithList Then
        For intIndice = 0 To 2
            If optbytTipoDeInscricaoCadastral(intIndice).Value Then
                Exit For
            End If
        Next
        BuscaInscricaoCadastral (intIndice)
        LeDaTabelaParaObj "", dbl_Fiscais, strQueryFiscais
        BuscaDadosProprietario (intIndice)
    End If
    
End Sub

Private Sub BuscaInscricaoCadastral(intIndice As Integer)
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strQueryInscricao(intIndice), 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txtstrInscricaoCadastral.Text = !strInscricaoCadastral
                txt_strInscricaoCadastral.Text = !Descricao
            End With
        End If
    End If
    
End Sub

Private Sub dbcintOrdermServico_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintOrdermServico, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 651
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrSalvar
End Sub

Private Sub Form_Load()
    tab_3dPasta.Tab = 0
    optbytTipoDeInscricaoCadastral_Click (0)
    TrocaCorObjeto txtPKId, True
    TrocaCorObjeto txt_strInscricaoCadastral, True
    TrocaCorObjeto txt_Logradouro, True
    TrocaCorObjeto txt_Numero, True
    TrocaCorObjeto txt_Complemento, True
    TrocaCorObjeto txt_Bairro, True
    TrocaCorObjeto txt_Municipio, True
    TrocaCorObjeto txt_Distrito, True
    TrocaCorObjeto txt_UF, True
    TrocaCorObjeto txt_Cep, True
    
    'LeDaTabelaParaObj "", tdb_AutoDeInfracao, strQueryAutoDeInfracao
    dbcintOcorrencia.Tag = strQuerryOcorrencia & ";strDescricao"
    
    '''GUIA
    dbc_strInscricaoInicial.Tag = strQueryInscricaoGuia & ";strInscricaoCadastral"
    dbc_strInscricaoFinal.Tag = strQueryInscricaoGuia & ";strInscricaoCadastral"
    dbc_intMensagem1.Tag = strQueryMensagem & ";strDescricao"
    dbc_intMensagem2.Tag = strQueryMensagem & ";strDescricao"
    
End Sub

Private Function strQuerryOcorrencia() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrOcorrencia
    strSql = strSql & " WHERE "
    strSql = strSql & " intUtilizacaoDaOcorrencia = 1 "
    strSql = strSql & " ORDER BY strDescricao "
    strQuerryOcorrencia = strSql
End Function

Private Function strQueryInscricaoGuia() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String

    strSql = ""
    strSql = strSql & " SELECT PKId, strInscricaoCadastral "
    strSql = strSql & " FROM "
    strSql = strSql & gstrEconomico
    strSql = strSql & " WHERE "
    strSql = strSql & " dtmDataBaixa IS NULL " 'Verifica se existe data de baixa
    strSql = strSql & " ORDER BY "
'    strSql = strSql & " CONVERT(NUMERIC,strInscricaoCadastral) "
    strSql = strSql & gstrCONVERT(CDT_NUMERIC, "strInscricaoCadastral")
    strSql = strSql

    strQueryInscricaoGuia = strSql
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
'    strSql = strSql & "SELECT PKId, ltrim(rtrim(PKId)) + ' - ' + ltrim(rtrim(strDescricao)) as Descricao "
    strSql = strSql & "SELECT PKId, ltrim(rtrim(PKId)) " & strCONCAT & " ' - ' " & strCONCAT & " ltrim(rtrim(strDescricao)) as Descricao "
    strSql = strSql & " FROM " & gstrMensagem
    strSql = strSql & " ORDER BY PKId "

    strQueryMensagem = strSql
End Function

Private Function strQueryAutoDeInfracao() As String
    Dim strSql  As String

    strSql = ""

    strSql = strSql & " SELECT AUT.PKId, AUT.strInscricaoCadastral, CON.strNome "

    strSql = strSql & " FROM "
    strSql = strSql & gstrAutoDeInfracao & " AUT, "
    strSql = strSql & gstrContribuinte & " CON "

    strSql = strSql & " WHERE "
    strSql = strSql & " CON.PKId = AUT.intContribuinte "
    strSql = strSql & " ORDER BY AUT.PKId "
    strQueryAutoDeInfracao = strSql
    
End Function

Private Function strQueryDataComboOrdemServico(intIndex As Integer) As String

'******************************************************************************************
' Data: 12/05/2003
' Alteração: - Inseridos apelidos nas colunas relacionadas na cláusula SELECT e alterada a
'            cláusula ORDER BY de modo que esta utilizá-se os apelidos.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
'    strSql = strSql & "SELECT PKId, PKId "
    strSql = strSql & "SELECT PKId PKId1, PKId PKId2 "
    strSql = strSql & "FROM " & gstrOrdemServico & " "
    strSql = strSql & "WHERE bytOrigem = " & intIndex & " "
'    strSql = strSql & "ORDER BY PKId "
    strSql = strSql & "ORDER BY PKId1 "
    strQueryDataComboOrdemServico = strSql
End Function

Private Sub optbytTipoDeInscricaoCadastral_Click(Index As Integer)
    Dim strSql As String
    Dim intIndice As Integer
    
    optbytTipoDeInscricaoCadastral(Index).CausesValidation = True

    For intIndice = 0 To 2
        If intIndice <> Index Then
            optbytTipoDeInscricaoCadastral(intIndice).CausesValidation = False
        End If
    Next
    
    If Not blnRowColChange Then
        blnPrimeiraVez = False
        LimpaObjeto Me
    End If
    
    txt_strInscricaoCadastral.Text = ""
    
    Set dbl_Fiscais.RowSource = Nothing
    dbl_Fiscais.BoundText = 0
    
    LimpaEndereco
    
    Set xTaxa = New XArrayDB
    xTaxa.Clear
    xTaxa.ReDim 0, 0, 0, 2
    
    Set tdb_Taxas.Array = xTaxa
    tdb_Taxas.Rebind
    tdb_Taxas.Refresh
    
    dbcintComposicaoDaReceita.BoundText = ""
    dbcintOrdermServico.BoundText = ""
End Sub

Private Function strQuerryComposicao(Index As Integer) As String
    Dim strSql As String
    Dim Utilizacao As Integer
    
    Utilizacao = 0
    If Index = 0 Or Index = 1 Then
        Utilizacao = 1
    ElseIf Index = 2 Then
        Utilizacao = 2
    End If
    
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao "
    
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita
    
    strSql = strSql & " WHERE "
    strSql = strSql & " intUtilizacao = " & Utilizacao
    
    strSql = strSql & " ORDER BY strDescricao "
    strQuerryComposicao = strSql
End Function

Private Function strQueryFiscais() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT OSF.PKId, CON.strNome "

    strSql = strSql & " FROM "
    strSql = strSql & gstrFiscais & " FIS, "
    strSql = strSql & gstrOrdemServicoFiscal & " OSF, "
    strSql = strSql & gstrOrdemServico & " OS, "
    
    strSql = strSql & gstrContribuinte & " CON "
    
    strSql = strSql & " WHERE "
    strSql = strSql & " OS.PKID = " & dbcintOrdermServico.BoundText
    strSql = strSql & " AND OSF.intOrdemServico = OS.PKId "
    strSql = strSql & " AND FIS.PKId = OSF.intFiscal "
    strSql = strSql & " AND CON.PKId = FIS.intContribuinte "
    strSql = strSql & " ORDER BY strNome "

    strQueryFiscais = strSql
End Function

Private Sub BuscaDadosProprietario(intIndice As Integer)

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql          As String
    Dim strTabelaInscricao As String
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT CO.PKId, LG.strDescricao AS Logradouro, CO.intNumero, "
    strSql = strSql & " CO.strComplemento , CO.intCep, BA.strDescricao AS Bairro, UF.strSigla AS UF, "
    strSql = strSql & " CO.strDistritoC, CI.strDescricao AS Municipio "
    strSql = strSql & " FROM "
    
    If intIndice = 0 Then
        strTabelaInscricao = gstrImobiliario
    ElseIf intIndice = 1 Then
        strTabelaInscricao = gstrImobiliarioRural
    ElseIf intIndice = 2 Then
        strTabelaInscricao = gstrEconomico
    End If
    
'    strSQL = strSQL & strTabelaInscricao & " AS INS, "
    strSql = strSql & strTabelaInscricao & " INS, "
'    strSQL = strSQL & gstrOrdemServico & " AS OS, "
    strSql = strSql & gstrOrdemServico & " OS, "
'    strSQL = strSQL & gstrContribuinte & " AS CO, "
    strSql = strSql & gstrContribuinte & " CO, "
'    strSQL = strSQL & gstrLogradouro & " AS LG, "
    strSql = strSql & gstrLogradouro & " LG, "
'    strSQL = strSQL & gstrCidade & " AS CI, "
    strSql = strSql & gstrCidade & " CI, "
'    strSQL = strSQL & gstrBairro & " AS BA, "
    strSql = strSql & gstrBairro & " BA, "
'    strSQL = strSQL & gstrUF & " AS UF "
    strSql = strSql & gstrUF & " UF "
    strSql = strSql & " WHERE "
    strSql = strSql & " OS.PKId = " & dbcintOrdermServico.BoundText
    strSql = strSql & " AND INS.PKId = OS.intInscricaoCadastral "
    strSql = strSql & " AND CO.PKId = INS.intContribuinte "
    strSql = strSql & " AND CI.PKId = CO.intMunicipio "
    strSql = strSql & " AND LG.PKId = CO.intLogradouro "
    strSql = strSql & " AND BA.PKId = CO.intBairro "
    strSql = strSql & " AND UF.PKId = CO.intUF "
    
    LimpaEndereco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txtintContribuinte.Text = !Pkid
                txt_Logradouro.Text = !Logradouro
                txt_Numero.Text = !intNumero
                txt_Complemento.Text = gstrENulo(!strComplemento)
                txt_Bairro.Text = gstrENulo(!Bairro)
                txt_Municipio.Text = gstrENulo(!Municipio)
                txt_Distrito.Text = gstrENulo(!strDistritoC)
                txt_UF.Text = gstrENulo(!UF)
                txt_Cep.Text = gstrCEPFormatado(!intCep)
            End With
        End If
    End If
    
End Sub

Private Function LimpaEndereco()
    txt_Logradouro.Text = ""
    txt_Numero.Text = ""
    txt_Complemento.Text = ""
    txt_Bairro.Text = ""
    txt_Municipio.Text = ""
    txt_Distrito.Text = ""
    txt_UF.Text = ""
    txt_Cep.Text = ""
End Function

Private Function strQueryInscricao(Index As Integer) As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    
    strSql = ""
    If Index = 0 Or Index = 1 Then
'        strSql = strSql & " SELECT A.strInscricaoAnterior AS strInscricaoCadastral, LTRIM(RTRIM(A.strInscricaoAnterior)) + ' - ' +  LTRIM(RTRIM(C.strNome)) AS Descricao "
        strSql = strSql & " SELECT A.strInscricaoAnterior AS strInscricaoCadastral, LTRIM(RTRIM(A.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(C.strNome)) AS Descricao "
    ElseIf Index = 2 Then
'        strSql = strSql & " SELECT A.strInscricaoCadastral, LTRIM(RTRIM(A.strInscricaoCadastral)) + ' - ' +  LTRIM(RTRIM(C.strNome)) AS Descricao "
        strSql = strSql & " SELECT A.strInscricaoCadastral, LTRIM(RTRIM(A.strInscricaoCadastral)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(C.strNome)) AS Descricao "
    End If

    strSql = strSql & " FROM "

    If Index = 0 Then
        strSql = strSql & gstrImobiliario & " A, "
        strSql = strSql & gstrOrdemServico & " B, "
        strSql = strSql & gstrContribuinte & " C "
    ElseIf Index = 1 Then
        strSql = strSql & gstrImobiliarioRural & " A, "
        strSql = strSql & gstrOrdemServico & " B, "
        strSql = strSql & gstrContribuinte & " C "
    ElseIf Index = 2 Then
        strSql = strSql & gstrEconomico & " A, "
        strSql = strSql & gstrOrdemServico & " B, "
        strSql = strSql & gstrContribuinte & " C "
    End If
    
    strSql = strSql & " WHERE "
    strSql = strSql & " A.PKId = B.intInscricaoCadastral "
    strSql = strSql & " AND B.PKId = " & dbcintOrdermServico.BoundText
    strSql = strSql & " AND C.PKId = A.intContribuinte "
    strSql = strSql & " ORDER BY Descricao "

    strQueryInscricao = strSql
End Function

Private Sub dbcintComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbcintComposicaoDaReceita, Me, Area
    If Area = 2 Then
        MontaAtividade dbcintComposicaoDaReceita.BoundText
    End If
End Sub

Private Sub MontaAtividade(intComposicaoReceita As Integer)
    Dim strSql As String
    Dim adoRec As ADODB.Recordset
    Dim varAux As String
    
    On Error GoTo Err_Handle
    
    Set xTaxa = New XArrayDB
    xTaxa.Clear
    
    xTaxa.ReDim 0, 0, 0, 2
    
    strSql = ""
    strSql = strSql & " SELECT A.PKId, A.strDescricao FROM "
    strSql = strSql & gstrReceita & " A,"
    strSql = strSql & gstrValorCompoRec & " B"
    strSql = strSql & " WHERE A.PKId = B.intReceita "
    strSql = strSql & " AND B.intComposicaoDaReceita = " & intComposicaoReceita
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        With adoRec
            If Not .EOF Then
                xTaxa.ReDim 0, .RecordCount - 1, 0, 2
                Do While Not .EOF
                    varAux = !Pkid
                    xTaxa(.AbsolutePosition - 1, 0) = varAux
                    
                    varAux = False
                    xTaxa(.AbsolutePosition - 1, 1) = varAux
                
                    varAux = !strDescricao
                    xTaxa(.AbsolutePosition - 1, 2) = varAux
                    
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    Set tdb_Taxas.Array = xTaxa
    tdb_Taxas.Rebind
    tdb_Taxas.Refresh
    
    Exit Sub
Err_Handle:

End Sub

Private Function blnDadosGuiaOK() As Boolean

    If dbcintComposicaoDaReceita.MatchedWithList = False Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Composição da Receita deve ser selecionada."
        dbcintComposicaoDaReceita.SetFocus
        Exit Function
    End If

    If dbc_strInscricaoInicial.MatchedWithList = False Then
        ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para gerar a Guia de Arrecadação."
        dbc_strInscricaoInicial.SetFocus
        Exit Function
    End If
    
    If dbc_strInscricaoFinal.MatchedWithList = False Then
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

Private Sub tdb_AutoDeInfracao_Click()
    blnPrimeiraVez = True
End Sub

Private Sub tdb_AutoDeInfracao_FilterChange()
    blnPrimeiraVez = False
    gblnFilraCampos tdb_AutoDeInfracao
End Sub

Private Sub tdb_AutoDeInfracao_KeyPress(KeyAscii As Integer)
    Select Case tdb_AutoDeInfracao.Col
        Case 0, 1
            CaracterValido KeyAscii, "N", tdb_AutoDeInfracao
        Case 2
            CaracterValido KeyAscii, "A", tdb_AutoDeInfracao
    End Select
End Sub

Private Sub tdb_AutoDeInfracao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If blnPrimeiraVez Then
        blnRowColChange = True
        With tdb_AutoDeInfracao
            If Not .EOF And Not .BOF Then
                txtPKId.Text = tdb_AutoDeInfracao.Columns("PKId").Value
                LeDaTabelaParaObj gstrAutoDeInfracao, Me
                dbcintOrdermServico_Click (2)
                dbcintComposicaoDaReceita_Click (2)
                gCorLinhaSelecionada tdb_AutoDeInfracao
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
            End If
        End With
        blnRowColChange = False
    End If
    
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    If UCase(strModoOperacao) = gstrLocalizar Or UCase(strModoOperacao) = gstrPreencherLista Then
        Dim intAuxIndice As Integer
        For intAuxIndice = 0 To optbytTipoDeInscricaoCadastral.Count - 1
            If optbytTipoDeInscricaoCadastral(intAuxIndice).Value = True Then
                Exit For
            End If
        Next
        
        dbcintOrdermServico.Tag = strQueryDataComboOrdemServico(intAuxIndice) & ";PKId"
        dbcintComposicaoDaReceita.Tag = strQuerryComposicao(intAuxIndice) & ";strDescricao"
    
        ToolBarGeral strModoOperacao, gstrAutoDeInfracao, False, tdb_AutoDeInfracao, Me, mobjAux, strQueryAutoDeInfracao
        Exit Sub
    End If
    
    If tab_3dPasta.Tab = 0 Or tab_3dPasta.Tab = 1 Then
        If UCase(strModoOperacao) = UCase(gstrSalvar) Then
            If blnDatasOK = False Then
                Exit Sub
            End If
        End If
        If strModoOperacao = gstrNovo Or strModoOperacao = gstrSalvar Or strModoOperacao = gstrDeletar Then
            If ToolBarGeral(strModoOperacao, gstrAutoDeInfracao, False, tdb_AutoDeInfracao, Me, , strQueryAutoDeInfracao) Then
               blnPrimeiraVez = False
               Set dbl_Fiscais.RowSource = Nothing
               txt_strInscricaoCadastral.Text = ""
               LimpaEndereco
               dbcintComposicaoDaReceita.BoundText = ""
               Set xTaxa = New XArrayDB
                xTaxa.Clear
                xTaxa.ReDim 0, 0, 0, 2
                Set tdb_Taxas.Array = xTaxa
                tdb_Taxas.Rebind
                tdb_Taxas.Refresh
               Set dbcintComposicaoDaReceita.RowSource = Nothing
               HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
               optbytTipoDeInscricaoCadastral_Click (0)
            End If
        End If
    ElseIf tab_3dPasta.Tab = 2 Then
        If UCase(strModoOperacao) = UCase(gstrImprimir) Then
            If blnDadosGuiaOK Then
                Set gfrmFormularioQueEstaImprimindoGuia = Me
                rptGuiaDeArrecadacaoMunicipal.strImposto = dbcintComposicaoDaReceita.Text
'                ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, gstrQuerryRelatorioGuiaDeArrecadacao(dbc_strInscricaoInicial.Text, dbc_strInscricaoFinal.Text, txt_intExercicio.Text, dbcintComposicaoDaReceita.BoundText, , txt_DataDeVencimento.Text)
            End If
        ElseIf UCase(strModoOperacao) = UCase(gstrNovo) Then
            LimpaObjetos
        End If
    End If
    
    If strModoOperacao = gstrCalcularReajuste Then
        EfetuaCalculodeReceitasDiversas
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
End Sub

Private Function blnDatasOK() As Boolean
blnDatasOK = False
    
    If txtdtmLancamento.Text <> "" Then
        If gblnDataValida(txtdtmLancamento.Text) = False Then
            ExibeMensagem "A data de lançamento não é válida."
            txtdtmLancamento.SetFocus
            Exit Function
        End If
    End If
    
    If txtdtmVencimento.Text <> "" Then
        If gblnDataValida(txtdtmVencimento.Text) = False Then
            ExibeMensagem "A data de vencimento não é válida."
            txtdtmVencimento.SetFocus
            Exit Function
        End If
    End If

blnDatasOK = True
End Function

Private Sub tdb_Taxas_AfterColUpdate(ByVal ColIndex As Integer)
    tdb_Taxas.Update
End Sub

Private Sub tdb_Taxas_AfterUpdate()
    tdb_Taxas.Update
End Sub

Private Function strPKId() As String
    Dim strSql As String
    Dim i As Integer
    strSql = ""
    For i = 0 To xTaxa.Count(1) - 1

        If xTaxa(i, 2) = -1 Then
            If strSql <> "" Then
               strSql = strSql & ","
            End If
            strSql = strSql & xTaxa(i, 0)
        End If
    Next
    strPKId = strSql
End Function


Private Sub EfetuaCalculodeReceitasDiversas()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 06/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 07/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 08/05/2003
' Alteração: - Substituição da chamada à função CriaADO por uma chamada à função
'            ExecuteStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql                  As String
Dim strMsg                  As String
Dim adoRec                  As New ADODB.Recordset
Dim blnSelecionouReceita    As Boolean
Dim blnECalculada           As Boolean
Dim i                       As Integer
Dim dblValor                As Double
Dim dblValorCalculado       As Double
Dim dblValorParcelado       As Double
Dim dblIndexador            As Double
Dim lngSequencia            As Long
Dim datDataVencimento       As Date
Dim dblValorResto           As Double
Dim dblValorAliquota        As Double
Dim intParcelas               As Integer

    blnSelecionouReceita = False
    
    Set gobjBanco = New clsBanco
    
    If blnDadosOk Then
'        strSql = "sp_EfetuaCalculo '" & strPKId & "'," & dbcintComposicaoDaReceita.BoundText & ",11,0,NULL,0,0,0," & glngCodUsr
        strSql = gstrStoredProcedure("sp_EfetuaCalculo", "'" & strPKId & "'," & dbcintComposicaoDaReceita.BoundText & ",11,0,NULL,0,0,0," & glngCodUsr, True)
        
        Set gobjBanco = New clsBanco
'        If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
        If gobjBanco.ExecuteStoredProcedure(strSql, 10, adoRec) Then
            With adoRec
                If Not (.BOF And .EOF) Then
                    dblValorCalculado = (!dblValorCalculado)
                    dblIndexador = (!dblIndexador)
                End If
            End With
        End If
        
        If Val(txtintIntervaloParcela.Text) = 0 Then
            intParcelas = 1
        Else
            intParcelas = Val(txtintNumeroDeParcela.Text)
        End If
        
        strMsg = ""
        strMsg = strMsg & "Confirma o cálculo de " & gstrConvVrDoSql(dblValorCalculado) & Chr(10)
        If intParcelas > 1 Then
            strMsg = strMsg & "em " & intParcelas & " parcela(s) ?"
        Else
            strMsg = strMsg & "em uma  parcela ?"
        End If
        
        If Not gblnExclusaoGravacaoOk("", strMsg, True) Then
            Exit Sub
        End If
            
        gobjBanco.ExecutaBeginTrans
        Screen.MousePointer = vbHourglass
            
        'Pesquisa a sequência da composição da receita
        strSql = ""
'        strSql = strSql & " SELECT ISNULL(MAX(strSequencia),0) + 1 AS Maximo FROM " & gstrLancamentoCalculo
        strSql = strSql & " SELECT " & gstrISNULL("MAX(strSequencia)", "0") & " + 1 AS Maximo FROM " & gstrLancamentoCalculo
        strSql = strSql & " WHERE intComposicaoReceita = " & dbcintComposicaoDaReceita.BoundText
        strSql = strSql & " AND intContribuinte = " & txtintContribuinte.Text
        strSql = strSql & " AND intExercicio = " & txtintExercicio.Text
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 10, adoRec) Then
            lngSequencia = adoRec!Maximo
        End If
        
        strSql = ""
        
        If (bytDBType = EDatabases.Oracle) Then
            strSql = strSql & "DECLARE "
            strSql = strSql & "TYPE tp_csr IS REF CURSOR; "
            strSql = strSql & "csr tp_csr; "
            strSql = strSql & "V_Desconto NUMBER := 0; "
            strSql = strSql & "BEGIN "
            
        End If
        
        strSql = strSql & " INSERT INTO " & gstrLancamentoCalculo
        strSql = strSql & " (intExercicio, intContribuinte, intComposicaoReceita, intMensagem, strInscricaoCadastral, "
        strSql = strSql & " dtmLancamento, dtmVencimento, intNumeroDeParcelas, intIntervaloEntreParcelas, "
        strSql = strSql & " bitUtilizacaoDebito, intOcorrencia, bytOrigem, strSequencia, dtmDtAtualizacao, lngCodUsr ) VALUES ( "
        strSql = strSql & txtintExercicio.Text
        strSql = strSql & ", " & Val(txtintContribuinte.Text)
        strSql = strSql & ", " & dbcintComposicaoDaReceita.BoundText
        strSql = strSql & ", NULL" 'Mensagem - pode conter null
        strSql = strSql & ", '" & txtstrInscricaoCadastral.Text 'Inscrição cadastral (Para receitas diversas - código do contribuinte)
        strSql = strSql & "', " & gstrConvDtParaSql(txtdtmLancamento.Text)
        strSql = strSql & ", " & gstrConvDtParaSql(txtdtmVencimento.Text)
        strSql = strSql & ", " & intParcelas
        strSql = strSql & ", " & Val(txtintIntervaloParcela.Text)
        strSql = strSql & ", 2 " 'Utilização do débito = 2 - Econômicas
        strSql = strSql & ", " & Val(dbcintOcorrencia.BoundText) 'Ocorrência
        strSql = strSql & ", 2" 'Origem (Economico)
        strSql = strSql & ", " & CStr(lngSequencia)
'        strSql = strSql & ", GETDATE()"
        strSql = strSql & ", " & strGETDATE
        strSql = strSql & ", " & glngCodUsr
        strSql = strSql & " )"
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        'Gravar as Parcelas Taxas
'        strSQL = strSQL & " EXECUTE "
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "pe_EfetuaCalculo.", " EXECUTE ")
'        strSQL = strSQL & "sp_EfetuaCalculo '" & strPKId & "'," & dbcintComposicaoDaReceita.BoundText & ",21,"
        strSql = strSql & "sp_EfetuaCalculo" & IIf((bytDBType = EDatabases.Oracle), "(", " ") & _
            "'" & strPKId & "'," & dbcintComposicaoDaReceita.BoundText & ",21,"
        strSql = strSql & intParcelas & "," & gstrConvDtParaSql(txtdtmVencimento) & "," & txtintIntervaloParcela
'        strSql = strSql & ",0,0," & glngCodUsr
        strSql = strSql & ",0," & IIf((bytDBType = EDatabases.Oracle), " V_Desconto", " 0") & "," & glngCodUsr
              
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), ", csr); ", "")
        
        'Fim Gravar
        
        dblValorParcelado = (txtdblValorParcela + dblValorCalculado) / Val(intParcelas)
        dblValor = 0
        datDataVencimento = txtdtmVencimento.Text
        
        For i = 1 To intParcelas
        
            If i = intParcelas Then
                dblValorParcelado = (dblValorParcelado * intParcelas) - dblValorResto
            Else
                dblValorResto = gstrConvVrDoSql(dblValorResto + dblValorParcelado)
                dblValor = gstrConvVrDoSql(dblValorParcelado)
            End If
            
            strSql = strSql & " INSERT INTO " & gstrParcelaReceita
            strSql = strSql & " (intLancamentoCalculo, intComposicaoDaReceita, intNumeroParcela, dtmDataVencimento, "
            strSql = strSql & " dblValorParcela, bytDividaAjuizada, bytSimulado, bytPrescrita, "
            strSql = strSql & " bytCancelada, bytAtiva, bytSuspensaoDeExigencia, dtmDtAtualizacao, lngCodUsr) "
            strSql = strSql & " (SELECT MAX(PKId) "
            strSql = strSql & ", " & dbcintComposicaoDaReceita.BoundText
            strSql = strSql & ", " & i
            
            strSql = strSql & ", " & gstrConvDtParaSql(datDataVencimento)
            datDataVencimento = datDataVencimento + Val(txtintIntervaloParcela.Text)
                
            If i < Val(txtintIntervaloParcela.Text) Then
                strSql = strSql & ", " & gstrConvVrParaSql(gstrConvVrDoSql(dblValor))
            Else
                strSql = strSql & ", " & gstrConvVrParaSql(gstrConvVrDoSql(dblValorParcelado))
            End If
            strSql = strSql & ", 0" 'Dívida Ajuizada
            strSql = strSql & ", 0" 'Simulado
            strSql = strSql & ", 0" 'Prescrita
            strSql = strSql & ", 0" 'Cancelada
            strSql = strSql & ", 0" 'Divida Ativa
'            strSql = strSql & ",0, GETDATE()"
            strSql = strSql & ",0, " & strGETDATE
            strSql = strSql & ", " & glngCodUsr
            strSql = strSql & " FROM " & gstrLancamentoCalculo & ")"
        
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        Next i
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), " END;", "")
        
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSql, False) Then
            gobjBanco.ExecutaCommitTrans
            ExibeMensagem "Cálculo efetuado com sucesso!"
        Else
            gobjBanco.ExecutaRollbackTrans
        End If
        Screen.MousePointer = vbNormal
    End If

End Sub


Private Sub dbc_intMensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem1
End Sub

Private Sub dbc_intMensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem2
End Sub

Private Sub txt_DataDeVencimento_GotFocus()
    MarcaCampo txt_DataDeVencimento
End Sub

Private Sub txt_DataDeVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataDeVencimento
End Sub

Private Sub txt_DataDeVencimento_LostFocus()
    txt_DataDeVencimento.Text = gstrDataFormatada(txt_DataDeVencimento.Text)
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

Private Function blnDadosOk() As Boolean
    blnDadosOk = False
    Dim i As Integer
    
    If Val(Trim(txtdblValorParcela.Text)) = 0 Then
        ExibeMensagem "O campo " & lbldblValorParcela.Caption & " não pode ser zero nem nulo."
        txtdblValorParcela.SetFocus
        Exit Function
    End If
    If Not dbcintOcorrencia.MatchedWithList Then
        ExibeMensagem "O campo Ocorrência não pode ser Nulo"
        dbcintOcorrencia.SetFocus
        Exit Function
    End If
    If txtdtmLancamento.Text = "" Then
        ExibeMensagem "O campo " & lbldtmLancamento.Caption & " não pode ser nulo."
        txtdtmLancamento.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txtdtmLancamento.Text, True) Then
        txtdtmLancamento.SetFocus
        Exit Function
    End If
    If txtdtmVencimento.Text = "" Then
        ExibeMensagem "O campo " & lbldtmVencimento.Caption & " não pode ser nulo."
        txtdtmVencimento.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txtdtmVencimento.Text, True) Then
        txtdtmVencimento.SetFocus
        Exit Function
    ElseIf CVDate(txtdtmLancamento.Text) > CVDate(txtdtmVencimento.Text) Then
        ExibeMensagem "A " & lbldtmVencimento.Caption & " deve ser posterior a " & lbldtmLancamento.Caption & "."
        txtdtmVencimento.SetFocus
        Exit Function
    End If
    If Val(Trim(txtintIntervaloParcela.Text)) = 0 Then
        ExibeMensagem "O campo " & lblintIntervaloParcela.Caption & " não pode ser zero nem nulo."
        txtintIntervaloParcela.SetFocus
        Exit Function
    End If
    If txtintExercicio.Text = "" Then
        ExibeMensagem "O campo " & lblintExercicio.Caption & " não pode ser nulo."
        txtintExercicio.SetFocus
        Exit Function
    End If
    
    For i = 0 To xTaxa.Count(1) - 1
        If xTaxa(i, 2) = -1 Then
            blnDadosOk = True
            Exit Function
        End If
    Next
    ExibeMensagem "Selecione uma taxa para efetuar o cálculo!"
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


Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    Select Case tab_3dPasta.Tab
        Case 0
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
        Case 1
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
            If Trim(txtPKId.Text) <> "" Then
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
            End If
            txtdblValorParcela.SetFocus
        Case 2
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrNovo
    End Select
End Sub

Private Sub txtdblValorParcela_GotFocus()
    MarcaCampo txtdblValorParcela
End Sub

Private Sub txtdblValorParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValorParcela
End Sub

Private Sub txtdtmLancamento_GotFocus()
    MarcaCampo txtdtmLancamento
End Sub

Private Sub txtdtmLancamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmLancamento
End Sub

Private Sub txtdtmLancamento_LostFocus()
    txtdtmLancamento.Text = gstrDataFormatada(txtdtmLancamento.Text)
End Sub

Private Sub txtdtmVencimento_GotFocus()
    MarcaCampo txtdtmVencimento
End Sub

Private Sub txtdtmVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmVencimento
End Sub

Private Sub txtdtmVencimento_LostFocus()
    txtdtmVencimento.Text = gstrDataFormatada(txtdtmVencimento.Text)
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintIntervaloParcela_GotFocus()
    MarcaCampo txtintIntervaloParcela
End Sub

Private Sub txtintIntervaloParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintIntervaloParcela
End Sub

Private Sub txtintNumeroDeParcela_GotFocus()
    MarcaCampo txtintNumeroDeParcela
End Sub

Private Sub txtintNumeroDeParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumeroDeParcela
End Sub

Private Sub txtPKId_GotFocus()
    MarcaCampo txtPKId
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

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
    If Area = 2 And dbc_intMensagem1.MatchedWithList = True Then
        LeDoComboParaTXT1
    End If
End Sub

Private Sub dbc_intMensagem2_Click(Area As Integer)
    DropDownDataCombo dbc_intMensagem2, Me, Area
    If Area = 2 And dbc_intMensagem2.MatchedWithList = True Then
        LeDoComboParaTXT2
    End If
End Sub

Private Function LeDoComboParaTXT1()
Dim strSql As String
Dim adoResultado As ADODB.Recordset

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
Dim adoResultado As ADODB.Recordset

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
