VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadPagamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagamentos"
   ClientHeight    =   6330
   ClientLeft      =   -30
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "CadPagamentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9540
   Begin VB.TextBox txtPKId 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2910
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6255
      Left            =   60
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   60
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Capa de Lote"
      TabPicture(0)   =   "CadPagamentos.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lançamentos"
      TabPicture(1)   =   "CadPagamentos.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_Cadastro"
      Tab(1).Control(1)=   "tdb_Parcela"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "fra_DadosdoDebito"
      Tab(1).ControlCount=   4
      Begin VB.Frame fra_DadosdoDebito 
         Caption         =   " Dados do Débito "
         Height          =   1695
         Left            =   -74685
         TabIndex        =   24
         Top             =   2460
         Width           =   8235
         Begin VB.TextBox txtdblJuros 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1725
            TabIndex        =   25
            Top             =   210
            Width           =   1605
         End
         Begin VB.TextBox txtdblMulta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4185
            TabIndex        =   26
            Top             =   210
            Width           =   1605
         End
         Begin VB.TextBox txtdblCorrecao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1725
            TabIndex        =   27
            Top             =   570
            Width           =   1605
         End
         Begin VB.TextBox txtdblTotalPago 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1725
            TabIndex        =   29
            Top             =   900
            Width           =   1605
         End
         Begin VB.TextBox txtdblDesconto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4185
            TabIndex        =   28
            Top             =   570
            Width           =   1605
         End
         Begin MSDataListLib.DataCombo dbcintOcorrencia 
            Height          =   315
            Left            =   1725
            TabIndex        =   30
            Top             =   1260
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblintOcorrencia 
            AutoSize        =   -1  'True
            Caption         =   "Ocorrência"
            Height          =   195
            Left            =   885
            TabIndex        =   48
            Top             =   1320
            Width           =   780
         End
         Begin VB.Label lbldblJuros 
            AutoSize        =   -1  'True
            Caption         =   "Juros"
            Height          =   195
            Left            =   1290
            TabIndex        =   47
            Top             =   300
            Width           =   375
         End
         Begin VB.Label lbldblMulta 
            AutoSize        =   -1  'True
            Caption         =   "Multa"
            Height          =   195
            Left            =   3735
            TabIndex        =   46
            Top             =   255
            Width           =   390
         End
         Begin VB.Label lbldblCorrecao 
            AutoSize        =   -1  'True
            Caption         =   "Correção"
            Height          =   195
            Left            =   1020
            TabIndex        =   45
            Top             =   615
            Width           =   645
         End
         Begin VB.Label lbldblTotalPago 
            AutoSize        =   -1  'True
            Caption         =   "Total Pago"
            Height          =   195
            Left            =   885
            TabIndex        =   44
            Top             =   945
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
            Height          =   195
            Left            =   3435
            TabIndex        =   43
            Top             =   615
            Width           =   690
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados da Guia/Contribuinte"
         Height          =   1995
         Left            =   -74685
         TabIndex        =   13
         Top             =   420
         Width           =   8235
         Begin VB.CheckBox chk_IssQn 
            Caption         =   "ISSQN Variável"
            Height          =   195
            Left            =   6660
            TabIndex        =   19
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txt_strSequencia 
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
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   16
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtintExercicio 
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
            Left            =   1755
            MaxLength       =   4
            TabIndex        =   15
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtintNumeroParcela 
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
            Left            =   5355
            MaxLength       =   5
            TabIndex        =   17
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txt_dtmDataVencimento 
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
            Left            =   3705
            MaxLength       =   15
            TabIndex        =   22
            Top             =   1605
            Width           =   975
         End
         Begin VB.TextBox txt_dblValorParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5205
            TabIndex        =   23
            Top             =   1605
            Width           =   1365
         End
         Begin VB.TextBox txt_dtmDataLancamento 
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
            Left            =   1725
            MaxLength       =   50
            TabIndex        =   21
            Top             =   1605
            Width           =   975
         End
         Begin MSMask.MaskEdBox mskInscricaoCadastral 
            Height          =   285
            Left            =   1755
            TabIndex        =   14
            Top             =   240
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSDataListLib.DataCombo dbc_intComposicaoReceita 
            Height          =   315
            Left            =   1740
            TabIndex        =   18
            Top             =   900
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbc_intContribuinte 
            Height          =   315
            Left            =   1740
            TabIndex        =   20
            Top             =   1245
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblstrSequencia 
            AutoSize        =   -1  'True
            Caption         =   "Sequência"
            Height          =   195
            Left            =   2400
            TabIndex        =   49
            Top             =   615
            Width           =   765
         End
         Begin VB.Label lbl_InscricaoCadastral 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   300
            TabIndex        =   42
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label lbldtmExercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   975
            TabIndex        =   41
            Top             =   615
            Width           =   675
         End
         Begin VB.Label lbldblNumeroParcela 
            AutoSize        =   -1  'True
            Caption         =   "Número da Parcela"
            Height          =   195
            Left            =   3885
            TabIndex        =   40
            Top             =   615
            Width           =   1365
         End
         Begin VB.Label lblintContribuinte 
            AutoSize        =   -1  'True
            Caption         =   "Contribuinte"
            Height          =   195
            Left            =   780
            TabIndex        =   39
            Top             =   1305
            Width           =   840
         End
         Begin VB.Label lbldtmDataVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento"
            Height          =   195
            Left            =   2760
            TabIndex        =   38
            Top             =   1650
            Width           =   840
         End
         Begin VB.Label lbldblValorParcela 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   4740
            TabIndex        =   37
            Top             =   1650
            Width           =   360
         End
         Begin VB.Label lblintComposicaoReceita 
            AutoSize        =   -1  'True
            Caption         =   "Origem da Receita"
            Height          =   195
            Left            =   300
            TabIndex        =   36
            Top             =   960
            Width           =   1320
         End
         Begin VB.Label lbldtmDataLancamento 
            AutoSize        =   -1  'True
            Caption         =   "Data do Lançamento"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1590
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Da Guia/Banco"
         Height          =   4875
         Left            =   1043
         TabIndex        =   0
         Top             =   728
         Width           =   7365
         Begin VB.TextBox txtCapaDeLote 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1755
            TabIndex        =   1
            Top             =   480
            Width           =   1605
         End
         Begin VB.TextBox txt_Utilizacao 
            Height          =   285
            Left            =   1755
            TabIndex        =   5
            Top             =   1920
            Width           =   1545
         End
         Begin VB.TextBox txtdtmPagamento 
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
            Left            =   1755
            MaxLength       =   15
            TabIndex        =   6
            Top             =   2280
            Width           =   975
         End
         Begin VB.Frame Frame3 
            Caption         =   "Fechamento"
            Height          =   1635
            Left            =   1755
            TabIndex        =   7
            Top             =   2700
            Width           =   5175
            Begin VB.TextBox txt_Juros 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   945
               TabIndex        =   8
               Top             =   300
               Width           =   1605
            End
            Begin VB.TextBox txt_Multa 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3405
               TabIndex        =   9
               Top             =   300
               Width           =   1605
            End
            Begin VB.TextBox txt_Correcao 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   945
               TabIndex        =   10
               Top             =   660
               Width           =   1605
            End
            Begin VB.TextBox txt_Desconto 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3405
               TabIndex        =   11
               Top             =   660
               Width           =   1605
            End
            Begin VB.TextBox txt_ValorTotal 
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
               Left            =   945
               MaxLength       =   15
               TabIndex        =   12
               Top             =   1140
               Width           =   1605
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Juros"
               Height          =   195
               Left            =   510
               TabIndex        =   54
               Top             =   390
               Width           =   375
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Multa"
               Height          =   195
               Left            =   2955
               TabIndex        =   53
               Top             =   345
               Width           =   390
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Correção"
               Height          =   195
               Left            =   240
               TabIndex        =   52
               Top             =   705
               Width           =   645
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Desconto"
               Height          =   195
               Left            =   2655
               TabIndex        =   51
               Top             =   705
               Width           =   690
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total"
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   1185
               Width           =   765
            End
         End
         Begin MSDataListLib.DataCombo dbcintBanco 
            Height          =   315
            Left            =   1755
            TabIndex        =   2
            Top             =   810
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintAgencia 
            Height          =   315
            Left            =   1755
            TabIndex        =   3
            Top             =   1170
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintConta 
            Height          =   315
            Left            =   1755
            TabIndex        =   4
            Top             =   1530
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblCapaDeLote 
            AutoSize        =   -1  'True
            Caption         =   "Capa de Lote"
            Height          =   195
            Left            =   720
            TabIndex        =   60
            Top             =   525
            Width           =   960
         End
         Begin VB.Label lblintConta 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   1140
            TabIndex        =   59
            Top             =   1590
            Width           =   540
         End
         Begin VB.Label lblintAgencia 
            AutoSize        =   -1  'True
            Caption         =   "Agência"
            Height          =   195
            Left            =   1095
            TabIndex        =   58
            Top             =   1230
            Width           =   585
         End
         Begin VB.Label lblintBanco 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   1215
            TabIndex        =   57
            Top             =   870
            Width           =   465
         End
         Begin VB.Label lblintUtilizacao 
            AutoSize        =   -1  'True
            Caption         =   "Utilização"
            Height          =   195
            Left            =   960
            TabIndex        =   56
            Top             =   1950
            Width           =   690
         End
         Begin VB.Label lbldtmPagamento 
            AutoSize        =   -1  'True
            Caption         =   "Data de Pagamento"
            Height          =   195
            Left            =   225
            TabIndex        =   55
            Top             =   2325
            Width           =   1425
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Parcela 
         Height          =   1995
         Left            =   -74700
         TabIndex        =   31
         Top             =   4200
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   3519
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Inscrição Cadastral"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Exercício"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Sequência"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Parcela"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "intComposicaoDaReceita"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Origem da Receita"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "intContribuinte"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Contribuinte"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Data Lançamento"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Vencimento"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Valor"
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Juros"
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Multa"
         Columns(12).DataField=   ""
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Correção"
         Columns(13).DataField=   ""
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Desconto"
         Columns(14).DataField=   ""
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Total Pago"
         Columns(15).DataField=   ""
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "intOcorrencia"
         Columns(16).DataField=   ""
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Data Pagamento"
         Columns(17).DataField=   ""
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "PKId_ParcelaTaxa"
         Columns(18).DataField=   ""
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   19
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=19"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2540"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2461"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=2"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1349"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1270"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1508"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1429"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=1191"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1111"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=0"
         Splits(0)._ColumnProps(31)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(32)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(34)=   "Column(5).Width=2831"
         Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=2752"
         Splits(0)._ColumnProps(37)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=0"
         Splits(0)._ColumnProps(39)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(40)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(41)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(43)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(44)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(45)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(46)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(48)=   "Column(7).Width=3651"
         Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=3572"
         Splits(0)._ColumnProps(51)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=0"
         Splits(0)._ColumnProps(53)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(54)=   "Column(8).Width=1482"
         Splits(0)._ColumnProps(55)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(8)._WidthInPix=1402"
         Splits(0)._ColumnProps(57)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(58)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(59)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(60)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(61)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(62)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(63)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(64)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(65)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(66)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(67)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(68)=   "Column(9).AllowFocus=0"
         Splits(0)._ColumnProps(69)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(70)=   "Column(10).Width=2037"
         Splits(0)._ColumnProps(71)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(72)=   "Column(10)._WidthInPix=1958"
         Splits(0)._ColumnProps(73)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(74)=   "Column(10).AllowSizing=0"
         Splits(0)._ColumnProps(75)=   "Column(10)._ColStyle=2"
         Splits(0)._ColumnProps(76)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(77)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(78)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(79)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(80)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(82)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(83)=   "Column(11).AllowSizing=0"
         Splits(0)._ColumnProps(84)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(85)=   "Column(11).AllowFocus=0"
         Splits(0)._ColumnProps(86)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(87)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(88)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(89)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(90)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(91)=   "Column(12).AllowSizing=0"
         Splits(0)._ColumnProps(92)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(93)=   "Column(12).AllowFocus=0"
         Splits(0)._ColumnProps(94)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(95)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(96)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(97)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(98)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(99)=   "Column(13).AllowSizing=0"
         Splits(0)._ColumnProps(100)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(101)=   "Column(13).AllowFocus=0"
         Splits(0)._ColumnProps(102)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(103)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(104)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(105)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(106)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(107)=   "Column(14).AllowSizing=0"
         Splits(0)._ColumnProps(108)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(109)=   "Column(14).AllowFocus=0"
         Splits(0)._ColumnProps(110)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(111)=   "Column(15).Width=2037"
         Splits(0)._ColumnProps(112)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(113)=   "Column(15)._WidthInPix=1958"
         Splits(0)._ColumnProps(114)=   "Column(15)._EditAlways=0"
         Splits(0)._ColumnProps(115)=   "Column(15)._ColStyle=2"
         Splits(0)._ColumnProps(116)=   "Column(15).AllowFocus=0"
         Splits(0)._ColumnProps(117)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(118)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(119)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(120)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(121)=   "Column(16)._EditAlways=0"
         Splits(0)._ColumnProps(122)=   "Column(16).AllowSizing=0"
         Splits(0)._ColumnProps(123)=   "Column(16).Visible=0"
         Splits(0)._ColumnProps(124)=   "Column(16).AllowFocus=0"
         Splits(0)._ColumnProps(125)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(126)=   "Column(17).Width=2725"
         Splits(0)._ColumnProps(127)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(128)=   "Column(17)._WidthInPix=2646"
         Splits(0)._ColumnProps(129)=   "Column(17)._EditAlways=0"
         Splits(0)._ColumnProps(130)=   "Column(17).AllowSizing=0"
         Splits(0)._ColumnProps(131)=   "Column(17).Visible=0"
         Splits(0)._ColumnProps(132)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(133)=   "Column(18).Width=2725"
         Splits(0)._ColumnProps(134)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(135)=   "Column(18)._WidthInPix=2646"
         Splits(0)._ColumnProps(136)=   "Column(18)._EditAlways=0"
         Splits(0)._ColumnProps(137)=   "Column(18).AllowSizing=0"
         Splits(0)._ColumnProps(138)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(139)=   "Column(18).Order=19"
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
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=39"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=106,.parent=13,.alignment=1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=13,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=98,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=95,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=96,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=97,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=94,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=90,.parent=13,.alignment=0"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=87,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=88,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=89,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=86,.parent=13,.alignment=0"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=82,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=79,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=80,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=81,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=13,.alignment=0"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=62,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=58,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=55,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=56,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=57,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=54,.parent=13"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=51,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=52,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=53,.parent=17"
         _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=50,.parent=13"
         _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=47,.parent=14"
         _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=48,.parent=15"
         _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=49,.parent=17"
         _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=43,.parent=14"
         _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=44,.parent=15"
         _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=45,.parent=17"
         _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=28,.parent=13"
         _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=25,.parent=14"
         _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=26,.parent=15"
         _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=27,.parent=17"
         _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=32,.parent=13"
         _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=29,.parent=14"
         _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=30,.parent=15"
         _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=31,.parent=17"
         _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=110,.parent=13"
         _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
         _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
         _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
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
         _StyleDefs(124) =   ":id=38,.parent=33,.bgcolor=&H8000000E&,.fgcolor=&H80000012&"
         _StyleDefs(125) =   "Named:id=39:EvenRow"
         _StyleDefs(126) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(127) =   "Named:id=40:OddRow"
         _StyleDefs(128) =   ":id=40,.parent=33"
         _StyleDefs(129) =   "Named:id=41:RecordSelector"
         _StyleDefs(130) =   ":id=41,.parent=34"
         _StyleDefs(131) =   "Named:id=42:FilterBar"
         _StyleDefs(132) =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_Cadastro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   -73590
         TabIndex        =   34
         Top             =   1920
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmCadPagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X                               As XArrayDB

Private Function strQueryConta() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
    strSql = ""
'    strSQL = strSQL & "SELECT C.PKId, RTRIM(C.strConta) + ' - ' + RTRIM(C.strDigitoVerificador) AS Conta "
    strSql = strSql & "SELECT C.PKId, RTRIM(C.strConta) " & strCONCAT & " ' - ' " & strCONCAT & " RTRIM(C.strDigitoVerificador) AS Conta "
    strSql = strSql & "FROM " & gstrContaBancaria & " C "
    strSql = strSql & "WHERE C.intAgencia = " & dbcintAgencia.BoundText & " "
    strSql = strSql & "AND C.intBanco = " & dbcintBanco.BoundText & " "
    strSql = strSql & "AND C.blnContaPublica = 1 "
    strSql = strSql & "Order By C.strConta"
    strQueryConta = strSql
End Function

Private Sub dbc_intComposicaoReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicaoReceita, Me, Area
    VerificaCampos
End Sub

Private Sub dbc_intComposicaoReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicaoReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicaoReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicaoReceita
End Sub

Private Sub dbc_intContribuinte_Click(Area As Integer)
    DropDownDataCombo dbc_intContribuinte, Me, Area
End Sub

Private Sub dbc_intContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intContribuinte_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intContribuinte
End Sub

Private Sub dbcintAgencia_Click(Area As Integer)
    DropDownDataCombo dbcintAgencia, Me, Area
    If Area = 2 And dbcintAgencia.MatchedWithList Then
        LeDaTabelaParaObj gstrContaBancaria, dbcintConta, strQueryConta
    End If
End Sub

Private Sub dbcintAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintAgencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintAgencia_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintAgencia
End Sub

Private Function strQueryAgencia() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT AG.PKId, AG.strDescricao FROM "
    strSql = strSql & gstrAgencia & " AG "
    strSql = strSql & "WHERE AG.intBanco = " & dbcintBanco.BoundText
    strSql = strSql & " ORDER BY AG.strDescricao"
    strQueryAgencia = strSql
End Function

Private Sub dbcintBanco_Click(Area As Integer)
    DropDownDataCombo dbcintBanco, Me, Area
    If Area = 2 And dbcintBanco.MatchedWithList Then
        LeDaTabelaParaObj gstrAgencia, dbcintAgencia, strQueryAgencia
        dbcintConta.BoundText = ""
        Set dbcintConta.DataSource = Nothing
    End If
End Sub

Private Sub dbcintBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintBanco, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBanco_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", dbcintBanco
End Sub

Private Sub dbcintConta_Click(Area As Integer)
    DropDownDataCombo dbcintConta, Me, Area
End Sub

Private Sub dbcintConta_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintConta, Me, , KeyCode, Shift
End Sub

Private Sub dbcintConta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintConta
End Sub

Private Sub dbcintOcorrencia_Click(Area As Integer)
    DropDownDataCombo dbcintOcorrencia, Me, Area
End Sub

Private Sub dbcintOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintOcorrencia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintOcorrencia
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 680
End Sub

Private Sub Form_Load()
TrocaCorObjeto txt_Utilizacao, True
TrocaCorObjeto txt_Juros, True
TrocaCorObjeto txt_Multa, True
TrocaCorObjeto txt_Correcao, True
TrocaCorObjeto txt_Desconto, True
TrocaCorObjeto txt_ValorTotal, True
TrocaCorObjeto dbc_intContribuinte, True
TrocaCorObjeto txt_dtmDataLancamento, True
TrocaCorObjeto txt_dtmDataVencimento, True
TrocaCorObjeto txt_dblValorParcela, True

HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar

LeDaTabelaParaObj gstrBanco, dbcintBanco
LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoReceita, "SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita & " ORDER BY strDescricao "
txt_Utilizacao.Text = "BAIXA"

LeDaTabelaParaObj gstrOcorrencia, dbcintOcorrencia, strQueryOcorrencia
'LeDaTabelaParaObj gstrOcorrencia, dbc_intContribuinte, "SELECT PKId, strNome FROM " & gstrContribuinte & " ORDER BY strNome"
dbc_intContribuinte.Tag = "SELECT PKId, strNome FROM " & gstrContribuinte & " ORDER BY strNome;strNome"
Set X = New XArrayDB
End Sub

Private Function strQueryOcorrencia() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo CAST do SQL Server pela função gstrCONVERT.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
    strSql = ""
'    strSQL = strSQL & "SELECT O.PKID, RTRIM(CAST(O.intCodigo AS CHAR)) " & strCONCAT & " ' - ' " & strCONCAT & " O.strDescricao AS Ocorrencia "
    strSql = strSql & "SELECT O.PKID, RTRIM(" & gstrCONVERT(CDT_VARCHAR, "O.intCodigo") & ") " & strCONCAT & " ' - ' " & strCONCAT & " O.strDescricao AS Ocorrencia "
    strSql = strSql & "FROM " & gstrOcorrencia & " O, "
    strSql = strSql & gstrUtilizacaoDaOcorrencia & " U "
    strSql = strSql & "WHERE U.PKId = O.intUtilizacaodaOcorrencia "
    strSql = strSql & "AND U.PKId = 2"  'Baixa
    strSql = strSql & "ORDER BY O.intCodigo"

    strQueryOcorrencia = strSql
End Function

Private Function blnDadosOk() As Boolean
blnDadosOk = False
    If txtCapaDeLote.Text = "" Then
        ExibeMensagem "A capa de lote tem que ser informada."
        txtCapaDeLote.SetFocus
        Exit Function
    End If
    
    If txt_Utilizacao.Text = "" Then
        ExibeMensagem "O campo utilização tem que ser informado."
        txt_Utilizacao.SetFocus
        Exit Function
    End If
    
    If Not gblnDataValida(txtdtmPagamento) Then
        ExibeMensagem "Data de pagamento inválida."
        txtdtmPagamento.SetFocus
        Exit Function
    End If
    
    If dbcintBanco.MatchedWithList = False Then
        ExibeMensagem "O banco tem que ser informado."
        dbcintBanco.SetFocus
        Exit Function
    End If
    
    If dbcintAgencia.MatchedWithList = False Then
        ExibeMensagem "A agência tem que ser informada."
        dbcintAgencia.SetFocus
        Exit Function
    End If
    
    If dbcintConta.MatchedWithList = False Then
        ExibeMensagem "A conta bancária tem que ser informada."
        dbcintConta.SetFocus
        Exit Function
    End If
    
    If Trim(mskInscricaoCadastral.ClipText) = "" Then
        ExibeMensagem "A inscrição cadastral tem que ser informada."
        mskInscricaoCadastral.SetFocus
        Exit Function
    End If
    
    If Trim(txtintExercicio) = "" Then
        ExibeMensagem "Exercício tem que ser informado."
        txtintExercicio.SetFocus
        Exit Function
    End If
    
    If dbc_intComposicaoReceita.MatchedWithList = False Then
        ExibeMensagem "A origem da receita tem que ser informada."
        dbc_intComposicaoReceita.SetFocus
        Exit Function
    End If
    
    If dbc_intContribuinte.MatchedWithList = False Then
        ExibeMensagem "O contribuinte tem que ser informado."
        dbc_intContribuinte.SetFocus
        Exit Function
    End If
    
    If dbcintOcorrencia.MatchedWithList = False Then
        ExibeMensagem "A ocorrência tem que ser informada."
        dbcintOcorrencia.SetFocus
        Exit Function
    End If
    
    If Trim(txtintNumeroParcela) = "" Then
        ExibeMensagem "Número da parcela inválido."
        txtintNumeroParcela.SetFocus
        Exit Function
    End If
    
    If txt_strSequencia.Text = "" Then
        ExibeMensagem "O campo sequência tem que ser informado."
        txt_strSequencia.SetFocus
        Exit Function
    End If
    
    If txt_dblValorParcela.Enabled = True Then
        Select Case CStr(gvntConvVrDoSql(txt_dblValorParcela))
            Case "0", ",", ""
                ExibeMensagem "O valor da parcela tem que ser informado."
                txt_dblValorParcela.SetFocus
                Exit Function
        End Select
    End If
    
    Select Case CStr(gvntConvVrDoSql(txtdblTotalPago))
        Case "0", ",", ""
            ExibeMensagem "O valor para total pago tem que ser informado."
            txtdblTotalPago.SetFocus
            Exit Function
    End Select
    
    If txt_dtmDataVencimento.Enabled = True Then
        If Not gblnDataValida(txt_dtmDataVencimento) Then
            ExibeMensagem "Data de vencimento da parcela inválida."
            txt_dtmDataVencimento.SetFocus
            Exit Function
        End If
    End If
    
    If txt_dtmDataLancamento.Enabled = True Then
        If Not gblnDataValida(txt_dtmDataLancamento) Then
            ExibeMensagem "Data de lançamento inválida."
            txt_dtmDataLancamento.SetFocus
            Exit Function
        End If
    End If
    
blnDadosOk = True
End Function


Private Sub mskInscricaoCadastral_GotFocus()
tab_3dPasta.Tab = 1
End Sub

Private Sub mskInscricaoCadastral_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", mskInscricaoCadastral
End Sub


Private Sub CarregaParcelaReceita()

'******************************************************************************************
' Data: 08/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    Dim adoRec As ADODB.Recordset
    
    On Error GoTo err_CarregaParcelaReceita

    Screen.MousePointer = vbHourglass

    strSql = ""
    strSql = strSql & " SELECT B.PKId as PKId_LancamentoCalculo, A.*, B.* FROM " & gstrParcelaReceita & " A, "
    strSql = strSql & gstrLancamentoCalculo & " B "
    strSql = strSql & " WHERE B.PKId = A.intLancamentoCalculo "
    strSql = strSql & " AND B.strInscricaoCadastral = '" & mskInscricaoCadastral.ClipText & "'"
    strSql = strSql & " AND B.intExercicio = " & Val(txtintExercicio.Text)
    strSql = strSql & " AND B.strSequencia = '" & Val(txt_strSequencia.Text) & "'"
    strSql = strSql & " AND B.intComposicaoReceita = " & dbc_intComposicaoReceita.BoundText
    'If Val(txtintNumeroParcela.Text) > 0 Then
    strSql = strSql & " AND A.intNumeroParcela = " & Val(txtintNumeroParcela.Text)
    'strSql = strSql & " AND (A.strSituacao = 'A' OR strSituacao IS NULL) "
    'End If
    
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Set gobjBanco = New clsBanco
'                gobjBanco.Execute "sp_calculoMultaJuros " & !PKId_LancamentoCalculo & "," & gstrConvDtParaSql(txtdtmPagamento)
                gobjBanco.Execute gstrStoredProcedure("sp_calculoMultaJuros", !PKId_LancamentoCalculo & "," & gstrConvDtParaSql(txtdtmPagamento))
            
                strSql = ""
                strSql = strSql & " SELECT B.PKId as PKId_LancamentoCalculo, A.*, B.* FROM " & gstrParcelaReceita & " A, "
                strSql = strSql & gstrLancamentoCalculo & " B "
                strSql = strSql & " WHERE B.PKId = A.intLancamentoCalculo "
                strSql = strSql & " AND B.PKId = " & !PKId_LancamentoCalculo
                
                'strSql = strSql & " AND B.strInscricaoCadastral = '" & mskInscricaoCadastral.ClipText & "'"
                'strSql = strSql & " AND B.intExercicio = " & Val(txtintExercicio.Text)
                'strSql = strSql & " AND B.strSequencia = '" & Val(txt_strSequencia.Text) & "'"
                'strSql = strSql & " AND B.intComposicaoReceita = " & dbc_intComposicaoReceita.BoundText
                
                strSql = strSql & " AND A.intNumeroParcela = " & Val(txtintNumeroParcela.Text)
'                strSql = strSql & " AND (A.strSituacao = 'A' OR strSituacao IS NULL) "
            End If
        End With
    End If
    
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
            
                Set gobjBanco = New clsBanco
'                gobjBanco.Execute "sp_calculoMultaJuros " & !PKId_LancamentoCalculo & "," & gstrConvDtParaSql(txtdtmPagamento)
                gobjBanco.Execute gstrStoredProcedure("sp_calculoMultaJuros", !PKId_LancamentoCalculo & "," & gstrConvDtParaSql(txtdtmPagamento))
            
                txtdblJuros.Text = "0,00"
                txtdblMulta.Text = "0,00"
                
                dbc_intComposicaoReceita.BoundText = gstrVerificaCampoNulo(!intComposicaoReceita)
                txt_dtmDataVencimento = gstrDataFormatada(gstrVerificaCampoNulo(!dtmDataVencimento))
                
                txtdblJuros.Text = gstrConvVrDoSql(!dblJuros)
                txtdblMulta.Text = gstrConvVrDoSql(!dblMulta)
                
                txt_dblValorParcela = gvntConvVrDoSql(gstrVerificaCampoNulo(!dblValorParcela))
                txtdblTotalPago.Text = gvntConvVrDoSql(CDbl(txt_dblValorParcela.Text) + CDbl(txtdblJuros.Text) + CDbl(txtdblMulta.Text))
                
                PreencherListaDeOpcoes dbc_intContribuinte, gstrVerificaCampoNulo(!intContribuinte)
                dbc_intContribuinte.BoundText = gstrVerificaCampoNulo(!intContribuinte)
                txt_dtmDataLancamento = gstrDataFormatada(gstrVerificaCampoNulo(!dtmLancamento))

                txtPKId = gstrVerificaCampoNulo(!PKId_LancamentoCalculo)
            Else
                'dbc_intComposicaoReceita.BoundText = ""
                txt_dtmDataVencimento = ""
                txt_dblValorParcela = ""
                dbc_intContribuinte = ""
                TrocaCorObjeto dbc_intContribuinte, False
                txt_dtmDataLancamento = ""
                txtPKId = ""
            End If
        End With
    End If
    
err_CarregaParcelaReceita:
Screen.MousePointer = vbDefault
End Sub

Private Sub mskInscricaoCadastral_LostFocus()
VerificaCampos
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
If tab_3dPasta.Tab = 1 Then
    VerificaCampos
End If
End Sub

Private Sub tdb_Parcela_Click()
gCorLinhaSelecionada tdb_Parcela
End Sub

Private Sub txt_strSequencia_GotFocus()
MarcaCampo txt_strSequencia
End Sub

Private Sub txt_strSequencia_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "A", txt_strSequencia
End Sub

Private Sub txtCapaDeLote_GotFocus()
MarcaCampo txtCapaDeLote
End Sub

Private Sub txtCapaDeLote_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "V", txtCapaDeLote
End Sub

Private Sub txtCapaDeLote_LostFocus()
txtCapaDeLote.Text = gstrConvVrDoSql(txtCapaDeLote, 2)
End Sub

Private Sub txtdblCorrecao_GotFocus()
MarcaCampo txtdblCorrecao
End Sub

Private Sub txtdblCorrecao_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "V", txtdblCorrecao
End Sub

Private Sub txtdblCorrecao_LostFocus()
txtdblCorrecao.Text = gstrConvVrDoSql(txtdblCorrecao.Text)
SomaNoValorTotal
End Sub

Private Sub txtdblDesconto_GotFocus()
MarcaCampo txtdblDesconto
End Sub

Private Sub txtdblDesconto_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "V", txtdblDesconto
End Sub

Private Sub txtdblDesconto_LostFocus()
txtdblDesconto.Text = gstrConvVrDoSql(txtdblDesconto.Text)
SomaNoValorTotal
End Sub

Private Sub txtdblJuros_GotFocus()
MarcaCampo txtdblJuros
End Sub

Private Sub txtdblJuros_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "V", txtdblJuros
End Sub

Private Sub SomaNoValorTotal()
Dim dblValor As Double

dblValor = 0

If Trim(txt_dblValorParcela) <> "" Then
    dblValor = txt_dblValorParcela

    If Trim(txtdblJuros) <> "" Then
        dblValor = dblValor + txtdblJuros
    End If
    If Trim(txtdblMulta) <> "" Then
        dblValor = dblValor + txtdblMulta
    End If
    If Trim(txtdblCorrecao) <> "" Then
        dblValor = dblValor + txtdblCorrecao
    End If
    If Trim(txtdblDesconto) <> "" Then
        dblValor = dblValor - txtdblDesconto
    End If
    
    txtdblTotalPago.Text = gstrConvVrDoSql(dblValor, 2)
End If
End Sub

Private Sub txtdblJuros_LostFocus()
txtdblJuros.Text = gstrConvVrDoSql(txtdblJuros.Text)
SomaNoValorTotal
End Sub

Private Sub txtdblMulta_GotFocus()
MarcaCampo txtdblMulta
End Sub

Private Sub txtdblMulta_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "V", txtdblMulta
End Sub

Private Sub txtdblMulta_LostFocus()
txtdblMulta.Text = gstrConvVrDoSql(txtdblMulta.Text)
SomaNoValorTotal
End Sub

Private Sub txtdblTotalPago_GotFocus()
MarcaCampo txtdblTotalPago
End Sub

Private Sub txtdblTotalPago_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "V", txtdblTotalPago
End Sub

Private Sub txtdblTotalPago_LostFocus()
txtdblTotalPago.Text = gstrConvVrDoSql(txtdblTotalPago.Text)
End Sub

Private Sub txtdtmPagamento_GotFocus()
MarcaCampo txtdtmPagamento
End Sub

Private Sub txtdtmPagamento_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "D", txtdtmPagamento
End Sub

Private Sub txtdtmPagamento_LostFocus()
txtdtmPagamento.Text = gstrDataFormatada(txtdtmPagamento.Text)
End Sub

Private Sub txtintExercicio_GotFocus()
MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintExercicio_LostFocus()
VerificaCampos
End Sub

Private Sub txtintNumeroParcela_GotFocus()
MarcaCampo txtintNumeroParcela
End Sub

Private Sub txtintNumeroParcela_KeyPress(KeyAscii As Integer)
CaracterValido KeyAscii, "N", txtintNumeroParcela
End Sub

Private Sub txtintNumeroParcela_LostFocus()
VerificaCampos
End Sub

Private Sub VerificaCampos()
If Trim(mskInscricaoCadastral.ClipText) <> "" And _
   Val(txtintExercicio.Text) > 0 And _
   Trim(txtintNumeroParcela.Text) <> "" And _
   dbc_intComposicaoReceita.MatchedWithList Then
    CarregaParcelaReceita
End If
End Sub

Public Sub MantemForm(strModoOperacao As String)
Select Case UCase(strModoOperacao)
    Case UCase(gstrNovo)
        LimpaFormulario
    
    Case UCase(gstrSalvar)
        If blnDadosOk Then
            SomaNoValorTotal
            MontaArray
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
        End If
        
    Case UCase(gstrDeletar)
    
    Case UCase(gstrCalcularReajuste)
        If VerificaLancamento Then
            If GravaPagamentos Then
                'LimpaFormulario
            End If
        End If
    Case UCase(gstrPreencherLista)
        PreencherListaDeOpcoes Me.ActiveControl
End Select
End Sub

Private Sub LimpaFormulario()
    txtCapaDeLote.Text = ""
    txt_Utilizacao.Text = ""
    txtdtmPagamento = ""
    dbcintBanco.BoundText = ""
    dbcintAgencia.BoundText = ""
    dbcintConta.BoundText = ""
    mskInscricaoCadastral.Text = ""
    txtintExercicio = ""
    dbc_intComposicaoReceita.BoundText = ""
    dbc_intContribuinte.BoundText = ""
    dbcintOcorrencia.BoundText = ""
    txtintNumeroParcela = ""
    txt_strSequencia.Text = ""
    txt_dblValorParcela = ""
    txtdblTotalPago = ""
    txt_dtmDataVencimento = ""
    txt_dtmDataLancamento = ""
    txt_Multa = ""
    txt_Juros = ""
    txt_Correcao = ""
    txt_Desconto = ""
    txtdblMulta = ""
    txtdblJuros = ""
    txtdblCorrecao = ""
    txtdblDesconto = ""
    txtCapaDeLote.SetFocus
End Sub

Private Function GravaPagamentos() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 13/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql As String
Dim i As Integer
Dim intLinhas As Integer

intLinhas = X.Count(1) - 1

strSql = ""

strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")

For i = 0 To intLinhas
    If Trim(txtPKId) <> "" Then

        strSql = strSql & " UPDATE " & gstrParcelaReceita
        strSql = strSql & " SET strSituacao = 'P'"
        strSql = strSql & ", dtmDataPagamento = " & gstrConvDtParaSql(txtdtmPagamento)
        strSql = strSql & ", dblJuros = " & gstrConvVrParaSql(X(i, 11))
        strSql = strSql & ", dblMulta = " & gstrConvVrParaSql(X(i, 12))
        strSql = strSql & ", dblTotalPago = " & gstrConvVrParaSql(X(i, 15))
        strSql = strSql & " WHERE intLancamentoCalculo = " & X(i, 18)
        strSql = strSql & " AND intNumeroParcela = " & gstrConvVrParaSql(X(i, 3))
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        strSql = strSql & " UPDATE " & gstrParcelaTaxa
        strSql = strSql & " SET strSituacao = 'P'"
        strSql = strSql & ", dtmDataPagamento = " & gstrConvDtParaSql(txtdtmPagamento)
        strSql = strSql & ", dblJuros = " & gstrConvVrParaSql(X(i, 11))
        strSql = strSql & ", dblMulta = " & gstrConvVrParaSql(X(i, 12))
        strSql = strSql & ", dblTotalPago = " & gstrConvVrParaSql(X(i, 15))
        strSql = strSql & " WHERE intLancamentoCalculo = " & X(i, 18)
        strSql = strSql & " AND intNumeroParcela = " & gstrConvVrParaSql(X(i, 3))
        
        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        If gstrConvVrParaSql(X(i, 3)) = 0 Then
            strSql = strSql & " UPDATE " & gstrParcelaReceita
            strSql = strSql & " SET strSituacao = 'E'"
            strSql = strSql & " WHERE intLancamentoCalculo = " & X(i, 18)
            strSql = strSql & " AND intNumeroParcela <> " & gstrConvVrParaSql(X(i, 3))
            
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
            
            strSql = strSql & " UPDATE " & gstrParcelaTaxa
            strSql = strSql & " SET strSituacao = 'E'"
            strSql = strSql & " WHERE intLancamentoCalculo = " & X(i, 18)
            strSql = strSql & " AND intNumeroParcela <> " & gstrConvVrParaSql(X(i, 3))
        
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        Else
            strSql = strSql & " UPDATE " & gstrParcelaReceita
            strSql = strSql & " SET strSituacao = 'E'"
            strSql = strSql & " WHERE intLancamentoCalculo = " & X(i, 18)
            strSql = strSql & " AND intNumeroParcela = 0"
            
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
            
            strSql = strSql & " UPDATE " & gstrParcelaTaxa
            strSql = strSql & " SET strSituacao = 'E'"
            strSql = strSql & " WHERE intLancamentoCalculo = " & X(i, 18)
            strSql = strSql & " AND intNumeroParcela = 0"
        
            strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")
        
        End If
    End If
    
    strSql = strSql & " INSERT INTO " & gstrPagamentoParcela
    strSql = strSql & " (strInscricaoCadastral, intExercicio, strSequencia, intNumeroDaParcela, intComposicaoDaReceita,"
    strSql = strSql & " intContribuinte, intBanco, intAgencia, intContaBancaria, dtmDataLancamento, dtmDataVencimento, "
    strSql = strSql & " dblValorParcela , dblJuros, dblMulta, dblCorrecao, dblDesconto, dblTotalPago, intOcorrencia, "
    strSql = strSql & " dtmDataPagamento, dtmDtAtualizacao, lngCodUsr) VALUES ( "
    
    'Inscrição
    strSql = strSql & gstrConvVrParaSql(X(i, 0))
    'Exercício
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 1))
    'Sequência
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 2))
    'Número da parcela
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 3))
    'Composição da receita
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 4))
    'Contribuinte
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 6))
    'Banco
    strSql = strSql & "," & gstrConvVrParaSql(dbcintBanco.BoundText)
    'Agencia
    strSql = strSql & "," & gstrConvVrParaSql(dbcintAgencia.BoundText)
    'Conta Bancária
    strSql = strSql & "," & gstrConvVrParaSql(dbcintConta.BoundText)
    'Data de Lançamento
    strSql = strSql & "," & gstrConvDtParaSql(X(i, 8))
    'Data de vencimento
    strSql = strSql & "," & gstrConvDtParaSql(X(i, 9))
    'Valor da Parcela
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 10))
    'Juros
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 11))
    'Multa
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 12))
    'Correção
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 13))
    'Desconto
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 14))
    'Total pago
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 15))
    'Ocorrência
    strSql = strSql & "," & gstrConvVrParaSql(X(i, 16))
    'Data do Pagamento
    strSql = strSql & "," & gstrConvDtParaSql(X(i, 17))
    'Atualização da tabela
'    strSql = strSql & ", GETDATE()"
    strSql = strSql & ", " & strGETDATE
    'Usuário
    strSql = strSql & "," & glngCodUsr
    strSql = strSql & ")"

    strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "; ", "")

Next i

strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "END; ", "")

Set gobjBanco = New clsBanco
gobjBanco.ExecutaBeginTrans

Set gobjBanco = New clsBanco
If Not gobjBanco.Execute(strSql, False) Then
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaRollbackTrans
Else
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaCommitTrans
End If

End Function

Private Function VerificaLancamento() As Boolean
If txtCapaDeLote.Text <> txt_ValorTotal.Text Then
    VerificaLancamento = False
    ExibeMensagem "Os valores lançados não fecham com o valor da capa de lote!"
Else
    VerificaLancamento = True
End If
End Function

Private Sub MontaArray()
Dim i As Integer
Dim varAux As Variant

i = X.Count(1)
X.ReDim 0, i, 0, 18

varAux = mskInscricaoCadastral.ClipText
X(i, 0) = varAux

varAux = txtintExercicio.Text
X(i, 1) = varAux

varAux = txt_strSequencia.Text
X(i, 2) = varAux

varAux = txtintNumeroParcela.Text
X(i, 3) = varAux

varAux = dbc_intComposicaoReceita.BoundText
X(i, 4) = varAux

varAux = dbc_intComposicaoReceita.Text
X(i, 5) = varAux

varAux = dbc_intContribuinte.BoundText
X(i, 6) = varAux

varAux = dbc_intContribuinte.Text
X(i, 7) = varAux

varAux = txt_dtmDataLancamento.Text
X(i, 8) = varAux

varAux = txt_dtmDataVencimento.Text
X(i, 9) = varAux

varAux = gstrConvVrDoSql(txt_dblValorParcela.Text)
X(i, 10) = varAux

If Trim(txtdblJuros.Text) <> "" Then
    If Trim(txt_Juros) <> "" Then
        txt_Juros = CDbl(txt_Juros) + CDbl(txtdblJuros.Text)
    Else
        txt_Juros = txtdblJuros.Text
    End If
End If
txt_Juros = gstrConvVrDoSql(txt_Juros)

varAux = gstrConvVrDoSql(txtdblJuros.Text)
X(i, 11) = varAux

If Trim(txtdblMulta.Text) <> "" Then
    If Trim(txt_Multa) <> "" Then
        txt_Multa = CDbl(txt_Multa) + CDbl(txtdblMulta.Text)
    Else
        txt_Multa = txtdblMulta.Text
    End If
End If
txt_Multa = gstrConvVrDoSql(txt_Multa)

varAux = gstrConvVrDoSql(txtdblMulta.Text)
X(i, 12) = varAux

If Trim(txtdblCorrecao.Text) <> "" Then
    If Trim(txt_Correcao) <> "" Then
        txt_Correcao = CDbl(txt_Correcao) + CDbl(txtdblCorrecao.Text)
    Else
        txt_Correcao = txtdblCorrecao.Text
    End If
End If
txt_Correcao = gstrConvVrDoSql(txt_Correcao)

varAux = gstrConvVrDoSql(txtdblCorrecao.Text)
X(i, 13) = varAux

If Trim(txtdblDesconto.Text) <> "" Then
    If Trim(txt_Desconto) <> "" Then
        txt_Desconto = CDbl(txt_Desconto) + CDbl(txtdblDesconto.Text)
    Else
        txt_Desconto = txtdblDesconto.Text
    End If
End If
txt_Desconto = gstrConvVrDoSql(txt_Desconto)

varAux = gstrConvVrDoSql(txtdblDesconto.Text)
X(i, 14) = varAux

If Trim(txtdblTotalPago) <> "" Then
    If Trim(txt_ValorTotal) <> "" Then
        txt_ValorTotal = CDbl(txt_ValorTotal) + CDbl(txtdblTotalPago)
    Else
        txt_ValorTotal = txtdblTotalPago
    End If
End If
varAux = gstrConvVrDoSql(txtdblTotalPago.Text)
X(i, 15) = varAux

varAux = dbcintOcorrencia.BoundText
X(i, 16) = varAux

varAux = txtdtmPagamento.Text
X(i, 17) = varAux

varAux = txtPKId.Text
X(i, 18) = varAux

Set tdb_Parcela.Array = X
tdb_Parcela.ReBind
tdb_Parcela.Refresh

LimpaCampos
End Sub

Private Sub LimpaCampos()
txtdblJuros = ""
txtdblMulta = ""
txtdblCorrecao = ""
txtdblDesconto = ""
txtdblTotalPago = ""
dbcintOcorrencia.BoundText = ""
mskInscricaoCadastral.Text = ""
mskInscricaoCadastral.SetFocus
End Sub
