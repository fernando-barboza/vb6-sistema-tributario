VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MsDatLst.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadAtualizaValores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atualização de Valores"
   ClientHeight    =   6840
   ClientLeft      =   1320
   ClientTop       =   2190
   ClientWidth     =   8940
   Icon            =   "frmCadAtualizaValores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5025
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   8864
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parâmetros"
      TabPicture(0)   =   "frmCadAtualizaValores.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblComposicao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblexercicio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtPKId"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcintcomposicaoreceita"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtintexercicio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_Referencia"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_corte"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_calculo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Composicao"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_Receitas"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Multa"
      TabPicture(1)   =   "frmCadAtualizaValores.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblexercicio1"
      Tab(1).Control(1)=   "lblComposicao1"
      Tab(1).Control(2)=   "lblacima"
      Tab(1).Control(3)=   "lbldias"
      Tab(1).Control(4)=   "lblCobrar"
      Tab(1).Control(5)=   "lblPorcento"
      Tab(1).Control(6)=   "lvw_Itens"
      Tab(1).Control(7)=   "txt_intexercicio"
      Tab(1).Control(8)=   "txt_intcomposicaoreceita"
      Tab(1).Control(9)=   "txt_Acima"
      Tab(1).Control(10)=   "txt_cobrar"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Dados para cobrança"
      TabPicture(2)   =   "frmCadAtualizaValores.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblContaBancaria"
      Tab(2).Control(1)=   "lblstrParcelaOpcional"
      Tab(2).Control(2)=   "lblstrParcela"
      Tab(2).Control(3)=   "dbcintContaBancaria"
      Tab(2).Control(4)=   "cmd_ContaBancaria"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtstrParcelaOpcional"
      Tab(2).Control(6)=   "txtstrParcela"
      Tab(2).ControlCount=   7
      Begin VB.TextBox txtstrParcela 
         Height          =   1575
         Left            =   -74880
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   3270
         Width           =   8565
      End
      Begin VB.TextBox txtstrParcelaOpcional 
         Height          =   1575
         Left            =   -74880
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   1320
         Width           =   8565
      End
      Begin VB.CommandButton cmd_ContaBancaria 
         Height          =   300
         Left            =   -71190
         Picture         =   "frmCadAtualizaValores.frx":1096
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Contas Bancarias"
         Top             =   540
         Width           =   360
      End
      Begin VB.Frame fra_Receitas 
         Caption         =   "Receitas"
         Height          =   645
         Left            =   75
         TabIndex        =   6
         Top             =   870
         Width           =   8610
         Begin MSDataListLib.DataCombo dbcintReceitaMulta 
            Height          =   315
            HelpContextID   =   1
            Left            =   510
            TabIndex        =   8
            Top             =   210
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintReceitaJuros 
            Height          =   315
            HelpContextID   =   1
            Left            =   3045
            TabIndex        =   10
            Top             =   210
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintReceitaCorrecao 
            Height          =   315
            HelpContextID   =   1
            Left            =   6765
            TabIndex        =   12
            Top             =   210
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblintReceitaMulta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Multa"
            Height          =   195
            Left            =   90
            TabIndex        =   7
            Top             =   270
            Width           =   390
         End
         Begin VB.Label lblintReceitaJuros 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Juros"
            Height          =   195
            Left            =   2625
            TabIndex        =   9
            Top             =   270
            Width           =   375
         End
         Begin VB.Label lblintReceitaCorrecao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Correção Monetária"
            Height          =   195
            Left            =   5325
            TabIndex        =   11
            Top             =   300
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmd_Composicao 
         Height          =   300
         Left            =   7020
         Picture         =   "frmCadAtualizaValores.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Composição da Receita"
         Top             =   435
         Width           =   360
      End
      Begin VB.TextBox txt_cobrar 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   -70200
         MaxLength       =   22
         TabIndex        =   36
         Top             =   1170
         Width           =   1590
      End
      Begin VB.TextBox txt_Acima 
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
         Left            =   -71865
         MaxLength       =   4
         TabIndex        =   35
         Top             =   1170
         Width           =   510
      End
      Begin VB.TextBox txt_intcomposicaoreceita 
         Height          =   315
         Left            =   -73110
         TabIndex        =   33
         Top             =   540
         Width           =   5145
      End
      Begin VB.TextBox txt_intexercicio 
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
         Left            =   -66780
         MaxLength       =   4
         TabIndex        =   34
         Top             =   540
         Width           =   495
      End
      Begin VB.Frame fra_calculo 
         Caption         =   "Formas de Cálculo "
         Height          =   2355
         Left            =   90
         TabIndex        =   23
         Top             =   2535
         Width           =   8625
         Begin MSDataListLib.DataCombo dbcinttipoformacalculoprincipal 
            Height          =   315
            HelpContextID   =   1
            Left            =   1710
            TabIndex        =   25
            Top             =   360
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcinttipoformacalculomulta 
            Height          =   315
            HelpContextID   =   1
            Left            =   1710
            TabIndex        =   27
            Top             =   750
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcinttipoformacalculojuros 
            Height          =   315
            HelpContextID   =   1
            Left            =   1710
            TabIndex        =   29
            Top             =   1110
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcinttipoformacalculotipojuros 
            Height          =   315
            HelpContextID   =   1
            Left            =   1710
            TabIndex        =   30
            Top             =   1500
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcinttipoformacalculocorrecao 
            Height          =   315
            HelpContextID   =   1
            Left            =   1710
            TabIndex        =   32
            Top             =   1860
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblCorrecaoMonetaria 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Correção Momentária"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1950
            Width           =   1515
         End
         Begin VB.Label lblJuros 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Juros"
            Height          =   195
            Left            =   1260
            TabIndex        =   28
            Top             =   1170
            Width           =   375
         End
         Begin VB.Label lblMulta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Multa"
            Height          =   195
            Left            =   1245
            TabIndex        =   26
            Top             =   810
            Width           =   390
         End
         Begin VB.Label lblPrincipal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Principal"
            Height          =   195
            Left            =   1035
            TabIndex        =   24
            Top             =   420
            Width           =   600
         End
      End
      Begin VB.Frame fra_corte 
         Caption         =   "Corte"
         Height          =   975
         Left            =   4590
         TabIndex        =   19
         Top             =   1545
         Width           =   4125
         Begin VB.CheckBox chkbitcancela 
            Caption         =   "Cancelar se menor que "" 0 """
            Height          =   225
            Left            =   540
            TabIndex        =   22
            Top             =   690
            Width           =   3435
         End
         Begin VB.ComboBox cbointcorte 
            Height          =   315
            ItemData        =   "frmCadAtualizaValores.frx":12D2
            Left            =   540
            List            =   "frmCadAtualizaValores.frx":12D4
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   330
            Width           =   3465
         End
         Begin VB.Label lblTipoCorte 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   390
            Width           =   315
         End
      End
      Begin VB.Frame fra_Referencia 
         Caption         =   "Indexador Referência"
         Height          =   975
         Left            =   90
         TabIndex        =   13
         Top             =   1545
         Width           =   4455
         Begin VB.CommandButton cmd_indexador 
            Height          =   315
            Left            =   1860
            Picture         =   "frmCadAtualizaValores.frx":12D6
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Ativa Cadastro Indexador Econômico"
            Top             =   330
            Width           =   360
         End
         Begin VB.TextBox txtdblvalor 
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
            Left            =   2700
            MaxLength       =   22
            TabIndex        =   18
            Top             =   330
            Width           =   1665
         End
         Begin MSDataListLib.DataCombo dbcintIndexadorEconomico 
            Height          =   315
            Left            =   840
            TabIndex        =   15
            Top             =   330
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblValor 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   2280
            TabIndex        =   17
            Top             =   390
            Width           =   360
         End
         Begin VB.Label lblIndexador 
            AutoSize        =   -1  'True
            Caption         =   "Indexador"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   390
            Width           =   705
         End
      End
      Begin VB.TextBox txtintexercicio 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   8190
         MaxLength       =   4
         TabIndex        =   5
         Top             =   435
         Width           =   495
      End
      Begin MSDataListLib.DataCombo dbcintcomposicaoreceita 
         Height          =   315
         Left            =   1890
         TabIndex        =   2
         Top             =   435
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ListView lvw_Itens 
         Height          =   2025
         Left            =   -73492
         TabIndex        =   45
         Top             =   2070
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   3572
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Acima de "
            Object.Width           =   5221
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "%"
            Object.Width           =   5203
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbcintContaBancaria 
         Height          =   315
         HelpContextID   =   1
         Left            =   -73650
         TabIndex        =   46
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   -2147483643
         Text            =   ""
      End
      Begin VB.Label lblstrParcela 
         Caption         =   "Instruções para parcelamento"
         Height          =   225
         Left            =   -74850
         TabIndex        =   51
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label lblstrParcelaOpcional 
         Caption         =   "Instruções para parcelas opcionais"
         Height          =   225
         Left            =   -74850
         TabIndex        =   50
         Top             =   1020
         Width           =   2595
      End
      Begin VB.Label lblContaBancaria 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente"
         Height          =   195
         Left            =   -74820
         TabIndex        =   48
         Top             =   645
         Width           =   1065
      End
      Begin VB.Label lblPorcento 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   -68520
         TabIndex        =   44
         Top             =   1260
         Width           =   120
      End
      Begin VB.Label lblCobrar 
         AutoSize        =   -1  'True
         Caption         =   "Cobrar"
         Height          =   195
         Left            =   -70860
         TabIndex        =   43
         Top             =   1230
         Width           =   465
      End
      Begin VB.Label lbldias 
         AutoSize        =   -1  'True
         Caption         =   "dias"
         Height          =   195
         Left            =   -71280
         TabIndex        =   42
         Top             =   1230
         Width           =   285
      End
      Begin VB.Label lblacima 
         AutoSize        =   -1  'True
         Caption         =   "Acima de "
         Height          =   195
         Left            =   -72660
         TabIndex        =   41
         Top             =   1230
         Width           =   705
      End
      Begin VB.Label lblComposicao1 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   -74835
         TabIndex        =   39
         Top             =   630
         Width           =   1695
      End
      Begin VB.Label lblexercicio1 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   -67530
         TabIndex        =   38
         Top             =   630
         Width           =   675
      End
      Begin VB.Label lblexercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   7440
         TabIndex        =   4
         Top             =   525
         Width           =   675
      End
      Begin VB.Label lblComposicao 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   525
         Width           =   1695
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1665
      Left            =   60
      TabIndex        =   40
      Top             =   5100
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   2937
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
      Columns(1).Caption=   "Composição da Receita"
      Columns(1).DataField=   "strComposicao"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Exercicio"
      Columns(2).DataField=   "intExercicio"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Indexador"
      Columns(3).DataField=   "strIndexador"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Valor"
      Columns(4).DataField=   "dblValor"
      Columns(4).NumberFormat=   "FormatText Event"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   16
      Columns(5)._MaxComboItems=   5
      Columns(5).ValueItems(0)._DefaultItem=   0
      Columns(5).ValueItems(0).Value=   "1"
      Columns(5).ValueItems(0).Value.vt=   8
      Columns(5).ValueItems(0).DisplayValue=   "Sim"
      Columns(5).ValueItems(0).DisplayValue.vt=   8
      Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(1)._DefaultItem=   0
      Columns(5).ValueItems(1).Value=   "0"
      Columns(5).ValueItems(1).Value.vt=   8
      Columns(5).ValueItems(1).DisplayValue=   "Não"
      Columns(5).ValueItems(1).DisplayValue.vt=   8
      Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems.Count=   2
      Columns(5).Caption=   "Corte"
      Columns(5).DataField=   "Intcorte"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Principal"
      Columns(6).DataField=   "Principal"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Multa"
      Columns(7).DataField=   "Multa"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Juros"
      Columns(8).DataField=   "Juros"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Tipo de Juros"
      Columns(9).DataField=   "TipoJuros"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Correção Monetária"
      Columns(10).DataField=   "Correcao"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=4710"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4630"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1376"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1296"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1640"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1561"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=196610"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=926"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=847"
      Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=0"
      Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(41)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(46)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(49)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(51)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(52)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(54)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(56)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(59)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=78,.parent=13,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=0"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
      _StyleDefs(80)  =   "Named:id=33:Normal"
      _StyleDefs(81)  =   ":id=33,.parent=0"
      _StyleDefs(82)  =   "Named:id=34:Heading"
      _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   ":id=34,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=36:Selected"
      _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=37:Caption"
      _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(91)  =   "Named:id=38:HighlightRow"
      _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=39:EvenRow"
      _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(95)  =   "Named:id=40:OddRow"
      _StyleDefs(96)  =   ":id=40,.parent=33"
      _StyleDefs(97)  =   "Named:id=41:RecordSelector"
      _StyleDefs(98)  =   ":id=41,.parent=34"
      _StyleDefs(99)  =   "Named:id=42:FilterBar"
      _StyleDefs(100) =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadAtualizaValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando       As Boolean
Dim mblnAlterandoAux    As Boolean
Dim mblnAlterandoLista  As Boolean
Dim mobjAux             As Object
Dim mblnSelecionou      As Boolean
Dim mblnClickOk         As Boolean
Dim mblnPrimeiraVez     As Boolean
Dim bytOrdenacao        As Byte
Dim blnOrdenacaoAsc     As Boolean
Dim mobjLista           As Object
Dim intPkid             As Long

Private Function strQuery() As String
    
    Dim strSql  As String
    
    strSql = ""
    
    strSql = strSql & "Select "
    strSql = strSql & "PA.Pkid, "
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CR.intCodigo") & strCONCAT & "' - '" & strCONCAT & " CR.strDescricao strComposicao, "
    strSql = strSql & "PA.Intexercicio, "
    strSql = strSql & "IE.STRNOME strIndexador, "
    strSql = strSql & "PA.Dblvalor, "
    strSql = strSql & "PA.Intcorte, "
    strSql = strSql & "PA.Bitcancela, "
    strSql = strSql & "P.Strabreviatura as Principal, "
    strSql = strSql & "M.Strabreviatura as Multa, "
    strSql = strSql & "J.Strabreviatura as Juros, "
    strSql = strSql & "TJ.Strabreviatura as TipoJuros, "
    strSql = strSql & "c.Strabreviatura as Correcao, "
    strSql = strSql & "PA.Intcomposicaoreceita, "
    strSql = strSql & "PA.Intindexadoreconomico, "
    strSql = strSql & "PA.Inttipoformacalculoprincipal, "
    strSql = strSql & "PA.INTTIPOFORMACALCULOMULTA, "
    strSql = strSql & "PA.Inttipoformacalculojuros, "
    strSql = strSql & "PA.Inttipoformacalculotipojuros, "
    strSql = strSql & "PA.Inttipoformacalculocorrecao, "
    strSql = strSql & "PA.strParcelaOpcional, "
    strSql = strSql & "PA.strParcela "
    
    strSql = strSql & "From "
    
    strSql = strSql & gstrParametroAtualizacao & " PA, "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
    strSql = strSql & gstrIndexadorEconomico & " IE, "
    
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 1 )  P , "
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 2 )  M , "
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 3 )  J , "
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 4 )  TJ , "
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 5 )  C "
    
    strSql = strSql & "Where "
    
    strSql = strSql & "CR.Pkid = PA.Intcomposicaoreceita            AND "
    strSql = strSql & "IE.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " PA.Intindexadoreconomico AND "
    strSql = strSql & "P.Pkid = PA.Inttipoformacalculoprincipal     AND "
    strSql = strSql & "M.Pkid = PA.INTTIPOFORMACALCULOMULTA         AND "
    strSql = strSql & "J.Pkid = PA.Inttipoformacalculojuros         AND "
    strSql = strSql & "TJ.Pkid = PA.Inttipoformacalculotipojuros    AND "
    strSql = strSql & "c.Pkid = PA.Inttipoformacalculocorrecao "
    
    If mblnClickOk = False Then
        
        If dbcintcomposicaoreceita.MatchedWithList Then
            strSql = strSql & " AND PA.Intcomposicaoreceita = " & dbcintcomposicaoreceita.BoundText
        End If
        
        If Len(Trim(txtintexercicio)) = 4 Then
            strSql = strSql & " AND PA.intExercicio = " & txtintexercicio.Text
        End If
        
        If dbcintIndexadorEconomico.MatchedWithList Then
            strSql = strSql & " AND PA.Intindexadoreconomico = " & dbcintIndexadorEconomico.BoundText
        End If
        
        If Trim(txtdblvalor) <> "" Then
            strSql = strSql & " AND PA.Dblvalor = " & gstrConvVrParaSql(txtdblvalor)
        End If
        
        If Trim(cbointcorte.Text) <> "" Then
            strSql = strSql & " AND PA.intcorte = " & cbointcorte.ItemData(cbointcorte.ListIndex)
        End If
        
        If chkbitcancela.Value Then
            strSql = strSql & " AND PA.Bitcancela = " & IIf(chkbitcancela.Value, 1, 0)
        End If
        
        If dbcintReceitaMulta.MatchedWithList Then
            strSql = strSql & " AND PA.intreceitamulta= " & dbcintReceitaMulta.BoundText
        End If
        
        If dbcintReceitaJuros.MatchedWithList Then
            strSql = strSql & " AND PA.intReceitaJuros = " & dbcintReceitaJuros.BoundText
        End If
        
        If dbcintReceitaCorrecao.MatchedWithList Then
            strSql = strSql & " AND PA.intReceitaCorrecao = " & dbcintReceitaCorrecao.BoundText
        End If
        
        If dbcinttipoformacalculoprincipal.MatchedWithList Then
            strSql = strSql & " AND PA.inttipoformacalculoprincipal= " & dbcinttipoformacalculoprincipal.BoundText
        End If
        
        If dbcinttipoformacalculomulta.MatchedWithList Then
            strSql = strSql & " AND PA.inttipoformacalculomulta = " & dbcinttipoformacalculomulta.BoundText
        End If
        
        If dbcinttipoformacalculojuros.MatchedWithList Then
            strSql = strSql & " AND PA.inttipoformacalculojuros = " & dbcinttipoformacalculojuros.BoundText
        End If
        
        If dbcinttipoformacalculotipojuros.MatchedWithList Then
            strSql = strSql & " AND PA.inttipoformacalculotipojuros= " & dbcinttipoformacalculotipojuros.BoundText
        End If
        
        If dbcinttipoformacalculocorrecao.MatchedWithList Then
            strSql = strSql & " AND PA.inttipoformacalculocorrecao = " & dbcinttipoformacalculocorrecao.BoundText
        End If
    End If
    
    Select Case bytOrdenacao
    Case Is = 1
        strSql = strSql & "Order By strComposicao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    Case Is = 2
        strSql = strSql & "Order By strIndexador" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    Case Is = 3
        strSql = strSql & "Order By PA.Dblvalor" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    Case Is = 4
        strSql = strSql & "Order By PA.Intcorte" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    Case Is = 5
        strSql = strSql & "Order By Principal" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    Case Is = 6
        strSql = strSql & "Order By Multa" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    Case Is = 7
        strSql = strSql & "Order By Juros" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    Case Is = 8
        strSql = strSql & "Order By TipoJuros" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    Case Is = 9
        strSql = strSql & "Order By Correcao" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSql
End Function

Private Function strQueryAplicar() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT PKId FROM "
    strSql = strSql & gstrParametroAtualizacao & " "
    strQueryAplicar = strSql
End Function

Private Sub cbointcorte_Click()
    If cbointcorte.ListIndex = 0 Then
        chkbitcancela.Enabled = False
        chkbitcancela.Value = False
    Else
        chkbitcancela.Enabled = True
    End If
End Sub

Private Sub cmd_Composicao_Click()
    CarregaForm frmCadComposicaoDaReceita, dbcintcomposicaoreceita
End Sub

Private Sub cmd_ContaBancaria_Click()
    CarregaForm frmCadContasBancarias, dbcintContaBancaria
End Sub

Private Sub cmd_indexador_Click()
    CarregaForm frmIndexadorEconomico, dbcintIndexadorEconomico
End Sub

Private Sub dbcintcomposicaoreceita_Change()
    If dbcintcomposicaoreceita.MatchedWithList Then
        txt_intcomposicaoreceita.Text = dbcintcomposicaoreceita.Text
    End If
End Sub

Private Sub dbcintcomposicaoreceita_Click(Area As Integer)
    DropDownDataCombo dbcintcomposicaoreceita, Me, Area
End Sub

Private Sub dbcintcomposicaoreceita_GotFocus()
    MarcaCampo dbcintcomposicaoreceita
End Sub

Private Sub dbcintcomposicaoreceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintcomposicaoreceita, Me, , KeyCode, Shift
End Sub

Private Sub dbcintcomposicaoreceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintcomposicaoreceita
End Sub

Private Sub dbcintcontabancaria_Click(Area As Integer)
    DropDownDataCombo dbcintContaBancaria, Me, Area
End Sub

Private Sub dbcintContaBancaria_GotFocus()
    MarcaCampo dbcintContaBancaria
End Sub

Private Sub dbcintContaBancaria_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintContaBancaria, Me, , KeyCode, Shift
End Sub

Private Sub dbcintContaBancaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintContaBancaria
End Sub

Private Sub dbcintReceitaMulta_Click(Area As Integer)
    DropDownDataCombo dbcintReceitaMulta, Me, Area
End Sub

Private Sub dbcintReceitaMulta_GotFocus()
    MarcaCampo dbcintReceitaMulta
    dbcintReceitaMulta.Tag = strQueryReceitas(False) & ";strSigla"
End Sub

Private Sub dbcintReceitaMulta_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintReceitaMulta, Me, , KeyCode, Shift
End Sub

Private Sub dbcintReceitaMulta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintReceitaMulta
End Sub

Private Sub dbcintReceitaJuros_Click(Area As Integer)
    DropDownDataCombo dbcintReceitaMulta, Me, Area
End Sub

Private Sub dbcintReceitaJuros_GotFocus()
    MarcaCampo dbcintReceitaJuros
    dbcintReceitaJuros.Tag = strQueryReceitas(False) & ";strSigla"
End Sub

Private Sub dbcintReceitaJuros_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintReceitaJuros, Me, , KeyCode, Shift
End Sub

Private Sub dbcintReceitaJuros_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintReceitaJuros
End Sub

Private Sub dbcintReceitaCorrecao_Click(Area As Integer)
    DropDownDataCombo dbcintReceitaCorrecao, Me, Area
End Sub

Private Sub dbcintReceitaCorrecao_GotFocus()
    MarcaCampo dbcintReceitaCorrecao
    dbcintReceitaCorrecao.Tag = strQueryReceitas(False) & ";strSigla"
End Sub

Private Sub dbcintReceitaCorrecao_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintReceitaCorrecao, Me, , KeyCode, Shift
End Sub

Private Sub dbcintReceitaCorrecao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintReceitaCorrecao
End Sub

Private Sub dbcintindexadoreconomico_Click(Area As Integer)
    DropDownDataCombo dbcintIndexadorEconomico, Me, Area
End Sub

Private Sub dbcintindexadoreconomico_GotFocus()
    MarcaCampo dbcintIndexadorEconomico
End Sub

Private Sub dbcintindexadoreconomico_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintIndexadorEconomico, Me, , KeyCode, Shift
End Sub

Private Sub dbcintindexadoreconomico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintIndexadorEconomico
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1151
    VirificaGradeListView Me
    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
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

Private Sub Form_Load()
    
    bytOrdenacao = 1: blnOrdenacaoAsc = True
    VerificaObjParaAplicar mobjAux
    PreencheCombos
    
    dbcintcomposicaoreceita.Tag = strQueryComposicao & ";strDescricao"
    dbcintIndexadorEconomico.Tag = strQueryIndexEconomico & ";strAbreviatura"
    dbcintContaBancaria.Tag = strQueryContaCorrente & ";strConta"
    
    TrocaCorObjeto txt_intcomposicaoreceita, True
    TrocaCorObjeto txt_intexercicio, True
    mblnAlterando = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
End Sub

Private Sub lvw_Itens_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txt_Acima.Text = lvw_Itens.SelectedItem.Text
    txt_cobrar.Text = lvw_Itens.SelectedItem.SubItems(1)
    mblnAlterandoLista = True
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3dPasta.Tab = 1 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
End Sub

Private Sub tab_3dPasta_LostFocus()
    dbcintContaBancaria.SetFocus
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    mblnClickOk = True
End Sub

Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 3 Then
        Value = gstrConvVrDoSql(Value, 6)
    End If
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
    mblnClickOk = False
End Sub


Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            txtPKId.Text = .Columns("PKID").Value
            
            dbcintReceitaJuros.Tag = strQueryReceitas(True) & ";strSigla"
            dbcintReceitaMulta.Tag = strQueryReceitas(True) & ";strSigla"
            dbcintReceitaCorrecao.Tag = strQueryReceitas(True) & ";strSigla"
            LeDaTabelaParaObj gstrParametroAtualizacao, Me
            
            dbcintReceitaJuros.Tag = ""
            dbcintReceitaMulta.Tag = ""
            dbcintReceitaCorrecao.Tag = ""
            
            mblnClickOk = False
            mblnAlterandoLista = False
            txt_intcomposicaoreceita.Text = dbcintcomposicaoreceita.Text
            txt_intexercicio.Text = txtintexercicio.Text
            lvw_Itens.ListItems.Clear
            PreencheListMulta
            If mblnPrimeiraVez Then
                gCorLinhaSelecionada tdb_Lista
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSql As String
    
    Select Case UCase(strModoOperacao)
    Case UCase(gstrIncluirItem)
        IncluirItemNoGrid
    Case UCase(gstrExcluirItem)
        ExcluirItemNoGrid
    Case UCase(gstrImprimir)
        ImprimeRelatorio rptAtualizaValores, strQueryRelatorio
    Case gstrSalvar
        If blnDadosOk = False Then Exit Sub
        If mblnAlterando Then
            mblnAlterandoAux = mblnAlterando
            intPkid = txtPKId
        Else
            mblnAlterandoAux = False
        End If
        If ToolBarGeral(strModoOperacao, gstrParametroAtualizacao, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery) Then
            Set gobjBanco = New clsBanco
            If lvw_Itens.ListItems.Count > 0 Then
                gobjBanco.Execute StrSalvaItem
            End If
            mblnAlterando = False
            'LeDaTabelaParaObj "", tdb_Lista, strQuery
            Limpa_Controles Me, True, True, False, True, True
            tab_3dPasta.Tab = 0
            dbcintcomposicaoreceita.SetFocus
        End If
    Case gstrDeletar
        If blnDadosOk = False Then Exit Sub
        If gblnExclusaoGravacaoOk(gstrDeletar, "Confirma exclusão ") = True Then
            Set gobjBanco = New clsBanco
            gobjBanco.Execute Deletar
            Limpa_Controles Me, True, True, False, True, True
            mblnAlterando = False
            LeDaTabelaParaObj "", tdb_Lista, strQuery
            tab_3dPasta.Tab = 0
            dbcintcomposicaoreceita.SetFocus
        End If
    Case gstrNovo
        'ToolBarGeral strModoOperacao, gstrParametroAtualizacao, mblnAlterando, tdb_Lista, _
        Me, mobjAux, strQuery
        If tab_3dPasta.Tab = 0 Then
            Limpa_Controles Me, True, True, False, True, True
            dbcintcomposicaoreceita.SetFocus
            cbointcorte.ListIndex = -1
            Set dbcintReceitaJuros.RowSource = Nothing
            Set dbcintReceitaMulta.RowSource = Nothing
            Set dbcintReceitaCorrecao.RowSource = Nothing
        Else
            txt_cobrar.Text = ""
            txt_Acima.Text = ""
            txt_Acima.SetFocus
        End If
        
        mblnAlterando = False
    Case gstrLocalizar
        LeDaTabelaParaObj "", tdb_Lista, strQuery
    Case Else
        ToolBarGeral strModoOperacao, gstrParametroAtualizacao, mblnAlterando, tdb_Lista, _
        Me, mobjAux, strQuery
        
    End Select
    
End Sub

Private Sub txt_Acima_GotFocus()
    MarcaCampo txt_Acima
End Sub

Private Sub txt_Acima_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_Acima
End Sub

Private Sub txt_cobrar_GotFocus()
    MarcaCampo txt_cobrar
End Sub

Private Sub txt_cobrar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_cobrar
    If InStr(Me.txt_cobrar.Text, ",") > 0 Then 'Procura a virgula dentro da string
    If Len(Mid(Me.txt_cobrar.Text, InStr(Me.txt_cobrar.Text, ","))) > 6 And KeyAscii <> 13 Then 'conta o numero de casas depois da virgula e trava na sexta casa
    KeyAscii = 0
End If
End If
End Sub

Private Sub txt_cobrar_LostFocus()
    txt_cobrar = gstrConvVrDoSql(txt_cobrar, 6)
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblvalor
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblvalor
End Sub

Private Sub txtdblValor_LostFocus()
    txtdblvalor = gstrConvVrDoSql(txtdblvalor, 6)
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintexercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintexercicio
End Sub

Private Sub txtintExercicio_LostFocus()
    txt_intexercicio.Text = txtintexercicio.Text
    dbcintReceitaMulta.Text = ""
    dbcintReceitaMulta.ListField = ""
    dbcintReceitaJuros.Text = ""
    dbcintReceitaJuros.ListField = ""
    dbcintReceitaCorrecao.Text = ""
    dbcintReceitaCorrecao.ListField = ""
End Sub

Private Sub txtPKId_GotFocus()
    MarcaCampo txtPKId
End Sub

Function strQueryRelatorio() As String
    Dim strSql As String
    strSql = ""
    strSql = "SELECT"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "CR.intCodigo") & strCONCAT & "' - '" & strCONCAT & " CR.strDescricao strComposicao, "
    strSql = strSql & "IE.STRNOME strIndexador, "
    strSql = strSql & "PA.Dblvalor Dblvalor, "
    strSql = strSql & "PA.Intcorte Intcorte, "
    strSql = strSql & "P.Strabreviatura Principal, "
    strSql = strSql & "M.Strabreviatura Multa, "
    strSql = strSql & "J.Strabreviatura Juros, "
    strSql = strSql & "TJ.Strabreviatura TipoJuros, "
    strSql = strSql & "C.Strabreviatura Correcao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrParametroAtualizacao & " PA, "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
    strSql = strSql & gstrIndexadorEconomico & " IE, "
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 1 )  P , "
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 2 )  M , "
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 3 )  J , "
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 4 )  TJ , "
    strSql = strSql & "(Select Pkid, Strabreviatura From " & gstrTipoFormaCalculo & " Where intTipo = 5 )  C"
    strSql = strSql & " WHERE "
    strSql = strSql & "CR.Pkid = PA.Intcomposicaoreceita AND "
    strSql = strSql & "IE.Pkid = PA.Intindexadoreconomico AND "
    strSql = strSql & "P.Pkid = PA.Inttipoformacalculoprincipal AND "
    strSql = strSql & "M.Pkid = PA.INTTIPOFORMACALCULOMULTA AND "
    strSql = strSql & "J.Pkid = PA.Inttipoformacalculojuros AND "
    strSql = strSql & "TJ.Pkid = PA.Inttipoformacalculotipojuros AND "
    strSql = strSql & "C.Pkid = PA.Inttipoformacalculocorrecao"
    strSql = strSql & " ORDER BY "
    strSql = strSql & "CR.intcodigo, "
    strSql = strSql & "CR.strDescricao"
    strQueryRelatorio = strSql
End Function

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    If Not dbcintcomposicaoreceita.MatchedWithList Then
        ExibeMensagem "O campo " & lblComposicao.Caption & " é obrigatório."
        dbcintcomposicaoreceita.SetFocus
        Exit Function
    ElseIf Len(Trim(txtintexercicio)) <> 4 Then
        ExibeMensagem "O campo " & lblexercicio.Caption & " é obrigatório."
        txtintexercicio.SetFocus
        Exit Function
    ElseIf Not dbcintIndexadorEconomico.MatchedWithList Then
        ExibeMensagem "O campo " & lblIndexador.Caption & " é obrigatório."
        dbcintIndexadorEconomico.SetFocus
        Exit Function
    ElseIf Trim(txtdblvalor) = "" Then
        ExibeMensagem "O campo valor é obrigatório."
        txtdblvalor.SetFocus
        Exit Function
    ElseIf Trim(cbointcorte.Text) = "" Then
        ExibeMensagem "O campo tipo de corte é obrigatório."
        cbointcorte.SetFocus
        Exit Function
    ElseIf Not dbcinttipoformacalculoprincipal.MatchedWithList Then
        ExibeMensagem "O campo Principal é obrigatório."
        dbcinttipoformacalculoprincipal.SetFocus
        Exit Function
    ElseIf Not dbcinttipoformacalculomulta.MatchedWithList Then
        ExibeMensagem "O campo Multa é obrigatório."
        dbcinttipoformacalculomulta.SetFocus
        Exit Function
    ElseIf Not dbcinttipoformacalculojuros.MatchedWithList Then
        ExibeMensagem "O campo Juros é obrigatório."
        dbcinttipoformacalculojuros.SetFocus
        Exit Function
    ElseIf Not dbcinttipoformacalculotipojuros.MatchedWithList Then
        ExibeMensagem "O campo Tipo de Juros é obrigatório."
        dbcinttipoformacalculotipojuros.SetFocus
        Exit Function
    ElseIf Not dbcinttipoformacalculocorrecao.MatchedWithList Then
        ExibeMensagem "O campo Correção é obrigatório."
        dbcinttipoformacalculocorrecao.SetFocus
        Exit Function
    ElseIf lvw_Itens.ListItems.Count <= 0 Then
        ExibeMensagem "Pelo menos um Multa é obrigatória."
        txt_Acima.SetFocus
        Exit Function
    End If
    blnDadosOk = True
    
End Function

Private Function strQueryComposicao() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid,"
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "intCodigo") & strCONCAT & "' - '" & strCONCAT & " strDescricao AS strDescricao"
    strSql = strSql & " FROM "
    strSql = strSql & gstrComposicaoDaReceita
    strSql = strSql & " ORDER BY strDescricao"
    
    strQueryComposicao = strSql
    
End Function

Private Function strQueryCalPrincipal() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid, "
    strSql = strSql & "strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTipoFormaCalculo
    strSql = strSql & " Where intTipo = 1 "
    strSql = strSql & "ORDER BY strDescricao"
    
    strQueryCalPrincipal = strSql
    
End Function

Private Function strQueryCalMulta() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid, "
    strSql = strSql & "strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTipoFormaCalculo
    strSql = strSql & " Where intTipo = 2 "
    strSql = strSql & "ORDER BY strDescricao"
    
    strQueryCalMulta = strSql
    
End Function

Private Function strQueryCalJuros() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid, "
    strSql = strSql & "strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTipoFormaCalculo
    strSql = strSql & " Where intTipo = 3 "
    strSql = strSql & "ORDER BY strDescricao"
    
    strQueryCalJuros = strSql
    
End Function

Private Function strQueryCalTipoJuros() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid, "
    strSql = strSql & "strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTipoFormaCalculo
    strSql = strSql & " Where intTipo = 4 "
    strSql = strSql & "ORDER BY strDescricao"
    
    strQueryCalTipoJuros = strSql
    
End Function


Private Function strQueryCalCorrecao() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid, "
    strSql = strSql & "strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTipoFormaCalculo
    strSql = strSql & " Where intTipo = 5 "
    strSql = strSql & "ORDER BY strDescricao"
    
    strQueryCalCorrecao = strSql
    
End Function

Private Function strQueryReceitas(blnSemExercicio As Boolean) As String
    Dim strSql As String
    
    strSql = ""
    'TRI0836
    strSql = "SELECT R.Pkid, "
    strSql = strSql & "R.strSigla "
    strSql = strSql & "FROM "
    strSql = strSql & gstrReceita & " R "
    
    If Not blnSemExercicio Then
        strSql = strSql & ", " & gstrReceitasExercicio & " RE "
        strSql = strSql & " WHERE RE.intReceita = R.Pkid "
        'strSql = strSql & " AND RE.intExercicio = " & gstrENulo(txtintExercicio.Text, , True)
    End If
    strSql = strSql & " GROUP BY R.Pkid,R.strSigla "
    strSql = strSql & " ORDER BY R.strSigla"
    'TRI0836
    
    
    strQueryReceitas = strSql
    
End Function


Private Sub PreencheCombos()
    
    cbointcorte.AddItem "Não"
    cbointcorte.ItemData(cbointcorte.NewIndex) = 0
    cbointcorte.AddItem "Sim"
    cbointcorte.ItemData(cbointcorte.NewIndex) = 1
    
    dbcinttipoformacalculoprincipal.Tag = strQueryCalPrincipal & ";strDescricao"
    dbcinttipoformacalculomulta.Tag = strQueryCalMulta & ";strDescricao"
    dbcinttipoformacalculojuros.Tag = strQueryCalJuros & ";strDescricao"
    dbcinttipoformacalculotipojuros.Tag = strQueryCalTipoJuros & ";strDescricao"
    dbcinttipoformacalculocorrecao.Tag = strQueryCalCorrecao & ";strDescricao"
    
    LeDaTabelaParaObj "", dbcinttipoformacalculoprincipal, strQueryCalPrincipal
    LeDaTabelaParaObj "", dbcinttipoformacalculomulta, strQueryCalMulta
    LeDaTabelaParaObj "", dbcinttipoformacalculojuros, strQueryCalJuros
    LeDaTabelaParaObj "", dbcinttipoformacalculotipojuros, strQueryCalTipoJuros
    LeDaTabelaParaObj "", dbcinttipoformacalculocorrecao, strQueryCalCorrecao
    
    dbcinttipoformacalculoprincipal.Text = ""
    dbcinttipoformacalculomulta.Text = ""
    dbcinttipoformacalculojuros.Text = ""
    dbcinttipoformacalculotipojuros.Text = ""
    dbcinttipoformacalculocorrecao.Text = ""
    
End Sub

Private Function strQueryIndexEconomico() As String
    Dim strSql As String
    
    strSql = "SELECT Pkid,"
    strSql = strSql & " strabreviatura "
    strSql = strSql & " FROM "
    strSql = strSql & gstrIndexadorEconomico
    strSql = strSql & " ORDER BY strAbreviatura"
    
    strQueryIndexEconomico = strSql
    
End Function

Private Function ExcluirItemNoGrid()
    With lvw_Itens
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
End Function

Private Function IncluirItemNoGrid()
    Dim intInd          As Integer
    
    If Trim(txt_Acima) = "" Then
        ExibeMensagem "A quantidade de dias deve ser preenchida corretamente."
        txt_Acima.SetFocus
        Exit Function
    ElseIf Trim(txt_cobrar) = "" Then
        ExibeMensagem "A porcentagem deve ser preenchida corretamente."
        txt_cobrar.SetFocus
        Exit Function
    End If
    With lvw_Itens
        If mblnAlterandoLista Then
            For intInd = 1 To .ListItems.Count
                If .SelectedItem.Index <> intInd Then
                    If Trim(txt_Acima) = .ListItems(intInd).Text Then
                        ExibeMensagem "Não é possível incluir itens com quantidade de dias iguais"
                        Exit Function
                    End If
                End If
            Next
            .SelectedItem.Text = txt_Acima.Text
            .SelectedItem.SubItems(1) = gstrConvVrDoSql(txt_cobrar.Text, 6)
            mblnAlterandoLista = False
        Else
            For intInd = 1 To .ListItems.Count
                If Trim(txt_Acima) = .ListItems(intInd).Text Then
                    ExibeMensagem "Não é possível incluir itens com quantidade de dias iguais"
                    Exit Function
                End If
            Next
            
            Set mobjLista = .ListItems.Add(, , txt_Acima.Text)
            mobjLista.SubItems(1) = gstrConvVrDoSql(txt_cobrar.Text, 6)
        End If
    End With
    txt_Acima.Text = ""
    txt_cobrar.Text = ""
    txt_Acima.SetFocus
End Function

Private Function StrSalvaItem() As String
    Dim strSql  As String
    Dim intInd  As Integer
    
    
    strSql = ""
    strSql = IIf(bytDBType = Oracle, "Begin ", "")
    If mblnAlterandoAux Then
        strSql = strSql & " Delete from " & gstrParametroAtualizacaoMulta & " Where intparametroatualizacao = " & intPkid
        strSql = strSql & IIf(bytDBType = Oracle, ";", "")
    End If
    
    With lvw_Itens
        For intInd = 1 To .ListItems.Count
            strSql = strSql & " INSERT INTO "
            
            strSql = strSql & gstrParametroAtualizacaoMulta & " ("
            
            strSql = strSql & "intparametroatualizacao, "
            strSql = strSql & "intdias, "
            strSql = strSql & "dblvalor, "
            strSql = strSql & "dtmDtAtualizacao, "
            strSql = strSql & "lngCodUsr) "
            
            strSql = strSql & "Values("
            If mblnAlterandoAux Then
                strSql = strSql & intPkid & ", "
            Else
                strSql = strSql & glngPegaUltimaChave(gstrParametroAtualizacao, "pkid") & ", "
            End If
            strSql = strSql & .ListItems(intInd).Text & ", "
            strSql = strSql & gstrConvVrParaSql(.ListItems(intInd).SubItems(1)) & ", "
            strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSql = strSql & glngCodUsr & " "
            strSql = strSql & ")" & IIf(bytDBType = Oracle, ";", "")
        Next
    End With
    strSql = strSql & IIf(bytDBType = Oracle, " End;", "")
    StrSalvaItem = strSql
End Function

Private Sub PreencheListMulta()
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    strSql = strSql & "Select "
    strSql = strSql & "intdias, "
    strSql = strSql & "dblvalor "
    strSql = strSql & "From "
    strSql = strSql & gstrParametroAtualizacaoMulta & " "
    strSql = strSql & "Where "
    strSql = strSql & "Intparametroatualizacao = " & txtPKId & " "
    strSql = strSql & "Order By intDias "
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            Do While Not adoResultado.EOF
                Set mobjLista = lvw_Itens.ListItems.Add(, , gstrENulo(adoResultado!intDias))
                mobjLista.SubItems(1) = gstrConvVrDoSql(gstrENulo(adoResultado!dblValor), 6)
                adoResultado.MoveNext
            Loop
        End If
    End If
End Sub

Private Function Deletar() As String
    Dim strSql As String
    
    strSql = ""
    strSql = IIf(bytDBType = Oracle, "Begin", "")
    
    strSql = strSql & " Delete From " & gstrParametroAtualizacaoMulta & " Where intparametroatualizacao = " & tdb_Lista.Columns("pkid").Value & IIf(bytDBType = Oracle, " ;", "")
    strSql = strSql & " Delete From " & gstrParametroAtualizacao & " Where Pkid = " & tdb_Lista.Columns("pkid").Value & IIf(bytDBType = Oracle, " ;", "")
    
    strSql = strSql & IIf(bytDBType = Oracle, "End;", "")
    
    Deletar = strSql
End Function

Private Function strQueryContaCorrente() As String
    Dim strSql As String
    
    strSql = "SELECT CB.Pkid, "
    strSql = strSql & "strConta ContaCorrente"
    strSql = strSql & " FROM " & gstrContaBancaria & " CB, "
    strSql = strSql & gstrPlanoConta & " PC"
    strSql = strSql & " Where"
    strSql = strSql & " CB.Pkid = PC.Intcontabancaria"
    strSql = strSql & " ORDER BY intNumeroConta, strDigitoVerificador"
    
    strQueryContaCorrente = strSql
    
End Function

