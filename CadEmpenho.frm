VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadEmpenho 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empenhos"
   ClientHeight    =   8385
   ClientLeft      =   2325
   ClientTop       =   1845
   ClientWidth     =   9840
   HasDC           =   0   'False
   HelpContextID   =   10
   Icon            =   "CadEmpenho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   9840
   Begin VB.Frame fra_empenho 
      Height          =   540
      Left            =   120
      TabIndex        =   184
      Top             =   360
      Width           =   9615
      Begin VB.TextBox txtintnumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   735
         MaxLength       =   25
         OLEDragMode     =   1  'Automatic
         TabIndex        =   1
         Top             =   180
         Width           =   1425
      End
      Begin VB.TextBox txtdblSaldoEmpenho 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6150
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtintExercicioEmpenho 
         Height          =   300
         Left            =   3150
         TabIndex        =   2
         Top             =   180
         Width           =   765
      End
      Begin VB.TextBox txtdtmData 
         Height          =   285
         Left            =   4560
         OLEDragMode     =   1  'Automatic
         TabIndex        =   3
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8055
         MaxLength       =   25
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         TabIndex        =   5
         Top             =   180
         Width           =   1425
      End
      Begin VB.Label lbl_NumeroEmpenho 
         AutoSize        =   -1  'True
         Caption         =   $"CadEmpenho.frx":1042
         Height          =   195
         Left            =   120
         TabIndex        =   185
         Top             =   225
         Width           =   7875
      End
   End
   Begin VB.TextBox txt_SubPrograma 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   183
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txt_TipoCredito 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   182
      Top             =   0
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox txt_UnidadeOrcamentaria 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   181
      Top             =   0
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox txtPKId 
      Height          =   285
      Left            =   8500
      TabIndex        =   180
      Text            =   "txtPKId"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt_Orgao 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   179
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txt_Subunidade 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   178
      Top             =   0
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txt_Funcao 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   177
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.TextBox txt_Programa 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   176
      Top             =   0
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txt_Projetoatividade 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   175
      Top             =   0
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txt_Subfuncao 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   174
      Top             =   0
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txt_ElementoDespesa 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   173
      Top             =   -15
      Visible         =   0   'False
      Width           =   2715
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6465
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   11404
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Empenho"
      TabPicture(0)   =   "CadEmpenho.frx":10DE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tab_3DGeral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Subempenho"
      TabPicture(1)   =   "CadEmpenho.frx":10FA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frm_liqAutomatico"
      Tab(1).Control(1)=   "chc_LiquidarAutomaticamente"
      Tab(1).Control(2)=   "txt_CodHistoricoSub"
      Tab(1).Control(3)=   "fra_Parcela"
      Tab(1).Control(4)=   "fra_HistoricoSubEmpenho"
      Tab(1).Control(5)=   "cbo_HistoricoSubEmpenho"
      Tab(1).Control(6)=   "cmd_HistoricoSubEmpenho"
      Tab(1).Control(7)=   "lvw_ListaSubempenho"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Complemento"
      TabPicture(2)   =   "CadEmpenho.frx":1116
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt_CodHistoricoComp"
      Tab(2).Control(1)=   "txt_ValorComplemento"
      Tab(2).Control(2)=   "cmd_HistoricoComplemento"
      Tab(2).Control(3)=   "txt_DataComplemento"
      Tab(2).Control(4)=   "cbo_HistoricoComplemento"
      Tab(2).Control(5)=   "fra_HistoricoComplemento"
      Tab(2).Control(6)=   "lvw_Complemento"
      Tab(2).Control(7)=   "lbl_DataComplemento"
      Tab(2).Control(8)=   "lbl_ValorComplemento"
      Tab(2).Control(9)=   "lbl_TotalComplemento"
      Tab(2).Control(10)=   "lblTotalComplemento"
      Tab(2).Control(11)=   "lbl_Numero"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Liqüidação"
      TabPicture(3)   =   "CadEmpenho.frx":1132
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txt_CodHistoricoLiq"
      Tab(3).Control(1)=   "txt_DataVencto"
      Tab(3).Control(2)=   "cmd_EventoLiq"
      Tab(3).Control(3)=   "cbo_intEventoLiq"
      Tab(3).Control(4)=   "txt_codEventoLiq"
      Tab(3).Control(5)=   "fr_descontos"
      Tab(3).Control(6)=   "txt_dblValorAux"
      Tab(3).Control(7)=   "cbo_HistoricoLiquidacao"
      Tab(3).Control(8)=   "txt_DataLiuidacao"
      Tab(3).Control(9)=   "cmd_HistoricoLiquidacao"
      Tab(3).Control(10)=   "fra_HistoricoLiquidacao"
      Tab(3).Control(11)=   "tab_3DPastaLiquidacao"
      Tab(3).Control(12)=   "lblParcela"
      Tab(3).Control(13)=   "lblCodEventoContabilLiq"
      Tab(3).Control(14)=   "lbl_Parcela"
      Tab(3).Control(15)=   "lbl_ValorAux"
      Tab(3).Control(16)=   "lblLiquido"
      Tab(3).Control(17)=   "lbl_Liquido"
      Tab(3).Control(18)=   "lbl_DataLiquidacao"
      Tab(3).ControlCount=   19
      TabCaption(4)   =   "Anulação"
      TabPicture(4)   =   "CadEmpenho.frx":114E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "tab_3DAnulacao"
      Tab(4).ControlCount=   1
      Begin VB.Frame frm_liqAutomatico 
         BorderStyle     =   0  'None
         Caption         =   "0"
         Height          =   705
         Left            =   -74850
         TabIndex        =   218
         Top             =   1650
         Width           =   4455
         Begin VB.TextBox txt_dtmVenctoLiqAutomatica 
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
            Left            =   720
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   68
            ToolTipText     =   "Previsão de pagamento"
            Top             =   0
            Width           =   990
         End
         Begin VB.TextBox txt_strNotasFiscaisLiqAutomatica 
            Height          =   285
            Left            =   2730
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   69
            Text            =   "CadEmpenho.frx":116A
            Top             =   0
            Width           =   1665
         End
         Begin VB.TextBox txt_codEventoLiqAutomatica 
            Height          =   315
            Left            =   955
            MaxLength       =   15
            TabIndex        =   70
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox cbo_intEventoLiqAutomatica 
            Height          =   315
            Left            =   1695
            TabIndex        =   71
            Top             =   360
            Width           =   2370
         End
         Begin VB.CommandButton cmd_EventoLiqAutomatica 
            Height          =   300
            Left            =   4080
            Picture         =   "CadEmpenho.frx":117C
            Style           =   1  'Graphical
            TabIndex        =   72
            Tag             =   "247"
            ToolTipText     =   "Clique para cadastar convênio"
            Top             =   375
            Width           =   330
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Vencto"
            Height          =   195
            Left            =   30
            TabIndex        =   250
            ToolTipText     =   "Previsão de pagamento"
            Top             =   90
            Width           =   510
         End
         Begin VB.Label lbl_strNotasFiscaisLiqAutomatica 
            AutoSize        =   -1  'True
            Caption         =   "Nota Fiscal"
            Height          =   195
            Left            =   1830
            TabIndex        =   220
            Top             =   45
            Width           =   795
         End
         Begin VB.Label lbl_EventoLiqAutomatica 
            AutoSize        =   -1  'True
            Caption         =   "E. Contabil"
            Height          =   195
            Left            =   90
            TabIndex        =   219
            Top             =   420
            Width           =   765
         End
      End
      Begin VB.CheckBox chc_LiquidarAutomaticamente 
         Alignment       =   1  'Right Justify
         Caption         =   "Liquidar Automaticamente"
         Height          =   225
         Left            =   -67950
         TabIndex        =   251
         Top             =   930
         Width           =   2475
      End
      Begin VB.TextBox txt_CodHistoricoLiq 
         Height          =   315
         Left            =   -70590
         TabIndex        =   96
         Top             =   2030
         Width           =   555
      End
      Begin VB.TextBox txt_CodHistoricoComp 
         Height          =   315
         Left            =   -72840
         TabIndex        =   82
         Top             =   1860
         Width           =   555
      End
      Begin VB.TextBox txt_CodHistoricoSub 
         Height          =   315
         Left            =   -70260
         TabIndex        =   74
         Top             =   2070
         Width           =   495
      End
      Begin VB.TextBox txt_DataVencto 
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
         Left            =   -71970
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   90
         Top             =   1290
         Width           =   990
      End
      Begin VB.CommandButton cmd_EventoLiq 
         Height          =   300
         Left            =   -65700
         Picture         =   "CadEmpenho.frx":1506
         Style           =   1  'Graphical
         TabIndex        =   248
         Tag             =   "247"
         ToolTipText     =   "Clique para cadastar convênio"
         Top             =   975
         Width           =   330
      End
      Begin VB.ComboBox cbo_intEventoLiq 
         Height          =   315
         Left            =   -68625
         TabIndex        =   89
         Top             =   960
         Width           =   2910
      End
      Begin VB.TextBox txt_codEventoLiq 
         Height          =   315
         Left            =   -69405
         MaxLength       =   15
         TabIndex        =   88
         Top             =   960
         Width           =   765
      End
      Begin VB.Frame fr_descontos 
         Height          =   795
         Left            =   -73080
         TabIndex        =   235
         Top             =   1530
         Width           =   2415
         Begin VB.TextBox txt_dblDesconto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   94
            Text            =   "0,00"
            Top             =   465
            Width           =   1035
         End
         Begin VB.Label lblRetencao 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Left            =   1080
            OLEDropMode     =   1  'Manual
            TabIndex        =   93
            Tag             =   "1"
            Top             =   165
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lbl_Retencao 
            AutoSize        =   -1  'True
            Caption         =   "Retenção"
            Height          =   195
            Left            =   240
            TabIndex        =   239
            Top             =   210
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lbl_Extra 
            AutoSize        =   -1  'True
            Caption         =   "Extra"
            Height          =   195
            Left            =   540
            TabIndex        =   238
            Top             =   210
            Width           =   360
         End
         Begin VB.Label lblExtra 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Left            =   1080
            OLEDropMode     =   1  'Manual
            TabIndex        =   237
            Tag             =   "1"
            Top             =   195
            Width           =   1035
         End
         Begin VB.Label lbl_Desconto 
            AutoSize        =   -1  'True
            Caption         =   "Orçamentario"
            Height          =   195
            Left            =   90
            TabIndex        =   236
            Top             =   540
            Width           =   945
         End
      End
      Begin VB.TextBox txt_dblValorAux 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74280
         TabIndex        =   91
         Top             =   1440
         Width           =   1065
      End
      Begin TabDlg.SSTab tab_3DGeral 
         Height          =   5475
         Left            =   90
         TabIndex        =   6
         Top             =   900
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   9657
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Principal"
         TabPicture(0)   =   "CadEmpenho.frx":1890
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "tab_3DEmpenho"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fra_Processo"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fra_CodEventoContabil"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "fra_dotacao"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "fra_NumeroReserva"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Itens"
         TabPicture(1)   =   "CadEmpenho.frx":18AC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txt_SubTotalItem"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "txt_PkidItem"
         Tab(1).Control(2)=   "dbc_intStrMarca"
         Tab(1).Control(3)=   "txt_strdescricaodetalhada"
         Tab(1).Control(4)=   "txt_dblValorEstimado"
         Tab(1).Control(5)=   "txt_strObsItem"
         Tab(1).Control(6)=   "txt_dblQuantidade"
         Tab(1).Control(7)=   "txt_intCatalogoMaterialServico"
         Tab(1).Control(8)=   "txt_intCodigo"
         Tab(1).Control(9)=   "txt_intUnidadedeMedida"
         Tab(1).Control(10)=   "lvw_Itens"
         Tab(1).Control(11)=   "lbl_SubTotalItem"
         Tab(1).Control(12)=   "lblstrMarca"
         Tab(1).Control(13)=   "lblstrdescricaodetalhada"
         Tab(1).Control(14)=   "lbldblValorEstimado"
         Tab(1).Control(15)=   "lblstrObsItem"
         Tab(1).Control(16)=   "lblintUnidadedeMedida"
         Tab(1).Control(17)=   "lbldblQuantidade"
         Tab(1).Control(18)=   "lblintCodigoTipoMaterial"
         Tab(1).ControlCount=   19
         Begin VB.TextBox txt_SubTotalItem 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   330
            Left            =   -67110
            MaxLength       =   500
            TabIndex        =   252
            TabStop         =   0   'False
            Top             =   4920
            Width           =   1635
         End
         Begin VB.TextBox txt_PkidItem 
            Height          =   285
            Left            =   -74640
            TabIndex        =   65
            Top             =   540
            Visible         =   0   'False
            Width           =   525
         End
         Begin MSDataListLib.DataCombo dbc_intStrMarca 
            Height          =   315
            Left            =   -67650
            TabIndex        =   58
            Top             =   450
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.TextBox txt_strdescricaodetalhada 
            Height          =   1155
            Left            =   -73650
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   63
            Top             =   1710
            Width           =   7995
         End
         Begin VB.TextBox txt_dblValorEstimado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -71865
            MaxLength       =   25
            TabIndex        =   60
            Top             =   885
            Width           =   1335
         End
         Begin VB.TextBox txt_strObsItem 
            Height          =   330
            Left            =   -73650
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   1290
            Width           =   7995
         End
         Begin VB.TextBox txt_dblQuantidade 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -73635
            MaxLength       =   10
            TabIndex        =   59
            Top             =   900
            Width           =   765
         End
         Begin VB.TextBox txt_intCatalogoMaterialServico 
            Height          =   315
            Left            =   -72480
            TabIndex        =   57
            Top             =   450
            Width           =   4215
         End
         Begin VB.TextBox txt_intCodigo 
            Height          =   315
            Left            =   -73620
            TabIndex        =   56
            Top             =   450
            Width           =   1065
         End
         Begin VB.TextBox txt_intUnidadedeMedida 
            Height          =   315
            Left            =   -68940
            TabIndex        =   61
            Top             =   870
            Width           =   3285
         End
         Begin VB.Frame fra_NumeroReserva 
            Caption         =   " Número da Reserva "
            Height          =   930
            Left            =   0
            TabIndex        =   192
            Top             =   1605
            Width           =   7065
            Begin VB.CommandButton cmd_Reserva 
               Height          =   300
               Left            =   1020
               Picture         =   "CadEmpenho.frx":18C8
               Style           =   1  'Graphical
               TabIndex        =   16
               Tag             =   "204"
               ToolTipText     =   "Clique para consultar programa de trabalho"
               Top             =   420
               Width           =   330
            End
            Begin VB.TextBox txt_Reservado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1530
               MaxLength       =   25
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               OLEDropMode     =   1  'Manual
               TabIndex        =   17
               Top             =   450
               Width           =   1335
            End
            Begin VB.TextBox txt_Cancelado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2865
               MaxLength       =   25
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               OLEDropMode     =   1  'Manual
               TabIndex        =   18
               Top             =   450
               Width           =   1335
            End
            Begin VB.TextBox txt_Empenhado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4215
               MaxLength       =   25
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               OLEDropMode     =   1  'Manual
               TabIndex        =   19
               Top             =   450
               Width           =   1335
            End
            Begin VB.TextBox txt_Saldo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Enabled         =   0   'False
               Height          =   285
               Left            =   5565
               MaxLength       =   25
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               OLEDropMode     =   1  'Manual
               TabIndex        =   20
               Top             =   450
               Width           =   1335
            End
            Begin VB.ComboBox cbointReservaDotacao 
               Height          =   315
               Left            =   105
               TabIndex        =   15
               Top             =   420
               Width           =   915
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Valor                      Cancelado            Empenhado           Saldo"
               Height          =   195
               Left            =   1575
               TabIndex        =   193
               Top             =   240
               Width           =   4410
            End
         End
         Begin VB.Frame fra_dotacao 
            Caption         =   "Dotação"
            Height          =   540
            Left            =   45
            TabIndex        =   188
            Top             =   420
            Width           =   9465
            Begin VB.ComboBox cboProgramaTrabalho 
               Height          =   315
               Left            =   1365
               Sorted          =   -1  'True
               TabIndex        =   8
               ToolTipText     =   "Código do programa de trabalho"
               Top             =   180
               Width           =   3735
            End
            Begin VB.ComboBox cboCodigoReduzido 
               Height          =   315
               ItemData        =   "CadEmpenho.frx":1C52
               Left            =   60
               List            =   "CadEmpenho.frx":1C54
               TabIndex        =   7
               ToolTipText     =   "Código do programa de trabalho"
               Top             =   180
               Width           =   1305
            End
            Begin VB.CommandButton cmd_ProgramaTrabalho 
               Height          =   300
               Left            =   5100
               Picture         =   "CadEmpenho.frx":1C56
               Style           =   1  'Graphical
               TabIndex        =   9
               Tag             =   "204"
               ToolTipText     =   "Clique para consultar programa de trabalho"
               Top             =   180
               Width           =   330
            End
            Begin VB.TextBox txt_ValorProgramaTrabalho 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Enabled         =   0   'False
               Height          =   285
               Left            =   6165
               MaxLength       =   25
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               OLEDropMode     =   1  'Manual
               TabIndex        =   10
               Top             =   180
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txt_SaldoDotacao 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Enabled         =   0   'False
               Height          =   285
               Left            =   8055
               MaxLength       =   25
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               OLEDropMode     =   1  'Manual
               TabIndex        =   11
               Top             =   180
               Width           =   1335
            End
            Begin VB.TextBox txt_TotalDotado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Enabled         =   0   'False
               Height          =   285
               Left            =   6180
               MaxLength       =   25
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               OLEDropMode     =   1  'Manual
               TabIndex        =   189
               Top             =   180
               Width           =   1335
            End
            Begin VB.Label lbl_ValorProgramaTrabalho 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
               Height          =   195
               Left            =   6255
               TabIndex        =   191
               Top             =   225
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label lbl_TotalDotado 
               AutoSize        =   -1  'True
               Caption         =   "Tot. Emp.                               Saldo"
               Height          =   195
               Left            =   5475
               TabIndex        =   190
               Top             =   225
               Width           =   2490
            End
         End
         Begin VB.Frame fra_CodEventoContabil 
            Caption         =   "Evento Contábil "
            Height          =   585
            Left            =   45
            TabIndex        =   187
            Top             =   990
            Width           =   9465
            Begin VB.ComboBox cbo_intEvento 
               Height          =   315
               Left            =   1260
               OLEDropMode     =   1  'Manual
               TabIndex        =   13
               Top             =   200
               Width           =   7830
            End
            Begin VB.TextBox txt_codEvento 
               Height          =   315
               Left            =   60
               MaxLength       =   15
               OLEDropMode     =   1  'Manual
               TabIndex        =   12
               Top             =   210
               Width           =   1185
            End
            Begin VB.CommandButton cmd_Evento 
               Height          =   300
               Left            =   9090
               Picture         =   "CadEmpenho.frx":1FE0
               Style           =   1  'Graphical
               TabIndex        =   14
               Tag             =   "247"
               ToolTipText     =   "Clique para cadastar convênio"
               Top             =   195
               Width           =   330
            End
         End
         Begin VB.Frame fra_Processo 
            Caption         =   " Processo "
            Height          =   945
            Left            =   7245
            TabIndex        =   186
            Top             =   1605
            Width           =   1905
            Begin VB.TextBox txtbitDigito 
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
               Left            =   1470
               MaxLength       =   2
               TabIndex        =   23
               Top             =   390
               Width           =   285
            End
            Begin VB.TextBox txtintExercicio 
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
               Left            =   990
               MaxLength       =   4
               TabIndex        =   22
               Top             =   390
               Width           =   465
            End
            Begin VB.TextBox txtstrCodigo 
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
               Left            =   150
               MaxLength       =   15
               MultiLine       =   -1  'True
               TabIndex        =   21
               Top             =   390
               Width           =   825
            End
         End
         Begin TabDlg.SSTab tab_3DEmpenho 
            Height          =   2505
            Left            =   60
            TabIndex        =   194
            Top             =   2625
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   4419
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Informações Básicas"
            TabPicture(0)   =   "CadEmpenho.frx":236A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbl_ItemDespesa"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "frmCredorTipo"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "dbcintItemDespesa"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "cbo_Historico"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "cmd_Historico"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "fra_Historico"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "cmd_ItemDespesa"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "txt_intCodItemDespesa"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "txt_CodHistorico"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).ControlCount=   9
            TabCaption(1)   =   "Informações Complementar"
            TabPicture(1)   =   "CadEmpenho.frx":2386
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_Homologacao"
            Tab(1).Control(1)=   "lbl_NumeroSolicitacao"
            Tab(1).Control(2)=   "lbl_Licitacao"
            Tab(1).Control(3)=   "lbl_Modalidade"
            Tab(1).Control(4)=   "lbl_Contrato"
            Tab(1).Control(5)=   "lbl_Fundo"
            Tab(1).Control(6)=   "lblStrCondPagto"
            Tab(1).Control(7)=   "lblstrLocEntrega"
            Tab(1).Control(8)=   "lblstrPrazoEntrega"
            Tab(1).Control(9)=   "dbcintFundo"
            Tab(1).Control(10)=   "dbcintModalidade"
            Tab(1).Control(11)=   "fra_Convenio"
            Tab(1).Control(12)=   "txtdtmHomologacao"
            Tab(1).Control(13)=   "txtstrModalidade"
            Tab(1).Control(14)=   "txtstrSolicitacao"
            Tab(1).Control(15)=   "txtstrContrato"
            Tab(1).Control(16)=   "txtstrLicitacao"
            Tab(1).Control(17)=   "cmd_Fundo"
            Tab(1).Control(18)=   "txtstrEmbasamento"
            Tab(1).Control(18).Enabled=   0   'False
            Tab(1).Control(19)=   "txtstrCondPagto"
            Tab(1).Control(20)=   "txtstrLocEntrega"
            Tab(1).Control(21)=   "txtstrPrazoEntrega"
            Tab(1).ControlCount=   22
            TabCaption(2)   =   "Sub-Elementos"
            TabPicture(2)   =   "CadEmpenho.frx":23A2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "lblItemSubElemento"
            Tab(2).Control(1)=   "Label5"
            Tab(2).Control(2)=   "dbcItemDespSubElemento"
            Tab(2).Control(3)=   "lvwSubElemento"
            Tab(2).Control(4)=   "cmd_ItemDespSubElemento"
            Tab(2).Control(5)=   "txtDblValorSubElemento"
            Tab(2).Control(6)=   "txtItemDespSubElemento"
            Tab(2).ControlCount=   7
            Begin VB.TextBox txt_CodHistorico 
               Height          =   315
               Left            =   4860
               TabIndex        =   33
               Top             =   2040
               Width           =   555
            End
            Begin VB.TextBox txt_intCodItemDespesa 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1080
               TabIndex        =   29
               Top             =   1740
               Width           =   705
            End
            Begin VB.TextBox txtItemDespSubElemento 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   -73590
               TabIndex        =   51
               Top             =   420
               Width           =   705
            End
            Begin VB.TextBox txtDblValorSubElemento 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   -67140
               MaxLength       =   25
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               TabIndex        =   54
               Top             =   450
               Width           =   1425
            End
            Begin VB.CommandButton cmd_ItemDespSubElemento 
               Height          =   300
               Left            =   -68220
               Picture         =   "CadEmpenho.frx":23BE
               Style           =   1  'Graphical
               TabIndex        =   53
               Tag             =   "244"
               ToolTipText     =   "Clique para cadastar itens de despesa"
               Top             =   420
               Width           =   330
            End
            Begin VB.TextBox txtstrPrazoEntrega 
               Height          =   315
               Left            =   -68040
               MaxLength       =   50
               TabIndex        =   46
               Top             =   1050
               Width           =   2355
            End
            Begin VB.TextBox txtstrLocEntrega 
               Height          =   315
               Left            =   -68040
               MaxLength       =   50
               TabIndex        =   41
               Top             =   540
               Width           =   2355
            End
            Begin VB.TextBox txtstrCondPagto 
               Height          =   315
               Left            =   -70500
               MaxLength       =   50
               TabIndex        =   40
               Top             =   540
               Width           =   2355
            End
            Begin VB.TextBox txtstrEmbasamento 
               Height          =   285
               Left            =   -69825
               MaxLength       =   15
               TabIndex        =   214
               TabStop         =   0   'False
               Top             =   -30
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.CommandButton cmd_ItemDespesa 
               Height          =   300
               Left            =   4500
               Picture         =   "CadEmpenho.frx":2748
               Style           =   1  'Graphical
               TabIndex        =   31
               Tag             =   "244"
               ToolTipText     =   "Clique para cadastar itens de despesa"
               Top             =   1740
               Width           =   330
            End
            Begin VB.CommandButton cmd_Fundo 
               Height          =   330
               Left            =   -68475
               Picture         =   "CadEmpenho.frx":2AD2
               Style           =   1  'Graphical
               TabIndex        =   45
               Tag             =   "249"
               ToolTipText     =   "Clique para cadastar fundo"
               Top             =   1080
               Width           =   330
            End
            Begin VB.TextBox txtstrLicitacao 
               Height          =   315
               Left            =   -74850
               MaxLength       =   15
               TabIndex        =   42
               Top             =   1110
               Width           =   1035
            End
            Begin VB.TextBox txtstrContrato 
               Height          =   315
               Left            =   -74850
               MaxLength       =   15
               TabIndex        =   36
               Top             =   540
               Width           =   1035
            End
            Begin VB.TextBox txtstrSolicitacao 
               Height          =   315
               Left            =   -73740
               MaxLength       =   15
               TabIndex        =   43
               Top             =   1110
               Width           =   1245
            End
            Begin VB.TextBox txtstrModalidade 
               Height          =   315
               Left            =   -72735
               MaxLength       =   15
               TabIndex        =   38
               Top             =   540
               Width           =   960
            End
            Begin VB.TextBox txtdtmHomologacao 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   -71640
               TabIndex        =   39
               Top             =   540
               Width           =   1035
            End
            Begin VB.Frame fra_Historico 
               Caption         =   " Histórico "
               Height          =   1035
               Left            =   4890
               TabIndex        =   199
               Top             =   960
               Width           =   4485
               Begin VB.TextBox txtstrHistorico 
                  Height          =   765
                  Left            =   60
                  MaxLength       =   4000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   32
                  Top             =   180
                  Width           =   4365
               End
            End
            Begin VB.CommandButton cmd_Historico 
               Height          =   300
               Left            =   9030
               Picture         =   "CadEmpenho.frx":2E5C
               Style           =   1  'Graphical
               TabIndex        =   35
               Tag             =   "248"
               ToolTipText     =   "Clique para cadastar histórico"
               Top             =   2055
               Width           =   330
            End
            Begin VB.Frame fra_Convenio 
               Caption         =   "Convênio"
               Height          =   615
               Left            =   -74880
               TabIndex        =   195
               Top             =   1440
               Width           =   9255
               Begin VB.CommandButton cmd_Convenio 
                  Height          =   300
                  Left            =   4200
                  Picture         =   "CadEmpenho.frx":31E6
                  Style           =   1  'Graphical
                  TabIndex        =   48
                  Tag             =   "247"
                  ToolTipText     =   "Clique para cadastar convênio"
                  Top             =   210
                  Width           =   330
               End
               Begin VB.TextBox txt_DataFinal 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000000&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   5640
                  MaxLength       =   25
                  MultiLine       =   -1  'True
                  OLEDragMode     =   1  'Automatic
                  OLEDropMode     =   1  'Manual
                  TabIndex        =   49
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.TextBox txt_SaldoConvenio 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000000&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   7800
                  MaxLength       =   25
                  MultiLine       =   -1  'True
                  OLEDragMode     =   1  'Automatic
                  OLEDropMode     =   1  'Manual
                  TabIndex        =   50
                  Top             =   240
                  Width           =   1335
               End
               Begin MSDataListLib.DataCombo dbcintConvenio 
                  Height          =   315
                  Left            =   930
                  TabIndex        =   47
                  Top             =   210
                  Width           =   3225
                  _ExtentX        =   5689
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
               End
               Begin VB.Label lbl_Convenio 
                  AutoSize        =   -1  'True
                  Caption         =   "Descrição"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   198
                  Top             =   330
                  Width           =   720
               End
               Begin VB.Label lblDataFinal 
                  AutoSize        =   -1  'True
                  Caption         =   "Data Final"
                  Height          =   195
                  Left            =   4830
                  TabIndex        =   197
                  Top             =   330
                  Width           =   720
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Saldo"
                  Height          =   195
                  Left            =   7320
                  TabIndex        =   196
                  Top             =   330
                  Width           =   405
               End
            End
            Begin MSDataListLib.DataCombo dbcintModalidade 
               Height          =   315
               Left            =   -73710
               TabIndex        =   37
               Top             =   540
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dbcintFundo 
               Height          =   315
               Left            =   -72375
               TabIndex        =   44
               Top             =   1110
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cbo_Historico 
               Height          =   315
               Left            =   5460
               TabIndex        =   34
               Top             =   2040
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dbcintItemDespesa 
               Height          =   315
               Left            =   1785
               TabIndex        =   30
               Top             =   1740
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin MSComctlLib.ListView lvwSubElemento 
               Height          =   1185
               Left            =   -74970
               TabIndex        =   55
               Top             =   840
               Width           =   9225
               _ExtentX        =   16272
               _ExtentY        =   2090
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
                  Text            =   "Código "
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Descrição"
                  Object.Width           =   11765
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Valor"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSDataListLib.DataCombo dbcItemDespSubElemento 
               Height          =   315
               Left            =   -72855
               TabIndex        =   52
               Top             =   420
               Width           =   4590
               _ExtentX        =   8096
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Frame frmCredorTipo 
               BorderStyle     =   0  'None
               Height          =   1305
               Left            =   150
               TabIndex        =   242
               Top             =   420
               Width           =   9255
               Begin VB.TextBox txt_intNContribuinte 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   885
                  TabIndex        =   24
                  Top             =   150
                  Width           =   705
               End
               Begin VB.CommandButton cmd_Tipo 
                  Height          =   300
                  Left            =   4350
                  Picture         =   "CadEmpenho.frx":3570
                  Style           =   1  'Graphical
                  TabIndex        =   28
                  ToolTipText     =   "Clique para cadastar tipo"
                  Top             =   765
                  Width           =   330
               End
               Begin VB.CommandButton cmd_Credor 
                  Height          =   300
                  Left            =   8850
                  Picture         =   "CadEmpenho.frx":38FA
                  Style           =   1  'Graphical
                  TabIndex        =   26
                  Tag             =   "15"
                  ToolTipText     =   "Clique para cadastar contribuinte"
                  Top             =   165
                  Width           =   330
               End
               Begin MSDataListLib.DataCombo dbcintTipo 
                  Height          =   315
                  Left            =   885
                  TabIndex        =   27
                  Top             =   750
                  Width           =   3450
                  _ExtentX        =   6085
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
               End
               Begin MSDataListLib.DataCombo dbcintCredor 
                  Height          =   315
                  Left            =   1590
                  TabIndex        =   25
                  Top             =   150
                  Width           =   7275
                  _ExtentX        =   12832
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
               End
               Begin VB.Label lbl_Tipo 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo"
                  Height          =   195
                  Left            =   540
                  TabIndex        =   244
                  Top             =   820
                  Width           =   315
               End
               Begin VB.Label lbl_Fornecedor 
                  AutoSize        =   -1  'True
                  Caption         =   "Credor"
                  Height          =   195
                  Left            =   390
                  TabIndex        =   243
                  Top             =   270
                  Width           =   465
               End
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
               Height          =   195
               Left            =   -67560
               TabIndex        =   241
               Top             =   495
               Width           =   360
            End
            Begin VB.Label lblItemSubElemento 
               AutoSize        =   -1  'True
               Caption         =   "Item de Despesa"
               Height          =   195
               Left            =   -74850
               TabIndex        =   240
               ToolTipText     =   "Item de despesa"
               Top             =   495
               Width           =   1200
            End
            Begin VB.Label lblstrPrazoEntrega 
               AutoSize        =   -1  'True
               Caption         =   "Prazo de Entrega"
               Height          =   195
               Left            =   -68040
               TabIndex        =   217
               Top             =   840
               Width           =   1230
            End
            Begin VB.Label lblstrLocEntrega 
               AutoSize        =   -1  'True
               Caption         =   "Local de Entrega"
               Height          =   195
               Left            =   -68040
               TabIndex        =   216
               Top             =   330
               Width           =   1215
            End
            Begin VB.Label lblStrCondPagto 
               AutoSize        =   -1  'True
               Caption         =   "Condição Pagamento"
               Height          =   195
               Left            =   -70500
               TabIndex        =   215
               Top             =   330
               Width           =   1530
            End
            Begin VB.Label lbl_Fundo 
               AutoSize        =   -1  'True
               Caption         =   "Fundo"
               Height          =   195
               Left            =   -72345
               TabIndex        =   206
               Top             =   900
               Width           =   450
            End
            Begin VB.Label lbl_Contrato 
               AutoSize        =   -1  'True
               Caption         =   "Contrato"
               Height          =   195
               Left            =   -74850
               TabIndex        =   205
               Top             =   330
               Width           =   600
            End
            Begin VB.Label lbl_Modalidade 
               AutoSize        =   -1  'True
               Caption         =   "Modalidade"
               Height          =   195
               Left            =   -73710
               TabIndex        =   204
               Top             =   330
               Width           =   825
            End
            Begin VB.Label lbl_Licitacao 
               AutoSize        =   -1  'True
               Caption         =   "Licitação"
               Height          =   195
               Left            =   -74850
               TabIndex        =   203
               Top             =   900
               Width           =   645
            End
            Begin VB.Label lbl_NumeroSolicitacao 
               AutoSize        =   -1  'True
               Caption         =   "Pedido  Empenho"
               Height          =   195
               Left            =   -73755
               TabIndex        =   202
               Top             =   900
               Width           =   1260
            End
            Begin VB.Label lbl_Homologacao 
               AutoSize        =   -1  'True
               Caption         =   "Homologação"
               Height          =   195
               Left            =   -71640
               TabIndex        =   201
               Top             =   330
               Width           =   990
            End
            Begin VB.Label lbl_ItemDespesa 
               AutoSize        =   -1  'True
               Caption         =   "I.de Despesa"
               Height          =   195
               Left            =   60
               TabIndex        =   200
               ToolTipText     =   "Item de despesa"
               Top             =   1860
               Width           =   945
            End
         End
         Begin MSComctlLib.ListView lvw_Itens 
            Height          =   1695
            Left            =   -74940
            TabIndex        =   64
            Top             =   3090
            Width           =   9465
            _ExtentX        =   16695
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Pkid"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Código "
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Descrição"
               Object.Width           =   7937
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Marca"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Quantidade"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Vlr. Unitário"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Unid. Medida"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Observação"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Complemento"
               Object.Width           =   17639
            EndProperty
         End
         Begin VB.Label lbl_SubTotalItem 
            AutoSize        =   -1  'True
            Caption         =   "Subtotal"
            Height          =   195
            Left            =   -67800
            TabIndex        =   253
            Top             =   4980
            Width           =   585
         End
         Begin VB.Label lblstrMarca 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            Height          =   195
            Left            =   -68190
            TabIndex        =   213
            Top             =   540
            Width           =   450
         End
         Begin VB.Label lblstrdescricaodetalhada 
            AutoSize        =   -1  'True
            Caption         =   "Complemento"
            Height          =   195
            Left            =   -74730
            TabIndex        =   212
            Top             =   1710
            Width           =   960
         End
         Begin VB.Label lbldblValorEstimado 
            AutoSize        =   -1  'True
            Caption         =   "Vlr. Unitário"
            Height          =   195
            Left            =   -72825
            TabIndex        =   211
            Top             =   960
            Width           =   810
         End
         Begin VB.Label lblstrObsItem 
            AutoSize        =   -1  'True
            Caption         =   "Observação"
            Height          =   195
            Left            =   -74610
            TabIndex        =   210
            Top             =   1350
            Width           =   870
         End
         Begin VB.Label lblintUnidadedeMedida 
            AutoSize        =   -1  'True
            Caption         =   "Unidade de Medida"
            Height          =   195
            Left            =   -70440
            TabIndex        =   209
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lbldblQuantidade 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            Height          =   195
            Left            =   -74565
            TabIndex        =   208
            Top             =   960
            Width           =   825
         End
         Begin VB.Label lblintCodigoTipoMaterial 
            AutoSize        =   -1  'True
            Caption         =   "Item"
            Height          =   195
            Left            =   -74100
            TabIndex        =   207
            Top             =   555
            Width           =   300
         End
      End
      Begin VB.TextBox txt_ValorComplemento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74430
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   79
         Top             =   1410
         Width           =   1395
      End
      Begin VB.ComboBox cbo_HistoricoLiquidacao 
         Height          =   315
         Left            =   -70020
         Sorted          =   -1  'True
         TabIndex        =   97
         ToolTipText     =   "Histórico padrão"
         Top             =   2030
         Width           =   4305
      End
      Begin VB.TextBox txt_DataLiuidacao 
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
         Left            =   -71970
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   87
         Top             =   960
         Width           =   990
      End
      Begin VB.CommandButton cmd_HistoricoLiquidacao 
         Height          =   300
         Left            =   -65670
         Picture         =   "CadEmpenho.frx":3C84
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Clique para cadastar histórico"
         Top             =   2030
         Width           =   330
      End
      Begin VB.Frame fra_HistoricoLiquidacao 
         Caption         =   " Histórico "
         Height          =   705
         Left            =   -70590
         TabIndex        =   160
         Top             =   1290
         Width           =   5235
         Begin VB.TextBox txt_HistoricoLiquidacao 
            Height          =   530
            Left            =   0
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   95
            Top             =   180
            Width           =   5235
         End
      End
      Begin TabDlg.SSTab tab_3DPastaLiquidacao 
         Height          =   3705
         Left            =   -74970
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   2370
         Width           =   9705
         _ExtentX        =   17119
         _ExtentY        =   6535
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Parcela "
         TabPicture(0)   =   "CadEmpenho.frx":400E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lbl_ContaRetencao"
         Tab(0).Control(1)=   "lvw_Liquidacao"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Extra"
         TabPicture(1)   =   "CadEmpenho.frx":402A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lbl_Conta"
         Tab(1).Control(1)=   "lbl_valorExtra"
         Tab(1).Control(2)=   "lvw_Extra"
         Tab(1).Control(3)=   "cbo_ContaExtra"
         Tab(1).Control(4)=   "cbo_DescricaoExtra"
         Tab(1).Control(5)=   "txt_ValorExtra"
         Tab(1).Control(6)=   "cmd_ContaExtra"
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Retenção"
         TabPicture(2)   =   "CadEmpenho.frx":4046
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label2"
         Tab(2).Control(1)=   "Label3"
         Tab(2).Control(2)=   "lvw_Retencao"
         Tab(2).Control(3)=   "cbo_ContaRetencao"
         Tab(2).Control(4)=   "cbo_DescricaoRetencao"
         Tab(2).Control(5)=   "txt_ValorRetencao"
         Tab(2).Control(6)=   "cmd_ContaRetencao"
         Tab(2).ControlCount=   7
         TabCaption(3)   =   "Orçamentario"
         TabPicture(3)   =   "CadEmpenho.frx":4062
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lbl_ContaOrcamentario"
         Tab(3).Control(1)=   "lbl_valorOrcamentario"
         Tab(3).Control(2)=   "lvw_Orcamentario"
         Tab(3).Control(3)=   "cmd_ContaOrcamentario"
         Tab(3).Control(4)=   "txt_ValorOrcamentario"
         Tab(3).Control(5)=   "cbo_DescricaoOrcamentario"
         Tab(3).Control(6)=   "cbo_ContaOrcamentario"
         Tab(3).ControlCount=   7
         TabCaption(4)   =   "Notas Fiscais"
         TabPicture(4)   =   "CadEmpenho.frx":407E
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "lbl_ValorTotal"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "lbl_Total"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "lbl_NotasFiscais"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "lbl_dblValorNF"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "lbl_dtmDataNF"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "lvw_NotasFiscais"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).Control(6)=   "txt_strNotasFiscais"
         Tab(4).Control(6).Enabled=   0   'False
         Tab(4).Control(7)=   "txt_dblValorNF"
         Tab(4).Control(7).Enabled=   0   'False
         Tab(4).Control(8)=   "txt_dtmDataNF"
         Tab(4).Control(8).Enabled=   0   'False
         Tab(4).ControlCount=   9
         Begin VB.ComboBox cbo_ContaOrcamentario 
            Height          =   315
            Left            =   -74280
            Sorted          =   -1  'True
            TabIndex        =   111
            Top             =   720
            Width           =   1725
         End
         Begin VB.ComboBox cbo_DescricaoOrcamentario 
            Height          =   315
            Left            =   -72600
            TabIndex        =   112
            Top             =   720
            Width           =   4875
         End
         Begin VB.TextBox txt_ValorOrcamentario 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -66840
            MaxLength       =   25
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   114
            Top             =   720
            Width           =   1425
         End
         Begin VB.CommandButton cmd_ContaOrcamentario 
            Height          =   300
            Left            =   -67740
            Picture         =   "CadEmpenho.frx":409A
            Style           =   1  'Graphical
            TabIndex        =   113
            Tag             =   "761"
            ToolTipText     =   "Clique para cadastar conta"
            Top             =   720
            Width           =   330
         End
         Begin VB.TextBox txt_dtmDataNF 
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
            Left            =   555
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   116
            Top             =   480
            Width           =   990
         End
         Begin VB.TextBox txt_dblValorNF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2055
            MaxLength       =   25
            MultiLine       =   -1  'True
            TabIndex        =   117
            Top             =   480
            Width           =   1425
         End
         Begin VB.TextBox txt_strNotasFiscais 
            Height          =   480
            Left            =   4755
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   118
            Text            =   "CadEmpenho.frx":4424
            Top             =   480
            Width           =   2925
         End
         Begin VB.CommandButton cmd_ContaRetencao 
            Height          =   300
            Left            =   -67740
            Picture         =   "CadEmpenho.frx":4436
            Style           =   1  'Graphical
            TabIndex        =   108
            Tag             =   "761"
            ToolTipText     =   "Clique para cadastar conta"
            Top             =   720
            Width           =   330
         End
         Begin VB.TextBox txt_ValorRetencao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -66840
            MaxLength       =   25
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   109
            Top             =   720
            Width           =   1425
         End
         Begin VB.ComboBox cbo_DescricaoRetencao 
            Height          =   315
            Left            =   -72600
            TabIndex        =   107
            Top             =   720
            Width           =   4875
         End
         Begin VB.ComboBox cbo_ContaRetencao 
            Height          =   315
            Left            =   -74280
            Sorted          =   -1  'True
            TabIndex        =   106
            Top             =   720
            Width           =   1725
         End
         Begin VB.CommandButton cmd_ContaExtra 
            Height          =   300
            Left            =   -67740
            Picture         =   "CadEmpenho.frx":47C0
            Style           =   1  'Graphical
            TabIndex        =   103
            Tag             =   "322"
            ToolTipText     =   "Clique para cadastar conta"
            Top             =   720
            Width           =   330
         End
         Begin VB.TextBox txt_ValorExtra 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   -66840
            MaxLength       =   25
            MultiLine       =   -1  'True
            TabIndex        =   104
            Top             =   720
            Width           =   1425
         End
         Begin VB.ComboBox cbo_DescricaoExtra 
            Height          =   315
            Left            =   -72600
            TabIndex        =   102
            Top             =   720
            Width           =   4875
         End
         Begin VB.ComboBox cbo_ContaExtra 
            Height          =   315
            Left            =   -74280
            Sorted          =   -1  'True
            TabIndex        =   101
            Top             =   720
            Width           =   1725
         End
         Begin MSComctlLib.ListView lvw_Liquidacao 
            Height          =   2970
            Left            =   -74970
            TabIndex        =   100
            Top             =   660
            Width           =   9585
            _ExtentX        =   16907
            _ExtentY        =   5239
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Parcela"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Data"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Vencimento"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Situação"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Nº OP"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Total OP"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Data Pagto"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Tipo"
               Object.Width           =   2028
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Liquidação"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Histórico"
               Object.Width           =   5027
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Flag"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "intEventoContabil"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Extra 
            Height          =   2520
            Left            =   -74970
            TabIndex        =   105
            Top             =   1110
            Width           =   9585
            _ExtentX        =   16907
            _ExtentY        =   4445
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Conta"
               Object.Width           =   2911
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descrição"
               Object.Width           =   11289
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "."
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Retencao 
            Height          =   2520
            Left            =   -74970
            TabIndex        =   110
            Top             =   1110
            Width           =   9585
            _ExtentX        =   16907
            _ExtentY        =   4445
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Conta"
               Object.Width           =   2911
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descrição"
               Object.Width           =   11289
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_NotasFiscais 
            Height          =   2640
            Left            =   30
            TabIndex        =   120
            Top             =   990
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   4657
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Data"
               Object.Width           =   2911
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Nota(s) Fiscal(is)"
               Object.Width           =   11289
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "."
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Orcamentario 
            Height          =   2520
            Left            =   -74970
            TabIndex        =   115
            Top             =   1110
            Width           =   9585
            _ExtentX        =   16907
            _ExtentY        =   4445
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Conta"
               Object.Width           =   2911
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descrição"
               Object.Width           =   11289
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lbl_valorOrcamentario 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   -67260
            TabIndex        =   234
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lbl_ContaOrcamentario 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   -74790
            TabIndex        =   233
            Top             =   750
            Width           =   420
         End
         Begin VB.Label lbl_dtmDataNF 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   150
            TabIndex        =   232
            Top             =   585
            Width           =   345
         End
         Begin VB.Label lbl_dblValorNF 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   1635
            TabIndex        =   231
            Top             =   585
            Width           =   360
         End
         Begin VB.Label lbl_NotasFiscais 
            AutoSize        =   -1  'True
            Caption         =   "Nota Fiscal"
            Height          =   195
            Left            =   3810
            TabIndex        =   230
            Top             =   585
            Width           =   795
         End
         Begin VB.Label lbl_Total 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Left            =   7740
            TabIndex        =   229
            Top             =   645
            Width           =   360
         End
         Begin VB.Label lbl_ValorTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00 "
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   8145
            TabIndex        =   119
            Top             =   600
            Width           =   1425
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   -74790
            TabIndex        =   169
            Top             =   750
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   -67260
            TabIndex        =   168
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lbl_ContaRetencao 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   -74790
            TabIndex        =   167
            Top             =   1050
            Width           =   420
         End
         Begin VB.Label lbl_valorExtra 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   -67260
            TabIndex        =   166
            Top             =   750
            Width           =   360
         End
         Begin VB.Label lbl_Conta 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   -74790
            TabIndex        =   165
            Top             =   750
            Width           =   420
         End
      End
      Begin VB.Frame fra_Parcela 
         Caption         =   " Parcelas "
         Height          =   1305
         Left            =   -74910
         TabIndex        =   155
         Top             =   1110
         Width           =   4545
         Begin VB.CommandButton cmd_Periodo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4110
            Picture         =   "CadEmpenho.frx":4B4A
            Style           =   1  'Graphical
            TabIndex        =   151
            ToolTipText     =   "Clique para cadastar período"
            Top             =   720
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.ComboBox cboPeriodo 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2550
            TabIndex        =   150
            Text            =   "cboPeriodo"
            Top             =   720
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.TextBox txt_ValorParcela 
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
            Left            =   2790
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   67
            Top             =   180
            Width           =   1665
         End
         Begin VB.TextBox txt_DataParcela 
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
            Left            =   780
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   66
            ToolTipText     =   "Previsão de pagamento"
            Top             =   180
            Width           =   990
         End
         Begin VB.TextBox txtNumParcela 
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
            Enabled         =   0   'False
            Height          =   285
            Left            =   1020
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   149
            Top             =   720
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lbl_Periodo 
            AutoSize        =   -1  'True
            Caption         =   "Período"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1920
            TabIndex        =   159
            Top             =   840
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lbl_ValorParcela 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   2310
            TabIndex        =   158
            Top             =   270
            Width           =   360
         End
         Begin VB.Label lbl_DataParcelamento 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   90
            TabIndex        =   157
            ToolTipText     =   "Previsão de pagamento"
            Top             =   270
            Width           =   345
         End
         Begin VB.Label lbl_NumeroParcelas 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   156
            Top             =   810
            Visible         =   0   'False
            Width           =   825
         End
      End
      Begin VB.CommandButton cmd_HistoricoComplemento 
         Height          =   300
         Left            =   -65700
         Picture         =   "CadEmpenho.frx":4ED4
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Clique para cadastar histórico"
         Top             =   1860
         Width           =   330
      End
      Begin VB.TextBox txt_DataComplemento 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   -74430
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   1080
         Width           =   990
      End
      Begin VB.ComboBox cbo_HistoricoComplemento 
         Height          =   315
         Left            =   -72240
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         TabIndex        =   83
         ToolTipText     =   "Histórico padrão"
         Top             =   1860
         Width           =   6495
      End
      Begin VB.Frame fra_HistoricoComplemento 
         Caption         =   " Histórico "
         Height          =   855
         Left            =   -72840
         TabIndex        =   154
         Top             =   960
         Width           =   7485
         Begin VB.TextBox txt_HistoricoComplemento 
            Height          =   675
            Left            =   0
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   81
            Top             =   180
            Width           =   7485
         End
      End
      Begin VB.Frame fra_HistoricoSubEmpenho 
         Caption         =   " Histórico "
         Height          =   855
         Left            =   -70260
         TabIndex        =   153
         Top             =   1110
         Width           =   4905
         Begin VB.TextBox txt_HistoricoSubEmpenho 
            Height          =   675
            Left            =   0
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Top             =   180
            Width           =   4905
         End
      End
      Begin VB.ComboBox cbo_HistoricoSubEmpenho 
         Height          =   315
         Left            =   -69750
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         TabIndex        =   75
         Text            =   "cboHistoricoSubEmpenho"
         ToolTipText     =   "Histórico padrão"
         Top             =   2070
         Width           =   4005
      End
      Begin VB.CommandButton cmd_HistoricoSubEmpenho 
         Height          =   300
         Left            =   -65700
         Picture         =   "CadEmpenho.frx":525E
         Style           =   1  'Graphical
         TabIndex        =   76
         Tag             =   "248"
         ToolTipText     =   "Clique para cadastar histórico"
         Top             =   2070
         Width           =   330
      End
      Begin MSComctlLib.ListView lvw_ListaSubempenho 
         Height          =   3570
         Left            =   -74940
         TabIndex        =   77
         Top             =   2460
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   6297
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parcela"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Data"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Situação"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Tipo"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Histórico"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Flag Situacao"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "intEventoContabil"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lvw_Complemento 
         Height          =   3570
         Left            =   -74940
         TabIndex        =   85
         Top             =   2250
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   6297
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parcela"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Data"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Situação"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Tipo"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Histórico"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Flag Situacao"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "intEventoContabil"
            Object.Width           =   0
         EndProperty
      End
      Begin TabDlg.SSTab tab_3DAnulacao 
         Height          =   5445
         Left            =   -74880
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   900
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   9604
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Principal"
         TabPicture(0)   =   "CadEmpenho.frx":55E8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txt_CodHistoricoAnl"
         Tab(0).Control(1)=   "frmSubElementoEstorno"
         Tab(0).Control(2)=   "fra_EventoContabil"
         Tab(0).Control(3)=   "cbo_HistoricoAnulacao"
         Tab(0).Control(4)=   "cmd_HistoricoAnulacao"
         Tab(0).Control(5)=   "txt_ValorAnulacao"
         Tab(0).Control(6)=   "fra_HistoricoAnulacao"
         Tab(0).Control(7)=   "txt_DataAnulucao"
         Tab(0).Control(8)=   "lvw_Anulacao"
         Tab(0).Control(9)=   "lbl_DataAnulacao"
         Tab(0).Control(10)=   "lbl_valorAnulacao"
         Tab(0).ControlCount=   11
         TabCaption(1)   =   "Itens"
         TabPicture(1)   =   "CadEmpenho.frx":5604
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblintCodigoTipoMaterialAnulacao"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lbldblQuantidadeAnulacao"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lblstrObsItemAnulacao"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lblstrdescricaodetalhadaAnulacao"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lblstrMarcaAnulacao"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "lvw_ItensAnulacao"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "dbc_intStrMarcaAnulacao"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "cbo_intCatalogoMaterialServicoAnulacao"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "cbo_intCodigoAnulacao"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txt_intUnidadedeMedidaAnulacao"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "txt_dblQuantidadeAnulacao"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "txt_strObsItemAnulacao"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "txt_dblValorEstimadoAnulacao"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txt_strdescricaodetalhadaAnulacao"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txt_PkidItemAnulacao"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).ControlCount=   15
         Begin VB.TextBox txt_CodHistoricoAnl 
            Height          =   315
            Left            =   -73020
            TabIndex        =   125
            Top             =   1260
            Width           =   555
         End
         Begin VB.Frame frmSubElementoEstorno 
            Caption         =   "Sub-Elemento"
            Height          =   1515
            Left            =   -74940
            TabIndex        =   245
            Top             =   1590
            Width           =   9405
            Begin VB.ComboBox dbcItemDespCodSubElementoEst 
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   128
               Top             =   180
               Width           =   1695
            End
            Begin VB.TextBox txtDblValorSubElementoEst 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   7860
               MaxLength       =   25
               MultiLine       =   -1  'True
               OLEDragMode     =   1  'Automatic
               TabIndex        =   131
               Top             =   210
               Width           =   1425
            End
            Begin VB.CommandButton Command1 
               Height          =   300
               Left            =   6720
               Picture         =   "CadEmpenho.frx":5620
               Style           =   1  'Graphical
               TabIndex        =   130
               Tag             =   "244"
               ToolTipText     =   "Clique para cadastar itens de despesa"
               Top             =   180
               Width           =   360
            End
            Begin MSComctlLib.ListView lvwSubElementoEst 
               Height          =   855
               Left            =   60
               TabIndex        =   132
               Top             =   600
               Width           =   9255
               _ExtentX        =   16325
               _ExtentY        =   1508
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
                  Text            =   "Código "
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Descrição"
                  Object.Width           =   11818
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   3
                  Text            =   "Valor"
                  Object.Width           =   2540
               EndProperty
            End
            Begin MSDataListLib.DataCombo dbcItemDespSubElementoEst 
               Height          =   315
               Left            =   3255
               TabIndex        =   129
               Top             =   180
               Width           =   3480
               _ExtentX        =   6138
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
               Height          =   195
               Left            =   7440
               TabIndex        =   247
               Top             =   255
               Width           =   360
            End
            Begin VB.Label lblItemSubElementoEst 
               AutoSize        =   -1  'True
               Caption         =   "Item de Despesa"
               Height          =   195
               Left            =   120
               TabIndex        =   246
               ToolTipText     =   "Item de despesa"
               Top             =   255
               Width           =   1200
            End
         End
         Begin VB.Frame fra_EventoContabil 
            Caption         =   " Evento Contabil "
            Height          =   690
            Left            =   -73020
            TabIndex        =   144
            Top             =   1650
            Width           =   7485
            Begin VB.CommandButton cmd_EventoAnul 
               Height          =   300
               Left            =   7020
               Picture         =   "CadEmpenho.frx":59AA
               Style           =   1  'Graphical
               TabIndex        =   147
               TabStop         =   0   'False
               Tag             =   "247"
               ToolTipText     =   "Clique para cadastar convênio"
               Top             =   225
               Width           =   330
            End
            Begin VB.TextBox txt_CodEventoAnul 
               Height          =   315
               Left            =   165
               MaxLength       =   15
               TabIndex        =   145
               Top             =   225
               Width           =   765
            End
            Begin VB.ComboBox cbo_intEventoAnul 
               Height          =   315
               Left            =   930
               TabIndex        =   146
               Top             =   225
               Width           =   6090
            End
         End
         Begin VB.TextBox txt_PkidItemAnulacao 
            Height          =   285
            Left            =   360
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   540
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txt_strdescricaodetalhadaAnulacao 
            Height          =   1155
            Left            =   1350
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   1710
            Width           =   7995
         End
         Begin VB.TextBox txt_dblValorEstimadoAnulacao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3135
            MaxLength       =   25
            TabIndex        =   138
            Top             =   885
            Width           =   1335
         End
         Begin VB.TextBox txt_strObsItemAnulacao 
            Height          =   330
            Left            =   1350
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   140
            Top             =   1290
            Width           =   7995
         End
         Begin VB.TextBox txt_dblQuantidadeAnulacao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1365
            MaxLength       =   10
            TabIndex        =   137
            Top             =   900
            Width           =   765
         End
         Begin VB.TextBox txt_intUnidadedeMedidaAnulacao 
            Height          =   315
            Left            =   6060
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   870
            Width           =   3285
         End
         Begin VB.ComboBox cbo_HistoricoAnulacao 
            Height          =   315
            Left            =   -72420
            OLEDragMode     =   1  'Automatic
            Sorted          =   -1  'True
            TabIndex        =   126
            ToolTipText     =   "Histórico padrão"
            Top             =   1260
            Width           =   6525
         End
         Begin VB.CommandButton cmd_HistoricoAnulacao 
            Height          =   300
            Left            =   -65880
            Picture         =   "CadEmpenho.frx":5D34
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "Clique para cadastar histórico"
            Top             =   1260
            Width           =   330
         End
         Begin VB.TextBox txt_ValorAnulacao 
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
            Left            =   -74490
            MaxLength       =   15
            MultiLine       =   -1  'True
            TabIndex        =   123
            Top             =   840
            Width           =   1395
         End
         Begin VB.Frame fra_HistoricoAnulacao 
            Caption         =   " Histórico "
            Height          =   855
            Left            =   -73020
            TabIndex        =   221
            Top             =   360
            Width           =   7485
            Begin VB.TextBox txt_HistoricoAnulacao 
               Height          =   675
               Left            =   0
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   124
               Top             =   180
               Width           =   7485
            End
         End
         Begin VB.TextBox txt_DataAnulucao 
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
            Left            =   -74490
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   122
            Top             =   480
            Width           =   990
         End
         Begin VB.ComboBox cbo_intCodigoAnulacao 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   450
            Width           =   855
         End
         Begin VB.ComboBox cbo_intCatalogoMaterialServicoAnulacao 
            Height          =   315
            Left            =   2340
            Style           =   2  'Dropdown List
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   450
            Width           =   4365
         End
         Begin MSDataListLib.DataCombo dbc_intStrMarcaAnulacao 
            Height          =   315
            Left            =   7350
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   450
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvw_ItensAnulacao 
            Height          =   1695
            Left            =   60
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   3090
            Width           =   9465
            _ExtentX        =   16695
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Pkid"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Código "
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Descrição"
               Object.Width           =   7937
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Marca"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Quantidade"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Vlr. Unitário"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Unid. Medida"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Observação"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Complemento"
               Object.Width           =   17639
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Anulacao 
            Height          =   2535
            Left            =   -74940
            TabIndex        =   148
            Top             =   2415
            Width           =   9420
            _ExtentX        =   16616
            _ExtentY        =   4471
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Parcela"
               Object.Width           =   1323
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Data"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Valor"
               Object.Width           =   1941
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Situação"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Tipo"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Anulação"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Histórico"
               Object.Width           =   3968
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Flag"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "intEventoContabil"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Emp.Anulação"
               Object.Width           =   2117
            EndProperty
         End
         Begin VB.Label lblstrMarcaAnulacao 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            Height          =   195
            Left            =   6810
            TabIndex        =   228
            Top             =   540
            Width           =   450
         End
         Begin VB.Label lblstrdescricaodetalhadaAnulacao 
            AutoSize        =   -1  'True
            Caption         =   "Complemento"
            Height          =   195
            Left            =   270
            TabIndex        =   227
            Top             =   1710
            Width           =   960
         End
         Begin VB.Label lblstrObsItemAnulacao 
            AutoSize        =   -1  'True
            Caption         =   "Observação"
            Height          =   195
            Left            =   390
            TabIndex        =   226
            Top             =   1350
            Width           =   870
         End
         Begin VB.Label lbldblQuantidadeAnulacao 
            AutoSize        =   -1  'True
            Caption         =   "Quantide                           Vlr. Unitário                                  Unidade de medida"
            Height          =   195
            Left            =   435
            TabIndex        =   225
            Top             =   960
            Width           =   5580
         End
         Begin VB.Label lblintCodigoTipoMaterialAnulacao 
            AutoSize        =   -1  'True
            Caption         =   "Item"
            Height          =   195
            Left            =   990
            TabIndex        =   224
            Top             =   510
            Width           =   300
         End
         Begin VB.Label lbl_DataAnulacao 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   -74895
            TabIndex        =   223
            Top             =   570
            Width           =   345
         End
         Begin VB.Label lbl_valorAnulacao 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   -74910
            TabIndex        =   222
            Top             =   930
            Width           =   360
         End
      End
      Begin VB.Label lblParcela 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   -74280
         OLEDropMode     =   1  'Manual
         TabIndex        =   86
         Tag             =   "1"
         Top             =   990
         Width           =   1065
      End
      Begin VB.Label lblCodEventoContabilLiq 
         AutoSize        =   -1  'True
         Caption         =   "Evento Contabil"
         Height          =   195
         Left            =   -70590
         TabIndex        =   249
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label lbl_DataComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   -74835
         TabIndex        =   172
         Top             =   1170
         Width           =   345
      End
      Begin VB.Label lbl_ValorComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -74850
         TabIndex        =   171
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label lbl_TotalComplemento 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   -74850
         TabIndex        =   170
         Top             =   1800
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblTotalComplemento 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   -74430
         OLEDropMode     =   1  'Manual
         TabIndex        =   80
         Tag             =   "1"
         Top             =   1740
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lbl_Parcela 
         AutoSize        =   -1  'True
         Caption         =   "Parcela                                           Data"
         Height          =   195
         Left            =   -74880
         TabIndex        =   164
         Top             =   1050
         Width           =   2820
      End
      Begin VB.Label lbl_ValorAux 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   -74700
         TabIndex        =   163
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label lblLiquido 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   -74280
         OLEDropMode     =   1  'Manual
         TabIndex        =   92
         Tag             =   "1"
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label lbl_Liquido 
         AutoSize        =   -1  'True
         Caption         =   "Líquido"
         Height          =   195
         Left            =   -74910
         TabIndex        =   162
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label lbl_DataLiquidacao 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   -72930
         TabIndex        =   161
         Top             =   1350
         Width           =   885
      End
      Begin VB.Label lbl_Numero 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   -71355
         TabIndex        =   152
         Top             =   3780
         Width           =   555
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1695
      Left            =   30
      TabIndex        =   143
      Top             =   6600
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   2990
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
      Columns(1).Caption=   "Número"
      Columns(1).DataField=   "intNumero"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Data"
      Columns(2).DataField=   "dtmData"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Valor"
      Columns(3).DataField=   "dblValor"
      Columns(3).NumberFormat=   "Standard"
      Columns(3).EditMaskUpdate=   -1  'True
      Columns(3).EditMaskRight=   -1  'True
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "CodigoReduzido"
      Columns(4).DataField=   "intCodigoReduzido"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Programa de Trabalho"
      Columns(5).DataField=   "strCodigo"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160664
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1508"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1429"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1826"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1746"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2355"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=8281"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=8202"
      Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
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
      MultiSelect     =   0
      DeadAreaBackColor=   13160664
      RowDividerColor =   13160664
      RowSubDividerColor=   13160664
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
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80000002&"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
      _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(61)  =   "Named:id=33:Normal"
      _StyleDefs(62)  =   ":id=33,.parent=0"
      _StyleDefs(63)  =   "Named:id=34:Heading"
      _StyleDefs(64)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   ":id=34,.wraptext=-1"
      _StyleDefs(66)  =   "Named:id=35:Footing"
      _StyleDefs(67)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   "Named:id=36:Selected"
      _StyleDefs(69)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(70)  =   "Named:id=37:Caption"
      _StyleDefs(71)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(72)  =   "Named:id=38:HighlightRow"
      _StyleDefs(73)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(74)  =   "Named:id=39:EvenRow"
      _StyleDefs(75)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(76)  =   "Named:id=40:OddRow"
      _StyleDefs(77)  =   ":id=40,.parent=33"
      _StyleDefs(78)  =   "Named:id=41:RecordSelector"
      _StyleDefs(79)  =   ":id=41,.parent=34"
      _StyleDefs(80)  =   "Named:id=42:FilterBar"
      _StyleDefs(81)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadEmpenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim strEvento                   As String
    Dim mblnAtivarPastas            As Boolean
    Dim mblnAbrindo                 As Boolean
    Dim mblnAlterandoHistorico      As Boolean
    Dim mblnClick                   As Boolean
    Dim mblnClickOk                 As Boolean
    Dim blnAlteraReserva            As Boolean
    Dim strUltimaData               As String
    Dim mblnDigitouQtdParcela       As Boolean
    Dim mblnAlterandoEmpenho        As Boolean
    Dim mblnAlterandoSubEmpenho     As Boolean
    Dim mblnAlterandoComplemento    As Boolean
    Dim mblnAlterandoNF             As Boolean
    Dim mobjAux                     As Object
    Dim mobjLista                   As Object
    Dim mblnAlterandoExtra          As Boolean
    Dim mblnAlterandoRetencao       As Boolean
    Dim mlngParcelaRetencao         As Long
    Dim mlngPKIdRetencao            As Long
    Dim mlngPKIdExtra               As Long
    Dim mlngPKIdOrcamentario        As Long
    Dim mlngEmpenhoAnulacao         As Long
    Dim mlngEmpenhoLiquidacao       As Long
    Dim mintCodigo                  As Integer
    Dim mdblValorExtra              As Double
    Dim mdblValorOrcamentario       As Double
    Dim mdblValorRetencao           As Double
    Dim mblnAtualizaTelaSubempenho  As Boolean
    Dim mblnselecionou              As Boolean
    Dim mstrNumero                  As String
    Dim mstrCodigo                  As String
    Dim mstrNFPkidExcluir           As String
    Dim mblnLimpaGrid               As Boolean
    Public mblnRestosAPagar         As Boolean
    Dim mblnEmpenhoEstorno          As Boolean
    Dim mblnPrimeiraVez             As Boolean
    Dim dtmDataReserva              As Date
    Dim blnItemDespesa              As Boolean
    Dim blnDataAutomatica           As Boolean
    Dim blnLiqAutomatica            As Boolean
    Dim mblnAlterandoOrcamentario   As Boolean
    Public strEmpInicial            As String
    Public strEmpFinal              As String
    Public strParcInicial           As String
    Public strParcFinal             As String
    Public blnSoEstorno             As Boolean
    Public blnAtivaFormImprime      As Boolean
    Public blnAlmoxarifado          As Boolean
    Public blnCompras               As Boolean
    Public blnFornecedor            As Boolean
    Public blnProcesso              As Boolean
    Public blnTesouraria            As Boolean
    Dim blnDesbilita2Click          As Boolean
    Dim intQuantItemAnulada         As Double
    Dim intQuantTotal               As Double
    Dim msgExtras                   As String
    Dim msgOrcamentario             As String
    Dim msgNotas                    As String
    Dim mblnCriarParcelaLiquidada   As Boolean
    Public intExercicioEmpenho      As Integer
    Dim blnMantemItemSel            As Boolean
    Dim blnMantemItemSelAnul        As Boolean
    Dim intItemIndex                As Integer
    Dim intItemIndexAnul            As Integer
    Dim dblSaldoItem                As Double
    Dim blnAlterandoItem            As Boolean
    
    Dim aryContas()
    Dim aryTpMov()
    Dim aryValor()
    
    Public mIntCodSeguranca         As Integer
    
    'Variaveis usadas para os Empenho importados de Compras
    Public blnImportadoPedidoEmpenho   As Boolean
    Public intNumPedidoEmpenho      As Long

Private Sub cbo_Historico_Change()
    txtstrHistorico = Trim(cbo_Historico.Text)
End Sub

Private Sub cbo_Historico_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 0
End Sub

Private Sub cbo_HistoricoAnulacao_Change()
    txt_HistoricoAnulacao = cbo_HistoricoAnulacao.Text
End Sub

Private Sub cbo_HistoricoAnulacao_Click()
    cbo_HistoricoAnulacao_Change
    txt_CodHistoricoAnl.Text = IIf(gstrItemData(cbo_HistoricoAnulacao) = 0, "", gstrItemData(cbo_HistoricoAnulacao))
End Sub

Private Sub cbo_HistoricoComplemento_Change()
    txt_HistoricoComplemento = Trim(cbo_HistoricoComplemento)
End Sub

Private Sub cbo_HistoricoComplemento_Click()
    cbo_HistoricoComplemento_Change
    txt_CodHistoricoComp.Text = IIf(gstrItemData(cbo_HistoricoComplemento) = 0, "", gstrItemData(cbo_HistoricoComplemento))
End Sub

Private Sub cbo_HistoricoLiquidacao_Change()
    txt_HistoricoLiquidacao = cbo_HistoricoLiquidacao.Text
End Sub

Private Sub cbo_HistoricoLiquidacao_Click()
    txt_HistoricoLiquidacao = cbo_HistoricoLiquidacao.Text
    txt_CodHistoricoLiq.Text = IIf(gstrItemData(cbo_HistoricoLiquidacao) = 0, "", gstrItemData(cbo_HistoricoLiquidacao))
End Sub

Private Sub cbo_HistoricoSubEmpenho_Click()
    txt_HistoricoSubEmpenho = cbo_HistoricoSubEmpenho.Text
    txt_CodHistoricoSub.Text = IIf(gstrItemData(cbo_HistoricoSubEmpenho) = 0, "", gstrItemData(cbo_HistoricoSubEmpenho))
End Sub

Private Sub cbo_intCatalogoMaterialServicoAnulacao_Click()
    cbo_intCodigoAnulacao.ListIndex = cbo_intCatalogoMaterialServicoAnulacao.ListIndex
End Sub

Private Sub cbo_intCodigoAnulacao_Click()
    CarregaItemAnulado True
    mAtivaPastaDeObjeto tab_3dPasta, 4, tab_3DAnulacao, 1
End Sub

Private Sub cbo_intevento_Click()
   leCodigoEvento txt_codEvento, cbo_intEvento
End Sub

Private Sub cbo_intEventoAnul_Click()
    leCodigoEvento txt_CodEventoAnul, cbo_intEventoAnul
    
End Sub

Private Sub cbo_intEventoAnul_GotFocus()
    If cbo_intEventoAnul.Text = "" Then txt_CodEventoAnul.Text = ""
End Sub

Private Sub cbo_intEventoAnul_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_intEventoAnul_LostFocus()
    If cbo_intEventoAnul.Text = "" Then txt_CodEventoAnul.Text = ""
End Sub

Private Sub cboCodigoReduzido_Click()
    
    cboProgramaTrabalho.ListIndex = gintIndiceCBO(cboProgramaTrabalho, _
                                    gstrItemData(cboCodigoReduzido))
                                    
    If cboProgramaTrabalho.ListIndex <> -1 Then
        preencheCboevento
    End If
                                    
End Sub

Private Sub cboCodigoReduzido_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DGeral, 0
End Sub

Private Sub cboCodigoReduzido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        preencheDotacaoByCodigo cboCodigoReduzido, cboProgramaTrabalho
        CaracterValido KeyAscii
        Exit Sub
    End If
   CaracterValido KeyAscii
   ProcuraTextoDigitado KeyAscii, cboCodigoReduzido, 1
End Sub

Private Sub cboCodigoReduzido_LostFocus()
    Dim strTextoDigitado As String
    Dim blnAchou         As Boolean
    Dim intContador      As Integer
    'preencheDotacaoByCodigo cboCodigoReduzido, cboProgramaTrabalho
    blnAchou = False
    With cboCodigoReduzido
        If .ListIndex = -1 Then
        strTextoDigitado = .Text
        For intContador = 0 To .ListCount - 1
            If .list(intContador) = strTextoDigitado Then
                If cboCodigoReduzido.Text <> "" Then
                    blnAchou = True
                    .ListIndex = intContador
                End If
            Exit For
            End If
        Next
        Else
        blnAchou = True
        End If
        
    End With
    If blnAchou = True Then
        cboProgramaTrabalho.ListIndex = gintIndiceCBO(cboProgramaTrabalho, gstrItemData(cboCodigoReduzido))
        preencheCboevento
    Else
        cboCodigoReduzido.Text = strTextoDigitado
        cboProgramaTrabalho.ListIndex = -1
    End If
End Sub

Private Sub cbo_ContaExtra_GotFocus()

   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
   VerificaTabAtivo
   If mblnselecionou = False Then
   mAtivaPastaDeObjeto tab_3dPasta, 3
   End If
End Sub

Private Sub cbo_DescricaoExtra_GotFocus()
   
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
   VerificaTabAtivo
End Sub

Private Sub cbointReservaDotacao_Click()
   'If Len(Trim(cbointReservaDotacao.Text)) > 0 Then
   If cbointReservaDotacao.ListIndex <> -1 Then
     PreencheDadosReserva
   Else
       txt_Reservado = ""
       txt_Cancelado = ""
       txt_Empenhado = ""
       txt_Saldo = ""
   End If
End Sub

Private Sub cbointReservaDotacao_KeyPress(KeyAscii As Integer)
    
   CaracterValido KeyAscii
   ProcuraTextoDigitado KeyAscii, cbointReservaDotacao, 1
End Sub

Private Sub cbo_ContaExtra_Click()
    cbo_DescricaoExtra.ListIndex = gintIndiceCBO(cbo_DescricaoExtra, _
                                  gstrItemData(cbo_ContaExtra))
End Sub

Private Sub cbo_ContaRetencao_Click()
    cbo_DescricaoRetencao.ListIndex = gintIndiceCBO(cbo_DescricaoRetencao, _
                                     gstrItemData(cbo_ContaRetencao))
End Sub

Private Sub cbo_DescricaoExtra_Click()
    cbo_ContaExtra.ListIndex = gintIndiceCBO(cbo_ContaExtra, _
                              gstrItemData(cbo_DescricaoExtra))
End Sub

Private Sub cbo_DescricaoRetencao_Click()
    cbo_ContaRetencao.ListIndex = gintIndiceCBO(cbo_ContaRetencao, _
                                     gstrItemData(cbo_DescricaoRetencao))
End Sub

Private Sub cbointReservaDotacao_LostFocus()
    Dim strTextoDigitado As String
    Dim blnAchou         As Boolean
    Dim intContador      As Integer
    'preencheDotacaoByCodigo cboCodigoReduzido, cboProgramaTrabalho
    blnAchou = False
    With cbointReservaDotacao
    strTextoDigitado = .Text
    If cbo_intEvento.ListIndex >= 0 Then LeTabelaReservaDotacao
        If .ListIndex = -1 Then
        For intContador = 0 To .ListCount - 1
            If .list(intContador) = strTextoDigitado Then
            blnAchou = True
            .ListIndex = intContador
            Exit For
            End If
        Next
        Else
        blnAchou = True
        End If
        
    End With
    If blnAchou = True And cbointReservaDotacao.ListCount > 0 Then
    preencheReservaDotacaoByCodigo cbointReservaDotacao.Text
    PreencheDadosReserva
    Else
       txt_Reservado = ""
       txt_Cancelado = ""
       txt_Empenhado = ""
       txt_Saldo = ""
       TrocaCorObjeto cboProgramaTrabalho, False
       TrocaCorObjeto cboCodigoReduzido, False
       cbointReservaDotacao.Text = ""
    End If
End Sub

Private Sub cboPeriodo_Click()
    If Val(txtNumParcela) > 1 And cboPeriodo.ListIndex > 0 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcular
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcular
    End If
End Sub

Private Sub cboProgramaTrabalho_Click()
    CalculaSaldoAtual
    
    
    If Not IsDate(txtDTMDATA) Then
        txt_SaldoDotacao = ""
        txt_TotalDotado = ""
    End If
    
    If cboProgramaTrabalho.ListIndex <> -1 And blnAlteraReserva Then
        preencheCboevento
    End If
    If blnAlteraReserva Then
        txt_Reservado = ""
        txt_Cancelado = ""
        txt_Empenhado = ""
        txt_Saldo = ""
        cbointReservaDotacao.Text = ""
    End If
End Sub

Private Sub cboProgramaTrabalho_DropDown()
   If Len(Trim(cboProgramaTrabalho)) > 0 Then
      If Left(cboProgramaTrabalho, 1) = "%" And Len(Trim(cboProgramaTrabalho)) > 1 Then
         LeTabelaProgramaTrabalho , cboProgramaTrabalho.Text
      End If
   End If
End Sub

Private Sub cboProgramaTrabalho_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cboProgramaTrabalho_LostFocus()

Dim strTextoDigitado As String
Dim lngPkidPrograma  As Long
Dim blnAchou         As Boolean
Dim intContador      As Integer
    
    blnAchou = False
    
    With cboProgramaTrabalho
        strTextoDigitado = .Text
        If .ListIndex <> -1 Then lngPkidPrograma = .ItemData(.ListIndex)
        If cbo_intEvento.ListIndex <> -1 Then LeTabelaProgramaTrabalho
    
        If .ListIndex = -1 Then
        For intContador = 0 To .ListCount - 1
            If lngPkidPrograma > 0 Then
                If .ItemData(intContador) = lngPkidPrograma Then
                    blnAchou = True
                    blnAlteraReserva = False
                    .ListIndex = intContador
                    Exit For
                End If
            Else
                If .list(intContador) = strTextoDigitado Then
                    blnAchou = True
                    blnAlteraReserva = False
                    .ListIndex = intContador
                    Exit For
                End If
            End If
        Next
        Else
            blnAchou = True
        End If
    End With
    
    If blnAchou = True Then
        cboCodigoReduzido.ListIndex = gintIndiceCBO(cboCodigoReduzido, gstrItemData(cboProgramaTrabalho))
        preencheCboevento
    Else
        cboProgramaTrabalho.Text = ""
        cboCodigoReduzido.ListIndex = -1
        txt_SaldoDotacao = ""
        txt_TotalDotado = ""
    End If
    
End Sub

Private Sub chc_LiquidarAutomaticamente_Click()
    If chc_LiquidarAutomaticamente.Value = Checked Then
        lbl_DataParcelamento.Top = 270
        txt_DataParcela.Top = 180
        lbl_ValorParcela.Top = 270
        txt_ValorParcela.Top = 180
        frm_liqAutomatico.Visible = True
    Else
        txt_ValorParcela.Top = 430
        txt_DataParcela.Top = 430
        lbl_ValorParcela.Top = 520
        lbl_DataParcelamento.Top = 520
        frm_liqAutomatico.Visible = False
    End If
    blnLiqAutomatica = CBool(Abs(chc_LiquidarAutomaticamente.Value))
End Sub

Private Sub cmd_ContaExtra_Click()
    CarregaForm frmCadPlanoConta, cbo_DescricaoExtra
End Sub

Private Sub cmd_ContaRetencao_Click()
    CarregaForm frmConPrevisaoDaReceita, cbo_DescricaoRetencao
End Sub

Private Sub cmd_Convenio_Click()
    CarregaForm frmCadConvenio, dbcintConvenio
End Sub

Private Sub cmd_Credor_Click()
    CarregaForm frmCadContribuinte, dbcintCredor
    frmCadContribuinte.Caption = "Cadastro de Credores"
    frmCadContribuinte.Tag = "Credor"
End Sub

Private Sub cmd_EventoAnul_Click()
    CarregaForm frmCadEvento, cbo_intEventoAnul, strQueryAplicarEventoAnul
End Sub

Private Sub cmd_Fundo_Click()
    CarregaForm frmCadFundo, dbcintFundo
End Sub

Private Sub cmd_Historico_Click()
    CarregaForm frmCadHistorico, cbo_Historico
End Sub

Private Sub cmd_HistoricoAnulacao_Click()
    CarregaForm frmCadHistorico, cbo_HistoricoAnulacao
End Sub

Private Sub cmd_HistoricoComplemento_Click()
    CarregaForm frmCadHistorico, cbo_HistoricoComplemento
End Sub

Private Sub cmd_HistoricoLiquidacao_Click()
    CarregaForm frmCadHistorico, cbo_HistoricoLiquidacao
End Sub

Private Sub cmd_HistoricoSubEmpenho_Click()
    CarregaForm frmCadHistorico, cbo_Historico
End Sub

Private Sub cmd_ItemDespesa_Click()
    CarregaForm frmCadItemDespesa, dbcintItemDespesa
End Sub

Private Sub cmd_Periodo_Click()
    CarregaForm frmCadPeriodo, cboPeriodo
End Sub

Private Sub cmd_ProgramaTrabalho_Click()
        LeTabelaProgramaTrabalho
        CarregaForm frmConProgramaDeTrabalho, cboProgramaTrabalho
End Sub

Private Sub cmd_Reserva_Click()
        LeTabelaReservaDotacao
        CarregaForm frmReservaDotacao, cbointReservaDotacao
End Sub

Private Sub cmd_Tipo_Click()
    CarregaForm frmCadTipoEmpenho, dbcintTipo
End Sub

Private Sub cbo_Historico_Click(Area As Integer)
    DropDownDataCombo cbo_Historico, Me, Area
    txt_CodHistorico.Text = IIf(gstrItemData(cbo_Historico) = 0, "", gstrItemData(cbo_Historico))
    txtstrHistorico.Text = cbo_Historico.Text
End Sub

Private Sub dbc_intStrMarca_GotFocus()
mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DGeral, 1
End Sub

Private Sub dbcintConvenio_Click(Area As Integer)
    leConvenio
End Sub

'Private Sub dbcintConvenio_Click(Area As Integer)
'    LeSaldoConvenio gstrItemData(dbcintConvenio), 0, txt_SaldoConvenio, txt_DataFinal
'End Sub

'Private Sub dbcintConvenio_GotFocus()
'    mAtivaPastaDeObjeto tab_3DPasta, 0, tab_3DEmpenho, 1
'End Sub

Private Sub dbcintConvenio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintConvenio
End Sub

Private Sub dbcintCredor_Click(Area As Integer)
    If Area = 2 Then
        txt_intNContribuinte = LeCDCCredor(dbcintCredor.BoundText)
        'txt_intNContribuinte = dbcintCredor.BoundText
    End If
    
End Sub


'Private Sub dbcintCredor_Click(Area As Integer)
'    DropDownDataCombo dbcintCredor, Me, Area
'End Sub

'Private Sub dbcintCredor_GotFocus()
'    mAtivaPastaDeObjeto tab_3DPasta, 0, tab_3DEmpenho, 0
'End Sub

'Private Sub dbcintCredor_KeyDown(KeyCode As Integer, Shift As Integer)
'    DropDownDataCombo dbcintCredor, Me, , KeyCode, Shift
'End Sub

Private Sub dbcintCredor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintFundo_GotFocus()
mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 1
End Sub

Private Sub dbcintFundo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbcintFundo
    
End Sub

Private Sub cbo_Historico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub


Private Sub dbcintItemDespesa_Change()
    If Val(dbcintItemDespesa.BoundText) > 0 Then
        If dbcintItemDespesa.MatchedWithList Then
            txt_intCodItemDespesa = LeCoditemDespesa(dbcintItemDespesa.BoundText)
            txt_intCodItemDespesa = gvntFormatacaoEspecifica(txt_intCodItemDespesa, 4)
        End If
    End If
       

    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 0
       
End Sub

Private Sub dbcintItemDespesa_Click(Area As Integer)
    If Area = 2 Then
       If Val(dbcintItemDespesa.BoundText) > 0 Then
          txt_intCodItemDespesa = LeCoditemDespesa(dbcintItemDespesa.BoundText)
       End If
       txt_intCodItemDespesa = gvntFormatacaoEspecifica(txt_intCodItemDespesa, 4)
       
    End If
End Sub

Private Sub dbcintItemDespesa_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 0
End Sub

Private Sub dbcintItemDespesa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintModalidade_GotFocus()
     mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 1
End Sub

Private Sub dbcintTipo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcItemDespSubElemento_Change()
    On Error Resume Next
    If Val(dbcItemDespSubElemento.BoundText) > 0 Then
        txtItemDespSubElemento = gvntFormatacaoEspecifica(LeCoditemDespesa(dbcItemDespSubElemento.BoundText), 4)
    Else
        txtItemDespSubElemento = ""
    End If
End Sub

Private Sub dbcItemDespSubElemento_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        txtItemDespSubElemento = gvntFormatacaoEspecifica(LeCoditemDespesa(dbcItemDespSubElemento.BoundText), 4)
    End If
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = 241
    VirificaGradeListView Me, mblnAlterandoEmpenho
       
    If mblnselecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    End If

    If mblnAlterandoEmpenho Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrDeletar
        'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir, gstrDeletar
        'HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
    End If
    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    
    VerificaTabAtivo
    
    If blnDataAutomatica = True Then
        DataAutomatica
    End If
    TrocaCorObjeto txt_intUnidadedeMedida, True
    TrocaCorObjeto txt_strdescricaodetalhada, True
    
    If mblnRestosAPagar Then
        TrocaCorObjeto txt_CodEventoAnul, False
        TrocaCorObjeto cbo_intEventoAnul, False
        TrocaCorObjeto cmd_EventoAnul, False
'        If mblnAlterandoEmpenho Then
'           HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
'        Else
'           HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar
'        End If
    Else
        TrocaCorObjeto txt_CodEventoAnul, True
        TrocaCorObjeto cbo_intEventoAnul, True
        TrocaCorObjeto cmd_EventoAnul, True
    End If
    'Alteração M4RC3LØ 14/03/2003 3 LN DOWN
    TrocaCorObjeto txtstrCodigo, False
    TrocaCorObjeto txtintExercicio, False
    TrocaCorObjeto txtbitDigito, False
    mblnAtivarPastas = True
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImportarDados
    mblnAtivarPastas = False
End Sub

Private Sub VerificaTabAtivo()
    
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrNovo, gstrSalvar, gstrImprimir, gstrPreencherLista, gstrLocalizar, gstrImportarDados
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar, gstrCancelar, gstrIncluirItem, gstrExcluirItem, gstrCalcular
    
    
    If tab_3dPasta.Tab = 0 Then
        If Not blnImportadoPedidoEmpenho Then
            If tab_3DGeral.Tab = 1 And Not mblnAlterandoEmpenho Then
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
            End If
        End If
        If tab_3DGeral.Tab = 0 And tab_3DEmpenho.Tab = 2 And mblnAlterandoEmpenho = False Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
        End If
    ElseIf tab_3dPasta.Tab = 1 Then
        TrocaCorObjeto cbo_HistoricoSubEmpenho, False
    ElseIf tab_3dPasta.Tab = 2 Then
        TrocaCorObjeto cbo_HistoricoComplemento, False
    ElseIf tab_3dPasta.Tab = 3 Then
        If tab_3DPastaLiquidacao.Tab = 0 And Not lvw_Liquidacao.SelectedItem Is Nothing Then
            If lvw_Liquidacao.SelectedItem.ListSubItems(4) = "Programada" Then
                HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
            End If
            If lvw_Liquidacao.SelectedItem.ListSubItems(4) = "Liquidada" Then
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCancelar
            End If
            If lvw_Liquidacao.SelectedItem.ListSubItems(4) = "Paga" Then
                HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar, gstrIncluirItem, gstrExcluirItem
            End If
        ElseIf tab_3DPastaLiquidacao.Tab = 1 Or tab_3DPastaLiquidacao.Tab = 3 Or tab_3DPastaLiquidacao.Tab = 4 Then
            If Not lvw_Liquidacao.SelectedItem Is Nothing Then
                If lvw_Liquidacao.SelectedItem.ListSubItems(4) = "Programada" Then
                    TrocaCorObjeto cbo_HistoricoLiquidacao, False
                    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem, gstrCancelar
                    TrocaCorObjeto txt_ValorExtra, False
                    TrocaCorObjeto txt_ValorOrcamentario, False
                    TrocaCorObjeto txt_dblValorNF, False
                ElseIf lvw_Liquidacao.SelectedItem.ListSubItems(4) = "Liquidada" Then
                    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem, gstrCancelar, gstrSalvar
                Else
                    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrNovo
                End If
            
            ElseIf mblnCriarParcelaLiquidada Then
               HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
            End If
        
        End If
    ElseIf tab_3dPasta.Tab = 4 Then
        TrocaCorObjeto cbo_HistoricoAnulacao, False
        If tab_3DAnulacao.Tab = 1 Then
            If Not lvw_Anulacao.SelectedItem Is Nothing Then
                If lvw_Anulacao.SelectedItem.SubItems(3) = "Programada" Then
                    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
                End If
            End If
        Else
            If Not lvw_Anulacao.SelectedItem Is Nothing Then
                If (lvw_Anulacao.SelectedItem.SubItems(3) = "Programada" Or lvw_Anulacao.SelectedItem.SubItems(3) = "Paga") And Val(txt_ValorAnulacao) <> 0 Then
                    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
                    If frmSubElementoEstorno.Visible = True Then
                        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem, gstrSalvar
                    End If
                End If
            End If
        End If
        
    End If
    
    
End Sub

Private Function strQuery()
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT EP.PKId, EP.intNumero , "
    strSQL = strSQL & "EP.dtmData,EP.dblValor , PT.intCodigoReduzido, PT.strCodigo "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
    strSQL = strSQL & "WHERE EP.intProgramaTrabalho = PT.PKId "
    If mblnRestosAPagar Then
       strSQL = strSQL & " AND EP.intExercicioRP = " & gintExercicio
       If Len(Trim(txtintExercicioEmpenho)) > 0 Then
          strSQL = strSQL & " AND EP.intExercicio =  " & txtintExercicioEmpenho
       End If
    Else
       'strSQL = strSQL & " AND EP.intExercicio = " & gintExercicio
       strSQL = strSQL & " AND PT.intExercicio = " & gintExercicio
    End If
    If cbointReservaDotacao.ListIndex > -1 Then
       strSQL = strSQL & " AND intReservaDotacao = " & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex)
    End If
    If cboCodigoReduzido.ListIndex > -1 Then
       strSQL = strSQL & " AND intProgramaTrabalho = " & cboProgramaTrabalho.ItemData(cboProgramaTrabalho.ListIndex)
    End If
    strSQL = strSQL & " ORDER BY EP.intNumero, EP.dtmData, PT.strCodigo "
    strQuery = strSQL
End Function

Private Function strQueryLocalizar()
    
    Dim strSQL  As String
    
    strSQL = ""
    
    strSQL = strSQL & "SELECT DISTINCT EP.PKId, EP.intNumero , "
    strSQL = strSQL & "EP.dtmData,EP.dblValor , PT.intCodigoReduzido, PT.strCodigo "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT, "
    strSQL = strSQL & gstrSubEmpenhoNF & " SUNF, "
    strSQL = strSQL & gstrSubempenho & " SU "

    strSQL = strSQL & "WHERE EP.intProgramaTrabalho = PT.PKId "
    strSQL = strSQL & " AND EP.Pkid = SU.Intempenho "
    strSQL = strSQL & " AND SU.Pkid " & strOUTJSQLServer & "= SUNF.intSubEmpenho " & strOUTJOracle
    
    If mblnRestosAPagar Then
       strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "EP.DTMDATA") & " < " & gintExercicio
       strSQL = strSQL & " AND EP.intExercicioRP  >= " & gintExercicio
       
       If Trim(txtintExercicioEmpenho) <> "" Then
            'orc1376
            'strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "EP.DTMDATA") & " = " & txtintExercicioEmpenho
            strSQL = strSQL & " AND EP.intExercicioEmpenho = " & txtintExercicioEmpenho
       End If
    
    Else
       'orc1376
       'strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "EP.DTMDATA") & " = " & gintExercicio
       strSQL = strSQL & " AND intExercicioEmpenho = " & gintExercicio
       strSQL = strSQL & " AND (EP.intExercicioRP IS NULL or EP.intExercicioRP > " & gintExercicio & ")"
        
    End If
    
    If cbointReservaDotacao.ListIndex > -1 And cbointReservaDotacao.ListCount > 0 Then
       strSQL = strSQL & " AND intReservaDotacao = " & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex)
    End If
    
    If cboCodigoReduzido.ListIndex > -1 Then
       strSQL = strSQL & " AND intProgramaTrabalho = " & cboProgramaTrabalho.ItemData(cboProgramaTrabalho.ListIndex)
    End If
    
    If cbo_intEvento.ListIndex > -1 Then
       strSQL = strSQL & " AND EP.intEvento = " & cbo_intEvento.ItemData(cbo_intEvento.ListIndex)
    End If
    
    'so coloca os campos de nota fiscal na consulta caso o usuário estiver na guia da nota
    If tab_3dPasta.Tab = 3 And tab_3DPastaLiquidacao.Tab = 4 Then
        If Trim(txt_dblValorNF.Text) <> "" Then
           strSQL = strSQL & " AND SUNF.dblValorNf = " & gstrConvVrParaSql(txt_dblValorNF)
        End If
        
        If Trim(txt_dtmDataNF.Text) <> "" Then
           strSQL = strSQL & " AND SUNF.dtmData = " & gstrConvDtParaSql(txt_dtmDataNF)
        End If
        
        If Trim(txt_strNotasFiscais.Text) <> "" Then
           strSQL = strSQL & " AND UPPER(SUNF.strNotaFiscal) LIKE ('" & UCase(Trim(txt_strNotasFiscais)) & "%')"
        End If
    End If
    
    strSQL = strSQL & " ORDER BY EP.intNumero, EP.dtmData, PT.strCodigo "
    
    strQueryLocalizar = strSQL
    
End Function

Private Function strQueryAplicar() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strContaContabil, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " "
    strSQL = strSQL & "WHERE  ABS(blnAnalitica) = 1"
    strQueryAplicar = strSQL
End Function

Private Sub LeTabelaEmpenho(Optional strFiltro As String)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT EP.dtmData, EP.PKId, EP.intNumero, "
    strSQL = strSQL & "EP.dblValor, PT.strCodigo, PT.intCodigoReduzido "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
    strSQL = strSQL & "WHERE EP.intProgramaTrabalho = PT.PKId "
    strSQL = strSQL & "AND PT.intExercicio " & IIf(mblnRestosAPagar = False, "=", "<") & gintExercicio & " "
    strSQL = strSQL & strFiltro & " "
    strSQL = strSQL & "ORDER BY EP.intNumero, EP.dtmData, PT.strCodigo "
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        Set tdb_Lista.DataSource = adoResultado
        tdb_Lista.Refresh
        If adoResultado.EOF = False Then
            adoResultado.MoveLast
            strUltimaData = gstrDataFormatada(adoResultado!DTMDATA)
            txtDTMDATA = strUltimaData
        End If
        
    End If
End Sub

Private Function strEmpenho() As String
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT MAX(intNumero) AS Empenho FROM " & gstrEmpenho
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            strEmpenho = adoResultado!Empenho
        End If
    End If
End Function

Private Sub GravaEmpenho()
    Dim strSQL        As String
    Dim strFiltro     As String
    Dim lngDigitadoEmpenho As Long  'Esta variável foi adicionada para evitar que
                                    ' durante a gravação o campo txtintNumero seja alterada por algum evento
                                    ' como lost focus
    
    'lngDigitadoEmpenho = CLng(txtintNumero.Text)
    lngDigitadoEmpenho = txtintNumero
    
    If gblnExclusaoGravacaoOk(IIf(mblnAlterandoEmpenho, "A", "I"), " do Empenho") Then
        If mblnAlterandoEmpenho Then
              strSQL = strQueryAlteraEmpenho
              Set gobjBanco = New clsBanco
              gobjBanco.Execute strSQL
              Exit Sub
        End If
        If blnDadosOk Then
        
            If Not mblnAlterandoEmpenho Then
ProximoCodigo:
                
                If gblnExisteCodigo(2, gstrEmpenho, "intNumero", CStr(lngDigitadoEmpenho), gstrDATEPART(strYEAR, "dtmData"), "'" & Val(IIf(mblnRestosAPagar = False, gintExercicio, txtintExercicioEmpenho)) & "'") Then
                    'mstrCodigo = (gstrProximoCodigo(txtintNumero, gstrEmpenho, "intNumero", gintCodSeguranca, gstrDATEPART(strYEAR, "dtmData"), Val(gintExercicio), , True))
                    mstrCodigo = GeraProximoDeEmpenho
                    If MsgBox("O número de empenho informado já se encontra cadastrado. Deseja usar o número " & mstrCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                        If txtintNumero.Enabled Then txtintNumero.SetFocus
                        Exit Sub
                    Else
                        txtintNumero.Text = mstrCodigo
                        lngDigitadoEmpenho = mstrCodigo
                        GoTo ProximoCodigo
                    End If
                Else
                    If gblnExisteCodigo(2, gstrSubempenho, "intEmpenhoAnulacao", CStr(lngDigitadoEmpenho), gstrDATEPART(strYEAR, "dtmData"), "'" & Val(IIf(mblnRestosAPagar = False, gintExercicio, txtintExercicioEmpenho)) & "'") Then
                        mstrCodigo = GeraProximoDeEmpenho
                        If MsgBox("O número de empenho informado já se encontra cadastrado. Deseja usar o número " & mstrCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                            If txtintNumero.Enabled Then txtintNumero.SetFocus
                            Exit Sub
                        Else
                            txtintNumero.Text = mstrCodigo
                            lngDigitadoEmpenho = mstrCodigo
                            GoTo ProximoCodigo
                        End If
                    End If
                End If
        
        
           End If
        
            strSQL = strQueryIncluiEmpenho
        
        
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSQL, True) Then
             
                If cbointReservaDotacao.ListIndex > -1 Then GravaReservaDotacaoLiberada
                
                If tab_3DEmpenho.TabVisible(2) = True Then GravaSubElementos glngRetornaPkidTabelaPai("seq" & gstrEmpenho, gstrEmpenho), , lvwSubElemento
            
                If lvw_Itens.ListItems.Count >= 1 Then
                    gobjBanco.Execute (StrSalvaIten(lngDigitadoEmpenho))
                End If
            
                'If lvw_Itens.ListItems.Count >= 1 Then
                '  StrSalvaIten (txtintNumero)
                'End If
            
                If mblnAlterandoEmpenho = False Then
                    gGravaHistoricoContribuinte strEmpenho, _
                                       CLng(gstrItemData(dbcintCredor)), _
                                       "Empenho Global", _
                                       "Orçamentário", _
                                       CDbl(gstrConvVrParaSql(txtdblValor))
                    If Not mblnRestosAPagar Then
                        If tab_3DEmpenho.TabVisible(2) = True Then 'grava contas de subElemento
                            If Not GeraMovimentosByEvento(gstrItemData(cbo_intEvento), txtDTMDATA, Str(CDbl(txtdblValor)), txtstrHistorico, CStr(lngDigitadoEmpenho), "3", aryContas, aryTpMov, , aryValor) Then
                              ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                            End If
                        Else
                            If Not GeraMovimentosByEvento(gstrItemData(cbo_intEvento), txtDTMDATA, Str(CDbl(txtdblValor)), txtstrHistorico, CStr(lngDigitadoEmpenho), "3") Then
                              ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                            End If
                        End If
                    End If
                End If
                
                
                mintCodigo = lngDigitadoEmpenho
                LimpaObjeto Me, mblnAlterandoEmpenho
                LimpaDadosReserva
                cbo_Historico.Text = ""
                txt_intNContribuinte = ""
                txtstrCodigo = ""
                txtbitDigito = ""
                txtintExercicio = ""
                
                If mblnLimpaGrid Then
                    mblnLimpaGrid = False
                    mstrNumero = Space$(0)
                    Set tdb_Lista.DataSource = Nothing
                End If
                
                mstrNumero = mstrNumero & mintCodigo & ","
                
                TrocaCorObjeto cbo_intEvento, False
                TrocaCorObjeto txt_codEvento, False
                txt_codEvento.Enabled = True
                txt_codEvento.BackColor = vbWindowBackground
                cbo_intEvento.ListIndex = -1
                
                If tdb_Lista.Text <> "" Then
                    strFiltro = " AND intNumero IN (" & Mid(mstrNumero, 1, Len(mstrNumero) - 1) & ")"
                Else
                    strFiltro = " AND intNumero = " & mintCodigo
                End If
                LeTabelaEmpenho (strFiltro)
                HabilitaDesabilitaTab tab_3dPasta, False
                
                'txtintNumero_GotFocus
                VerificaTabNovo
            Else
                If Not mblnAlterandoEmpenho Then
                'Somente irá sugerir o outro código
                ' se a operação não for alteração.
                    GoTo ProximoCodigo
                End If
            End If
        End If
    End If
End Sub

Sub HabilitaDesabilitaTab(mtab_3DPasta As SSTab, blnFlag As Boolean)
    Dim bytInd  As Byte
    For bytInd = 1 To mtab_3DPasta.Tabs - 1
        mtab_3DPasta.TabEnabled(bytInd) = blnFlag
    Next
    
    If mblnRestosAPagar Then
       mtab_3DPasta.TabEnabled(1) = mblnRestosAPagar
    End If
    
    
    txt_DataParcela.Visible = tab_3dPasta.TabEnabled(1)
    txt_dtmVenctoLiqAutomatica.Visible = tab_3dPasta.TabEnabled(1)
    txt_ValorParcela.Visible = tab_3dPasta.TabEnabled(1)
    txt_strNotasFiscaisLiqAutomatica.Visible = tab_3dPasta.TabEnabled(1)
    txt_codEventoLiqAutomatica.Visible = tab_3dPasta.TabEnabled(1)
    cbo_intEventoLiqAutomatica.Visible = tab_3dPasta.TabEnabled(1)
    cmd_EventoLiqAutomatica.Visible = tab_3dPasta.TabEnabled(1)
    txt_HistoricoSubEmpenho.Visible = tab_3dPasta.TabEnabled(1)
    cbo_HistoricoSubEmpenho.Visible = tab_3dPasta.TabEnabled(1)
    cmd_HistoricoSubEmpenho.Visible = tab_3dPasta.TabEnabled(1)
    lvw_ListaSubempenho.Visible = tab_3dPasta.TabEnabled(1)
    
    txt_DataComplemento.Visible = tab_3dPasta.TabEnabled(2)
    txt_ValorComplemento.Visible = tab_3dPasta.TabEnabled(2)
    lblTotalComplemento.Visible = tab_3dPasta.TabEnabled(2)
    txt_HistoricoComplemento.Visible = tab_3dPasta.TabEnabled(2)
    cbo_HistoricoComplemento.Visible = tab_3dPasta.TabEnabled(2)
    cmd_HistoricoComplemento.Visible = tab_3dPasta.TabEnabled(2)
    
    txt_dblValorAux.Visible = tab_3dPasta.TabEnabled(3)
    txt_DataLiuidacao.Visible = tab_3dPasta.TabEnabled(3)
    txt_dblDesconto.Visible = tab_3dPasta.TabEnabled(3)
    txt_codEventoLiq.Visible = tab_3dPasta.TabEnabled(3)
    cbo_intEventoLiq.Visible = tab_3dPasta.TabEnabled(3)
    cmd_EventoLiq.Visible = tab_3dPasta.TabEnabled(3)
    txt_HistoricoLiquidacao.Visible = tab_3dPasta.TabEnabled(3)
    cbo_HistoricoLiquidacao.Visible = tab_3dPasta.TabEnabled(3)
    cmd_HistoricoLiquidacao.Visible = tab_3dPasta.TabEnabled(3)
    lvw_Liquidacao.Visible = tab_3dPasta.TabEnabled(3)
    cbo_ContaExtra.Visible = tab_3dPasta.TabEnabled(3)
    cbo_DescricaoExtra.Visible = tab_3dPasta.TabEnabled(3)
    cmd_ContaExtra.Visible = tab_3dPasta.TabEnabled(3)
    txt_ValorExtra.Visible = tab_3dPasta.TabEnabled(3)
    lvw_Extra.Visible = tab_3dPasta.TabEnabled(3)
    cbo_ContaOrcamentario.Visible = tab_3dPasta.TabEnabled(3)
    cbo_DescricaoOrcamentario.Visible = tab_3dPasta.TabEnabled(3)
    cmd_ContaOrcamentario.Visible = tab_3dPasta.TabEnabled(3)
    txt_ValorOrcamentario.Visible = tab_3dPasta.TabEnabled(3)
    lvw_Orcamentario.Visible = tab_3dPasta.TabEnabled(3)
    txt_dtmDataNF.Visible = tab_3dPasta.TabEnabled(3)
    txt_dblValorNF.Visible = tab_3dPasta.TabEnabled(3)
    txt_strNotasFiscais.Visible = tab_3dPasta.TabEnabled(3)
    lvw_NotasFiscais.Visible = tab_3dPasta.TabEnabled(3)
    
    txt_DataAnulucao.Visible = tab_3dPasta.TabEnabled(4)
    txt_ValorAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    txt_HistoricoAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    cbo_HistoricoAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    cmd_HistoricoAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    dbcItemDespCodSubElementoEst.Visible = tab_3dPasta.TabEnabled(4)
    dbcItemDespSubElementoEst.Visible = tab_3dPasta.TabEnabled(4)
    Command1.Visible = tab_3dPasta.TabEnabled(4)
    txtDblValorSubElementoEst.Visible = tab_3dPasta.TabEnabled(4)
    lvwSubElementoEst.Visible = tab_3dPasta.TabEnabled(4)
    lvw_Anulacao.Visible = tab_3dPasta.TabEnabled(4)
    cbo_intCodigoAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    cbo_intCatalogoMaterialServicoAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    dbc_intStrMarcaAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    txt_dblQuantidadeAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    txt_dblValorEstimadoAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    txt_intUnidadedeMedidaAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    txt_strObsItemAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    txt_strdescricaodetalhadaAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    lvw_ItensAnulacao.Visible = tab_3dPasta.TabEnabled(4)
    
End Sub

Private Function blnDadosOk()
    Dim strSQL            As String
    Dim dtmDtEncerramento As Date
    Dim objControl        As Object
    Dim adoCampo          As ADODB.Field
    Dim adoResultado      As ADODB.Recordset
    Dim dblvalorAcumulado As Double
    Dim i                 As Integer
    Dim strTemp           As String
    Dim cdblSomatorio     As Double
    
    
    blnDadosOk = False
    If Val(Trim(txtintNumero)) < 0 Then
        ExibeMensagem "O Número do Empenho não pode ser negativo."
        If txtintNumero.Enabled Then txtintNumero.SetFocus
        Exit Function
    End If
    
    'Alterado pendencia orc1572
    If Len(Trim(cbointReservaDotacao)) > 0 Then
        'Pego o Saldo da Stored Procedure pois o saldo do campo não leva em consideração
        'lancamentos futuros.
        cdblSomatorio = gdblSaldoDotacaoAtual(cboCodigoReduzido.ItemData(cboCodigoReduzido.ListIndex)) + CDbl(gstrConvVrDoSql(IIf(Len(Trim(txt_Saldo.Text)) = 0, "0,00", txt_Saldo), 2))
        If cdblSomatorio < CDbl(gstrConvVrDoSql(IIf(Len(Trim(txtdblValor.Text)) = 0, "0,00", txtdblValor), 2)) Then
            MsgBox "O saldo da dotaçao somado ao saldo da reserva não pode ser inferior ao valor do empenho", vbInformation + vbOKOnly, "Saldo Dotação"
            If txtdblValor.Enabled Then txtdblValor.SetFocus
            Exit Function
        End If
    Else
        If Not VerificaSaldoDotacao Then
            ExibeMensagem "O Valor do empenho não pode ser maior do que o valor atual da dotação."
            If txtdblValor.Enabled Then txtdblValor.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtDTMDATA.Text) = "" Then
        ExibeMensagem "A Data do Empenho tem que ser informada."
        If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
        Exit Function
    End If
    
    If cbo_intEvento.ListIndex = -1 Then
        ExibeMensagem "O Evento Contabil tem que ser informado."
        If cbo_intEvento.Enabled Then cbo_intEvento.SetFocus
        Exit Function
    End If
    
    strEvento = cbo_intEvento.Text
    cbo_intevento_LostFocus
    
    strTemp = cboProgramaTrabalho.Text
    
    'cboProgramaTrabalho_LostFocus
    
    If cboProgramaTrabalho.Text <> strTemp Then
       cboProgramaTrabalho.Text = strTemp
       ExibeMensagem "Dotação inválida"
       If cboProgramaTrabalho.Enabled Then cboProgramaTrabalho.SetFocus
       Exit Function
    End If
    
    strTemp = cbointReservaDotacao.Text
    cbointReservaDotacao_LostFocus
    If cbointReservaDotacao.Text <> strTemp Then
       cbointReservaDotacao.Text = strTemp
       ExibeMensagem "Reserva inválida"
       If cbointReservaDotacao.Enabled Then cbointReservaDotacao.SetFocus
       Exit Function
    End If
   
    If cboProgramaTrabalho.ListIndex = -1 Then
        ExibeMensagem "Dotação inválida"
        Exit Function
    End If
        
    If Trim(txtdblValor.Text) = "" Then
        ExibeMensagem "O Valor do Empenho tem que ser informado."
        If txtdblValor.Enabled Then txtdblValor.SetFocus
        Exit Function
    End If
    
    If Not dbcintTipo.MatchedWithList Then
        ExibeMensagem "Tipo de empenho inválido."
        If dbcintTipo.Enabled Then dbcintTipo.SetFocus
        Exit Function
    End If
       
    If blnValidarProcesso Then
       If Len(Trim(txtstrCodigo)) > 0 Or Len(Trim(txtbitDigito)) > 0 Or Len(Trim(txtintExercicio)) > 0 Then
          If Not VerificaEmpenhoProcesso(Trim(txtstrCodigo), Val(txtbitDigito), Val(txtintExercicio)) Then
             ExibeMensagem "Processo não localizado."
             If txtstrCodigo.Enabled Then txtstrCodigo.SetFocus
             Exit Function
          End If
       End If
    End If
    
    If mblnRestosAPagar Then
    
        If Trim(txtintNumero) = "" Then
            ExibeMensagem "É necessário digitar um número"
            txtintNumero.SetFocus
            Exit Function
        End If
    
        If Val(txtintExercicioEmpenho) >= gintExercicio Then
            ExibeMensagem "É necessário que o Exercício informado seja inferior ao exercício vigente"
            Exit Function
        End If
        
        If Not Year(txtDTMDATA) = txtintExercicioEmpenho Then
            ExibeMensagem "A Data informada não pertence ao Exercício Informado"
            txtDTMDATA.SetFocus
            Exit Function
        End If
        
        strSQL = "SELECT * FROM " & gstrEmpenho & " WHERE intNumero = " & txtintNumero & _
                 " AND " & gstrDATEPART("YYYY", "dtmData") & " = " & txtintExercicioEmpenho
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If Not adoResultado.EOF And Not adoResultado.BOF Then
                ExibeMensagem "Já existe um empenho de Número " & txtintNumero & " no exercício " & txtintExercicioEmpenho
                Exit Function
            End If
        End If
        
    Else

        'Orc677
        If Right(txtDTMDATA, 4) <> gintExercicio Then
            ExibeMensagem "A data do empenho não equivale ao exercício corrente."
            If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
            Exit Function
        End If

    End If
    
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM " & gstrEmpenho & " "
    strSQL = strSQL & "WHERE PKID = (SELECT MAX(PKId) FROM " & gstrEmpenho & ")"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            For Each adoCampo In adoResultado.Fields
                For Each objControl In Me.Controls
                  If UCase(Mid(adoCampo.Name, 4)) = UCase(Trim(Mid(objControl.Name, 4))) _
                       And adoCampo.Name <> "PKId" _
                       And adoCampo.Name <> "intNumero" Then
                       If CBool(adoCampo.Attributes And adFldIsNullable) = False Then
                          If Trim(objControl) = "" Then
                              ExibeMensagem "Dado incorreto. A coluna '" & Mid(adoCampo.Name, 4) _
                                          & "' não aceita nulo."
                            
                               If objControl.Enabled Then
                                  If objControl.Enabled Then objControl.SetFocus
                               End If
                               Exit Function
                           End If
                       End If
                   End If
                   
                Next
            Next
        End If
    End If
    
    If Not mblnAlterandoEmpenho Then
        If cbointReservaDotacao.ListIndex < 0 Then
           
           If Not mblnRestosAPagar Then
               If SaldoDotacaoAtual(gstrItemData(cboProgramaTrabalho, True), Val(Month(CDate(txtDTMDATA))), IIf(mblnRestosAPagar = False, gintExercicio, txtintExercicioEmpenho), CDbl(txtdblValor)) = Empty Then
                     If txtdblValor.Enabled Then txtdblValor.SetFocus
                     Exit Function
               End If
        
               CalculaSaldoAtual
        
               If Val(gstrConvVrParaSql(txtdblValor)) > Val(gstrConvVrParaSql(txt_SaldoDotacao)) And mblnAlterandoEmpenho = False Then
                   ExibeMensagem "O valor do Empenho não poder ser superior ao Saldo da Dotação."
                   If txtDTMDATA.Enabled Then
                       txtdblValor.SetFocus
                   End If
                   Exit Function
               End If
           End If
        Else
        
           If dtmDataReserva > CDate(txtDTMDATA) Then
              ExibeMensagem "A data do empenho não pode ser menor que a data da reserva."
              Exit Function
           End If
           
           'Alterado para quando o o empenho for maior que o sldo da reserva
           If Val(gstrConvVrParaSql(txtdblValor)) > Val(gstrConvVrParaSql(txt_Saldo)) And mblnAlterandoEmpenho = False Then
                If SaldoDotacaoAtual(gstrItemData(cboProgramaTrabalho, True), Val(Month(CDate(txtDTMDATA))), gintExercicio, CDbl(Val(gstrConvVrParaSql(txtdblValor)) - Val(gstrConvVrParaSql(txt_Saldo)))) = Empty Then
                    If txtdblValor.Enabled Then
                        If txtdblValor.Enabled Then txtdblValor.SetFocus
                    End If
                    Exit Function
                End If
           End If
           
        End If
    End If
    If txtdblValor.Text = "" Then
        ExibeMensagem "O valor tem que ser informado."
        If txtdblValor.Enabled Then txtdblValor.SetFocus
        Exit Function
    End If
    
    'Orc677
    If IsDate(txtdtmHomologacao.Text) Then
        If Right(txtdtmHomologacao, 4) <> gintExercicio Then
            ExibeMensagem "A data de homologação não equivale ao exercício corrente."
            If txtdtmHomologacao.Enabled Then txtdtmHomologacao.SetFocus
            Exit Function
        End If
    End If
    
    Set gobjBanco = New clsBanco
    
    strSQL = ""
    strSQL = strSQL & "Select bytItemDespesaObrigatorio from " & gstrConfiguracaoGeral & " Where Pkid = 1 "
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                blnItemDespesa = IIf(Not IsNull(!bytItemDespesaObrigatorio), .Fields("bytItemDespesaObrigatorio").Value, False)
            End If
        End With
    End If
    If blnItemDespesa = True Then
    
    
        If tab_3DEmpenho.TabVisible(2) = True Then
            For i = 1 To lvwSubElemento.ListItems.Count
                dblvalorAcumulado = dblvalorAcumulado + Val(gstrConvVrParaSql(lvwSubElemento.ListItems(i).SubItems(3)))
            Next
            
            If Val(gstrConvVrParaSql(dblvalorAcumulado)) <> Val(gstrConvVrParaSql(txtdblValor)) Then
                ExibeMensagem "O valor acumulado dos Sub-Elmentos tem que ser o mesmo do Empenho."
                tab_3DEmpenho.Tab = 2
                If txtItemDespSubElemento.Enabled Then txtItemDespSubElemento.SetFocus
                Exit Function
            End If
              
            If GeraArraysSubElmento(gstrItemData(cbo_intEvento), False, lvwSubElemento) = False Then
                Exit Function
            End If
        Else
    
            If Not dbcintItemDespesa.MatchedWithList Then
                ExibeMensagem "Item de Despesa inválido."
                If dbcintItemDespesa.Enabled Then dbcintItemDespesa.SetFocus
                Exit Function
            End If
'            If dbcintItemDespesa.Text = "" Then
'                ExibeMensagem "O Item de despesa tem que ser informado."
'                tab_3DEmpenho.Tab = 0
'                If dbcintItemDespesa.Enabled Then dbcintItemDespesa.SetFocus
'                Exit Function
'            End If
        End If
    End If
    If dbcintCredor.Text = "" Then
        ExibeMensagem "O Credor tem que ser informado."
        tab_3DEmpenho.Tab = 0
        If dbcintCredor.Enabled Then dbcintCredor.SetFocus
        Exit Function
    End If
    
    If dbcintTipo.Text = "" Then
        ExibeMensagem "O Tipo tem que ser informado."
        tab_3DEmpenho.Tab = 0
        If dbcintTipo.Enabled Then dbcintTipo.SetFocus
        Exit Function
    End If
    
    If gblnDataValida(txtDTMDATA) = False Then
        ExibeMensagem "A data do Empenho tem que ser informada corretamente."
        If txtDTMDATA.Enabled Then
            If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
        End If
        Exit Function
    ElseIf Not VerificaAdiantamentos Then
       Exit Function
    ElseIf mblnAlterandoEmpenho = False And (Year(txtDTMDATA) <> gintExercicio) And Not mblnRestosAPagar Then
        ExibeMensagem "A data do empenho tem que estar dentro do ano de " & gintExercicio & "."
        If txtDTMDATA.Enabled Then
            If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
        End If
        Exit Function
    ElseIf dbcintConvenio.MatchedWithList Then
        If Val(gstrConvVrParaSql(txtdblValor)) > Val(gstrConvVrParaSql(txt_SaldoConvenio)) Then
            ExibeMensagem "O valor do Empenho não poder ser superior ao Saldo do Convênio."
            If txtDTMDATA.Enabled Then
                If txtdblValor.Enabled Then txtdblValor.SetFocus
            End If
            Exit Function
        ElseIf gblnDataValida(txtDTMDATA) Then
            If (CVDate(txtDTMDATA) > CVDate(txt_DataFinal)) And mblnAlterandoEmpenho = False Then
                ExibeMensagem "A data do Empenho não pode ser superior à data final do convênio."
                If txtDTMDATA.Enabled Then
                    If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
                End If
                Exit Function
            End If
        End If
        
    ElseIf gblnDataValida(strUltimaData) Then
        If (CVDate(strUltimaData) > CVDate(txtDTMDATA)) And mblnAlterandoEmpenho = False Then
            ExibeMensagem "A data do Empenho não pode ser inferior a data do último empenho."
            If txtDTMDATA.Enabled Then
                If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
            End If
            Exit Function
        End If
    End If
    If lvw_Itens.ListItems.Count >= 1 Then
        'If Trim(txtstrCondPagto.Text) = "" Then
            'ExibeMensagem "A Condição de Pagamento é obrigatória quando é incluso Itens no Empenho"
            'If txtstrCondPagto.Enabled Then txtstrCondPagto.SetFocus
            'Exit Function
        'End If
        'If Trim(txtstrLocEntrega.Text) = "" Then
        '    ExibeMensagem "O o local de entrega é obrigatório quando é incluso Itens no Empenho"
        '    If txtstrLocEntrega.Enabled Then txtstrLocEntrega.SetFocus
        '    Exit Function
        'End If
        'If Trim(txtstrPrazoEntrega.Text) = "" Then
        '    ExibeMensagem "O Prazo de Entrega é obrigatório quando é incluso Itens no Empenho"
        '    If txtstrPrazoEntrega.Enabled Then txtstrPrazoEntrega.SetFocus
        '    Exit Function
        'End If
    End If
    
    If lvw_Itens.ListItems.Count >= 1 Then
        If VlTotalItem <> CCur(txtdblValor) Then
            ExibeMensagem "O valor total de Itens deve ser igual o valor do Empenho"
            tab_3DGeral.Tab = 1
            Exit Function
        End If
    End If
    
    If Not mblnRestosAPagar Then
        dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
            
        If dtmDtEncerramento = Empty Then
           Exit Function
        Else
            If Not Val(txtPKId) <> 0 Then
                If CDate(txtDTMDATA) <= dtmDtEncerramento Then
                    ExibeMensagem "A data do Empenho deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
                    If txtDTMDATA.Enabled Then txtDTMDATA.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    'Orc1168
    'Valor de Empenho não pode ser maior que Total Reserva Dotação
    'If Val(gstrConvVrParaSql(Me.txtdblValor.Text)) >= Val(gstrConvVrParaSql(Me.txt_Reservado.Text)) Then
    'If Val(Me.txtdblValor.Text) > Val(Me.txt_Reservado.Text) Then
    '    ExibeMensagem "Não é permitido que o Valor de Empenho seja" + vbCrLf + _
                      "maior ou igual que o Valor da Reserva Dotação."
    '    Exit Function
    'End If
    
    blnDadosOk = True
End Function

Private Sub DeletaEmpenho()
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "DELETE " & gstrEmpenho & " "
    strSQL = strSQL & "WHERE PKId = " & txtPKId
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSQL
End Sub

Private Function strQueryIncluiEmpenho() As String

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL       As String

    strSQL = ""
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    strSQL = strSQL & "INSERT INTO " & gstrEmpenho & " ("
    strSQL = strSQL & "intNumero, intExercicioEmpenho, dtmData, intProgramaTrabalho, dblValor, "
    strSQL = strSQL & "intTipo, strContrato, strEmbasamento, strLicitacao, "
    strSQL = strSQL & "strSolicitacao, strModalidade, dtmHomologacao, "
    strSQL = strSQL & "intFundo, intConvenio, intItemDespesa, intCredor, "
    strSQL = strSQL & "strHistorico, dtmDtAtualizacao, intevento ,lngCodUsr, "
    
    If mblnRestosAPagar Then
        strSQL = strSQL & "intExercicioRP, "
    End If
    
    strSQL = strSQL & "intReservaDotacao, intModalidade, strCodigo, bitDigito, STRCONDPAGTO, STRLOCENTREGA, STRPRAZOENTREGA, intExercicio) "
    


'    strSql = strSql & "SELECT ISNULL(MAX(intNumero), 0) + 1, "
    strSQL = strSQL & " VALUES (" & txtintNumero & ", "
    strSQL = strSQL & txtintExercicioEmpenho & ", "
    strSQL = strSQL & gstrConvDtParaSql(txtDTMDATA) & ", "
    strSQL = strSQL & gstrItemData(cboProgramaTrabalho, True) & ", "
    strSQL = strSQL & gstrConvVrParaSql(txtdblValor) & ", "
    strSQL = strSQL & gstrItemData(dbcintTipo) & ", "
    strSQL = strSQL & "'" & txtstrContrato & "', "
    strSQL = strSQL & "'" & txtstrEmbasamento & "', "
    strSQL = strSQL & "'" & txtstrLicitacao & "', "
    strSQL = strSQL & "'" & txtstrsolicitacao & "', "
    strSQL = strSQL & "'" & txtstrModalidade & "', "
    strSQL = strSQL & gstrConvDtParaSql(txtdtmHomologacao) & ", "
    strSQL = strSQL & gstrItemData(dbcintFundo, True) & ", "
    strSQL = strSQL & gstrItemData(dbcintConvenio, True) & ", "
    'strSQL = strSQL & gstrItemData(dbcintItemDespesa, True) & ", "
    strSQL = strSQL & IIf(Trim(dbcintItemDespesa.Text) <> "", gstrItemData(dbcintItemDespesa, True), "NULL") & ", "
    
    strSQL = strSQL & gstrItemData(dbcintCredor, True) & ", "
    strSQL = strSQL & "'" & txtstrHistorico & "', "
    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSQL = strSQL & gstrItemData(cbo_intEvento) & ", "
    strSQL = strSQL & glngCodUsr & ", "
    
    If mblnRestosAPagar Then
        strSQL = strSQL & gintExercicio & ", "
    End If
    
    If cbointReservaDotacao.ListIndex < 0 Then
       strSQL = strSQL & "NULL, "
    Else
       strSQL = strSQL & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex) & ","
    End If
    strSQL = strSQL & gstrItemData(dbcintModalidade, True) & ", "
    strSQL = strSQL & "'" & Trim(txtstrCodigo) & "', "
    strSQL = strSQL & IIf(Trim(txtbitDigito) <> "", Trim(txtbitDigito), "NULL") & ", "
    strSQL = strSQL & "'" & IIf(Trim(txtStrcondpagto.Text) <> "", txtStrcondpagto.Text, "") & "', "
    strSQL = strSQL & "'" & IIf(Trim(txtStrlocentrega.Text) <> "", txtStrlocentrega.Text, "") & "', "
    strSQL = strSQL & "'" & IIf(Trim(txtStrprazoentrega.Text) <> "", txtStrprazoentrega.Text, "") & "', "
    strSQL = strSQL & IIf(Trim(txtintExercicio) <> "", Trim(txtintExercicio), "NULL") & ");"
    
    
    CalculaSaldoAtual
    
    strSQL = strSQL & " UPDATE " & gstrSubempenho
    strSQL = strSQL & " SET dblEmpenhadoAteData = " & gstrConvVrParaSql(CDbl(txt_TotalDotado) + CDbl(txtdblValor))
    'strSql = strSql & ", dblSaldoAtual = " & gstrConvVrParaSql(CDbl(txt_SaldoDotacao) - CDbl(txtdblValor))
    strSQL = strSQL & ", dblSaldoAtual = " & gstrConvVrParaSql(SaldoDotacaoSoEmpenho(gstrItemData(cboCodigoReduzido), txtDTMDATA.Text) - CDbl(txtdblValor))
    
    
    strSQL = strSQL & " WHERE intNumero = 0 AND "
    strSQL = strSQL & " intEmpenho = (SELECT MAX(PKID) FROM " & gstrEmpenho & " );"
    

    
    If blnImportadoPedidoEmpenho Then
       strSQL = strSQL & SalvaEmpenhoCompras(txtintNumero)
       blnImportadoPedidoEmpenho = False
    End If
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    
    strQueryIncluiEmpenho = strSQL

End Function

Private Function strQueryAlteraEmpenho() As String
    Dim strSQL As String
    strSQL = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
        
    strSQL = strSQL & "UPDATE " & gstrEmpenho & " SET "
    strSQL = strSQL & "strContrato = '" & txtstrContrato & "', "
    strSQL = strSQL & "strLicitacao = '" & txtstrLicitacao & "', "
    strSQL = strSQL & "strModalidade = '" & txtstrModalidade & "', "
    strSQL = strSQL & "strEmbasamento = '" & txtstrEmbasamento & "', "
    strSQL = strSQL & "strSolicitacao = '" & txtstrsolicitacao & "', "
    strSQL = strSQL & "strHistorico = '" & txtstrHistorico & "', "
    strSQL = strSQL & "intModalidade = " & gstrItemData(dbcintModalidade, True) & ", "
    strSQL = strSQL & "intCredor = " & gstrItemData(dbcintCredor, True) & ", "
    strSQL = strSQL & "intTipo = " & gstrItemData(dbcintTipo) & ", "
    strSQL = strSQL & "intFundo = " & gstrItemData(dbcintFundo, True) & ", "
    strSQL = strSQL & "intConvenio = " & gstrItemData(dbcintConvenio, True) & ", "
    
    'strSQL = strSQL & "intItemDespesa = " & gstrItemData(dbcintItemDespesa, True) & ", "
    strSQL = strSQL & "intItemDespesa = " & IIf(Trim(dbcintItemDespesa.Text) <> "", gstrItemData(dbcintItemDespesa, True), "NULL") & ", "
    
    strSQL = strSQL & "dtmHomologacao = " & gstrConvDtParaSql(txtdtmHomologacao) & ", "
    strSQL = strSQL & "intevento = " & gstrItemData(cbo_intEvento) & ", "
    strSQL = strSQL & "strCodigo = '" & Trim(txtstrCodigo) & "', "
    strSQL = strSQL & "bitDigito = '" & Trim(txtbitDigito) & "', "
    strSQL = strSQL & "intExercicio = '" & Trim(txtintExercicio) & "', "
    
    strSQL = strSQL & "strCondPagto = '" & Trim(txtStrcondpagto) & "', "
    strSQL = strSQL & "strLocEntrega = '" & Trim(txtStrlocentrega) & "' , "
    strSQL = strSQL & "strPrazoEntrega = '" & Trim(txtStrprazoentrega) & "' "
    strSQL = strSQL & " WHERE PKID = " & tdb_Lista.Columns(0).Value & "; "
    
    strSQL = strSQL & " UPDATE " & gstrSubempenho & " SET strHistorico = '" & txtstrHistorico & "' "
    strSQL = strSQL & "WHERE intEmpenho = " & tdb_Lista.Columns(0).Value & " AND "
    strSQL = strSQL & "intNumero = 0 AND bytSituacao = 1 "
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; END; ", "")
    
    strQueryAlteraEmpenho = strSQL
End Function

Private Sub VerificaTabGravacao()
    tab_3dPasta.SetFocus
    Select Case tab_3dPasta.Tab
    Case 0 'Empenho
'        If Not mblnRestosAPagar Then
           GravaEmpenho
'        Else
          'AlteraEmpenho no formulário Restos a Pagar
'          AlteraEmpenho
'        End If
    Case 1 'Subempenho (parcela)
        GravaSubEmpenho
    Case 2 'Complemento de empenho
        GravaComplemento
    Case 3 'Liquidação de parcela
        If tab_3dPasta.TabEnabled(2) = True Then
            VerificaGravacaoLiquidacao
        End If
    Case 4
        GravaAnulacao
        
    End Select
    
    LeSubEmpenho lvw_Anulacao, 4, txtPKId
    LeSubEmpenho lvw_ListaSubempenho, , txtPKId
    LeSubEmpenho lvw_Liquidacao, 2, txtPKId
    
    If blnMantemItemSel = True Then
        If intItemIndex = -1 Or intItemIndex > lvw_Liquidacao.ListItems.Count Then Exit Sub
        lvw_Liquidacao.ListItems(intItemIndex).Selected = True
        blnMantemItemSel = False
    End If
    
    If blnMantemItemSelAnul = True Then
        If intItemIndexAnul = -1 Or intItemIndexAnul > lvw_Anulacao.ListItems.Count Then Exit Sub
        lvw_Anulacao.ListItems(intItemIndexAnul).Selected = True
        blnMantemItemSelAnul = False
    End If
    
End Sub

Private Function blnDadosAnulacaoOk() As Boolean
    Dim dtmDtEncerramento As Date
    Dim i                 As Integer
    Dim valorTotalAnulado As Double
    Dim strSQL            As String
    Dim adoResultado      As ADODB.Recordset
    Dim dblvalorAcumulado  As Double
    
    
    intItemIndexAnul = lvw_Anulacao.SelectedItem.Index
    blnMantemItemSelAnul = True 'variavel usada para manter o grid selecionado neste item, pois o mesmo é atualizado e o item selecionado muda
    
    'VERIFICA SE ESTA ANULADA E PERMITE A ALTERÇÃO DO HISTÓRICO
    If UCase(lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).SubItems(3)) = "CANC. RP" Or UCase(lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).SubItems(3)) = "ANULADA" Then
    
        mblnAlterandoHistorico = True
        
        'ORC677
        If Year(CDate(txt_DataAnulucao)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data da anulação tem que estar no exercício de " & gintExercicio & "."
            If txt_DataAnulucao.Enabled Then txt_DataAnulucao.SetFocus
            Exit Function
        End If
        
    Else
        mblnAlterandoHistorico = False
    
        If tab_3DAnulacao.TabEnabled(1) = True Then
        
            'M6R Pergunta se exitem itens selecionados somente quando for parcela 0
            If lvw_ItensAnulacao.ListItems.Count = 0 And lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).Text = 0 Then
                ExibeMensagem "É necessário selecionar algum item para anular ."
                tab_3DAnulacao.Tab = 1
                Exit Function
            End If
            
            If lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).Text = 0 Then
                For i = 1 To lvw_ItensAnulacao.ListItems.Count
                    valorTotalAnulado = valorTotalAnulado + (CDbl(lvw_ItensAnulacao.ListItems(i).SubItems(4)) _
                                        * CDbl(lvw_ItensAnulacao.ListItems(i).SubItems(5)))
                Next
                If Val(gstrConvVrParaSql(txt_ValorAnulacao)) = 0 Then txt_ValorAnulacao = gstrConvVrDoSql(valorTotalAnulado)
            Else
                valorTotalAnulado = Val(gstrConvVrParaSql(txt_ValorAnulacao))
            End If
        End If
        'REMOVIDO 26/04/04 POR M4RC3LØ
        'M6R PERGUNTAR PARA O FLAVIO
'        If lvw_Anulacao.SelectedItem.Text <> "0" And VerificaNotasFiscais(lvw_Anulacao.SelectedItem.Tag) = True Then
'           ExibeMensagem "Não é possivel Liquidar parcelas que possuem Notas Fiscais."
'           Exit Function
'        End If
        
        'M4RC3LØ VERIFICA SE O EVENTO CONTABIL PERTENCE AO EXERCICIO
        strSQL = ""
        If (cbo_intEventoAnul.ListIndex >= 0) And (mblnRestosAPagar = True) Then
            strSQL = "SELECT intExercicio FROM " & gstrEvento & " WHERE Pkid = " & cbo_intEventoAnul.ItemData(cbo_intEventoAnul.ListIndex)
            
            Set gobjBanco = New clsBanco
            
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.RecordCount > 0 Then
                   If adoResultado!intExercicio <> "" Then
                    If Trim(Val(txtintExercicioEmpenho.Text)) <> Val(adoResultado!intExercicio) Then
                        ExibeMensagem "Evento Contabil Incorreto !!!"
                        Exit Function
                    End If
                   Else
                        ExibeMensagem "Evento Contabil Incorreto !!!"
                        Exit Function
                   End If
                End If
            End If
        End If
        
        If Not IsDate(txt_DataAnulucao) Then
            ExibeMensagem "É necessário informar a data de cancelamento."
            If txt_DataAnulucao.Enabled Then txt_DataAnulucao.SetFocus
            Exit Function
        End If
        
        If lvw_Anulacao.SelectedItem Is Nothing Then
           ExibeMensagem "É nescessário selecionar uma parcela na lista abaixo."
           Exit Function
        ElseIf Val(lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index)) = 0 And (cbo_intEventoAnul.ListIndex = -1 And mblnRestosAPagar) Then
            ExibeMensagem "O Evento Contabil deve ser selecionado."
            If cbo_intEventoAnul.Enabled Then cbo_intEventoAnul.SetFocus
            Exit Function
        ElseIf CVDate(txt_DataAnulucao) <= VerificaDataEncerramento("EC", gintExercicio) Then
            ExibeMensagem "Data da anulação tem que ser superior ao fechamento."
            If txt_DataAnulucao.Enabled Then txt_DataAnulucao.SetFocus
            Exit Function
        ElseIf gblnDataValida(txt_DataAnulucao) = False Then
            ExibeMensagem "Data da anulação tem que ser informada corretamente."
            If txt_DataAnulucao.Enabled Then txt_DataAnulucao.SetFocus
            Exit Function
        ElseIf CVDate(txt_DataAnulucao) < CVDate(txtDTMDATA) Then
            ExibeMensagem "Data da anulação não ser inferior a data do empenho."
            If txt_DataAnulucao.Enabled Then txt_DataAnulucao.SetFocus
            Exit Function
        ElseIf Val(gstrConvVrParaSql(gstrConvVrDoSql(valorTotalAnulado, 2))) <> Val(gstrConvVrParaSql(txt_ValorAnulacao)) And tab_3DAnulacao.TabEnabled(1) = True Then
            ExibeMensagem "O Valor da somatória dos itens anulados não pode ser diferente do total a ser anulado." _
            & vbNewLine & "Somátoria do Itens Anulados = " & gstrConvVrDoSql(valorTotalAnulado)
            If txt_ValorAnulacao.Enabled Then txt_ValorAnulacao.SetFocus
            Exit Function
        ElseIf CDbl(txt_ValorAnulacao) > CDbl(lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).SubItems(2)) Then
            ExibeMensagem "O Valor a ser Anulado não pode ser maior que o saldo da parcela."
            Exit Function
        ElseIf UCase(lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).SubItems(3)) = "CANC. RP" Then
            ExibeMensagem "Parcela já anulada."
            Exit Function
    'Rotina dasabilitada por motivo de cancelamento de RP
    '    ElseIf mblnRestosAPagar And lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index) = 0 Then
    '        ExibeMensagem "Esta parcela não pode ser anulada."
    '        Exit Function
        ElseIf mblnRestosAPagar And UCase(lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).SubItems(3)) = "LIQUIDADA" Then
           If Val(Right(lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).SubItems(1), 4)) < gintExercicio Then
              ExibeMensagem "Não é possível anular uma liquidacao feita no exercicio anterior."
              Exit Function
           End If
        End If
        
        'ORC677
        If IsDate(txt_DataAnulucao.Text) Then
            If Year(CDate(txt_DataAnulucao)) <> CInt(gintExercicio) Then
                ExibeMensagem "A data da anulação tem que estar no exercício de " & gintExercicio & "."
                If txt_DataAnulucao.Enabled Then txt_DataAnulucao.SetFocus
                Exit Function
            End If
        End If
        
        dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
            
        If dtmDtEncerramento = Empty Then
           Exit Function
        Else
           If CDate(txt_DataAnulucao) <= dtmDtEncerramento Then
              ExibeMensagem "A data da anulação deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
              If txt_DataAnulucao.Enabled Then txt_DataAnulucao.SetFocus
              Exit Function
           End If
        End If
        
        
        If frmSubElementoEstorno.Visible = True Then
            For i = 1 To lvwSubElementoEst.ListItems.Count
                dblvalorAcumulado = dblvalorAcumulado + Val(gstrConvVrParaSql(lvwSubElementoEst.ListItems(i).SubItems(3)))
            Next
             
            If Val(gstrConvVrParaSql(txt_ValorAnulacao)) <> dblvalorAcumulado Then
                ExibeMensagem "O valor acumulado dos Sub-Elementos deve ser o mesmo da Anulação."
                If dbcItemDespCodSubElementoEst.Enabled Then dbcItemDespCodSubElementoEst.SetFocus
                Exit Function
            End If
            
            If GeraArraysSubElmento(gstrItemData(cbo_intEvento), True, lvwSubElementoEst) = False Then
                Exit Function
            End If
            
        End If
        
        
    End If
    blnMantemItemSelAnul = False 'não é necessário manter pois vai gravar e atualizar o grid sem o risco de perder informaçaoes
    blnDadosAnulacaoOk = True
    
End Function

Private Function CalculaSaldoDevolucaoParcelaAdiantamento(ByVal strpkidParcela) As Double
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = "SELECT dblvalor "
    strSQL = strSQL & " - (Select " & gstrISNULL("SUM(dblvalor)", "0") & "  FROM " & gstrControleAdiantamento & " WHERE intParcela = " & strpkidParcela & ")"
    strSQL = strSQL & " Saldo"
    strSQL = strSQL & " FROM " & gstrSubempenho
    strSQL = strSQL & " WHERE PKID = " & strpkidParcela
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        CalculaSaldoDevolucaoParcelaAdiantamento = Val(gstrConvVrParaSql(gstrENulo(adoResultado!Saldo)))
    End If
    
End Function

Sub GravaAnulacao()
    Dim strSQL                  As String
    Dim dblValorParcAExcluir    As Double
    Dim dblValorParcASomar      As Double
    Dim strEmpenhoAnulacao      As String
    Dim blnAnula                As Boolean
    Dim i                       As Integer
    Dim strAux                  As String
    Dim adoResultado            As ADODB.Recordset
    Dim adoResultadoAux         As ADODB.Recordset
    
        If blnDadosAnulacaoOk Then
       'VERIFICANDO SE E ALTERAÇÃO DE HISTÓRICO
       If mblnAlterandoHistorico Then
           If gblnExclusaoGravacaoOk("", "Deseja alterar o histórico") = True Then
                strSQL = ""
                strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
                strSQL = strSQL & "strHistorico = "
                strSQL = strSQL & "'" & Trim(txt_HistoricoAnulacao) & "' "
                strSQL = strSQL & "WHERE PKId = " & lvw_Anulacao.SelectedItem.Tag
                Set gobjBanco = New clsBanco
                gobjBanco.Execute (strSQL)
                LeSubEmpenho lvw_Anulacao, 4, txtPKId
                LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                LeSubEmpenho lvw_Liquidacao, 2, txtPKId
           End If
       Else
       
       
       
       
       
       'GRAVAÇAO DE ANULAÇÃO
       If blnEmpenhadoCompras Then
            With lvw_Anulacao
                If .ListItems(.SelectedItem.Index).Text = 0 Then
                    If Val(gstrConvVrParaSql(txt_ValorAnulacao)) > Val(gstrConvVrParaSql(.ListItems(.SelectedItem.Index).SubItems(2))) Then
                        ExibeMensagem "O valor da anulação não pode ser maior que o valor da parcela"
                        Exit Sub
                    End If
                                       
                End If
            End With
       End If
              
       Set gobjBanco = New clsBanco
       With lvw_Anulacao
          dblValorParcAExcluir = Val(gstrConvVrParaSql(.ListItems(.SelectedItem.Index).SubItems(2)))
          If UCase(.ListItems(.SelectedItem.Index).SubItems(4)) <> "COMPLEMENTO" Then
                  dblValorParcASomar = dblValorParcAExcluir
          Else
                  dblValorParcASomar = 0
          End If
          
          If dblValorParcASomar = 0 And UCase(.ListItems(.SelectedItem.Index).SubItems(4)) <> "COMPLEMENTO" Then
             Beep
             ExibeMensagem "Esta parcela não possui saldo a ser anulado."
             Exit Sub
          End If
           
           'compara a anulacao com o total a devolver
           If Val(gstrConvVrParaSql(txt_ValorAnulacao)) > CalculaSaldoDevolucaoParcelaAdiantamento(.SelectedItem.Tag) Then
                ExibeMensagem "O valor a ser anulado não pode ser maior que o saldo para Devolução (" & gstrConvVrDoSql(CalculaSaldoDevolucaoParcelaAdiantamento(.SelectedItem.Tag)) & ")."
                Exit Sub
           End If
           
           
           If gblnExclusaoGravacaoOk("", "Confirma anulação da parcela?", True) Then
              If lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).Text = 0 Or (.ListItems(.SelectedItem.Index).ListSubItems(3).Text = "Paga") Then
                 
                 strSQL = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
                 
                 If GerarEmpenhoDeEstorno And Not mblnRestosAPagar Then
                 
                    strEmpenhoAnulacao = GeraProximoDeEmpenho
ProximoCodigo:
                    If gblnExisteCodigo(2, gstrEmpenho, "intNumero", strEmpenhoAnulacao, gstrDATEPART(strYEAR, "dtmData"), "'" & Val(gintExercicio) & "'") Then
                          strEmpenhoAnulacao = GeraProximoDeEmpenho
                          GoTo ProximoCodigo
                    Else
                       If gblnExisteCodigo(2, gstrSubempenho, "intEmpenhoAnulacao", strEmpenhoAnulacao, gstrDATEPART(strYEAR, "dtmData"), "'" & Val(gintExercicio) & "'") Then
                          strEmpenhoAnulacao = GeraProximoDeEmpenho
                          GoTo ProximoCodigo
                       End If
                    End If
                    strSQL = strSQL & "INSERT INTO " & gstrSubempenho & " ("
                    strSQL = strSQL & "intEmpenho, intNumero, dtmData, dtmAnulacao,"
                    strSQL = strSQL & "dblValor, bytSituacao, bytTipo, strHistorico, "
                    strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr,intEmpenhoAnulacao) "
                    strSQL = strSQL & "SELECT " & txtPKId & ", "
                    strSQL = strSQL & "MAX(intNumero) - MAX(intNumero) , "
                    strSQL = strSQL & gstrConvDtParaSql(txt_DataAnulucao) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(txt_DataAnulucao) & ", "
                    strSQL = strSQL & gstrConvVrParaSql(txt_ValorAnulacao) & ", 4, "
                    'bytTipo para anulação de parcelas pagas M6R
                    strSQL = strSQL & IIf(lvw_Anulacao.SelectedItem.SubItems(3) = "Paga", 3, 1) & ", "
                    strSQL = strSQL & "'" & txt_HistoricoAnulacao & "', "
                    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSQL = strSQL & glngCodUsr & ", " & strEmpenhoAnulacao & " FROM " & gstrSubempenho & " "
                    strSQL = strSQL & "WHERE intEmpenho = " & txtPKId & ";"
                 
                 Else
                    
                    strSQL = strSQL & "INSERT INTO " & gstrSubempenho & " ("
                    strSQL = strSQL & "intEmpenho, intNumero, dtmData, dtmAnulacao,"
                    strSQL = strSQL & "dblValor, bytSituacao, bytTipo, strHistorico, "
                    strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr) "
                    strSQL = strSQL & "SELECT " & txtPKId & ", "
                    strSQL = strSQL & "MAX(intNumero) - MAX(intNumero) , "
                    strSQL = strSQL & gstrConvDtParaSql(txt_DataAnulucao) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(txt_DataAnulucao) & ", "
                    strSQL = strSQL & gstrConvVrParaSql(txt_ValorAnulacao) & ", 4, "
                    'bytTipo para anulação de parcelas pagas M6R
                    strSQL = strSQL & IIf(lvw_Anulacao.SelectedItem.SubItems(3) = "Paga", 3, 1) & ", "
                    strSQL = strSQL & "'" & txt_HistoricoAnulacao & "', "
                    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                    strSQL = strSQL & glngCodUsr & " FROM " & gstrSubempenho & " "
                    strSQL = strSQL & "WHERE intEmpenho = " & txtPKId & ";"
                 
                 End If
                 
                 
                 
                 strSQL = strSQL & gstrGravaItenAnulado
                 CalculaSaldoAtual
                 
                 strSQL = strSQL & " UPDATE " & gstrSubempenho
                 strSQL = strSQL & " SET dblEmpenhadoAteData = " & gstrConvVrParaSql(Val(gstrConvVrParaSql(txt_TotalDotado)) - Val(gstrConvVrParaSql(txt_ValorAnulacao)))
                 strSQL = strSQL & ", dblSaldoAtual = " & gstrConvVrParaSql(Val(gstrConvVrParaSql(txt_SaldoDotacao)) + Val(gstrConvVrParaSql(txt_ValorAnulacao)))
                 strSQL = strSQL & " WHERE intNumero = 0 AND "
                 strSQL = strSQL & " PKID = (SELECT MAX(PKID) FROM " & gstrSubempenho & " );"
                 
                 'Grava Nota fiscal negativa M7R
                     
                 If lvw_Anulacao.SelectedItem.SubItems(3) = "Paga" Then
                     strSQL = strSQL & " INSERT INTO " & gstrSubEmpenhoNF & " ("
                     strSQL = strSQL & "intSubEmpenho, dtmData, dblValorNF, "
                     strSQL = strSQL & "strNotaFiscal, dtmDtAtualizacao, lngCodUsr) VALUES "
                     strSQL = strSQL & "( " & lvw_Anulacao.SelectedItem.Tag & " , "
                     strSQL = strSQL & gstrConvDtParaSql(txt_DataAnulucao) & ", "
                     strSQL = strSQL & gstrConvVrParaSql(txt_ValorAnulacao * -1) & ", "
                     strSQL = strSQL & "'Nota de Anulação de Despesa', "
                     strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & " "
                     strSQL = strSQL & "); "
                  
                     'M8R PEN_701
                     
                     strAux = "SELECT intEvento FROM " & gstrSubempenho & " WHERE Pkid  = " & lvw_Anulacao.SelectedItem.Tag
                     
                     Set gobjBanco = New clsBanco
                     
                     gobjBanco.CriaADO strAux, 5, adoResultado
                     
                     If frmSubElementoEstorno.Visible = False Then
                        If Not GeraMovimentosByEvento(adoResultado!intEvento, txt_DataAnulucao, Str(CDbl((txt_ValorAnulacao) * (-1))), "Nota Fiscal de Anulação", txtintNumero, "3", , , , , , False) Then
                           ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                        End If
                     Else
                        If Not GeraMovimentosByEvento(adoResultado!intEvento, txt_DataAnulucao, Str(CDbl((txt_ValorAnulacao) * (-1))), "Nota Fiscal de Anulação", txtintNumero, "3", aryContas, aryTpMov, , aryValor, , False) Then
                           ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                        End If
                     End If
                     
                     If Not blnGravaMovLiq(Val(txtPKId), _
                                           lvw_Anulacao.SelectedItem.Tag, _
                                           gstrItemData(cboProgramaTrabalho, True), _
                                           txt_DataAnulucao, txt_ValorAnulacao * (-1), _
                                           "***Cancelamento de Liquidação***") Then
                                             '* (-1)
                                             'txt_DataAnulucao, "-" & txt_ValorAnulacao,
                            ExibeMensagem "Não foi possível completar o cancelamento desta Liquidação para este empenho"
                     End If
                     
                  
                  End If
                                  
                 strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
                     
                 
                 If gobjBanco.Execute(strSQL) Then
                    strSQL = ""
                    strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
                    strSQL = strSQL & "dblValor = dblValor - " & gstrConvVrParaSql(txt_ValorAnulacao) & ","
                    strSQL = strSQL & "dtmDtAtualizacao = "
                    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
                    strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
                    strSQL = strSQL & "WHERE PKId = " & lvw_Anulacao.SelectedItem.Tag
                    If gobjBanco.Execute(strSQL) Then
                       'Pen_780 17/07/2004
                       If cbointReservaDotacao.ListIndex > -1 Then
                           'FAZ O VALOR VOLTAR PARA A RESERVA
                           GeraAnulacaoReservaDotacaoLiberada
                           'FAZ O VALOR VOLTAR DIRETO PARA A dotacao
                           CancelarReservaDotacao gstrItemData(cbointReservaDotacao), txt_DataAnulucao, txt_ValorAnulacao, "Cancelamento automático Reserva /Anulação de Empenho Nº: " & Trim(txtintNumero) & "/" & Trim(txtintExercicioEmpenho), True
                       End If
                       PreencheDadosReserva
                    End If
                 End If
                 
                 If Not mblnRestosAPagar Then
                    If frmSubElementoEstorno.Visible = False Then
                        If Not GeraMovimentosByEvento(gstrItemData(cbo_intEvento), txt_DataAnulucao, Str(CDbl((txt_ValorAnulacao) * (-1))), txt_HistoricoAnulacao, txtintNumero, "3", , , , , , False) Then
                            ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                        End If
                    Else
                        If Not GeraMovimentosByEvento(gstrItemData(cbo_intEvento), txt_DataAnulucao, Str(CDbl((txt_ValorAnulacao) * (-1))), txt_HistoricoAnulacao, txtintNumero, "3", aryContas, aryTpMov, , aryValor, , False) Then
                            ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                        End If
                    End If
                 Else
                    'STRORIGEM 12 PARA CANCELAMENTO DE RESTOS A PAGAR
                    If Not GeraMovimentosByEvento(gstrItemData(cbo_intEventoAnul), txt_DataAnulucao, Str(CDbl(txt_ValorAnulacao)), txt_HistoricoAnulacao, txtintNumero, "12", , , , , , True) Then
                        ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                    Else
                        strSQL = ""
                        strSQL = strSQL & " UPDATE " & gstrSubempenho & " SET "
                        strSQL = strSQL & " intEvento = " & gstrItemData(cbo_intEventoAnul)
                        strSQL = strSQL & " WHERE PKId = (SELECT MAX(PKID) FROM " & gstrSubempenho & " )"
                        
                        gobjBanco.Execute strSQL
                    End If
                 End If
                 If frmSubElementoEstorno.Visible = True Then GravaSubElementos , glngRetornaPkidTabelaPai("seq" & gstrSubempenho, gstrSubempenho), lvwSubElementoEst
                 
                 LeSubEmpenho lvw_Anulacao, 4, txtPKId
                 'Refresh na list view da liquidação para não gerar erro na gravação de uma parcela
                 'de liquidação após uma de anulação.
                 LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                 LeSubEmpenho lvw_Liquidacao, 2, txtPKId
                 PreencheSaldoEmpenho
                 mblnAtualizaTelaSubempenho = True
              Else
                 
                 strSQL = ""
                 
                 If tab_3DAnulacao.TabEnabled(1) And bytDBType = EDatabases.Oracle Then
                    strSQL = "BEGIN "
                 End If
                 
                 
                 strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
                 strSQL = strSQL & "dtmAnulacao = "
                 strSQL = strSQL & gstrConvDtParaSql(txt_DataAnulucao) & ", "
                 strSQL = strSQL & "bytSituacao = 4, " '4=Anulada
                 strSQL = strSQL & "strHistorico = "
                 strSQL = strSQL & "'" & Trim(txt_HistoricoAnulacao) & "', "
                 strSQL = strSQL & "dtmDtAtualizacao = "
                 strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
                 strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
                 strSQL = strSQL & "WHERE PKId = " & lvw_Anulacao.SelectedItem.Tag
                 
                 If tab_3DAnulacao.TabEnabled(1) And bytDBType = EDatabases.Oracle Then
                    strSQL = strSQL & ";"
                    strSQL = strSQL & gstrGravaItenAnulado(lvw_Anulacao.SelectedItem.Tag)
                    strSQL = strSQL & " END; "
                 End If
                 
                 
                 
                 If gobjBanco.Execute(strSQL) Then
                    strSQL = "UPDATE " & gstrSubempenho & " "
                    strSQL = strSQL & "SET dblValor = dblValor +  "
                    strSQL = strSQL & gstrConvVrParaSql(dblValorParcASomar) & " "
                    strSQL = strSQL & "WHERE PKId = " & .ListItems(1).Tag
                    If gobjBanco.Execute(strSQL) Then
                       
                       If UCase(.ListItems(.SelectedItem.Index).SubItems(4)) = "COMPLEMENTO" Then
                          If frmSubElementoEstorno.Visible = False Then
                                If Not GeraMovimentosByEvento(gstrItemData(cbo_intEvento), txt_DataAnulucao, Str(CDbl((txt_ValorAnulacao) * (-1))), txt_HistoricoAnulacao, txtintNumero, "3") Then
                                   ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                                End If
                          Else
                                If Not GeraMovimentosByEvento(gstrItemData(cbo_intEvento), txt_DataAnulucao, Str(CDbl((txt_ValorAnulacao) * (-1))), txt_HistoricoAnulacao, txtintNumero, "3", aryContas, aryTpMov, , aryValor) Then
                                   ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                                End If
                          End If
                          
                       End If
                       LeSubEmpenho lvw_Liquidacao, 2, txtPKId
                       LeSubEmpenho lvw_Anulacao, 4, txtPKId
                       LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                       PreencheSaldoEmpenho
                       mblnAtualizaTelaSubempenho = True
                    End If
                 End If
              End If
              
            If blnEmpenhadoCompras Then
                blnAnula = True
                If lvw_Anulacao.ListItems(1).SubItems(2) <> "0,00" Then blnAnula = False
                For i = 2 To lvw_Anulacao.ListItems.Count
                    If UCase(lvw_Anulacao.ListItems(i).SubItems(3)) <> "ANULADA" And UCase(lvw_Anulacao.ListItems(i).SubItems(3)) <> "CANC. RP" Then blnAnula = False
                Next
                If blnAnula Then
                    If SalvaAnulacaoCompras = False Then
                        ExibeMensagem "Ocorreu um erro ao gravar a Anulação de Empenho na Requisição de Compras."
                    End If
                    
                    blnAnula = False
                End If
            End If
            lvw_Anulacao.ListItems(1).Selected = True
            lvw_Anulacao_ItemClick lvw_Anulacao.ListItems(1)
           Else
                intItemIndexAnul = lvw_Anulacao.SelectedItem.Index
                blnMantemItemSelAnul = True 'variavel usada para manter o grid selecionado neste item, pois o mesmo é atualizado e o item selecionado muda
           End If
           

       End With
       End If
    End If
End Sub

Private Function blnDadoLiquidacaoOK(Optional ByVal blnGravouExtras As Boolean, _
                                     Optional ByVal blnGravouNotas As Boolean, _
                                     Optional ByVal blnSemMsg As Boolean, _
                                     Optional ByVal blnGravouRetOrcamentaria As Boolean) As Boolean
    Dim dblSoma           As Double
    Dim intCont           As Integer
    Dim intParcela        As Integer
    Dim dtmDtEncerramento As Date
    Dim msgLiquidacao     As String
    Dim mstrMsgErro       As String
    Dim adoResultado      As ADODB.Recordset
    Dim strSQL            As String
    
    If gblnDataValida(txt_DataLiuidacao) = False Then
        
        If Not blnSemMsg Then
            ExibeMensagem "Data da liquidação incorreta."
            Exit Function
        End If
        
        msgLiquidacao = "Data da liquidação é incorreta."
        GoTo saida
        
    ElseIf CVDate(txt_DataLiuidacao) < CVDate(txtDTMDATA) Then
        
        If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
        If Not blnSemMsg Then
            ExibeMensagem "Data da liquidação não pode ser menor que a data do empenho."
            Exit Function
        End If
        
        msgLiquidacao = "Data da liquidação não pode ser menor que a data do empenho."
        GoTo saida
    End If
    'M4RC3LØ VERIFICA SE O EVENTO CONTABIL PERTENCE AO EXERCICIO
    If cbo_intEventoLiq.ListIndex >= 0 Then
        strSQL = ""
        strSQL = "SELECT intExercicio FROM " & gstrEvento & " WHERE Pkid = " & cbo_intEventoLiq.ItemData(cbo_intEventoLiq.ListIndex)
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If adoResultado.RecordCount > 0 Then
                If Trim(Val(txtintExercicioEmpenho.Text)) <> Val(adoResultado!intExercicio) Then
                    If Not blnSemMsg Then
                        ExibeMensagem "Evento Contabil incorreto !!!"
                        Exit Function
                    End If
                    
                    msgLiquidacao = "o Evento Contabil esta incorreto."
                    GoTo saida
                End If
            End If
        End If
    End If
    
    If cbo_intEventoLiq.ListIndex = -1 Then
        If cbo_intEventoLiq.Enabled Then cbo_intEventoLiq.SetFocus
        
        If Not blnSemMsg Then
            ExibeMensagem "A Evento Contabil tem que ser informado."
            Exit Function
        End If
                
        msgLiquidacao = "A Evento Contabil tem que ser informado."
        GoTo saida
    End If
    
    For intCont = 1 To lvw_Liquidacao.ListItems.Count
       If lvw_Liquidacao.ListItems(intCont).Text <> "0" Then
          intParcela = 1
       End If
    Next
    
    If intParcela = 1 And lvw_Liquidacao.ListItems(lvw_Liquidacao.SelectedItem.Index).Text = "0" Then
        If Not blnSemMsg Then
            ExibeMensagem "Este empenho possui mais de uma parcela e não poderá ser liquidado."
            Exit Function
        End If
       
        msgLiquidacao = "Este empenho possui mais de uma parcela."
        GoTo saida
    End If
    
    If Val(gstrConvVrParaSql(lvw_Liquidacao.SelectedItem.ListSubItems(3))) = 0 Then
        If Not blnSemMsg Then
            ExibeMensagem "Não é possível liquidar uma Parcela com valor zero."
            Exit Function
        End If
        
        msgLiquidacao = "Não é possível liquidar uma Parcela com valor zero."
        GoTo saida
    End If
    
    
    
    If Val(gstrConvVrParaSql(lbl_ValorTotal)) <> Val(gstrConvVrParaSql(txt_dblValorAux)) Then
        If Not blnSemMsg Then
            ExibeMensagem "O valor total da(s) Nota(s) Fiscal(is) deve ser igual ao valor liquidado."
            Exit Function
        End If
        
        msgLiquidacao = "O valor total da(s) Nota(s) Fiscal(is) deve ser igual ao valor liquidado."
        GoTo saida
    End If
        
    dtmDtEncerramento = VerificaDataEncerramento("EF", gintExercicio)
    
    If dtmDtEncerramento = Empty Then
       blnDadoLiquidacaoOK = True
       Exit Function
    Else
       If CDate(txt_DataLiuidacao) <= dtmDtEncerramento Then
            If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
            If Not blnSemMsg Then
                ExibeMensagem "A data da liquidação deve ser maior que a data de último encerramento financeiro (" & dtmDtEncerramento & ")."
                Exit Function
            End If
            
            msgLiquidacao = "A data da liquidação deve ser maior que a data de último encerramento financeiro (" & dtmDtEncerramento & ")."
            GoTo saida
       End If
    End If
    
    
    'ORC677
    If IsDate(txt_DataLiuidacao.Text) Then
        If Year(CDate(txt_DataLiuidacao)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data de liquidação tem que estar no exercício de " & gintExercicio & "."
            If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
            Exit Function
        End If
    
        If CDate(txt_DataLiuidacao) < CDate(lvw_Liquidacao.SelectedItem.SubItems(1)) Then
            ExibeMensagem "A data de liquidação não pode ser menor que a data de criação da parcela."
            If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
            Exit Function
        End If
        
       'A mensagem de critica esta dentro da rotina que é usada tambem na rotina que cria parcela diretamente da guia liquidação
        If gblnMaiorSubEmpenhoLiq(txt_DataLiuidacao) = False Then
               Exit Function
        End If
    End If
    
    
    
    If gblnDataValida(txt_DataVencto) = False Then
        ExibeMensagem "A Data do Vencimento é inválida."
        If txt_DataVencto.Enabled Then txt_DataVencto.SetFocus
        Exit Function
    ElseIf CVDate(txt_DataLiuidacao) > CVDate(txt_DataVencto) Then
        ExibeMensagem "A Data do Vencimento não poder ser inferior a data da Liquidação."
        If txt_DataVencto.Enabled Then txt_DataVencto.SetFocus
        Exit Function
    End If
    
    
    blnDadoLiquidacaoOK = True
    
    Exit Function
saida:


    If Trim(msgNotas) <> "" Then mstrMsgErro = msgNotas
    
    If Not blnGravouExtras Then
        If Trim(msgExtras) <> "" Then mstrMsgErro = mstrMsgErro & vbNewLine & msgExtras
    End If
    
    If Not blnGravouRetOrcamentaria Then
        If Trim(msgOrcamentario) <> "" Then mstrMsgErro = mstrMsgErro & vbNewLine & msgOrcamentario
    End If

    msgLiquidacao = "A parcela não pode ser Liquidada porque " & msgLiquidacao
    mstrMsgErro = mstrMsgErro & vbNewLine & msgLiquidacao
    ExibeMensagem mstrMsgErro
    
    
    msgOrcamentario = ""
    msgExtras = ""
End Function

Private Function blnGravouLiquidacao(strConta As String, _
                                     lngChave As Long, _
                                     txtdblValor As String, _
                                     bytTipo As Byte, _
                                     blnAlterando As Boolean) As Boolean
    Dim strSQL  As String
    strSQL = ""
    
'      If blnLiquidouParcela Then
          If blnAlterando Then
              strSQL = strSQL & "UPDATE " & gstrSubempenhoLiquidado & " SET "
              strSQL = strSQL & "intConta = " & Val(strConta) & ", "
              strSQL = strSQL & "dblValor = " & gstrConvVrParaSql(txtdblValor) & ", "
              strSQL = strSQL & "dtmDtAtualizacao = "
              strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
              strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
              strSQL = strSQL & "WHERE PKId = " & lngChave
          Else
              strSQL = strSQL & "INSERT INTO " & gstrSubempenhoLiquidado & " ("
              strSQL = strSQL & "intParcela, intConta, dblValor, bytTipo, "
              strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr "
              strSQL = strSQL & ") VALUES ("
              strSQL = strSQL & lvw_Liquidacao.SelectedItem.Tag & ", "
              strSQL = strSQL & Val(strConta) & ", " & gstrConvVrParaSql(txtdblValor) & ", "
              strSQL = strSQL & bytTipo & ", "
              strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
              strSQL = strSQL & glngCodUsr & ")"
          End If
          Set gobjBanco = New clsBanco
          If gobjBanco.Execute(strSQL) Then
              blnGravouLiquidacao = True
          End If
 '     End If
    
End Function

Private Function blnGravouLiquidacaoOrcamentario(strConta As String, _
                                     lngChave As Long, _
                                     txtdblValor As String, _
                                     bytTipo As Byte, _
                                     blnAlterando As Boolean) As Boolean
    Dim strSQL  As String
    strSQL = ""
    
'      If blnLiquidouParcela Then
          If blnAlterando Then
              strSQL = strSQL & "UPDATE " & gstrSubEmpRetencaoOrcamentaria & " SET "
              strSQL = strSQL & "INTCODIGOORCAMENTARIO = " & Val(strConta) & ", "
              strSQL = strSQL & "dblValor = " & gstrConvVrParaSql(txtdblValor) & ", "
              strSQL = strSQL & "dtmDtAtualizacao = "
              strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
              strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
              strSQL = strSQL & "WHERE PKId = " & lngChave
          Else
              strSQL = strSQL & "INSERT INTO " & gstrSubEmpRetencaoOrcamentaria & " ("
              strSQL = strSQL & "intParcela, INTCODIGOORCAMENTARIO, dblValor, "
              strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr "
              strSQL = strSQL & ") VALUES ("
              strSQL = strSQL & lvw_Liquidacao.SelectedItem.Tag & ", "
              strSQL = strSQL & Val(strConta) & ", " & gstrConvVrParaSql(txtdblValor) & ", "
              strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
              strSQL = strSQL & glngCodUsr & ")"
          End If
          Set gobjBanco = New clsBanco
          If gobjBanco.Execute(strSQL) Then
              blnGravouLiquidacaoOrcamentario = True
          End If
 '     End If
    
End Function


Private Function blnLiquidouParcela() As Boolean
    Dim strSQL      As String
    Dim lngParcela  As Long
    If Val(lvw_Liquidacao.ListItems(lvw_Liquidacao.SelectedItem.Index).SubItems(8)) = 2 Then 'Liquidada
        blnLiquidouParcela = True
    Else
        lngParcela = lvw_Liquidacao.SelectedItem.Tag
        strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
        strSQL = strSQL & "dblDesconto = "
        strSQL = strSQL & gstrConvVrParaSql(txt_dblDesconto) & ", "
        strSQL = strSQL & "dtmLiquidacao = "
        strSQL = strSQL & gstrConvDtParaSql(txt_DataLiuidacao) & ", "
        strSQL = strSQL & "bytSituacao = 2, " '2= LIQUIDADA
        'strSql = strSql & "strHistorico = "
        'strSql = strSql & "'" & Trim(txt_HistoricoLiquidacao) & "', "
        strSQL = strSQL & "dtmVencimento = "
        strSQL = strSQL & gstrConvDtParaSql(txt_DataVencto) & ", "
        strSQL = strSQL & "dtmDtAtualizacao = "
        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
        strSQL = strSQL & "intevento = " & gstrItemData(cbo_intEventoLiq) & ", "
        strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
        strSQL = strSQL & "WHERE PKId = " & lvw_Liquidacao.SelectedItem.Tag
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSQL) Then
            
            If Not GeraMovimentosByEvento(gstrItemData(cbo_intEventoLiq), txt_DataLiuidacao, Str(CDbl(txt_dblValorAux)), txt_HistoricoLiquidacao, txtintNumero, "3", , , , , True) Then
                  ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
            End If
            
            If Not blnGravaMovLiq(Val(txtPKId), lvw_Liquidacao.SelectedItem.Tag, gstrItemData(cboProgramaTrabalho, True), txt_DataLiuidacao, txt_dblValorAux, Trim(txt_HistoricoLiquidacao)) Then
               ExibeMensagem "Não foi possível completar a Liquidação deste empenho."
            End If
               
            LeSubEmpenho lvw_Liquidacao, 2, txtPKId
            LeSubEmpenho lvw_Anulacao, 4, txtPKId
            LeSubEmpenho lvw_ListaSubempenho, , txtPKId
            Call gblnEncontroItemNoListView(lvw_Liquidacao, CStr(lngParcela), lvwTag)
            mblnAtualizaTelaSubempenho = True
            blnLiquidouParcela = True
        Else
            ExibeMensagem "Ocorreu erro ao atualizar a parcela como Liquidada." & vbNewLine & "A gravação foi abortada e nenhum registro foi gravado."
        End If
    End If
End Function

Private Function blnDadosExtraOk() As Boolean

    If Val(gstrConvVrParaSql(txt_ValorExtra)) > Val(gstrConvVrParaSql(lblLiquido)) Then ' - Val(gstrConvVrParaSql(lblExtra))) Then
        ExibeMensagem "O valor extra não pode ser maior que o valor da parcela."
        Exit Function
    End If

    If Val(gstrConvVrParaSql(txt_ValorExtra)) = 0 Then
        ExibeMensagem "O valor tem ser informado corretamente."
        If txt_ValorExtra.Enabled Then txt_ValorExtra.SetFocus
        Exit Function
    ElseIf Not (Val(gstrConvVrParaSql(lblLiquido)) > 0) Then
        If lvw_Extra.ListItems.Count > 0 Then
            ExibeMensagem "O valor da parcela tem que ser superior ao total dos lançamentos."
        Else
            ExibeMensagem "O valor da parcela tem que ser superior ao valor do lançamento."
        End If
        If txt_ValorExtra.Enabled Then txt_ValorExtra.SetFocus
        Exit Function
    ElseIf cbo_ContaExtra.ListIndex = -1 Then
        ExibeMensagem "A conta tem ser informada corretamente."
        If cbo_ContaExtra.Enabled Then cbo_ContaExtra.SetFocus
        Exit Function
    End If
    blnDadosExtraOk = True
End Function

Private Function blnDadosOrcamentarioOk() As Boolean

    If Val(gstrConvVrParaSql(txt_ValorOrcamentario)) > Val(gstrConvVrParaSql(lblLiquido)) Then ' - Val(gstrConvVrParaSql(txt_dblDesconto))) Then
        ExibeMensagem "O valor do desconto orçamentario não pode ser maior que o valor da parcela."
        Exit Function
    End If

    If Val(gstrConvVrParaSql(txt_ValorOrcamentario)) = 0 Then
        ExibeMensagem "O valor tem ser informado corretamente."
        If txt_ValorOrcamentario.Enabled Then txt_ValorOrcamentario.SetFocus
        Exit Function
'    ElseIf Not (Val(gstrConvVrParaSql(txt_dblDesconto)) > 0) Then
'        If lvw_Orcamentario.ListItems.Count > 0 Then
'            ExibeMensagem "O valor da parcela tem que ser superior ao total dos lançamentos."
'        Else
'            ExibeMensagem "O valor da parcela tem que ser superior ao valor do lançamento."
'        End If
'        If txt_ValorOrcamentario.Enabled Then txt_ValorOrcamentario.SetFocus
'        Exit Function
    ElseIf cbo_ContaOrcamentario.ListIndex = -1 Then
        ExibeMensagem "A conta tem ser informada corretamente."
        If cbo_ContaOrcamentario.Enabled Then cbo_ContaOrcamentario.SetFocus
        Exit Function
    End If
    blnDadosOrcamentarioOk = True
End Function
Private Function blnDadosRetencaoOk() As Boolean
    If Val(gstrConvVrParaSql(txt_ValorRetencao)) = 0 Then
        ExibeMensagem "O valor tem ser informado corretamente."
        If txt_ValorRetencao.Enabled Then txt_ValorRetencao.SetFocus
        Exit Function
    ElseIf Not (Val(gstrConvVrParaSql(lblLiquido)) > 0) Then
        If lvw_Retencao.ListItems.Count > 0 Then
            ExibeMensagem "O valor da parcela tem que ser superior ao total dos lançamentos."
        Else
            ExibeMensagem "O valor da parcela tem que ser superior ao valor do lançamento."
        End If
        If txt_ValorRetencao.Enabled Then txt_ValorRetencao.SetFocus
        Exit Function
    ElseIf cbo_ContaRetencao.ListIndex = -1 Then
        ExibeMensagem "A conta tem ser informada corretamente."
        If cbo_ContaRetencao.Enabled Then cbo_ContaRetencao.SetFocus
        Exit Function
    End If
    blnDadosRetencaoOk = True
End Function

Private Sub VerificaGravacaoLiquidacao()
   Dim intInd As Integer
   Dim mblnGravouNotasFiscais As Boolean
   Dim mblnGravouExtras As Boolean
   Dim mblnGravouOrcamentario As Boolean
   
   
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
        
    If mblnCriarParcelaLiquidada Then
        intItemIndex = -1
    Else
        intItemIndex = lvw_Liquidacao.SelectedItem.Index
    End If
    
    
    If Trim(lblParcela.Caption) = "" Then
        lvw_Liquidacao.SelectedItem = Nothing
    End If
    
    If lvw_Liquidacao.SelectedItem Is Nothing And mblnCriarParcelaLiquidada Then
       If GravaSubEmpenhoLiquidado = False Then Exit Sub
       lvw_Liquidacao.ListItems(lvw_Liquidacao.ListItems.Count).Selected = True
       intItemIndex = lvw_Liquidacao.SelectedItem.Index
    End If
    
    DoEvents
    mblnGravouNotasFiscais = GravaNotasFiscais
    
    If lvw_Liquidacao.SelectedItem Is Nothing Then
        ExibeMensagem "É necessário selecionar uma parcela para liquidar"
        If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
        
        Set gobjBanco = New clsBanco
        
           gobjBanco.ExecutaRollbackTrans
        
        LeSubEmpenho lvw_Liquidacao, 2, txtPKId
        LeSubEmpenho lvw_ListaSubempenho, , txtPKId
        LeSubEmpenho lvw_Anulacao, 4, txtPKId
        lvw_Liquidacao.ListItems(intItemIndex).Selected = True
        
        Exit Sub
    Else
        If lvw_Liquidacao.SelectedItem.SubItems(4) <> "Programada" Then
            ExibeMensagem "Somente parcelas programadas podem ser liquidadas"
            If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
            
            Set gobjBanco = New clsBanco
            
            gobjBanco.ExecutaRollbackTrans
            
            LeSubEmpenho lvw_Liquidacao, 2, txtPKId
            LeSubEmpenho lvw_ListaSubempenho, , txtPKId
            LeSubEmpenho lvw_Anulacao, 4, txtPKId
            lvw_Liquidacao.ListItems(intItemIndex).Selected = True
            
            Exit Sub
        End If
    End If
    
    
    
    If Val(gstrConvVrParaSql(txt_dblDesconto)) < 0 Then
        ExibeMensagem "O Valor de desconto não pode ser negativo."
        If txt_dblDesconto.Enabled Then txt_dblDesconto.SetFocus
        
        Set gobjBanco = New clsBanco
        
        gobjBanco.ExecutaRollbackTrans
        
        LeSubEmpenho lvw_Liquidacao, 2, txtPKId
        LeSubEmpenho lvw_ListaSubempenho, , txtPKId
        LeSubEmpenho lvw_Anulacao, 4, txtPKId
        lvw_Liquidacao.ListItems(intItemIndex).Selected = True
        
        Exit Sub
    End If
    
    
    'grava extra
    'If gblnExclusaoGravacaoOk(IIf(mblnAlterandoExtra, "A", "I")) Then
    msgExtras = ""
        With lvw_Extra
           For intInd = 1 To .ListItems.Count
               .ListItems(intInd).Selected = True
               If .ListItems(intInd).SubItems(3) = "0" Then
                  If blnGravouLiquidacao(.SelectedItem.Tag, _
                                         mlngPKIdExtra, _
                                         .ListItems(intInd).SubItems(2), 1, _
                                         mblnAlterandoExtra) Then
                                         
                      mblnGravouExtras = True
                      msgExtras = "Gravação de Extras : " & "Valores Extras gravados com sucesso."
                      'LimpaDadosExtra
                      'LeLiquidacaoExtra
                      'AtualizaListView lvw_Extra, txt_ValorExtra, lblExtra
                  Else
                      mblnGravouExtras = False
                      msgExtras = "Erro ao gravar Extras : " & "Valores extras não foram gravados com sucesso." & vbNewLine & "A gravação foi abortada e nenhum lançamento foi gravado."
                      
                      Set gobjBanco = New clsBanco
                      
                      gobjBanco.ExecutaRollbackTrans
                      
                      ExibeMensagem msgExtras
                      
                      LeSubEmpenho lvw_Liquidacao, 2, txtPKId
                      'Atualiza ListView do subempenho após erro de gravação para naum gerar erro no saldo da parcela 0
                      LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                      LeSubEmpenho lvw_Anulacao, 4, txtPKId
                      PreencheSaldoEmpenho
                      lvw_Liquidacao.ListItems(intItemIndex).Selected = True
                      
                      Exit Sub
                  End If
               End If
           Next
           
              If mblnGravouExtras And blnGerarDespesaExtra Then
                 If Not GravaDespesaExtra Then
                    ExibeMensagem "Não foi possível gravar as Despesas Extra Orçamentárias porque há contas que não possuem vínculo com um Credor. Verifique as contas e seus vínculos"
                    ExibeMensagem "Liquidação cancelada"
                                        
                    Set gobjBanco = New clsBanco
                      
                    gobjBanco.ExecutaRollbackTrans
                      
                    LeSubEmpenho lvw_Liquidacao, 2, txtPKId
                    LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                    LeSubEmpenho lvw_Anulacao, 4, txtPKId
                    PreencheSaldoEmpenho
                    lvw_Liquidacao.ListItems(intItemIndex).Selected = True
                      
                    Exit Sub
                 End If
              End If
              
        End With
        
    'End If
    
'GRAVA RETENCAO ORCAMENTARIA ------------------------------------------------------
        msgOrcamentario = ""
        With lvw_Orcamentario
           For intInd = 1 To .ListItems.Count
               .ListItems(intInd).Selected = True
               If .ListItems(intInd).SubItems(3) = "0" Then
                  If blnGravouLiquidacaoOrcamentario(.SelectedItem.Tag, _
                                                     mlngPKIdExtra, _
                                                     .ListItems(intInd).SubItems(2), 1, _
                                                     mblnAlterandoExtra) Then
                                         
                      mblnGravouOrcamentario = True
                      msgOrcamentario = "Gravação de Retenção Orçamentaria : " & "Valores de Retenção Orcammentária gravados com sucesso."
                      'LimpaDadosExtra
                      'LeLiquidacaoExtra
                      'AtualizaListView lvw_Extra, txt_ValorExtra, lblExtra
                  Else
                      mblnGravouOrcamentario = False
                      msgOrcamentario = "Erro ao gravar Retenção Orçamentaria : " & "Valores de Retenção Orcamentária não foram gravados com sucesso." & vbNewLine & "A gravação foi abortada e nenhum lançamento foi gravado."
                      
                      Set gobjBanco = New clsBanco
                      
                      gobjBanco.ExecutaRollbackTrans
                      
                      ExibeMensagem msgOrcamentario
                      
                      LeSubEmpenho lvw_Liquidacao, 2, txtPKId
                      
                      LeSubEmpenho lvw_Anulacao, 4, txtPKId
                      'Atualiza ListView do subempenho após erro de gravação para naum gerar erro no saldo da parcela 0
                      LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                      PreencheSaldoEmpenho
                      lvw_Liquidacao.ListItems(intItemIndex).Selected = True
                      
                      Exit Sub
                  End If
               End If
           Next
        End With
        

'----------------------------------------------------------------------------------
 
    
    
    
'    Liquidação de parcela
    If blnDadoLiquidacaoOK(mblnGravouExtras, mblnGravouNotasFiscais, True, mblnGravouOrcamentario) Then
        'If gblnExclusaoGravacaoOk("", "Confirma liquidação da parcela?", True) Then
            If blnLiquidouParcela Then
                txt_DataLiuidacao = ""
                cbo_HistoricoLiquidacao = ""
                cbo_intEventoLiq.ListIndex = -1
                If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
            Else
               
               Set gobjBanco = New clsBanco
               
               gobjBanco.ExecutaRollbackTrans
                  
               LeSubEmpenho lvw_Anulacao, 4, txtPKId
               LeSubEmpenho lvw_Liquidacao, 2, txtPKId
               'Atualiza ListView do subempenho após erro de gravação para naum gerar erro no saldo da parcela 0
               LeSubEmpenho lvw_ListaSubempenho, , txtPKId
               PreencheSaldoEmpenho
               lvw_Liquidacao.ListItems(intItemIndex).Selected = True
               
               Exit Sub
            End If
        'End If
    Else
       
       Set gobjBanco = New clsBanco
       
       gobjBanco.ExecutaRollbackTrans
       
       LeSubEmpenho lvw_Liquidacao, 2, txtPKId
       'Atualiza ListView do subempenho após erro de gravação para naum gerar erro no saldo da parcela 0
       LeSubEmpenho lvw_ListaSubempenho, , txtPKId
       LeSubEmpenho lvw_Anulacao, 4, txtPKId
       PreencheSaldoEmpenho
       
       blnMantemItemSel = True 'variavel usada para manter o grid selecionado neste item, pois o mesmo é atualizado e o item selecionado muda
       
       
       Exit Sub
    End If
            
    
'        If blnDadosRetencaoOk And blnDadoLiquidacaoOK Then
'            If blnGravouLiquidacao(gstrItemData(cbo_ContaRetencao), _
'                                   mlngPKIdRetencao, _
'                                   txt_ValorRetencao, 2, _
'                                   mblnAlterandoRetencao) Then
'                LimpaDadosRetencao
'                LeLiquidacaoRetencao
'            End If
'        End If
   
   Set gobjBanco = New clsBanco
   
   gobjBanco.ExecutaCommitTrans
   LeLiquidacaoOrcamentario
   LeLiquidacaoExtra
   
   If mblnCriarParcelaLiquidada Then
      LimpaDadosLiquidacao
      mblnCriarParcelaLiquidada = False
   End If
End Sub

Sub LimpaDadosExtra(Optional blnFechaCampos As Boolean)
    
    TrocaCorObjeto cmd_ContaExtra, blnFechaCampos
    TrocaCorObjeto cbo_ContaExtra, blnFechaCampos
    TrocaCorObjeto cbo_DescricaoExtra, blnFechaCampos
    TrocaCorObjeto txt_ValorExtra, blnFechaCampos
    txt_ValorExtra = ""
    cbo_ContaExtra.ListIndex = -1
    If cbo_ContaExtra.Enabled Then cbo_ContaExtra.SetFocus
    mblnAlterandoExtra = False
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
End Sub

Sub LimpaDadosOrcamentario(Optional blnFechaCampos As Boolean)
    
    TrocaCorObjeto cmd_ContaOrcamentario, blnFechaCampos
    TrocaCorObjeto cbo_ContaOrcamentario, blnFechaCampos
    TrocaCorObjeto cbo_DescricaoOrcamentario, blnFechaCampos
    TrocaCorObjeto txt_ValorOrcamentario, blnFechaCampos
    txt_ValorOrcamentario = ""
    cbo_ContaOrcamentario.ListIndex = -1
    If cbo_ContaOrcamentario.Enabled Then cbo_ContaOrcamentario.SetFocus
    mblnAlterandoOrcamentario = False
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
End Sub


Sub LimpaDadosRetencao()
    txt_ValorRetencao = ""
    cbo_ContaRetencao.ListIndex = -1
    If cbo_ContaRetencao.Enabled Then cbo_ContaRetencao.SetFocus
    mblnAlterandoRetencao = False
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
End Sub
Sub LimpaDadosReserva()
   
   txt_Reservado = ""
   txt_Cancelado = ""
   txt_Empenhado = ""
   txt_Saldo = ""
   cbointReservaDotacao.Text = ""
   TrocaCorObjeto cbointReservaDotacao, False
   TrocaCorObjeto cmd_Reserva, False
   
End Sub
Sub LimpaDadosLiquidacao()
    txt_DataLiuidacao = ""
    txt_DataVencto = ""
    txt_HistoricoLiquidacao = ""
    lblParcela = ""
    lblRetencao = "0,00"
    lblExtra = "0,00"
    lblLiquido = "0,00"
    txt_dblDesconto = "0,00"
    txt_dblValorAux = ""
    tab_3DPastaLiquidacao.Tab = 0
    HabilitaDesabilitaTab tab_3DPastaLiquidacao, False
End Sub

Sub GravaComplemento()

    If blnComplementoOK Then
        If mblnAlterandoComplemento Then
            If gblnExclusaoGravacaoOk("", "Confirma alteração deste complemento?", True) Then
               If blnAtualizouParcela(lvw_Complemento.SelectedItem.Tag, _
                                      txt_DataComplemento, _
                                      txt_ValorComplemento, _
                                      txt_HistoricoComplemento) = False Then
                   Exit Sub
               End If
               LimpaTelaComplemento
            End If
        Else
            If gblnExclusaoGravacaoOk("", "Confirma criação deste complemento?", True) Then
               If blnIncluiuParcela(txt_DataComplemento, txt_ValorComplemento, _
                                    txt_HistoricoComplemento, 2) = False Then
                  Exit Sub
               Else
                  
                  If Not GeraMovimentosByEvento(gstrItemData(cbo_intEvento), txt_DataComplemento, Str(CDbl(txt_ValorComplemento)), txt_HistoricoComplemento, txtintNumero, "3") Then
                     ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                  End If
                  
               End If
           End If
        End If
        LeSubEmpenho lvw_Complemento, 5, txtPKId
        mblnAtualizaTelaSubempenho = True
        LimpaTelaComplemento
    End If
End Sub

Private Function blnComplementoOK()
    Dim dtmDtEncerramento        As Date
    Dim dblComplemetoMaisEmpenho As Double
    dblComplemetoMaisEmpenho = Val(gstrConvVrParaSql(txtdblValor)) _
                             + Val(gstrConvVrParaSql(lblTotalComplemento))
    If Trim(txt_DataComplemento.Text) = "" Then
        ExibeMensagem "A data de previsão de pagamento deve ser preenchida"
        If txt_DataComplemento.Enabled Then
            txt_DataComplemento.SetFocus
        End If
        Exit Function
    End If
                              
                              
                              
    If Val(gstrConvVrParaSql(txt_ValorComplemento)) = 0 Then
        ExibeMensagem "O valor do complemento tem que ser informado corretamente."
        If txt_ValorComplemento.Enabled Then txt_ValorComplemento.SetFocus
        Exit Function
   'ElseIf dblComplemetoMaisEmpenho > Val(gstrConvVrParaSql(txt_SaldoDotacao)) Then
    ElseIf SaldoDotacaoAtual(gstrItemData(cboProgramaTrabalho, True), Val(Month(CDate(txt_DataComplemento))), gintExercicio, CDbl(txt_ValorComplemento)) = Empty Then
        If txt_ValorComplemento.Enabled Then txt_ValorComplemento.SetFocus
        Exit Function
    ElseIf Val(gstrConvVrParaSql(txt_ValorComplemento)) > Val(gstrConvVrParaSql(txt_SaldoDotacao)) Then
        ExibeMensagem "Não há saldo suficiente para fazer este complemento."
        If txt_ValorComplemento.Enabled Then txt_ValorComplemento.SetFocus
        Exit Function
    ElseIf gblnDataValida(txt_DataComplemento) = False Then
        ExibeMensagem "A data de previsão de pagamento do complemento está incorreta."
        If txt_DataComplemento.Enabled Then txt_DataComplemento.SetFocus
        Exit Function
    ElseIf CVDate(txt_DataComplemento) < CVDate(txtDTMDATA) Then
        ExibeMensagem "A data de previsão de pagamento do complemento não poder ser inferior a data do empenho."
        If txt_DataComplemento.Enabled Then txt_DataComplemento.SetFocus
        Exit Function
    End If
    
    dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
        
    If dtmDtEncerramento = Empty Then
       Exit Function
    Else
       If CDate(txt_DataComplemento) <= dtmDtEncerramento Then
          ExibeMensagem "A data de previsão de pagamento do complemento deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
          If txt_DataComplemento.Enabled Then txt_DataComplemento.SetFocus
          Exit Function
       End If
    End If
    
    'ORC677
    If IsDate(txt_DataComplemento.Text) Then
        If Year(CDate(txt_DataComplemento)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data da de complemento tem que estar no exercício de " & gintExercicio & "."
            If txt_DataComplemento.Enabled Then txt_DataComplemento.SetFocus
            Exit Function
        End If
    End If
    
    blnComplementoOK = True
End Function

Private Function blnDadoDaParcelaOk(dblValor As Double) As Boolean
    Dim dtmDtEncerramento As Date
    Dim adoResultado      As ADODB.Recordset
    Dim strSQL            As String
    
'    If mblnAlterandoSubEmpenho Then
'        ExibeMensagem "Não é permitido alterar os dados de uma parcela."
'        Exit Function
    If UCase(lvw_ListaSubempenho.ListItems(1).SubItems(3)) = "PAGA" Then
        ExibeMensagem "Não é possível inserir parcelas quando o empenho está Pago."
        Exit Function
    ElseIf lvw_ListaSubempenho.ListItems(1).SubItems(3) = "Liquidada" Then
        ExibeMensagem "Não é possível inserir parcelas quando o empenho está liquidado."
        Exit Function
    ElseIf dblValor < 0 Then
        ExibeMensagem "A soma das parcelas não pode superar o valor do empenhado."
        If txt_ValorParcela.Enabled Then txt_ValorParcela.SetFocus
        Exit Function
    ElseIf gblnDataValida(txt_DataParcela) = False Then
        ExibeMensagem "Data da parcela incorreta."
        If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
        Exit Function
    ElseIf Not mblnAlterandoSubEmpenho And lvw_ListaSubempenho.ListItems.Count > 0 And _
            CDate(txt_DataParcela) < gdtmMaiorSubEmpenho Then
        ExibeMensagem "Data da nova parcela não pode ser inferior à data da última parcela."
        If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
        Exit Function
    ElseIf CVDate(txtDTMDATA) > CVDate(txt_DataParcela) Then
        ExibeMensagem "Data da parcela não poder ser inferior a data do empenho."
        If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
        Exit Function
    End If
    
    
    'ORC677
    If IsDate(txt_DataParcela.Text) Then
        If Year(CDate(txt_DataParcela)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data da parcela tem que estar no exercício de " & gintExercicio & "."
            If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
            Exit Function
        End If
    End If

    dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
        
    If dtmDtEncerramento = Empty Then
       Exit Function
    Else
       If CDate(txt_DataParcela) <= dtmDtEncerramento Then
          ExibeMensagem "A data da parcela deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
          If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
          Exit Function
       End If
    End If
    
    If blnLiqAutomatica And tab_3dPasta.Tab = 1 Then
       
        dtmDtEncerramento = VerificaDataEncerramento("EF", gintExercicio)
            
        If dtmDtEncerramento = Empty Then
           Exit Function
        Else
           If CDate(txt_DataParcela) <= dtmDtEncerramento Then
              ExibeMensagem "A data da liquidação deve ser maior que a data de último encerramento Financeiro (" & dtmDtEncerramento & ")."
              If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
              Exit Function
           End If
        End If
       
       
       If Trim(txt_strNotasFiscaisLiqAutomatica.Text) = "" And Not mblnAlterandoSubEmpenho Then
            ExibeMensagem "A Nota Fiscal tem que ser informada."
            If txt_strNotasFiscaisLiqAutomatica.Enabled Then txt_strNotasFiscaisLiqAutomatica.SetFocus
            Exit Function
        End If
        
       'A mensagem de critica esta dentro da rotina que é usada tambem na rotina que cria parcela diretamente da guia liquidação
        If gblnMaiorSubEmpenhoLiq(txt_DataParcela) = False Then
               Exit Function
        End If
        
        
        If gblnDataValida(txt_dtmVenctoLiqAutomatica) = False Then
            ExibeMensagem "A Data do Vencimento é inválida."
            If txt_dtmVenctoLiqAutomatica.Enabled Then txt_dtmVenctoLiqAutomatica.SetFocus
            Exit Function
        ElseIf CVDate(txt_DataParcela) > CVDate(txt_dtmVenctoLiqAutomatica) Then
            ExibeMensagem "A Data do Vencimento não poder ser inferior a data da Liquidação."
            If txt_dtmVenctoLiqAutomatica.Enabled Then txt_dtmVenctoLiqAutomatica.SetFocus
            Exit Function
        End If
        
        
        
        'M4RC3LØ VERIFICA SE O EVENTO CONTABIL PERTENCE AO EXERCICIO
        If cbo_intEventoLiqAutomatica.ListIndex >= 0 Then
            strSQL = ""
            strSQL = "SELECT intExercicio FROM " & gstrEvento & " WHERE Pkid = " & cbo_intEventoLiqAutomatica.ItemData(cbo_intEventoLiqAutomatica.ListIndex)
            
            Set gobjBanco = New clsBanco
            
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.RecordCount > 0 Then
                    If Trim(Val(txtintExercicioEmpenho.Text)) <> Val(adoResultado!intExercicio) Then
                        ExibeMensagem "Evento Contabil incorreto !!!"
                        Exit Function
                    End If
                End If
            End If
        End If
        
        
        
        
        
        If cbo_intEventoLiqAutomatica.ListIndex = -1 Then
            ExibeMensagem "O Evento Contabil tem que ser informado."
            If cbo_intEventoLiqAutomatica.Enabled Then cbo_intEventoLiqAutomatica.SetFocus
            Exit Function
        End If
    End If
    
    
    blnDadoDaParcelaOk = True
End Function

Private Function gdtmMaiorSubEmpenho() As Date
    Dim mdtRetorno As Date
    Dim mobjItemChild As ListItem
    
    mdtRetorno = CDate("01/01/" & gintExercicio)
    
    For Each mobjItemChild In lvw_ListaSubempenho.ListItems
        If CDate(mobjItemChild.SubItems(1)) > mdtRetorno Then
            If Not (mobjItemChild.SubItems(3) = "Canc. RP" Or mobjItemChild.SubItems(3) = "Anulada" Or mobjItemChild.SubItems(3) = "Anul. Desp.") Then
                mdtRetorno = CDate(mobjItemChild.SubItems(1))
            End If
        End If
    Next
    gdtmMaiorSubEmpenho = mdtRetorno
End Function

Private Function gblnMaiorSubEmpenhoLiq(ByVal mdtmDataComparada As String) As Boolean
    Dim mdtRetorno As Date
    Dim strSQL  As String
    Dim adoResultado As ADODB.Recordset
    
    mdtRetorno = CDate("01/01/" & gintExercicio)
    
    strSQL = "Select MAX(su.dtmliquidacao) dtmliquidacao FROM " & gstrSubempenho & " su where su.bytsituacao = 2 and su.intempenho = " & txtPKId
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            If Not IsNull(adoResultado!dtmLiquidacao) Then
                mdtRetorno = adoResultado!dtmLiquidacao
            End If
        End If
    End If
    
    If CDate(mdtmDataComparada) < mdtRetorno Then
        ExibeMensagem "A data de liquidação não pode ser menor que a data da última liquidação (" & CStr(mdtRetorno) & ")."
        gblnMaiorSubEmpenhoLiq = False
    Else
        gblnMaiorSubEmpenhoLiq = True
    End If

End Function




Sub GravaSubEmpenho()
    Dim dblValor            As Double
    Dim dblValorParcAnt     As Double
    Dim dblValorParcAtu     As Double
    Dim intNumDeParcela     As Integer
    Dim lngNumDeParcelaAnt  As Long
        
    
    intNumDeParcela = lvw_ListaSubempenho.ListItems.Count
    With lvw_ListaSubempenho
        dblValorParcAtu = Val(gstrConvVrParaSql(txt_ValorParcela))
        'dblValorParcAnt = ProcuraValorParcAnt(lngNumDeParcelaAnt)
        lngNumDeParcelaAnt = .ListItems(1).Tag
        dblValorParcAnt = CDbl(.ListItems(1).ListSubItems(2).Text)
        dblValor = dblValorParcAnt - dblValorParcAtu
        If blnDadoDaParcelaOk(dblValor) Then
           If gblnExclusaoGravacaoOk(IIf(mblnAlterandoSubEmpenho, "A", "I"), " desta Parcela") Then
               If Not mblnAlterandoSubEmpenho Then
                  If blnAtualizouParcela(lngNumDeParcelaAnt, _
                                      txt_DataParcela, _
                                      dblValor, _
                                      txt_HistoricoSubEmpenho, _
                                      intNumDeParcela) Then
                  'If dblValor > 0 Then
   '                    If lvw_ListaSubempenho.SelectedItem.Index = intNumDeParcela Then
                          'intNumDeParcela = intNumDeParcela + 1
                          'txt_DataParcela = DateAdd("m", 1, txt_DataParcela)
                                  
                          Dim strValorNovaParc As String
                        If txt_ValorParcela.Text <> "" Then
                            strValorNovaParc = CDbl(txt_ValorParcela.Text)
                        End If
                           
                        If strValorNovaParc = "" Then
                            ExibeMensagem "Valor da parcela está incorreto"
                            Exit Sub
                        End If
                                            
                          
                                  
                          If blnIncluiuParcela(txt_DataParcela, _
                                               strValorNovaParc, _
                                               txt_HistoricoSubEmpenho, 1) = False Then
                              ExibeMensagem "Ocorreram erros durante a gravação da parcela"
                              'Exit Sub
                          End If
   '                    Else
   '                        If gblnEncontroItemNoListView(lvw_ListaSubempenho, _
   '                                                      .ListItems(.SelectedItem.Index + 1).Tag, _
   '                                                      lvwTag) Then
   '                            txtdtmdataParcela = .ListItems(.SelectedItem.Index).SubItems(1)
   '                            dblValorParcAnt = Val(gstrConvVrParaSql(.ListItems(.SelectedItem. _
   '                                              Index).SubItems(2)))
   '                            dblValor = dblValor + dblValorParcAnt
   '                            If blnAtualizouParcela(lvw_ListaSubempenho.SelectedItem.Tag, _
   '                                                   txtdtmdataParcela, _
   '                                                   dblValor, _
   '                                                   txtstrhistoricoSubEmpenho) = False Then
   '                                Exit Sub
   '                            End If
   '                        End If
   '                    End If
                  'End If
                   'OrganizaNumParcelas
                   LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                   LeSubEmpenho lvw_Anulacao, 4, txtPKId
                   LeSubEmpenho lvw_Liquidacao, 2, txtPKId
                   PreencheSaldoEmpenho
                   LimpaTelaSubempenho
                  Else
                    ExibeMensagem "Ocorreram erros durante a atualização da parcela"
                  End If
              Else
                 
                If blnModificaParcela(lvw_ListaSubempenho.SelectedItem.Tag, _
                                      txt_DataParcela, _
                                      dblValor, _
                                      txt_HistoricoSubEmpenho, _
                                      intNumDeParcela) Then
                     
                     ExibeMensagem "Parcela modificada com sucesso."
                     LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                     LeSubEmpenho lvw_Anulacao, 4, txtPKId
                     LeSubEmpenho lvw_Liquidacao, 2, txtPKId
                Else
                    ExibeMensagem "Ocorreram erros durante a modificação da parcela."
                End If
              End If
           End If
        End If
    End With
End Sub

Private Function ProcuraValorParcAnt(lngNumDeParcelaAnt As Long) As Double
Dim intCont As Integer
    
    With lvw_ListaSubempenho
        
        intCont = .ListItems.Count
        
        Do While intCont > 0 And .ListItems(intCont).SubItems(6) <> 1
            intCont = intCont - 1
        Loop
        
        lngNumDeParcelaAnt = .ListItems(intCont).Tag
        If intCont > 0 Then ProcuraValorParcAnt = .ListItems(intCont).SubItems(2)
        
    End With
End Function

Private Function LeSubEmpenho(lvw_ListaParcela As ListView, Optional bytFlag As Byte, Optional strPKId As String)
    Dim intCont As Integer
    Dim strSQL  As String
    Dim adoResultado As ADODB.Recordset
    
    If strPKId = "" Then Exit Function
    LeDaTabelaParaObj "", lvw_ListaParcela, strQuerySubempenho(bytFlag, strPKId)
    DesmarcaIntemListView lvw_ListaParcela
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar, gstrCancelar
    
    If lvw_ListaParcela.Name = "lvw_Liquidacao" Then
        For intCont = 1 To lvw_ListaParcela.ListItems.Count
            lvw_ListaParcela.ListItems(intCont).ListSubItems(6).Text = gstrConvVrDoSql(lvw_ListaParcela.ListItems(intCont).ListSubItems(6).Text, 2)
            If lvw_ListaParcela.ListItems(intCont).ListSubItems(4).Text = "Liquidada" Then
                strSQL = "Select strHistorico from " & gstrmovliq & " tml, "
                strSQL = strSQL & "(select Max(pkid) pkid  from "
                strSQL = strSQL & "(select pkid, strHistorico from " & gstrmovliq & " where intParcela=" & lvw_ListaParcela.ListItems(intCont).Tag & " ) dados) maiorPkid "
                strSQL = strSQL & "where tml.pkid = maiorPkid.pkid"
                 
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                         
                   lvw_ListaParcela.ListItems(intCont).ListSubItems(11).Text = gstrENulo(adoResultado!STRHISTORICO)
                   
                  
                End If
                
            End If
        Next
    End If
    
    If lvw_ListaParcela.Name = "lvw_Anulacao" Then
        For intCont = 1 To lvw_ListaParcela.ListItems.Count
            lvw_ListaParcela.ListItems(intCont).ListSubItems(2).Text = gstrConvVrDoSql(lvw_ListaParcela.ListItems(intCont).ListSubItems(2).Text, 2)
        Next
    End If
    
    If lvw_ListaParcela.Name = "lvw_ListaSubempenho" Then
        For intCont = 1 To lvw_ListaParcela.ListItems.Count
            lvw_ListaParcela.ListItems(intCont).ListSubItems(2).Text = gstrConvVrDoSql(lvw_ListaParcela.ListItems(intCont).ListSubItems(2).Text, 2)
        Next
    End If
    
    For intCont = 1 To lvw_ListaParcela.ListItems.Count
       If lvw_ListaParcela.ListItems(intCont).Text = "" Then
          lvw_ListaParcela.ListItems(intCont).Text = 0
       End If
    Next
End Function

Sub LimpaTelaSubempenho()
    TrocaCorObjeto txt_DataParcela, False
    TrocaCorObjeto txt_dtmVenctoLiqAutomatica, False
    TrocaCorObjeto txt_ValorParcela, False
    TrocaCorObjeto txt_strNotasFiscaisLiqAutomatica, False
    TrocaCorObjeto cmd_EventoLiqAutomatica, False
    TrocaCorObjeto cbo_intEventoLiqAutomatica, False
    TrocaCorObjeto txt_codEventoLiqAutomatica, False
    txt_CodHistoricoSub.Text = ""
    txt_DataParcela = ""
    txt_dtmVenctoLiqAutomatica = ""
    txt_ValorParcela = ""
    txtNumParcela = ""
    cboPeriodo.ListIndex = -1
    txt_HistoricoSubEmpenho = txtstrHistorico
    cbo_HistoricoSubEmpenho.ListIndex = -1
    mblnAlterandoSubEmpenho = False
    cbo_intEventoLiqAutomatica.ListIndex = -1
    txt_strNotasFiscaisLiqAutomatica.Text = ""
    If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
End Sub

Private Sub VerificaTabExclusao()
    Select Case tab_3dPasta.Tab
    Case 1
        DeletaParcela
    Case 2
        
    Case 3
        VerificaTabLiquidacao
    Case 4
       'VerificaTabDeletaAnulacao
    End Select
End Sub

Sub VerificaTabLiquidacao()
    If tab_3DPastaLiquidacao.Tab = 1 Then
        If blnDeletouLancamento(1, gstrItemData(cbo_ContaExtra)) Then
            LimpaDadosExtra
            LeLiquidacaoExtra
        End If
    ElseIf tab_3DPastaLiquidacao.Tab = 2 Then
        If blnDeletouLancamento(2, gstrItemData(cbo_ContaRetencao)) Then
            LimpaDadosRetencao
            LeLiquidacaoRetencao
        End If
    ElseIf tab_3DPastaLiquidacao.Tab = 3 Then
        If blnDeletouLancamentoOrcamentario Then
            LimpaDadosOrcamentario
            LeLiquidacaoOrcamentario
        End If
     
    End If
End Sub

Sub DeletaParcela()
    Dim strSQL                  As String
    Dim dblValorParcAExcluir    As Double
    Dim dblValorParcASomar      As Double
    Dim lngPKIdAExcluir         As Long
    Dim lngPKIdASomar           As Long
    With lvw_ListaSubempenho
        dblValorParcAExcluir = Val(gstrConvVrParaSql(.ListItems(.SelectedItem.Index).SubItems(2)))
        lngPKIdAExcluir = .SelectedItem.Tag
        If blnHaItemParaSomar(.SelectedItem.Index, _
                              lngPKIdASomar, _
                              dblValorParcASomar) Then
            If UCase(.ListItems(.SelectedItem.Index).SubItems(4)) <> "COMPLEMENTO" Then
               'dblValorParcASomar = dblValorParcASomar + dblValorParcAExcluir
               dblValorParcASomar = dblValorParcAExcluir
            Else
               dblValorParcASomar = 0
            End If
            
            If gblnExclusaoGravacaoOk("E", "Confirma Exclusão da Parcela?", True) Then
                strSQL = ""
                strSQL = strSQL & "DELETE " & gstrSubempenho & " "
                strSQL = strSQL & "WHERE PKId = " & lngPKIdAExcluir
                Set gobjBanco = New clsBanco
                If gobjBanco.Execute(strSQL) Then
'------ Atualiza uma parcela com a soma dos valores
                   strSQL = "UPDATE " & gstrSubempenho & " "
                   strSQL = strSQL & "SET dblValor = dblValor +  "
                   strSQL = strSQL & gstrConvVrParaSql(dblValorParcASomar) & " "
                   strSQL = strSQL & "WHERE PKId = " & lngPKIdASomar
                   If gobjBanco.Execute(strSQL) Then
                       'OrganizaNumParcelas
                       LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                       PreencheSaldoEmpenho
                       LimpaTelaSubempenho
                       VerificaComplemento
                   End If
                End If
            End If
        Else
            ExibeMensagem "Não há parcela programada para somar o valor da parcela excluída."
        End If
    End With
End Sub

Private Function blnHaItemParaSomar(intIndiceSelecionado As Integer, _
                                    lngPKIdASomar As Long, _
                                    dblValorParcASomar As Double) As Boolean
    Dim intInd            As Integer
    Dim dtmDtEncerramento As Date
    With lvw_ListaSubempenho
        If UCase(.ListItems(.SelectedItem.Index).SubItems(3)) <> "PROGRAMADA" Then
            ExibeMensagem "Não é permitido excluir parcela anulada ou liquidada."
            Exit Function
        End If
        
        dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
        
        If dtmDtEncerramento = Empty Then
           Exit Function
        Else
           If CDate(.ListItems(.SelectedItem.Index).SubItems(1)) <= dtmDtEncerramento Then
              ExibeMensagem "Não é permitido excluir parcela com data inferior ou igual a data de último encerramento (" & dtmDtEncerramento & ")."
              Exit Function
           End If
        End If
        
        If intIndiceSelecionado = 1 Then Exit Function
        
        For intInd = .ListItems.Count To (intIndiceSelecionado + 1) Step -1
            If .ListItems(intInd).SubItems(6) = 1 Then
                blnHaItemParaSomar = True
                'lngPKIdASomar = .ListItems(intInd).Tag
                lngPKIdASomar = .ListItems(1).Tag
                dblValorParcASomar = Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(2)))
                Exit Function
            End If
        Next
        For intInd = (intIndiceSelecionado - 1) To 1 Step -1
            If .ListItems(intInd).SubItems(6) = 1 Then
                blnHaItemParaSomar = True
                'lngPKIdASomar = .ListItems(intInd).Tag
                lngPKIdASomar = .ListItems(1).Tag
                dblValorParcASomar = Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(2)))
                Exit Function
            End If
        Next
    End With
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
    
End Sub

Private Sub lblLiquido_Change()
    If Val(gstrConvVrParaSql(lblLiquido)) < 0 Then
        ExibeMensagem "O Valor Líquido não pode ser negativo."
        If txt_dblDesconto.Enabled Then txt_dblDesconto.SetFocus
        txt_dblDesconto = Mid(txt_dblDesconto, 1, Len(txt_dblDesconto) - 1)
    End If

End Sub


Private Function blnTestaLiquidacao() As Boolean
        Dim adoResultado     As ADODB.Recordset
    
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT SE.PKId, SE.intNumero "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP, " & gstrSubempenho & " SE "
    strSQL = strSQL & "WHERE SE.bytSituacao = 2 "
    strSQL = strSQL & "AND SE.Pkid = " & lvw_Liquidacao.SelectedItem.Tag
    strSQL = strSQL & " AND EP.intNumero = " & Trim(txtintNumero.Text)
    
    strSQL = strSQL & " AND NOT SE.PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela) "
    strSQL = strSQL & " ORDER BY SE.intNumero"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            blnTestaLiquidacao = True
        Else
            blnTestaLiquidacao = False
        End If
    Else
        blnTestaLiquidacao = False
    End If
    
End Function

Private Sub lvw_Liquidacao_DblClick()
    If mobjAux Is Nothing = False Then
        MantemForm gstrAplicar
    End If
End Sub

Private Sub lvw_Liquidacao_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 3, tab_3DPastaLiquidacao, 0
End Sub

Private Sub lvw_Orcamentario_GotFocus()
   VerificaTabAtivo
   mAtivaPastaDeObjeto tab_3dPasta, 3, tab_3DPastaLiquidacao, 3
End Sub


Private Sub lvw_Orcamentario_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    TrocaCorObjeto cbo_ContaOrcamentario, True
    TrocaCorObjeto cbo_DescricaoOrcamentario, True
    TrocaCorObjeto txt_ValorOrcamentario, True
    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrExcluirItem
    
    mblnAlterandoOrcamentario = True
    With lvw_Orcamentario.ListItems(lvw_Orcamentario.SelectedItem.Index)
        txt_ValorOrcamentario = .SubItems(2)
        mdblValorOrcamentario = gstrConvVrDoSql(.SubItems(2))
        If .SubItems(3) = "0" Then
            cbo_ContaOrcamentario.ListIndex = gintIndiceCBO(cbo_ContaOrcamentario, _
                          .Tag)
        Else
            cbo_ContaOrcamentario.ListIndex = gintIndiceCBO(cbo_ContaOrcamentario, _
                                      .SubItems(3))
        End If
        mlngPKIdOrcamentario = lvw_Orcamentario.SelectedItem.Tag
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
    End With
    
    VerificaTabAtivo
End Sub


Private Sub tab_3DGeral_GotFocus()
'mAtivaPastaDeObjeto tab_3dPasta, 0

 
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub txt_codEvento_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DGeral, 0
End Sub
Private Sub txt_codEventoLiq_GotFocus()

mAtivaPastaDeObjeto tab_3dPasta, 3

End Sub

Private Sub txt_CodHistorico_GotFocus()
    MarcaCampo txt_CodHistorico
End Sub

Private Sub txt_CodHistorico_LostFocus()
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    Set gobjBanco = New clsBanco
    
    strSQL = "SELECT h.StrDescricao FROM " & gstrHistorico & " H WHERE H.STRCODIGO = '" & Me.txt_CodHistorico.Text & "'"
    gobjBanco.CriaADO strSQL, 10, adoResultado
    With adoResultado
        If Not .EOF Then
            cbo_Historico.Text = gstrENulo(!strDescricao)
            txtstrHistorico.Text = gstrENulo(!strDescricao)
        Else
            cbo_Historico.Text = ""
            txtstrHistorico.Text = ""
        End If
    End With
End Sub

Private Sub txt_CodHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CodHistorico
End Sub

Private Sub txt_CodHistoricoAnl_GotFocus()
    MarcaCampo txt_CodHistoricoAnl
End Sub

Private Sub txt_CodHistoricoAnl_LostFocus()
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    Set gobjBanco = New clsBanco
    
    strSQL = "SELECT h.StrDescricao FROM " & gstrHistorico & " H WHERE H.STRCODIGO = '" & txt_CodHistoricoAnl.Text & "'"
    gobjBanco.CriaADO strSQL, 10, adoResultado
    With adoResultado
        If Not .EOF Then
            txt_HistoricoAnulacao.Text = gstrENulo(!strDescricao)
            cbo_HistoricoAnulacao.Text = gstrENulo(!strDescricao)
        Else
            txt_HistoricoAnulacao.Text = ""
            cbo_HistoricoAnulacao.Text = ""
        End If
    End With
End Sub

Private Sub txt_CodHistoricoAnl_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CodHistoricoAnl
End Sub
Private Sub txt_CodHistoricoComp_GotFocus()
    MarcaCampo txt_CodHistoricoComp
End Sub

Private Sub txt_CodHistoricoComp_LostFocus()
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    Set gobjBanco = New clsBanco
    
    strSQL = "SELECT h.StrDescricao FROM " & gstrHistorico & " H WHERE H.STRCODIGO = '" & txt_CodHistoricoComp.Text & "'"
    gobjBanco.CriaADO strSQL, 10, adoResultado
    With adoResultado
        If Not .EOF Then
            cbo_HistoricoComplemento.Text = gstrENulo(!strDescricao)
            txt_HistoricoComplemento.Text = gstrENulo(!strDescricao)
        Else
            cbo_HistoricoComplemento.Text = ""
            txt_HistoricoComplemento.Text = ""
        End If
    End With
End Sub

Private Sub txt_CodHistoricoComp_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CodHistoricoComp
End Sub

Private Sub txt_CodHistoricoLiq_GotFocus()
    MarcaCampo txt_CodHistoricoLiq
End Sub

Private Sub txt_CodHistoricoLiq_LostFocus()
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    Set gobjBanco = New clsBanco
    
    strSQL = "SELECT h.StrDescricao FROM " & gstrHistorico & " H WHERE H.STRCODIGO = '" & txt_CodHistoricoLiq.Text & "'"
    gobjBanco.CriaADO strSQL, 10, adoResultado
    With adoResultado
        If Not .EOF Then
            txt_HistoricoLiquidacao.Text = gstrENulo(!strDescricao)
            cbo_HistoricoLiquidacao.Text = gstrENulo(!strDescricao)
        Else
            txt_HistoricoLiquidacao.Text = ""
            cbo_HistoricoLiquidacao.Text = ""
        End If
    End With
End Sub

Private Sub txt_CodHistoricoLiq_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CodHistoricoLiq
End Sub
Private Sub txt_CodHistoricoSub_GotFocus()
    MarcaCampo txt_CodHistoricoSub
End Sub

Private Sub txt_CodHistoricoSub_LostFocus()
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    Set gobjBanco = New clsBanco
    
    strSQL = "SELECT h.StrDescricao FROM " & gstrHistorico & " H WHERE H.STRCODIGO = '" & txt_CodHistoricoSub.Text & "'"
    gobjBanco.CriaADO strSQL, 10, adoResultado
    With adoResultado
        If Not .EOF Then
            cbo_HistoricoSubEmpenho.Text = gstrENulo(!strDescricao)
            txt_HistoricoSubEmpenho.Text = gstrENulo(!strDescricao)
        Else
            cbo_HistoricoSubEmpenho.Text = ""
            txt_HistoricoSubEmpenho.Text = ""
        End If
    End With
End Sub

Private Sub txt_CodHistoricoSub_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CodHistoricoSub
End Sub

Private Sub txt_DataVencto_GotFocus()
    MarcaCampo txt_DataVencto
    If mblnselecionou = False Then
    mAtivaPastaDeObjeto tab_3dPasta, 3
    End If
End Sub

Private Sub txt_DataVencto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataVencto
End Sub

Private Sub txt_DataVencto_LostFocus()

    txt_DataVencto = gstrDataFormatada(txt_DataVencto)
        
End Sub

Private Sub txt_dblValorAux_Change()
    lblLiquido = gstrConvVrDoSql(Val(gstrConvVrParaSql(txt_dblValorAux)) - _
                                 (Val(gstrConvVrParaSql(lblExtra)) + _
                                  Val(gstrConvVrParaSql(txt_dblDesconto))))

End Sub

Private Sub lvw_Itens_ItemClick(ByVal Item As MSComctlLib.ListItem)
    HabilitaDesabilitaTab tab_3DGeral, True

    With lvw_Itens
        txt_PkidItem = .SelectedItem.Tag
        
    txt_intCodigo.Text = .SelectedItem.SubItems(1)
    txt_intCatalogoMaterialServico.Text = .SelectedItem.SubItems(2)
    dbc_intStrMarca = .SelectedItem.SubItems(3)
    txt_dblQuantidade.Text = .SelectedItem.SubItems(4)
    txt_dblValorEstimado.Text = gstrConvVrDoSql(.SelectedItem.SubItems(5), 5)
    txt_intUnidadedeMedida.Text = .SelectedItem.SubItems(6)
    txt_strObsItem.Text = .SelectedItem.SubItems(7)
    txt_strdescricaodetalhada.Text = .SelectedItem.SubItems(8)
        
        
    End With
    blnAlterandoItem = True
    dblSaldoItem = gstrConvVrDoSql(txt_dblValorEstimado.Text, 5) * gstrConvVrDoSql(txt_dblQuantidade.Text, 5)
    VerificaTabAtivo
End Sub

Private Sub tab_3DAnulacao_Click(PreviousTab As Integer)
    VerificaTabAtivo
    If mblnAbrindo Then
    If txtintNumero.Enabled = True Then txtintNumero.SetFocus
    mblnAbrindo = False
    End If
End Sub

Private Sub tab_3DGeral_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        If Not blnImportadoPedidoEmpenho Then
            If mblnAlterandoEmpenho = False Then
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
                If txt_intCodigo.Enabled Then txt_intCodigo.SetFocus
            End If
        End If
    ElseIf PreviousTab = 1 Then
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
    VerificaTabAtivo
    If tab_3DGeral.Tab = 1 And txt_intCodigo.Enabled = True Then
            txt_intCodigo.SetFocus
    End If
    
End Sub

Private Sub txt_CodEventoAnul_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_CodEventoAnul_LostFocus()
    PreencheEventobyCodigo txt_CodEventoAnul, cbo_intEventoAnul, "12"
End Sub

Private Sub txt_dblQuantidade_GotFocus()
    MarcaCampo txt_dblQuantidade
End Sub

Private Sub txt_dblQuantidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_dblQuantidade
End Sub

Private Sub txt_dblValorEstimado_GotFocus()
    MarcaCampo txt_dblValorEstimado
End Sub

Private Sub txt_dblValorEstimado_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorEstimado
End Sub

Private Sub txt_dblValorEstimado_LostFocus()
    txt_dblValorEstimado = gstrConvVrDoSql(txt_dblValorEstimado, 5)
End Sub

Private Sub txt_dtmVenctoLiqAutomatica_GotFocus()
    MarcaCampo txt_dtmVenctoLiqAutomatica
    If mblnAbrindo = False Then
    mAtivaPastaDeObjeto tab_3dPasta, 1
    End If
End Sub

Private Sub txt_dtmVenctoLiqAutomatica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmVenctoLiqAutomatica
End Sub

Private Sub txt_dtmVenctoLiqAutomatica_LostFocus()
    txt_dtmVenctoLiqAutomatica = gstrDataFormatada(txt_dtmVenctoLiqAutomatica)
End Sub
Private Sub txt_HistoricoComplemento_GotFocus()
If tab_3dPasta.Tab <> 2 Then
    mAtivaPastaDeObjeto tab_3dPasta, 2
    If txt_DataComplemento.Enabled = True Then
        txt_DataComplemento.SetFocus
 End If
End If
End Sub

Private Sub txt_intCatalogoMaterialServico_GotFocus()
    MarcaCampo txt_intCatalogoMaterialServico
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DGeral, 1
End Sub

Private Sub txt_intCatalogoMaterialServico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intCatalogoMaterialServico
End Sub

Private Sub txt_intCodigo_GotFocus()
    MarcaCampo txt_intCodigo
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DGeral, 1
End Sub

Private Sub txt_intCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intCodigo
End Sub

Private Sub txt_intCodigo_LostFocus()
    If Trim(txt_intCodigo.Text) <> "" Then
        ValidaItem (True)
    Else
        LimpaCamposItem
    End If
End Sub

Private Sub txt_intCodItemDespesa_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii
End Sub

Private Sub txt_intCodItemDespesa_LostFocus()
Dim strPKId As String
Dim strSQL As String
   
   
   txt_intCodItemDespesa = gstrValorSemMascara(txt_intCodItemDespesa)
   
   strPKId = LeCoditemDespesa(, txt_intCodItemDespesa)
   
   
   If strPKId = "" Then
        dbcintItemDespesa.BoundText = ""
        Exit Sub
   End If
   
    If Len(Trim(txt_intCodItemDespesa)) > 0 Then
        If dbcintItemDespesa.Enabled Then dbcintItemDespesa.SetFocus
            strSQL = "SELECT IT.PKID,"
           strSQL = strSQL & " IT.STRDescricao"
           strSQL = strSQL & " FROM "
           strSQL = strSQL & gstrItemDespesa & " IT "
           strSQL = strSQL & " WHERE IT.PKID = " & strPKId
        
           'Cláudio
           'LeDaTabelaParaObj gstrContribuinte, dbcintCredor, "SELECT PKID, strNome FROM " & gstrContribuinte & _
                                                          " WHERE PKID = " & txt_intNContribuinte
                                                          
           LeDaTabelaParaObj gstrItemDespesa, dbcintItemDespesa, strSQL
                                                               
           dbcintItemDespesa.BoundText = strPKId
    End If
    
    txt_intCodItemDespesa = gstrValorSemMascara(txt_intCodItemDespesa)
    txt_intCodItemDespesa = gvntFormatacaoEspecifica(txt_intCodItemDespesa, 4)

End Sub

Private Sub txtintExercicioEmpenho_Change()
    LeTabelaReservaDotacao
    LimpaDadosReserva
End Sub

Private Sub txtintExercicioEmpenho_LostFocus()
Dim intIdx As Integer
   If mblnRestosAPagar Then
   
   
'      If Len(Trim(txtintNumero)) > 0 And BuscaEmpenho(txtintExercicioEmpenho) <> Empty Then
'
'         If Val(txtintExercicioEmpenho) = 0 Or Val(txtintExercicioEmpenho) >= CInt(gintExercicio) Then
'            Exit Sub
'         End If
'
'         mblnClickOk = False
'         mblnAlterandoEmpenho = True
'         mblnAtualizaTelaSubempenho = True
'         gCorLinhaSelecionada tdb_Lista
'         txtPKId.Text = BuscaEmpenho(txtintExercicioEmpenho)
'         HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
'         HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcular
'         HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir
'         mblnSelecionou = True
'         tab_3DPastaLiquidacao.Tab = 0
'         VerificaSubempenho
'         VerificaComplemento
'         VerificaLiquidacao
'         VerificaAnulacao
'
'         If mblnEmpenhoEstorno Then
'            tab_3DPasta.Tab = 4
'            For intIdx = 1 To lvw_Anulacao.ListItems.Count
'                If lvw_Anulacao.ListItems(intIdx).SubItems(9) = txtintNumero Then
'                   lvw_Anulacao.ListItems(intIdx).Selected = True
'                End If
'            Next
'         End If
'
'         LeEmpenho (Str(BuscaEmpenho(txtintExercicioEmpenho)))
'      End If
   End If

End Sub

Private Sub txt_intNContribuinte_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 0
End Sub



Private Sub txt_strNotasFiscaisLiqAutomatica_GotFocus()
    MarcaCampo txt_strNotasFiscaisLiqAutomatica
End Sub

Private Sub txt_strObsItem_GotFocus()
    MarcaCampo txt_strObsItem
End Sub

Private Sub txt_strObsItem_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strObsItem
End Sub

Private Sub txt_dblDesconto_Change()
lblLiquido = gstrConvVrDoSql(Val(gstrConvVrParaSql(txt_dblValorAux)) - _
                                 (Val(gstrConvVrParaSql(lblExtra)) + _
                                  Val(gstrConvVrParaSql(txt_dblDesconto))))
End Sub

Private Sub txt_dblDesconto_GotFocus()
    MarcaCampo txt_dblDesconto
End Sub


Private Sub txt_dblDesconto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblDesconto
End Sub

Private Sub txt_dblDesconto_LostFocus()
    txt_dblDesconto = gstrConvVrDoSql(txt_dblDesconto, 2)
End Sub

Private Sub txt_ValorOrcamentario_LostFocus()
    txt_ValorOrcamentario = gstrConvVrDoSql(txt_ValorOrcamentario)
End Sub



Private Sub txtItemDespSubElemento_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 2
End Sub

Private Sub txtItemDespSubElemento_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "N", txtItemDespSubElemento
End Sub

Private Sub txtItemDespSubElemento_LostFocus()
    Dim strPKId As String
    Dim strSQL As String
   
   txtItemDespSubElemento = gstrValorSemMascara(txtItemDespSubElemento)
   txtItemDespSubElemento = gvntFormatacaoEspecifica(txtItemDespSubElemento, 4)
   
   strPKId = LeCoditemDespesa(, txtItemDespSubElemento)
   
   
   If strPKId = "" Then
        dbcItemDespSubElemento.BoundText = ""
        Exit Sub
   Else
        PreencherListaDeOpcoes dbcItemDespSubElemento, strPKId
        dbcItemDespSubElemento_Click 2
   End If
   
End Sub

Private Sub txtstrCodigo_GotFocus()
mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DGeral, 0
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub
Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub
Private Sub txtbitDigito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigito
End Sub
Private Sub lbl_Codevento_GotFocus()
    If txt_codEvento.Enabled = True Then txt_codEvento.SetFocus
End Sub

Private Sub lblExtra_Change()
lblLiquido = gstrConvVrDoSql(Val(gstrConvVrParaSql(txt_dblValorAux)) - _
                                 (Val(gstrConvVrParaSql(lblExtra)) + _
                                  Val(gstrConvVrParaSql(txt_dblDesconto))))
End Sub

Private Sub lvw_Anulacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    With lvw_Anulacao
        
        cbo_HistoricoAnulacao.ListIndex = -1
        
        If Val(.ListItems(.SelectedItem.Index)) = 0 Then
            TrocaCorObjeto txt_CodEventoAnul, False
            TrocaCorObjeto cbo_intEventoAnul, False
            TrocaCorObjeto cmd_EventoAnul, False
        Else
            txt_CodEventoAnul.Text = ""
            cbo_intEventoAnul.ListIndex = -1
            TrocaCorObjeto txt_CodEventoAnul, True
            TrocaCorObjeto cbo_intEventoAnul, True
            TrocaCorObjeto cmd_EventoAnul, True
        End If
        
        If Trim(.ListItems(.SelectedItem.Index).SubItems(5)) <> "" Then
            txt_DataAnulucao = .ListItems(.SelectedItem.Index).SubItems(5)
        Else
            txt_DataAnulucao = .ListItems(.SelectedItem.Index).SubItems(1)
        End If
        
        txt_HistoricoAnulacao = .ListItems(.SelectedItem.Index).SubItems(6)
        txt_ValorAnulacao = .ListItems(.SelectedItem.Index).SubItems(2)

       If Trim(.ListItems(.SelectedItem.Index)) = "0" Then
          TrocaCorObjeto txt_ValorAnulacao, False
          If Val(.ListItems(.SelectedItem.Index).SubItems(8)) > 0 Then
             'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
             cbo_intEventoAnul.ListIndex = gintIndiceCBO(cbo_intEventoAnul, gstrVerificaCampoNulo(.ListItems(.SelectedItem.Index).SubItems(8)))
          Else
             'HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
             txt_CodEventoAnul.Text = ""
             cbo_intEventoAnul.ListIndex = -1
          End If
       Else
          'M6R PERMITIR SELECIONAR VALOR TAMBEM PARA PARCELAS PAGAS INTNUMERO > 0
          If lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).ListSubItems(3).Text = "Paga" Then
            TrocaCorObjeto txt_ValorAnulacao, False
          Else
            TrocaCorObjeto txt_ValorAnulacao, True
          End If
       End If
       
       LeSubElementos lvwSubElementoEst, , lvw_Anulacao.SelectedItem.Tag
       limpaDadosSubElementosEst
       PreencheCboItemAnulado
       If Not .SelectedItem Is Nothing Then
            If tab_3DAnulacao.TabEnabled(1) = True And .ListItems(.SelectedItem.Index).SubItems(3) = "Programada" Then
                 AbreFechaCamposAnulacao True
             Else
                 AbreFechaCamposAnulacao False
            End If
       End If
       
       If tab_3DAnulacao.TabEnabled(1) = True Then
            LimpaCamposAnulado
            PreencheGridItemAnulado .SelectedItem.Tag
       End If

    End With
    VerificaTabAtivo
    
End Sub


Private Sub lvw_Complemento_ItemClick(ByVal Item As MSComctlLib.ListItem)
   mblnAlterandoComplemento = True
   With lvw_Complemento
        txt_DataComplemento = .ListItems(.SelectedItem.Index).SubItems(1)
        txt_HistoricoComplemento = .ListItems(.SelectedItem.Index).SubItems(5)
        txt_ValorComplemento = .ListItems(.SelectedItem.Index).SubItems(2)
'        If UCase(.ListItems(.SelectedItem.Index).SubItems(3)) <> "PROGRAMADA" Then
'           TrocaCorObjeto txt_ValorComplemento, True
'        Else
'           TrocaCorObjeto txt_ValorComplemento, False
'        End If
        TrocaCorObjeto txt_ValorComplemento, True
    End With
End Sub

Private Sub lvw_Extra_GotFocus()
   
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
   VerificaTabAtivo
   mAtivaPastaDeObjeto tab_3dPasta, 3, tab_3DPastaLiquidacao, 1
   
End Sub

Private Sub lvw_Extra_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    TrocaCorObjeto cbo_ContaExtra, True
    TrocaCorObjeto cbo_DescricaoExtra, True
    TrocaCorObjeto txt_ValorExtra, True
    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrExcluirItem
    
    mblnAlterandoExtra = True
    With lvw_Extra.ListItems(lvw_Extra.SelectedItem.Index)
        txt_ValorExtra = .SubItems(2)
        mdblValorExtra = gstrConvVrDoSql(.SubItems(2))
        
        If .SubItems(3) = "0" Then
            cbo_ContaExtra.ListIndex = gintIndiceCBO(cbo_ContaExtra, _
                                      .Tag)
        Else
            cbo_ContaExtra.ListIndex = gintIndiceCBO(cbo_ContaExtra, _
                                      .SubItems(3))
        End If
        

                                  
        mlngPKIdExtra = lvw_Extra.SelectedItem.Tag
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
    End With
    
    VerificaTabAtivo
End Sub

Private Sub lvw_Liquidacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnAtivarPastas = False
    HabilitaDesabilitaTab tab_3DPastaLiquidacao, True
    
    With lvw_Liquidacao
        
        mblnCriarParcelaLiquidada = False
        
        lblRetencao = "0,00"
        lblExtra = "0,00"
        lblLiquido = "0,00"
        txt_dblDesconto = "0,00"
        txt_dblValorAux = ""
        
        txt_DataLiuidacao = .ListItems(.SelectedItem.Index).SubItems(9)
        
        txt_HistoricoLiquidacao = .ListItems(.SelectedItem.Index).SubItems(10)
        lblParcela = .ListItems(.SelectedItem.Index).Text
        txt_dblValorAux = .ListItems(.SelectedItem.Index).SubItems(3)
        txt_DataVencto = .ListItems(.SelectedItem.Index).SubItems(2)
        lblExtra = strValorRetencao(1)
        lblRetencao = strValorRetencao(2)
        
        
        LeDaTabelaParaObj "", lvw_NotasFiscais, strQueryNotasFiscais
        LeLiquidacaoExtra
        LeLiquidacaoOrcamentario
        txt_dblDesconto = gstrConvVrDoSql(dblValorDesconto(.SelectedItem.Tag), 2)
        If gintIndiceCBO(cbo_intEventoLiq, gstrVerificaCampoNulo(.ListItems(.SelectedItem.Index).SubItems(12))) <> -1 Then
            cbo_intEventoLiq.ListIndex = gintIndiceCBO(cbo_intEventoLiq, gstrVerificaCampoNulo(.ListItems(.SelectedItem.Index).SubItems(12)))
        Else
            If cbo_intEventoLiq.ListCount = 1 And Trim(.ListItems(.SelectedItem.Index).SubItems(4)) = "Programada" Then
                cbo_intEventoLiq.ListIndex = 0
            Else
                cbo_intEventoLiq.ListIndex = -1
            End If
        End If
        
        TrocaCorObjeto txt_dblValorAux, True
        
        If Trim(.ListItems(.SelectedItem.Index).SubItems(4)) <> "Programada" Then
            'desbilita controles tela liquidação
            TrocaCorObjeto txt_DataLiuidacao, True
            TrocaCorObjeto txt_DataVencto, True
            TrocaCorObjeto txt_dblDesconto, True
            TrocaCorObjeto txt_codEventoLiq, True
            TrocaCorObjeto cbo_intEventoLiq, True
            TrocaCorObjeto cmd_EventoLiq, True
            TrocaCorObjeto txt_HistoricoLiquidacao, True
            TrocaCorObjeto cbo_HistoricoLiquidacao, True
            TrocaCorObjeto cmd_HistoricoLiquidacao, True
            'desbilita controles tela extra
            TrocaCorObjeto cbo_ContaExtra, True
            TrocaCorObjeto cbo_DescricaoExtra, True
            TrocaCorObjeto cmd_ContaExtra, True
            TrocaCorObjeto txt_ValorExtra, True
            'desbilita controles notas fiscais
            TrocaCorObjeto txt_dtmDataNF, True
            TrocaCorObjeto txt_dblValorNF, True
            TrocaCorObjeto txt_strNotasFiscais, True
            
        ElseIf Trim(.ListItems(.SelectedItem.Index).SubItems(4)) = "Programada" Then
            TrocaCorObjeto txt_DataLiuidacao, False
            TrocaCorObjeto txt_DataVencto, False
            TrocaCorObjeto txt_dblDesconto, True
            TrocaCorObjeto cbo_intEventoLiq, False
            TrocaCorObjeto txt_codEventoLiq, False
            TrocaCorObjeto cmd_EventoLiq, False
            TrocaCorObjeto txt_HistoricoLiquidacao, False
            TrocaCorObjeto cbo_HistoricoLiquidacao, False
            TrocaCorObjeto cmd_HistoricoLiquidacao, False
            
            'desbilita controles tela extra
            TrocaCorObjeto cbo_ContaExtra, False
            TrocaCorObjeto cbo_DescricaoExtra, False
            TrocaCorObjeto cmd_ContaExtra, False
            TrocaCorObjeto txt_ValorExtra, False
            'desbilita controles notas fiscais
            TrocaCorObjeto txt_dtmDataNF, False
            TrocaCorObjeto txt_dblValorNF, False
            TrocaCorObjeto txt_strNotasFiscais, False
            
            txt_codEventoLiq.Enabled = True
            txt_codEventoLiq.BackColor = vbWindowBackground
            txt_DataLiuidacao = ""
            
        End If
        
        LimpaDadosNF IIf(Trim(.ListItems(.SelectedItem.Index).SubItems(4)) = "Liquidada", True, False)
        LimpaDadosExtra IIf(Trim(.ListItems(.SelectedItem.Index).SubItems(4)) = "Liquidada", True, False)
        'If Trim(.ListItems(.SelectedItem.Index).SubItems(3)) = "Liquidada" Then
        '      cbo_intEventoLiq.ListIndex = 0
        'End If
    
    End With
    VerificaTabAtivo
        mblnAtivarPastas = True
End Sub

Private Function strValorRetencao(bytTipo As Byte) As String

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strValorRetencao = 0
    strSQL = ""
'    strSql = strSql & "SELECT ISNULL(SUM(dblValor), 0) AS Soma FROM "
    strSQL = strSQL & "SELECT " & gstrISNULL("SUM(dblValor)", "0") & " AS Soma FROM "
    strSQL = strSQL & gstrSubempenhoLiquidado & " "
    strSQL = strSQL & "WHERE intParcela = " & lvw_Liquidacao.SelectedItem.Tag & " "
    strSQL = strSQL & "AND bytTipo = " & bytTipo
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                strValorRetencao = gstrConvVrDoSql(!SOMA)
            End If
        End With
    End If
End Function

Private Sub lvw_NotasFiscais_GotFocus()
   
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
   VerificaTabAtivo
   mAtivaPastaDeObjeto tab_3dPasta, 3, tab_3DPastaLiquidacao, 4
End Sub

Private Sub lvw_NotasFiscais_ItemClick(ByVal Item As MSComctlLib.ListItem)
   
   TrocaCorObjeto txt_dtmDataNF, True
   TrocaCorObjeto txt_dblValorNF, True
   TrocaCorObjeto txt_strNotasFiscais, True
   
   txt_dtmDataNF = lvw_NotasFiscais.SelectedItem
   txt_dblValorNF = lvw_NotasFiscais.SelectedItem.SubItems(1)
   txt_strNotasFiscais = lvw_NotasFiscais.SelectedItem.SubItems(2)
   mblnAlterandoNF = True
   VerificaTabAtivo
End Sub

Private Sub lvw_Retencao_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mlngPKIdRetencao = lvw_Retencao.SelectedItem.Tag
    mblnAlterandoRetencao = True
    With lvw_Retencao.ListItems(lvw_Retencao.SelectedItem.Index)
        txt_ValorRetencao = .SubItems(2)
        cbo_ContaRetencao.ListIndex = gintIndiceCBO(cbo_ContaRetencao, .SubItems(3))
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
    End With
End Sub

Private Sub tab_3DEmpenho_Click(PreviousTab As Integer)
   
   If tab_3DEmpenho.Tab = 0 Then
      On Error Resume Next
      If dbcintItemDespesa.Enabled Then dbcintItemDespesa.SetFocus
      HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
   ElseIf tab_3DEmpenho.Tab = 1 Then
      If txtstrContrato.Enabled Then txtstrContrato.SetFocus
      HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
   End If
   VerificaTabAtivo
End Sub

Private Sub tab_3DPastaLiquidacao_Click(PreviousTab As Integer)
    VerificaPastaLiquidacao tab_3DPastaLiquidacao.Tab
    'TrocaCorObjeto txt_ValorExtra, False
    VerificaTabAtivo
End Sub

Private Sub VerificaPastaLiquidacao(bytTab As Byte)
    Select Case bytTab
    Case 0
        'HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem
        'HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrExcluirItem
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCancelar
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrNovo
    Case 1
        'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
        'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
        VerificaContaExtra
        If Not lvw_Liquidacao.SelectedItem Is Nothing Then
            If Trim(lvw_Liquidacao.SelectedItem.ListSubItems(4)) = "Liquidada" Then
                LimpaDadosExtra True
            Else
                LimpaDadosExtra
            End If
        End If
        

        
    Case 2
        'HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem
        'HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrExcluirItem
        VerificaContaRetencao
    Case 3
        VerificaContaOrcamentario
        If Not lvw_Liquidacao.SelectedItem Is Nothing Then
            If Trim(lvw_Liquidacao.SelectedItem.ListSubItems(4)) = "Liquidada" Then
                LimpaDadosOrcamentario True
            Else
                LimpaDadosOrcamentario
            End If
        End If
    Case 4
       'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
       'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
       
       txt_dblValorNF = txt_dblValorAux
       txt_dtmDataNF = txt_DataLiuidacao
       txt_strNotasFiscais = "Sem Nota Fiscal"
       If txt_dtmDataNF.Enabled = True Then
       txt_dtmDataNF.SetFocus
       End If
    End Select
End Sub

Private Sub VerificaContaExtra()
    Static lngEmpenho  As Long
    Static lngParcela  As Long
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrNovo, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar
    If cbo_DescricaoExtra.ListCount = 0 Then
        LePlanoContaGeral cbo_ContaExtra, cbo_DescricaoExtra, "OU", "RT", "EO"
    End If
    If Not lvw_Liquidacao.SelectedItem Is Nothing Then
        If lngEmpenho <> Val(tdb_Lista.Columns(0).Value) Or lngParcela <> Val(lvw_Liquidacao.SelectedItem.Tag) Then
            lngEmpenho = Val(tdb_Lista.Columns(0))
            With lvw_Liquidacao
                lngParcela = Val(.SelectedItem.Tag)
            End With
            LeLiquidacaoExtra
        ElseIf gblnHaItemMarcadoLista(lvw_Extra) And lvw_Extra.ListItems.Count > 0 Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
        End If
    End If
End Sub

Private Sub VerificaContaOrcamentario()
    Static lngEmpenho  As Long
    Static lngParcela  As Long
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrNovo, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar
    If cbo_DescricaoOrcamentario.ListCount = 0 Then
        LePrevisaoReceitaGeral cbo_ContaOrcamentario, cbo_DescricaoOrcamentario
    End If
    If Not lvw_Liquidacao.SelectedItem Is Nothing Then
        If lngEmpenho <> Val(tdb_Lista.Columns(0).Value) Or lngParcela <> Val(lvw_Liquidacao.SelectedItem.Tag) Then
            lngEmpenho = Val(tdb_Lista.Columns(0))
            With lvw_Liquidacao
                lngParcela = Val(.SelectedItem.Tag)
            End With
            LeLiquidacaoOrcamentario
        ElseIf gblnHaItemMarcadoLista(lvw_Orcamentario) And lvw_Orcamentario.ListItems.Count > 0 Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
        End If
    End If
End Sub

Private Sub VerificaContaRetencao()
    Static lngEmpenho  As Long
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrNovo, gstrSalvar
    If cbo_DescricaoRetencao.ListCount = 0 Then
        LePrevisaoReceitaGeral cbo_ContaRetencao, cbo_DescricaoRetencao
    End If
    If lngEmpenho <> tdb_Lista.Columns(0) Or _
       mlngParcelaRetencao <> lvw_Liquidacao.SelectedItem.Tag Then
        lngEmpenho = tdb_Lista.Columns(0)
        mlngParcelaRetencao = lvw_Liquidacao.SelectedItem.Tag
        LeLiquidacaoRetencao
    ElseIf gblnHaItemMarcadoLista(lvw_Retencao) And lvw_Extra.ListItems.Count > 0 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
    End If
End Sub

Sub LeLiquidacaoRetencao()
    Dim objList         As Object
    Dim strSQL          As String
    Dim dblSoma         As Double
    Dim adoResultado    As ADODB.Recordset
    lvw_Retencao.ListItems.Clear
    strSQL = ""
    strSQL = strSQL & "SELECT CO.strCodigoOrcamentario, CO.strDescricao, "
    strSQL = strSQL & "PL.PKId, PL.dblValor, PL.intConta "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrCodigoOrcamentario & " CO, "
    strSQL = strSQL & gstrSubempenhoLiquidado & " PL "
    strSQL = strSQL & "WHERE PL.intConta = CO.PKId "
    strSQL = strSQL & "AND PL.bytTipo = 2 "
    strSQL = strSQL & "AND PL.intParcela = " & lvw_Liquidacao.SelectedItem.Tag & " "
    strSQL = strSQL & "ORDER BY CO.strCodigoOrcamentario"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set objList = lvw_Retencao.ListItems.Add(, , _
                              gvntFormatacaoEspecifica(!strCodigoOrcamentario))
                objList.SubItems(1) = !strDescricao
                objList.SubItems(2) = gstrConvVrDoSql(!dblValor)
                objList.SubItems(3) = !intConta
                dblSoma = dblSoma + gstrConvVrDoSql(!dblValor)
                objList.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
    lblRetencao = gstrConvVrDoSql(dblSoma)
End Sub

Private Sub LeLiquidacaoExtra()
    Dim objList         As Object
    Dim dblExtra        As Double
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    lvw_Extra.ListItems.Clear
    strSQL = ""
    strSQL = strSQL & "SELECT PC.strContaContabil, PC.strDescricao, "
    strSQL = strSQL & "PL.PKId, PL.dblValor, PL.intConta "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrSubempenhoLiquidado & " PL "
    strSQL = strSQL & "WHERE PL.intConta = PC.PKId "
    strSQL = strSQL & "AND PL.bytTipo = 1 "
    strSQL = strSQL & "AND PL.intParcela = " & lvw_Liquidacao.SelectedItem.Tag & " "
    strSQL = strSQL & "ORDER BY PC.strContaContabil"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set objList = lvw_Extra.ListItems.Add(, , _
                              gvntFormatacaoEspecifica(!strContaContabil))
                objList.SubItems(1) = !strDescricao
                objList.SubItems(2) = gstrConvVrDoSql(!dblValor)
                objList.SubItems(3) = !intConta
                dblExtra = dblExtra + gstrConvVrDoSql(!dblValor)
                objList.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
    lblExtra = gstrConvVrDoSql(dblExtra)
End Sub

Private Function GeraArraysSubElmento(ByVal strPikdEvento As String, ByVal mblnEstorno As Boolean, lwvObj As ListView) As Boolean
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim adoResultadoAux As ADODB.Recordset
    Dim intCaracsitem   As Integer
    Dim i               As Integer
    Dim intCont         As Integer
    Dim strCodConta     As String
    Dim strAux          As String
      
    ReDim aryContas(0)
    ReDim aryTpMov(0)
    ReDim aryValor(0)
   
   intCaracsitem = Len(Replace(gstrMascaraItemDespesa, ".", ""))
   intCont = 1
            
   strSQL = "SELECT "  'Contas de Credito
       strSQL = strSQL & "PC.strContaContabil, "
       strSQL = strSQL & "EVD.intContaContabil, 1 AS TpMov , "
       strSQL = strSQL & "PC.blnPatrimonial, "
       strSQL = strSQL & "PC.bytMovimentaSistema, "
       strSQL = strSQL & "'C' DebCred "
       
   strSQL = strSQL & "FROM "
       strSQL = strSQL & gstrEventoContaContabilCredito & " EVD, "
       strSQL = strSQL & gstrPlanoConta & " PC "
   strSQL = strSQL & " WHERE "
       strSQL = strSQL & "EVD.intEvento = " & strPikdEvento
       strSQL = strSQL & " AND EVD.intContaContabil = PC.Pkid"
       strSQL = strSQL & " AND EVD.bytContaGrupo = 1"


   strSQL = strSQL & " UNION ALL SELECT "  'Contas de Debito
       strSQL = strSQL & "PC.strContaContabil, "
       strSQL = strSQL & "EVD.intContaContabil, 1 AS TpMov , "
       strSQL = strSQL & "PC.blnPatrimonial, "
       strSQL = strSQL & "PC.bytMovimentaSistema, "
       strSQL = strSQL & "'D' DebCred "
       
   strSQL = strSQL & "FROM "
       strSQL = strSQL & gstrEventoContaContabilDebito & " EVD, "
       strSQL = strSQL & gstrPlanoConta & " PC "
   strSQL = strSQL & " WHERE "
       strSQL = strSQL & "EVD.intEvento = " & strPikdEvento
       strSQL = strSQL & " AND EVD.intContaContabil = PC.Pkid"
       strSQL = strSQL & " AND EVD.bytContaGrupo = 1"


    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
        
            If adoResultado.RecordCount > 0 Then
            ReDim aryContas(1 To adoResultado.RecordCount * lwvObj.ListItems.Count)
            ReDim aryTpMov(1 To adoResultado.RecordCount * lwvObj.ListItems.Count)
            ReDim aryValor(1 To adoResultado.RecordCount * lwvObj.ListItems.Count)
            End If
        
            Do While Not .EOF
                For i = 1 To lwvObj.ListItems.Count
                                   
                    strCodConta = Mid(Replace(gvntFormatacaoEspecifica(!strContaContabil, 1), ".", ""), Len(gstrDigitoDespesa) + 1)
                    
                    strCodConta = gstrDigitoDespesa & Replace(gvntFormatacaoEspecifica(strCodConta, 3), ".", "")
                    'strCodConta = gstrDigitoDespesa & strCodConta

                    strCodConta = Mid(strCodConta, 1, Len(strCodConta) - intCaracsitem) & Replace(lwvObj.ListItems(i).SubItems(1), ".", "")
                    
                    strAux = "SELECT PKID FROM " & gstrPlanoConta & " PC "
                    strAux = strAux & " WHERE "
                    strAux = strAux & " PC.strContaContabil LIKE '" & strCodConta & "%'"
                    
                    If gobjBanco.CriaADO(strAux, 5, adoResultadoAux) Then
                        If Not adoResultadoAux.EOF Then
                        aryContas(intCont) = Val(adoResultadoAux!Pkid)
                        aryTpMov(intCont) = IIf(adoResultado!debCred = "C", 0, 1) 'Crédito
                        
                        aryValor(intCont) = Replace(Str( _
                        IIf(mblnEstorno, _
                        CDbl(lwvObj.ListItems(i).SubItems(3)) * -1, _
                        CDbl(lwvObj.ListItems(i).SubItems(3))) _
                        ), ".", ",")
                        intCont = intCont + 1
                        Else
                            ExibeMensagem "Nem todos os Sub-Elementos formam uma Conta quando associados as Contas de Grupo do Evento selecionado."
                            Exit Function
                        End If
                    Else
                            ExibeMensagem "Ocorreu um erro ao verificar os sub-elementos."
                            Exit Function
                    End If

                Next
                .MoveNext
            Loop
        End With
    End If
    GeraArraysSubElmento = True

End Function

Private Sub LeSubElementos(lwvObj As ListView, Optional strPkidEmpenho As String, Optional strpkidParcela As String)
    Dim objList         As Object
    Dim dblExtra        As Double
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    lwvObj.ListItems.Clear
    strSQL = ""
    strSQL = strSQL & "SELECT ID.PKID ,"
    strSQL = strSQL & "ID.strCodigo, ID.strDescricao , SE.Dblvalor "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrSubElementoEmpenho & " SE, "
    strSQL = strSQL & gstrItemDespesa & " ID "
    strSQL = strSQL & "WHERE ID.pkid = SE.intItemDespesa "
    If strPkidEmpenho <> "" Then
        strSQL = strSQL & " AND SE.intEmpenho = " & strPkidEmpenho
    Else
        strSQL = strSQL & " AND SE.intparcela = " & strpkidParcela
    End If
    strSQL = strSQL & " ORDER BY strcodigo"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set objList = lwvObj.ListItems.Add(, , !Pkid)
                objList.SubItems(1) = gvntFormatacaoEspecifica(!strCodigo, 4)
                objList.SubItems(2) = !strDescricao
                objList.SubItems(3) = gstrConvVrDoSql(!dblValor)
                objList.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
End Sub

Private Sub LeLiquidacaoOrcamentario()
    Dim objList         As Object
    Dim dblOrcamentario As Double
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    lvw_Orcamentario.ListItems.Clear
    strSQL = ""
    strSQL = strSQL & "SELECT PC.strCodigoOrcamentario, PC.strDescricao, "
    strSQL = strSQL & "PL.PKId, PL.dblValor, PL.INTCODIGOORCAMENTARIO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrCodigoOrcamentario & " PC, "
    strSQL = strSQL & gstrSubEmpRetencaoOrcamentaria & " PL "
    strSQL = strSQL & "WHERE PL.INTCODIGOORCAMENTARIO = PC.PKId "
    strSQL = strSQL & "AND PL.intParcela = " & lvw_Liquidacao.SelectedItem.Tag & " "
    strSQL = strSQL & "ORDER BY PC.strCodigoOrcamentario"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                Set objList = lvw_Orcamentario.ListItems.Add(, , _
                              gvntFormatacaoEspecifica(!strCodigoOrcamentario))
                objList.SubItems(1) = !strDescricao
                objList.SubItems(2) = gstrConvVrDoSql(!dblValor)
                objList.SubItems(3) = !INTCODIGOORCAMENTARIO
                'dblOrcamentario = dblOrcamentario + gstrConvVrDoSql(!DBLVALOR)
                objList.Tag = !Pkid
                .MoveNext
            Loop
        End With
    End If
    'txt_dblDesconto = gstrConvVrDoSql(dblOrcamentario)
End Sub


Private Sub tab_3DPastaLiquidacao_LostFocus()
   HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem
   HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrExcluirItem
End Sub

Private Sub tdb_Lista_Click()
    mblnAtivarPastas = False
    'Alteração M4RC3LØ 1 LN DONW
    TrocaCorObjeto txtintExercicioEmpenho, True
    mblnPrimeiraVez = False
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        If blnDesbilita2Click = True Then
            tdb_Lista_RowColChange 0, 0
        Else
            blnDesbilita2Click = False
        End If
    End If
    mblnAtivarPastas = True
    mblnAbrindo = True
End Sub

Private Sub tdb_Lista_DblClick()
    
    'MantemForm gstrAplicar
    
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = True
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    mblnAtivarPastas = False
    If mblnPrimeiraVez Then
       mblnPrimeiraVez = False
       mblnAtivarPastas = False
       Exit Sub
    End If
    
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            mblnselecionou = True
            mblnAlterandoEmpenho = True
            mblnAtualizaTelaSubempenho = True
            blnDesbilita2Click = True
            gCorLinhaSelecionada tdb_Lista
            txtPKId = tdb_Lista.Columns(0).Value
            carregaEmpenho (tdb_Lista.Columns(0).Value)
            Habilitando_Modalidade
        End If
    End With
    mAtivaPastaDeObjeto tab_3dPasta, 0
    mblnAtivarPastas = True
End Sub

Private Sub carregaEmpenho(ByVal strEmpenhoPKID As String)
    LeEmpenho (strEmpenhoPKID)
    txtPKId.Text = strEmpenhoPKID
    PreencheGridItem (strEmpenhoPKID)
    DoEvents
    'If cbointReservaDotacao.ListCount = 1 Then cbointReservaDotacao.ListIndex = 0
    'If cboProgramaTrabalho.ListCount = 1 Then cboProgramaTrabalho.ListIndex = 0
    If blnEmpenhadoCompras Or mblnRestosAPagar Then
       'TrocaCorObjeto txtstrCodigo, True
        TrocaCorObjeto txtintExercicio, True
        TrocaCorObjeto txtbitDigito, True
        TrocaCorObjeto txt_intNContribuinte, True
        TrocaCorObjeto dbcintCredor, True
        TrocaCorObjeto cmd_Credor, True
        'TrocaCorObjeto dbcintModalidade, True
        'TrocaCorObjeto txtstrModalidade, True
        TrocaCorObjeto txtstrsolicitacao, True
        TrocaCorObjeto dbcintTipo, True
        TrocaCorObjeto cmd_Tipo, True
    Else
        TrocaCorObjeto txtstrCodigo, False
        TrocaCorObjeto txtintExercicio, False
        TrocaCorObjeto txtbitDigito, False
        TrocaCorObjeto txt_intNContribuinte, True
        TrocaCorObjeto dbcintCredor, True
        TrocaCorObjeto cmd_Credor, True
        'TrocaCorObjeto dbcintModalidade, True
        'TrocaCorObjeto txtstrModalidade, True
        TrocaCorObjeto txtstrsolicitacao, True
        TrocaCorObjeto dbcintTipo, True
        TrocaCorObjeto cmd_Tipo, True
    End If
    
    'Pen_687_ORC_100
    HabilitaControlesLiquidacao True
    
    VerificaSubempenho
    VerificaComplemento
    VerificaLiquidacao
    VerificaAnulacao
    LimpaCamposItem
    'M4R
    If mobjAux Is Nothing = False Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    End If
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcular
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    
    tab_3DPastaLiquidacao.Tab = 0
    tab_3DEmpenho.Tab = 0
    tab_3DGeral.Tab = 0
    tab_3DAnulacao.Tab = 0
    tab_3DAnulacao.TabEnabled(1) = False
    'Alteração M4RC3LØ 3 Ln Down
    
    txt_ValorProgramaTrabalho.Text = ""
    txt_SaldoDotacao.Text = ""
    txt_TotalDotado.Text = ""
    
    TrocaCorObjeto txtstrCodigo, False
    TrocaCorObjeto txtintExercicio, False
    TrocaCorObjeto txtbitDigito, False
    TrocaCorObjeto cboProgramaTrabalho, True
    TrocaCorObjeto cboCodigoReduzido, True
    
End Sub


Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub LimpaCamposItem()
    txt_PkidItem.Text = ""
    txt_intCodigo.Text = ""
    txt_intCatalogoMaterialServico.Text = ""
    dbc_intStrMarca.Text = ""
    dbc_intStrMarca.BoundText = ""
    txt_dblQuantidade.Text = ""
    txt_dblValorEstimado.Text = ""
    txt_intUnidadedeMedida.Text = ""
    txt_strObsItem.Text = ""
    txt_strdescricaodetalhada.Text = ""
End Sub

Private Function Habilitando_Modalidade() As Boolean
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
    
    strSQL = strSQL & "SELECT EP.PKID "
    strSQL = strSQL & "FROM " & gstrPedidoDeEmpenho & " EP ," & gstrEmpenhoContrato & " EC "
    strSQL = strSQL & "WHERE EP.pkid =" & txtPKId.Text
    strSQL = strSQL & " AND EP.PKID = EC.INTPEDIDOEMPENHO "
        
       Set gobjBanco = New clsBanco
       If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
            If Not adoResultado.EOF Then
            Habilitando_Modalidade = True
            Else
            Habilitando_Modalidade = False
            End If
        End If
            
End Function


Private Sub LeEmpenho(strPKId As String)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM " & gstrEmpenho & " "
    strSQL = strSQL & "WHERE PKId = " & strPKId
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txtintNumero = !INTNUMERO
                txtDTMDATA = gstrDataFormatada(!DTMDATA)
                
                
                LeTabelaProgramaTrabalho (CStr(!intProgramaTrabalho))
                cboProgramaTrabalho.ListIndex = gintIndiceCBO(cboProgramaTrabalho, !intProgramaTrabalho)
                   
                
                
                'If cbo_intEvento.ListCount = 0 Then
                '   preencheCboevento (CStr(gstrVerificaCampoNulo(!intEvento)))
                'End If
                
                TrocaCorObjeto cbointReservaDotacao, True
                TrocaCorObjeto cmd_Reserva, True
                If Not IsNull(!intReservaDotacao) Then
                   TrocaCorObjeto txtintNumero, True
                
                   cbo_intEvento.ListIndex = gintIndiceCBO(cbo_intEvento, gstrVerificaCampoNulo(!intEvento))
                   LeTabelaReservaDotacao (CStr(!intReservaDotacao))
                   cbointReservaDotacao.ListIndex = gintIndiceCBO(cbointReservaDotacao, !intReservaDotacao)
                   PreencheDadosReserva
                   txtdblValor = gstrConvVrDoSql(!dblValor)
                   txtstrContrato = gstrVerificaCampoNulo(!strContrato)
                   txtstrModalidade = gstrVerificaCampoNulo(!strModalidade)
                   If Habilitando_Modalidade = True Then
                            TrocaCorObjeto dbcintModalidade, True
                            TrocaCorObjeto txtstrModalidade, True
                    Else
                            TrocaCorObjeto dbcintModalidade, False
                            TrocaCorObjeto txtstrModalidade, False
                    End If
                   txtstrEmbasamento = gstrVerificaCampoNulo(!strEmbasamento)
                   txtdtmHomologacao = gstrDataFormatada(!dtmHomologacao)
                   txtstrLicitacao = gstrVerificaCampoNulo(!strLicitacao)
                   txt_intNContribuinte = LeCDCCredor(gstrVerificaCampoNulo(!intCredor))
                   txtstrCodigo = gstrVerificaCampoNulo(!strCodigo)
                   txtbitDigito = gstrVerificaCampoNulo(!bitDigito)
                   txtintExercicio = gstrVerificaCampoNulo(!intExercicio)
                   PreencherListaDeOpcoes dbcintTipo, gstrVerificaCampoNulo(!intTipo)
                   PreencherListaDeOpcoes dbcintConvenio, gstrVerificaCampoNulo(!intConvenio)
                   PreencherListaDeOpcoes dbcintModalidade, gstrVerificaCampoNulo(!intModalidade)
                   PreencherListaDeOpcoes dbcintFundo, gstrVerificaCampoNulo(!intFundo)
                   PreencherListaDeOpcoes dbcintItemDespesa, gstrVerificaCampoNulo(!intItemDespesa)
                   PreencherListaDeOpcoes dbcintCredor, gstrVerificaCampoNulo(!intCredor)
                   txtstrsolicitacao = gstrVerificaCampoNulo(!strSolicitacao)
                   txtstrHistorico = gstrVerificaCampoNulo(!STRHISTORICO)
                    'orc1376
                   txtintExercicioEmpenho = CInt(!intExercicioEmpenho)
                   txtStrcondpagto.Text = gstrVerificaCampoNulo(!Strcondpagto)
                   txtStrlocentrega.Text = gstrVerificaCampoNulo(!Strlocentrega)
                   txtStrprazoentrega.Text = gstrVerificaCampoNulo(!Strprazoentrega)
                   
                   LeSubElementos lvwSubElemento, gstrVerificaCampoNulo(!Pkid)
                   CarregaCombosSubElementosEst gstrVerificaCampoNulo(!Pkid)
                                   
                   TrocaCorObjeto txtdblValor, True
                   TrocaCorObjeto txtDTMDATA, True
                   TrocaCorObjeto cboCodigoReduzido, True
                   TrocaCorObjeto cboProgramaTrabalho, True
                   TrocaCorObjeto cmd_ProgramaTrabalho, True
                   TrocaCorObjeto cbo_intEvento, True
                   TrocaCorObjeto txt_codEvento, True
                   TrocaCorObjeto cmd_Evento, True
                   HabilitaDesabilitaTab tab_3dPasta, True
                Else
                   TrocaCorObjeto txtintNumero, True
                   
                   LimpaDadosReserva
                   TrocaCorObjeto cbointReservaDotacao, True
                   TrocaCorObjeto cmd_Reserva, True
                   LeTabelaProgramaTrabalho (CStr(!intProgramaTrabalho))
                   cboProgramaTrabalho.ListIndex = gintIndiceCBO(cboProgramaTrabalho, !intProgramaTrabalho)
                   cbo_intEvento.ListIndex = gintIndiceCBO(cbo_intEvento, gstrVerificaCampoNulo(!intEvento))
                   txtdblValor = gstrConvVrDoSql(!dblValor)
                   txtstrContrato = gstrVerificaCampoNulo(!strContrato)
                   txtstrModalidade = gstrVerificaCampoNulo(!strModalidade)
                   If Habilitando_Modalidade = True Then
                            TrocaCorObjeto dbcintModalidade, True
                            TrocaCorObjeto txtstrModalidade, True
                    Else
                            TrocaCorObjeto dbcintModalidade, False
                            TrocaCorObjeto txtstrModalidade, False
                    End If
                   txtstrEmbasamento = gstrVerificaCampoNulo(!strEmbasamento)
                   txtdtmHomologacao = gstrDataFormatada(!dtmHomologacao)
                   txtstrLicitacao = gstrVerificaCampoNulo(!strLicitacao)
                   txt_intNContribuinte = LeCDCCredor(gstrVerificaCampoNulo(!intCredor))
                   txtstrCodigo = gstrVerificaCampoNulo(!strCodigo)
                   txtbitDigito = gstrVerificaCampoNulo(!bitDigito)
                   txtintExercicio = gstrVerificaCampoNulo(!intExercicio)
                   PreencherListaDeOpcoes dbcintTipo, gstrVerificaCampoNulo(!intTipo)
                   PreencherListaDeOpcoes dbcintModalidade, gstrVerificaCampoNulo(!intModalidade)
                   PreencherListaDeOpcoes dbcintConvenio, gstrVerificaCampoNulo(!intConvenio)
                   PreencherListaDeOpcoes dbcintFundo, gstrVerificaCampoNulo(!intFundo)
                   PreencherListaDeOpcoes dbcintItemDespesa, gstrVerificaCampoNulo(!intItemDespesa)
                   PreencherListaDeOpcoes dbcintCredor, gstrVerificaCampoNulo(!intCredor)
                   txtstrsolicitacao = gstrVerificaCampoNulo(!strSolicitacao)
                   txtstrHistorico = gstrVerificaCampoNulo(!STRHISTORICO)
                   
                   LeSubElementos lvwSubElemento, gstrVerificaCampoNulo(!Pkid)
                   CarregaCombosSubElementosEst gstrVerificaCampoNulo(!Pkid)
                   txtintExercicioEmpenho = Year(CDate(!DTMDATA))
                   txtStrcondpagto.Text = gstrVerificaCampoNulo(!Strcondpagto)
                   txtStrlocentrega.Text = gstrVerificaCampoNulo(!Strlocentrega)
                   txtStrprazoentrega.Text = gstrVerificaCampoNulo(!Strprazoentrega)
                   
                   TrocaCorObjeto txtdblValor, True
                   TrocaCorObjeto txtDTMDATA, True
                   TrocaCorObjeto cboCodigoReduzido, True
                   TrocaCorObjeto cboProgramaTrabalho, True
                   TrocaCorObjeto cmd_ProgramaTrabalho, True
                   TrocaCorObjeto cbo_intEvento, True
                   TrocaCorObjeto txt_codEvento, True
                   TrocaCorObjeto cmd_Evento, True
                   HabilitaDesabilitaTab tab_3dPasta, True
               End If
            End If
        End With
    End If
End Sub

Private Sub lvw_ListaSubempenho_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_ListaSubempenho
        mblnAlterandoSubEmpenho = True
        
        TrocaCorObjeto txt_DataParcela, True
        TrocaCorObjeto txt_dtmVenctoLiqAutomatica, True
        TrocaCorObjeto txt_ValorParcela, True
        TrocaCorObjeto txt_strNotasFiscaisLiqAutomatica, True
        TrocaCorObjeto cmd_EventoLiqAutomatica, True
        TrocaCorObjeto cbo_intEventoLiqAutomatica, True
        TrocaCorObjeto txt_codEventoLiqAutomatica, True
        
        txt_DataParcela = .ListItems(.SelectedItem.Index).SubItems(1)
        txt_ValorParcela = .ListItems(.SelectedItem.Index).SubItems(2)
        txt_HistoricoSubEmpenho = .ListItems(.SelectedItem.Index).SubItems(5)
        'If Val(.ListItems(.SelectedItem.Index).SubItems(6)) = 1 Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
        'Else
        '    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar
        'End If
    End With
End Sub

Public Sub SelecionaLiquidacao()
       mblnClickOk = True
       DoEvents
       tdb_Lista_RowColChange 0, 0
       tab_3dPasta.Tab = 3
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
If mblnselecionou = False Then
    VeriticaTabClicado tab_3dPasta.Tab
End If
   If tab_3dPasta.Tab = 3 Then
        If cbo_intEventoLiq.ListCount = 1 Then
            cbo_intEventoLiq.ListIndex = 0
        End If
        
        If Not lvw_Liquidacao.SelectedItem Is Nothing Then
            lvw_Liquidacao_ItemClick lvw_Liquidacao.SelectedItem
        End If
        
    ElseIf tab_3dPasta.Tab = 1 Then
        If cbo_intEventoLiqAutomatica.ListCount = 1 Then
            cbo_intEventoLiqAutomatica.ListIndex = 0
        End If
    End If
    PreencheSaldoEmpenho
    VerificaTabAtivo
End Sub


Private Sub PreencheSaldoEmpenho()

Dim strSQL As String
Dim adoResultado As ADODB.Recordset


    If mblnAlterandoEmpenho = True Then
        If lvw_ListaSubempenho.ListItems.Count >= 1 Then
            If lvw_ListaSubempenho.ListItems(1).ListSubItems.Count >= 2 Then
                txtdblSaldoEmpenho = lvw_ListaSubempenho.ListItems(1).SubItems(2)
            End If
        End If
       
        strSQL = "select bytsituacao "
        strSQL = strSQL & " From " & gstrSubempenho
        strSQL = strSQL & " Where INTEMPENHO = " & txtPKId.Text
        strSQL = strSQL & " and intnumero=0 "
    
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        
            If Not adoResultado.EOF Then
                If adoResultado("bytsituacao") = 3 Then
                    txtdblSaldoEmpenho = "0,00"
                End If
            End If
        End If
        
        strSQL = "SELECT bytSituacao, intNumero"
        strSQL = strSQL & " FROM " & gstrSubempenho
        strSQL = strSQL & " WHERE intEmpenho  = " & txtPKId.Text
        'strSQL = strSQL & " AND intNumero = 0 "
        strSQL = strSQL & " ORDER BY intNumero, bytSituacao"
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
            If Not adoResultado.EOF Then
                adoResultado.MoveFirst
                If adoResultado.RecordCount = 1 Then
                    If Val(adoResultado.Fields("bytsituacao")) = 2 And Val(adoResultado.Fields("intNumero")) = 0 Then
                        txtdblSaldoEmpenho = "0,00"
                    End If
                End If
            End If
        End If
    Else
        txtdblSaldoEmpenho = ""
    End If
End Sub

Sub VeriticaTabClicado(bytTab As Byte)
    If bytTab = 0 Then
        TrocaCorObjeto txtintNumero, mblnAlterandoEmpenho
        TrocaCorObjeto txtdblValor, mblnAlterandoEmpenho
        TrocaCorObjeto txtDTMDATA, mblnAlterandoEmpenho
        TrocaCorObjeto cboCodigoReduzido, mblnAlterandoEmpenho
        TrocaCorObjeto cboProgramaTrabalho, mblnAlterandoEmpenho
        TrocaCorObjeto cmd_ProgramaTrabalho, mblnAlterandoEmpenho
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar, gstrCancelar
        
    Else
        TrocaCorObjeto txtintNumero, True
        TrocaCorObjeto txtdblValor, True
        TrocaCorObjeto txtDTMDATA, True
        TrocaCorObjeto cboCodigoReduzido, True
        TrocaCorObjeto cboProgramaTrabalho, True
        TrocaCorObjeto cmd_ProgramaTrabalho, True
    End If
    
    Select Case bytTab
    Case 0
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrNovo, gstrSalvar
    Case 1
        VerificaSubempenho
    Case 2
        VerificaComplemento
    Case 3
        'Pen_687_ORC_100
        If tab_3dPasta.TabEnabled(2) = True Then
            VerificaLiquidacao
        End If
    Case 4
        VerificaAnulacao
    End Select
End Sub

Sub VerificaAnulacao()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrNovo
    If (txt_DataAnulucao.Enabled) And (txt_DataAnulucao.Visible) Then txt_DataAnulucao.SetFocus
    LeSubEmpenho lvw_Anulacao, 4, txtPKId
    If cbo_HistoricoAnulacao.ListCount = 0 Then
        LeDaTabelaParaObj "", cbo_HistoricoAnulacao, "Select strcodigo, strDescricao FROM " & gstrHistorico & " ORDER BY strDescricao"
        TrocaCorObjeto txt_ValorAnulacao, True
    End If
    If mlngEmpenhoAnulacao <> txtPKId Or _
       lvw_Anulacao.ListItems.Count = 0 Then
        mlngEmpenhoAnulacao = txtPKId
        txt_DataAnulucao = ""
        txt_ValorAnulacao = ""
        txt_HistoricoAnulacao = ""
    End If
    With lvw_Anulacao
        If .ListItems.Count > 0 Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar
        End If
    End With
End Sub

Private Function strQueryPeriodo() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao FROM "
    strSQL = strSQL & gstrPeriodo & " "
    strSQL = strSQL & "WHERE bytPeriodo <> 3"
    strQueryPeriodo = strSQL
End Function



Private Sub ProcuraPeriodoMensal()
    Dim bytInd  As Byte
    With cboPeriodo
        If .ListCount > 0 Then
            For bytInd = 0 To .ListCount - 1
                cboPeriodo.ListIndex = bytInd
                If UCase(Trim(cboPeriodo.Text)) = "MENSAL" Then
                    Exit Sub
                End If
            Next
        End If
    End With
End Sub

Private Sub VerificaSubempenho()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrNovo
    If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
    If cbo_HistoricoSubEmpenho.ListCount = 0 Then
        LeDaTabelaParaObj "", cbo_HistoricoSubEmpenho, "Select strcodigo, strDescricao FROM " & gstrHistorico & " ORDER BY strDescricao"
        LeDaTabelaParaObj "", cboPeriodo, strQueryPeriodo
    End If
    If mblnAtualizaTelaSubempenho Then
        mblnAtualizaTelaSubempenho = False
        LeSubEmpenho lvw_ListaSubempenho, , txtPKId
        PreencheSaldoEmpenho
        LimpaTelaSubempenho
    End If
    With lvw_ListaSubempenho
         If .ListItems.Count = 1 Then
            If .ListItems(1).Text = 1 Then
                TrocaCorObjeto txtNumParcela, False
                TrocaCorObjeto cboPeriodo, False
                TrocaCorObjeto cmd_Periodo, False
                ProcuraPeriodoMensal
            End If
        Else
            TrocaCorObjeto txtNumParcela, True
            TrocaCorObjeto cboPeriodo, True
            TrocaCorObjeto cmd_Periodo, True
        End If
        If gblnHaItemMarcadoLista(lvw_ListaSubempenho) And .ListItems.Count > 1 Then
            If Val(.ListItems(.SelectedItem.Index).SubItems(6)) = 1 Then
                HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrNovo, gstrSalvar
            Else
                HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrNovo, gstrSalvar
            End If
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo
        End If
    End With
    
    HabilitaDesabilitaBotao1 False, gstrDeletar
    
End Sub

Private Sub VerificaLiquidacao()

    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrNovo
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
    
    TrocaCorObjeto cbo_intEventoLiq, False
    TrocaCorObjeto txt_codEventoLiq, False
    txt_codEventoLiq.Enabled = True
    txt_codEventoLiq.BackColor = vbWindowBackground
    cbo_intEventoLiq.ListIndex = -1
    
    If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
    
    LeSubEmpenho lvw_Liquidacao, 2, txtPKId
    LeSubEmpenho lvw_Anulacao, 4, txtPKId
    LeSubEmpenho lvw_ListaSubempenho, , txtPKId
    
    If cbo_HistoricoLiquidacao.ListCount = 0 Then
        LeDaTabelaParaObj "", cbo_HistoricoLiquidacao, "Select strcodigo, strDescricao FROM " & gstrHistorico & " ORDER BY strDescricao"
    End If
    If mlngEmpenhoLiquidacao <> txtPKId Or lvw_Liquidacao.ListItems.Count = 0 Then
        mlngEmpenhoLiquidacao = txtPKId
        LimpaDadosLiquidacao
    End If
    With lvw_Liquidacao
        If gblnHaItemMarcadoLista(lvw_Liquidacao) And .ListItems.Count > 1 Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
        End If
        If .ListItems.Count > 0 Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar
        End If
    End With
End Sub

Sub AtualizaTotalComplemento()
    Dim intInd  As Integer
    Dim dblSoma As Double
    With lvw_Complemento
        For intInd = 1 To .ListItems.Count
            dblSoma = dblSoma + .ListItems(intInd).SubItems(2)
        Next
    End With
    lblTotalComplemento = gstrConvVrDoSql(dblSoma)
End Sub

Private Sub VerificaComplemento()
    Static lngEmpenho   As Long
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrNovo
    If cbo_HistoricoComplemento.ListCount = 0 Then
        LeDaTabelaParaObj "", cbo_HistoricoComplemento, "Select strcodigo, strDescricao FROM " & gstrHistorico & " ORDER BY strDescricao"
    End If
    'If lngEmpenho <> tdb_Lista.Columns(0).Text Then
        lngEmpenho = txtPKId 'tdb_Lista.Columns(0).Text
        LeSubEmpenho lvw_Complemento, 5, txtPKId
        LimpaTelaComplemento
        AtualizaTotalComplemento
    'End If
    With lvw_Complemento
        If .ListItems.Count = 1 Then
            If .ListItems(1).Text = 1 Then
                TrocaCorObjeto txtNumParcela, False
                TrocaCorObjeto cboPeriodo, False
                TrocaCorObjeto cmd_Periodo, False
                ProcuraPeriodoMensal
            End If
        End If
        If gblnHaItemMarcadoLista(lvw_Complemento) And .ListItems.Count > 1 Then
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrDeletar
        Else
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
        End If
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
    End With
End Sub

Private Function strQuerySubempenho(bytFlag As Byte, strPKId As String) As String
    Dim strSQL      As String
    
strSQL = ""

If bytFlag = 2 Then
    
    strSQL = "select SE.PKID,"
    
        
    strSQL = strSQL _
        & " se.intnumero, " _
        & " SE.DTMDATA, " _
        & " SE.dtmvencimento, " _
        & " SE.DBLVALOR, "
        
   
    If bytDBType = SQLServer Then
        strSQL = strSQL _
            & " (CASE bytSituacao " _
            & " WHEN 1 THEN 'Programada' " _
            & " WHEN 2 THEN  'Liquidada' " _
            & " WHEN 3 THEN 'Paga' " _
            & " WHEN 4 THEN " _
            & "     Case SE.INTNUMERO " _
            & "         WHEN 0 THEN " _
            & "             Case isnull(se.intevento, 0) " _
            & "                WHEN 0 THEN 'Anulada' " _
            & "                Else " _
            & "                     'Canc. RP' " _
            & "             End " _
            & "         Else " _
            & "             'Anulada' " _
            & "         End " _
            & " END )  AS strSituacao, "
    ElseIf bytDBType = Oracle Then
        strSQL = strSQL _
            & " DECODE(bytSituacao, 1, 'Programada', 2, 'Liquidada', " _
            & " 3, 'Paga', 4, DECODE(se.intnumero,0, " _
            & " DECODE(SE.intEvento, NULL,'Anulada','Canc. RP'), " _
            & " 'Anulada'))  AS strSituacao, "
    End If
    
    strSQL = strSQL _
        & " (select OP.INTNUMERO from " & gstrOrdemPagamento & " OP , " _
        & gstrOrdemPagamentoEmpenho & " OPE where " _
        & " OP.PKID " & strOUTJOracle & " = ope.intordempagamento " _
        & " AND OPE.INTPARCELA = Se.Pkid " _
        & " AND " & gstrISNULL("OP.bytCancelado", "0") & "  = 0 " _
        & " Union " _
        & " select OP.INTNUMERO from " & gstrOrdemPagamento & " OP , " _
        & gstrOrdemPagamentoResto & " OPE where " _
        & " OP.PKID " & strOUTJOracle & " = ope.intordempagamento " _
        & " AND OPE.INTPARCELA = Se.Pkid " _
        & " AND " & gstrISNULL("OP.bytCancelado", "0") & " = 0 " _
        & " ) ORDEM, "
  
    strSQL = strSQL _
        & " (select dblValor  from " & gstrOrdemPagamento & " OP , " _
        & gstrOrdemPagamentoEmpenho & " OPE where " _
        & " OP.PKID " & strOUTJOracle & " = ope.intordempagamento " _
        & " AND OPE.INTPARCELA = Se.Pkid " _
        & " AND  " & gstrISNULL("OP.bytCancelado", "0") & " = 0 " _
        & " Union " _
        & " select OPE.DblValor from " & gstrOrdemPagamento & " OP , " _
        & gstrOrdemPagamentoResto & " OPE where " _
        & " OP.PKID " & strOUTJOracle & " = ope.intordempagamento " _
        & " AND OPE.INTPARCELA = Se.Pkid " _
        & " AND " & gstrISNULL("OP.bytCancelado", "0") & " = 0 " _
        & "  ) dblTotalOp, "
    
    
    
    strSQL = strSQL _
        & " dtmPagamento, ("
    
    strSQL = strSQL & gstrCASEWHEN("SE.bytTipo", "1,'Normal',2,'Complemento',3,'Anul. Desp.'")
    
    strSQL = strSQL & ") As strTipo, "
    
    'strSql = strSql _
        & " ( " _
        & " Case SE.bytTipo " _
        & "     WHEN 1 THEN 'Normal' " _
        & "     WHEN 2 THEN 'Complemento' " _
        & "     WHEN 3 THEN 'Anul. Desp.' " _
        & " End " _
        & " )  AS strTipo, "
   
    strSQL = strSQL _
        & " dtmLiquidacao, " _
        & " SE.strHistorico, " _
        & " SE.bytSituacao, " _
        & gstrISNULL("SE.intEvento", "0") & " As intEvento "
    
    strSQL = strSQL _
        & " from " & gstrEmpenho & " e, " & gstrSubempenho & " se " _
        & " Where e.Pkid = SE.INTEMPENHO " _
        & " and Se.intempenho = " & strPKId _
        & "  AND bytSituacao < 4 ORDER BY se.intNumero " _
        & " , se.dtmData, bytSituacao "

    strQuerySubempenho = strSQL

Else
    

    strSQL = ""
    strSQL = strSQL & "SELECT SU.PKId, SU.intNumero, SU.dtmData, SU.dblValor,"

    If bytDBType = Oracle Then
        strSQL = strSQL & "DECODE(SU.bytSituacao, 1, 'Programada', 2, 'Liquidada', 3, 'Paga', 4, DECODE(SU.intNumero,0, DECODE(SU.intEvento, NULL,'Anulada','Canc. RP'),'Anulada'))  AS strSituacao, "
    Else
     '   strSql = strSql & gstrCASEWHEN("SU.bytSituacao", "1, 'Programada', 2, 'Liquidada', 3, 'Paga', 4, 'Anulada'") & " AS strSituacao, "
        strSQL = strSQL _
            & " (CASE bytSituacao " _
            & " WHEN 1 THEN 'Programada' " _
            & " WHEN 2 THEN  'Liquidada' " _
            & " WHEN 3 THEN 'Paga' " _
            & " WHEN 4 THEN " _
            & "     Case SU.INTNUMERO " _
            & "         WHEN 0 THEN " _
            & "             Case isnull(su.intEvento, 0) " _
            & "                WHEN 0 THEN 'Anulada' " _
            & "                Else " _
            & "                    'Canc. RP'  " _
            & "             End " _
            & "         Else " _
            & "             'Anulada' " _
            & "         End " _
            & " END )  AS strSituacao, "
        
    End If

    strSQL = strSQL & gstrCASEWHEN("SU.bytTipo", _
        "1, 'Normal', 2, 'Complemento',3,'Anul. Desp.'") & " AS strTipo, "

    If bytFlag = 2 Then 'Liquidada
        strSQL = strSQL & "SU.dtmLiquidacao, "
    ElseIf bytFlag = 4 Then 'Anulada
        strSQL = strSQL & "SU.dtmAnulacao, "
    End If

    strSQL = strSQL & "SU.strHistorico, "
    strSQL = strSQL & "SU.bytSituacao," & gstrISNULL("SU.intEvento", "0", "SU.intEvento") & " AS intEvento "

    If bytFlag = 4 Then strSQL = strSQL & ", " & gstrISNULL("SU.intEmpenhoAnulacao", "''", "SU.intEmpenhoAnulacao") & " AS intEmpenhoAnulacao "
    strSQL = strSQL & " FROM " & gstrSubempenho & " SU, " & gstrEmpenho & " EP, " & gstrTipoEmpenho & " TP "
    strSQL = strSQL & "WHERE SU.intEmpenho = " & strPKId
    strSQL = strSQL & " AND EP.Pkid = SU.intEmpenho AND EP.Inttipo = TP.Pkid "
    If bytFlag = 2 Then
        strSQL = strSQL & " AND SU.bytSituacao < 4" 'Programada, Liquidada e paga
    ElseIf bytFlag = 4 Then
    
        If blnVerificaParamAnul Then
            strSQL = strSQL & " AND SU.bytSituacao IN (3) AND "
        Else
            strSQL = strSQL & " AND SU.bytSituacao IN (3) AND EP.Pkid = SU.intEmpenho AND EP.Inttipo = TP.Pkid AND TP.Bytadiantamento = 1 AND "
        End If
        
        
        If bytDBType = Oracle Then
            strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "EP.dtmData, 'yyyy'")
        ElseIf bytDBType = SQLServer Then
            strSQL = strSQL & " YEAR(EP.dtmData) "
        End If
        strSQL = strSQL & " = " & CStr(gintExercicio)
        
    ElseIf bytFlag = 5 Then
        strSQL = strSQL & " AND SU.bytTipo = 2" 'Liquidada
    End If
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "SELECT PKId, intNumero, dtmData, dblValor,"

    If bytDBType = Oracle Then
        strSQL = strSQL & "DECODE(bytSituacao, 1, 'Programada', 2, 'Liquidada', 3, 'Paga', 4, DECODE(intnumero,0, DECODE(intEvento, null,'Anulada','Canc. RP'),'Anulada'))  AS strSituacao, "
    Else
    'strSql = strSql & gstrCASEWHEN("bytSituacao", "1, 'Programada', 2, 'Liquidada', 3, 'Paga', 4, 'Anulada'") & " AS strSituacao, "
         strSQL = strSQL _
             & " (CASE bytSituacao " _
             & " WHEN 1 THEN 'Programada' " _
             & " WHEN 2 THEN  'Liquidada' " _
             & " WHEN 3 THEN 'Paga' " _
             & " WHEN 4 THEN " _
             & "     Case INTNUMERO " _
             & "         WHEN 0 THEN " _
             & "             Case isnull(INTEVENTO, 0) " _
             & "                WHEN 0 THEN 'Anulada' " _
             & "                Else " _
             & "                     'Canc. RP' " _
             & "             End " _
             & "         Else " _
             & "             'Anulada' " _
             & "         End " _
             & " END )  AS strSituacao, "
    End If
    
    strSQL = strSQL & gstrCASEWHEN("bytTipo", _
        "1, 'Normal', 2, 'Complemento',3,'Anul. Desp.'") & " AS strTipo, "

    If bytFlag = 2 Then 'Liquidada
        strSQL = strSQL & "dtmLiquidacao, "
    ElseIf bytFlag = 4 Then 'Anulada
        strSQL = strSQL & "dtmAnulacao, "
    End If

    strSQL = strSQL & "strHistorico, "
    strSQL = strSQL & "bytSituacao," & gstrISNULL("intEvento", "0", "intEvento") & " AS intEvento "

    If bytFlag = 4 Then strSQL = strSQL & ", " & gstrISNULL("intEmpenhoAnulacao", "''", "intEmpenhoAnulacao") & " AS intEmpenhoAnulacao "
    strSQL = strSQL & " FROM " & gstrSubempenho & " "
    strSQL = strSQL & gstrEmpenho & " "
    strSQL = strSQL & "WHERE intEmpenho = " & strPKId
    If bytFlag = 2 Then
        strSQL = strSQL & " AND bytSituacao < 4" 'Programada, Liquidada e paga
    ElseIf bytFlag = 4 Then
        strSQL = strSQL & " AND bytSituacao IN (1,4)" 'Programada e cancelada
    ElseIf bytFlag = 5 Then
        strSQL = strSQL & " AND bytTipo = 2" 'Liquidada
    End If


    strSQL = strSQL & " ORDER BY "
    If bytDBType = Oracle Then
        strSQL = strSQL & " intNumero, bytsituacao, dtmData"
    ElseIf bytDBType = SQLServer Then
        strSQL = strSQL & " SU.intNumero, bytsituacao, SU.dtmData"
    End If
'    strSql = strSql & ", bytSituacao "
    strQuerySubempenho = strSQL
    
End If
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Select Case UCase(strModoOperacao)
        Case gstrNovo
            VerificaTabNovo
        Case gstrSalvar
            VerificaTabGravacao
        Case gstrDeletar
            VerificaTabExclusao
        Case gstrImprimir
            If tab_3dPasta.Tab = 4 Then
                If Not lvw_Anulacao.SelectedItem Is Nothing Then
                   If UCase(lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).SubItems(3)) = "CANC. RP" Then
                      ImprimeCancelamentoRP (lvw_Anulacao.ListItems(lvw_Anulacao.SelectedItem.Index).Tag)
                   End If
                End If
            ElseIf Len(Trim(strEmpInicial)) = 0 Then
               If blnAtivaFormImprime Then
                  frm_IntervaloDeEmpenho.MantemForm (gstrImprimir)
               Else
                  CarregaForm frm_IntervaloDeEmpenho
                  frm_IntervaloDeEmpenho.txt_EmpInicial = txtintNumero
                  frm_IntervaloDeEmpenho.txt_EmpFinal = txtintNumero
                  frm_IntervaloDeEmpenho.txt_ParcInicial = 0
                  frm_IntervaloDeEmpenho.txt_ParcFinal = 0
               End If
               
            Else
               ImprimeRelatorio rptNotaDeEmpenho, strQueryRelatorio
               strEmpInicial = Space$(0)
               strEmpFinal = Space$(0)
               strParcInicial = Space$(0)
               strParcFinal = Space$(0)
               blnSoEstorno = False
            End If
        
        Case UCase(gstrCalcular)
            CalculaParcela
        Case UCase(gstrImportarDados)
           CarregaForm frmImportEmpenho
        Case UCase(gstrCancelar)
            If Not VerificaOrdemDePagamento Then
               CancelaLiquidacao
            Else
               ExibeMensagem "Esta liquidação possui Ordem de Pagamento e sua liquidação não poderá ser desfeita."
            End If
        Case UCase(gstrIncluirItem)
            VerificaTabParaIncluir
        Case UCase(gstrAtualizar)
            Form_Activate
        Case UCase(gstrExcluirItem)
            VerificaTabParaExcluir
            CalculaSubTotalItem (gstrExcluirItem)
        Case UCase(gstrPreencherLista)
             
            If Me.ActiveControl.Name = cboCodigoReduzido.Name Or Me.ActiveControl.Name = cboProgramaTrabalho.Name Then
               If Left(cboProgramaTrabalho, 1) = "%" And Len(Trim(cboProgramaTrabalho)) > 1 Then
                  LeTabelaProgramaTrabalho , cboProgramaTrabalho.Text
               Else
                  LeTabelaProgramaTrabalho
               End If
                
            End If
            If Me.ActiveControl.Name = cbointReservaDotacao.Name Then
               LeTabelaReservaDotacao
            End If
            
            If Me.ActiveControl.Name = cbo_intEvento.Name Then
                preencheCboevento
            End If
            
            If Me.ActiveControl.Name = cbo_intEventoLiq.Name Then
                preencheCboeventoLiq
            End If
            
            If Me.ActiveControl.Name = cbo_intEventoAnul.Name Then
                LeDaTabelaParaObj gstrEvento, cbo_intEventoAnul, strQueryAplicarEventoAnul
            End If
            
            If Me.ActiveControl.Name = cbo_intEventoLiqAutomatica.Name Then
                preencheCboeventoLiqAutomatica
            End If
            
         Case gstrFechar
            Unload Me
    End Select
    If UCase(strModoOperacao) = UCase(gstrLocalizar) Or UCase(strModoOperacao) = UCase(gstrPreencherLista) Then
     ToolBarGeral strModoOperacao, gstrEmpenho, mblnAlterandoEmpenho, tdb_Lista, Me, mobjAux, strQueryLocalizar, strQueryLocalizar
     
     
    End If
    If UCase(strModoOperacao) = UCase(gstrAplicar) Then
        
        If dbcintCredor.BoundText = frmCadOrdemPagamento.dbcintCredor.BoundText Or frmCadOrdemPagamento.dbcintCredor.BoundText = "" Then 'Verifica o Credor
            If lvw_Liquidacao.SelectedItem.ListSubItems.Item(4).Text = "Programada" Then 'Verifica situção da Parcela
                ExibeMensagem "Somente parcelas Liquidadas podem ser selecionadas para uma Ordem de Pagamento."
            Else
                If blnTestaLiquidacao Then
                    frmCadOrdemPagamento.mblnTelaEmpenho = True
                    frmCadOrdemPagamento.strpkidParcela = lvw_Liquidacao.SelectedItem.Tag
                    ToolBarGeral strModoOperacao, gstrEmpenho, mblnAlterandoEmpenho, tdb_Lista, Me, mobjAux, strQueryLocalizar, strQueryLocalizar
                Else
                    ExibeMensagem "Esta parcela já pertence a uma ordem de pagamento."
                End If
            End If
        Else
            ExibeMensagem "Este empenho não pertence ao credor já selecionado na Ordem de Pagamento corrente."
        End If
    End If
    
    If UCase(strModoOperacao) = gstrLocalizar Then
       mblnLimpaGrid = True
    End If
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
    If strModoOperacao = gstrNovo Then
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
    End If
End Sub

Private Sub CancelaLiquidacao()
    Dim strSQL      As String
    Dim strMsg      As String
    Dim strParcela  As String
    Dim strDtmliquidacao  As String
    
    strParcela = lvw_Liquidacao.SelectedItem.Tag
       
    strMsg = "Confirma cancelamento da liquidação?"
    If gblnExclusaoGravacaoOk("", strMsg, True) Then
    
        strDtmliquidacao = CStr(lvw_Liquidacao.ListItems(lvw_Liquidacao.SelectedItem.Index).SubItems(9))
        'If Year(CDate(strDtmliquidacao)) <> CInt(gintExercicio) Then
        strDtmliquidacao = frmDataPrompt.DataPrompt("Insira a data do cancelamento.", CDate(strDtmliquidacao), "EO", "a Parcela", gintExercicio)
        'End If
        
        
        'ORC677
        If IsDate(strDtmliquidacao) Then
            If Year(CDate(strDtmliquidacao)) <> CInt(gintExercicio) Then
                ExibeMensagem "A data de liquidação tem que estar no exercício de " & gintExercicio & "."
                Exit Sub
            End If
        End If
        
        
        If strDtmliquidacao = "" Then Exit Sub
           
    
        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
        
        strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
        strSQL = strSQL & "dtmLiquidacao = NULL, "
        strSQL = strSQL & "bytSituacao = 1, " '1 = Programada
        strSQL = strSQL & "dbldesconto = 0, " '1 = Programada
        strSQL = strSQL & "dtmVencimento = NULL, "
        strSQL = strSQL & "dtmDtAtualizacao = "
        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
        strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
        strSQL = strSQL & "WHERE PKId = " & lvw_Liquidacao.SelectedItem.Tag & "; "
'--------Deleta tabela de parcela liquidada
        strSQL = strSQL & "DELETE " & gstrSubempenhoLiquidado & " "
        strSQL = strSQL & "WHERE intParcela = " & lvw_Liquidacao.SelectedItem.Tag & "; "
'-------Deleta tabela de Desconto Orçamentario
        strSQL = strSQL & "DELETE " & gstrSubEmpRetencaoOrcamentaria & " "
        strSQL = strSQL & "WHERE intParcela = " & lvw_Liquidacao.SelectedItem.Tag & "; "
'-------Deleta as notas fiscais pertencentes a esta liquidação
        strSQL = strSQL & "DELETE " & gstrSubEmpenhoNF & " "
        strSQL = strSQL & "WHERE intSubEmpenho = " & lvw_Liquidacao.SelectedItem.Tag
        
        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; END; ", "")
        
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSQL) Then
            
            If Not GeraMovimentosByEvento(gstrItemData(cbo_intEventoLiq), strDtmliquidacao, Str(CDbl(txt_dblValorAux) * (-1)), txt_HistoricoLiquidacao, txtintNumero, "3") Then
               ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
            End If
             
            If Not blnGravaMovLiq(Val(txtPKId), lvw_Liquidacao.SelectedItem.Tag, gstrItemData(cboProgramaTrabalho, True), strDtmliquidacao, "-" & txt_dblValorAux, "***Cancelamento De Liquidação***") Then
               ExibeMensagem "Não foi possível completar o cancelamento desta Liquidação para este empenho."
            End If
             
            tab_3DPastaLiquidacao.TabEnabled(4) = False
            txt_dtmDataNF.Visible = False
            txt_dblValorNF.Visible = False
            txt_strNotasFiscais.Visible = False
            lbl_ValorTotal.Visible = False
            lvw_NotasFiscais.Visible = False
            
            LeSubEmpenho lvw_Anulacao, 4, txtPKId
            LeSubEmpenho lvw_Liquidacao, 2, txtPKId
            LeSubEmpenho lvw_ListaSubempenho, , txtPKId
            PreencheSaldoEmpenho
            Call gblnEncontroItemNoListView(lvw_Liquidacao, strParcela, lvwTag)
            lblExtra = 0
            lblRetencao = 0
            If tab_3DPastaLiquidacao.Tab = 1 Then
                tab_3DPastaLiquidacao.Tab = 0
            End If
            lvw_Liquidacao.ListItems(1).Selected = True
            cbo_DescricaoExtra.ListIndex = -1
            cbo_ContaExtra.ListIndex = -1
            txt_ValorExtra.Text = ""
            lvw_Extra.ListItems.Clear
            If Not lvw_Liquidacao.SelectedItem Is Nothing Then
                lvw_Liquidacao_ItemClick lvw_Liquidacao.SelectedItem
            Else
                lblRetencao = "0,00"
                lblExtra = "0,00"
                lblLiquido = "0,00"
                txt_dblDesconto = "0,00"
                txt_dblValorAux = "0,00"
                txt_DataLiuidacao = ""
            End If
        End If
    End If
End Sub

Private Function blnAtualizouParcela(lngChave As Long, _
                                     vntDataParcela, _
                                     vntValorParcela, _
                                     vntHistoricoSubEmpenho, _
                                     Optional intNumParc As Integer) As Boolean
    Dim strSQL      As String
    strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
    'strSql = strSql & "dtmData = " & gstrConvDtParaSql(vntDataParcela) & ", "
    strSQL = strSQL & "dblValor = " & gstrConvVrParaSql(vntValorParcela) & ", "
    'strSQL = strSQL & "strHistorico = '" & vntHistoricoSubEmpenho & "', "
    strSQL = strSQL & "dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
    strSQL = strSQL & "WHERE PKId = " & lngChave
    
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSQL) Then
        blnAtualizouParcela = True
    End If
End Function

Sub VerificaPeriodo(strPeriodo As String, intIntervalo As Integer)

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT intIntervalo, "
'    strSql = strSql & "CASE bytPeriodo "
'    strSql = strSql & "WHEN 0 THEN 'm' "
'    strSql = strSql & "WHEN 1 THEN 'd' "
'    strSql = strSql & "WHEN 2 THEN 'yyyy' "
'    strSql = strSql & "WHEN 3 THEN 'h' "
'    strSql = strSql & "END AS strPeriodo "
    strSQL = strSQL & gstrCASEWHEN("bytPeriodo", _
        "0, 'm', 1, 'd', 2, 'yyyy', 3, 'h'") & " AS strPeriodo "
    
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPeriodo & " "
    strSQL = strSQL & "WHERE PKId = " & gstrItemData(cboPeriodo)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                strPeriodo = gstrENulo(!strPeriodo)
                intIntervalo = gstrENulo(!intIntervalo)
            End If
        End With
    End If
End Sub

Sub CalculaParcela()
    Dim strMsg          As String
    Dim intInd          As Integer
    Dim strPeriodo      As String
    Dim intIntervalo    As Integer
    Dim dblVrDiferente  As Double
    VerificaPeriodo strPeriodo, intIntervalo
    strMsg = "Confirma divisão do empenho em " & Val(txtNumParcela) & " parcelas"
    If gblnExclusaoGravacaoOk("A", strMsg, True) Then
        txt_ValorParcela = gstrConvVrDoSql(txtdblValor / txtNumParcela)
        '------ verifica se a soma das parcelas é igual ao valor do empenho -----
        dblVrDiferente = Val(gstrConvVrParaSql(txt_ValorParcela)) * Val(txtNumParcela)
        If dblVrDiferente > txtdblValor Then
            dblVrDiferente = gstrConvVrDoSql(dblVrDiferente - txtdblValor)
            dblVrDiferente = txt_ValorParcela - dblVrDiferente
        Else
            dblVrDiferente = gstrConvVrDoSql(txtdblValor - dblVrDiferente)
            dblVrDiferente = dblVrDiferente + txt_ValorParcela
        End If
        
        '------------------------------------------------------------------------
        If gblnDataValida(txt_DataParcela) = False Then
            txt_DataParcela = txtDTMDATA
        End If
        For intInd = 1 To Val(txtNumParcela)
            If intInd = 1 Then
                If blnAtualizouParcela(lvw_ListaSubempenho.SelectedItem.Tag, _
                                       txt_DataParcela, _
                                       txt_ValorParcela, _
                                       txt_HistoricoSubEmpenho) = False Then
                    PreencheSaldoEmpenho
                    Exit Sub
                End If
            Else
                If intInd = Val(txtNumParcela) Then
                    txt_ValorParcela = dblVrDiferente
                End If
                If blnIncluiuParcela(txt_DataParcela, _
                                     txt_ValorParcela, _
                                     txt_HistoricoSubEmpenho, 1) = False Then
                    Exit Sub
                End If
            End If
            txt_DataParcela = gstrDataFormatada(DateAdd(strPeriodo, intIntervalo, txt_DataParcela))
        Next
        LeSubEmpenho lvw_ListaSubempenho, , txtPKId
        PreencheSaldoEmpenho
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcular
    End If
End Sub

Private Function blnIncluiuParcela(vntDataParcela, _
                                   vntValorParcela, _
                                   vntHistoricoSubEmpenho, _
                                   bytTipo As Byte)
    Dim strSQL  As String
        strSQL = ""
    
'    If blnLiqAutomatica And tab_3DPasta.Tab = 1 Then
'        'liquidaçao de parcela automatica
'        strSql = strSql & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
'    End If
    
    strSQL = strSQL & "INSERT INTO " & gstrSubempenho & " ("
    strSQL = strSQL & "intEmpenho, intNumero, dtmData, "
    strSQL = strSQL & "dblValor, bytSituacao, bytTipo, strHistorico, "
    strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr"
    
    If blnLiqAutomatica And tab_3dPasta.Tab = 1 Then
        'liquidaçao de parcela automatica
        strSQL = strSQL & ",dtmLiquidacao,intevento,dtmVencimento"
    End If
    
    strSQL = strSQL & ")"
    strSQL = strSQL & "SELECT " & txtPKId & ", "
    strSQL = strSQL & "MAX(intNumero) + 1, "
    strSQL = strSQL & gstrConvDtParaSql(vntDataParcela) & ", "
    strSQL = strSQL & gstrConvVrParaSql(vntValorParcela) & ", " & IIf(blnLiqAutomatica And tab_3dPasta.Tab = 1, "2", "1") & ", "
    strSQL = strSQL & bytTipo & ", "
    strSQL = strSQL & "'" & vntHistoricoSubEmpenho & "', "
    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr
    
    If blnLiqAutomatica And tab_3dPasta.Tab = 1 Then
        'liquidaçao de parcela automatica
        strSQL = strSQL & ", " & gstrConvDtParaSql(vntDataParcela) & ", " & gstrItemData(cbo_intEventoLiqAutomatica) & ", " & gstrConvDtParaSql(txt_dtmVenctoLiqAutomatica)
    End If

    strSQL = strSQL & " FROM " & gstrSubempenho & " "
    strSQL = strSQL & "WHERE intEmpenho = " & txtPKId
    
    If Not gobjBanco.Execute(strSQL) Then
        ExibeMensagem "Ocorreram erros durante a gravação da parcela"
    End If
    
    If blnLiqAutomatica And tab_3dPasta.Tab = 1 Then
    
        strSQL = "INSERT INTO " & gstrSubEmpenhoNF & " ("
        strSQL = strSQL & "intSubEmpenho, dtmData, dblValorNF, "
        strSQL = strSQL & "strNotaFiscal, dtmDtAtualizacao, lngCodUsr) "
        strSQL = strSQL & "(SELECT MAX(PKID) , "
        strSQL = strSQL & gstrConvDtParaSql(vntDataParcela) & ", "
        strSQL = strSQL & gstrConvVrParaSql(vntValorParcela) & ", '"
        strSQL = strSQL & txt_strNotasFiscaisLiqAutomatica & "', "
        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & " "
        strSQL = strSQL & "FROM " & gstrSubempenho & " )"
        
        If Not gobjBanco.Execute(strSQL) Then
            ExibeMensagem "Ocorreram erros durante a gravação da nota fiscal"
        End If
    End If
    
    
        
    If blnLiqAutomatica And tab_3dPasta.Tab = 1 Then
        If Not GeraMovimentosByEvento(gstrItemData(cbo_intEventoLiqAutomatica), CStr(vntDataParcela), Str(CDbl(vntValorParcela)), txt_HistoricoSubEmpenho, txtintNumero, "3") Then
              ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
        End If
        
        If Not blnGravaMovLiq(Val(txtPKId), glngRetornaPkidTabelaPai("seq" & gstrSubempenho, gstrSubempenho), gstrItemData(cboProgramaTrabalho, True), vntDataParcela, vntValorParcela, vntHistoricoSubEmpenho) Then
           ExibeMensagem "Ocorreram erros durante a gravação dos Movimentos de Liquidação."
        End If
        
    End If
    
    blnIncluiuParcela = True
        
End Function

Sub VerificaTabNovo()
    Select Case tab_3dPasta.Tab
    Case 0
        txt_CodHistorico.Text = ""
        LimpaObjeto Me, mblnAlterandoEmpenho
        Limpa_Controles Me, True, False, False, True, True
        LimpaDadosReserva
        blnImportadoPedidoEmpenho = False
        Set cbo_Historico.DataSource = Nothing
        TrocaCorObjeto cmd_Credor, False
        TrocaCorObjeto cmd_Evento, False
        TrocaCorObjeto cmd_Reserva, False
        TrocaCorObjeto cmd_ProgramaTrabalho, False
        TrocaCorObjeto txtstrsolicitacao, False
        TrocaCorObjeto txtStrlocentrega, False
        TrocaCorObjeto txtStrcondpagto, False
        TrocaCorObjeto txtStrprazoentrega, False
        TrocaCorObjeto dbcintTipo, False
        TrocaCorObjeto cmd_Tipo, False
        
        'M4R
        If mblnRestosAPagar Then
           TrocaCorObjeto txtintExercicioEmpenho, False
        Else
           txtintExercicioEmpenho.Text = gintExercicio
        End If
        txt_intNContribuinte = ""
        txtstrCodigo = ""
        txtbitDigito = ""
        txtintExercicio = ""
        
        TrocaCorObjeto cboCodigoReduzido, False
        TrocaCorObjeto cboProgramaTrabalho, False
        
        HabilitaDesabilitaTab tab_3dPasta, False
        'txtdtmData = strUltimaData
        
        TrocaCorObjeto cbo_intEvento, False
        TrocaCorObjeto txt_codEvento, False
        
        TrocaCorObjeto dbcintCredor, False
        TrocaCorObjeto cmd_Credor, False
        txt_intNContribuinte.Text = ""
        txt_intNContribuinte.Enabled = True
        txt_intNContribuinte.BackColor = vbWindowBackground
        TrocaCorObjeto dbcintModalidade, False
        TrocaCorObjeto txtstrModalidade, False
        TrocaCorObjeto txtdtmHomologacao, False

        
        TrocaCorObjeto cbointReservaDotacao, False
        TrocaCorObjeto cmd_Reserva, False

        TrocaCorObjeto cmd_Evento, False
        TrocaCorObjeto cmd_ProgramaTrabalho, False
        
        txt_codEvento.BackColor = cbo_intEvento.BackColor
        cbointReservaDotacao.BackColor = vbWindowBackground
        cboCodigoReduzido.BackColor = vbWindowBackground
        cboProgramaTrabalho.BackColor = vbWindowBackground
        txt_codEvento.BackColor = vbWindowBackground
        cbo_intEvento.BackColor = vbWindowBackground
        txt_codEvento.Enabled = True
        cbointReservaDotacao.Enabled = True
        cboCodigoReduzido.Enabled = True
        cboProgramaTrabalho.Enabled = True
        txt_codEvento.Enabled = True
        cbo_intEvento.Enabled = True
        cbo_intEvento.ListIndex = -1
        tab_3DEmpenho.Tab = 0
        tab_3DGeral.Tab = 0
        
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
        VerificaDataAutomatica
        If blnDataAutomatica = True Then
            DataAutomatica
        Else
            txtDTMDATA.Text = ""
        End If
        
        TrocaCorObjeto frmCadEmpenho.txt_intCodigo, False
        TrocaCorObjeto frmCadEmpenho.txt_intCatalogoMaterialServico, False
        TrocaCorObjeto frmCadEmpenho.txt_dblQuantidade, False
        TrocaCorObjeto frmCadEmpenho.txt_dblValorEstimado, False
        TrocaCorObjeto frmCadEmpenho.txt_intUnidadedeMedida, False
        TrocaCorObjeto frmCadEmpenho.txt_strObsItem, False
        TrocaCorObjeto frmCadEmpenho.txt_strdescricaodetalhada, False
        TrocaCorObjeto frmCadEmpenho.dbc_intStrMarca, False
        
        If Not mblnRestosAPagar Then txtintExercicioEmpenho = gintExercicio
        proximoCodigoEmpenho
        Screen.MousePointer = vbDefault
        
        'Pen_687_ORC_100
        tab_3dPasta.TabEnabled(3) = True
        txt_dblValorAux.Visible = True
        txt_DataLiuidacao.Visible = True
        txt_dblDesconto.Visible = True
        txt_codEventoLiq.Visible = True
        cbo_intEventoLiq.Visible = True
        cmd_EventoLiq.Visible = True
        txt_HistoricoLiquidacao.Visible = True
        cbo_HistoricoLiquidacao.Visible = True
        cmd_HistoricoLiquidacao.Visible = True
        lvw_Liquidacao.Visible = True
        cbo_ContaExtra.Visible = True
        cbo_DescricaoExtra.Visible = True
        cmd_ContaExtra.Visible = True
        txt_ValorExtra.Visible = True
        lvw_Extra.Visible = True
        cbo_ContaOrcamentario.Visible = True
        cbo_DescricaoOrcamentario.Visible = True
        cmd_ContaOrcamentario.Visible = True
        txt_ValorOrcamentario.Visible = True
        lvw_Orcamentario.Visible = True
        txt_dtmDataNF.Visible = True
        txt_dblValorNF.Visible = True
        txt_strNotasFiscais.Visible = True
        lvw_NotasFiscais.Visible = True
        HabilitaControlesLiquidacao False
        
        frmCadEmpenho.Visible = True
        If Not mblnRestosAPagar Then ProximaData

        
    Case 1
        LimpaTelaSubempenho
    Case 2
        LimpaTelaComplemento
    Case 3
        VerificaNovaLiquidacao
    Case 4
        Select Case tab_3DAnulacao.Tab
           Case 0: LimpaTelaAnulacao
           Case 1: LimpaCamposAnulado
        End Select
        
    End Select
    PreencheSaldoEmpenho
    
    If txtintNumero.Enabled And txtintNumero.Visible Then
       txtintNumero.SetFocus
    End If
    dblSaldoItem = 0
End Sub

Private Sub LimpaTelaComplemento()
    txt_DataComplemento = ""
    txt_ValorComplemento = ""
    txt_HistoricoComplemento = ""
    cbo_HistoricoComplemento.Text = ""
    txt_CodHistoricoComp.Text = ""
    cbo_HistoricoComplemento.ListIndex = -1
    TrocaCorObjeto txt_ValorComplemento, False
    mblnAlterandoComplemento = False
    If txt_DataComplemento.Enabled Then txt_DataComplemento.SetFocus
    AtualizaTotalComplemento
End Sub

Sub VerificaNovaLiquidacao()
    
    Select Case tab_3DPastaLiquidacao.Tab
    Case 0
        mblnCriarParcelaLiquidada = True
        txt_CodHistoricoLiq.Text = ""
        lvw_Liquidacao.SelectedItem = Nothing
        LimpaDadosLiquidacao
        HabilitaDesabilitaTab tab_3DPastaLiquidacao, True
        TrocaCorObjeto txt_DataLiuidacao, False
        TrocaCorObjeto txt_DataVencto, False
        TrocaCorObjeto txt_dblValorAux, False
        lvw_Extra.ListItems.Clear
        lvw_NotasFiscais.ListItems.Clear
        LimpaDadosExtra
        LimpaDadosOrcamentario
        LimpaDadosNF
        txt_DataLiuidacao.SetFocus
        lbl_ValorTotal = "0,00"
        'M4R
        TrocaCorObjeto cbo_intEventoLiq, False
        TrocaCorObjeto txt_codEventoLiq, False
        TrocaCorObjeto cmd_EventoLiq, False
        TrocaCorObjeto txt_HistoricoLiquidacao, False
        TrocaCorObjeto cbo_HistoricoLiquidacao, False
        TrocaCorObjeto txt_dblDesconto, True
        DoEvents
        tab_3DPastaLiquidacao.Tab = 0
        
        If mobjAux Is Nothing = False Then
            If Trim(frmCadOrdemPagamento.txtDTMDATA.Text) <> "" Then
                txt_DataLiuidacao = frmCadOrdemPagamento.txtDTMDATA
            End If
        End If
    Case 1
        LimpaDadosExtra
    Case 2
        LimpaDadosRetencao
    Case 3
        LimpaDadosOrcamentario
    Case 4
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
        
        TrocaCorObjeto txt_dtmDataNF, False
        TrocaCorObjeto txt_dblValorNF, False
        TrocaCorObjeto txt_strNotasFiscais, False
        LimpaDadosNF
    End Select
End Sub

Private Sub VerificaParametroSubElmentos()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim resultado       As Integer

    strSQL = ""
    strSQL = strSQL & " Select bytempenhosubElementos "
    strSQL = strSQL & " FROM " & gstrConfiguracaoGeral
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
       With adoResultado
        If Not .EOF Then
            If !bytempenhosubElementos = 0 Or mblnRestosAPagar Then
                tab_3DEmpenho.TabVisible(2) = False
                txtItemDespSubElemento.Visible = False
                dbcItemDespSubElemento.Visible = False
                cmd_ItemDespSubElemento.Visible = False
                txtDblValorSubElemento.Visible = False
                lvwSubElemento.Visible = False
                frmSubElementoEstorno.Visible = False
            Else
                lbl_ItemDespesa.Visible = False
                dbcintItemDespesa.Visible = False
                txt_intCodItemDespesa.Visible = False
                cmd_ItemDespesa.Visible = False
                'frmCredorTipo.Top = 540
                tab_3DEmpenho.TabVisible(2) = True
                txtItemDespSubElemento.Visible = True
                dbcItemDespSubElemento.Visible = True
                cmd_ItemDespSubElemento.Visible = True
                txtDblValorSubElemento.Visible = True
                lvwSubElemento.Visible = True
                frmSubElementoEstorno.Visible = True
                lvw_Anulacao.Top = frmSubElementoEstorno.Top + frmSubElementoEstorno.Height + 100
                lvw_Anulacao.Height = tab_3DAnulacao.Height - lvw_Anulacao.Top - 100
            End If
        End If
       End With
    Else
        
    End If
End Sub


Private Sub Form_Load()
    
    mblnAtivarPastas = False
    mblnAbrindo = True
    frmCadOrdemPagamento.mblnTelaEmpenho = False
    mblnClick = False
    blnImportadoPedidoEmpenho = False
    dbcintTipo.Tag = "SELECT PKID, strDescricao FROM " & gstrTipoEmpenho & " ORDER BY strDescricao;strDescricao"
    dbcintConvenio.Tag = "SELECT PKID, strDescricao FROM " & gstrConvenio & " ORDER BY strDescricao;strDescricao"
    dbcintFundo.Tag = "SELECT PKID, strDescricao FROM " & gstrFundo & " ORDER BY strDescricao;strDescricao"
    dbcintCredor.Tag = "SELECT PKID, strNome FROM " & gstrContribuinte & " ORDER BY strNome;strNome"
    
    dbc_intStrMarca.Tag = "Select Pkid, strMarca From " & gstrMarcas & " ORDER BY strMarca;strMarca"
    dbcintItemDespesa.Tag = "SELECT PKID, strDescricao FROM " & gstrItemDespesa & " ORDER BY strDescricao;strDescricao"
    dbcItemDespSubElemento.Tag = "SELECT PKID, strDescricao FROM " & gstrItemDespesa & " ORDER BY strDescricao;strDescricao"
    cbo_Historico.Tag = "SELECT strCodigo, strDescricao FROM " & gstrHistorico & " ORDER BY strDescricao;strDescricao"
    dbcintModalidade.Tag = "SELECT PKID, strCodigo FROM " & gstrComprasLicitacao & " ORDER BY strCodigo;strCodigo"
    TrocaCorObjeto txtintExercicioEmpenho, Not mblnRestosAPagar
    'VerificaListaAutomatica gstrEmpenho, tdb_Lista, strQuery
    VerificaListaAutomatica gstrEmpenho, tdb_Lista, strQueryLocalizar
    VerificaObjParaAplicar mobjAux
    HabilitaDesabilitaTab tab_3dPasta, False
    mblnAlterandoEmpenho = False
    mblnAtualizaTelaSubempenho = True
    mlngEmpenhoAnulacao = 0
    mlngEmpenhoLiquidacao = 0
    
    
    lvw_Liquidacao.ColumnHeaders(10).Position = 3
    LimpaObjeto Me, mblnAlterandoEmpenho
    preencheCboevento
    If cbo_intEvento.ListCount = 1 Then
        cbo_intEvento.ListIndex = 0
    End If
    
    preencheCboeventoLiq
    preencheCboeventoLiqAutomatica
    LeDaTabelaParaObj gstrEvento, cbo_intEventoAnul, strQueryAplicarEventoAnul
     
    LeTabelaProgramaTrabalho
    
    fra_EventoContabil.Visible = mblnRestosAPagar
    
    If cbo_intEventoLiq.ListCount = 1 Then
        cbo_intEventoLiq.ListIndex = 0
    End If
    
    cbointReservaDotacao.Enabled = True
    If mblnRestosAPagar Then
       DesabilitaControlesRP
       Me.Caption = "Restos a Pagar"
    Else
       Me.Caption = "Empenhos"
       txtintExercicioEmpenho.Text = gintExercicio
    End If
    
    VerificaDataAutomatica
    VerificaParametroSubElmentos
    
    If blnLiqAutomatica Then
        frm_liqAutomatico.Visible = True
    Else
        txt_ValorParcela.Top = txt_ValorParcela.Top + 250
        txt_DataParcela.Top = txt_ValorParcela.Top
        lbl_ValorParcela.Top = lbl_ValorParcela.Top + 250
        lbl_DataParcelamento.Top = lbl_ValorParcela.Top
        frm_liqAutomatico.Visible = False
    End If
    Me.chc_LiquidarAutomaticamente.Value = Abs(CInt(blnLiqAutomatica))
    
    tab_3DPastaLiquidacao.TabVisible(2) = False
    cbo_ContaRetencao.Visible = False
    cbo_DescricaoRetencao.Visible = False
    cmd_ContaRetencao.Visible = False
    txt_ValorRetencao.Visible = False
    lvw_Retencao.Visible = False
    tab_3DAnulacao.TabEnabled(1) = False
      
    'If Not mblnRestosAPagar Then txtintExercicioEmpenho = gintExercicio
    tab_3DPastaLiquidacao.Tab = 0
    tab_3DEmpenho.Tab = 0
    tab_3DGeral.Tab = 0
    tab_3DAnulacao.Tab = 0
    tab_3dPasta.Tab = 0
    tab_3DAnulacao.TabEnabled(1) = False
    
    'Pen_687_ORC_100 ===============================================================================
        tab_3dPasta.TabEnabled(3) = True
        txt_dblValorAux.Visible = True
        txt_DataLiuidacao.Visible = True
        txt_dblDesconto.Visible = True
        txt_codEventoLiq.Visible = True
        cbo_intEventoLiq.Visible = True
        cmd_EventoLiq.Visible = True
        txt_HistoricoLiquidacao.Visible = True
        cbo_HistoricoLiquidacao.Visible = True
        cmd_HistoricoLiquidacao.Visible = True
        lvw_Liquidacao.Visible = True
        cbo_ContaExtra.Visible = True
        cbo_DescricaoExtra.Visible = True
        cmd_ContaExtra.Visible = True
        txt_ValorExtra.Visible = True
        lvw_Extra.Visible = True
        cbo_ContaOrcamentario.Visible = True
        cbo_DescricaoOrcamentario.Visible = True
        cmd_ContaOrcamentario.Visible = True
        txt_ValorOrcamentario.Visible = True
        lvw_Orcamentario.Visible = True
        txt_dtmDataNF.Visible = True
        txt_dblValorNF.Visible = True
        txt_strNotasFiscais.Visible = True
        lvw_NotasFiscais.Visible = True
        HabilitaControlesLiquidacao False
    
        frmCadEmpenho.Visible = True
        
    HabilitaControlesLiquidacao False
    'Pen_687_ORC_100 ===============================================================================

    mblnAtivarPastas = True
        txtintNumero.SetFocus
        
    If Not mblnRestosAPagar Then ProximaData
    TrocaCorDeFundoObjeto True, txt_SubTotalItem
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnselecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub LeTabelaProgramaTrabalho(Optional strFiltroPKID As Variant, Optional strLIKE As Variant)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim intIndicePkid   As Integer
   
       mblnAtivarPastas = False
   
'   If cbo_intEvento.ListIndex <> -1 Then
      cboProgramaTrabalho.Clear
      cboCodigoReduzido.Clear
      
      
      strSQL = ""
      strSQL = strSQL & "SELECT PT.PKId, PT.intCodigoReduzido, PT.strCodigo, "
      strSQL = strSQL & " ED.strCodigoElementoDespesa "
      strSQL = strSQL & " FROM " & gstrProgramaDeTrabalho & " PT, "
      strSQL = strSQL & gstrElementoDespesa & " ED "
      strSQL = strSQL & " WHERE PT.intElementoDespesa = ED.PKID "
      
'      If cbo_intEvento.ListIndex <> -1 Then
'            strSql = strSql & " AND " & strSUBSTRING & "(ED.strCodigoElementoDespesa,1," & Len(BuscaCodigosPeloEvento(gstrItemData(cbo_intEvento), gstrDigitoDespesa, "D", 2)) & ") = '" & _
'                                                                      BuscaCodigosPeloEvento(gstrItemData(cbo_intEvento), gstrDigitoDespesa, "D", 2) & "'"
'      End If
        
      If Not IsMissing(strFiltroPKID) Then
         strSQL = strSQL & " AND PT.PKID = " & strFiltroPKID
      Else
         strSQL = strSQL & " AND PT.intExercicio = "
         '10/03/05 - Consiste data em branco
         If IsDate(txtDTMDATA) Then
            strSQL = strSQL & Year(txtDTMDATA)
         Else
            strSQL = strSQL & gintExercicio
         End If
      End If
      
      If Not IsMissing(strLIKE) Then
         strSQL = strSQL & " AND PT.strcodigo LIKE '" & strLIKE & "%'"
      End If
      
      strSQL = strSQL & " ORDER BY PT.strCodigo"
      
      Set gobjBanco = New clsBanco
      If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
          With adoResultado
              Do While Not .EOF
                  If Not IsNull(!strCodigo) Then
                    cboProgramaTrabalho.AddItem !strCodigo
                    cboProgramaTrabalho.ItemData(cboProgramaTrabalho.NewIndex) = !Pkid
                  End If
                  .MoveNext
              Loop
          End With
      End If
      
      strSQL = Mid(strSQL, 1, Len(strSQL) - 21) + "ORDER BY intCodigoReduzido"
      
      If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
          With adoResultado
              Do While Not .EOF
                  If Not IsNull(!intCodigoReduzido) Then
                     cboCodigoReduzido.AddItem (!intCodigoReduzido)
                  Else
                     cboCodigoReduzido.AddItem ""
                  End If
                  cboCodigoReduzido.ItemData(cboCodigoReduzido.NewIndex) = !Pkid
                  .MoveNext
              Loop
          End With
      End If
'   Else
'      If Not mblnAlterandoEmpenho Then ExibeMensagem "Informe o Evento Contábil antes de escolher a Dotação."
'   End If

        mblnAtivarPastas = True
End Sub
Private Sub LeTabelaReservaDotacao(Optional strFiltroPKID As Variant)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    
    If mblnRestosAPagar Then Exit Sub
    'If cbo_intEvento.ListIndex <> -1 Then
       cbointReservaDotacao.Clear
      

           'strsql = "SELECT * FROM ("
           strSQL = strSQL & "SELECT RD.PKId, RD.intNumero,  (RD.dblValor - " & gstrISNULL("SUM(RDL.dblValor)", "0") & ") dblValor,"
           strSQL = strSQL & " PT.intCodigoReduzido, PT.strCodigo,ED.strCodigoElementoDespesa "
           strSQL = strSQL & " FROM " & gstrProgramaDeTrabalho & " PT, "
           strSQL = strSQL & gstrElementoDespesa & " ED, "
           strSQL = strSQL & gstrReservaDotacao & " RD, "
           strSQL = strSQL & gstrReservaDotacaoLiberada & " RDL "
           strSQL = strSQL & " WHERE RD.PKID " & strOUTJSQLServer & "= RDL.intReservaDotacao" & strOUTJOracle & " AND PT.Pkid = RD.intProgramaTrabalho AND PT.intElementoDespesa = ED.Pkid "
           
           If cbo_intEvento.ListIndex > -1 Then
                strSQL = strSQL & " AND " & strSUBSTRING & "(ED.strCodigoElementoDespesa,1," & Len(BuscaCodigosPeloEvento(gstrItemData(cbo_intEvento), gstrDigitoDespesa, "D", 2)) & ") = '" & _
                                              BuscaCodigosPeloEvento(gstrItemData(cbo_intEvento), gstrDigitoDespesa, "D", 2) & "'"
           End If
           If txtintExercicioEmpenho = "" Then
                txtintExercicioEmpenho = gintExercicio
            End If
            
            strSQL = strSQL & " AND PT.intExercicio = " & txtintExercicioEmpenho
            strSQL = strSQL & " AND RD.intExercicioReserva = " & txtintExercicioEmpenho
                
        
           
           If Not IsMissing(strFiltroPKID) Then
              strSQL = strSQL & " AND RD.PKID = " & strFiltroPKID
           End If
           
           
           strSQL = strSQL & " GROUP BY RD.PKID, RD.intNumero, RD.dblValor, PT.intCodigoReduzido, PT.strCodigo, ED.strCodigoElementoDespesa "
           'strsql = strsql & " WHERE RD.dblValor > 0 "
          
          If Not mblnAlterandoEmpenho Then
                strSQL = strSQL & " HAVING (RD.dblValor -  " & gstrISNULL("SUM(RDL.dblValor)", "0") & " ) > 0"
          End If
           
           strSQL = strSQL & " ORDER BY RD.intNumero"
          
       
          
       Set gobjBanco = New clsBanco
       If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
           With adoResultado
               Do While Not .EOF
                    cbointReservaDotacao.AddItem !INTNUMERO
                    cbointReservaDotacao.ItemData(cbointReservaDotacao.NewIndex) = !Pkid
                    .MoveNext
               Loop
           End With
       End If
'    Else
'
'       If Not mblnAlterandoEmpenho Then ExibeMensagem "É necessário informar o Evento Contábil antes de informar a Reserva."
       
'    End If
End Sub

Private Sub preencheReservaDotacaoByCodigo(ByVal strCodigo As String)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim resultado       As Integer

    strSQL = ""
    strSQL = strSQL & "SELECT PKId, intNumero "
    strSQL = strSQL & "FROM " & gstrReservaDotacao & " "
    strSQL = strSQL & "WHERE " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio
    If strCodigo <> "" Then
        strSQL = strSQL & " AND intNumero = " & strCodigo
    End If
    strSQL = strSQL & " ORDER BY intNumero"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
       With adoResultado
        If .RecordCount >= 1 Then
           cbointReservaDotacao.Clear
           cbointReservaDotacao.AddItem !INTNUMERO
           cbointReservaDotacao.ItemData(cbointReservaDotacao.NewIndex) = !Pkid
           cbointReservaDotacao.ListIndex = 0
        End If
       End With
    Else
        cbointReservaDotacao.ListIndex = -1
        cbointReservaDotacao.Text = strCodigo
    End If
End Sub



Private Sub txtintExercicioEmpenho_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicioEmpenho
End Sub

Private Sub txt_intNContribuinte_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "N", txt_intNContribuinte
End Sub

Private Sub txt_intNContribuinte_LostFocus()
Dim strPKId As String
Dim strSQL As String
   
   
   strPKId = LeCDCCredor(, txt_intNContribuinte)
   
   
   If strPKId = "" Then
        dbcintCredor.BoundText = ""
        Exit Sub
   End If
   
    If Len(Trim(txt_intNContribuinte)) > 0 Then
        If dbcintCredor.Enabled Then dbcintCredor.SetFocus
            strSQL = "SELECT CO.PKID,"
           strSQL = strSQL & " CO.STRNOME"
           strSQL = strSQL & " FROM "
           strSQL = strSQL & gstrContribuinte & " CO, "
           strSQL = strSQL & gstrItens & " IT, "
           strSQL = strSQL & gstrModuloContribuinte & " MC"
           strSQL = strSQL & " WHERE CO.PKID = " & strPKId & " AND"
           strSQL = strSQL & " IT.PKId = MC.intItem AND"
           strSQL = strSQL & " MC.intContribuinte = CO.Pkid AND"
           strSQL = strSQL & " IT.Pkid =" & gintModulo & " AND CO.BLNINATIVO = 0"
        
           'Cláudio
           'LeDaTabelaParaObj gstrContribuinte, dbcintCredor, "SELECT PKID, strNome FROM " & gstrContribuinte & _
                                                          " WHERE PKID = " & txt_intNContribuinte
                                                          
           LeDaTabelaParaObj gstrContribuinte, dbcintCredor, strSQL
                                                               
           dbcintCredor.BoundText = strPKId
    End If
  
End Sub

Private Sub txtdtmData_GotFocus()
    MarcaCampo txtDTMDATA
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDTMDATA
End Sub

Private Sub txtdtmData_LostFocus()
    Dim blnMantemValor_blnAlteraReserva As Boolean

    txtDTMDATA = gstrDataFormatada(txtDTMDATA)
    If Not IsDate(txtDTMDATA) And txtDTMDATA <> "" Then
           ExibeMensagem txtDTMDATA & " não é uma data válida"
           txtDTMDATA.SetFocus
    Else
        If Len(Trim(Me.txtDTMDATA.Text)) > 0 Then
            Me.txtintExercicioEmpenho.Text = Year(CDate(Me.txtDTMDATA.Text))
        End If
    End If
    
    blnMantemValor_blnAlteraReserva = blnAlteraReserva
    blnAlteraReserva = False
    cboProgramaTrabalho_Click
    blnAlteraReserva = blnMantemValor_blnAlteraReserva
End Sub

Private Sub txt_dtmDataNF_GotFocus()
    MarcaCampo txt_dtmDataNF
    'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
    'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
    VerificaTabAtivo
    If mblnselecionou = False And mblnAbrindo = False Then
        mAtivaPastaDeObjeto tab_3dPasta, 3, tab_3DPastaLiquidacao, 4
    End If
End Sub

Private Sub txt_dtmDataNF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataNF
End Sub

Private Sub txt_dtmDataNF_LostFocus()

    txt_dtmDataNF = gstrDataFormatada(txt_dtmDataNF)
    
End Sub

Private Sub txt_DataAnulucao_GotFocus()
    MarcaCampo txt_DataAnulucao
    If tab_3dPasta.TabEnabled(4) = False Then
       txt_codEvento.SetFocus
    End If
    If mblnAbrindo = False Then
    mAtivaPastaDeObjeto tab_3dPasta, 4, tab_3DAnulacao, 0
    Else
     mblnAbrindo = False
     If tab_3dPasta.Enabled Then tab_3dPasta.SetFocus
    End If
End Sub

Private Sub txt_DataAnulucao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataAnulucao
End Sub

Private Sub txt_DataAnulucao_LostFocus()

    txt_DataAnulucao = gstrDataFormatada(txt_DataAnulucao)
    
    'ORC677
    If IsDate(txt_DataAnulucao) Then
        If Year(CDate(txt_DataAnulucao)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data de Anulação tem que estar no exercício de " & gintExercicio & "."
            If txt_DataAnulucao.Enabled Then txt_DataAnulucao.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txt_DataComplemento_GotFocus()
    MarcaCampo txt_DataComplemento
    If mblnselecionou = False Then
    mAtivaPastaDeObjeto tab_3dPasta, 2
    End If
End Sub

Private Sub txt_DataComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataComplemento
End Sub

Private Sub txt_DataComplemento_LostFocus()

    txt_DataComplemento = gstrDataFormatada(txt_DataComplemento)
    
    'ORC677
    If IsDate(txt_DataComplemento) Then
        If Year(CDate(txt_DataComplemento)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data de Complemento tem que estar no exercício de " & gintExercicio & "."
            If txt_DataComplemento.Enabled Then txt_DataComplemento.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txt_DataLiuidacao_GotFocus()
    MarcaCampo txt_DataLiuidacao
    If mblnselecionou = False Then
    mAtivaPastaDeObjeto tab_3dPasta, 3
    End If
End Sub

Private Sub txt_DataLiuidacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataLiuidacao
End Sub

Private Sub txt_DataLiuidacao_LostFocus()

    txt_DataLiuidacao = gstrDataFormatada(txt_DataLiuidacao)
    
    'ORC677
    If IsDate(txt_DataLiuidacao) Then
        If Year(CDate(txt_DataLiuidacao)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data de Liquidação tem que estar no exercício de " & gintExercicio & "."
            If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txt_DataParcela_GotFocus()
    MarcaCampo txt_DataParcela
    If mblnAbrindo = False Then
    mAtivaPastaDeObjeto tab_3dPasta, 1
    'Else
    'mblnAbrindo = False
    End If
End Sub

Private Sub txt_DataParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataParcela
End Sub

Private Sub txt_DataParcela_LostFocus()

    txt_DataParcela = gstrDataFormatada(txt_DataParcela)
    
    'ORC677
    If IsDate(txt_DataParcela) Then
        If Year(CDate(txt_DataParcela)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data da parcela tem que estar no exercício de " & gintExercicio & "."
            If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txt_ElementoDespesa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtintNumero_GotFocus()
    proximoCodigoEmpenho
End Sub

Private Sub txtintNumero_LostFocus()
    Dim intIdx As Integer
    Dim strPkidEmpenho As String
    Dim strSQL As String
    Dim adoResultado  As New ADODB.Recordset
    
     mblnAbrindo = Not mblnAlterandoEmpenho
    
    If mblnRestosAPagar Or mblnAlterandoEmpenho Then Exit Sub
        
    If Len(Trim(txtintNumero)) > 0 And BuscaEmpenho <> Empty Then
        strPkidEmpenho = Str(BuscaEmpenho)
                    
        If Len(txtPKId.Text) = 0 Then
            txtPKId.Text = BuscaEmpenho
        End If
        
        If mblnEmpenhoEstorno Then
           tab_3dPasta.Tab = 4
           For intIdx = 1 To lvw_Anulacao.ListItems.Count
               If lvw_Anulacao.ListItems(intIdx).SubItems(9) = txtintNumero Then
                  lvw_Anulacao.ListItems(intIdx).Selected = True
               End If
           Next
        End If

        'Set adoResultado = tdb_Lista.DataSource
        SetaRecordset tdb_Lista.DataSource, adoResultado
        'strSql = adoResultado.Source
        Set gobjBanco = New clsBanco
           
        If adoResultado.State = 1 Then
            If Not adoResultado.EOF Then adoResultado.MoveFirst
            'roda o grid para selecionar o pkid
            Do While Not adoResultado.EOF
                If adoResultado!Pkid = Val(strPkidEmpenho) Then
                    While Not tdb_Lista.EOF
                        If tdb_Lista.Columns(0).Value = Val(strPkidEmpenho) Then
                            mblnClickOk = True
                            mblnAbrindo = True
                            mblnPrimeiraVez = False
                            mblnselecionou = True
                            tdb_Lista_RowColChange 0, 0
                            txtPKId.Text = tdb_Lista.Columns(0).Value
                            Exit Sub
                        End If
                        tdb_Lista.MoveNext
                    Wend
                    Exit Sub
                End If
                adoResultado.MoveNext
            Loop
        End If
    
        'caso o pkid encotrado na tenha no grid
      
        strSQL = "SELECT DISTINCT EP.PKId, EP.intNumero , "
        strSQL = strSQL & "EP.dtmData,EP.dblValor , PT.intCodigoReduzido, PT.strCodigo "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrEmpenho & " EP, "
        strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
        strSQL = strSQL & "WHERE EP.intProgramaTrabalho = PT.PKId "
        strSQL = strSQL & " AND EP.PKID  = " & strPkidEmpenho
        
        LeDaTabelaParaObj gstrEmpenho, tdb_Lista, strSQL
        mblnClickOk = True
        mblnPrimeiraVez = False
        tdb_Lista_RowColChange 0, 0
        
        If Habilitando_Modalidade = True Then
            TrocaCorObjeto dbcintModalidade, True
            TrocaCorObjeto txtstrModalidade, True
        Else
            TrocaCorObjeto dbcintModalidade, False
            TrocaCorObjeto txtstrModalidade, False
        End If
    End If

End Sub

Private Sub SetaRecordset(ByVal adoRecOrigem As ADODB.Recordset, ByRef adoRecdestino As ADODB.Recordset)
    Set adoRecdestino = adoRecOrigem
End Sub


Private Sub txtstrEmbasamento_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 1
    MarcaCampo txtstrEmbasamento
End Sub

Private Sub txtstrEmbasamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrEmbasamento
End Sub

Private Sub txt_Funcao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrHistorico_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 0
    MarcaCampo txtstrHistorico
End Sub

Private Sub txtstrHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrHistorico
End Sub

Private Sub txtstrModalidade_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 1
    MarcaCampo txtstrModalidade
End Sub

Private Sub txtstrModalidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrModalidade
End Sub

Private Sub txtstrContrato_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 1
    MarcaCampo txtstrContrato
End Sub

Private Sub txtstrContrato_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrContrato
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintNumero
End Sub

Private Sub txtstrLicitacao_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 1
    MarcaCampo txtstrLicitacao
End Sub

Private Sub txtstrLicitacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii ', "N", txtstrLicitacao
End Sub

Private Sub txtNumParcela_Change()
    If (Val(txtNumParcela) > 1 And mblnDigitouQtdParcela) And cboPeriodo.ListIndex > -1 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcular
        mblnDigitouQtdParcela = False
    Else
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcular
    End If
End Sub

Private Sub txtNumParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtNumParcela
    mblnDigitouQtdParcela = True
End Sub

Private Sub txt_strNotasFiscais_GotFocus()
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
   
   MarcaCampo txt_strNotasFiscais
    VerificaTabAtivo
End Sub

Private Sub txtstrSolicitacao_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 1
    MarcaCampo txtstrLicitacao
End Sub

Private Sub txtstrSolicitacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLicitacao
End Sub

Private Sub txt_Orgao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_Programa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_Projetoatividade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_Subfuncao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_SubPrograma_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_Subunidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_TipoCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_UnidadeOrcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
End Sub

Private Sub txtdblValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValor
End Sub

Private Sub txtdblValor_LostFocus()
      
   txtdblValor = gstrConvVrDoSql(txtdblValor)

End Sub

Private Sub txt_dblValorAux_GotFocus()
    MarcaCampo txt_dblValorAux
    mAtivaPastaDeObjeto tab_3dPasta, 3
End Sub

Private Sub txt_dblValorAux_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorAux
End Sub

Private Sub txt_dblValorAux_LostFocus()
    txt_dblValorAux = gstrConvVrDoSql(txt_dblValorAux)
End Sub

Private Sub txt_dblValorNF_GotFocus()
    MarcaCampo txt_dblValorNF
    'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
    'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
    VerificaTabAtivo
End Sub

Private Sub txt_dblValorNF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorNF
End Sub

Private Sub txt_dblValorNF_LostFocus()
    txt_dblValorNF = gstrConvVrDoSql(txt_dblValorNF)
End Sub

Private Sub txtdtmHomologacao_GotFocus()
    mAtivaPastaDeObjeto tab_3dPasta, 0, tab_3DEmpenho, 1
    MarcaCampo txtdtmHomologacao
End Sub

Private Sub txtdtmHomologacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmHomologacao
End Sub

Private Sub txtdtmHomologacao_LostFocus()

    txtdtmHomologacao = gstrDataFormatada(txtdtmHomologacao)
    
    'ORC677
    If IsDate(txtdtmHomologacao) Then
        If Year(CDate(txtdtmHomologacao)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data da Homologação tem que estar no exercício de " & gintExercicio & "."
            If txtdtmHomologacao.Enabled Then txtdtmHomologacao.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txt_ValorAnulacao_GotFocus()
    MarcaCampo txt_ValorAnulacao
End Sub

Private Sub txt_ValorAnulacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_ValorAnulacao
End Sub

Private Sub txt_ValorAnulacao_LostFocus()
    txt_ValorAnulacao = gstrConvVrDoSql(txt_ValorAnulacao)
End Sub

Private Sub txt_ValorComplemento_GotFocus()
    MarcaCampo txt_ValorComplemento
End Sub

Private Sub txt_ValorComplemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_ValorComplemento
End Sub

Private Sub txt_ValorComplemento_LostFocus()
    Dim intInd  As Integer
    Dim dblSoma As Double
    txt_ValorComplemento = gstrConvVrDoSql(txt_ValorComplemento)
    If mblnAlterandoComplemento Then
        With lvw_Complemento
            .ListItems(.SelectedItem.Index).SubItems(2) = txt_ValorComplemento
            For intInd = 1 To .ListItems.Count
                dblSoma = dblSoma + .ListItems(intInd).SubItems(2)
            Next
        End With
        lblTotalComplemento = gstrConvVrDoSql(dblSoma)
    Else
        lblTotalComplemento = gstrConvVrDoSql(Val(gstrConvVrParaSql(lblTotalComplemento)) + _
                                              Val(gstrConvVrParaSql(txt_ValorComplemento)))
    End If
End Sub

Private Sub txt_ValorExtra_GotFocus()
    MarcaCampo txt_ValorExtra
   ' HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem
   VerificaTabAtivo
End Sub

Private Sub txt_ValorExtra_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_ValorExtra
End Sub

Private Sub txt_ValorExtra_LostFocus()
    'txt_ValorExtra = gstrConvVrDoSql(txt_ValorExtra)
    'If mblnAlterandoExtra Then
    '    AtualizaListView lvw_Extra, txt_ValorExtra, lblExtra
    'Else
    '    lblExtra = gstrConvVrDoSql(Val(gstrConvVrParaSql(lblExtra)) + _
    '                               Val(gstrConvVrParaSql(txt_ValorExtra)))
    'End If
    txt_ValorExtra = gstrConvVrDoSql(txt_ValorExtra)
End Sub

Sub AtualizaListView(lvw_Lista As ListView, _
                     txtdblValor As TextBox, _
                     lblSoma As Label)
    Dim intInd  As Integer
    Dim dblSoma As Double
    With lvw_Lista
        .ListItems(.SelectedItem.Index).SubItems(2) = txtdblValor
        For intInd = 1 To .ListItems.Count
            dblSoma = dblSoma + .ListItems(intInd).SubItems(2)
        Next
    End With
    lblSoma = gstrConvVrDoSql(dblSoma)
End Sub

Private Sub txt_ValorParcela_GotFocus()
    MarcaCampo txt_ValorParcela
End Sub

Private Sub txt_ValorParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_ValorParcela
End Sub

Private Sub txt_ValorParcela_LostFocus()
    txt_ValorParcela = gstrConvVrDoSql(txt_ValorParcela)
End Sub

Private Sub txt_ValorRetencao_GotFocus()
    MarcaCampo txt_ValorRetencao
End Sub

Private Sub txt_ValorRetencao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_ValorRetencao
End Sub

Private Sub txt_ValorRetencao_LostFocus()
    txt_ValorRetencao = gstrConvVrDoSql(txt_ValorRetencao)
    If mblnAlterandoRetencao Then
        AtualizaListView lvw_Retencao, txt_ValorRetencao, lblRetencao
    Else
        lblRetencao = gstrConvVrDoSql(Val(gstrConvVrParaSql(lblRetencao)) + _
                                      Val(gstrConvVrParaSql(txt_ValorRetencao)))
    End If
End Sub

Private Function blnDeletouLancamento(bytTipo As Byte, strConta As String) As Boolean
    Dim strSQL  As String
    If gblnExclusaoGravacaoOk("", "Confirma Exclusão do lançamento?") Then
        strSQL = ""
        strSQL = strSQL & "DELETE " & gstrSubempenhoLiquidado & " "
        strSQL = strSQL & "WHERE PKID = " & lvw_Extra.SelectedItem.Tag & " "
       ' strSQL = strSQL & "AND intConta = " & Val(strConta) & " "
       ' strSQL = strSQL & "AND bytTipo = " & bytTipo
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSQL) Then
            blnDeletouLancamento = True
        End If
    End If
End Function

Private Function blnDeletouLancamentoOrcamentario() As Boolean
    Dim strSQL  As String
    If gblnExclusaoGravacaoOk("", "Confirma Exclusão do lançamento?") Then
        strSQL = ""
        strSQL = strSQL & "DELETE " & gstrSubEmpRetencaoOrcamentaria & " "
        strSQL = strSQL & "WHERE PKID = " & lvw_Orcamentario.SelectedItem.Tag & " "
       ' strSQL = strSQL & "AND intConta = " & Val(strConta) & " "
       ' strSQL = strSQL & "AND bytTipo = " & bytTipo
        Set gobjBanco = New clsBanco
        If gobjBanco.Execute(strSQL) Then
            blnDeletouLancamentoOrcamentario = True
        End If
    End If
End Function


Private Sub OrganizaNumParcelas()
Dim adoRec As ADODB.Recordset
Dim strSQL As String
Dim intCont As Integer

    Set gobjBanco = New clsBanco
    
    strSQL = ""
    strSQL = strSQL & " SELECT PKId FROM " & gstrSubempenho
    strSQL = strSQL & " WHERE intEmpenho = " & tdb_Lista.Columns(0)
    strSQL = strSQL & " AND intNumero > 0 "
    strSQL = strSQL & " ORDER BY intNumero"
    
    If gobjBanco.CriaADO(strSQL, 10, adoRec) And Not adoRec.EOF Then
        For intCont = 1 To adoRec.RecordCount - 1
            strSQL = ""
            strSQL = strSQL & " UPDATE " & gstrSubempenho & " SET "
            strSQL = strSQL & " intNumero=" & intCont & " WHERE PKId="
            strSQL = strSQL & adoRec("PKId") & " AND intNumero > 0"
            gobjBanco.Execute strSQL
            adoRec.MoveNext
        Next
    End If
End Sub

Private Function strTelefone() As String
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strTelefone = "''"
    
    
    If dbcintCredor.MatchedWithList Then
        strSQL = "Select " & gstrTOPnSQLServer(1)
        strSQL = strSQL & "A.strconteudo "
        strSQL = strSQL & "From "
        strSQL = strSQL & gstrEmpenho & " TA, "
        strSQL = strSQL & gstrFormaDeComunicacao & " A, "
        strSQL = strSQL & gstrTipoDeComunicacao & " B "
        strSQL = strSQL & "Where "
        strSQL = strSQL & "b.pkid = a.inttipodecomunicacao and "
        strSQL = strSQL & "b.inttipo in(6) and " 'Tipo Telefone Comercial
        strSQL = strSQL & "TA.intcredor = A.intcontribuinte AND "
        strSQL = strSQL & "TA.intcredor = " & Val(dbcintCredor.BoundText)
        
        strSQL = gstrTOPnOracle(strSQL, 1)
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
            If Not adoResultado.EOF Then
                strTelefone = "'" & gstrENulo(adoResultado!strConteudo) & "'"
                'gstrTelAutForn = "'" & gstrENulo(adoResultado!strConteudo) & "'"
            Else
                strTelefone = "'" & gstrENulo(adoResultado!strConteudo) & "'"
                'gstrTelAutForn = "'" & gstrENulo(adoResultado!strConteudo) & "'"
            End If
        Else
            strTelefone = "''"
            'gstrTelAutForn = "''"
        End If
    End If
End Function

Public Function strQueryRelatorio()

    If tdb_Lista.Columns(0).Value = "" Then
        ExibeMensagem "Não ha registro selecionado para ser impresso."
        Exit Function
    End If
      
    Dim strSQL  As String
    strSQL = ""
    
    If blnAlmoxarifado Then
        strSQL = "SELECT "
        
        strSQL = strSQL & "EP.intNumero intEmpenho, "
        strSQL = strSQL & "EP.PKID intEmpenhoPKID, "
        strSQL = strSQL & "'" & gstrMascaraItemDespesa & "' Mascara, "
        strSQL = strSQL & "'ALMOXARIFADO' strDestino, "
        strSQL = strSQL & "EP.strModalidade, "
        strSQL = strSQL & "RC.intPedidoEmpenho PedidoEmpenho, "
        strSQL = strSQL & "LO.strDescricao UnidadeCC, "
        strSQL = strSQL & "AC.strobjetoautorizacao, "
        strSQL = strSQL & strTelefone & " telefone, "
        strSQL = strSQL & "EP.strContrato, "
        strSQL = strSQL & "EP.strSolicitacao, "
        strSQL = strSQL & "EP.strCodigo, "
        strSQL = strSQL & "EP.intExercicio, "
        strSQL = strSQL & "EP.bitDigito, "
        strSQL = strSQL & "EP.dblValor dblValorEmpenho, "
        
        strSQL = strSQL & "RD.intNumero NumeroReserva, "
        strSQL = strSQL & "RD.intExercicioReserva , "
        strSQL = strSQL & "RD.strSolicitacao SolicitacaoReserva, "
        strSQL = strSQL & "RD.intExercicio ExercicioSolicitacaoReserva, "
        
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "EP.PKID") & strCONCAT & "'/1' Agrupamento, "
        
        ' Este campo é usado para ordenar os registros e para servir de grupo a relatórios externos que
        ' não exibam os subempenhos
        
        strSQL = strSQL & "EP.Strcondpagto, "
        strSQL = strSQL & "EP.Strlocentrega, "
        strSQL = strSQL & "EP.Strprazoentrega, "
        strSQL = strSQL & "EP.dtmdata DataEmpenho, "
        strSQL = strSQL & "PT.PKID CodigoProgramaTrabalho, "
        strSQL = strSQL & "PT.strCodigo strDotacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "PT.intCodigoReduzido, "
        strSQL = strSQL & "PJ.strcodigo CodProjetoAtividade, "
        strSQL = strSQL & "PT.intPrograma CodPrograma, "
        strSQL = strSQL & "PT.intSubFuncao CodSubFuncao, "
        strSQL = strSQL & "PT.intFuncao CodFuncao, "
        strSQL = strSQL & "PT.dblValor dblProgramaTrabalho, "
        strSQL = strSQL & "PT.intProjetoAtividade intProjetoAtividade, "
        strSQL = strSQL & "OG.strCodigo CodOrgao, "
        strSQL = strSQL & "OG.strDescricao strOrgao, "
        strSQL = strSQL & "UO.strCodigo CodUnidadeOrcamentaria, "
        strSQL = strSQL & "UO.strDescricao UnidadeOrcamentaria, "
        strSQL = strSQL & "ED.strCodigoElementoDespesa CodElementoDespesa, "
        strSQL = strSQL & "ED.strDescricao ElementoDespesa, "
        strSQL = strSQL & "CL.strCodigo CodCompraLicitacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "FR.strCodigo CodigoFonteRecurso, "
        strSQL = strSQL & "FR.strDescricao strFonteRecurso, "
        strSQL = strSQL & "CV.strCodigo CodigoConvenio, "
        strSQL = strSQL & "CV.strDescricao strConvenio, "
        strSQL = strSQL & "CT.CDC intCodigoContribuinte, "
        strSQL = strSQL & "CT.strNome, "
        strSQL = strSQL & "CT.strCNPJCPF, "
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "LG.STRDESCRICAO") & " strEndereco, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.StrlogradouroD", "CT.Strlogradouroc"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intNumero") & " INTNUMERO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intNumeroD", "CT.intNumeroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.strComplemento") & " STRCOMPLEMENTO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strComplementoD", "CT.strComplementoC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "MP.strDescricao") & " STRMUNICIPIO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strMunicipioD", "MPC.strDescricao"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "UF.strSigla") & " STRUF, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',UFD.strSigla", "UFC.strSigla"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "BR.strDescricao") & " STRBAIRRO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strbairroD", "CT.strbairroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intCEP") & " INTCEP, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intCEPD", "CT.intCEPC"))
        
        strSQL = strSQL & "CT.Strlogradouroc,"
        strSQL = strSQL & "CT.intNumeroC,"
        strSQL = strSQL & "CT.strComplementoC,"
        strSQL = strSQL & "MPC.strDescricao strMunicipioC,"
        strSQL = strSQL & "UFC.strSigla strUFC,"
        strSQL = strSQL & "CT.strBairroC,"
        strSQL = strSQL & "CT.intCEPC,"
        strSQL = strSQL & "SEP.PKID intPKIDParcela, "
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "SEP.PKID") & strCONCAT & "'/1' AgrupamentoParcela, "
        strSQL = strSQL & "SEP.intNumero intParcela, "
        '09/03/05 - Incluido tipo de anulacao p/ situacao 4
        strSQL = strSQL & "SEP.bytSituacao, SEP.bytTipo, "
      
        If blnSoEstorno Then
           strSQL = strSQL & "'NOTA DE ESTORNO' Nota, "
           strSQL = strSQL & "SEP.strHistorico, "
        Else
           strSQL = strSQL & "'NOTA DE EMPENHO' Nota, "
           strSQL = strSQL & gstrCASEWHEN("SEP.intNumero", "0, EP.strHistorico", "SEP.strHistorico") & " strHistorico, "
        End If
      
        strSQL = strSQL & "SEP.dtmData, "
        strSQL = strSQL & "SEP.intEmpenhoAnulacao, "
        strSQL = strSQL & "SEP.dblValor, "
        strSQL = strSQL & gstrISNULL("SEP.dblEmpenhadoAteData", "0") & " AS dblEmpenhadoAteData, "
        strSQL = strSQL & gstrISNULL("SEP.dblSaldoAtual", "0") & " AS dblSaldoAtual, "
        strSQL = strSQL & "CL.strDescricao strModalidadeLicitacao, "
        strSQL = strSQL & "IT.intcodigoitem AS CodItem, "
        strSQL = strSQL & "IT.Strdescricaoitem AS DescItem, "
        strSQL = strSQL & "IT.Strmarca AS MarcaItem, "
        strSQL = strSQL & "IT.STRUNIDADE as UnidItem, "
        strSQL = strSQL & "IT.DBLQUANTIDADE AS QuantItem, "
        strSQL = strSQL & "IT.Dblprecounitario AS PrecoItem, "
        strSQL = strSQL & "(IT.DBLQUANTIDADE * IT.Dblprecounitario) AS PrecoTotItem, "
        'Incluida a Data de Vencimento Ficha Orc1136 - Fernando
        strSQL = strSQL & "Sep.DtmVencimento DtmVencimento "
                    
        If bytDBType = EDatabases.Oracle Then
            strSQL = strSQL & "FROM " & gstrEmpenho & " EP,"
            strSQL = strSQL & gstrProgramaDeTrabalho & " PT,"
            strSQL = strSQL & gstrOrgao & " OG,"
            strSQL = strSQL & gstrProjeto & " PJ,"
            strSQL = strSQL & gstrFonteRecurso & " FR,"
            strSQL = strSQL & gstrConvenio & " CV,"
            strSQL = strSQL & gstrEmpenhoContrato & " PE, "
            strSQL = strSQL & gstrLocais & " LO, "
            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC, "
            strSQL = strSQL & gstrContribuinte & " CT,"
            strSQL = strSQL & gstrCidade & " MP,"
            strSQL = strSQL & gstrUF & " UF,"
            strSQL = strSQL & gstrBairro & " BR,"
            strSQL = strSQL & gstrSubempenho & " SEP,"
            strSQL = strSQL & gstrComprasLicitacao & " CL, "
            strSQL = strSQL & gstrUnidadeOrcamentaria & " UO, "
            strSQL = strSQL & gstrElementoDespesa & " ED, "
            If blnSoEstorno Then
               strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT ,"
            Else
               strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT , "
               'strSQL = strSQL & gstrItemEmpenho & " IT "
            End If
            
            strSQL = strSQL & gstrLogradouro & " LG,"
            strSQL = strSQL & gstrCidade & " MPC,"
            strSQL = strSQL & gstrUF & " UFC,"
            strSQL = strSQL & gstrUF & " UFD,"
            strSQL = strSQL & gstrReservaDotacao & " RD, "
            strSQL = strSQL & gstrRequisicaoCompras & " RC "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
                
                
            If blnSoEstorno Then
               strSQL = strSQL & "SEP.Pkid " & strOUTJSQLServer & "=IT.intSubEmpenho " & strOUTJOracle & " AND "
            Else
               strSQL = strSQL & "EP.Pkid " & strOUTJSQLServer & "= IT.Intempenho " & strOUTJOracle & " AND "
            End If
            
            strSQL = strSQL & "EP.PKID = SEP.intEmpenho AND "
            strSQL = strSQL & "EP.intProgramaTrabalho = PT.PKID AND "
            strSQL = strSQL & "CL.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intModalidade   AND "
            strSQL = strSQL & "PT.intOrgao = OG.pkid AND "
            strSQL = strSQL & "PT.intProjetoAtividade = PJ.PKID AND "
            strSQL = strSQL & "PT.intFonteRecurso = FR.PKID AND "
            strSQL = strSQL & "CV.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intConvenio AND "
            strSQL = strSQL & "PE.intPedidoEmpenho " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.Pkid AND "
            strSQL = strSQL & "RC.PKId " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intRequisicaodeCompra AND "
            strSQL = strSQL & "LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " RC.intCodigoCentroDeCusto2 AND "
            strSQL = strSQL & "AC.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intautorizacaodecompra AND "
            strSQL = strSQL & "EP.intCredor = CT.PKID AND "
            strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
            strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
            strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
            strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
            strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
            strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
            strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ") & " AND "
            
            strSQL = strSQL & "PT.intUnidadeOrcamentaria = UO.PKID AND "
            strSQL = strSQL & "PT.intElementoDespesa = ED.PKID"
            
            strSQL = strSQL & " AND LG.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intLogradouro "
            strSQL = strSQL & " AND MPC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipioC "
            strSQL = strSQL & " AND UFC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFC "
            strSQL = strSQL & " AND UFD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFD "
            strSQL = strSQL & " AND RD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intReservaDotacao "
            
        Else

            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrEmpenho & " EP "
            strSQL = strSQL & "INNER JOIN " & gstrSubempenho & " SEP ON (EP.PKID = SEP.intEmpenho) "
            strSQL = strSQL & "INNER JOIN " & gstrProgramaDeTrabalho & " PT ON (EP.intProgramaTrabalho = PT.PKID) "
            strSQL = strSQL & "LEFT JOIN " & gstrComprasLicitacao & " CL ON (CL.PKID  = EP.intModalidade) "
            strSQL = strSQL & "INNER JOIN " & gstrOrgao & " OG ON (PT.intOrgao = OG.PKID ) "
            strSQL = strSQL & "INNER JOIN " & gstrProjeto & " PJ ON (PT.intProjetoAtividade = PJ.PKID) "
            strSQL = strSQL & "INNER JOIN " & gstrFonteRecurso & " FR ON (PT.intFonteRecurso = FR.PKID) "
            strSQL = strSQL & "LEFT JOIN "
            strSQL = strSQL & gstrEmpenhoContrato & " PE ON (PE.intPedidoEmpenho = EP.PKId) LEFT JOIN "
            strSQL = strSQL & gstrRequisicaoCompras & " RC ON (RC.PKId = PE.intRequisicaodeCompra) LEFT JOIN "
            strSQL = strSQL & gstrLocais & " LO ON (LO.PKId = RC.intCodigoCentroDeCusto2) LEFT JOIN "
            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC ON (AC.PKId = PE.intAutorizacaodeCompra) LEFT JOIN "
            strSQL = strSQL & gstrConvenio & " CV ON (CV.PKID  = EP.intConvenio) "
            strSQL = strSQL & "INNER JOIN " & gstrContribuinte & " CT ON (EP.intCredor = CT.PKID) "
            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MP ON (MP.PKID  = CT.intMunicipio) "
            strSQL = strSQL & "LEFT JOIN " & gstrBairro & " BR ON (BR.PKID  = CT.intBairro) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UF ON (UF.PKID  = CT.intUF) "
            strSQL = strSQL & "LEFT JOIN " & gstrLogradouro & " LG ON (LG.PKID  = CT.intLogradouro) "
            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MPC ON (MPC.PKID  = CT.intMunicipioC) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFC ON (UFC.PKID  = CT.intUFC) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFD ON (UFD.PKID  = CT.intUFD) "
            strSQL = strSQL & "INNER JOIN " & gstrUnidadeOrcamentaria & " UO ON (PT.intUnidadeOrcamentaria = UO.PKID) "
            strSQL = strSQL & "INNER JOIN " & gstrElementoDespesa & " ED ON (PT.intElementoDespesa = ED.PKID) "
'----------------------------------------------------------------------------------------------------------------
             
'             strSQL = strSQL & " FROM "
'             strSQL = strSQL & gstrEmpenho & "EP "
'             strSQL = strSQL & "INNER JOIN " & gstrSubempenho & " SEP ON (EP.PKID = SEP.intEmpenho) "
'             strSQL = strSQL & "INNER JOIN " & gstrProgramaDeTrabalho & " PT ON (EP.intProgramaTrabalho = PT.PKID) "
'             strSQL = strSQL & "LEFT JOIN "
'             strSQL = strSQL & gstrComprasLicitacao & " CL "
'             strSQL = strSQL & "ON(CL.PKID = EP.intModalidade) "
'             strSQL = strSQL & "INNER JOIN " & gstrOrgao & " OG ON (PT.intOrgao = OG.PKID) "
'             strSQL = strSQL & "INNER JOIN " & gstrProjeto & " PJ ON (PT.intProjetoAtividade = PJ.PKID) "
'             strSQL = strSQL & "INNER JOIN " & gstrFonteRecurso & " FR ON (PT.intFonteRecurso = FR.PKID) LEFT JOIN "
'             strSQL = strSQL & gstrRequisicaoCompras & " RC "
'             strSQL = strSQL & "ON PT.pkid = RC.intProgramadeTrabalho INNER JOIN "
'             strSQL = strSQL & gstrEmpenhoContrato & " PE "
'             strSQL = strSQL & "ON PE.intPedidoEmpenho = EP.PKId AND RC.PKID = PE. INTREQUISICAODECOMPRA LEFT JOIN "
'             strSQL = strSQL & gstrLocais & " LO "
'             strSQL = strSQL & "ON(LO.PKId = RC.intCodigoCentroDeCusto2) LEFT JOIN "
'             strSQL = strSQL & gstrAutorizacaoDeCompra & " AC "
'             strSQL = strSQL & " ON(AC.PKId = PE.intAutorizacaodeCompra) LEFT JOIN "
'             strSQL = strSQL & gstrConvenio & " CV "
'             strSQL = strSQL & "ON(CV.PKID = EP.intConvenio) "
'             strSQL = strSQL & "INNER JOIN " & gstrContribuinte & " CT ON (EP.intCredor = CT.PKID) LEFT JOIN "
'             strSQL = strSQL & gstrCidade & " MP "
'             strSQL = strSQL & "ON(MP.PKID = CT.intMunicipio) LEFT JOIN "
'             strSQL = strSQL & gstrBairro & " BR ON(BR.PKID = CT.intBairro) LEFT JOIN "
'             strSQL = strSQL & gstrUF & " UF ON(UF.PKID = CT.intUF) LEFT JOIN "
'             strSQL = strSQL & gstrLogradouro & " LG "
'             strSQL = strSQL & "ON(LG.PKID = CT.intLogradouro) LEFT JOIN "
'             strSQL = strSQL & gstrCidade & " MPC "
'             strSQL = strSQL & "ON(MPC.PKID = CT.intMunicipioC) LEFT JOIN "
'             strSQL = strSQL & gstrUF & " UFC ON(UFC.PKID = CT.intUFC) LEFT JOIN "
'             strSQL = strSQL & gstrUF & " UFD ON(UFD.PKID = CT.intUFD) "
'             strSQL = strSQL & "INNER JOIN " & gstrUnidadeOrcamentaria & " UO ON (PT.intUnidadeOrcamentaria = UO.PKID) "
'             strSQL = strSQL & "INNER JOIN " & gstrElementoDespesa & " ED ON (PT.intElementoDespesa = ED.PKID) LEFT JOIN"

            strSQL = strSQL & "LEFT JOIN "
            
            If blnSoEstorno Then
                strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT "
            Else
                strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT  "
            End If
             
            If blnSoEstorno Then
                strSQL = strSQL & "ON (SEP.Pkid = IT.intSubEmpenho )"
            Else
                strSQL = strSQL & "ON (EP.Pkid = IT.Intempenho )"
            End If

            strSQL = strSQL & "LEFT JOIN " & gstrReservaDotacao & " RD ON (EP.intReservaDotacao  = RD.pkid) "
            
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
            strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
            strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
            strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
            strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ")
        End If
        
        strSQL = strSQL & AdicionaGroupByQueryRelatorio
    
    End If
'--------------------------------------------------------------------
    If blnCompras Then
        If strSQL = "" Then
            strSQL = "SELECT "
        Else
            strSQL = strSQL & "UNION ALL SELECT "
        End If
    
        strSQL = strSQL & "EP.intNumero intEmpenho, "
        strSQL = strSQL & "EP.PKID intEmpenhoPKID, "
        strSQL = strSQL & "'" & gstrMascaraItemDespesa & "' Mascara, "
        strSQL = strSQL & "'COMPRAS' strDestino, "
        strSQL = strSQL & "EP.strModalidade, "
        strSQL = strSQL & "RC.intPedidoEmpenho PedidoEmpenho, "
        strSQL = strSQL & "LO.strDescricao UnidadeCC, "
        strSQL = strSQL & "AC.strobjetoautorizacao, "
        strSQL = strSQL & strTelefone & " telefone, "
        strSQL = strSQL & "EP.strContrato, "
        strSQL = strSQL & "EP.strSolicitacao, "
        strSQL = strSQL & "EP.strCodigo, "
        strSQL = strSQL & "EP.intExercicio, "
        strSQL = strSQL & "EP.bitDigito, "
        strSQL = strSQL & "EP.dblValor  dblValorEmpenho, "
        strSQL = strSQL & "RD.intNumero NumeroReserva, "
        strSQL = strSQL & "RD.intExercicioReserva , "
        strSQL = strSQL & "RD.strSolicitacao SolicitacaoReserva, "
        strSQL = strSQL & "RD.intExercicio ExercicioSolicitacaoReserva, "
        
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "EP.PKID") & strCONCAT & "'/2' Agrupamento, "
        ' Este campo é usado para ordenar os registros e para servir de grupo a relatórios externos que
        ' não exibam os subempenhos
      
        strSQL = strSQL & "EP.Strcondpagto, "
        strSQL = strSQL & "EP.Strlocentrega, "
        strSQL = strSQL & "EP.Strprazoentrega, "
        strSQL = strSQL & "EP.dtmdata DataEmpenho, "
        strSQL = strSQL & "PT.PKID CodigoProgramaTrabalho, "
        strSQL = strSQL & "PT.strCodigo strDotacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "PT.intCodigoReduzido, "
        strSQL = strSQL & "PJ.strcodigo CodProjetoAtividade, "
        strSQL = strSQL & "PT.intPrograma CodPrograma, "
        strSQL = strSQL & "PT.intSubFuncao CodSubFuncao, "
        strSQL = strSQL & "PT.intFuncao CodFuncao, "
        strSQL = strSQL & "PT.dblValor dblProgramaTrabalho, "
        strSQL = strSQL & "PT.intProjetoAtividade intProjetoAtividade, "
        strSQL = strSQL & "OG.strCodigo CodOrgao, "
        strSQL = strSQL & "OG.strDescricao strOrgao, "
        strSQL = strSQL & "UO.strCodigo CodUnidadeOrcamentaria, "
        strSQL = strSQL & "UO.strDescricao UnidadeOrcamentaria, "
        strSQL = strSQL & "ED.strCodigoElementoDespesa CodElementoDespesa, "
        strSQL = strSQL & "ED.strDescricao ElementoDespesa, "
        strSQL = strSQL & "CL.strCodigo CodCompraLicitacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "FR.strCodigo CodigoFonteRecurso, "
        strSQL = strSQL & "FR.strDescricao strFonteRecurso, "
        strSQL = strSQL & "CV.strCodigo CodigoConvenio, "
        strSQL = strSQL & "CV.strDescricao strConvenio, "
        strSQL = strSQL & "CT.CDC intCodigoContribuinte, "
        strSQL = strSQL & "CT.strNome, "
        strSQL = strSQL & "CT.strCNPJCPF, "
      
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "LG.STRDESCRICAO") & " strEndereco, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.StrlogradouroD", "CT.Strlogradouroc"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intNumero") & " INTNUMERO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intNumeroD", "CT.intNumeroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.strComplemento") & " STRCOMPLEMENTO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strComplementoD", "CT.strComplementoC"))
              
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "MP.strDescricao") & " STRMUNICIPIO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strMunicipioD", "MPC.strDescricao"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "UF.strSigla") & " STRUF, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',UFD.strSigla", "UFC.strSigla"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "BR.strDescricao") & " STRBAIRRO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strbairroD", "CT.strbairroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intCEP") & " INTCEP, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intCEPD", "CT.intCEPC"))
        
        strSQL = strSQL & "CT.Strlogradouroc,"
        strSQL = strSQL & "CT.intNumeroC,"
        strSQL = strSQL & "CT.strComplementoC,"
        strSQL = strSQL & "MPC.strDescricao strMunicipioC,"
        strSQL = strSQL & "UFC.strSigla strUFC,"
        strSQL = strSQL & "CT.strBairroC,"
        strSQL = strSQL & "CT.intCEPC,"
        strSQL = strSQL & "SEP.PKID intPKIDParcela, "
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "SEP.PKID") & strCONCAT & "'/2' AgrupamentoParcela, "
        strSQL = strSQL & "SEP.intNumero intParcela, "
        strSQL = strSQL & "SEP.bytSituacao, "
        strSQL = strSQL & "SEP.bytTipo, "
        
        If blnSoEstorno Then
           strSQL = strSQL & "'NOTA DE ESTORNO' Nota, "
           strSQL = strSQL & "SEP.strHistorico, "
        Else
           strSQL = strSQL & "'NOTA DE EMPENHO' Nota, "
           strSQL = strSQL & gstrCASEWHEN("SEP.intNumero", "0, EP.strHistorico", "SEP.strHistorico") & " strHistorico, "
        End If
      
        strSQL = strSQL & "SEP.dtmData, "
        strSQL = strSQL & "SEP.intEmpenhoAnulacao, "
        strSQL = strSQL & "SEP.dblValor, "
        strSQL = strSQL & gstrISNULL("SEP.dblEmpenhadoAteData", "0") & " AS dblEmpenhadoAteData, "
        strSQL = strSQL & gstrISNULL("SEP.dblSaldoAtual", "0") & " AS dblSaldoAtual, "
        strSQL = strSQL & "CL.strDescricao strModalidadeLicitacao, "
        strSQL = strSQL & "IT.intcodigoitem AS CodItem,"
        strSQL = strSQL & "IT.Strdescricaoitem AS DescItem,"
        strSQL = strSQL & "IT.Strmarca AS MarcaItem,"
        strSQL = strSQL & "IT.STRUNIDADE as UnidItem,"
        strSQL = strSQL & "IT.DBLQUANTIDADE AS QuantItem,"
        strSQL = strSQL & "IT.Dblprecounitario AS PrecoItem,"
        strSQL = strSQL & "(IT.DBLQUANTIDADE * IT.Dblprecounitario) AS PrecoTotItem, "
        'Incluida a Data de Vencimento Ficha Orc1136 - Fernando
        strSQL = strSQL & "Sep.DtmVencimento DtmVencimento "
        
        If bytDBType = EDatabases.Oracle Then
            strSQL = strSQL & "FROM " & gstrEmpenho & " EP,"
            strSQL = strSQL & gstrProgramaDeTrabalho & " PT,"
            strSQL = strSQL & gstrOrgao & " OG,"
            strSQL = strSQL & gstrProjeto & " PJ,"
            strSQL = strSQL & gstrFonteRecurso & " FR,"
            strSQL = strSQL & gstrConvenio & " CV,"
            strSQL = strSQL & gstrEmpenhoContrato & " PE, " 'ON (PE.intNumeroEmpenho = EP.PKId) INNER JOIN "
            strSQL = strSQL & gstrLocais & " LO, " 'ON (LO.PKId = PE.intUnidadeCentroDeCusto) INNER JOIN "
            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC, " 'ON (AC.PKId = PE.intautorizacaodecompra) LEFT JOIN "
            strSQL = strSQL & gstrContribuinte & " CT,"
            strSQL = strSQL & gstrCidade & " MP,"
            strSQL = strSQL & gstrUF & " UF,"
            strSQL = strSQL & gstrBairro & " BR,"
            strSQL = strSQL & gstrSubempenho & " SEP,"
            strSQL = strSQL & gstrComprasLicitacao & " CL, "
            strSQL = strSQL & gstrUnidadeOrcamentaria & " UO, "
            strSQL = strSQL & gstrElementoDespesa & " ED, "
            If blnSoEstorno Then
                strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT ,"
            Else
                strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT , "
            End If
            
            strSQL = strSQL & gstrLogradouro & " LG,"
            strSQL = strSQL & gstrCidade & " MPC,"
            strSQL = strSQL & gstrUF & " UFC,"
            strSQL = strSQL & gstrUF & " UFD, "
            strSQL = strSQL & gstrReservaDotacao & " RD, "
            strSQL = strSQL & gstrRequisicaoCompras & " RC "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
           
            If blnSoEstorno Then
                strSQL = strSQL & "SEP.Pkid " & strOUTJSQLServer & "=IT.intSubEmpenho " & strOUTJOracle & " AND "
            Else
                strSQL = strSQL & "EP.Pkid " & strOUTJSQLServer & "= IT.Intempenho " & strOUTJOracle & " AND "
            End If
           
            strSQL = strSQL & "EP.PKID = SEP.intEmpenho AND "
            strSQL = strSQL & "EP.intProgramaTrabalho = PT.PKID AND "
            strSQL = strSQL & "CL.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intModalidade   AND "
            strSQL = strSQL & "PT.intOrgao = OG.PKID AND "
            strSQL = strSQL & "PT.intProjetoAtividade = PJ.PKID AND "
            strSQL = strSQL & "PT.intFonteRecurso = FR.PKID AND "
            strSQL = strSQL & "CV.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intConvenio AND "
            strSQL = strSQL & "PE.intPedidoEmpenho " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.Pkid AND "
            strSQL = strSQL & "RC.PKId " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intRequisicaodeCompra AND "
            strSQL = strSQL & "LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " RC.intCodigoCentroDeCusto2 AND "
            strSQL = strSQL & "AC.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intautorizacaodecompra AND "
            strSQL = strSQL & "EP.intCredor = CT.PKID AND "
            strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
            strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
            strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
            strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
            strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
            strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
            strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ") & " AND "
            strSQL = strSQL & "PT.intUnidadeOrcamentaria = UO.PKID AND "
            strSQL = strSQL & "PT.intElementoDespesa = ED.PKID"
            strSQL = strSQL & " AND LG.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intLogradouro "
            strSQL = strSQL & " AND MPC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipioC "
            strSQL = strSQL & " AND UFC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFC "
            strSQL = strSQL & " AND UFD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFD "
            strSQL = strSQL & " AND RD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intReservaDotacao "
         Else
            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrEmpenho & " EP "
            strSQL = strSQL & "INNER JOIN " & gstrSubempenho & " SEP ON (EP.PKID = SEP.intEmpenho) "
            strSQL = strSQL & "INNER JOIN " & gstrProgramaDeTrabalho & " PT ON (EP.intProgramaTrabalho = PT.PKID) "
            strSQL = strSQL & "LEFT JOIN " & gstrComprasLicitacao & " CL ON (CL.PKID  = EP.intModalidade) "
            strSQL = strSQL & "INNER JOIN " & gstrOrgao & " OG ON (PT.intOrgao = OG.PKID ) "
            strSQL = strSQL & "INNER JOIN " & gstrProjeto & " PJ ON (PT.intProjetoAtividade = PJ.PKID) "
            strSQL = strSQL & "INNER JOIN " & gstrFonteRecurso & " FR ON (PT.intFonteRecurso = FR.PKID) "
            strSQL = strSQL & "LEFT JOIN "
            strSQL = strSQL & gstrEmpenhoContrato & " PE ON (PE.intPedidoEmpenho = EP.PKId) LEFT JOIN "
            strSQL = strSQL & gstrRequisicaoCompras & " RC ON (RC.PKId = PE.intRequisicaodeCompra) LEFT JOIN "
            strSQL = strSQL & gstrLocais & " LO ON (LO.PKId = RC.intCodigoCentroDeCusto2) LEFT JOIN "
            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC ON (AC.PKId = PE.intAutorizacaodeCompra) LEFT JOIN "
            strSQL = strSQL & gstrConvenio & " CV ON (CV.PKID  = EP.intConvenio) "
            strSQL = strSQL & "INNER JOIN " & gstrContribuinte & " CT ON (EP.intCredor = CT.PKID) "
            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MP ON (MP.PKID  = CT.intMunicipio) "
            strSQL = strSQL & "LEFT JOIN " & gstrBairro & " BR ON (BR.PKID  = CT.intBairro) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UF ON (UF.PKID  = CT.intUF) "
            strSQL = strSQL & "LEFT JOIN " & gstrLogradouro & " LG ON (LG.PKID  = CT.intLogradouro) "
            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MPC ON (MPC.PKID  = CT.intMunicipioC) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFC ON (UFC.PKID  = CT.intUFC) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFD ON (UFD.PKID  = CT.intUFD) "
            strSQL = strSQL & "INNER JOIN " & gstrUnidadeOrcamentaria & " UO ON (PT.intUnidadeOrcamentaria = UO.PKID) "
            strSQL = strSQL & "INNER JOIN " & gstrElementoDespesa & " ED ON (PT.intElementoDespesa = ED.PKID) "
           
            strSQL = strSQL & "LEFT JOIN "
            
            If blnSoEstorno Then
                strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT "
            Else
                strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT  "
            End If
             
            If blnSoEstorno Then
                strSQL = strSQL & "ON (SEP.Pkid = IT.intSubEmpenho ) "
            Else
                strSQL = strSQL & "ON (EP.Pkid = IT.Intempenho ) "
            End If
            strSQL = strSQL & "LEFT JOIN " & gstrReservaDotacao & " RD ON (EP.intReservaDotacao  = RD.pkid) "
            
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
            strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
            strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
            strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
            strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ")
            
        End If
        
        strSQL = strSQL & AdicionaGroupByQueryRelatorio
        
    End If
'----------------------------------------------------------------------
    If blnFornecedor Then
        If strSQL = "" Then
            strSQL = "SELECT "
        Else
            strSQL = strSQL & "UNION ALL SELECT "
        End If
      
        strSQL = strSQL & "EP.intNumero intEmpenho, "
        strSQL = strSQL & "EP.PKID intEmpenhoPKID, "
        strSQL = strSQL & "'" & gstrMascaraItemDespesa & "' Mascara, "
        strSQL = strSQL & "'FORNECEDOR' strDestino, "
        strSQL = strSQL & "EP.strModalidade, "
        strSQL = strSQL & "RC.intPedidoEmpenho PedidoEmpenho, "
        strSQL = strSQL & "LO.strDescricao UnidadeCC, "
        strSQL = strSQL & "AC.strobjetoautorizacao, "
        strSQL = strSQL & strTelefone & " telefone, "
        strSQL = strSQL & "EP.strContrato, "
        strSQL = strSQL & "EP.strSolicitacao, "
        strSQL = strSQL & "EP.strCodigo, "
        strSQL = strSQL & "EP.intExercicio, "
        strSQL = strSQL & "EP.bitDigito, "
        
        strSQL = strSQL & "EP.dblValor dblValorEmpenho, "
        strSQL = strSQL & "RD.intNumero NumeroReserva, "
        strSQL = strSQL & "RD.intExercicioReserva , "
        strSQL = strSQL & "RD.strSolicitacao SolicitacaoReserva, "
        strSQL = strSQL & "RD.intExercicio ExercicioSolicitacaoReserva, "
        
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "EP.PKID") & strCONCAT & "'/3' Agrupamento, "
        ' Este campo é usado para ordenar os registros e para servir de grupo a relatórios externos que
        ' não exibam os subempenhos
        
        strSQL = strSQL & "EP.Strcondpagto, "
        strSQL = strSQL & "EP.Strlocentrega, "
        strSQL = strSQL & "EP.Strprazoentrega, "
        strSQL = strSQL & "EP.dtmdata DataEmpenho, "
        strSQL = strSQL & "PT.PKID CodigoProgramaTrabalho, "
        strSQL = strSQL & "PT.strCodigo strDotacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "PT.intCodigoReduzido, "
        strSQL = strSQL & "PJ.strcodigo CodProjetoAtividade, "
        strSQL = strSQL & "PT.intPrograma CodPrograma, "
        strSQL = strSQL & "PT.intSubFuncao CodSubFuncao, "
        strSQL = strSQL & "PT.intFuncao CodFuncao, "
        strSQL = strSQL & "PT.dblValor dblProgramaTrabalho, "
        strSQL = strSQL & "PT.intProjetoAtividade intProjetoAtividade, "
        strSQL = strSQL & "OG.strCodigo CodOrgao, "
        strSQL = strSQL & "OG.strDescricao strOrgao, "
        strSQL = strSQL & "UO.strCodigo CodUnidadeOrcamentaria, "
        strSQL = strSQL & "UO.strDescricao UnidadeOrcamentaria, "
        strSQL = strSQL & "ED.strCodigoElementoDespesa CodElementoDespesa, "
        strSQL = strSQL & "ED.strDescricao ElementoDespesa, "
        strSQL = strSQL & "CL.strCodigo CodCompraLicitacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "FR.strCodigo CodigoFonteRecurso, "
        strSQL = strSQL & "FR.strDescricao strFonteRecurso, "
        strSQL = strSQL & "CV.strCodigo CodigoConvenio, "
        strSQL = strSQL & "CV.strDescricao strConvenio, "
        strSQL = strSQL & "CT.CDC intCodigoContribuinte, "
        strSQL = strSQL & "CT.strNome, "
        strSQL = strSQL & "CT.strCNPJCPF, "
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "LG.STRDESCRICAO") & " strEndereco, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.StrlogradouroD", "CT.Strlogradouroc"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intNumero") & " INTNUMERO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intNumeroD", "CT.intNumeroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.strComplemento") & " STRCOMPLEMENTO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strComplementoD", "CT.strComplementoC"))
              
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "MP.strDescricao") & " STRMUNICIPIO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strMunicipioD", "MPC.strDescricao"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "UF.strSigla") & " STRUF, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',UFD.strSigla", "UFC.strSigla"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "BR.strDescricao") & " STRBAIRRO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strbairroD", "CT.strbairroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intCEP") & " INTCEP, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intCEPD", "CT.intCEPC"))
        
        strSQL = strSQL & "CT.Strlogradouroc,"
        strSQL = strSQL & "CT.intNumeroC,"
        strSQL = strSQL & "CT.strComplementoC,"
        strSQL = strSQL & "MPC.strDescricao strMunicipioC,"
        strSQL = strSQL & "UFC.strSigla strUFC,"
        strSQL = strSQL & "CT.strBairroC,"
        strSQL = strSQL & "CT.intCEPC,"
        strSQL = strSQL & "SEP.PKID intPKIDParcela, "
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "SEP.PKID") & strCONCAT & "'/3' AgrupamentoParcela, "
        strSQL = strSQL & "SEP.intNumero intParcela, "
        strSQL = strSQL & "SEP.bytSituacao, "
        strSQL = strSQL & "SEP.bytTipo, "
        
        If blnSoEstorno Then
           strSQL = strSQL & "'NOTA DE ESTORNO' Nota, "
           strSQL = strSQL & "SEP.strHistorico, "
        Else
           strSQL = strSQL & "'NOTA DE EMPENHO' Nota, "
           strSQL = strSQL & gstrCASEWHEN("SEP.intNumero", "0, EP.strHistorico", "SEP.strHistorico") & " strHistorico, "
        End If
      
        strSQL = strSQL & "SEP.dtmData, "
        strSQL = strSQL & "SEP.intEmpenhoAnulacao, "
        strSQL = strSQL & "SEP.dblValor, "
        strSQL = strSQL & gstrISNULL("SEP.dblEmpenhadoAteData", "0") & " AS dblEmpenhadoAteData, "
        strSQL = strSQL & gstrISNULL("SEP.dblSaldoAtual", "0") & " AS dblSaldoAtual, "
        strSQL = strSQL & "CL.strDescricao strModalidadeLicitacao, "
        strSQL = strSQL & "IT.intcodigoitem AS CodItem,"
        strSQL = strSQL & "IT.Strdescricaoitem AS DescItem,"
        strSQL = strSQL & "IT.Strmarca AS MarcaItem,"
        strSQL = strSQL & "IT.STRUNIDADE as UnidItem,"
        strSQL = strSQL & "IT.DBLQUANTIDADE AS QuantItem,"
        strSQL = strSQL & "IT.Dblprecounitario AS PrecoItem,"
        strSQL = strSQL & "(IT.DBLQUANTIDADE * IT.Dblprecounitario) AS PrecoTotItem, "
        'Incluida a Data de Vencimento Ficha Orc1136 - Fernando
        strSQL = strSQL & "Sep.DtmVencimento DtmVencimento "
        
  
        If bytDBType = EDatabases.Oracle Then
            strSQL = strSQL & "FROM " & gstrEmpenho & " EP,"
            strSQL = strSQL & gstrProgramaDeTrabalho & " PT,"
            strSQL = strSQL & gstrOrgao & " OG,"
            strSQL = strSQL & gstrProjeto & " PJ,"
            strSQL = strSQL & gstrFonteRecurso & " FR,"
            strSQL = strSQL & gstrConvenio & " CV,"
            strSQL = strSQL & gstrEmpenhoContrato & " PE, " 'ON (PE.intNumeroEmpenho = EP.PKId) INNER JOIN "
            strSQL = strSQL & gstrLocais & " LO, " 'ON (LO.PKId = PE.intUnidadeCentroDeCusto) INNER JOIN "
            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC, " 'ON (AC.PKId = PE.intautorizacaodecompra) LEFT JOIN "
            strSQL = strSQL & gstrContribuinte & " CT,"
            strSQL = strSQL & gstrCidade & " MP,"
            strSQL = strSQL & gstrUF & " UF,"
            strSQL = strSQL & gstrBairro & " BR,"
            strSQL = strSQL & gstrSubempenho & " SEP,"
            strSQL = strSQL & gstrComprasLicitacao & " CL, "
            strSQL = strSQL & gstrUnidadeOrcamentaria & " UO, "
            strSQL = strSQL & gstrElementoDespesa & " ED, "
                  
            If blnSoEstorno Then
               strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT ,"
            Else
               strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT , "
            End If
          
            strSQL = strSQL & gstrLogradouro & " LG,"
            strSQL = strSQL & gstrCidade & " MPC,"
            strSQL = strSQL & gstrUF & " UFC,"
            strSQL = strSQL & gstrUF & " UFD, "
            strSQL = strSQL & gstrReservaDotacao & " RD, "
            strSQL = strSQL & gstrRequisicaoCompras & " RC "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
          
          
            If blnSoEstorno Then
                strSQL = strSQL & "SEP.Pkid " & strOUTJSQLServer & "=IT.intSubEmpenho " & strOUTJOracle & " AND "
            Else
                strSQL = strSQL & "EP.Pkid " & strOUTJSQLServer & "= IT.Intempenho " & strOUTJOracle & " AND "
            End If
          
            strSQL = strSQL & "EP.PKID = SEP.intEmpenho AND "
            strSQL = strSQL & "EP.intProgramaTrabalho = PT.PKID AND "
            strSQL = strSQL & "CL.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intModalidade   AND "
            strSQL = strSQL & "PT.intOrgao = OG.PKID AND "
            strSQL = strSQL & "PT.intProjetoAtividade = PJ.PKID AND "
            strSQL = strSQL & "PT.intFonteRecurso = FR.PKID AND "
            strSQL = strSQL & "CV.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intConvenio AND "
            strSQL = strSQL & "PE.intPedidoEmpenho " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.Pkid AND "
            strSQL = strSQL & "RC.PKId " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intRequisicaodeCompra AND "
            strSQL = strSQL & "LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " RC.intCodigoCentroDeCusto2 AND "
            strSQL = strSQL & "AC.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intautorizacaodecompra AND "
            strSQL = strSQL & "EP.intCredor = CT.PKID AND "
            strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
            strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
            strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
            strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
            strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
            strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
            strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ") & " AND "
          
            strSQL = strSQL & "PT.intUnidadeOrcamentaria = UO.PKID AND "
            strSQL = strSQL & "PT.intElementoDespesa = ED.PKID"
          
            strSQL = strSQL & " AND LG.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intLogradouro "
            strSQL = strSQL & " AND MPC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipioC "
            strSQL = strSQL & " AND UFC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFC "
            strSQL = strSQL & " AND UFD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFD "
            strSQL = strSQL & " AND RD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intReservaDotacao "
          
        Else
'            strSQL = strSQL & "FROM "
'            strSQL = strSQL & gstrEmpenho & " EP "
'            strSQL = strSQL & "INNER JOIN " & gstrSubempenho & " SEP ON (EP.PKID = SEP.intEmpenho) "
'            strSQL = strSQL & "INNER JOIN " & gstrProgramaDeTrabalho & " PT ON (EP.intProgramaTrabalho = PT.PKID) "
'            strSQL = strSQL & "LEFT JOIN " & gstrComprasLicitacao & " CL ON (CL.PKID  = EP.intModalidade) "
'            strSQL = strSQL & "INNER JOIN " & gstrOrgao & " OG ON (PT.intOrgao = OG.PKID ) "
'            strSQL = strSQL & "INNER JOIN " & gstrProjeto & " PJ ON (PT.intProjetoAtividade = PJ.PKID) "
'            strSQL = strSQL & "INNER JOIN " & gstrFonteRecurso & " FR ON (PT.intFonteRecurso = FR.PKID) "
'            strSQL = strSQL & "LEFT JOIN "
'            strSQL = strSQL & gstrEmpenhoContrato & " PE ON (PE.intPedidoEmpenho = EP.PKId) LEFT JOIN "
'            strSQL = strSQL & gstrRequisicaoCompras & " RC ON (RC.PKId = PE.intRequisicaodeCompra) LEFT JOIN "
'            strSQL = strSQL & gstrLocais & " LO ON (LO.PKId = RC.intCodigoCentroDeCusto2) LEFT JOIN "
'            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC ON (AC.PKId = PE.intAutorizacaodeCompra) LEFT JOIN "
'            strSQL = strSQL & gstrConvenio & " CV ON (CV.PKID  = EP.intConvenio) "
'            strSQL = strSQL & "INNER JOIN " & gstrContribuinte & " CT ON (EP.intCredor = CT.PKID) "
'            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MP ON (MP.PKID  = CT.intMunicipio) "
'            strSQL = strSQL & "LEFT JOIN " & gstrBairro & " BR ON (BR.PKID  = CT.intBairro) "
'            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UF ON (UF.PKID  = CT.intUF) "
'            strSQL = strSQL & "LEFT JOIN " & gstrLogradouro & " LG ON (LG.PKID  = CT.intLogradouro) "
'            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MPC ON (MPC.PKID  = CT.intMunicipioC) "
'            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFC ON (UFC.PKID  = CT.intUFC) "
'            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFD ON (UFD.PKID  = CT.intUFD) "
'
'            strSQL = strSQL & "INNER JOIN " & gstrUnidadeOrcamentaria & " UO ON (PT.intUnidadeOrcamentaria = UO.PKID) "
'            strSQL = strSQL & "INNER JOIN " & gstrElementoDespesa & " ED ON (PT.intElementoDespesa = ED.PKID) "
'
'            strSQL = strSQL & "LEFT JOIN "
'
            strSQL = strSQL & " FROM "
             strSQL = strSQL & gstrEmpenho & " EP "
             strSQL = strSQL & "INNER JOIN " & gstrSubempenho & " SEP ON (EP.PKID = SEP.intEmpenho) "
             strSQL = strSQL & "INNER JOIN " & gstrProgramaDeTrabalho & " PT ON (EP.intProgramaTrabalho = PT.PKID) "
             strSQL = strSQL & "LEFT JOIN "
             strSQL = strSQL & gstrComprasLicitacao & " CL "
             strSQL = strSQL & "ON(CL.PKID = EP.intModalidade) "
             strSQL = strSQL & "INNER JOIN " & gstrOrgao & " OG ON (PT.intOrgao = OG.PKID) "
             strSQL = strSQL & "INNER JOIN " & gstrProjeto & " PJ ON (PT.intProjetoAtividade = PJ.PKID) "
             strSQL = strSQL & "INNER JOIN " & gstrFonteRecurso & " FR ON (PT.intFonteRecurso = FR.PKID) "
'             strSQL = strSQL & "LEFT JOIN " & gstrRequisicaoCompras & " RC "
             strSQL = strSQL & "LEFT JOIN "
             strSQL = strSQL & gstrEmpenhoContrato & " PE ON (PE.intPedidoEmpenho = EP.PKId) LEFT JOIN "
             strSQL = strSQL & gstrRequisicaoCompras & " RC ON (RC.PKId = PE.intRequisicaodeCompra) LEFT JOIN "
'             strSQL = strSQL & "ON PT.pkid = RC.intProgramadeTrabalho LEFT JOIN "
'             strSQL = strSQL & gstrEmpenhoContrato & " PE "
'             strSQL = strSQL & "ON PE.intPedidoEmpenho = EP.PKId AND RC.PKID = PE. INTREQUISICAODECOMPRA LEFT JOIN "
             strSQL = strSQL & gstrLocais & " LO "
             strSQL = strSQL & "ON(LO.PKId = RC.intCodigoCentroDeCusto2) LEFT JOIN "
             strSQL = strSQL & gstrAutorizacaoDeCompra & " AC "
             strSQL = strSQL & " ON(AC.PKId = PE.intAutorizacaodeCompra) LEFT JOIN "
             strSQL = strSQL & gstrConvenio & " CV "
             strSQL = strSQL & "ON(CV.PKID = EP.intConvenio) "
             strSQL = strSQL & "INNER JOIN " & gstrContribuinte & " CT ON (EP.intCredor = CT.PKID) LEFT JOIN "
             strSQL = strSQL & gstrCidade & " MP "
             strSQL = strSQL & "ON(MP.PKID = CT.intMunicipio) LEFT JOIN "
             strSQL = strSQL & gstrBairro & " BR ON(BR.PKID = CT.intBairro) LEFT JOIN "
             strSQL = strSQL & gstrUF & " UF ON(UF.PKID = CT.intUF) LEFT JOIN "
             strSQL = strSQL & gstrLogradouro & " LG "
             strSQL = strSQL & "ON(LG.PKID = CT.intLogradouro) LEFT JOIN "
             strSQL = strSQL & gstrCidade & " MPC "
             strSQL = strSQL & "ON(MPC.PKID = CT.intMunicipioC) LEFT JOIN "
             strSQL = strSQL & gstrUF & " UFC ON(UFC.PKID = CT.intUFC) LEFT JOIN "
             strSQL = strSQL & gstrUF & " UFD ON(UFD.PKID = CT.intUFD) "
             strSQL = strSQL & "INNER JOIN " & gstrUnidadeOrcamentaria & " UO ON (PT.intUnidadeOrcamentaria = UO.PKID) "
             strSQL = strSQL & "INNER JOIN " & gstrElementoDespesa & " ED ON (PT.intElementoDespesa = ED.PKID) LEFT JOIN"


            If blnSoEstorno Then
                strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT "
            Else
                strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT  "
                 'strSQL = strSQL & gstrItemEmpenho & " IT "
            End If
             
            If blnSoEstorno Then
                strSQL = strSQL & "ON (SEP.Pkid = IT.intSubEmpenho ) "
            Else
                strSQL = strSQL & "ON (EP.Pkid = IT.Intempenho ) "
            End If
            
            strSQL = strSQL & "LEFT JOIN " & gstrReservaDotacao & " RD ON (EP.intReservaDotacao  = RD.pkid) "
            
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
            strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
            strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
            strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
            strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ")
        
        End If
        strSQL = strSQL & AdicionaGroupByQueryRelatorio
    End If
    
'--------------------------------------------------------------------------------------------------------------
   
    If blnProcesso Then
        If strSQL = "" Then
            strSQL = "SELECT "
        Else
            strSQL = strSQL & "UNION ALL SELECT "
        End If
      
        strSQL = strSQL & "EP.intNumero intEmpenho, "
        strSQL = strSQL & "EP.PKID intEmpenhoPKID, "
        strSQL = strSQL & "'" & gstrMascaraItemDespesa & "' Mascara, "
        strSQL = strSQL & "'PROCESSO' strDestino, "
        strSQL = strSQL & "EP.strModalidade, "
        strSQL = strSQL & "RC.intPedidoEmpenho PedidoEmpenho, "
        strSQL = strSQL & "LO.strDescricao UnidadeCC, "
        strSQL = strSQL & "AC.strobjetoautorizacao, "
        strSQL = strSQL & strTelefone & " telefone, "
        strSQL = strSQL & "EP.strContrato, "
        strSQL = strSQL & "EP.strSolicitacao, "
        strSQL = strSQL & "EP.strCodigo, "
        strSQL = strSQL & "EP.intExercicio, "
        strSQL = strSQL & "EP.bitDigito, "
        strSQL = strSQL & "EP.dblValor dblValorEmpenho, "
        
        strSQL = strSQL & "RD.intNumero NumeroReserva, "
        strSQL = strSQL & "RD.intExercicioReserva , "
        strSQL = strSQL & "RD.strSolicitacao SolicitacaoReserva, "
        strSQL = strSQL & "RD.intExercicio ExercicioSolicitacaoReserva, "
        
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "EP.PKID") & strCONCAT & "'/4' Agrupamento, "
        ' Este campo é usado para ordenar os registros e para servir de grupo a relatórios externos que
        ' não exibam os subempenhos
        
        strSQL = strSQL & "EP.Strcondpagto, "
        strSQL = strSQL & "EP.Strlocentrega, "
        strSQL = strSQL & "EP.Strprazoentrega, "
        strSQL = strSQL & "EP.dtmdata DataEmpenho, "
        strSQL = strSQL & "PT.PKID CodigoProgramaTrabalho, "
        strSQL = strSQL & "PT.strCodigo strDotacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "PT.intCodigoReduzido, "
        strSQL = strSQL & "PJ.strcodigo CodProjetoAtividade, "
        strSQL = strSQL & "PT.intPrograma CodPrograma, "
        strSQL = strSQL & "PT.intSubFuncao CodSubFuncao, "
        strSQL = strSQL & "PT.intFuncao CodFuncao, "
        strSQL = strSQL & "PT.dblValor dblProgramaTrabalho, "
        strSQL = strSQL & "PT.intProjetoAtividade intProjetoAtividade, "
        strSQL = strSQL & "OG.strCodigo CodOrgao, "
        strSQL = strSQL & "OG.strDescricao strOrgao, "
        strSQL = strSQL & "UO.strCodigo CodUnidadeOrcamentaria, "
        strSQL = strSQL & "UO.strDescricao UnidadeOrcamentaria, "
        strSQL = strSQL & "ED.strCodigoElementoDespesa CodElementoDespesa, "
        strSQL = strSQL & "ED.strDescricao ElementoDespesa, "
        strSQL = strSQL & "CL.strCodigo CodCompraLicitacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "FR.strCodigo CodigoFonteRecurso, "
        strSQL = strSQL & "FR.strDescricao strFonteRecurso, "
        strSQL = strSQL & "CV.strCodigo CodigoConvenio, "
        strSQL = strSQL & "CV.strDescricao strConvenio, "
        strSQL = strSQL & "CT.CDC intCodigoContribuinte, "
        strSQL = strSQL & "CT.strNome, "
        strSQL = strSQL & "CT.strCNPJCPF, "
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "LG.STRDESCRICAO") & " strEndereco, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.StrlogradouroD", "CT.Strlogradouroc"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intNumero") & " INTNUMERO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intNumeroD", "CT.intNumeroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.strComplemento") & " STRCOMPLEMENTO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strComplementoD", "CT.strComplementoC"))
              
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "MP.strDescricao") & " STRMUNICIPIO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strMunicipioD", "MPC.strDescricao"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "UF.strSigla") & " STRUF, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',UFD.strSigla", "UFC.strSigla"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "BR.strDescricao") & " STRBAIRRO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strbairroD", "CT.strbairroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intCEP") & " INTCEP, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intCEPD", "CT.intCEPC"))
        
        strSQL = strSQL & "CT.Strlogradouroc,"
        strSQL = strSQL & "CT.intNumeroC,"
        strSQL = strSQL & "CT.strComplementoC,"
        strSQL = strSQL & "MPC.strDescricao strMunicipioC,"
        strSQL = strSQL & "UFC.strSigla strUFC,"
        strSQL = strSQL & "CT.strBairroC,"
        strSQL = strSQL & "CT.intCEPC,"
        strSQL = strSQL & "SEP.PKID intPKIDParcela, "
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "SEP.PKID") & strCONCAT & "'/4' AgrupamentoParcela, "
        strSQL = strSQL & "SEP.intNumero intParcela, "
        strSQL = strSQL & "SEP.bytSituacao, "
        strSQL = strSQL & "SEP.bytTipo, "
        
        If blnSoEstorno Then
           strSQL = strSQL & "'NOTA DE ESTORNO' Nota, "
           strSQL = strSQL & "SEP.strHistorico, "
        Else
           strSQL = strSQL & "'NOTA DE EMPENHO' Nota, "
           strSQL = strSQL & gstrCASEWHEN("SEP.intNumero", "0, EP.strHistorico", "SEP.strHistorico") & " strHistorico, "
        End If
      
        strSQL = strSQL & "SEP.dtmData, "
        strSQL = strSQL & "SEP.intEmpenhoAnulacao, "
        strSQL = strSQL & "SEP.dblValor, "
        strSQL = strSQL & gstrISNULL("SEP.dblEmpenhadoAteData", "0") & " AS dblEmpenhadoAteData, "
        strSQL = strSQL & gstrISNULL("SEP.dblSaldoAtual", "0") & " AS dblSaldoAtual, "
        strSQL = strSQL & "CL.strDescricao strModalidadeLicitacao, "
        strSQL = strSQL & "IT.intcodigoitem AS CodItem,"
        strSQL = strSQL & "IT.Strdescricaoitem AS DescItem,"
        strSQL = strSQL & "IT.Strmarca AS MarcaItem,"
        strSQL = strSQL & "IT.STRUNIDADE as UnidItem,"
        strSQL = strSQL & "IT.DBLQUANTIDADE AS QuantItem,"
        strSQL = strSQL & "IT.Dblprecounitario AS PrecoItem,"
        strSQL = strSQL & "(IT.DBLQUANTIDADE * IT.Dblprecounitario) AS PrecoTotItem, "
        'Incluida a Data de Vencimento Ficha Orc1136 - Fernando
        strSQL = strSQL & "Sep.DtmVencimento DtmVencimento "
  
        If bytDBType = EDatabases.Oracle Then
            strSQL = strSQL & "FROM " & gstrEmpenho & " EP,"
            strSQL = strSQL & gstrProgramaDeTrabalho & " PT,"
            strSQL = strSQL & gstrOrgao & " OG,"
            strSQL = strSQL & gstrProjeto & " PJ,"
            strSQL = strSQL & gstrFonteRecurso & " FR,"
            strSQL = strSQL & gstrConvenio & " CV,"
            strSQL = strSQL & gstrEmpenhoContrato & " PE, " 'ON (PE.intNumeroEmpenho = EP.PKId) INNER JOIN "
            strSQL = strSQL & gstrLocais & " LO, " 'ON (LO.PKId = PE.intUnidadeCentroDeCusto) INNER JOIN "
            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC, " 'ON (AC.PKId = PE.intautorizacaodecompra) LEFT JOIN "
            strSQL = strSQL & gstrContribuinte & " CT,"
            strSQL = strSQL & gstrCidade & " MP,"
            strSQL = strSQL & gstrUF & " UF,"
            strSQL = strSQL & gstrBairro & " BR,"
            strSQL = strSQL & gstrSubempenho & " SEP,"
            strSQL = strSQL & gstrComprasLicitacao & " CL, "
            strSQL = strSQL & gstrUnidadeOrcamentaria & " UO, "
            strSQL = strSQL & gstrElementoDespesa & " ED, "
            If blnSoEstorno Then
                strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT ,"
            Else
                strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT , "
            End If
         
            strSQL = strSQL & gstrLogradouro & " LG,"
            strSQL = strSQL & gstrCidade & " MPC,"
            strSQL = strSQL & gstrUF & " UFC,"
            strSQL = strSQL & gstrUF & " UFD, "
            strSQL = strSQL & gstrReservaDotacao & " RD, "
            strSQL = strSQL & gstrRequisicaoCompras & " RC "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
         
            If blnSoEstorno Then
               strSQL = strSQL & "SEP.Pkid " & strOUTJSQLServer & "=IT.intSubEmpenho " & strOUTJOracle & " AND "
            Else
               strSQL = strSQL & "EP.Pkid " & strOUTJSQLServer & "= IT.Intempenho " & strOUTJOracle & " AND "
            End If
         
            strSQL = strSQL & "EP.PKID = SEP.intEmpenho AND "
            strSQL = strSQL & "EP.intProgramaTrabalho = PT.PKID AND "
            strSQL = strSQL & "CL.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intModalidade   AND "
            strSQL = strSQL & "PT.intOrgao = OG.PKID AND "
            strSQL = strSQL & "PT.intProjetoAtividade = PJ.PKID AND "
            strSQL = strSQL & "PT.intFonteRecurso = FR.PKID AND "
            strSQL = strSQL & "CV.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intConvenio AND "
            strSQL = strSQL & "PE.intPedidoEmpenho " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.Pkid AND "
            strSQL = strSQL & "RC.PKId " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intRequisicaodeCompra AND "
            strSQL = strSQL & "LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " RC.intCodigoCentroDeCusto2 AND "
            strSQL = strSQL & "AC.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intautorizacaodecompra AND "
            strSQL = strSQL & "EP.intCredor = CT.PKID AND "
            strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
            strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
            strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
            strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
            strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
            strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
            strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ") & " AND "
            strSQL = strSQL & "PT.intUnidadeOrcamentaria = UO.PKID AND "
            strSQL = strSQL & "PT.intElementoDespesa = ED.PKID"
            strSQL = strSQL & " AND LG.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intLogradouro "
            strSQL = strSQL & " AND MPC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipioC "
            strSQL = strSQL & " AND UFC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFC "
            strSQL = strSQL & " AND UFD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFD "
            strSQL = strSQL & " AND RD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intReservaDotacao "
        Else
            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrEmpenho & " EP "
            strSQL = strSQL & "INNER JOIN " & gstrSubempenho & " SEP ON (EP.PKID = SEP.intEmpenho) "
            strSQL = strSQL & "INNER JOIN " & gstrProgramaDeTrabalho & " PT ON (EP.intProgramaTrabalho = PT.PKID) "
            strSQL = strSQL & "LEFT JOIN " & gstrComprasLicitacao & " CL ON (CL.PKID  = EP.intModalidade) "
            strSQL = strSQL & "INNER JOIN " & gstrOrgao & " OG ON (PT.intOrgao = OG.PKID ) "
            strSQL = strSQL & "INNER JOIN " & gstrProjeto & " PJ ON (PT.intProjetoAtividade = PJ.PKID) "
            strSQL = strSQL & "INNER JOIN " & gstrFonteRecurso & " FR ON (PT.intFonteRecurso = FR.PKID) "
            strSQL = strSQL & "LEFT JOIN "
            strSQL = strSQL & gstrEmpenhoContrato & " PE ON (PE.intPedidoEmpenho = EP.PKId) LEFT JOIN "
            strSQL = strSQL & gstrRequisicaoCompras & " RC ON (RC.PKId = PE.intRequisicaodeCompra) LEFT JOIN "
            strSQL = strSQL & gstrLocais & " LO ON (LO.PKId = RC.intCodigoCentroDeCusto2) LEFT JOIN "
            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC ON (AC.PKId = PE.intAutorizacaodeCompra) LEFT JOIN "
            strSQL = strSQL & gstrConvenio & " CV ON (CV.PKID  = EP.intConvenio) "
            strSQL = strSQL & "INNER JOIN " & gstrContribuinte & " CT ON (EP.intCredor = CT.PKID) "
            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MP ON (MP.PKID  = CT.intMunicipio) "
            strSQL = strSQL & "LEFT JOIN " & gstrBairro & " BR ON (BR.PKID  = CT.intBairro) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UF ON (UF.PKID  = CT.intUF) "
            strSQL = strSQL & "LEFT JOIN " & gstrLogradouro & " LG ON (LG.PKID  = CT.intLogradouro) "
            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MPC ON (MPC.PKID  = CT.intMunicipioC) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFC ON (UFC.PKID  = CT.intUFC) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFD ON (UFD.PKID  = CT.intUFD) "
            
            strSQL = strSQL & "INNER JOIN " & gstrUnidadeOrcamentaria & " UO ON (PT.intUnidadeOrcamentaria = UO.PKID) "
            strSQL = strSQL & "INNER JOIN " & gstrElementoDespesa & " ED ON (PT.intElementoDespesa = ED.PKID) "
            
            strSQL = strSQL & "LEFT JOIN "
            
            If blnSoEstorno Then
                strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT "
            Else
                strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT  "
            End If
             
            If blnSoEstorno Then
                strSQL = strSQL & "ON (SEP.Pkid = IT.intSubEmpenho )"
            Else
                strSQL = strSQL & "ON (EP.Pkid = IT.Intempenho )"
            End If
            
            strSQL = strSQL & "LEFT JOIN " & gstrReservaDotacao & " RD ON (EP.intReservaDotacao  = RD.pkid) "
            
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
            strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
            strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
            strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
            strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ")
        End If
        
        strSQL = strSQL & AdicionaGroupByQueryRelatorio
    
    End If
'-----------------------------------------------------------------------
    If blnTesouraria Then
        If strSQL = "" Then
            strSQL = "SELECT "
        Else
            strSQL = strSQL & "UNION ALL SELECT "
        End If
      
        strSQL = strSQL & "EP.intNumero intEmpenho, "
        strSQL = strSQL & "EP.PKID intEmpenhoPKID, "
        strSQL = strSQL & "'" & gstrMascaraItemDespesa & "' Mascara, "
        strSQL = strSQL & "'TESOURARIA' strDestino, "
        strSQL = strSQL & "EP.strModalidade, "
        strSQL = strSQL & "RC.intPedidoEmpenho PedidoEmpenho, "
        strSQL = strSQL & "LO.strDescricao UnidadeCC, "
        strSQL = strSQL & "AC.strobjetoautorizacao, "
        strSQL = strSQL & strTelefone & " telefone, "
        strSQL = strSQL & "EP.strContrato, "
        strSQL = strSQL & "EP.strSolicitacao, "
        strSQL = strSQL & "EP.strCodigo, "
        strSQL = strSQL & "EP.intExercicio, "
        strSQL = strSQL & "EP.bitDigito, "
        strSQL = strSQL & "EP.dblValor dblValorEmpenho, "
        
        strSQL = strSQL & "RD.intNumero NumeroReserva, "
        strSQL = strSQL & "RD.intExercicioReserva , "
        strSQL = strSQL & "RD.strSolicitacao SolicitacaoReserva, "
        strSQL = strSQL & "RD.intExercicio ExercicioSolicitacaoReserva, "
        
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "EP.PKID") & strCONCAT & "'/5' Agrupamento, "
        ' Este campo é usado para ordenar os registros e para servir de grupo a relatórios externos que
        ' não exibam os subempenhos
        
        strSQL = strSQL & "EP.Strcondpagto, "
        strSQL = strSQL & "EP.Strlocentrega, "
        strSQL = strSQL & "EP.Strprazoentrega, "
        strSQL = strSQL & "EP.dtmdata DataEmpenho, "
        strSQL = strSQL & "PT.PKID CodigoProgramaTrabalho, "
        strSQL = strSQL & "PT.strCodigo strDotacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "PT.intCodigoReduzido, "
        strSQL = strSQL & "PJ.strcodigo CodProjetoAtividade, "
        strSQL = strSQL & "PT.intPrograma CodPrograma, "
        strSQL = strSQL & "PT.intSubFuncao CodSubFuncao, "
        strSQL = strSQL & "PT.intFuncao CodFuncao, "
        strSQL = strSQL & "PT.dblValor dblProgramaTrabalho, "
        strSQL = strSQL & "PT.intProjetoAtividade intProjetoAtividade, "
        strSQL = strSQL & "OG.strCodigo CodOrgao, "
        strSQL = strSQL & "OG.strDescricao strOrgao, "
        strSQL = strSQL & "UO.strCodigo CodUnidadeOrcamentaria, "
        strSQL = strSQL & "UO.strDescricao UnidadeOrcamentaria, "
        strSQL = strSQL & "ED.strCodigoElementoDespesa CodElementoDespesa, "
        strSQL = strSQL & "ED.strDescricao ElementoDespesa, "
        strSQL = strSQL & "CL.strCodigo CodCompraLicitacao, "
        strSQL = strSQL & "PJ.strDescricao strProjeto, "
        strSQL = strSQL & "FR.strCodigo CodigoFonteRecurso, "
        strSQL = strSQL & "FR.strDescricao strFonteRecurso, "
        strSQL = strSQL & "CV.strCodigo CodigoConvenio, "
        strSQL = strSQL & "CV.strDescricao strConvenio, "
        strSQL = strSQL & "CT.CDC intCodigoContribuinte, "
        strSQL = strSQL & "CT.strNome, "
        strSQL = strSQL & "CT.strCNPJCPF, "
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "LG.STRDESCRICAO") & " strEndereco, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.StrlogradouroD", "CT.Strlogradouroc"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intNumero") & " INTNUMERO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intNumeroD", "CT.intNumeroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.strComplemento") & " STRCOMPLEMENTO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strComplementoD", "CT.strComplementoC"))
              
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "MP.strDescricao") & " STRMUNICIPIO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strMunicipioD", "MPC.strDescricao"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "UF.strSigla") & " STRUF, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',UFD.strSigla", "UFC.strSigla"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "BR.strDescricao") & " STRBAIRRO, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strbairroD", "CT.strbairroC"))
        
        strSQL = strSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intCEP") & " INTCEP, "
        strSQL = Replace(strSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intCEPD", "CT.intCEPC"))
        
        strSQL = strSQL & "CT.Strlogradouroc,"
        strSQL = strSQL & "CT.intNumeroC,"
        strSQL = strSQL & "CT.strComplementoC,"
        strSQL = strSQL & "MPC.strDescricao strMunicipioC,"
        strSQL = strSQL & "UFC.strSigla strUFC,"
        strSQL = strSQL & "CT.strBairroC,"
        strSQL = strSQL & "CT.intCEPC,"
        strSQL = strSQL & "SEP.PKID intPKIDParcela, "
        strSQL = strSQL & gstrCONVERT(CDT_NVARCHAR, "SEP.PKID") & strCONCAT & "'/5' AgrupamentoParcela, "
        strSQL = strSQL & "SEP.intNumero intParcela, "
        strSQL = strSQL & "SEP.bytSituacao, "
        strSQL = strSQL & "SEP.bytTipo, "
        
        If blnSoEstorno Then
           strSQL = strSQL & "'NOTA DE ESTORNO' Nota, "
           strSQL = strSQL & "SEP.strHistorico, "
        Else
           strSQL = strSQL & "'NOTA DE EMPENHO' Nota, "
           strSQL = strSQL & gstrCASEWHEN("SEP.intNumero", "0, EP.strHistorico", "SEP.strHistorico") & " strHistorico, "
        End If
        
        strSQL = strSQL & "SEP.dtmData, "
        strSQL = strSQL & "SEP.intEmpenhoAnulacao, "
        strSQL = strSQL & "SEP.dblValor, "
        strSQL = strSQL & gstrISNULL("SEP.dblEmpenhadoAteData", "0") & " AS dblEmpenhadoAteData, "
        strSQL = strSQL & gstrISNULL("SEP.dblSaldoAtual", "0") & " AS dblSaldoAtual, "
        strSQL = strSQL & "CL.strDescricao strModalidadeLicitacao, "
        strSQL = strSQL & "IT.intcodigoitem AS CodItem,"
        strSQL = strSQL & "IT.Strdescricaoitem AS DescItem,"
        strSQL = strSQL & "IT.Strmarca AS MarcaItem,"
        strSQL = strSQL & "IT.STRUNIDADE as UnidItem,"
        strSQL = strSQL & "IT.DBLQUANTIDADE AS QuantItem,"
        strSQL = strSQL & "IT.Dblprecounitario AS PrecoItem,"
        strSQL = strSQL & "(IT.DBLQUANTIDADE * IT.Dblprecounitario) AS PrecoTotItem, "
        'Incluida a Data de Vencimento Ficha Orc1136 - Fernando
        strSQL = strSQL & "Sep.DtmVencimento DtmVencimento "
        
        If bytDBType = EDatabases.Oracle Then
            strSQL = strSQL & "FROM " & gstrEmpenho & " EP,"
            strSQL = strSQL & gstrProgramaDeTrabalho & " PT,"
            strSQL = strSQL & gstrOrgao & " OG,"
            strSQL = strSQL & gstrProjeto & " PJ,"
            strSQL = strSQL & gstrFonteRecurso & " FR,"
            strSQL = strSQL & gstrConvenio & " CV,"
            strSQL = strSQL & gstrEmpenhoContrato & " PE, " 'ON (PE.intNumeroEmpenho = EP.PKId) INNER JOIN "
            strSQL = strSQL & gstrLocais & " LO, " 'ON (LO.PKId = PE.intUnidadeCentroDeCusto) INNER JOIN "
            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC, " 'ON (AC.PKId = PE.intautorizacaodecompra) LEFT JOIN "
            strSQL = strSQL & gstrContribuinte & " CT,"
            strSQL = strSQL & gstrCidade & " MP,"
            strSQL = strSQL & gstrUF & " UF,"
            strSQL = strSQL & gstrBairro & " BR,"
            strSQL = strSQL & gstrSubempenho & " SEP,"
            strSQL = strSQL & gstrComprasLicitacao & " CL, "
            strSQL = strSQL & gstrUnidadeOrcamentaria & " UO, "
            strSQL = strSQL & gstrElementoDespesa & " ED, "
            If blnSoEstorno Then
                strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT ,"
            Else
                strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT , "
            End If
            
            strSQL = strSQL & gstrLogradouro & " LG,"
            strSQL = strSQL & gstrCidade & " MPC,"
            strSQL = strSQL & gstrUF & " UFC,"
            strSQL = strSQL & gstrUF & " UFD, "
            strSQL = strSQL & gstrReservaDotacao & " RD, "
            strSQL = strSQL & gstrRequisicaoCompras & " RC "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
         
         
            If blnSoEstorno Then
               strSQL = strSQL & "SEP.Pkid " & strOUTJSQLServer & "=IT.intSubEmpenho " & strOUTJOracle & " AND "
            Else
               strSQL = strSQL & "EP.Pkid " & strOUTJSQLServer & "= IT.Intempenho " & strOUTJOracle & " AND "
            End If
            
            strSQL = strSQL & "EP.PKID = SEP.intEmpenho AND "
            strSQL = strSQL & "EP.intProgramaTrabalho = PT.PKID AND "
            strSQL = strSQL & "CL.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intModalidade   AND "
            strSQL = strSQL & "PT.intOrgao = OG.PKID AND "
            strSQL = strSQL & "PT.intProjetoAtividade = PJ.PKID AND "
            strSQL = strSQL & "PT.intFonteRecurso = FR.PKID AND "
            strSQL = strSQL & "CV.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intConvenio AND "
            strSQL = strSQL & "PE.intPedidoEmpenho " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.Pkid AND "
            strSQL = strSQL & "RC.PKId " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intRequisicaodeCompra AND "
            strSQL = strSQL & "LO.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " RC.intCodigoCentroDeCusto2 AND "
            strSQL = strSQL & "AC.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " PE.intautorizacaodecompra AND "
            strSQL = strSQL & "EP.intCredor = CT.PKID AND "
            strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
            strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
            strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
            strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
            strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
            strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
            strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ") & " AND "
            strSQL = strSQL & "PT.intUnidadeOrcamentaria = UO.PKID AND "
            strSQL = strSQL & "PT.intElementoDespesa = ED.PKID"
            strSQL = strSQL & " AND LG.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intLogradouro "
            strSQL = strSQL & " AND MPC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipioC "
            strSQL = strSQL & " AND UFC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFC "
            strSQL = strSQL & " AND UFD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFD "
            strSQL = strSQL & " AND RD.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " EP.intReservaDotacao "
        Else
            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrEmpenho & " EP "
            strSQL = strSQL & "INNER JOIN " & gstrSubempenho & " SEP ON (EP.PKID = SEP.intEmpenho) "
            strSQL = strSQL & "INNER JOIN " & gstrProgramaDeTrabalho & " PT ON (EP.intProgramaTrabalho = PT.PKID) "
            strSQL = strSQL & "LEFT JOIN " & gstrComprasLicitacao & " CL ON (CL.PKID  = EP.intModalidade) "
            strSQL = strSQL & "INNER JOIN " & gstrOrgao & " OG ON (PT.intOrgao = OG.PKID ) "
            strSQL = strSQL & "INNER JOIN " & gstrProjeto & " PJ ON (PT.intProjetoAtividade = PJ.PKID) "
            strSQL = strSQL & "INNER JOIN " & gstrFonteRecurso & " FR ON (PT.intFonteRecurso = FR.PKID) "
            strSQL = strSQL & "LEFT JOIN "
            strSQL = strSQL & gstrEmpenhoContrato & " PE ON (PE.intPedidoEmpenho = EP.PKId) LEFT JOIN "
            strSQL = strSQL & gstrRequisicaoCompras & " RC ON (RC.PKId = PE.intRequisicaodeCompra) LEFT JOIN "
            strSQL = strSQL & gstrLocais & " LO ON (LO.PKId = RC.intCodigoCentroDeCusto2) LEFT JOIN "
            strSQL = strSQL & gstrAutorizacaoDeCompra & " AC ON (AC.PKId = PE.intAutorizacaodeCompra) LEFT JOIN "
            strSQL = strSQL & gstrConvenio & " CV ON (CV.PKID  = EP.intConvenio) "
            strSQL = strSQL & "INNER JOIN " & gstrContribuinte & " CT ON (EP.intCredor = CT.PKID) "
            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MP ON (MP.PKID  = CT.intMunicipio) "
            strSQL = strSQL & "LEFT JOIN " & gstrBairro & " BR ON (BR.PKID  = CT.intBairro) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UF ON (UF.PKID  = CT.intUF) "
            strSQL = strSQL & "LEFT JOIN " & gstrLogradouro & " LG ON (LG.PKID  = CT.intLogradouro) "
            strSQL = strSQL & "LEFT JOIN " & gstrCidade & " MPC ON (MPC.PKID  = CT.intMunicipioC) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFC ON (UFC.PKID  = CT.intUFC) "
            strSQL = strSQL & "LEFT JOIN " & gstrUF & " UFD ON (UFD.PKID  = CT.intUFD) "
            
            strSQL = strSQL & "INNER JOIN " & gstrUnidadeOrcamentaria & " UO ON (PT.intUnidadeOrcamentaria = UO.PKID) "
            strSQL = strSQL & "INNER JOIN " & gstrElementoDespesa & " ED ON (PT.intElementoDespesa = ED.PKID) "
            
            strSQL = strSQL & "LEFT JOIN "
            
            If blnSoEstorno Then
               strSQL = strSQL & "(SELECT IE.Pkid, IEA.intSubEmpenho, IEA.DBLQUANTIDADE, IEA.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE, " & gstrItemEmpenhoAnulado & " IEA WHERE IE.PKID = IEA.INTITEMEMPENHO) IT "
            Else
               strSQL = strSQL & "(SELECT IE.Pkid, IE.DBLQUANTIDADE, IE.Dblprecounitario, IE.Intempenho, IE.STRUNIDADE , IE.Strmarca ,IE.intcodigoitem, IE.Strdescricaoitem FROM " & gstrItemEmpenho & " IE) IT  "
            End If
          
             If blnSoEstorno Then
                strSQL = strSQL & "ON (SEP.Pkid = IT.intSubEmpenho )"
                   
             Else
                strSQL = strSQL & "ON (EP.Pkid = IT.Intempenho )"
             End If
            
             strSQL = strSQL & "LEFT JOIN " & gstrReservaDotacao & " RD ON (EP.intReservaDotacao  = RD.pkid) "
            
             strSQL = strSQL & " WHERE "
             strSQL = strSQL & gstrDATEPART(strYEAR, "EP.dtmData") & " = " & intExercicioEmpenho & " AND "
             strSQL = strSQL & "(EP.intNumero BETWEEN " & strEmpInicial & " AND " & strEmpFinal & " OR  "
             strSQL = strSQL & "SEP.intEmpenhoAnulacao BETWEEN " & strEmpInicial & " AND " & strEmpFinal & ") AND  "
             strSQL = strSQL & "SEP.intNumero BETWEEN " & strParcInicial & " AND " & strParcFinal & " AND "
             strSQL = strSQL & IIf(blnSoEstorno, "SEP.bytSituacao = 4 ", "SEP.bytSituacao <> 4 ")
        End If
        
        strSQL = strSQL & AdicionaGroupByQueryRelatorio
    
    End If
    
    strSQL = strSQL & " ORDER BY intEmpenho, intParcela, Agrupamento, CodItem "
    
    strQueryRelatorio = strSQL
 
End Function



Private Sub LeTabelaProgramaTrabalhoParaReserva(strPKIDProgTrab As String)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    If cbo_intEvento.ListIndex <> -1 Then
       
       cboProgramaTrabalho.Clear
       cboCodigoReduzido.Clear
       
       strSQL = ""
       strSQL = strSQL & "SELECT PT.PKId, PT.intCodigoReduzido, PT.strCodigo, "
       strSQL = strSQL & " ED.strCodigoElementoDespesa "
       strSQL = strSQL & " FROM " & gstrProgramaDeTrabalho & " PT, "
       strSQL = strSQL & gstrElementoDespesa & " ED "
       strSQL = strSQL & " WHERE PT.intElementoDespesa = ED.PKID AND "
       strSQL = strSQL & strSUBSTRING & "(ED.strCodigoElementoDespesa,1," & Len(BuscaCodigosPeloEvento(gstrItemData(cbo_intEvento), gstrDigitoDespesa, "D", 2)) & ") = '" & _
                                        BuscaCodigosPeloEvento(gstrItemData(cbo_intEvento), gstrDigitoDespesa, "D", 2) & "'"
       strSQL = strSQL & " AND PT.intExercicio = " & gintExercicio
         
       strSQL = strSQL & " AND PT.PKID = " & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex)
         
       strSQL = strSQL & " ORDER BY PT.strCodigo"
         
       Set gobjBanco = New clsBanco
       If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
           With adoResultado
               Do While Not .EOF
                   If Not IsNull(!intCodigoReduzido) Then
                     cboProgramaTrabalho.AddItem !strCodigo
                     cboProgramaTrabalho.ItemData(cboProgramaTrabalho.NewIndex) = !Pkid
                   End If
                   .MoveNext
               Loop
           End With
       End If
       
       strSQL = Mid(strSQL, 1, Len(strSQL) - 21) + "ORDER BY intCodigoReduzido"
       
       If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
           With adoResultado
               Do While Not .EOF
                   If Not IsNull(!intCodigoReduzido) Then
                      cboCodigoReduzido.AddItem (!intCodigoReduzido)
                      cboCodigoReduzido.ItemData(cboCodigoReduzido.NewIndex) = !Pkid
                   End If
                   .MoveNext
               Loop
           End With
       End If
      TrocaCorObjeto cboProgramaTrabalho, True
      TrocaCorObjeto cboCodigoReduzido, True
   Else
      If Not mblnAlterandoEmpenho Then ExibeMensagem "É necessário informar o Evento Contábil antes de informar o número de Reserva."
      txt_Reservado = ""
      txt_Cancelado = ""
      txt_Empenhado = ""
      txt_Saldo = ""
      cbointReservaDotacao.Clear
   End If
End Sub

Private Function PreencheDadosReserva()
   Dim strSQL       As String
   Dim adoResultado As ADODB.Recordset
   
   
   If Trim(cbointReservaDotacao.Text) = "" Then Exit Function
   
   strSQL = "SELECT * FROM " & gstrReservaDotacao
   strSQL = strSQL & " WHERE intNumero = " & cbointReservaDotacao.Text
   strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "dtmData") & " = "
   If IsDate(txtDTMDATA) Then
       strSQL = strSQL & Year(txtDTMDATA)
   Else
       strSQL = strSQL & gintExercicio
   End If
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
      If Not adoResultado.EOF Then
         dtmDataReserva = CDate(gstrDataFormatada(adoResultado!DTMDATA))
           
         txt_Reservado = gstrConvVrDoSql(adoResultado!dblValor)
      Else
         txt_Reservado = "0,00"
      End If
   
 End If
   adoResultado.Close
   
   strSQL = "SELECT " & gstrISNULL("SUM(dblValor)", "0", "SUM(dblValor)") & " AS dblValor FROM " & gstrReservaDotacaoLiberada
   'strSql = strSql & " WHERE intFlag = 0 AND intReservaDotacao = " & cbointReservaDotacao.Text
   strSQL = strSQL & " WHERE intFlag = 0 AND intReservaDotacao = " & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex)
   
   If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
      If Not adoResultado.EOF Then
         txt_Cancelado = gstrConvVrDoSql(adoResultado!dblValor)
      Else
         txt_Cancelado = "0,00"
      End If
   End If
   adoResultado.Close
   
   strSQL = "SELECT " & gstrISNULL("SUM(dblValor)", "0", "SUM(dblValor)") & " dblValor"
   strSQL = strSQL & " FROM " & gstrReservaDotacaoLiberada
   strSQL = strSQL & " WHERE intFlag = 1 AND intReservaDotacao = "
   'strSql = strSql & cbointReservaDotacao.Text
   strSQL = strSQL & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex)
   
   If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
      If Not adoResultado.EOF Then
         txt_Empenhado = gstrConvVrDoSql(adoResultado!dblValor)
      Else
         txt_Empenhado = "0,00"
      End If
   End If
   
   adoResultado.Close
   
   txt_Saldo = gstrConvVrDoSql(CDbl(txt_Reservado) - (CDbl(txt_Cancelado) + CDbl(txt_Empenhado)))
   
          
       strSQL = "SELECT PT.PKID FROM " & gstrProgramaDeTrabalho & " PT, " & gstrReservaDotacao & " RD "
       strSQL = strSQL & " WHERE RD.intProgramaTrabalho = PT.PKID AND RD.pkid = " & gstrItemData(cbointReservaDotacao)
       
        
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If Not adoResultado.EOF Then
    '           LeTabelaProgramaTrabalho (CStr(adoResultado!Pkid))
            
                'PreencherListaDeOpcoes cboProgramaTrabalho, adoResultado!Pkid
                'PreencherListaDeOpcoes cboCodigoReduzido, adoResultado!Pkid
             'cboCodigoReduzido.ListIndex = -1
                         
             blnAlteraReserva = False
             cboProgramaTrabalho.ListIndex = gintIndiceCBO(cboProgramaTrabalho, adoResultado!Pkid)
             cboCodigoReduzido.ListIndex = gintIndiceCBO(cboCodigoReduzido, adoResultado!Pkid)
    
             If cboCodigoReduzido.ListIndex = -1 Then
                 LeTabelaProgramaTrabalho
                 cboProgramaTrabalho.ListIndex = gintIndiceCBO(cboProgramaTrabalho, adoResultado!Pkid)
                 cboCodigoReduzido.ListIndex = gintIndiceCBO(cboCodigoReduzido, adoResultado!Pkid)
    
             End If
             
             preencheCboevento
          End If
       End If
       
       adoResultado.Close
       blnAlteraReserva = True
       Set gobjBanco = Nothing
  
  End Function

Private Function blnVerificaParamAnul() As Boolean
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT bytanuladespesaqualquerempenho "
    strSQL = strSQL & "FROM " & gstrConfiguracaoGeral

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            blnVerificaParamAnul = adoResultado!bytAnulaDespesaQualquerEmpenho
        End If
    End If
End Function


Private Function GeraAnulacaoReservaDotacaoLiberada()
   Dim strSQL As String
   Dim dblValorReservDotLiberado As Double
   Dim dblValorReservDot As Double
   Dim ValorMaximoLiberar As Double
   Dim ValorLiberado As Double
   
   dblValorReservDotLiberado = VerificaValorReservaDotacaoLib
   
   If dblValorReservDotLiberado < CDbl(txt_ValorAnulacao) Then
       ValorLiberado = dblValorReservDotLiberado
   Else
       ValorLiberado = CDbl(txt_ValorAnulacao)
   End If
   
   
   strSQL = "INSERT INTO " & gstrReservaDotacaoLiberada & " ("
   strSQL = strSQL & "intReservaDotacao, intNumero, dtmData, dblValor, "
   strSQL = strSQL & "strHistorico, dtmDtAtualizacao, lngCodUsr,intFlag, intEmpenho "
   strSQL = strSQL & ") (SELECT "
   strSQL = strSQL & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex) & ", " & gstrISNULL("MAX(intNumero)", "0") & " + 1, "
   strSQL = strSQL & gstrConvDtParaSql(txt_DataAnulucao) & ", "
   strSQL = strSQL & gstrConvVrParaSql(ValorLiberado) & " * (-1), "
   strSQL = strSQL & "'" & txt_HistoricoAnulacao & "', "
   strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
   strSQL = strSQL & glngCodUsr & ", 1, " & txtPKId & " "
   strSQL = strSQL & "FROM " & gstrReservaDotacaoLiberada & " "
   strSQL = strSQL & "WHERE intReservaDotacao = " & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex) & ")"
   
   Set gobjBanco = New clsBanco
   gobjBanco.Execute (strSQL)
   Set gobjBanco = Nothing
   
End Function

Private Function VerificaValorReservaDotacao() As Double
    Dim strSQL As String
    Dim adoResultado  As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT DblValor "
    strSQL = strSQL & "FROM " & gstrReservaDotacao
    strSQL = strSQL & " WHERE PKID = " & gstrItemData(cbointReservaDotacao)
    
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
        If Not adoResultado.EOF Then
            VerificaValorReservaDotacao = gstrConvVrDoSql(adoResultado!dblValor)
        End If
    End If
End Function

Private Function VerificaValorReservaDotacaoLib() As Double
    Dim strSQL As String
    Dim adoResultado  As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT " & gstrISNULL("SUM(RL.DblValor)", "0") & " DblValor "
    strSQL = strSQL & "FROM " & gstrReservaDotacao & " RD ,"
    strSQL = strSQL & gstrReservaDotacaoLiberada & " RL "
    strSQL = strSQL & " WHERE RD.PKID = " & gstrItemData(cbointReservaDotacao)
    strSQL = strSQL & " AND RL.intReservaDotacao = RD.PKID "
    
    
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
        If Not adoResultado.EOF Then
            VerificaValorReservaDotacaoLib = gstrConvVrDoSql(adoResultado!dblValor)
        
        End If
    End If
End Function


Private Function GravaReservaDotacaoLiberada()
   Dim strSQL                      As String
   Dim dblValorReservDotLiberado   As Double
   Dim dblValorReservDot           As Double
   Dim ValorMaximoLiberar          As Double
   Dim ValorLiberado               As Double
   Dim dblValorCancelamento           As Double
   
   dblValorReservDotLiberado = VerificaValorReservaDotacaoLib
   dblValorReservDot = VerificaValorReservaDotacao
   
   ValorMaximoLiberar = dblValorReservDot - dblValorReservDotLiberado
   
   If ValorMaximoLiberar < CDbl(txtdblValor) Then
       ValorLiberado = ValorMaximoLiberar
   Else
       ValorLiberado = CDbl(txtdblValor)
   End If
   
   strSQL = "INSERT INTO " & gstrReservaDotacaoLiberada & " ("
   strSQL = strSQL & "intReservaDotacao, intNumero, dtmData, dblValor, "
   strSQL = strSQL & "strHistorico, dtmDtAtualizacao, lngCodUsr,intFlag, intEmpenho "
   strSQL = strSQL & ") (SELECT "
   strSQL = strSQL & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex) & ", " & gstrISNULL("MAX(intNumero)", "0") & " + 1, "
   strSQL = strSQL & gstrConvDtParaSql(txtDTMDATA) & ", "
   strSQL = strSQL & gstrConvVrParaSql(ValorLiberado) & ", "
   strSQL = strSQL & "'" & txtstrHistorico & "', "
   strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
   strSQL = strSQL & glngCodUsr & ", 1, " & glngRetornaPkidTabelaPai("seq" & gstrEmpenho, gstrEmpenho) & " "
   strSQL = strSQL & "FROM " & gstrReservaDotacaoLiberada & " "
   strSQL = strSQL & "WHERE intReservaDotacao = " & cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex) & ")"
   
   Set gobjBanco = New clsBanco
   gobjBanco.Execute (strSQL)
   Set gobjBanco = Nothing
   
   'FAZ O CANCELAMENTO DA RESERVA DE DOTAÇÃO RETORNANDO O VALOR PARA A VERBA
   If Val(gstrConvVrParaSql(txtdblValor)) < Val(gstrConvVrParaSql(txt_Saldo)) Then
       dblValorCancelamento = Val(gstrConvVrParaSql(txt_Saldo)) - Val(gstrConvVrParaSql(txtdblValor))
       CancelarReservaDotacao cbointReservaDotacao.ItemData(cbointReservaDotacao.ListIndex), txtDTMDATA.Text, dblValorCancelamento, "Cancelamento automático do saldo da reserva *Tela Empenho", True
   End If
   
End Function

Private Function gstrPkidEmpenhoInsercao() As String
    Dim strSQL As String
    Dim adoResultado  As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT pkid "
    strSQL = strSQL & "FROM " & gstrEmpenho
    strSQL = strSQL & " WHERE intnumero = " & txtintNumero.Text
    
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSQL, 20, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_DataFinal.Text = gstrDataFormatada(adoResultado!dtmDataAplicacaoFinal)
            txt_SaldoConvenio.Text = gstrConvVrDoSql(adoResultado!dblValor)
        End If
    End If
    
    

End Function


Private Function GravaSubElementos(Optional PKIDEmpenho As String, Optional pkidParcela As String, Optional lvwObj As ListView)
   Dim strSQL As String
   Dim i      As Integer
   
   
    For i = 1 To lvwObj.ListItems.Count
        
        strSQL = "INSERT INTO " & gstrSubElementoEmpenho & " ("
        strSQL = strSQL & IIf(PKIDEmpenho <> "", "intEmpenho", "intparcela")
        strSQL = strSQL & ",intItemDespesa,DblValor, dtmDtAtualizacao, lngCodUsr)"
        
        strSQL = strSQL & " Values ( "
        
        strSQL = strSQL & IIf(PKIDEmpenho <> "", PKIDEmpenho, pkidParcela) & ","
        strSQL = strSQL & lvwObj.ListItems(i).Tag & ","
        strSQL = strSQL & gstrConvVrParaSql(lvwObj.ListItems(i).SubItems(3)) & ","
        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
        strSQL = strSQL & glngCodUsr
        
        strSQL = strSQL & " ) "
        
        Set gobjBanco = New clsBanco
        gobjBanco.Execute (strSQL)
        Set gobjBanco = Nothing
                
    Next
   
End Function


Private Function VerificaAdiantamentos() As Boolean

   Dim strSQL        As String
   Dim adoResultado  As ADODB.Recordset
   Dim intQuantidade As Long
   Dim dblValor      As Double
   
   strSQL = "SELECT * FROM " & gstrTipoEmpenho
   strSQL = strSQL & " WHERE PKID = " & gstrItemData(dbcintTipo, True)
   strSQL = strSQL & " AND bytAdiantamento = 1"
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
         
      If Not adoResultado.EOF Then
         adoResultado.Close
         
         strSQL = "SELECT * FROM " & gstrParametrosContabeis
         strSQL = strSQL & " WHERE strCodigo = '3'"
         
         Set gobjBanco = New clsBanco
         
         If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If adoResultado.EOF Then
               ExibeMensagem "É necessário informar a Quantidade de adiantamentos permitidos para cada credor."
               adoResultado.Close
               Exit Function
            Else
               intQuantidade = adoResultado!dblValorParametro
            End If
         End If
         
         adoResultado.Close
      
         strSQL = "SELECT * FROM " & gstrParametrosContabeis
         strSQL = strSQL & " WHERE strCodigo = '6'"
         
         Set gobjBanco = New clsBanco
         
         If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If adoResultado.EOF Then
               ExibeMensagem "É necessário informar o valor máximo do adiantamento."
               adoResultado.Close
               Exit Function
            Else
               dblValor = Val(gstrConvVrParaSql(adoResultado!dblValorParametro))
            End If
         End If
         
         adoResultado.Close
      
         strSQL = "SELECT COUNT(*) AS Quantidade FROM " & gstrEmpenho & " EP, " & gstrSubempenho & " SEP, "
         strSQL = strSQL & gstrTipoEmpenho & " TEP "
         strSQL = strSQL & " WHERE TEP.bytAdiantamento = 1 AND "
         strSQL = strSQL & " EP.intTipo = TEP.PKID AND EP.PKID = SEP.intEmpenho AND "
         strSQL = strSQL & " SEP.intNumero = 0 AND bytSituacao IN (1,2) AND "
         strSQL = strSQL & " EP.intCredor = " & gstrItemData(dbcintCredor, True)
         
         Set gobjBanco = New clsBanco
         
         If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If Not adoResultado.EOF Then
               If intQuantidade <= adoResultado!Quantidade Then
                  ExibeMensagem "Este credor ultrapassou a quantidade de adiantamentos permitidos."
                  If dbcintCredor.Enabled Then dbcintCredor.SetFocus
                  adoResultado.Close
                  Exit Function
               End If
            End If
         End If
         
         If dblValor < Val(gstrConvVrParaSql(txtdblValor)) Then
            ExibeMensagem "O valor deste empenho supera o valor permitido para cada adiantamento."
            If txtdblValor.Enabled Then txtdblValor.SetFocus
            Exit Function
         End If
      
      End If
      adoResultado.Close
   End If
   
   VerificaAdiantamentos = True
         
End Function
Private Function DesabilitaControlesRP()

   TrocaCorObjeto cbointReservaDotacao, mblnRestosAPagar
   TrocaCorObjeto cmd_Reserva, mblnRestosAPagar
'   TrocaCorObjeto cboCodigoReduzido, mblnRestosAPagar
'   TrocaCorObjeto cboProgramaTrabalho, mblnRestosAPagar
'   TrocaCorObjeto cmd_ProgramaTrabalho, mblnRestosAPagar
'   TrocaCorObjeto dbcintItemDespesa, mblnRestosAPagar
'   TrocaCorObjeto cmd_ItemDespesa, mblnRestosAPagar
'   TrocaCorObjeto txt_intNContribuinte, mblnRestosAPagar
'   TrocaCorObjeto txtstrCodigo, mblnRestosAPagar
'   TrocaCorObjeto txtbitDigito, mblnRestosAPagar
'   TrocaCorObjeto txtintExercicio, mblnRestosAPagar
'   TrocaCorObjeto dbcintCredor, mblnRestosAPagar
'   TrocaCorObjeto cmd_Credor, mblnRestosAPagar
'   TrocaCorObjeto txt_codEvento, mblnRestosAPagar
'   TrocaCorObjeto cbo_intEvento, mblnRestosAPagar
'   TrocaCorObjeto cmd_Evento, mblnRestosAPagar
'   TrocaCorObjeto cmd_Credor, mblnRestosAPagar
'   TrocaCorObjeto dbcintTipo, mblnRestosAPagar
'   TrocaCorObjeto cmd_Tipo, mblnRestosAPagar
'   TrocaCorObjeto txtStrhistorico, mblnRestosAPagar
'   TrocaCorObjeto txtstrCodigo, mblnRestosAPagar
'   TrocaCorObjeto txtbitDigito, mblnRestosAPagar
'   TrocaCorObjeto txtintExercicio, mblnRestosAPagar
'   TrocaCorObjeto cbo_Historico, mblnRestosAPagar
'   TrocaCorObjeto cmd_Historico, mblnRestosAPagar
'   TrocaCorObjeto txtstrContrato, mblnRestosAPagar
'   TrocaCorObjeto txtstrEmbasamento, mblnRestosAPagar
'   TrocaCorObjeto txtstrModalidade, mblnRestosAPagar
'   TrocaCorObjeto txtdtmHomologacao, mblnRestosAPagar
'   TrocaCorObjeto txtstrLicitacao, mblnRestosAPagar
'   TrocaCorObjeto txtstrSolicitacao, mblnRestosAPagar
'   TrocaCorObjeto dbcintFundo, mblnRestosAPagar
'   TrocaCorObjeto cmd_Fundo, mblnRestosAPagar
'   TrocaCorObjeto dbcintConvenio, mblnRestosAPagar
'   TrocaCorObjeto cmd_Convenio, mblnRestosAPagar
'   tab_3dPasta.TabEnabled(2) = Not mblnRestosAPagar
'    txt_DataComplemento.Visible = tab_3dPasta.TabEnabled(2)
'    txt_ValorComplemento.Visible = tab_3dPasta.TabEnabled(2)
'    lblTotalComplemento.Visible = tab_3dPasta.TabEnabled(2)
'    txt_HistoricoComplemento.Visible = tab_3dPasta.TabEnabled(2)
'    cbo_HistoricoComplemento.Visible = tab_3dPasta.TabEnabled(2)
'    cmd_HistoricoComplemento.Visible = tab_3dPasta.TabEnabled(2)
    
End Function


Sub leConvenio()
    Dim strSQL As String
    Dim adoResultado  As ADODB.Recordset
    
    If dbcintConvenio.BoundText = "" Then
        txt_DataFinal.Text = ""
        txt_SaldoConvenio.Text = ""
        Exit Sub
    End If
    
    strSQL = ""
    strSQL = ""
    strSQL = strSQL & "SELECT dtmDataAplicacaoFinal, dblValor "
    strSQL = strSQL & "FROM " & gstrConvenio
    strSQL = strSQL & " WHERE PKID = " & dbcintConvenio.BoundText
    
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSQL, 20, adoResultado) Then
        If Not adoResultado.EOF Then
            txt_DataFinal.Text = gstrDataFormatada(adoResultado!dtmDataAplicacaoFinal)
            txt_SaldoConvenio.Text = gstrConvVrDoSql(adoResultado!dblValor)
        End If
    End If
    
    
End Sub
Private Function LimpaTelaAnulacao()

    txt_DataAnulucao = ""
    txt_ValorAnulacao = ""
    txt_HistoricoAnulacao = ""
    txt_CodHistoricoAnl.Text = ""
    cbo_HistoricoAnulacao.ListIndex = -1
    txt_CodEventoAnul.Text = ""
    cbo_intEventoAnul.ListIndex = -1

    'TrocaCorObjeto cbo_HistoricoAnulacao, False
    If txt_DataAnulucao.Enabled Then txt_DataAnulucao.SetFocus
    tab_3DAnulacao.TabEnabled(1) = False
    
End Function

Private Sub txt_codEvento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_codEvento_LostFocus()
Dim strEventoDaReserva As String
Dim strSQL As String
Dim adoResultado As ADODB.Recordset

    If Not IsNumeric(txt_codEvento) Then
        txt_codEvento = ""
    End If
    
    If cboProgramaTrabalho.ListIndex <> -1 Then
        strEventoDaReserva = EventoDaReserva
        
        If Len(Trim(strEventoDaReserva)) = 0 Then
            strEventoDaReserva = "''"
        End If
        
        strSQL = ""
        strSQL = strSQL & " SELECT 1 " & IIf(bytDBType = Oracle, " FROM DUAL ", " ")
        strSQL = strSQL & " WHERE " & gstrENulo(txt_codEvento, , True)
        strSQL = strSQL & " IN "
        strSQL = strSQL & " (SELECT strCodigo FROM " & gstrEvento
        strSQL = strSQL & " WHERE PKID IN (" & strEventoDaReserva & "))"
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If adoResultado.EOF Then
                txt_codEvento = ""
            End If
        End If
        
    End If
    
    PreencheEventobyCodigo txt_codEvento, cbo_intEvento, "2"
    cbo_intevento_LostFocus
End Sub

Private Sub cbo_intevento_GotFocus()
    If cbo_intEvento.Text = "" Then txt_codEvento.Text = ""
    strEvento = cbo_intEvento.Text
End Sub


Private Sub cbo_intevento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_intevento_LostFocus()
Dim strTextoDigitado As String
Dim blnAchou         As Boolean
Dim intContador      As Integer

If strEvento = cbo_intEvento.Text Then Exit Sub
    blnAchou = False
    With cbo_intEvento
        If .ListIndex = -1 Then
        strTextoDigitado = .Text
        For intContador = 0 To .ListCount - 1
            If .list(intContador) = strTextoDigitado Then
            blnAchou = True
            .ListIndex = intContador
            Exit For
            End If
        Next
        Else
        blnAchou = True
        End If

    End With
    If blnAchou = True Then
       LeTabelaReservaDotacao
       cbointReservaDotacao_Click
    Else
        cbo_intEvento.Text = ""
        cbointReservaDotacao.Clear
        txt_codEvento.Text = ""
    End If
    
    
    'If cbo_intEvento.ListIndex = -1 Then
    '   txt_codEvento.Text = ""
    'Else
    '  LeTabelaProgramaTrabalho
    'End If
End Sub

Private Sub cmd_Evento_Click()
    CarregaForm frmCadEvento, cbo_intEvento, strQueryAplicarEvento
End Sub

Private Function strQueryAplicarEvento() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEvento & " "
    strSQL = strSQL & "WHERE intTipoEvento = 2 "
    strQueryAplicarEvento = strSQL
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra
End Function

Private Sub preencheCboevento(Optional strFiltroPKID As Variant)
Dim strSQL As String
Dim strPKIds As String

    If IsMissing(strFiltroPKID) Then
        strSQL = "SELECT PKID, strDescricao FROM " & gstrEvento & " WHERE intTipoEvento=2"
    Else
        strSQL = "SELECT PKID, strDescricao FROM " & gstrEvento & " WHERE intTipoEvento=2 AND PKID = " & strFiltroPKID
    End If
    
    If cboProgramaTrabalho.ListIndex <> -1 Then
        strPKIds = EventoDaReserva
        If Len(Trim(strPKIds)) > 0 Then
            strSQL = strSQL & " AND PKID in (" & strPKIds & ")"
        Else
            strSQL = strSQL & " AND 1 <> 1" ' Para não retornar ninguém
        End If
    End If
    
    LeDaTabelaParaObj gstrEvento, cbo_intEvento, strSQL
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra
    
    
    If cbo_intEvento.ListCount = 1 Then
        cbo_intEvento.ListIndex = 0
    Else
        cbo_intEvento.Text = ""
    End If
    
    
End Sub

Private Sub txt_codeventoLiq_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_codeventoLiq_LostFocus()
    PreencheEventobyCodigo txt_codEventoLiq, cbo_intEventoLiq, "7"
End Sub

Private Sub cbo_inteventoLiq_Click()
    leCodigoEvento txt_codEventoLiq, cbo_intEventoLiq
End Sub

Private Sub cbo_inteventoLiq_GotFocus()
    If cbo_intEventoLiq.Text = "" Then txt_codEventoLiq.Text = ""
End Sub

Private Sub cbo_inteventoLiq_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_inteventoLiq_LostFocus()
    If cbo_intEventoLiq.Text = "" Then txt_codEventoLiq.Text = ""
End Sub

Private Sub cmd_EventoLiq_Click()
    CarregaForm frmCadEvento, cbo_intEventoLiq, strQueryAplicarEventoLiq
End Sub

Private Sub txt_codeventoLiqAutomatica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txt_codeventoLiqAutomatica_LostFocus()
    PreencheEventobyCodigo txt_codEventoLiqAutomatica, cbo_intEventoLiqAutomatica, "7"
End Sub

Private Sub cbo_inteventoLiqAutomatica_Click()
    leCodigoEvento txt_codEventoLiqAutomatica, cbo_intEventoLiqAutomatica
End Sub

Private Sub cbo_inteventoLiqAutomatica_GotFocus()
    If cbo_intEventoLiqAutomatica.Text = "" Then txt_codEventoLiqAutomatica.Text = ""
End Sub


Private Sub cbo_inteventoLiqAutomatica_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_inteventoLiqAutomatica_LostFocus()
    If cbo_intEventoLiqAutomatica.Text = "" Then txt_codEventoLiqAutomatica.Text = ""
End Sub

Private Sub cmd_EventoLiqAutomatica_Click()
    CarregaForm frmCadEvento, cbo_intEventoLiqAutomatica, strQueryAplicarEventoLiq
End Sub

Private Function strQueryAplicarEventoLiq() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEvento & " "
    strSQL = strSQL & "WHERE intTipoEvento = 7 "
    strQueryAplicarEventoLiq = strSQL
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra, , 6-alteracoes orcamentarias
    '              7-Liquidação
End Function

Private Sub preencheCboeventoLiq()
    Dim strSQL As String
    'M6R (Perguntar sobre o ano)
    strSQL = "SELECT PKID, strDescricao FROM " & gstrEvento & " WHERE intTipoEvento = 7 "
    'Verifica o ano do evento de acordo com a tela Empenho ou Resto a Pagar
    If mblnRestosAPagar Then
       strSQL = strSQL & " AND intExercicio < " & gintExercicio
    Else
       strSQL = strSQL & " AND intExercicio = " & gintExercicio
    End If
    LeDaTabelaParaObj gstrEvento, cbo_intEventoLiq, strSQL
    
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra , 6-alteracoes orcamentarias
    '              7-Liquidação
End Sub

Private Sub preencheCboeventoLiqAutomatica()
    Dim strSQL As String
    'M6R (Perguntar sobre o ano)
    strSQL = "SELECT PKID, strDescricao FROM " & gstrEvento & " WHERE intTipoEvento = 7 "
    'Verifica o ano do evento de acordo com a tela Empenho ou Resto a Pagar
    If mblnRestosAPagar Then
       strSQL = strSQL & " AND intExercicio < " & gintExercicio
    Else
       strSQL = strSQL & " AND intExercicio = " & gintExercicio
    End If
    LeDaTabelaParaObj gstrEvento, cbo_intEventoLiqAutomatica, strSQL
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra , 6-alteracoes orcamentarias
    '              7-Liquidação
End Sub

Private Sub proximoCodigoEmpenho()
    If Not mblnRestosAPagar Then
       'gstrProximoCodigo txtintNumero, gstrEmpenho, "intNumero", gintCodSeguranca, gstrDATEPART(strYEAR, "dtmData"), Val(gintExercicio), , , , , gstrDATEPART(strYEAR, "dtmData"), Val(gintExercicio)
       txtintNumero = GeraProximoDeEmpenho(True)
       MarcaCampo txtintNumero
    End If
End Sub

Public Sub LeByLinhaSelecionada()
    mblnClickOk = True
    mblnPrimeiraVez = False
    tdb_Lista_RowColChange 0, 0
    tab_3dPasta.Tab = 3
    
End Sub
Private Sub VerificaNF()

   If blnDadosOKNF Then
      If mblnAlterandoNF Then
         lvw_NotasFiscais.SelectedItem.Text = txt_dtmDataNF.Text
         lvw_NotasFiscais.SelectedItem.SubItems(1) = gstrConvVrDoSql(txt_dblValorNF)
         lvw_NotasFiscais.SelectedItem.SubItems(2) = txt_strNotasFiscais
         lvw_NotasFiscais.SelectedItem.SubItems(3) = "1"
      Else
         Set mobjLista = lvw_NotasFiscais.ListItems.Add(, , txt_dtmDataNF.Text)
         mobjLista.SubItems(1) = gstrConvVrDoSql(txt_dblValorNF)
         mobjLista.SubItems(2) = txt_strNotasFiscais
         mobjLista.SubItems(3) = "0"
         lbl_ValorTotal = gstrConvVrDoSql(CDbl(IIf(Len(Trim(lbl_ValorTotal)) = 0, "0", lbl_ValorTotal)) + CDbl(IIf(Len(Trim(txt_dblValorNF)) = 0, "0", txt_dblValorNF)))
         LimpaDadosNF
      End If
   End If
End Sub

Private Sub VerificaNFExcluir()
With lvw_NotasFiscais
        If .ListItems.Count > 0 Then
            If Len(Trim(.SelectedItem.Tag)) > 0 Then mstrNFPkidExcluir = mstrNFPkidExcluir & .SelectedItem.Tag & ","
            .ListItems.Remove .SelectedItem.Index
            mblnAlterandoNF = False
            lbl_ValorTotal = gstrConvVrDoSql(CDbl(IIf(Len(Trim(lbl_ValorTotal)) = 0, "0", lbl_ValorTotal)) - CDbl(IIf(Len(Trim(txt_dblValorNF)) = 0, "0", txt_dblValorNF)))
            LimpaDadosNF
        End If
    End With
End Sub

Private Function blnDadosOKNF() As Boolean
   
   If Trim(txt_dtmDataNF.Text) = "" Then
        ExibeMensagem "A Data da Nota Fiscal tem que ser informada."
        If txt_dtmDataNF.Enabled Then txt_dtmDataNF.SetFocus
        Exit Function
    End If
        
    If Trim(txt_dblValorNF.Text) = "" Then
        ExibeMensagem "O Valor da Nota Fiscal tem que ser informado."
        If txt_dblValorNF.Enabled Then txt_dblValorNF.SetFocus
        Exit Function
    End If
    
    If CDbl(txt_dblValorNF) = 0 Then
       ExibeMensagem "Não é possível inserir uma Nota fiscal com valor zero."
        If txt_dblValorNF.Enabled Then txt_dblValorNF.SetFocus
       Exit Function
    End If
    
    If Val(gstrConvVrParaSql(lbl_ValorTotal)) + Val(gstrConvVrParaSql(txt_dblValorNF)) > Val(gstrConvVrParaSql(txt_dblValorAux)) Then
       ExibeMensagem "O valor total da(s) Nota(s) Fiscal(is) não poder(m) ser superior ao valor liquidado."
       If txt_dtmDataNF.Enabled Then
          If txt_dblValorNF.Enabled Then txt_dblValorNF.SetFocus
       End If
       Exit Function
    End If
        
    If gblnDataValida(txt_dtmDataNF) = False Then
        ExibeMensagem "A data da Nota Fiscal tem que ser informada corretamente."
        If txt_dtmDataNF.Enabled Then
            If txt_dtmDataNF.Enabled Then txt_dtmDataNF.SetFocus
        End If
        Exit Function
    End If
    

'    If IsDate(txt_dtmDataNF.Text) Then
'        If Year(CDate(txt_dtmDataNF)) <> CInt(gintExercicio) Then
'            ExibeMensagem "A data da nota fiscal tem que estar no exercício de " & gintExercicio & "."
'            If txt_dtmDataNF.Enabled Then txt_dtmDataNF.SetFocus
'            Exit Function
'        End If
'    End If

    
    blnDadosOKNF = True
    
End Function
Private Sub LimpaDadosNF(Optional blnFechaCampos As Boolean)
   

   TrocaCorObjeto txt_dtmDataNF, blnFechaCampos
   TrocaCorObjeto txt_dblValorNF, blnFechaCampos
   TrocaCorObjeto txt_strNotasFiscais, blnFechaCampos
   mblnAlterandoNF = False
   txt_dblValorNF = ""
   txt_dtmDataNF = ""
   txt_strNotasFiscais = "Sem Nota Fiscal"
   If txt_dtmDataNF.Enabled Then txt_dtmDataNF.SetFocus
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem
   'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrExcluirItem

 End Sub
 
 Private Function GravaNotasFiscais() As Boolean
    Dim strSQL     As String
    Dim intInd     As Integer
    Dim intCont    As Integer
    Dim intParcela As Byte
    
    GravaNotasFiscais = False
    strSQL = ""
    
    If lvw_Liquidacao.SelectedItem Is Nothing Then
        Exit Function
    End If
    For intCont = 1 To lvw_Liquidacao.ListItems.Count
        If lvw_Liquidacao.ListItems(intCont).Text <> "0" Then
           intParcela = 1
        End If
    Next
    
    If Val(gstrConvVrParaSql(lbl_ValorTotal)) > Val(gstrConvVrParaSql(txt_dblValorAux)) Then
       'ExibeMensagem "O valor total da(s) Nota(s) Fiscal(is) deve ser igual ao valor liquidado."
       'If txt_dtmDataNF.Enabled Then txt_dtmDataNF.SetFocus
       msgNotas = "Erro ao gravar Notas Fiscais:  " & "O valor total da(s) Nota(s) Fiscal(is) não pode ser maior que o valor da parcela."
       Exit Function
    ElseIf intParcela = 1 And lvw_Liquidacao.ListItems(lvw_Liquidacao.SelectedItem.Index).Text = "0" Then
       Exit Function
    Else
       With lvw_NotasFiscais
           strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
           For intInd = 1 To .ListItems.Count
              If .ListItems(intInd).SubItems(3) = "0" Then 'somente os itens ainda não cadastrados
                  strSQL = strSQL & "INSERT INTO " & gstrSubEmpenhoNF & " ("
                  strSQL = strSQL & "intSubEmpenho, dtmData, dblValorNF, "
                  strSQL = strSQL & "strNotaFiscal, dtmDtAtualizacao, lngCodUsr) VALUES "
                  strSQL = strSQL & "( " & lvw_Liquidacao.SelectedItem.Tag & " , "
                  strSQL = strSQL & gstrConvDtParaSql(.ListItems(intInd).Text) & ", "
                  strSQL = strSQL & gstrConvVrParaSql(.ListItems(intInd).SubItems(1)) & ", '"
                  strSQL = strSQL & .ListItems(intInd).SubItems(2) & "', "
                  strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & " "
                  strSQL = strSQL & "); "
                  .ListItems(intInd).SubItems(3) = "1"
              End If
           Next
           strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
       End With
    End If
    
    
    If lvw_NotasFiscais.ListItems.Count = 0 Then
        msgNotas = "Erro ao gravar notas Fiscais:  " & "Não existe nenhuma Nota Fiscal para ser gravada."
    ElseIf Trim(strSQL) = "BEGIN END;" Then
        'ExibeMensagem "Não existe nenhum item novo para ser Gravado."
        'msgNotas = "Gravação de Notas Fiscais:  " & "Não existe nenhuma Nota Fiscal nova para ser Gravada."
        Exit Function
    End If
    
    
    If Trim(strSQL) = "BEGIN END;" Then Exit Function
    
    If strSQL <> "" Then
       Set gobjBanco = New clsBanco
       If gobjBanco.Execute(strSQL) Then
          If Len(Trim(mstrNFPkidExcluir)) > 0 Then
             mstrNFPkidExcluir = "(" & Mid(mstrNFPkidExcluir, 1, Len(mstrNFPkidExcluir) - 1) & ")"
             strSQL = "DELETE FROM " & gstrSubEmpenhoNF & " WHERE PKID IN " & mstrNFPkidExcluir
             mstrNFPkidExcluir = ""
             If Not gobjBanco.Execute(strSQL) Then
                'ExibeMensagem "Problemas durante a gravação das Notas Fiscais.Entre em contato com o fornecedor."
                msgNotas = "Erro ao gravar de Notas Fiscais:  " & "Problemas durante a gravação das Notas Fiscais.Entre em contato com o fornecedor."
                Exit Function
             End If
          End If
          LeDaTabelaParaObj "", lvw_NotasFiscais, strQueryNotasFiscais
          msgNotas = "Gravação de Notas Fiscais:  " & "Nota(s) Fiscal(is) gravada(s) com sucesso."
          'ExibeMensagem "Nota(s) Fiscal(is) gravada(s) com sucesso. "
          GravaNotasFiscais = True
       End If
   End If
End Function

Private Function strQueryNotasFiscais() As String

    Dim strSQL       As String
    Dim adoResultado As New ADODB.Recordset
    Dim dblTotal     As Double
    
    dblTotal = 0
    
    strSQL = "SELECT PKID, dtmData, dblValorNF, strNotaFiscal "
    strSQL = strSQL & " FROM " & gstrSubEmpenhoNF
    strSQL = strSQL & " WHERE intSubEmpenho = " & lvw_Liquidacao.SelectedItem.Tag
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
       With adoResultado
          While Not .EOF
             dblTotal = dblTotal + gstrConvVrDoSql(!dblValorNF)
             .MoveNext
          Wend
          
       End With
    End If
    
    lbl_ValorTotal = gstrConvVrDoSql(dblTotal)
    
    strQueryNotasFiscais = strSQL

End Function


Private Function VerificaNotasFiscais(ByVal strSubEmpenhoPKID As String) As Boolean

    Dim strSQL       As String
    Dim adoResultado As New ADODB.Recordset
    
    strSQL = "SELECT PKID, dtmData, dblValorNF, strNotaFiscal "
    strSQL = strSQL & " FROM " & gstrSubEmpenhoNF
    strSQL = strSQL & " WHERE intSubEmpenho = " & strSubEmpenhoPKID
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
       With adoResultado
          If Not .EOF Then VerificaNotasFiscais = True
       End With
    End If

End Function

Private Sub limpaDadosSubElementos()
    txtItemDespSubElemento.Text = ""
    dbcItemDespSubElemento.Text = ""
    txtDblValorSubElemento.Text = ""
End Sub

Private Sub txtDblValorSubElemento_GotFocus()
    MarcaCampo txtDblValorSubElemento
End Sub

Private Sub txtDblValorSubElemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtDblValorSubElemento
End Sub

Private Sub txtDblValorSubElemento_LostFocus()
    txtDblValorSubElemento = gstrConvVrDoSql(txtDblValorSubElemento)
End Sub

Private Sub dbcItemDespSubElemento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cmd_ItemDespSubElemento_Click()
    CarregaForm frmCadItemDespesa, dbcItemDespSubElemento
End Sub

Private Sub dbcItemDespCodSubElementoEst_Click()
    On Error Resume Next
    dbcItemDespSubElementoEst.BoundText = dbcItemDespCodSubElementoEst.ItemData(dbcItemDespCodSubElementoEst.ListIndex)
End Sub


Private Sub dbcItemDespSubElementoEst_Change()
    On Error Resume Next
    If Val(dbcItemDespSubElementoEst.BoundText) > 0 Then
        dbcItemDespCodSubElementoEst.ListIndex = gintIndiceCBO(dbcItemDespCodSubElementoEst, dbcItemDespSubElementoEst.BoundText)
    Else
        dbcItemDespCodSubElementoEst.ListIndex = -1
    End If
End Sub


Private Sub dbcItemDespSubElementoEst_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        dbcItemDespCodSubElementoEst.ListIndex = gintIndiceCBO(dbcItemDespCodSubElementoEst, dbcItemDespSubElementoEst.BoundText)
    End If
End Sub

Private Sub dbcItemDespSubElementoEst_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cmd_ItemDespSubElementoEst_Click()
    CarregaForm frmCadItemDespesa, dbcItemDespSubElementoEst
End Sub

Private Sub txtDblValorSubElementoEst_GotFocus()
    MarcaCampo txtDblValorSubElementoEst
End Sub

Private Sub txtDblValorSubElementoEst_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtDblValorSubElementoEst
End Sub

Private Sub txtDblValorSubElementoEst_LostFocus()
    txtDblValorSubElementoEst = gstrConvVrDoSql(txtDblValorSubElementoEst)
End Sub

Private Sub CarregaCombosSubElementosEst(ByVal strPkidEmpenho As String)
   Dim strSQL       As String
   Dim strAux       As String
   Dim i            As Integer
   
   strSQL = "SELECT "
   strSQL = strSQL & " ID.PKID, "
   strSQL = strSQL & " ID.strDescricao, "
   strSQL = strSQL & " ID.strCodigo "
   strSQL = strSQL & " FROM "
   strSQL = strSQL & gstrItemDespesa & " ID ,"
   strSQL = strSQL & gstrSubElementoEmpenho & " SE "
   strSQL = strSQL & " WHERE "
   strSQL = strSQL & " ID.PKID = SE.intItemDespesa "
   strSQL = strSQL & " AND SE.intEmpenho  = " & strPkidEmpenho
   
   dbcItemDespSubElementoEst.Tag = strSQL & " ORDER BY ID.strDescricao;strDescricao "
   
   PreencherListaDeOpcoes dbcItemDespSubElementoEst
   
   dbcItemDespSubElementoEst.Tag = ""
   
   strAux = "SELECT "
   strAux = strAux & " ID.PKID, "
   strAux = strAux & " ID.strCodigo "
   strAux = strAux & " FROM "
   strAux = strAux & gstrItemDespesa & " ID ,"
   strAux = strAux & gstrSubElementoEmpenho & " SE "
   strAux = strAux & " WHERE "
   strAux = strAux & " ID.PKID = SE.intItemDespesa "
   strAux = strAux & " AND SE.intEmpenho  = " & strPkidEmpenho
   strAux = strAux & " ORDER BY ID.strCodigo "
   
   LeDaTabelaParaObj "", dbcItemDespCodSubElementoEst, strAux
            
   For i = 0 To dbcItemDespCodSubElementoEst.ListCount - 1
        dbcItemDespCodSubElementoEst.list(i) = gvntFormatacaoEspecifica(dbcItemDespCodSubElementoEst.list(i), 4)
   Next
   
                
End Sub

Private Function blnDadosSubElementos() As Boolean
    Dim i As Integer
    Dim dblvalorAcumulado  As Double
    
    For i = 1 To lvwSubElemento.ListItems.Count
        dblvalorAcumulado = dblvalorAcumulado + Val(gstrConvVrParaSql(lvwSubElemento.ListItems(i).SubItems(3)))
    Next
    
    If Val(dbcItemDespSubElemento.BoundText) = 0 Then
        ExibeMensagem "É necessário selecionar um sub-Elmento válido."
        If txtDblValorSubElemento.Enabled Then txtDblValorSubElemento.SetFocus
        Exit Function
    End If

    If gblnEncontroItemNoListView(lvwSubElemento, CStr(dbcItemDespSubElemento.BoundText), 2, 0) = True Then
        ExibeMensagem "O Item de Despesa selecionado já se encontra na lista."
        If txtDblValorSubElemento.Enabled Then txtDblValorSubElemento.SetFocus
        Exit Function
    End If


    If Val(gstrConvVrParaSql(txtDblValorSubElemento)) = 0 Then
        ExibeMensagem "O valor não pode ser Zero."
        If txtDblValorSubElemento.Enabled Then txtDblValorSubElemento.SetFocus
        Exit Function
    End If

    If Val(gstrConvVrParaSql(dblvalorAcumulado + Val(gstrConvVrParaSql(txtDblValorSubElemento)))) > Val(gstrConvVrParaSql(txtdblValor)) Then
        ExibeMensagem "O valor acumulado dos Sub-Elementos não pode ser superior ao valor do Empenho."
        If txtDblValorSubElemento.Enabled Then txtDblValorSubElemento.SetFocus
        Exit Function
    End If

    blnDadosSubElementos = True
End Function

Private Sub IncluirSubelementoNoGrid()
    Set mobjLista = lvwSubElemento.ListItems.Add(, , dbcItemDespSubElemento.BoundText)
    mobjLista.SubItems(1) = txtItemDespSubElemento
    mobjLista.SubItems(2) = dbcItemDespSubElemento.Text
    mobjLista.SubItems(3) = gstrConvVrDoSql(txtDblValorSubElemento)
    mobjLista.Tag = dbcItemDespSubElemento.BoundText
    limpaDadosSubElementos
End Sub

Private Function ExcluirSubElementosNoGrid()
    With lvwSubElemento
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
End Function

Private Function ExcluirSubElementosEstNoGrid()
    With lvwSubElementoEst
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
        End If
    End With
End Function

Private Function VerificaSaldoSubElementoEst(ByVal strPkidEmpenho As String, ByVal strPkiditemdespesa) As Double
   Dim strSQL       As String
   Dim adoResultado As New ADODB.Recordset
   
   strSQL = " SELECT SUM (TMP.DBLVALOR) DBLSALDO FROM ("
   strSQL = strSQL & "SELECT SUM(DBLVALOR) DBLVALOR "
   strSQL = strSQL & " FROM "
   strSQL = strSQL & gstrSubElementoEmpenho
   strSQL = strSQL & " WHERE "
   strSQL = strSQL & " intEmpenho = " & strPkidEmpenho
   strSQL = strSQL & " AND intItemDespesa = " & strPkiditemdespesa
   
   strSQL = strSQL & " UNION ALL "
   strSQL = strSQL & "SELECT " & gstrISNULL("SUM(SU.DBLVALOR)*-1", "0")
   strSQL = strSQL & " FROM "
   strSQL = strSQL & gstrSubElementoEmpenho & " SU, "
   strSQL = strSQL & gstrSubempenho & " SE "
   strSQL = strSQL & " WHERE "
   strSQL = strSQL & " SU.intParcela = SE.PKID "
   strSQL = strSQL & " AND SE.intEmpenho = " & strPkidEmpenho
   strSQL = strSQL & " AND intItemDespesa = " & strPkiditemdespesa
   strSQL = strSQL & " AND SE.intNumero = 0 "
   strSQL = strSQL & " AND SE.bytSituacao = 4 "
   strSQL = strSQL & " )TMP "
   
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            VerificaSaldoSubElementoEst = !dblSaldo
         End If
      End With
   End If

    
End Function

Private Function blnDadosSubElementosEst() As Boolean
    Dim i As Integer
    Dim dblvalorAcumulado  As Double
    
    
    For i = 1 To lvwSubElementoEst.ListItems.Count
        dblvalorAcumulado = dblvalorAcumulado + Val(gstrConvVrParaSql(lvwSubElementoEst.ListItems(i).SubItems(3)))
    Next
    
    If Val(dbcItemDespSubElementoEst.BoundText) = 0 Then
        ExibeMensagem "É necessário selecionar um sub-Elmento válido."
        If dbcItemDespCodSubElementoEst.Enabled Then dbcItemDespCodSubElementoEst.SetFocus
        Exit Function
    End If

    If gblnEncontroItemNoListView(lvwSubElementoEst, CStr(dbcItemDespSubElementoEst.BoundText), 2, 0) = True Then
        ExibeMensagem "O Item de Despesa selecionado já se encontra na lista."
        If txtDblValorSubElementoEst.Enabled Then txtDblValorSubElementoEst.SetFocus
        Exit Function
    End If

    If Val(gstrConvVrParaSql(txtDblValorSubElementoEst)) = 0 Then
        ExibeMensagem "O valor não pode ser Zero."
        If txtDblValorSubElementoEst.Enabled Then txtDblValorSubElementoEst.SetFocus
        Exit Function
    End If

    If Val(gstrConvVrParaSql(txtDblValorSubElementoEst.Text)) > Val(gstrConvVrParaSql(VerificaSaldoSubElementoEst(txtPKId, dbcItemDespSubElementoEst.BoundText))) Then
        ExibeMensagem "O valor não pode ser maior que o saldo restante para anular deste item de despesa (" & gstrConvVrDoSql(VerificaSaldoSubElementoEst(txtPKId, dbcItemDespSubElementoEst.BoundText)) & ") "
        If txtDblValorSubElementoEst.Enabled Then txtDblValorSubElementoEst.SetFocus
        Exit Function
    End If

    If dblvalorAcumulado + Val(gstrConvVrParaSql(txtDblValorSubElementoEst)) > Val(gstrConvVrParaSql(txt_ValorAnulacao)) Then
        ExibeMensagem "O valor acumulado dos Sub-Elementos não pode ser superior ao valor a ser Anulado."
        If txtDblValorSubElementoEst.Enabled Then txtDblValorSubElementoEst.SetFocus
        Exit Function
    End If

    blnDadosSubElementosEst = True
End Function


Private Sub IncluirSubelementoEstNoGrid()
    Set mobjLista = lvwSubElementoEst.ListItems.Add(, , dbcItemDespSubElementoEst.BoundText)
    mobjLista.SubItems(1) = dbcItemDespCodSubElementoEst.Text
    mobjLista.SubItems(2) = dbcItemDespSubElementoEst.Text
    mobjLista.SubItems(3) = gstrConvVrDoSql(txtDblValorSubElementoEst)
    mobjLista.Tag = dbcItemDespSubElementoEst.BoundText
    limpaDadosSubElementosEst
End Sub


Private Sub limpaDadosSubElementosEst()
    dbcItemDespCodSubElementoEst.ListIndex = -1
    dbcItemDespSubElementoEst.Text = ""
    txtDblValorSubElementoEst.Text = ""
    TrocaCorObjeto txtDblValorSubElementoEst, False
End Sub



Private Sub VerificaTabParaIncluir()


    If tab_3dPasta.Tab = 0 And tab_3DGeral.Tab = 0 And tab_3DEmpenho.Tab = 2 Then
        If blnDadosSubElementos = True Then
            IncluirSubelementoNoGrid
        End If
        Exit Sub
    End If
    
    If tab_3dPasta.Tab = 4 And tab_3DAnulacao.Tab = 0 Then
        If blnDadosSubElementosEst = True Then
            IncluirSubelementoEstNoGrid
        End If
        Exit Sub
    End If
    
    Select Case tab_3dPasta.Tab
        Case 0
            
            If blnDadosItemOK = True Then
                IncluirItemNoGrid
                If txt_intCodigo.Enabled Then txt_intCodigo.SetFocus
                CalculaSubTotalItem
                LimpaDadosItem
            End If
        Case 3
            Select Case tab_3DPastaLiquidacao.Tab
                Case 1
                    IncluiExtra
                Case 3
                    IncluiOrcamentario
                Case 4
                    VerificaNF
            End Select
        Case 4
            IncluirItemNoGridAnulacao
    End Select
End Sub
Private Sub VerificaTabParaExcluir()
    
    If tab_3DGeral.Tab = 0 And tab_3DEmpenho.Tab = 2 And Not mblnAlterandoEmpenho Then
        ExcluirSubElementosNoGrid
    End If
    
    If tab_3DAnulacao.Tab = 0 And tab_3dPasta.Tab = 4 Then
        ExcluirSubElementosEstNoGrid
    End If
    
    Select Case tab_3dPasta.Tab
        Case 0
            ExcluirItemNoGrid
        Case 3
            Select Case tab_3DPastaLiquidacao.Tab
            Case 1
               ExcluiExtra
            Case 3
               ExcluiOrcamentario
            Case 4
               VerificaNFExcluir
            End Select
        Case 4
            Select Case tab_3DAnulacao.Tab
            Case 1
                ExcluirItemNoGridAnulacao
            End Select
    End Select
End Sub
Private Sub IncluiExtra()
   Dim i As Integer
   If blnDadosExtraOk Then ' And blnDadoLiquidacaoOK Then
      If mblnAlterandoExtra Then
         lvw_Extra.SelectedItem.Text = cbo_ContaExtra.Text
         lvw_Extra.SelectedItem.SubItems(1) = cbo_DescricaoExtra
         lvw_Extra.SelectedItem.SubItems(2) = gstrConvVrDoSql(txt_ValorExtra)
         lvw_Extra.SelectedItem.SubItems(3) = "1"
         lvw_Extra.SelectedItem.Tag = gstrItemData(cbo_ContaExtra)
      Else
         Set mobjLista = lvw_Extra.ListItems.Add(, , cbo_ContaExtra.Text)
         mobjLista.SubItems(1) = cbo_DescricaoExtra
         mobjLista.SubItems(2) = gstrConvVrDoSql(txt_ValorExtra)
         mobjLista.SubItems(3) = "0"
         mobjLista.Tag = gstrItemData(cbo_ContaExtra)
         LimpaDadosExtra
      End If
   End If
   
   CalculaDescontos
   
End Sub
Private Sub ExcluiExtra()
    With lvw_Extra
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
            mblnAlterandoExtra = False
            LimpaDadosExtra
        End If
    End With
    
    CalculaDescontos
    
End Sub

Private Sub CalculaDescontos()
   Dim strDescontoAnt As String
   
   strDescontoAnt = txt_dblDesconto
   txt_dblDesconto = "0"
   Dim i As Integer
   
   For i = 1 To lvw_Orcamentario.ListItems.Count
        txt_dblDesconto = Val(gstrConvVrParaSql(txt_dblDesconto)) + Val(gstrConvVrParaSql(lvw_Orcamentario.ListItems(i).SubItems(2)))
   Next
   txt_dblDesconto = gstrConvVrDoSql(txt_dblDesconto)


    If lvw_Orcamentario.ListItems.Count = 0 Then
        txt_dblDesconto = strDescontoAnt
    End If

   lblExtra = "0"
   For i = 1 To lvw_Extra.ListItems.Count
        lblExtra = Val(gstrConvVrParaSql(lblExtra)) + Val(gstrConvVrParaSql(lvw_Extra.ListItems(i).SubItems(2)))
   Next
   lblExtra = gstrConvVrDoSql(lblExtra)
End Sub

Private Sub IncluiOrcamentario()
   
   If blnDadosOrcamentarioOk Then
      If mblnAlterandoOrcamentario Then
         With lvw_Orcamentario
            .SelectedItem.Text = cbo_ContaOrcamentario.Text
            .SelectedItem.SubItems(1) = cbo_DescricaoOrcamentario
            .SelectedItem.SubItems(2) = gstrConvVrDoSql(txt_ValorOrcamentario)
            .SelectedItem.SubItems(3) = "1"
            .SelectedItem.Tag = gstrItemData(cbo_ContaOrcamentario)
         End With
      Else
         Set mobjLista = lvw_Orcamentario.ListItems.Add(, , cbo_ContaOrcamentario.Text)
         mobjLista.SubItems(1) = cbo_DescricaoOrcamentario
         mobjLista.SubItems(2) = gstrConvVrDoSql(txt_ValorOrcamentario)
         mobjLista.SubItems(3) = "0"
         mobjLista.Tag = gstrItemData(cbo_ContaOrcamentario)
         LimpaDadosOrcamentario
      End If
   End If

   CalculaDescontos

End Sub
Private Sub ExcluiOrcamentario()
With lvw_Orcamentario
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
            mblnAlterandoOrcamentario = False
            LimpaDadosOrcamentario
        End If
    End With
    CalculaDescontos
End Sub

Private Function GeraProximoDeEmpenho(Optional blnNaoAtualizaData As Boolean) As Long
   Dim strSQL       As String
   Dim adoResultado As New ADODB.Recordset
   
   strSQL = "SELECT " & gstrISNULL("MAX(intNumero)", "0") & " AS Codigo FROM " & gstrEmpenho & " "
   'strSQL = strSQL & " WHERE " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio
   strSQL = strSQL & " WHERE intExercicioEmpenho = " & gintExercicio
   strSQL = strSQL & " UNION SELECT " & gstrISNULL("MAX(intEmpenhoAnulacao)", "0") & " AS Codigo FROM " & gstrSubempenho & " "
   strSQL = strSQL & " WHERE " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & " "
   strSQL = strSQL & " ORDER BY Codigo DESC "
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            GeraProximoDeEmpenho = !Codigo
         End If
      End With
   End If
   
   If IsMissing(blnNaoAtualizaData) Or blnNaoAtualizaData = False Then
        ProximaData
   End If
   
   GeraProximoDeEmpenho = GeraProximoDeEmpenho + 1
   
End Function

Private Function BuscaEmpenho(Optional strExercicioRP As Variant) As Long
   Dim strSQL As String
   Dim adoResultado As New ADODB.Recordset
   
   mblnEmpenhoEstorno = False
   If Len(Trim(txtintNumero)) > 0 Then
      If IsMissing(strExercicioRP) Then
         strSQL = "SELECT PKID FROM " & gstrEmpenho
         strSQL = strSQL & " WHERE intNumero = " & txtintNumero
         'Orc1376
         'strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio
         strSQL = strSQL & " AND intExercicioEmpenho = " & gintExercicio
      ElseIf strExercicioRP = "" Then
         strSQL = "SELECT PKID FROM " & gstrEmpenho
         strSQL = strSQL & " WHERE intNumero = " & txtintNumero
         'Orc1376
         'strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio
         strSQL = strSQL & " AND intExercicioEmpenho = " & gintExercicio
      Else
         strSQL = "SELECT EP.PKID FROM " & gstrEmpenho & " EP, "
         strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
         strSQL = strSQL & " WHERE EP.intNumero = " & txtintNumero
         strSQL = strSQL & " AND EP.intProgramaTrabalho = PT.PKID "
         strSQL = strSQL & " AND PT.intExercicio = " & strExercicioRP
      End If
      
      Set gobjBanco = New clsBanco
      If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
         With adoResultado
            If Not .EOF Then
               BuscaEmpenho = !Pkid
            Else
               adoResultado.Close
               
               strSQL = "SELECT intEmpenho FROM " & gstrSubempenho
               strSQL = strSQL & " WHERE intEmpenhoAnulacao = " & txtintNumero
               strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio
               If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                  With adoResultado
                     If Not .EOF Then
                        BuscaEmpenho = !INTEMPENHO
                        mblnEmpenhoEstorno = True
                     Else
                        BuscaEmpenho = Empty
                     End If
                  End With
               End If
            End If
         End With
      End If
   Else
      BuscaEmpenho = Empty
   End If
End Function

Private Sub VerificaTabDeletaAnulacao()
   Dim strSQL           As String
   Dim adoResultado     As New ADODB.Recordset
   Dim dblValorExcluido As Double
   
   strSQL = "SELECT intEmpenho, dblValor FROM " & gstrSubempenho
   strSQL = strSQL & " WHERE PKID = " & lvw_Anulacao.SelectedItem.Tag
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            strSQL = "UPDATE " & gstrSubempenho & " SET dblValor = dblValor + " & gstrConvVrParaSql(!dblValor)
            strSQL = strSQL & " WHERE intEmpenho = " & !INTEMPENHO & " AND bytSituacao = 1 "
            strSQL = strSQL & " AND intNumero = 0 "
            If gobjBanco.Execute(strSQL) Then
               strSQL = "DELETE FROM " & gstrSubempenho & " WHERE PKID = " & lvw_Anulacao.SelectedItem.Tag
               If gobjBanco.Execute(strSQL) Then
                  LeSubEmpenho lvw_Anulacao, 4, txtPKId
                  ExibeMensagem "Anulação excluída com sucesso."
               End If
            End If
         End If
      End With
   End If
   
   LimpaTelaAnulacao
   
End Sub
Private Function VerificaOrdemDePagamento() As Boolean
   Dim strSQL       As String
   Dim adoResultado As ADODB.Recordset
   
   strSQL = "SELECT OPE.* FROM " & gstrOrdemPagamentoEmpenho & " OPE, "
   strSQL = strSQL & gstrOrdemPagamento & " OP "
   strSQL = strSQL & " WHERE OP.PKID = OPE.intOrdemPagamento AND OPE.intParcela = " & lvw_Liquidacao.SelectedItem.Tag
   strSQL = strSQL & " AND (OP.bytCancelado = 0 OR OP.bytCancelado IS NULL) "
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            VerificaOrdemDePagamento = True
         End If
         .Close
      End With
   End If
   
   strSQL = "SELECT OPR.* FROM " & gstrOrdemPagamentoResto & " OPR, "
   strSQL = strSQL & gstrOrdemPagamento & " OP "
   strSQL = strSQL & " WHERE OP.PKID = OPR.intOrdemPagamento AND OPR.intParcela = " & lvw_Liquidacao.SelectedItem.Tag
   strSQL = strSQL & " AND (OP.bytCancelado = 0 OR OP.bytCancelado IS NULL) "
   
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            VerificaOrdemDePagamento = True
         End If
         .Close
      End With
   End If
   
   
End Function

Private Function VerificaLiqAutomatica() As Boolean
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    VerificaLiqAutomatica = False

    Set gobjBanco = New clsBanco
    strSQL = ""
    strSQL = strSQL & "Select bytGerarEmpenhoLiqAutomatica from " & gstrConfiguracaoGeral & " Where Pkid = 1 "
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                VerificaLiqAutomatica = IIf(Not IsNull(!bytGerarEmpenhoLiqAutomatica), .Fields("bytGerarEmpenhoLiqAutomatica").Value, False)
            End If
        End With
    End If
    
    VerificaLiqAutomatica
    
    Set adoResultado = Nothing
    Set gobjBanco = Nothing
End Function

Private Function VerificaDataAutomatica()
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    strSQL = ""
    strSQL = strSQL & "Select bytGerarEmpenhoDataAutomatica,bytGerarEmpenhoLiqAutomatica from " & gstrConfiguracaoGeral & " Where Pkid = 1 "
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                blnDataAutomatica = IIf(Not IsNull(!bytGerarEmpenhoDataAutomatica), .Fields("bytGerarEmpenhoDataAutomatica").Value, False)
                blnLiqAutomatica = IIf(Not IsNull(!bytGerarEmpenhoLiqAutomatica), .Fields("bytGerarEmpenhoLiqAutomatica").Value, False)
            End If
        End With
    End If
    Set adoResultado = Nothing
    Set gobjBanco = Nothing
End Function

Private Function DataAutomatica()
    Dim adoResultado As ADODB.Recordset
    Dim strSQL       As String
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT dtmFechamento FROM " & gstrFechamentoContabil & " WHERE strCodigo = '" & "EO" & "' AND intExercicio =" & gintExercicio, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            txtDTMDATA = CDate(gstrDataFormatada(adoResultado!dtmFechamento)) + 1
            If Weekday(CDate(gstrDataFormatada(adoResultado!dtmFechamento)) + 1) = 7 Then
                txtDTMDATA = CDate(gstrDataFormatada(adoResultado!dtmFechamento)) + 2
            ElseIf Weekday(CDate(gstrDataFormatada(adoResultado!dtmFechamento)) + 1) = 1 Then
                txtDTMDATA = CDate(gstrDataFormatada(adoResultado!dtmFechamento)) + 1
            End If
            TrocaCorObjeto txtDTMDATA, True
        End If
    End If
    adoResultado.Close
    Set adoResultado = Nothing
    Set gobjBanco = Nothing

End Function

Private Sub CalculaSaldoAtual()
   If Not IsDate(txtDTMDATA) Then Exit Sub
   LeProgramaTrabalho cboProgramaTrabalho, cboCodigoReduzido, _
                     txt_Orgao, txt_Subunidade, txt_Funcao, txt_Programa, _
                     txt_Projetoatividade, txt_UnidadeOrcamentaria, _
                     txt_TipoCredito, txt_Subfuncao, txt_SubPrograma, _
                     txt_ElementoDespesa, txt_SaldoDotacao, _
                     txt_TotalDotado, txt_ValorProgramaTrabalho, , , , , , , , , , txtDTMDATA
End Sub


Private Function IncluirItemNoGrid()
    Dim intInd          As Integer
   
'    With lvw_Itens
'        For intInd = 1 To .ListItems.Count
'            If Trim(txt_intCodigo) = .ListItems(intInd).SubItems(1) Then
'                ExibeMensagem "Não é possível incluir itens iguais"
'                Exit Function
'            End If
'        Next
'    End With


    If Not lvw_Itens.SelectedItem Is Nothing Then
        Set mobjLista = lvw_Itens.SelectedItem
    Else
        Set mobjLista = lvw_Itens.ListItems.Add(, , txt_PkidItem)
    End If
    
    mobjLista.SubItems(1) = txt_intCodigo.Text
    mobjLista.SubItems(2) = txt_intCatalogoMaterialServico.Text
    mobjLista.SubItems(3) = dbc_intStrMarca
    mobjLista.SubItems(4) = txt_dblQuantidade.Text
    mobjLista.SubItems(5) = gstrConvVrDoSql(txt_dblValorEstimado.Text, 5)
    mobjLista.SubItems(6) = txt_intUnidadedeMedida.Text
    mobjLista.SubItems(7) = txt_strObsItem.Text
    mobjLista.SubItems(8) = txt_strdescricaodetalhada.Text
    
    Set lvw_Itens.SelectedItem = Nothing
    
End Function

Private Function ExcluirItemNoGrid()
    With lvw_Itens
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
             
        End If
    End With
End Function
Private Function ValidaItem(Optional blnPreencheCampos As Boolean) As Boolean
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    ValidaItem = False
    strSQL = ""
    strSQL = strSQL & "Select CM.*, UM.Strdescricao as StrdescUnid from "
    strSQL = strSQL & gstrCatalogoMaterialServico & " CM, "
    strSQL = strSQL & gstrUnidadeMedida & " UM "
    strSQL = strSQL & " Where "
    strSQL = strSQL & "CM.Intunidadedemedida " & strOUTJSQLServer & "= UM.Pkid " & strOUTJOracle & " AND "
    strSQL = strSQL & "intcodigo = " & gstrConvVrParaSql(txt_intCodigo)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
                If .RecordCount >= 1 Then
                    If blnPreencheCampos = True Then
                        txt_PkidItem.Text = !Pkid
                        txt_intCatalogoMaterialServico.Text = Trim(gstrENulo(!strDescricao))
                        txt_intUnidadedeMedida.Text = Trim(gstrENulo(!StrdescUnid))
                        txt_strdescricaodetalhada.Text = Trim(gstrENulo(!strDescricaoDetalhada))
                    End If
                Else
                    ExibeMensagem "Código informado não existe"
                    MarcaCampo txt_intCodigo
                    If txt_intCodigo.Enabled Then txt_intCodigo.SetFocus
                    Exit Function
                End If
        End With
    End If
    ValidaItem = True
End Function

Private Function blnDadosItemOK() As Boolean

    blnDadosItemOK = False

    If Trim(txt_intCodigo) = "" Then
        ExibeMensagem "Código do Item tem que ser informado"
        If txt_intCodigo.Enabled Then txt_intCodigo.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intCatalogoMaterialServico) = "" Then
        ExibeMensagem "A Descrição do Item tem que ser informada "
        If txt_intCatalogoMaterialServico.Enabled Then txt_intCatalogoMaterialServico.SetFocus
        Exit Function
    End If
    
    If Trim(txt_dblQuantidade) = "" Then
        ExibeMensagem "Quantidade de Itens é obrigatória"
        If txt_dblQuantidade.Enabled Then txt_dblQuantidade.SetFocus
        Exit Function
    End If
    If Trim(txt_dblValorEstimado) = "" Then
        ExibeMensagem "Valor Estimado é obrigatório"
        If txt_dblValorEstimado.Enabled Then txt_dblValorEstimado.SetFocus
        Exit Function
    End If
    
    If ValidaItem(False) = False Then
        ExibeMensagem "Informações do Item estão inválidas"
        Exit Function
    End If
    
    blnDadosItemOK = True
End Function

Private Sub LimpaDadosItem()
    txt_PkidItem.Text = ""
    
    txt_intCodigo.Text = ""
    txt_intCatalogoMaterialServico.Text = ""
    dbc_intStrMarca.Text = ""
    txt_dblQuantidade.Text = ""
    txt_dblValorEstimado.Text = ""
    txt_intUnidadedeMedida.Text = ""
    txt_strObsItem.Text = ""
    txt_strdescricaodetalhada.Text = ""
End Sub

Private Function VlTotalItem() As Currency
    Dim intInd          As Integer
    Dim dblTotal        As Double
   
    
    With lvw_Itens
        For intInd = 1 To .ListItems.Count
            dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(5))) * Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(4)))
        Next
        VlTotalItem = gstrConvVrDoSql(CDbl(dblTotal))
    End With
    
End Function

Private Function StrSalvaIten(INTEMPENHO As Long) As String
    Dim strSQL  As String
    Dim intInd  As Integer
    
    strSQL = ""
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    With lvw_Itens
        For intInd = 1 To .ListItems.Count
'            If bytDBType = SQLServer Then
'                strSql = strSql & " declare @PKID int "
'                strSql = strSql & " SET @PKID= (Select Pkid From " & gstrEmpenho & " Where intNumero=" & INTEMPENHO
'                strSql = strSql & " AND  YEAR(dtmData)  = " & gintExercicio & " ) "
'            End If
            strSQL = strSQL & "INSERT INTO "
            strSQL = strSQL & gstrItemEmpenho & " ("
            strSQL = strSQL & "INTEMPENHO, "
            strSQL = strSQL & "INTCODIGOITEM, "
            strSQL = strSQL & "STRDESCRICAOITEM, "
            strSQL = strSQL & "STRMARCA, "
            strSQL = strSQL & "DBLQUANTIDADE, "
            strSQL = strSQL & "DBLPRECOUNITARIO, "
            strSQL = strSQL & "STRUNIDADE, "
            strSQL = strSQL & "STROBSERVACAO, "
            strSQL = strSQL & "STRDESCRICAODETALHADA, "
            strSQL = strSQL & "dtmDtAtualizacao, "
            strSQL = strSQL & "lngCodUsr) "
            strSQL = strSQL & "Values("
'            If bytDBType = Oracle Then
                strSQL = strSQL & glngRetornaPkidTabelaPai("seq" & gstrEmpenho, gstrEmpenho) & "," 'Pkid do Empenho
'            ElseIf bytDBType = SQLServer Then
'                strSql = strSql & "@PKID,"
'            End If
            strSQL = strSQL & .ListItems(intInd).SubItems(1) & ", " 'Código do Item
            strSQL = strSQL & "'" & Replace(.ListItems(intInd).SubItems(2), "'", "´") & "', "                 'Descrição
            strSQL = strSQL & "'" & IIf(.ListItems(intInd).SubItems(3) <> "", .ListItems(intInd).SubItems(3), "") & "', "               'Marca
            strSQL = strSQL & gstrConvVrParaSql(.ListItems(intInd).SubItems(4)) & ", " 'Quantidade
            strSQL = strSQL & gstrConvVrParaSql(.ListItems(intInd).SubItems(5)) & ", " 'Valor Estimado
            strSQL = strSQL & "'" & .ListItems(intInd).SubItems(6) & "', "                   'Unidade Medida
            strSQL = strSQL & "'" & IIf(.ListItems(intInd).SubItems(7) <> "", .ListItems(intInd).SubItems(7), "") & "', "               'Observação
            strSQL = strSQL & "'" & Replace(Mid(.ListItems(intInd).SubItems(8), 1, 3995), "'", "´") & "', "               'Descr. Detalhada (truncado ate 4000 posicoes)
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSQL = strSQL & glngCodUsr & " "
            strSQL = strSQL & ")" & IIf(bytDBType = Oracle, ";", "")
        Next
    End With
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
    StrSalvaIten = strSQL
End Function

Private Sub PreencheGridItem(intPkid As Long)
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "IE.PKID, "
    strSQL = strSQL & "Null, "
    strSQL = strSQL & "IE.intcodigoitem, "
    strSQL = strSQL & "IE.strdescricaoitem, "
    strSQL = strSQL & "IE.strmarca, "
    strSQL = strSQL & "IE.dblquantidade, "
    strSQL = strSQL & "IE.dblprecounitario, "
    strSQL = strSQL & "IE.strUnidade, "
    strSQL = strSQL & "IE.strobservacao, "
    strSQL = strSQL & "IE.strdescricaodetalhada "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrItemEmpenho & " IE "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "IE.INTEMPENHO =" & intPkid
    LeDaTabelaParaObj gstrItemEmpenho, lvw_Itens, strSQL
    PreencherSubTotalItens strSQL
End Sub

Private Function SaldoDotacaoSoEmpenho(ByVal intDotacao As String, ByVal strData As String) As Double
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    'dispara na gravação
    '´parametros: pkid conta, data la de cima
       
        strSQL = " SELECT SUM(TMP.DBLVALOR) DBLVALOR FROM ("

    '+1 - Saldo inicial
        strSQL = strSQL & "SELECT dblValor FROM " & gstrProgramaDeTrabalho
        strSQL = strSQL & " WHERE PKID = " & intDotacao & " AND intExercicio = " & gintExercicio
       
    '+2 - Suplementação reduzida
       strSQL = strSQL & " UNION "
       strSQL = strSQL & "SELECT " & gstrISNULL("SUM(DSR.dblValor)", "0") & " AS dblValor FROM " & gstrSuplementacaoReducao & " SR, " & gstrDotacaoSuplementadaReduzida & " DSR "
       strSQL = strSQL & " WHERE SR.PKID = DSR.intSuplementacaoReducao AND DSR.intProgramaTrabalho = " & intDotacao
       strSQL = strSQL & " AND DSR.bytOperacao = 2 "
       strSQL = strSQL & " AND SR.dtmDataDecreto <= " & gstrConvDtParaSql(strData)

       
    '+3 - Suplementação despesa
       strSQL = strSQL & " UNION "
       strSQL = strSQL & "SELECT " & gstrISNULL("SUM(SRD.dblValor)", "0") & " AS dblValor FROM " & gstrSuplementacaoReducao & " SR, " & gstrSuplementacaoReducaoDespesa & " SRD "
       strSQL = strSQL & " WHERE SR.PKID = SRD.intSuplementacaoReducao AND SRD.intProgramaTrabalho = " & intDotacao
       strSQL = strSQL & " AND SR.dtmDataDecreto <= " & gstrConvDtParaSql(strData)
       
    '-4 - Redução (anulado)
       strSQL = strSQL & " UNION "
       strSQL = strSQL & "SELECT " & gstrISNULL("SUM(DSR.dblValor * -1)", "0") & " AS dblValor FROM " & gstrSuplementacaoReducao & " SR, " & gstrDotacaoSuplementadaReduzida & " DSR "
       strSQL = strSQL & " WHERE SR.PKID = DSR.intSuplementacaoReducao AND DSR.intProgramaTrabalho = " & intDotacao
       strSQL = strSQL & " AND DSR.bytOperacao = 1 "
       strSQL = strSQL & " AND SR.dtmDataDecreto <= " & gstrConvDtParaSql(strData)
       
    '-5 - Reservado
       strSQL = strSQL & " UNION "
       strSQL = strSQL & "SELECT " & gstrISNULL("SUM(dblValor * -1)", "0") & " AS dblValor FROM " & gstrEmpenho
       strSQL = strSQL & " WHERE intReservaDotacao IS NULL AND intProgramaTrabalho = " & intDotacao
       strSQL = strSQL & " AND dtmData <= " & gstrConvDtParaSql(strData)

       
    '+6 - Estornado - anulacao
       strSQL = strSQL & " UNION "
       strSQL = strSQL & "SELECT " & gstrISNULL("SUM(SEP.dblValor)", "0") & " AS dblValor FROM " & gstrSubempenho & " SEP, "
       strSQL = strSQL & gstrEmpenho & " EP "
       strSQL = strSQL & " WHERE EP.intReservaDotacao IS NULL AND EP.PKID = SEP.intEmpenho AND "
       strSQL = strSQL & " EP.intProgramaTrabalho = " & intDotacao & " AND "
       strSQL = strSQL & " SEP.intNumero = 0 AND bytSituacao = 4" 'Estorno
       strSQL = strSQL & " AND SEP.dtmData <= " & gstrConvDtParaSql(strData)
       strSQL = strSQL & ") TMP "
       
       If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
          SaldoDotacaoSoEmpenho = adoResultado!dblValor
       End If
       
       adoResultado.Close
    

End Function


Private Function SalvaEmpenhoCompras(INTEMPENHO As Long) As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " UPDATE " & gstrEmpenhoContrato
    strSQL = strSQL & " SET intPedidoEmpenho=(Select Pkid From " & gstrEmpenho & " Where intNumero=" & INTEMPENHO & " AND " & gstrDATEPART(strYEAR, "dtmData") & " = " & gintExercicio & "  ) "
    strSQL = strSQL & "Where intrequisicaodecompra in(SELECT Pkid From " & gstrRequisicaoCompras & " Where intPedidoEmpenho =" & intNumPedidoEmpenho & ")"
    strSQL = strSQL & IIf(bytDBType = Oracle, ";", "")
    
    strSQL = strSQL & " UPDATE " & gstrRequisicaoCompras
    strSQL = strSQL & " SET strnumeroempenho = '" & INTEMPENHO & "/000." & Mid$(Year(CDate(txtDTMDATA)), 3, 2) & "', dtmdataempenho=" & gstrConvDtParaSql(txtDTMDATA)
    strSQL = strSQL & " Where intPedidoEmpenho = " & intNumPedidoEmpenho
    strSQL = strSQL & IIf(bytDBType = Oracle, ";", "")
    SalvaEmpenhoCompras = strSQL
End Function

Private Function blnEmpenhadoCompras() As Boolean
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    blnEmpenhadoCompras = False
    strSQL = ""
    strSQL = strSQL & "Select * "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrEmpenhoContrato
    strSQL = strSQL & " Where "
    strSQL = strSQL & "intPedidoEmpenho = " & txtPKId
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount >= 1 Then
            blnEmpenhadoCompras = True
        End If
    End If
End Function

Private Function SalvaAnulacaoCompras() As Boolean
    Dim strSQL As String
    SalvaAnulacaoCompras = False
    
    strSQL = ""
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    strSQL = strSQL & " UPDATE " & gstrRequisicaoCompras
    strSQL = strSQL & " SET strnumeroempenho = '', dtmdataempenho= NUll "
    strSQL = strSQL & "Where pkid in(SELECT intrequisicaodecompra From " & gstrEmpenhoContrato & " Where intPedidoEmpenho =" & txtPKId & ")"
    strSQL = strSQL & IIf(bytDBType = Oracle, ";", "")

    strSQL = strSQL & " UPDATE " & gstrEmpenhoContrato
    strSQL = strSQL & " SET intPedidoEmpenho = (SELECT intPedidoEmpenho FROM " & gstrRequisicaoCompras & " WHERE Pkid in(SELECT intrequisicaodecompra From " & gstrEmpenhoContrato & " Where intPedidoEmpenho =" & txtPKId & ") GROUP BY intPedidoEmpenho ) "
    strSQL = strSQL & "Where intpedidoempenho = " & txtPKId
    strSQL = strSQL & IIf(bytDBType = Oracle, ";", "")
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
    
    Set gobjBanco = New clsBanco
    SalvaAnulacaoCompras = gobjBanco.Execute(strSQL)

End Function

Private Function PreencheCboItemAnulado()
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    cbo_intCodigoAnulacao.Clear
    cbo_intCatalogoMaterialServicoAnulacao.Clear
    
    strSQL = ""
    strSQL = strSQL & "SELECT IE.PKid, CM.IntCodigo, CM.StrDescricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrItemEmpenho & " IE, "
    strSQL = strSQL & gstrCatalogoMaterialServico & " CM "
    strSQL = strSQL & " WHERE IE.IntEmpenho= " & txtPKId
    strSQL = strSQL & " AND IE.IntCodigoItem= CM.IntCodigo "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF Then
            tab_3DAnulacao.TabEnabled(1) = False
            Exit Function
        Else
            tab_3DAnulacao.TabEnabled(1) = True
            
            
            cbo_intCodigoAnulacao.AddItem "Todos"
            cbo_intCodigoAnulacao.ItemData(cbo_intCodigoAnulacao.ListCount - 1) = 0
            
            cbo_intCatalogoMaterialServicoAnulacao.AddItem "Todos"
            cbo_intCatalogoMaterialServicoAnulacao.ItemData(cbo_intCatalogoMaterialServicoAnulacao.ListCount - 1) = 0
            
        End If
        While Not adoResultado.EOF
            cbo_intCodigoAnulacao.AddItem adoResultado!intCodigo
            cbo_intCodigoAnulacao.ItemData(cbo_intCodigoAnulacao.ListCount - 1) = adoResultado!Pkid
            
            cbo_intCatalogoMaterialServicoAnulacao.AddItem adoResultado!strDescricao
            cbo_intCatalogoMaterialServicoAnulacao.ItemData(cbo_intCatalogoMaterialServicoAnulacao.ListCount - 1) = adoResultado!Pkid
            adoResultado.MoveNext
        Wend
    End If
    
End Function


Private Sub txt_dblQuantidadeAnulacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_dblQuantidadeAnulacao
    
End Sub

Private Sub txt_dblValorEstimadoAnulacao_GotFocus()
    MarcaCampo txt_dblValorEstimadoAnulacao
    
End Sub


Private Sub txt_dblValorEstimadoAnulacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorEstimadoAnulacao
End Sub

Private Sub txt_dblValorEstimadoAnulacao_LostFocus()
    txt_dblValorEstimadoAnulacao = gstrConvVrDoSql(txt_dblValorEstimadoAnulacao, 5)
End Sub


Private Sub txt_strObsItemAnulacao_GotFocus()
    MarcaCampo txt_strObsItemAnulacao
End Sub

Private Sub txt_strObsItemAnulacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strObsItemAnulacao
    
End Sub

Private Sub lvw_ItensAnulacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    HabilitaDesabilitaTab tab_3DAnulacao, True

    With lvw_ItensAnulacao
        txt_PkidItemAnulacao = .SelectedItem.Text
        'cbo_intCodigoAnulacao.Clear
        txt_intCodigo = Replace(.SelectedItem.SubItems(1), ".", "")
        'cbo_intCodigoAnulacao.AddItem Replace(.SelectedItem.SubItems(1), ".", "")
        'cbo_intCodigoAnulacao.ListIndex = 0
        'cbo_intCatalogoMaterialServicoAnulacao.Clear
        'cbo_intCatalogoMaterialServicoAnulacao.AddItem .SelectedItem.SubItems(2)
        txt_intCatalogoMaterialServico = .SelectedItem.SubItems(2)
        'cbo_intCatalogoMaterialServicoAnulacao.ListIndex = 0
        dbc_intStrMarcaAnulacao = .SelectedItem.SubItems(3)
        txt_dblQuantidadeAnulacao.Text = .SelectedItem.SubItems(4)
        txt_dblValorEstimadoAnulacao.Text = gstrConvVrDoSql(.SelectedItem.SubItems(5), 5)
        txt_intUnidadedeMedidaAnulacao.Text = .SelectedItem.SubItems(6)
        txt_strObsItemAnulacao.Text = .SelectedItem.SubItems(7)
        txt_strdescricaodetalhadaAnulacao.Text = .SelectedItem.SubItems(8)
    End With
        
    VerificaTabAtivo
End Sub

Private Sub LimpaCamposAnulado()
    cbo_intCodigoAnulacao.ListIndex = -1
    cbo_intCatalogoMaterialServicoAnulacao.ListIndex = -1
    dbc_intStrMarcaAnulacao = ""
    txt_dblQuantidadeAnulacao.Text = ""
    txt_dblValorEstimadoAnulacao.Text = ""
    txt_intUnidadedeMedidaAnulacao.Text = ""
    txt_strObsItemAnulacao.Text = ""
    txt_strdescricaodetalhadaAnulacao.Text = ""
    If cbo_intCodigoAnulacao.Enabled Then cbo_intCodigoAnulacao.SetFocus
    'lvw_ItensAnulacao.ListItems.Clear
End Sub


Private Sub PreencheGridItemAnulado(intPkid As Long)
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "IE.PKID, "
    strSQL = strSQL & "NULL, "
    strSQL = strSQL & "IE.intcodigoitem, "
    strSQL = strSQL & "IE.strdescricaoitem, "
    strSQL = strSQL & "IE.strmarca, "
    strSQL = strSQL & "EA.dblquantidade, "
    strSQL = strSQL & "EA.dblprecounitario, "
    strSQL = strSQL & "IE.strUnidade, "
    strSQL = strSQL & "EA.strobservacao, "
    strSQL = strSQL & "IE.strdescricaodetalhada "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrItemEmpenho & " IE, "
    strSQL = strSQL & gstrItemEmpenhoAnulado & " EA "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "IE.PKID = EA.intItemEmpenho "
    strSQL = strSQL & " AND EA.intSubEmpenho = " & intPkid
    LeDaTabelaParaObj gstrItemEmpenho, lvw_ItensAnulacao, strSQL
End Sub

Private Function IncluirItemNoGridAnulacao()
    Dim intInd          As Integer
    Dim strSQL          As String
    Dim objItem         As ListItem
   
    If blnDadosOKItemAnulado Then
    
        If cbo_intCodigoAnulacao.ListIndex = 0 Then
            strSQL = "SELECT "
            strSQL = strSQL & "IE.PKID, "
            strSQL = strSQL & "IE.PKID codItem, "
            strSQL = strSQL & "IE.intcodigoitem, "
            strSQL = strSQL & "IE.strdescricaoitem, "
            strSQL = strSQL & "IE.strmarca, "
            strSQL = strSQL & "IE.dblquantidade - " & gstrISNULL("EA.dblquantidade", "0") & " dblquantidade, "
            strSQL = strSQL & "IE.dblprecounitario, "
            strSQL = strSQL & "IE.strUnidade, "
            strSQL = strSQL & "IE.strobservacao, "
            strSQL = strSQL & "IE.strdescricaodetalhada "
            strSQL = strSQL & "FROM "
            If (bytDBType = Oracle) Then
                strSQL = strSQL & gstrItemEmpenho & " IE, "
                strSQL = strSQL & gstrItemEmpenhoAnulado & " EA "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "IE.PKID " & strOUTJSQLServer & "= EA.intItemEmpenho " & strOUTJOracle
                strSQL = strSQL & " AND "
            ElseIf (bytDBType = SQLServer) Then
                strSQL = strSQL & gstrItemEmpenho & " IE LEFT OUTER JOIN "
                strSQL = strSQL & gstrItemEmpenhoAnulado & " EA  ON (" & "IE.PKID = EA.intItemEmpenho" & ")"
                strSQL = strSQL & " WHERE "
            End If

            strSQL = strSQL & " IE.intEmpenho = " & txtPKId
            strSQL = strSQL & " AND (IE.dblquantidade - " & gstrISNULL("EA.dblquantidade", "0") & ") > 0 "
            LeDaTabelaParaObj gstrItemEmpenho, lvw_ItensAnulacao, strSQL
            For Each objItem In lvw_ItensAnulacao.ListItems
                objItem.Text = objItem.Tag
                objItem.SubItems(4) = Val(gstrConvVrParaSql(objItem.SubItems(4)))
            Next
            
        Else
            Set mobjLista = lvw_ItensAnulacao.ListItems.Add(, , gstrItemData(cbo_intCodigoAnulacao))
            mobjLista.Tag = gstrItemData(cbo_intCodigoAnulacao)
            mobjLista.SubItems(1) = gstrConvVrDoSql(cbo_intCodigoAnulacao.Text, 0, 0)
            mobjLista.SubItems(2) = cbo_intCatalogoMaterialServicoAnulacao.Text
            mobjLista.SubItems(3) = dbc_intStrMarcaAnulacao
            mobjLista.SubItems(4) = txt_dblQuantidadeAnulacao.Text
            mobjLista.SubItems(5) = gstrConvVrDoSql(txt_dblValorEstimadoAnulacao.Text, 5)
            mobjLista.SubItems(6) = txt_intUnidadedeMedidaAnulacao.Text
            mobjLista.SubItems(7) = txt_strObsItemAnulacao.Text
            mobjLista.SubItems(8) = txt_strdescricaodetalhadaAnulacao.Text
            LimpaCamposAnulado
        End If
    End If
End Function

Private Function blnDadosOKItemAnulado() As Boolean
    blnDadosOKItemAnulado = False
    Dim intInd            As Integer
    
'    With lvw_ItensAnulacao
'        For intInd = 1 To .ListItems.Count
'            If (Trim(cbo_intCodigoAnulacao) = .ListItems(intInd).SubItems(1)) Then
'                ExibeMensagem "Não é possível incluir itens iguais"
'                Exit Function
'            End If
'        Next
'    End With

    If gblnEncontroItemNoListView(lvw_ItensAnulacao, gstrItemData(cbo_intCodigoAnulacao), lvwTag) Then
        ExibeMensagem "Não é possível incluir itens iguais"
        Exit Function
    End If
    
    
    If cbo_intCodigoAnulacao.ListIndex = -1 Then
        ExibeMensagem "O item tem que ser informado."
        If cbo_intCodigoAnulacao.Enabled Then cbo_intCodigoAnulacao.SetFocus
        Exit Function
    End If
    
    If cbo_intCodigoAnulacao.ListIndex = 0 Then
        blnDadosOKItemAnulado = True
        Exit Function
    End If
    
    
    If gstrConvVrParaSql(Val(txt_dblQuantidadeAnulacao.Text)) = 0 Then
        ExibeMensagem "A quantidade não pode ser Zero."
        If txt_dblQuantidadeAnulacao.Enabled Then txt_dblQuantidadeAnulacao.SetFocus
        Exit Function
    End If
    
    If gstrConvVrParaSql(txt_dblQuantidadeAnulacao.Text) > (intQuantTotal - intQuantItemAnulada) Then
        ExibeMensagem "A quantidade anulada não pode ser superior a " & CStr(intQuantTotal - intQuantItemAnulada) & "."
        If txt_dblQuantidadeAnulacao.Enabled Then txt_dblQuantidadeAnulacao.SetFocus
        Exit Function
    End If
        
    If Val(gstrConvVrParaSql(txt_dblValorEstimadoAnulacao.Text)) = 0 Then
        ExibeMensagem "A Preço Unitário não pode ser Zero."
        If txt_dblQuantidadeAnulacao.Enabled Then txt_dblQuantidadeAnulacao.SetFocus
        Exit Function
    End If
    
    blnDadosOKItemAnulado = True

End Function


Private Function ExcluirItemNoGridAnulacao()
    With lvw_ItensAnulacao
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
            LimpaCamposAnulado
        End If
    End With
End Function


Private Sub AbreFechaCamposAnulacao(ByVal mblnFechar As Boolean)
    TrocaCorObjeto cbo_intCodigoAnulacao, Not mblnFechar
    TrocaCorObjeto cbo_intCatalogoMaterialServicoAnulacao, Not mblnFechar
    TrocaCorObjeto dbc_intStrMarcaAnulacao, True
    TrocaCorObjeto txt_dblQuantidadeAnulacao, Not mblnFechar
    TrocaCorObjeto txt_dblValorEstimadoAnulacao, Not mblnFechar
    TrocaCorObjeto txt_intUnidadedeMedidaAnulacao, True
    TrocaCorObjeto txt_strObsItemAnulacao, Not mblnFechar
    TrocaCorObjeto txt_strdescricaodetalhadaAnulacao, True
End Sub

Private Function gstrGravaItenAnulado(Optional ByVal strSubEmpenhoPKID As String) As String
    Dim strSQL  As String
    Dim intInd  As Integer
    
    strSQL = ""
    With lvw_ItensAnulacao
    
        For intInd = 1 To .ListItems.Count
            strSQL = strSQL & "INSERT INTO "
            strSQL = strSQL & gstrItemEmpenhoAnulado & " ("
            strSQL = strSQL & "IntSubEmpenho, "
            strSQL = strSQL & "IntItemEmpenho, "
            strSQL = strSQL & "DBLQUANTIDADE, "
            strSQL = strSQL & "STROBSERVACAO, "
            strSQL = strSQL & "DBLPRECOUNITARIO, "
            strSQL = strSQL & "dtmDtAtualizacao, "
            strSQL = strSQL & "lngCodUsr) "
            
            'campo intSubEmpenho
            
            If strSubEmpenhoPKID = "" Then
                strSQL = strSQL & "(Select MAX(SE.Pkid) "
                strSQL = strSQL & ", " & .ListItems(intInd).Text ' PKID itemEmpenho
                strSQL = strSQL & ", " & .ListItems(intInd).SubItems(4) ' quantidade
                strSQL = strSQL & ", '" & .ListItems(intInd).SubItems(7) ' obs
                strSQL = strSQL & "', " & gstrConvVrParaSql(.ListItems(intInd).SubItems(5))  ' preço unitario
                strSQL = strSQL & ", " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & glngCodUsr & " "
                strSQL = strSQL & " FROM " & gstrSubempenho & " SE "
                strSQL = strSQL & " WHERE SE.intEmpenho=" & txtPKId.Text
                strSQL = strSQL & " AND SE.bytSituacao = 4)" 'pkid subEmpenho
                strSQL = strSQL & IIf(bytDBType = Oracle, ";", "")
            Else
            strSQL = strSQL & "Values("
            strSQL = strSQL & strSubEmpenhoPKID
            strSQL = strSQL & ", " & .ListItems(intInd).Text ' PKID itemEmpenho
            strSQL = strSQL & ", " & .ListItems(intInd).SubItems(4) ' quantidade
            strSQL = strSQL & ", '" & .ListItems(intInd).SubItems(7) ' obs
            strSQL = strSQL & "', " & gstrConvVrParaSql(.ListItems(intInd).SubItems(5))  ' preço unitario
            strSQL = strSQL & ", " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSQL = strSQL & glngCodUsr & " "
            strSQL = strSQL & ")" & IIf(bytDBType = Oracle, ";", "")
            End If
        Next
    End With
    gstrGravaItenAnulado = strSQL
End Function

Private Function CarregaItemAnulado(Optional blnPreencheCampos As Boolean) As Boolean
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    If cbo_intCodigoAnulacao.ListIndex = -1 Then Exit Function
    
    If cbo_intCodigoAnulacao.ListIndex = 0 Then
        cbo_intCatalogoMaterialServicoAnulacao.ListIndex = 0
        
        txt_intUnidadedeMedidaAnulacao.Text = ""
        txt_strdescricaodetalhadaAnulacao.Text = ""

        dbc_intStrMarcaAnulacao.Text = ""
        txt_dblQuantidadeAnulacao.Text = ""
        txt_dblValorEstimadoAnulacao.Text = ""
        txt_strObsItemAnulacao.Text = ""
        
        CarregaItemAnulado = True
        Exit Function
    End If
    
    
    'ValidaItem = False
    strSQL = ""
    strSQL = strSQL & "SELECT CM.*, UM.Strdescricao as StrdescUnid FROM "
    strSQL = strSQL & gstrCatalogoMaterialServico & " CM, "
    strSQL = strSQL & gstrUnidadeMedida & " UM "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & "CM.Intunidadedemedida " & strOUTJSQLServer & "= UM.Pkid " & strOUTJOracle & " AND "
    strSQL = strSQL & "intcodigo = " & cbo_intCodigoAnulacao
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                If blnPreencheCampos = True Then
                    cbo_intCatalogoMaterialServicoAnulacao.ListIndex = gintIndiceCBO(cbo_intCatalogoMaterialServicoAnulacao, gstrItemData(cbo_intCodigoAnulacao))
                    txt_intUnidadedeMedidaAnulacao.Text = Trim(gstrENulo(!StrdescUnid))
                    txt_strdescricaodetalhadaAnulacao.Text = Trim(gstrENulo(!strDescricaoDetalhada))
                End If
            Else
                ExibeMensagem "Código informado não existe"
                If cbo_intCodigoAnulacao.Enabled Then cbo_intCodigoAnulacao.SetFocus
                Exit Function
            End If
        End With
    End If
    
    
    strSQL = ""
    strSQL = strSQL & "SELECT IE.strMarca, IE.dblQuantidade ,"
    strSQL = strSQL & " IE.dblPrecoUnitario, IE.strObservacao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrItemEmpenho & " IE "
    strSQL = strSQL & " WHERE IE.PKID = " & gstrItemData(cbo_intCodigoAnulacao)
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .RecordCount >= 1 Then
                If blnPreencheCampos = True Then
                    dbc_intStrMarcaAnulacao.Text = IIf(IsNull(!Strmarca), "", !Strmarca)
                    txt_dblQuantidadeAnulacao.Text = !dblQuantidade
                    txt_dblValorEstimadoAnulacao.Text = gstrConvVrDoSql(!Dblprecounitario, 5)
                    txt_strObsItemAnulacao.Text = IIf(IsNull(!strObservacao), "", !strObservacao)
                End If
            Else
                ExibeMensagem "Código informado não existe"
                If cbo_intCodigoAnulacao.Enabled Then cbo_intCodigoAnulacao.SetFocus
                Exit Function
            End If
        End With
    End If
    
    intQuantItemAnulada = VerificaQuantidadeItemAnulado
    intQuantTotal = CDbl(txt_dblQuantidadeAnulacao.Text)
    txt_dblQuantidadeAnulacao.Text = intQuantTotal - intQuantItemAnulada
    
    CarregaItemAnulado = True
    
End Function

Private Function VerificaQuantidadeItemAnulado() As Long
    Dim adoResultado    As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & gstrISNULL("SUM(EA.dblquantidade)", "0") & " Quantidade"
    'strSql = strSql & "EA.intItemEmpenho), "
    'strSql = strSql & "EA.dblprecounitario, "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrItemEmpenhoAnulado & " EA, "
    strSQL = strSQL & gstrSubempenho & " SE, "
    strSQL = strSQL & gstrEmpenho & " E "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " EA.intSubEmpenho = SE.PKID "
    strSQL = strSQL & " AND E.PKID = SE.intEmpenho "
    strSQL = strSQL & " AND E.PKID = " & txtPKId.Text
    strSQL = strSQL & " AND EA.intItemEmpenho = " & gstrItemData(cbo_intCodigoAnulacao)
    
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            VerificaQuantidadeItemAnulado = adoResultado!Quantidade
        End If
    End If
    
    
End Function

Private Function dblValorDesconto(lngPkidSubEmpenho As Long) As Double
Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset

    strSQL = "SELECT dblDesconto"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrSubempenho
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " Pkid = " & Val(lngPkidSubEmpenho)
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
    
        If Not adoResultado.EOF Then
            dblValorDesconto = gstrConvVrDoSql(Val(gstrConvVrParaSql(adoResultado!dblDesconto)), 2, , True)
        Else
            dblValorDesconto = gstrConvVrDoSql(0, 2)
        End If
        
    End If
    

End Function

Private Function strQueryAplicarEventoAnul() As String
Dim strSQL As String
    
    strSQL = "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEvento & " "
    strSQL = strSQL & "WHERE intTipoEvento = 12 "
    strQueryAplicarEventoAnul = strSQL
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra, , 6-alteracoes orcamentarias
    '              7-Liquidação
    '              12-Cancelamento de Restos à Pagar
End Function

Private Function GravaSubEmpenhoLiquidado() As Boolean
    Dim dblValor            As Double
    Dim dblValorParcAnt     As Double
    Dim dblValorParcAtu     As Double
    Dim intNumDeParcela     As Integer
    Dim lngNumDeParcelaAnt  As Long
    
    
    intNumDeParcela = lvw_ListaSubempenho.ListItems.Count
    With lvw_ListaSubempenho
        dblValorParcAtu = Val(gstrConvVrParaSql(txt_dblValorAux))
        'dblValorParcAnt = ProcuraValorParcAnt(lngNumDeParcelaAnt)
        lngNumDeParcelaAnt = .ListItems(1).Tag
        dblValorParcAnt = CDbl(.ListItems(1).ListSubItems(2).Text)
        dblValor = dblValorParcAnt - dblValorParcAtu
        If blnDadoDaParcelaLiquidadaOk(dblValor) Then
           If gblnExclusaoGravacaoOk(IIf(mblnAlterandoSubEmpenho, "A", "I"), " desta Parcela") Then
               If blnAtualizouParcela(lngNumDeParcelaAnt, _
                                      txt_DataLiuidacao, _
                                      dblValor, _
                                      txt_HistoricoLiquidacao, _
                                      intNumDeParcela) Then
                  'If dblValor > 0 Then
   '                    If lvw_ListaSubempenho.SelectedItem.Index = intNumDeParcela Then
                          'intNumDeParcela = intNumDeParcela + 1
                          'txt_DataParcela = DateAdd("m", 1, txt_DataParcela)
                          If blnIncluiuParcela(txt_DataLiuidacao, _
                                               txt_dblValorAux, _
                                               txt_HistoricoLiquidacao, 1) = False Then
                              PreencheSaldoEmpenho
                              Exit Function
                          End If
   '                    Else
   '                        If gblnEncontroItemNoListView(lvw_ListaSubempenho, _
   '                                                      .ListItems(.SelectedItem.Index + 1).Tag, _
   '                                                      lvwTag) Then
   '                            txtdtmdataParcela = .ListItems(.SelectedItem.Index).SubItems(1)
   '                            dblValorParcAnt = Val(gstrConvVrParaSql(.ListItems(.SelectedItem. _
   '                                              Index).SubItems(2)))
   '                            dblValor = dblValor + dblValorParcAnt
   '                            If blnAtualizouParcela(lvw_ListaSubempenho.SelectedItem.Tag, _
   '                                                   txtdtmdataParcela, _
   '                                                   dblValor, _
   '                                                   txtstrhistoricoSubEmpenho) = False Then
   '                                Exit Sub
   '                            End If
   '                        End If
   '                    End If
                  'End If
                   'OrganizaNumParcelas
                   LeSubEmpenho lvw_ListaSubempenho, , txtPKId
                   PreencheSaldoEmpenho
                   LeSubEmpenho lvw_Liquidacao, 2, txtPKId
                   LeSubEmpenho lvw_Anulacao, 4, txtPKId
                   GravaSubEmpenhoLiquidado = True
               End If
           End If
        End If
    End With
End Function


Private Function blnDadoDaParcelaLiquidadaOk(dblValor As Double) As Boolean
    Dim dtmDtEncerramento As Date
    
    If UCase(lvw_Liquidacao.ListItems(1).SubItems(4)) = "PAGA" Then
        ExibeMensagem "Não é possível inserir parcelas quando o empenho está Pago."
        Exit Function
    ElseIf lvw_Liquidacao.ListItems(1).SubItems(4) = "Liquidada" Then
        ExibeMensagem "Não é possível inserir parcelas quando o empenho está liquidado."
        Exit Function
    ElseIf dblValor < 0 Then
        ExibeMensagem "A soma das parcelas não pode superar o valor do empenhado."
        If txt_ValorParcela.Enabled Then txt_ValorParcela.SetFocus
        Exit Function
    ElseIf gblnDataValida(txt_DataLiuidacao) = False Then
        ExibeMensagem "Data da Liquidação incorreta."
        If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
        Exit Function
     'A mensagem de critica esta dentro da rotina que é usada tambem na guia de liquidação
     ElseIf gblnMaiorSubEmpenhoLiq(txt_DataLiuidacao) = False Then
            Exit Function
    ElseIf CVDate(txtDTMDATA) > CVDate(txt_DataLiuidacao) Then
        ExibeMensagem "Data da Liquidacao não poder ser inferior a data do empenho."
        If txt_DataParcela.Enabled Then txt_DataParcela.SetFocus
        Exit Function
    ElseIf gblnDataValida(txt_DataVencto) = False Then
        ExibeMensagem "A Data do Vencimento é inválida."
        If txt_DataVencto.Enabled Then txt_DataVencto.SetFocus
        Exit Function
    ElseIf CVDate(txt_DataLiuidacao) > CVDate(txt_DataVencto) Then
        ExibeMensagem "A Data do Vencimento não poder ser inferior a data da Liquidação."
        If txt_DataVencto.Enabled Then txt_DataVencto.SetFocus
        Exit Function
    End If
    
    dtmDtEncerramento = VerificaDataEncerramento("EO", gintExercicio)
        
    If dtmDtEncerramento = Empty Then
       Exit Function
    Else
       If CDate(txt_DataLiuidacao) <= dtmDtEncerramento Then
          ExibeMensagem "A data da Liquidação deve ser maior que a data de último encerramento Orçamentário (" & dtmDtEncerramento & ")."
          If txt_DataLiuidacao.Enabled Then txt_DataParcela.SetFocus
          Exit Function
       End If
    End If
    
    
    dtmDtEncerramento = VerificaDataEncerramento("EF", gintExercicio)
        
    If dtmDtEncerramento = Empty Then
       Exit Function
    Else
       If CDate(txt_DataLiuidacao) <= dtmDtEncerramento Then
          ExibeMensagem "A data da Liquidação deve ser maior que a data de último encerramento Financeiro (" & dtmDtEncerramento & ")."
          If txt_DataLiuidacao.Enabled Then txt_DataParcela.SetFocus
          Exit Function
       End If
    End If
    
    'ORC677
    If IsDate(txt_DataLiuidacao.Text) Then
        If Year(CDate(txt_DataLiuidacao.Text)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data da liquidação tem que estar no exercício de " & gintExercicio & "."
            If txt_DataLiuidacao.Enabled Then txt_DataLiuidacao.SetFocus
            Exit Function
        End If
    End If
    
    blnDadoDaParcelaLiquidadaOk = True
End Function

'Criada por M4RC3LØ 14/03/2004
Private Sub AlteraEmpenho()
Dim strSQL        As String

If blnDadosOk Then
    If gblnExclusaoGravacaoOk(IIf(mblnAlterandoEmpenho, "A", "I"), " do Empenho") Then
        If mblnAlterandoEmpenho Then
            strSQL = strQueryAlteraEmpenho
        End If
        Set gobjBanco = New clsBanco
        gobjBanco.Execute (strSQL)
    End If
End If
End Sub

Private Function ImprimeCancelamentoRP(lngParcela As Long)
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "CT.strNome,"
    strSQL = strSQL & "CT.CDC intContribuinte,"
    strSQL = strSQL & "CT.strCNPJCPF,"
    strSQL = strSQL & "CT.strLogradouroC strEndereco,"
    strSQL = strSQL & "CT.intNumero,"
    strSQL = strSQL & "CT.strComplemento strComplemento,"
    strSQL = strSQL & "MP.strDescricao strMunicipio,"
    strSQL = strSQL & "UF.strSigla strUF,"
    strSQL = strSQL & "BR.strDescricao strBairro,"
    strSQL = strSQL & "CP.intCEP,"
    strSQL = strSQL & "EP.intNumero intEmpenho,"
    strSQL = strSQL & "EP.dblValor  dblValorEmpenho,"
    strSQL = strSQL & "EP.strCodigo, "
    strSQL = strSQL & "EP.intExercicio intExercicioProcesso, "
    strSQL = strSQL & "EP.bitDigito, "
    strSQL = strSQL & "EP.Pkid PkidEmpenho,"
    strSQL = strSQL & "SEP.PKID PKIDParcela,"
    strSQL = strSQL & "SEP.intNumero intParcela,"
    strSQL = strSQL & "SEP.dblValor  dblValorParcela,"
    strSQL = strSQL & "SEP.strHistorico,"
    strSQL = strSQL & "SEP.dtmData,"
    strSQL = strSQL & "FR.strDescricao strFonteRecurso, "
    strSQL = strSQL & "PT.intCodigoReduzido, "
    strSQL = strSQL & "PT.intExercicio, "
    strSQL = strSQL & "PT.strCodigo strDotacao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContribuinte & " CT, "
    strSQL = strSQL & gstrCidade & " MP, "
    strSQL = strSQL & gstrUF & " UF, "
    strSQL = strSQL & gstrBairro & " BR, "
    strSQL = strSQL & gstrCeps & " CP, "
    strSQL = strSQL & gstrEmpenho & " EP, "
    strSQL = strSQL & gstrSubempenho & " SEP, "
    strSQL = strSQL & gstrFonteRecurso & " FR, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "EP.intCredor = CT.PKID AND "
    strSQL = strSQL & "SEP.intEmpenho = EP.PKID   AND "
    strSQL = strSQL & "EP.intProgramaTrabalho = PT.PKID AND "
    strSQL = strSQL & "PT.intFonteRecurso = FR.PKID AND "
    strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
    strSQL = strSQL & "CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & "  CT.intCep AND "
    strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
    strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
    strSQL = strSQL & "SEP.PKID = " & lngParcela & " "
    
    ImprimeRelatorio rptNotaDeCancelamento, strSQL
    
End Function

'Pen_687_ORC_100

Private Sub HabilitaControlesLiquidacao(blnHabilita As Boolean)

    If blnHabilita = True Then
        'Controles da Tab de Liquidação
        TrocaCorObjeto txt_DataLiuidacao, False, False
        TrocaCorObjeto txt_DataVencto, False, False
        TrocaCorObjeto txt_dblValorAux, False, False
        TrocaCorObjeto txt_dblDesconto, True, False
        TrocaCorObjeto txt_HistoricoLiquidacao, False, False
        TrocaCorObjeto cbo_HistoricoLiquidacao, False, False
        TrocaCorObjeto cbo_intEventoLiq, False, False
        TrocaCorObjeto txt_codEventoLiq, False, False
        tab_3DPastaLiquidacao.TabEnabled(0) = True
        lvw_Liquidacao.Visible = True
                
        tab_3DPastaLiquidacao.TabEnabled(1) = True
        cbo_ContaExtra.Visible = True
        cbo_DescricaoExtra.Visible = True
        cmd_ContaExtra.Visible = True
        txt_ValorExtra.Visible = True
        lvw_Extra.Visible = True
        
        tab_3DPastaLiquidacao.TabEnabled(3) = True
        cbo_ContaOrcamentario.Visible = True
        cbo_DescricaoOrcamentario.Visible = True
        cmd_ContaOrcamentario.Visible = True
        txt_ValorOrcamentario.Visible = True
        lvw_Orcamentario.Visible = True
        
        TrocaCorObjeto cmd_EventoLiq, False, False
        TrocaCorObjeto cmd_HistoricoLiquidacao, False, False
        TrocaCorObjeto lvw_NotasFiscais, False, False
    Else
        'Controles da Tab de Liquidação
        TrocaCorObjeto txt_DataLiuidacao, True, True
        TrocaCorObjeto txt_DataVencto, True, True
        TrocaCorObjeto txt_dblValorAux, True, True
        TrocaCorObjeto txt_dblDesconto, True, True
        TrocaCorObjeto txt_HistoricoLiquidacao, True, True
        TrocaCorObjeto cbo_HistoricoLiquidacao, True, True
        TrocaCorObjeto cbo_intEventoLiq, True, False
        TrocaCorObjeto txt_codEventoLiq, True, False
        tab_3DPastaLiquidacao.Tab = 4
        tab_3DPastaLiquidacao.TabEnabled(0) = False
        lvw_Liquidacao.Visible = False
        
        tab_3DPastaLiquidacao.TabEnabled(1) = False
        cbo_ContaExtra.Visible = False
        cbo_DescricaoExtra.Visible = False
        cmd_ContaExtra.Visible = False
        txt_ValorExtra.Visible = False
        lvw_Extra.Visible = False
        
        tab_3DPastaLiquidacao.TabEnabled(3) = False
        cbo_ContaOrcamentario.Visible = False
        cbo_DescricaoOrcamentario.Visible = False
        cmd_ContaOrcamentario.Visible = False
        txt_ValorOrcamentario.Visible = False
        lvw_Orcamentario.Visible = False
        
        tab_3DPastaLiquidacao.TabEnabled(4) = True
        txt_dtmDataNF.Visible = True
            txt_dblValorNF.Visible = True
            txt_strNotasFiscais.Visible = True
            lbl_ValorTotal.Visible = True
            lvw_NotasFiscais.Visible = True
        
        TrocaCorObjeto cmd_EventoLiq, True, True
        TrocaCorObjeto cmd_HistoricoLiquidacao, True, True
        TrocaCorObjeto lvw_NotasFiscais, True, True
        lbl_ValorTotal.Caption = ""
        txt_strNotasFiscais = ""
    End If

End Sub

Private Sub cbo_ContaOrcamentario_Click()
    cbo_DescricaoOrcamentario.ListIndex = gintIndiceCBO(cbo_DescricaoOrcamentario, _
                                  gstrItemData(cbo_ContaOrcamentario))
End Sub

Private Sub cbo_ContaOrcamentario_GotFocus()
   VerificaTabAtivo
   mAtivaPastaDeObjeto tab_3dPasta, 3, tab_3DPastaLiquidacao, 3
End Sub

Private Sub cbo_DescricaoOrcamentario_Click()
    cbo_ContaOrcamentario.ListIndex = gintIndiceCBO(cbo_ContaOrcamentario, _
                              gstrItemData(cbo_DescricaoOrcamentario))
End Sub

Private Sub cbo_DescricaoOrcamentario_GotFocus()
   VerificaTabAtivo
End Sub

Private Sub cmd_ContaOrcamentario_Click()
    CarregaForm frmConPrevisaoDaReceita, cbo_DescricaoOrcamentario
End Sub

Private Sub txt_ValorOrcamentario_GotFocus()
    MarcaCampo txt_ValorOrcamentario
   VerificaTabAtivo
End Sub

Private Sub txt_ValorOrcamentario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_ValorOrcamentario
End Sub
Public Function LeCoditemDespesa(Optional strPKId As String, Optional strCodigo As String)

Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
    
    If Trim(strPKId) = "" And Trim(strCodigo) = "" Then
        LeCoditemDespesa = ""
        Exit Function
    End If
    
    strSQL = ""
    strSQL = strSQL & "SELECT strCodigo , PKID"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrItemDespesa
    If strPKId <> "" Then
        strSQL = strSQL & " WHERE PKID = " & strPKId
    ElseIf strCodigo <> "" Then
        strSQL = strSQL & " WHERE " & strSUBSTRING & " ( strCodigo,1," & Len(gstrValorSemMascara(strCodigo)) & ") = '" & gstrValorSemMascara(strCodigo) & "'"
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            If strPKId <> "" Then
                LeCoditemDespesa = gstrENulo(adoResultado!strCodigo)
            ElseIf strCodigo <> "" Then
               LeCoditemDespesa = gstrENulo(adoResultado!Pkid)
            End If
        End If
        
    End If
End Function

Private Function blnGerarDespesaExtra() As Boolean
   Dim strSQL As String
   Dim adoResultado As ADODB.Recordset
   
   strSQL = "SELECT bytGerarDespesaExtra FROM " & gstrConfiguracaoGeral
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
   
      If gstrENulo(adoResultado!bytGerarDespesaExtra) > 0 Then
         
         blnGerarDespesaExtra = True
         
      End If
   
   End If
   
End Function

Private Function GravaDespesaExtra() As Boolean

   Dim strSQL As String
   Dim intInd As Integer
   Dim intCredor As Long
      
   If gblnExclusaoGravacaoOk("I", " Despesa extra-orçamentária") Then
      strSQL = ""
      strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
      
      With lvw_Extra
         For intInd = 1 To .ListItems.Count
            .ListItems(intInd).Selected = True
            
            intCredor = intCredorExtra(.ListItems(intInd).Tag)
            
            If intCredor = -1 Then
                GravaDespesaExtra = False
                Exit Function
            End If
             
            strSQL = strSQL & "INSERT INTO " & gstrDespesaExtraOrcamentaria & " ("
            strSQL = strSQL & "intNumero, intContribuinte, intContaContabil, "
            strSQL = strSQL & "dblValor, bytSituacao, dtmData, strHistorico, "
            strSQL = strSQL & "intExercicio, dtmDtAtualizacao, lngCodUsr) "
            strSQL = strSQL & " (SELECT MAX(intNumero) + 1 , "
            strSQL = strSQL & intCredor & ", "
            strSQL = strSQL & .ListItems(intInd).Tag & ", "
            strSQL = strSQL & gstrConvVrParaSql(.ListItems(intInd).SubItems(2)) & ", "
            strSQL = strSQL & " 0, "
            strSQL = strSQL & gstrConvDtParaSql(txt_DataLiuidacao) & ", "
            strSQL = strSQL & "'" & Trim(txt_HistoricoLiquidacao) & "', "
            strSQL = strSQL & gintExercicio & ", "
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ","
            strSQL = strSQL & glngCodUsr & " FROM " & gstrDespesaExtraOrcamentaria & ");"
         Next
      
      End With
      
      strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
      
      Set gobjBanco = New clsBanco
      If gobjBanco.Execute(strSQL) Then
         GravaDespesaExtra = True
      Else
         GravaDespesaExtra = False
      End If
       
   End If
      

End Function
Private Function blnModificaParcela(lngChave As Long, _
                                     vntDataParcela, _
                                     vntValorParcela, _
                                     vntHistoricoSubEmpenho, _
                                     Optional intNumParc As Integer) As Boolean
    Dim strSQL      As String
    strSQL = strSQL & "UPDATE " & gstrSubempenho & " SET "
    'strSql = strSql & "dtmData = " & gstrConvDtParaSql(vntDataParcela) & ", "
    'strSQL = strSQL & "dblValor = " & gstrConvVrParaSql(vntValorParcela) & ", "
    strSQL = strSQL & "strHistorico = '" & vntHistoricoSubEmpenho & "', "
    strSQL = strSQL & "dtmDtAtualizacao = " & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSQL = strSQL & "lngCodUsr = " & glngCodUsr & " "
    strSQL = strSQL & "WHERE PKId = " & lngChave
    
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSQL) Then
        blnModificaParcela = True
    End If
End Function

Private Function blnGravaMovLiq(PKIDEmpenho As Long, pkidParcela As Long, pkidProgramaDeTrabalho As Long, ByVal DTMDATA As String, ByVal dblValor As String, ByVal STRHISTORICO As String) As Boolean
   Dim strSQL As String
   
   strSQL = "INSERT INTO " & gstrmovliq
   strSQL = strSQL & " ( intEmpenho, intParcela, intProgramaTrabalho, dtmData, dblValor, strHistorico, dtmDtAtualizacao, lngCodUsr) VALUES "
   strSQL = strSQL & "(" & PKIDEmpenho & ", " & pkidParcela & ", " & pkidProgramaDeTrabalho & ", " & gstrConvDtParaSql(DTMDATA) & ", "
   strSQL = strSQL & gstrConvVrParaSql(dblValor) & ",'" & STRHISTORICO & "', " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & ")"
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.Execute(strSQL) Then
      blnGravaMovLiq = True
   End If
   
End Function

Private Sub mAtivaPastaDeObjeto(mtabPasta1 As SSTab, _
                              mbytTabAtivo1 As Byte, _
                     Optional mtabPasta2 As SSTab, _
                     Optional mbytTabAtivo2 As Byte)
If (mblnAtivarPastas = True) And (Not mtabPasta1 Is Nothing) Then
         If Not mtabPasta2 Is Nothing Then
               AtivaPastaDeObjeto mtabPasta1, mbytTabAtivo1, mtabPasta2, mbytTabAtivo2
         Else
               AtivaPastaDeObjeto mtabPasta1, mbytTabAtivo1
         End If
 End If
End Sub

Private Function intCredorExtra(intConta As Integer) As Long

' Retorna o número do Contribuinte vinculado com a conta informada na tabela CredorExtra
' IMPORTANTE: Se não for encontrado nenhum Contribuinte a função retorna -1

Dim strSQL As String
Dim adoResultado As ADODB.Recordset

    strSQL = "SELECT intCredor FROM " & gstrCredorExtra & " WHERE intPlanoConta = " & intConta
    
    On Error GoTo Tratamento
    
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF And Not adoResultado.BOF Then
            intCredorExtra = adoResultado!intCredor
        Else
            intCredorExtra = -1
        End If
    End If
    
    Exit Function
    
Tratamento:
    intCredorExtra = -1

End Function


Private Function GravaRestos()
    
    If blnDadosRestos Then
    End If
    
End Function

Private Function blnDadosRestos() As Boolean

'    If Trim(txtintExercicioEmpenho) = "" Then
'        ExibeMensagem "É necessário digitar o Exercício"
'        Exit Function
'    End If
'
'    If Not IsDate(txtDtmDada) Then
'        ExibeMensagem "É necessário digitar a Data"
'        Exit Function
'    End If
'
'    If Not IsNumeric(Trim(txtdblValor)) Then
'        If Trim(dblValor) = "" Then
'            ExibeMensagem "É necesságio informar o valor"
'        Else
'            ExibeMensagem "O Valor informado é inválido"
'        End If
'        Exit Function
'    End If
'
'
'
'    blnDadosRestos = True

End Function

Public Sub ProximaData()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
'Orc1376
'            strSQL = "Select dtmData from " & gstrEmpenho & " where intNumero = " & _
'                 "(Select Max(IntNumero) from " & gstrEmpenho & " where " & gstrDATEPART("YYYY", "dtmData") & " = " & CStr(gintExercicio) & ") " & _
'                 " and " & gstrDATEPART("YYYY", "dtmData") & " = " & CStr(gintExercicio)
            
            strSQL = "Select dtmData from " & gstrEmpenho & " where intNumero = " & _
                 "(Select Max(IntNumero) from " & gstrEmpenho & " where intExercicioEmpenho = " & CStr(gintExercicio) & ") " & _
                 " and intExercicioEmpenho = " & CStr(gintExercicio)
                 
        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
           With adoResultado
              If Not .EOF Then
                 txtDTMDATA = adoResultado!DTMDATA
              End If
           End With
        End If
End Sub

Private Function EventoDaReserva() As String
Dim strSQL As String
Dim strElemento As String
Dim strDigito As String
Dim adoResultado As ADODB.Recordset
Dim strResultado As String
Dim intContador As Integer


    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " ED.strCodigoElementoDespesa"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT,"
    strSQL = strSQL & gstrElementoDespesa & " ED"
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " PT.intElementoDespesa = ED.Pkid"
    strSQL = strSQL & " And PT.pkID = " & gstrItemData(cboProgramaTrabalho, True)
    
    Set gobjBanco = New clsBanco
    If Not gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        EventoDaReserva = ""
        Exit Function
    Else
        
        strElemento = adoResultado!strCodigoElementoDespesa
        
        strSQL = ""
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & " EV.PKID,"
        strSQL = strSQL & " EV.strDescricao,"
        strSQL = strSQL & " EVC.intContaContabil,"
        strSQL = strSQL & " PC.strContaContabil"
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrEvento & " EV,"
        strSQL = strSQL & gstrEventoContaContabilDebito & " EVC,"
        strSQL = strSQL & gstrPlanoConta & " PC"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " EV.Pkid = EVC.intEvento"
        strSQL = strSQL & " AND EVC.intContaContabil = PC.PKid"
        strSQL = strSQL & " AND " & strSUBSTRING & "(PC.strContaContabil, 1, " & Len(gstrDigitoDespesa) & ") = '" & gstrDigitoDespesa & "' "
        strSQL = strSQL & " AND EV.intTipoEvento = 2"
        If Not gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            EventoDaReserva = ""
        Else
            strResultado = ""
            While Not adoResultado.EOF
            
                strDigito = Mid(adoResultado!strContaContabil, Len(gstrDigitoDespesa) + 1)
                
                For intContador = 1 To Len(Mid(adoResultado!strContaContabil, Len(gstrDigitoDespesa) + 1))
                    If Mid(strDigito, Len(strDigito), 1) = "0" Or Len(strDigito) > 6 Then
                        strDigito = Mid(strDigito, 1, Len(strDigito) - 1)
                    Else
                        Exit For
                    End If
                                        
                Next
                
                If strDigito = Mid(strElemento, 1, Len(strDigito)) Then
                    strResultado = strResultado & adoResultado!Pkid & ", "
                End If
                adoResultado.MoveNext
                
            Wend
            
            If strResultado <> "" Then
                strResultado = Mid(strResultado, 1, Len(strResultado) - 2)
            End If
            EventoDaReserva = strResultado
        End If
    End If

End Function
Private Function AdicionaGroupByQueryRelatorio() As String
'*****************************************************************************************
'*   Programador:        Éder Henrique                                                   *
'*   Módulos:            Orçamentário                                                    *
'*   Data:               16/01/2006                                                      *
'*   Ficha:              orc1051                                                         *
'*   Objetivo:           Concantenar o Group By da Query da Function                     *
'*   strQueryRelatorio                                                                   *
'*****************************************************************************************

Dim sSQL As String

    sSQL = ""
    
    sSQL = sSQL & "GROUP BY "
    sSQL = sSQL & "EP.intNumero, EP.PKID, "
    sSQL = sSQL & "EP.strModalidade, EP.strContrato, "
    sSQL = sSQL & "RC.intPedidoEmpenho, "
    sSQL = sSQL & "LO.strDescricao, "
    sSQL = sSQL & "AC.strobjetoautorizacao, "
    sSQL = sSQL & "EP.strSolicitacao, EP.strCodigo, "
    sSQL = sSQL & "EP.intExercicio, EP.bitDigito, "
    sSQL = sSQL & "EP.dblValor, RD.intNumero, "
    sSQL = sSQL & "RD.intExercicioReserva, RD.strSolicitacao, "
    sSQL = sSQL & "RD.intExercicio, EP.Strcondpagto, "
    sSQL = sSQL & "EP.Strlocentrega, EP.Strprazoentrega, EP.dtmdata, "
    sSQL = sSQL & "PT.PKID, PT.strCodigo, PT.intCodigoReduzido, "
    sSQL = sSQL & "PJ.strcodigo, PT.intPrograma, PT.intProjetoAtividade,PJ.strDescricao, "
    sSQL = sSQL & "PT.intSubFuncao, PT.intFuncao, PT.dblValor, "
    sSQL = sSQL & "OG.strCodigo, OG.strDescricao, "
    sSQL = sSQL & "PJ.strDescricao, UO.strCodigo, "
    sSQL = sSQL & "UO.strDescricao, ED.strCodigoElementoDespesa, ED.strDescricao, "
    sSQL = sSQL & "CL.strCodigo, "
'    sSQL = sSQL & "SUP.strDecreto, SUP.dblValor2, "
'    sSQL = sSQL & "RED.strDecreto, RED.dblValor1, "
    sSQL = sSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "LG.STRDESCRICAO") & ", "
    sSQL = Replace(sSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.StrlogradouroD", "CT.Strlogradouroc"))
    sSQL = sSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intNumero") & ", "
    sSQL = Replace(sSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intNumeroD", "CT.intNumeroC"))
    sSQL = sSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.strComplemento") & ", "
    sSQL = Replace(sSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strComplementoD", "CT.strComplementoC"))
    sSQL = sSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "MP.strDescricao") & ", "
    sSQL = Replace(sSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strMunicipioD", "MPC.strDescricao"))
    sSQL = sSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "UF.strSigla") & ", "
    sSQL = Replace(sSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',UFD.strSigla", "UFC.strSigla"))
    sSQL = sSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "BR.strDescricao") & ", "
    sSQL = Replace(sSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.strbairroD", "CT.strbairroC"))
    sSQL = sSQL & gstrCASEWHEN(gstrISNULL("CT.intLogradouro", "-1"), "-1, StrTemp", "CT.intCEP") & ", "
    sSQL = Replace(sSQL, "StrTemp", gstrCASEWHEN(gstrISNULL("CT.Strlogradouroc", "'-1'"), "'-1',CT.intCEPD", "CT.intCEPC"))
    sSQL = sSQL & "FR.strCodigo, FR.strDescricao, CV.strCodigo, CV.strDescricao, "
    sSQL = sSQL & "CT.CDC, CT.strNome, "
    sSQL = sSQL & "CT.strCNPJCPF, "
    sSQL = sSQL & "CT.Strlogradouroc, CT.intNumeroC, CT.strComplementoC, "
    sSQL = sSQL & "MPC.strDescricao, UFC.strSigla, "
    sSQL = sSQL & "CT.strBairroC, CT.intCEPC, SEP.PKID, "
    sSQL = sSQL & "SEP.intNumero, SEP.bytSituacao, "
    sSQL = sSQL & "SEP.bytTipo, SEP.dtmData, "
    sSQL = sSQL & "SEP.intEmpenhoAnulacao, SEP.dblValor, "
    sSQL = sSQL & "SEP.dblEmpenhadoAteData, SEP.dblSaldoAtual, "
    If blnSoEstorno Then
       sSQL = sSQL & "SEP.strHistorico, "
    Else
       sSQL = sSQL & gstrCASEWHEN("SEP.intNumero", "0, EP.strHistorico", "SEP.strHistorico") & ", "
    End If
    sSQL = sSQL & "CL.strDescricao, IT.intcodigoitem, "
    sSQL = sSQL & "IT.Pkid, IT.Strdescricaoitem, IT.Strmarca, "
    sSQL = sSQL & "IT.STRUNIDADE, IT.DBLQUANTIDADE, "
    sSQL = sSQL & "IT.Dblprecounitario, "
    'Inserida a Data de Venciemnto - Ficha Orc1136 - Fernando
    sSQL = sSQL & "Sep.DtmVencimento "
'    sSQL = sSQL & gstrCONVERT(CDT_VARCHAR, "IT.strDescricaoDetalhada")
    AdicionaGroupByQueryRelatorio = sSQL

End Function

Private Function PegaDadosEmpenho(strCampo As String, strWhere As String) As String
'*****************************************************************************************
'*   Programador:        Éder Henrique                                                   *
'*   Módulos:            Orçamentário                                                    *
'*   Data:               16/01/2006                                                      *
'*   Ficha:              orc1145                                                         *
'*   Objetivo:           Concantenar o Group By da Query da Function                     *
'*   strQueryRelatorio                                                                   *
'*****************************************************************************************
    Dim adoResultado  As ADODB.Recordset
    Dim strSQL        As String
    
    
    strSQL = ""
    strSQL = strSQL & "SELECT " & strCampo
    strSQL = strSQL & "FROM " & gstrEmpenho & " "
    strSQL = strSQL & "WHERE" & strWhere
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            PegaDadosEmpenho = adoResultado.Fields(strCampo)
        End If
    End If


End Function

Private Function VerificaSaldoDotacao() As Boolean
Dim adoResultado  As ADODB.Recordset
Dim strSQL        As String

    'Alterado na pendencia orc1572
    VerificaSaldoDotacao = True
    strSQL = gstrStoredProcedure("sp_ProgTrabalhoParaEmpnho", Str(cboCodigoReduzido.ItemData(cboCodigoReduzido.ListIndex)), True, 50000)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If CDbl(IIf((Val(txtdblValor.Text) = 0), 0, txtdblValor.Text)) > CDbl(adoResultado.Fields("dblSaldo")) Then
            'If gstrConvVrDoSql(adoResultado.Fields("dblSaldo"), 2) < gstrConvVrDoSql(Val(gstrConvVrParaSql(txtdblValor.Text))) Then
                VerificaSaldoDotacao = False
                Exit Function
            End If
        End If
    End If
End Function

Private Function gdblSaldoDotacaoAtual(lngCodigoReduzido As Long) As Double
Dim adoResultado  As ADODB.Recordset
Dim strSQL        As String
    
    gdblSaldoDotacaoAtual = 0
    strSQL = gstrStoredProcedure("sp_ProgTrabalhoParaEmpnho", CStr(lngCodigoReduzido), True, 50000)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            gdblSaldoDotacaoAtual = CDbl(adoResultado.Fields("dblSaldo"))
        End If
    End If
End Function

Private Function CalculaSubTotalItem(Optional strOperacao As String)
    Dim dblValor As Double
    Dim dblQuantidade As Double
    Dim dblSubTotal As Double
    
    dblValor = gstrConvVrDoSql(txt_dblValorEstimado.Text, 5)
    dblQuantidade = gstrConvVrDoSql(txt_dblQuantidade.Text, 5)
    If txt_SubTotalItem = "" Then txt_SubTotalItem = 0
    dblSubTotal = gstrConvVrDoSql(txt_SubTotalItem, 5)
    If blnAlterandoItem Then dblSubTotal = dblSubTotal - dblSaldoItem: blnAlterandoItem = False
    If UCase(strOperacao) <> UCase(gstrExcluirItem) Then dblSubTotal = dblSubTotal + (dblValor * dblQuantidade)
    txt_SubTotalItem = gstrConvVrDoSql(dblSubTotal, 5)
End Function

Private Function PreencherSubTotalItens(strQuery As String)
    Dim adoResultado  As ADODB.Recordset
    Dim dblSaldo As Double
    Set gobjBanco = New clsBanco
    dblSaldo = 0
    If gobjBanco.CriaADO(strQuery, 10, adoResultado) Then
        If adoResultado.EOF = False Then
            While Not adoResultado.EOF
            dblSaldo = dblSaldo + (gstrConvVrDoSql(adoResultado.Fields("dblquantidade"), 5) * gstrConvVrDoSql(adoResultado.Fields("dblprecounitario"), 5))
            adoResultado.MoveNext
            Wend
        End If
    End If
    txt_SubTotalItem.Text = gstrConvVrDoSql(dblSaldo, 5)
End Function
