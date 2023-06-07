VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmArrecadacaoReceita 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrecadação de Receita"
   ClientHeight    =   7410
   ClientLeft      =   1590
   ClientTop       =   1920
   ClientWidth     =   9600
   HelpContextID   =   6
   Icon            =   "ArrecadacaoReceita.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   9600
   Begin VB.TextBox txtPKId 
      Height          =   285
      Left            =   6480
      TabIndex        =   66
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DDados 
      Height          =   7305
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   12885
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Arrecadação de Receita"
      TabPicture(0)   =   "ArrecadacaoReceita.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrNumero"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbldtmData"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintPlanoContas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintFundo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCodEventoContabil"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmd_BancoArrecadador"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtintNumero"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtdtmData"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmd_Fundo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_HistoricoSubEmpenho"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tdb_Lista"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmd_Historico"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cbointContaContabil"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dbc_strContaContabil"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbointFundo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cbo_intHistorico"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "tab_3DArecadarAnular"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "fra_Convenio"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmd_Evento"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cbointEvento"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_codEvento"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_CodHistorico"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      Begin VB.TextBox txt_CodHistorico 
         Height          =   285
         Left            =   6330
         MaxLength       =   5
         TabIndex        =   24
         Top             =   1950
         Width           =   585
      End
      Begin VB.TextBox txt_codEvento 
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   720
         Width           =   765
      End
      Begin VB.ComboBox cbointEvento 
         Height          =   315
         Left            =   2340
         TabIndex        =   7
         Top             =   720
         Width           =   3465
      End
      Begin VB.CommandButton cmd_Evento 
         Height          =   300
         Left            =   5800
         Picture         =   "ArrecadacaoReceita.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "247"
         ToolTipText     =   "Ativa Cadastro de Tipos de Envento"
         Top             =   735
         Width           =   330
      End
      Begin VB.Frame fra_Convenio 
         Caption         =   " Convênio "
         Height          =   915
         Left            =   90
         TabIndex        =   9
         Top             =   990
         Width           =   6135
         Begin VB.TextBox txt_Arrecadado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2835
            MaxLength       =   15
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   570
            Width           =   1335
         End
         Begin VB.TextBox txt_Saldo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4725
            MaxLength       =   15
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   570
            Width           =   1335
         End
         Begin VB.TextBox txt_ValorConvenio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   525
            MaxLength       =   15
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   570
            Width           =   1335
         End
         Begin VB.CommandButton cmd_Convenio 
            Height          =   300
            Left            =   5730
            Picture         =   "ArrecadacaoReceita.frx":13E8
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Tag             =   "247"
            ToolTipText     =   "Ativa Cadastro de Convênio"
            Top             =   210
            Width           =   330
         End
         Begin VB.ComboBox cbointConvenio 
            Height          =   315
            Left            =   900
            TabIndex        =   11
            Top             =   210
            Width           =   4815
         End
         Begin VB.Label lbl_Descricao 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   270
            Width           =   720
         End
         Begin VB.Label lbl_Arrecadado 
            AutoSize        =   -1  'True
            Caption         =   "Arrecadado"
            Height          =   195
            Left            =   1980
            TabIndex        =   15
            Top             =   600
            Width           =   825
         End
         Begin VB.Label lbl_Saldo 
            AutoSize        =   -1  'True
            Caption         =   "Saldo"
            Height          =   195
            Left            =   4260
            TabIndex        =   17
            Top             =   600
            Width           =   405
         End
         Begin VB.Label lbl_ValorConvenio 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   360
         End
      End
      Begin TabDlg.SSTab tab_3DArecadarAnular 
         Height          =   2985
         Left            =   90
         TabIndex        =   31
         Top             =   2650
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   5265
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Arrecadar "
         TabPicture(0)   =   "ArrecadacaoReceita.frx":1772
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl_TotalOrcamentario"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl_TotalExtra"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl_TotalGeral"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbl_Cancelado"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "tab_3DPasta"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txt_TotalOrcamentario"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txt_TotalExtraOrcamentario"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txt_TotalGeral"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txt_TotalCancelado"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Cancelar "
         TabPicture(1)   =   "ArrecadacaoReceita.frx":178E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txt_TotalGeralCancelado"
         Tab(1).Control(1)=   "tab_3DPastaCacelar"
         Tab(1).Control(2)=   "lbl_TotalGeralCancelado"
         Tab(1).ControlCount=   3
         Begin VB.TextBox txt_TotalCancelado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5910
            MaxLength       =   15
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   74
            Top             =   2610
            Width           =   1395
         End
         Begin VB.TextBox txt_TotalGeralCancelado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   -67170
            TabIndex        =   64
            Top             =   2610
            Width           =   1395
         End
         Begin VB.TextBox txt_TotalGeral 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   7830
            MaxLength       =   15
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   69
            Top             =   2610
            Width           =   1395
         End
         Begin VB.TextBox txt_TotalExtraOrcamentario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3660
            MaxLength       =   15
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   68
            Top             =   2610
            Width           =   1395
         End
         Begin VB.TextBox txt_TotalOrcamentario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1035
            MaxLength       =   15
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   67
            Top             =   2610
            Width           =   1395
         End
         Begin TabDlg.SSTab tab_3DPasta 
            Height          =   2205
            Left            =   60
            TabIndex        =   32
            Top             =   360
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   3889
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Orçamentária "
            TabPicture(0)   =   "ArrecadacaoReceita.frx":17AA
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbl_Valor"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lbl_ReceitaOrcamentaria"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lvw_Orcamentaria"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "cbostrOrcamentaria"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "cmd_PrevisaoDaReceita"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "cbointOrcamentaria"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "txt_dblValorOrcamentario"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "cbo_intCodigoReduzido"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).ControlCount=   8
            TabCaption(1)   =   "Extra-orçamentária "
            TabPicture(1)   =   "ArrecadacaoReceita.frx":17C6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "cmd_Orc"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "txt_dblValorExtraOrcamentaria"
            Tab(1).Control(2)=   "cbostrExtraOrcamentaria"
            Tab(1).Control(3)=   "cbointExtraOrcamentario"
            Tab(1).Control(4)=   "lvw_ExtraOrcamentaria"
            Tab(1).Control(5)=   "lbl_ReceitaExtra"
            Tab(1).Control(6)=   "lbl_ValorExtraOrcamentario"
            Tab(1).ControlCount=   7
            Begin VB.ComboBox cbo_intCodigoReduzido 
               Height          =   315
               Left            =   750
               Sorted          =   -1  'True
               TabIndex        =   34
               Top             =   390
               Width           =   1005
            End
            Begin VB.CommandButton cmd_Orc 
               Height          =   300
               Left            =   -66300
               Picture         =   "ArrecadacaoReceita.frx":17E2
               Style           =   1  'Graphical
               TabIndex        =   44
               TabStop         =   0   'False
               Tag             =   "322"
               ToolTipText     =   "Clique para cadastar conta"
               Top             =   390
               Width           =   330
            End
            Begin VB.TextBox txt_dblValorOrcamentario 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   750
               MaxLength       =   15
               MultiLine       =   -1  'True
               TabIndex        =   39
               Top             =   750
               Width           =   1695
            End
            Begin VB.TextBox txt_dblValorExtraOrcamentaria 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   -74250
               MaxLength       =   15
               MultiLine       =   -1  'True
               TabIndex        =   46
               Top             =   750
               Width           =   1335
            End
            Begin VB.ComboBox cbointOrcamentaria 
               Height          =   315
               Left            =   1800
               Sorted          =   -1  'True
               TabIndex        =   35
               Top             =   390
               Width           =   1875
            End
            Begin VB.CommandButton cmd_PrevisaoDaReceita 
               Height          =   300
               Left            =   8700
               Picture         =   "ArrecadacaoReceita.frx":1B6C
               Style           =   1  'Graphical
               TabIndex        =   37
               TabStop         =   0   'False
               Tag             =   "284"
               ToolTipText     =   "Ativa Cadastro de Previsão de Receita"
               Top             =   390
               Width           =   330
            End
            Begin VB.ComboBox cbostrOrcamentaria 
               Height          =   315
               Left            =   3720
               TabIndex        =   36
               Top             =   390
               Width           =   4950
            End
            Begin VB.ComboBox cbostrExtraOrcamentaria 
               Height          =   315
               Left            =   -72720
               TabIndex        =   43
               Top             =   390
               Width           =   6435
            End
            Begin VB.ComboBox cbointExtraOrcamentario 
               Height          =   315
               Left            =   -74250
               Sorted          =   -1  'True
               TabIndex        =   42
               Top             =   390
               Width           =   1575
            End
            Begin MSComctlLib.ListView lvw_Orcamentaria 
               Height          =   1035
               Left            =   90
               TabIndex        =   40
               Top             =   1080
               Width           =   8955
               _ExtentX        =   15796
               _ExtentY        =   1826
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
                  Text            =   "Receita"
                  Object.Width           =   2470
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Descrição"
                  Object.Width           =   8996
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Valor"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Text            =   "Cancelado"
                  Object.Width           =   1676
               EndProperty
            End
            Begin MSComctlLib.ListView lvw_ExtraOrcamentaria 
               Height          =   1035
               Left            =   -74910
               TabIndex        =   47
               Top             =   1080
               Width           =   8955
               _ExtentX        =   15796
               _ExtentY        =   1826
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
                  Text            =   "Conta"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Descrição"
                  Object.Width           =   8996
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Valor"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Cancelado"
                  Object.Width           =   1676
               EndProperty
            End
            Begin VB.Label lbl_ReceitaExtra 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Conta"
               Height          =   195
               Left            =   -74850
               TabIndex        =   41
               Top             =   420
               Width           =   555
            End
            Begin VB.Label lbl_ReceitaOrcamentaria 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Receita"
               Height          =   195
               Left            =   150
               TabIndex        =   33
               Top             =   420
               Width           =   555
            End
            Begin VB.Label lbl_Valor 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
               Height          =   195
               Left            =   345
               TabIndex        =   38
               Top             =   780
               Width           =   360
            End
            Begin VB.Label lbl_ValorExtraOrcamentario 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
               Height          =   195
               Left            =   -74655
               TabIndex        =   45
               Top             =   780
               Width           =   360
            End
         End
         Begin TabDlg.SSTab tab_3DPastaCacelar 
            Height          =   2205
            Left            =   -74940
            TabIndex        =   48
            Top             =   360
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   3889
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Orçamentária"
            TabPicture(0)   =   "ArrecadacaoReceita.frx":1EF6
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbl_ReceitaCancelar"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lblValor"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lvw_CodigoOrcamentarioCancelar"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "cbointOrcamentariaCancelar"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "cbostrOrcamentariaCancelar"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "txt_dblValorCancelamentoOrcamentario"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "cmd_PrevisaoReceitaCancelar"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "cbo_intCodigoReduzidoCancelar"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).ControlCount=   8
            TabCaption(1)   =   "Extra-orçamentária"
            TabPicture(1)   =   "ArrecadacaoReceita.frx":1F12
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_ContaCancelar"
            Tab(1).Control(1)=   "lblValorCancelamentoExtra"
            Tab(1).Control(2)=   "lvw_ContaContabilCancelar"
            Tab(1).Control(3)=   "cbointExtraOrcamentarioCancelar"
            Tab(1).Control(4)=   "cbostrExtraOrcamentariaCancelar"
            Tab(1).Control(5)=   "txt_dblValorExtraCancelamentoOrcamentaria"
            Tab(1).ControlCount=   6
            Begin VB.ComboBox cbo_intCodigoReduzidoCancelar 
               Height          =   315
               Left            =   750
               Sorted          =   -1  'True
               TabIndex        =   50
               Top             =   390
               Width           =   1005
            End
            Begin VB.CommandButton cmd_PrevisaoReceitaCancelar 
               Height          =   300
               Left            =   8685
               Picture         =   "ArrecadacaoReceita.frx":1F2E
               Style           =   1  'Graphical
               TabIndex        =   53
               TabStop         =   0   'False
               Tag             =   "284"
               ToolTipText     =   "Clique para cadastar previsão de receita"
               Top             =   390
               Width           =   330
            End
            Begin VB.TextBox txt_dblValorExtraCancelamentoOrcamentaria 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   -74250
               TabIndex        =   61
               Top             =   750
               Width           =   1335
            End
            Begin VB.TextBox txt_dblValorCancelamentoOrcamentario 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   750
               TabIndex        =   55
               Top             =   750
               Width           =   1695
            End
            Begin VB.ComboBox cbostrOrcamentariaCancelar 
               Height          =   315
               Left            =   3705
               TabIndex        =   52
               Top             =   390
               Width           =   4950
            End
            Begin VB.ComboBox cbointOrcamentariaCancelar 
               Height          =   315
               Left            =   1800
               Sorted          =   -1  'True
               TabIndex        =   51
               Top             =   390
               Width           =   1875
            End
            Begin VB.ComboBox cbostrExtraOrcamentariaCancelar 
               Height          =   315
               Left            =   -72720
               TabIndex        =   59
               Top             =   390
               Width           =   6765
            End
            Begin VB.ComboBox cbointExtraOrcamentarioCancelar 
               Height          =   315
               Left            =   -74250
               Sorted          =   -1  'True
               TabIndex        =   58
               Top             =   390
               Width           =   1575
            End
            Begin MSComctlLib.ListView lvw_ContaContabilCancelar 
               Height          =   1035
               Left            =   -74910
               TabIndex        =   62
               Top             =   1080
               Width           =   8955
               _ExtentX        =   15796
               _ExtentY        =   1826
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
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Receita"
                  Object.Width           =   2470
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Descrição"
                  Object.Width           =   10583
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Valor"
                  Object.Width           =   2117
               EndProperty
            End
            Begin MSComctlLib.ListView lvw_CodigoOrcamentarioCancelar 
               Height          =   1035
               Left            =   90
               TabIndex        =   56
               Top             =   1080
               Width           =   8955
               _ExtentX        =   15796
               _ExtentY        =   1826
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
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Conta"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Descrição"
                  Object.Width           =   10583
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   2
                  Text            =   "Valor"
                  Object.Width           =   2117
               EndProperty
            End
            Begin VB.Label lblValorCancelamentoExtra 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
               Height          =   195
               Left            =   -74715
               TabIndex        =   60
               Top             =   825
               Width           =   360
            End
            Begin VB.Label lblValor 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
               Height          =   195
               Left            =   345
               TabIndex        =   54
               Top             =   780
               Width           =   360
            End
            Begin VB.Label lbl_ReceitaCancelar 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Receita"
               Height          =   195
               Left            =   150
               TabIndex        =   49
               Top             =   420
               Width           =   555
            End
            Begin VB.Label lbl_ContaCancelar 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Conta"
               Height          =   195
               Left            =   -74850
               TabIndex        =   57
               Top             =   450
               Width           =   555
            End
         End
         Begin VB.Label lbl_Cancelado 
            AutoSize        =   -1  'True
            Caption         =   "Cancelado"
            Height          =   195
            Left            =   5100
            TabIndex        =   73
            Top             =   2655
            Width           =   765
         End
         Begin VB.Label lbl_TotalGeralCancelado 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Left            =   -67590
            TabIndex        =   63
            Top             =   2640
            Width           =   360
         End
         Begin VB.Label lbl_TotalGeral 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Left            =   7410
            TabIndex        =   72
            Top             =   2640
            Width           =   360
         End
         Begin VB.Label lbl_TotalExtra 
            AutoSize        =   -1  'True
            Caption         =   "Ex.Orçamentario"
            Height          =   195
            Left            =   2460
            TabIndex        =   71
            Top             =   2640
            Width           =   1170
         End
         Begin VB.Label lbl_TotalOrcamentario 
            AutoSize        =   -1  'True
            Caption         =   "Orçamentário"
            Height          =   195
            Left            =   60
            TabIndex        =   70
            Top             =   2640
            Width           =   945
         End
      End
      Begin VB.ComboBox cbo_intHistorico 
         Height          =   315
         Left            =   6930
         TabIndex        =   25
         Top             =   1930
         Width           =   2055
      End
      Begin VB.ComboBox cbointFundo 
         Height          =   315
         Left            =   750
         TabIndex        =   22
         Top             =   1930
         Width           =   5115
      End
      Begin VB.ComboBox dbc_strContaContabil 
         Height          =   315
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   2280
         Width           =   6765
      End
      Begin VB.ComboBox cbointContaContabil 
         Height          =   315
         ItemData        =   "ArrecadacaoReceita.frx":22B8
         Left            =   750
         List            =   "ArrecadacaoReceita.frx":22BA
         TabIndex        =   28
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Historico 
         Height          =   300
         Left            =   9045
         Picture         =   "ArrecadacaoReceita.frx":22BC
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "248"
         ToolTipText     =   "Ativa Cadastro de Histórico"
         Top             =   1930
         Width           =   330
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   1515
         Left            =   90
         TabIndex        =   65
         Top             =   5670
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   2672
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
         Columns(1).Caption=   "Nº da guia"
         Columns(1).DataField=   "intNumero"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   "dblTotal"
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Data"
         Columns(3).DataField=   "dtmData"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "intBanco"
         Columns(4).DataField=   "intPlanoConta"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Banco"
         Columns(5).DataField=   "strDescricao"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1693"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1614"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=514"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2461"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2381"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=79"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=9895"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=9816"
         Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.namedParent=42"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
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
         _StyleDefs(54)  =   "Named:id=33:Normal"
         _StyleDefs(55)  =   ":id=33,.parent=0"
         _StyleDefs(56)  =   "Named:id=34:Heading"
         _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(58)  =   ":id=34,.wraptext=-1"
         _StyleDefs(59)  =   "Named:id=35:Footing"
         _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=36:Selected"
         _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=37:Caption"
         _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(65)  =   "Named:id=38:HighlightRow"
         _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=39:EvenRow"
         _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(69)  =   "Named:id=40:OddRow"
         _StyleDefs(70)  =   ":id=40,.parent=33"
         _StyleDefs(71)  =   "Named:id=41:RecordSelector"
         _StyleDefs(72)  =   ":id=41,.parent=34"
         _StyleDefs(73)  =   "Named:id=42:FilterBar"
         _StyleDefs(74)  =   ":id=42,.parent=33"
      End
      Begin VB.Frame fra_HistoricoSubEmpenho 
         Caption         =   " Histórico "
         Height          =   1515
         Left            =   6300
         TabIndex        =   19
         Top             =   390
         Width           =   3105
         Begin VB.TextBox txtstrHistorico 
            Height          =   1185
            Left            =   150
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   240
            Width           =   2865
         End
      End
      Begin VB.CommandButton cmd_Fundo 
         Height          =   300
         Left            =   5880
         Picture         =   "ArrecadacaoReceita.frx":2646
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "249"
         ToolTipText     =   "Ativa Cadastro de Fundo"
         Top             =   1930
         Width           =   330
      End
      Begin VB.TextBox txtdtmData 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3300
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   380
         Width           =   1005
      End
      Begin VB.TextBox txtintNumero 
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   380
         Width           =   1005
      End
      Begin VB.CommandButton cmd_BancoArrecadador 
         Height          =   300
         Left            =   9045
         Picture         =   "ArrecadacaoReceita.frx":29D0
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "322"
         ToolTipText     =   "Clique para cadastar conta"
         Top             =   2295
         Width           =   330
      End
      Begin VB.Label lblCodEventoContabil 
         AutoSize        =   -1  'True
         Caption         =   "Evento Contabil"
         Height          =   195
         Left            =   285
         TabIndex        =   5
         Top             =   775
         Width           =   1125
      End
      Begin VB.Label lblintFundo 
         AutoSize        =   -1  'True
         Caption         =   "Fundo"
         Height          =   195
         Left            =   225
         TabIndex        =   21
         Top             =   1980
         Width           =   450
      End
      Begin VB.Label lblintPlanoContas 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   210
         TabIndex        =   27
         Top             =   2350
         Width           =   465
      End
      Begin VB.Label lbldtmData 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   2880
         TabIndex        =   3
         Top             =   405
         Width           =   345
      End
      Begin VB.Label lblstrNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número da guia"
         Height          =   195
         Left            =   285
         TabIndex        =   1
         Top             =   405
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmArrecadacaoReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mobjAux                                    As Object
Dim mobjLista                                  As Object
Dim mblnSelecionou                             As Boolean
Dim mblnClickOk                                As Boolean
Dim mblnLeLancamento                           As Boolean
Dim mblnAlterandoOrcamentarioCancelar          As Boolean
Dim mblnAlterandoExtraOrcamentariaCancelar     As Boolean
Dim mblnAlterandoOrcamentario                  As Boolean
Dim mblnAlterandoExtra                         As Boolean
Dim strDataOriginal                            As String
Dim mblnCarregaFormConta                       As Boolean
Public strDataInicial                          As String
Public strDataFinal                            As String
Public blnAtivaFormImprime                     As Boolean
Dim dblTotalOrcamentario                       As Double
Dim dblTotalExOrcamentario                     As Double
Dim aryContas()                                As Integer    'Array para ser utilizado na funcao de gravacao dos eventos
Dim aryTpMov()                                 As Byte       'Array que contera o tipo das contas que serão gravadas
Dim aryValor()                                 As String     'Array que contera o valor para as contas
Dim strFiltro                                  As String


Private Sub PreencheForm()
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT intNumero, dtmData, intConvenio, intFundo, "
    strSQL = strSQL & "strHistorico, intContaContabil , intevento "
    strSQL = strSQL & "FROM " & gstrArrecadacaoReceita & " "
    strSQL = strSQL & "WHERE PKId = " & Val(txtPKId)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                AtribuiValorDoSql txtintNumero, !INTNUMERO
                AtribuiValorDoSql txtdtmData, !DTMDATA
                strDataOriginal = txtdtmData
                txt_ValorConvenio = ""
                txt_Arrecadado = ""
                txt_Saldo = ""
                LeDaTabelaParaObj gstrConvenio, cbointConvenio
                cbointConvenio.ListIndex = gintIndiceCBO(cbointConvenio, gstrVerificaCampoNulo(!intConvenio))
                LeDaTabelaParaObj gstrFundo, cbointFundo
                cbointFundo.ListIndex = gintIndiceCBO(cbointFundo, gstrVerificaCampoNulo(!intFundo))
                AtribuiValorDoSql txtstrHistorico, !STRHISTORICO
                LePlanoContaGeralBanco cbointContaContabil, dbc_strContaContabil, "DC"
                cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, gstrVerificaCampoNulo(!intContaContabil))
                cbointEvento.ListIndex = gintIndiceCBO(cbointEvento, gstrVerificaCampoNulo(!intEvento))
            End If
        End With
    End If
    LeDaTabelaParaObj "", lvw_Orcamentaria, strQueryList(0)
    
    LeDaTabelaParaObj "", lvw_CodigoOrcamentarioCancelar, strQueryList(1)
    LeDaTabelaParaObj "", lvw_ExtraOrcamentaria, strQueryListExtra(0)
    
    LeDaTabelaParaObj "", lvw_ContaContabilCancelar, strQueryListExtra(1)
    
    txt_TotalOrcamentario.Text = gstrConvVrDoSql(CalculaTotal(txtPKId, 0, 0))
    'txt_TotalGeral.Text = gstrConvVrDoSql(CalculaTotal(txtPKId, 0))
    txt_TotalGeralCancelado.Text = gstrConvVrDoSql(CalculaTotal(txtPKId, 1, 0))
    txt_TotalCancelado.Text = gstrConvVrDoSql(CalculaTotal(txtPKId, 1, 0))
    txt_TotalGeralCancelado.Text = gstrConvVrDoSql(CDbl(txt_TotalGeralCancelado) + CDbl(CalculaTotal(txtPKId, 1, 1)))
    txt_TotalCancelado.Text = gstrConvVrDoSql(CDbl(txt_TotalCancelado) + CDbl(CalculaTotal(txtPKId, 1, 1)))
    
    txt_TotalExtraOrcamentario.Text = gstrConvVrDoSql(CalculaTotal(txtPKId, 0, 1))
    txt_TotalGeral.Text = gstrConvVrDoSql(CDbl(txt_TotalOrcamentario) + _
    CDbl(txt_TotalExtraOrcamentario.Text) - _
    CDbl(txt_TotalCancelado))
    'HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar
    'HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCancelar
    'tab_3DArecadarAnular.TabEnabled(1) = True
End Sub

Private Function gblnDadosOK() As Boolean
    Dim dtmDtEncerramento As Date
    
    'NÚMERO DA GUIA
    If Len(Trim(txtintNumero.Text)) = 0 Then
        ExibeMensagem "O número da Guia deve ser informado."
        txtintNumero.SetFocus
        Exit Function
        'DATA
    ElseIf gblnDataValida(txtdtmData) = False Then
        ExibeMensagem "A data está incorreta."
        txtdtmData.SetFocus
        Exit Function
    End If
    'CODIGO DO EVENTO
    If cbointEvento.ListIndex = -1 Then
        VerificaControlesByEvento False, True
    Else
        VerificaControlesByEvento True, True
    End If
    'TOTAL GERAL / SALDO
    If Not mblnAlterandoOrcamentario And cbointConvenio.ListIndex <> -1 _
    And Val(gstrConvVrParaSql(txt_TotalGeral)) > Val(gstrConvVrParaSql(txt_Saldo)) Then
    ExibeMensagem "O total arrecadado não pode ser superior ao saldo do convênio."
    lvw_Orcamentaria.SetFocus
    Exit Function
    'FUNDO
    'BANCO
ElseIf dbc_strContaContabil.ListIndex = -1 Then
    ExibeMensagem "O banco arrecadador não foi informado."
    If dbc_strContaContabil.Enabled Then dbc_strContaContabil.SetFocus
    'cbointContaContabil.SetFocus
    Exit Function
    'RECEITA
ElseIf Val(gstrConvVrParaSql(txt_TotalGeral)) = 0 And Val(gstrConvVrParaSql(txt_TotalGeralCancelado)) = 0 Then
    ExibeMensagem "Não há receita para gravar."
    lvw_Orcamentaria.SetFocus
    Exit Function
ElseIf mblnAlterandoOrcamentario And Right(strDataOriginal, 7) <> Right(txtdtmData, 7) Then
    ExibeMensagem "O mês e ano deste movimento não podem ser alterados."
    txtdtmData.SetFocus
    Exit Function
ElseIf cbointConvenio.ListIndex <> -1 Then
    If Not VerificaDtIntervaloConvenio(txtdtmData) Then
        ExibeMensagem "A data de movimento não pertence ao intervalo de datas do convênio informado."
        txtdtmData.SetFocus
        Exit Function
    End If
End If

dtmDtEncerramento = VerificaDataEncerramento("EF", gintExercicio)

If dtmDtEncerramento = Empty Then
    Exit Function
Else
    If CDate(txtdtmData) <= dtmDtEncerramento Then
        ExibeMensagem "A data do lançamento deve ser maior que a data de último encerramento (" & dtmDtEncerramento & ")."
        If txtdtmData.Enabled Then txtdtmData.SetFocus
        Exit Function
    End If
End If


'Orc677
If Right(txtdtmData, 4) <> gintExercicio Then
    ExibeMensagem "A data não equivale a data do exercício corrente."
    If txtdtmData.Enabled Then txtdtmData.SetFocus
    Exit Function
End If


gblnDadosOK = True
End Function

Private Function strQueryCancelar(lvw_Lita As ListView, bytTipo As Byte)
    Dim strSQL  As String
    Dim intInd  As Integer
    With lvw_Lita
        For intInd = 1 To .ListItems.Count
            strSQL = strSQL & "UPDATE " & gstrContaArrecadacaoReceita & " SET "
            strSQL = strSQL & "bytCancelado = 1, "
            strSQL = strSQL & "dtmDataCancelamento = "
            strSQL = strSQL & gstrConvDtParaSql(txtdtmData) & " "
            strSQL = strSQL & "WHERE bytTipo = " & bytTipo & " "
            strSQL = strSQL & "AND intArrecadacao = " & Val(txtPKId) & " "
            strSQL = strSQL & "AND intConta = " & .ListItems(intInd).Tag & ";"
        Next
    End With
    strQueryCancelar = strSQL
End Function

Private Function strQueryConta(lvw_Lita As ListView, bytTipo As Byte, bytLancamento As Byte)
    Dim strSQL  As String
    Dim intInd  As Integer
    With lvw_Lita
        For intInd = 1 To .ListItems.Count
            strSQL = strSQL & "INSERT INTO " & gstrContaArrecadacaoReceita & " ("
            strSQL = strSQL & "intArrecadacao, intConta, dblValorOrcamentario, "
            strSQL = strSQL & "bytTipo, bytCancelado, dtmDtAtualizacao, lngCodUsr) "
            strSQL = strSQL & "(SELECT MAX(PKId), "
            strSQL = strSQL & .ListItems(intInd).Tag & " , "
            strSQL = strSQL & gstrConvVrParaSql(.ListItems(intInd).SubItems(2)) & ", "
            'strSQL = strSQL & bytTipo & ", 0, "
            strSQL = strSQL & bytLancamento & ", " & bytTipo & ", "
            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & " "
            strSQL = strSQL & "FROM " & gstrArrecadacaoReceita & "); "
        Next
    End With
    strQueryConta = strSQL
End Function

Private Sub GravaArrecadacao()
    
    '******************************************************************************************
    ' Data: 09/06/2003
    ' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
    ' Responsável: Everton Bianchini
    '------------------------------------------------------------------------------------------
    ' Data: 11/06/2003
    ' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL
    '            permitindo, assim, a execução de múltiplos comandos SQL de uma única vez.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    Dim strSQL          As String
    Dim strCodigo       As String
    Dim intIdx          As Integer
    Dim dblSomaValores  As Double
    Dim intVlrConvenio  As Integer
    Dim strOrigemMovimento As String
    Dim strImprimeAux   As String
    Dim intInd          As Integer
    If gblnDadosOK Then
        If gblnExclusaoGravacaoOk(IIf(mblnAlterandoOrcamentario, "A", "I"), "desta receita", False) Then
            
            If mblnAlterandoOrcamentario Then
                strSQL = strQueryAlteraReceita
            End If
            If Not mblnAlterandoOrcamentario Then
                
ProximoCodigo:
                
                If gblnExisteCodigo(2, gstrArrecadacaoReceita, "intNumero", txtintNumero, "intExercicio", "'" & Val(gintExercicio) & "'") Then
                    strCodigo = (gstrProximoCodigo(txtintNumero, gstrArrecadacaoReceita, "intNumero", gintCodSeguranca, "intExercicio", Val(gintExercicio), , True))
                    If MsgBox("O número da Guia informado já se encontra cadastrado. Deseja usar o número " & strCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                        txtintNumero.SetFocus
                        Exit Sub
                    Else
                        txtintNumero.Text = strCodigo
                        GoTo ProximoCodigo
                    End If
                End If
                strSQL = strQueryIncluiReceita
            End If
            
            intVlrConvenio = VerificaValorConvenio
            If cbointConvenio.Text <> "" Then
                If intVlrConvenio = 0 Then
                    ExibeMensagem "O valor do convênio não pode ser menor que o valor total da arrecadação."
                    Exit Sub
                ElseIf intVlrConvenio = 1 Then
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaBeginTrans
                    Set gobjBanco = New clsBanco
                    If Not gobjBanco.Execute("UPDATE " & gstrConvenio & " SET dtmDataTerminio=" & gstrConvDtParaSql(txtdtmData) & " WHERE PKID=" & gstrItemData(cbointConvenio, True)) Then
                        ExibeMensagem "Houve problemas ao tentar atualizar o registro do convênio. A gravação será cancelada."
                        Set gobjBanco = New clsBanco
                        gobjBanco.ExecutaRollbackTrans
                        Exit Sub
                    Else
                        Set gobjBanco = New clsBanco
                        gobjBanco.ExecutaCommitTrans
                    End If
                End If
            End If
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSQL) Then
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaCommitTrans
                
                'Rotina criada para Gerar Movimentos do Evento Contábil
                
                If cbointEvento.ListIndex = -1 Then
                    ReDim aryContas(1 To lvw_ExtraOrcamentaria.ListItems.Count + lvw_ContaContabilCancelar.ListItems.Count + 1)
                    ReDim aryTpMov(1 To lvw_ExtraOrcamentaria.ListItems.Count + lvw_ContaContabilCancelar.ListItems.Count + 1)
                    ReDim aryValor(1 To lvw_ExtraOrcamentaria.ListItems.Count + lvw_ContaContabilCancelar.ListItems.Count + 1)
                    'Grava somente os valores para cada conta extra e não pelas contas do evento
                    If lvw_ExtraOrcamentaria.ListItems.Count > 0 Then
                        With lvw_ExtraOrcamentaria
                            For intIdx = 1 To .ListItems.Count
                                .ListItems(intIdx).Selected = True
                                aryContas(intIdx) = Val(.SelectedItem.Tag)
                                aryTpMov(intIdx) = 0 'Crédito
                                aryValor(intIdx) = Replace(Str(CDbl(.ListItems(intIdx).SubItems(2))), ".", ",")
                                dblSomaValores = Val(gstrConvVrParaSql(dblSomaValores)) + Val(gstrConvVrParaSql(.ListItems(intIdx).SubItems(2)))
                                
                            Next
                        End With
                    End If
                    'Grava somente os valores de cancelamento para cada conta extra e não pelas contas do evento
                    If lvw_ContaContabilCancelar.ListItems.Count > 0 Then
                        With lvw_ContaContabilCancelar
                            For intIdx = 1 To .ListItems.Count
                                .ListItems(intIdx).Selected = True
                                aryContas(intIdx + lvw_ExtraOrcamentaria.ListItems.Count) = Val(.SelectedItem.Tag)
                                aryTpMov(intIdx + lvw_ExtraOrcamentaria.ListItems.Count) = 0 'Crédito
                                aryValor(intIdx + lvw_ExtraOrcamentaria.ListItems.Count) = Replace(Str(CDbl(.ListItems(intIdx).SubItems(2)) * (-1)), ".", ",")
                                dblSomaValores = Val(gstrConvVrParaSql(dblSomaValores)) + (Val(gstrConvVrParaSql(.ListItems(intIdx).SubItems(2))) * (-1))
                            Next
                        End With
                    End If
                    aryContas(lvw_ExtraOrcamentaria.ListItems.Count + lvw_ContaContabilCancelar.ListItems.Count + 1) = gstrItemData(cbointContaContabil)
                    aryTpMov(lvw_ExtraOrcamentaria.ListItems.Count + lvw_ContaContabilCancelar.ListItems.Count + 1) = 1 'Débito
                    aryValor(lvw_ExtraOrcamentaria.ListItems.Count + lvw_ContaContabilCancelar.ListItems.Count + 1) = Replace(Str(dblSomaValores), ".", ",")
                    If Not mblnAlterandoOrcamentario Then
                        If Not GeraMovimentosByEvento(gstrItemData(cbointEvento), txtdtmData, "0", Trim(txtstrHistorico.Text), txtintNumero.Text, "6", aryContas, aryTpMov, True, aryValor) Then
                            ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                        End If
                    End If
                Else
                    
                    'Grava os valores orçamentários
                    ReDim aryContas(1)
                    ReDim aryTpMov(1)
                    
                    If Not mblnAlterandoOrcamentario Then
                        If lvw_Orcamentaria.ListItems.Count > 0 Then
                            aryContas(1) = gstrItemData(cbointContaContabil)
                            aryTpMov(1) = 1 'Débito
                            If Not GeraMovimentosByEvento(gstrItemData(cbointEvento), txtdtmData, Str(CDbl(IIf(Len(Trim(txt_TotalOrcamentario)) = 0, "0,00", txt_TotalOrcamentario))), Trim(txtstrHistorico.Text), txtintNumero.Text, "6", aryContas, aryTpMov) Then
                                ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                            End If
                        End If
                        'Grava os valores de cancelamentos orçamentários
                        If lvw_CodigoOrcamentarioCancelar.ListItems.Count > 0 Then
                            aryContas(1) = gstrItemData(cbointContaContabil)
                            aryTpMov(1) = 1 'Débito
                            If Not GeraMovimentosByEvento(gstrItemData(cbointEvento), txtdtmData, Str(CDbl(IIf(Len(Trim(txt_TotalGeralCancelado)) = 0, "0,00", txt_TotalGeralCancelado)) * (-1)), Trim(txtstrHistorico.Text), txtintNumero.Text, "6", aryContas, aryTpMov) Then
                                ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                            End If
                        End If
                    End If
                End If
                
                strFiltro = strFiltro & txtintNumero & ","
                
                LeDaTabelaParaObj gstrArrecadacaoReceita, tdb_Lista, strQuery(strFiltro)
                
                'Impressão da Autenticação de Movimentação
                If Not mblnAlterandoOrcamentario Then
                    If frmPagamento.blnFitaAutenticadoraOK Then
                        'Inicio da Impressão
                        If Trim(cbointContaContabil.Text) <> "" Then
                            strOrigemMovimento = "BCO"
                        End If
                        'Arrecadar
                        If tab_3DArecadarAnular.TabEnabled(0) = True Then
                            If tab_3DPasta.TabEnabled(0) = True Then
                                'Orçamentário
                                If Trim(strOrigemMovimento) = "" Then strOrigemMovimento = "ROR"
                                With Me.lvw_Orcamentaria
                                    strImprimeAux = strRegistroAutenticacao(txtintNumero.Text)
                                    For intInd = 1 To .ListItems.Count
                                        frmPagamento.ImprimeFitaAutenticadora strImprimeAux & Right(String(13 - Len(gstrConvVrDoSql(.ListItems(intInd).SubItems(2))), "0") & gstrConvVrDoSql(.ListItems(intInd).SubItems(2)), 13) & " " & strOrigemMovimento
                                    Next
                                End With
                            ElseIf tab_3DPasta.TabEnabled(1) = True Then
                                'Extra
                                If Trim(strOrigemMovimento) = "" Then strOrigemMovimento = "REX"
                                With Me.lvw_ExtraOrcamentaria
                                    strImprimeAux = strRegistroAutenticacao(txtintNumero.Text)
                                    For intInd = 1 To .ListItems.Count
                                        frmPagamento.ImprimeFitaAutenticadora strImprimeAux & Right(String(13 - Len(gstrConvVrDoSql(.ListItems(intInd).SubItems(2))), "0") & gstrConvVrDoSql(.ListItems(intInd).SubItems(2)), 13) & " " & strOrigemMovimento
                                    Next
                                End With
                            End If
                        End If
                        'Cancelar
                        If tab_3DArecadarAnular.TabEnabled(1) = True Then
                            If tab_3DPastaCacelar.TabEnabled(0) = True Then
                                'Orçamentário
                                If Trim(strOrigemMovimento) = "" Then strOrigemMovimento = "ROR"
                                With Me.lvw_CodigoOrcamentarioCancelar
                                    strImprimeAux = strRegistroAutenticacao(txtintNumero.Text)
                                    For intInd = 1 To .ListItems.Count
                                        frmPagamento.ImprimeFitaAutenticadora strImprimeAux & Right(String(13 - Len(gstrConvVrDoSql(.ListItems(intInd).SubItems(2))), "0") & gstrConvVrDoSql(.ListItems(intInd).SubItems(2)), 13) & " " & strOrigemMovimento
                                    Next
                                End With
                            ElseIf tab_3DPastaCacelar.TabEnabled(1) = True Then
                                'Extra
                                If Trim(strOrigemMovimento) = "" Then strOrigemMovimento = "REX"
                                With Me.lvw_ContaContabilCancelar
                                    strImprimeAux = strRegistroAutenticacao(txtintNumero.Text)
                                    For intInd = 1 To .ListItems.Count
                                        frmPagamento.ImprimeFitaAutenticadora strImprimeAux & Right(String(13 - Len(gstrConvVrDoSql(.ListItems(intInd).SubItems(2))), "0") & gstrConvVrDoSql(.ListItems(intInd).SubItems(2)), 13) & " " & strOrigemMovimento
                                    Next
                                End With
                            End If
                        End If
                    End If
                End If
                
                LimpaDadosArrecadacao
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
            End If
        End If
    End If
End Sub

Private Function CalculaTotal(strId As String, Index As Integer, intTipo As Integer)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
    strSQL = "SELECT SUM(dblValorOrcamentario) AS TOTAL " & _
    "FROM " & gstrContaArrecadacaoReceita & " " & _
    "WHERE bytCancelado = " & Index & _
    " AND bytTipo = " & intTipo & _
    "AND intArrecadacao = " & strId
    
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF And Not IsNull(adoResultado!Total) Then
            CalculaTotal = adoResultado!Total
        Else
            CalculaTotal = 0
        End If
    End If
End Function

Private Sub cbo_intCodigoReduzido_Click()
    'cbointOrcamentaria.ListIndex = cbo_intCodigoReduzido.ListIndex
    cbointOrcamentaria.ListIndex = gintIndiceCBO(cbointOrcamentaria, _
    gstrItemData(cbo_intCodigoReduzido))
End Sub

Private Sub cbo_intCodigoReduzido_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_intCodigoReduzido_LostFocus()
    glIgualaContas cbo_intCodigoReduzido, cbointOrcamentaria
End Sub

Private Sub cbo_intCodigoReduzido_Validate(Cancel As Boolean)
    Dim intFor As Integer
    
    For intFor = 0 To cbo_intCodigoReduzido.ListCount - 1
        If cbo_intCodigoReduzido.Text = cbo_intCodigoReduzido.list(intFor) And Len(Trim(cbo_intCodigoReduzido.Text)) > 0 Then
            cbo_intCodigoReduzido.ListIndex = intFor
            Exit For
        End If
    Next
End Sub

Private Sub cbo_intCodigoReduzidoCancelar_Click()
    cbointOrcamentaria.ListIndex = gintIndiceCBO(cbointOrcamentaria, _
    gstrItemData(cbo_intCodigoReduzido))
    'cbointOrcamentariaCancelar.ListIndex = cbo_intCodigoReduzidoCancelar.ListIndex
End Sub

Private Sub cbo_intCodigoReduzidoCancelar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbo_intCodigoReduzidoCancelar_LostFocus()
    glIgualaContas cbo_intCodigoReduzidoCancelar, cbointOrcamentariaCancelar
End Sub

Private Sub cbo_intCodigoReduzidoCancelar_Validate(Cancel As Boolean)
    Dim intFor As Integer
    
    For intFor = 0 To cbo_intCodigoReduzidoCancelar.ListCount - 1
        If cbo_intCodigoReduzidoCancelar.Text = cbo_intCodigoReduzidoCancelar.list(intFor) And Len(Trim(cbo_intCodigoReduzidoCancelar.Text)) > 0 Then
            cbo_intCodigoReduzidoCancelar.ListIndex = intFor
            Exit For
        End If
    Next
End Sub

Private Sub cbointEvento_Click()
    leCodigoEvento txt_codEvento, cbointEvento
    lvw_Orcamentaria.ListItems.Clear
    lvw_CodigoOrcamentarioCancelar.ListItems.Clear
    
    cbo_intCodigoReduzido.Clear
    cbointOrcamentaria.Clear
    cbostrOrcamentaria.Clear
    cbo_intCodigoReduzidoCancelar.Clear
    cbointOrcamentariaCancelar.Clear
    cbostrOrcamentariaCancelar.Clear
    
    LePrevisaoReceitaGeral cbointOrcamentaria, cbostrOrcamentaria, strQueryBuscaByEvento
    LePrevisaoReceitaGeral cbointOrcamentariaCancelar, cbostrOrcamentariaCancelar, strQueryBuscaByEvento
    LeDaTabelaParaObj "", cbo_intCodigoReduzido, strQueryCodigoReduzido
    LeDaTabelaParaObj "", cbo_intCodigoReduzidoCancelar, strQueryCodigoReduzido
    
    If cbointOrcamentaria.ListCount = 1 Then
        cbointOrcamentaria.ListIndex = 0
    End If
    
    If cbointOrcamentariaCancelar.ListCount = 1 Then
        cbointOrcamentariaCancelar.ListIndex = 0
        cbo_intCodigoReduzidoCancelar.ListIndex = gintIndiceCBO(cbo_intCodigoReduzido, _
        gstrItemData(cbointOrcamentaria))
        
    End If
    
    txt_TotalOrcamentario.Text = gstrConvVrDoSql("0")
    txt_TotalExtraOrcamentario.Text = gstrConvVrDoSql("0")
    txt_TotalCancelado.Text = gstrConvVrDoSql("0")
    txt_TotalGeral.Text = gstrConvVrDoSql("0")
    txt_TotalGeralCancelado.Text = gstrConvVrDoSql("0")
    
End Sub

Private Sub cbointEvento_GotFocus()
    If cbointEvento.Text = "" Then txt_codEvento.Text = ""
End Sub

Private Sub cbointEvento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointEvento_LostFocus()
    If cbointEvento.ListIndex = -1 Then
        txt_codEvento.Text = ""
        VerificaControlesByEvento (False)
    Else
        VerificaControlesByEvento (True), True
    End If
End Sub

Private Sub cbointContaContabil_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointContaContabil_LostFocus()
    glIgualaContas cbointContaContabil, dbc_strContaContabil
End Sub

Private Sub cbointContaContabil_Validate(Cancel As Boolean)
    Dim intFor As Integer
    
    For intFor = 0 To cbointContaContabil.ListCount - 1
        If cbointContaContabil.Text = cbointContaContabil.list(intFor) Then
            cbointContaContabil.ListIndex = intFor
            Exit For
        End If
    Next
End Sub

Private Sub cbointConvenio_Click()
    LeSaldoConvenio gstrItemData(cbointConvenio), 1, txt_Saldo, , txt_ValorConvenio, txt_Arrecadado
End Sub

Private Sub cbointConvenio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointConvenio_LostFocus()
    If Len(Trim(cbointConvenio)) = 0 Then
        txt_ValorConvenio = ""
        txt_Arrecadado = ""
        txt_Saldo = ""
    End If
    
End Sub

Private Sub cbointConvenio_Validate(Cancel As Boolean)
    If cbointConvenio.ListIndex = -1 Then
        cmd_Fundo.Enabled = True
        TrocaCorObjeto cbointFundo, False
        cbointContaContabil.ListIndex = -1
        dbc_strContaContabil.ListIndex = -1
        TrocaCorObjeto cbointContaContabil, False
        TrocaCorObjeto dbc_strContaContabil, False
        cmd_BancoArrecadador.Enabled = True
    Else
        cmd_Fundo.Enabled = False
        TrocaCorObjeto cbointFundo, True
        VerificaBancoConvenio
    End If
    
End Sub

Private Sub cbointExtraOrcamentario_GotFocus()
    'AtivaPastaDeObjeto tab_3DArecadarAnular, 0, tab_3DPasta, 1
End Sub

Private Sub cbointExtraOrcamentario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointExtraOrcamentarioCancelar_GotFocus()
    'AtivaPastaDeObjeto tab_3DArecadarAnular, 1, tab_3DPasta, 1
End Sub

Private Sub cbointExtraOrcamentarioCancelar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointFundo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointFundo_LostFocus()
    
    If cbointFundo.ListIndex = -1 Then
        cmd_Convenio.Enabled = True
        TrocaCorObjeto cbointConvenio, False
        cbointContaContabil.ListIndex = -1
        dbc_strContaContabil.ListIndex = -1
        TrocaCorObjeto cbointContaContabil, False
        TrocaCorObjeto dbc_strContaContabil, False
        cmd_BancoArrecadador.Enabled = True
    Else
        cmd_Convenio.Enabled = False
        TrocaCorObjeto cbointConvenio, True
        VerificaBancoFundo
    End If
    
End Sub

Private Sub cbointOrcamentaria_GotFocus()
    AtivaPastaDeObjeto tab_3DArecadarAnular, 0, tab_3DPasta, 0
End Sub

Private Sub cbointOrcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointOrcamentariaCancelar_Click()
    glIgualaContas cbointOrcamentariaCancelar, _
    cbostrOrcamentariaCancelar, _
    lvw_CodigoOrcamentarioCancelar, _
    mblnAlterandoOrcamentarioCancelar
    
    'If cbointOrcamentariaCancelar.ListIndex > -1 Then
    '    cbo_intCodigoReduzidoCancelar.ListIndex = cbointOrcamentariaCancelar.ListIndex
    'End If
End Sub

Private Sub cbointExtraOrcamentarioCancelar_Click()
    glIgualaContas cbointExtraOrcamentarioCancelar, _
    cbostrExtraOrcamentariaCancelar, _
    lvw_ContaContabilCancelar, _
    mblnAlterandoExtraOrcamentariaCancelar
End Sub

Private Sub cbointOrcamentariaCancelar_GotFocus()
    AtivaPastaDeObjeto tab_3DArecadarAnular, 1, tab_3DPasta, 0
End Sub

Private Sub cbointOrcamentariaCancelar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbostrExtraOrcamentaria_GotFocus()
    'AtivaPastaDeObjeto tab_3DArecadarAnular, 0, tab_3DPasta, 1
End Sub

Private Sub cbostrExtraOrcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbostrExtraOrcamentariaCancelar_GotFocus()
    'AtivaPastaDeObjeto tab_3DArecadarAnular, 1, tab_3DPasta, 1
End Sub

Private Sub cbostrExtraOrcamentariaCancelar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbostrOrcamentaria_GotFocus()
    AtivaPastaDeObjeto tab_3DArecadarAnular, 0, tab_3DPasta, 0
End Sub

Private Sub cbostrOrcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbostrOrcamentariaCancelar_Click()
    glIgualaContas cbostrOrcamentariaCancelar, _
    cbointOrcamentariaCancelar, _
    lvw_CodigoOrcamentarioCancelar, _
    mblnAlterandoOrcamentarioCancelar
End Sub

Private Sub cbostrExtraOrcamentariaCancelar_Click()
    glIgualaContas cbostrExtraOrcamentariaCancelar, _
    cbointExtraOrcamentarioCancelar, _
    lvw_ContaContabilCancelar, _
    mblnAlterandoExtraOrcamentariaCancelar
End Sub

Private Sub cbointOrcamentaria_Click()
    
    glIgualaContas cbointOrcamentaria, _
    cbostrOrcamentaria, _
    lvw_Orcamentaria, _
    mblnAlterandoOrcamentario
    
    '    If cbo_intCodigoReduzido.ListIndex > -1 Then
    'cbo_intCodigoReduzido.ListIndex = cbointOrcamentaria.ListIndex
    '    End If
    cbostrOrcamentaria.ListIndex = gintIndiceCBO(cbostrOrcamentaria, _
    gstrItemData(cbointOrcamentaria))
    cbo_intCodigoReduzido.ListIndex = gintIndiceCBO(cbo_intCodigoReduzido, _
    gstrItemData(cbointOrcamentaria))
End Sub

Private Sub cbostrOrcamentaria_Click()
    '    glIgualaContas cbostrOrcamentaria, _
    '                   cbointOrcamentaria, _
    '                   lvw_Orcamentaria, _
    '                   mblnAlterandoOrcamentario
    cbointOrcamentaria.ListIndex = gintIndiceCBO(cbointOrcamentaria, _
    gstrItemData(cbostrOrcamentaria))
End Sub

Private Sub cbointExtraOrcamentario_Click()
    glIgualaContas cbointExtraOrcamentario, _
    cbostrExtraOrcamentaria, _
    lvw_ExtraOrcamentaria, _
    mblnAlterandoExtra
End Sub

Private Sub cbostrExtraOrcamentaria_Click()
    glIgualaContas cbostrExtraOrcamentaria, _
    cbointExtraOrcamentario, _
    lvw_ExtraOrcamentaria, _
    mblnAlterandoExtra
End Sub

Private Sub cbointContaContabil_Click()
    glIgualaContas cbointContaContabil, dbc_strContaContabil
End Sub

Private Sub cbostrOrcamentariaCancelar_GotFocus()
    AtivaPastaDeObjeto tab_3DArecadarAnular, 1, tab_3DPasta, 0
End Sub

Private Sub cbostrOrcamentariaCancelar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub


Private Sub cmd_Evento_Click()
    CarregaForm frmCadEvento, cbointEvento, strQueryAplicarEvento
End Sub

Private Sub cmd_Orc_Click()
    CarregaForm frmCadPlanoConta, cbostrExtraOrcamentaria, strQueryExtraOrc()
End Sub

Private Sub cmd_PrevisaoDaReceita_Click()
    LePrevisaoReceitaGeral cbointOrcamentaria, cbostrOrcamentaria
    CarregaForm frmConPrevisaoDaReceita, cbostrOrcamentaria
End Sub

Private Sub cmd_PrevisaoReceitaCancelar_Click()
    LePrevisaoReceitaGeral cbointOrcamentariaCancelar, cbostrOrcamentariaCancelar
    CarregaForm frmConPrevisaoDaReceita, cbostrOrcamentariaCancelar
End Sub


Private Sub dbc_strContaContabil_Click()
    Dim tempIndice As Integer
    
    On Error GoTo Problema
    
    tempIndice = dbc_strContaContabil.ListIndex
    cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, _
    gstrItemData(dbc_strContaContabil))
    
    If cbointContaContabil.ListIndex = -1 Then
        LePlanoContaGeralBanco cbointContaContabil, dbc_strContaContabil, "DC"
        dbc_strContaContabil.ListIndex = tempIndice
        cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, _
        gstrItemData(dbc_strContaContabil))
    End If
    
    If dbc_strContaContabil.ListIndex = -1 Then cbointContaContabil.ListIndex = -1
    
    Exit Sub
    
Problema:
    If Err.Number = 380 Then
        Exit Sub
    End If
    
End Sub

Private Sub cbo_intHistorico_Click()
    'txtstrHistorico = cbo_intHistorico.Text
    Dim adoResultado As ADODB.Recordset
    Set gobjBanco = New clsBanco
    gobjBanco.CriaADO "SELECT h.strcodigo FROM tblhistorico H WHERE H.STRdescricao = '" & Me.cbo_intHistorico.Text & "'", 10, adoResultado
    With adoResultado
        If Not .EOF Then
            Me.txt_CodHistorico.Text = gstrENulo(!strCodigo)
            Me.txtstrHistorico.Text = Me.cbo_intHistorico.Text
        End If
    End With
End Sub

Private Sub cbo_intHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Function strQuery(Optional ByVal strFiltro As String) As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT AR.PKId, AR.intNumero, AR.dtmData, PC.strDescricao, "
    strSQL = strSQL & " ((SELECT " & gstrISNULL("SUM(CA.dblValorOrcamentario)", "0")
    strSQL = strSQL & " FROM " & gstrContaArrecadacaoReceita & " CA WHERE CA.intArrecadacao = AR.PKId AND "
    strSQL = strSQL & "CA.bytCancelado = 0 ) - "
    strSQL = strSQL & " (SELECT " & gstrISNULL("SUM(CA.dblValorOrcamentario)", "0")
    strSQL = strSQL & " FROM " & gstrContaArrecadacaoReceita & " CA WHERE CA.intArrecadacao = AR.PKId AND "
    strSQL = strSQL & "CA.bytCancelado = 1 )) AS dblTotal "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA, "
    strSQL = strSQL & gstrArrecadacaoReceita & " AR, "
    strSQL = strSQL & gstrPlanoConta & " PC "
    strSQL = strSQL & "WHERE AR.intContaContabil = PC.PKId "
    strSQL = strSQL & "AND CA.intArrecadacao = AR.PKId "
    strSQL = strSQL & "AND AR.intExercicio = " & gintExercicio & " "
    
    If strFiltro <> "" Then
        strSQL = strSQL & " AND AR.intNumero IN (" & Mid(strFiltro, 1, Len(strFiltro) - 1) & ")"
    End If
    If strFiltro = "" And Me.cbointEvento.Text <> "" Then
        strSQL = strSQL & " AND AR.intEvento = '" & Me.cbointEvento.ItemData(Me.cbointEvento.ListIndex) & "'"
    End If
    strSQL = strSQL & "GROUP BY AR.PKId, AR.intNumero, AR.dtmData, PC.strDescricao "
    strSQL = strSQL & "ORDER BY AR.intNumero"
    strQuery = strSQL
End Function

Private Function strQueryOrc(campo As String) As String
    campo = "SELECT CO.PKId, " & _
    campo & " " & _
    "FROM " & gstrCodigoOrcamentario & " CO , " & _
    gstrPrevisaoDaReceita & " PR " & _
    "WHERE CO.PKId = PR.intCodigoOrcamentario"
    
    strQueryOrc = campo
End Function

Private Function strQueryExtraOrc() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrPlanoConta & " "
    strSQL = strSQL & "WHERE blnExtraOrcamentaria = 1 "
    strSQL = strSQL & "AND ABS(blnAnalitica) = 1"
    strQueryExtraOrc = strSQL
End Function

Private Function strQueryList(Index As Integer) As String
    
    '******************************************************************************************
    ' Data: 09/06/2003
    ' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    Dim strSQL          As String
    Dim strAuxTabela    As String
    Dim strAuxCampo     As String
    '    If Index = 0 Then
    '        strAuxTabela = gstrCodigoOrcamentario
    '        strAuxCampo = "strCodigoOrcamentario"
    '    Else
    '        strAuxTabela = gstrPlanoConta
    '        strAuxCampo = "strContaContabil"
    '    End If
    
    strAuxTabela = gstrCodigoOrcamentario
    strAuxCampo = "strCodigoOrcamentario"
    
    strSQL = ""
    strSQL = strSQL & "SELECT CO.PKId, CO." & strAuxCampo & ", "
    strSQL = strSQL & "CO.strDescricao, CAR.dblValorOrcamentario  "
    '    strSql = strSql & "CASE bytCancelado "
    '    strSql = strSql & "WHEN 0 THEN 'Não' "
    '    strSql = strSql & "WHEN 1 THEN 'Sim' "
    '    strSql = strSql & "END AS strCancelado "
    'strSQL = strSQL & gstrCASEWHEN("bytCancelado", _
    '    "0, 'Não', 1, 'Sim'") & " AS strCancelado "
    
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CAR, "
    strSQL = strSQL & strAuxTabela & " CO "
    strSQL = strSQL & "WHERE CAR.intConta = CO.PKId "
    'strSQL = strSQL & "AND CAR.bytTipo = " & Index & " "
    strSQL = strSQL & "AND CAR.bytCancelado = " & Index & " "
    strSQL = strSQL & "AND CAR.intArrecadacao = " & Val(txtPKId)
    strQueryList = strSQL
End Function
Private Function strQueryListExtra(Index As Integer) As String
    
    '******************************************************************************************
    ' Data: 09/06/2003
    ' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    Dim strSQL          As String
    Dim strAuxTabela    As String
    Dim strAuxCampo     As String
    '    If Index = 0 Then
    '        strAuxTabela = gstrCodigoOrcamentario
    '        strAuxCampo = "strCodigoOrcamentario"
    '    Else
    '        strAuxTabela = gstrPlanoConta
    '        strAuxCampo = "strContaContabil"
    '    End If
    
    strAuxTabela = gstrPlanoConta
    strAuxCampo = "strContaContabil"
    
    strSQL = ""
    strSQL = strSQL & "SELECT PL.PKId, PL." & strAuxCampo & ", "
    strSQL = strSQL & "PL.strDescricao, CAR.dblValorOrcamentario  "
    '    strSql = strSql & "CASE bytCancelado "
    '    strSql = strSql & "WHEN 0 THEN 'Não' "
    '    strSql = strSql & "WHEN 1 THEN 'Sim' "
    '    strSql = strSql & "END AS strCancelado "
    'strSQL = strSQL & gstrCASEWHEN("bytCancelado", _
    '    "0, 'Não', 1, 'Sim'") & " AS strCancelado "
    
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CAR, "
    strSQL = strSQL & strAuxTabela & " PL "
    strSQL = strSQL & "WHERE CAR.intConta = PL.PKId "
    'strSQL = strSQL & "AND CAR.bytTipo = " & Index & " "
    strSQL = strSQL & "AND CAR.bytCancelado = " & Index & " "
    strSQL = strSQL & "AND CAR.intArrecadacao = " & Val(txtPKId)
    strQueryListExtra = strSQL
End Function

Private Sub dbc_strContaContabil_GotFocus()
    If mblnCarregaFormConta = True Then
        mblnCarregaFormConta = False
        If dbc_strContaContabil.ListIndex = -1 Then cbointContaContabil.ListIndex = -1
    End If
End Sub

Private Sub dbc_strContaContabil_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 285
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
    If mblnAlterandoOrcamentario Then
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    Else
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    End If
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImportarDados
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
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImportarDados
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
    VerificaListaAutomatica "", tdb_Lista, strQuery
    mblnAlterandoOrcamentario = False
    
    lvw_Orcamentaria.ColumnHeaders.Remove (4)
    lvw_Orcamentaria.ColumnHeaders(2).Width = Val(5999.8116)
    lvw_ExtraOrcamentaria.ColumnHeaders.Remove (4)
    lvw_ExtraOrcamentaria.ColumnHeaders(2).Width = Val(5999.8116)
    
    preencheCboevento
    
    If cbointEvento.ListCount = 1 Then
        cbointEvento.ListIndex = 0
    End If
    
    LePlanoContaGeralBanco cbointContaContabil, dbc_strContaContabil, "DC"
    
    txtdtmData = VerificaDataEncerramento("EF", gintExercicio) + 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub

Private Sub lvw_CodigoOrcamentarioCancelar_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub lvw_ExtraOrcamentaria_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '    With lvw_ExtraOrcamentaria
    '        cbostrExtraOrcamentaria.ListIndex = gintIndiceCBO(cbostrExtraOrcamentaria, _
    '                                                    .SelectedItem.Tag)
    '        txt_dblValorExtraOrcamentaria = .ListItems(.SelectedItem.Index).SubItems(2)
    '    End With
End Sub

Private Sub lvw_Orcamentaria_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '    With lvw_Orcamentaria
    '        cbostrOrcamentaria.ListIndex = gintIndiceCBO(cbostrOrcamentaria, _
    '                                                     .SelectedItem.Tag)
    '        txt_dblValorOrcamentario = .ListItems(.SelectedItem.Index).SubItems(2)
    '    End With
End Sub

Private Sub LeLancamento()
    Dim strSQL       As String
    'Le lançamento extra-orçamentário
    strSQL = ""
    strSQL = strSQL & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " PC, "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA "
    strSQL = strSQL & "WHERE CA.intConta = PC.PKId "
    strSQL = strSQL & "AND CA.bytTipo = 1 "
    strSQL = strSQL & "AND CA.bytCancelado = 0 "
    strSQL = strSQL & "AND CA.intArrecadacao = " & Val(txtPKId)
    LePlanoContaGeral cbointExtraOrcamentarioCancelar, cbostrExtraOrcamentariaCancelar, strSQL
    'Le lançamento orçamentário
    strSQL = ""
    strSQL = strSQL & "SELECT CO.PKId, CO.strCodigoOrcamentario, CO.strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrCodigoOrcamentario & " CO, "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " CA "
    strSQL = strSQL & "WHERE CA.intConta = CO.PKId "
    strSQL = strSQL & "AND CA.bytTipo = 0 "
    strSQL = strSQL & "AND CA.bytCancelado = 0 "
    strSQL = strSQL & "AND CA.intArrecadacao = " & Val(txtPKId)
    'LePrevisaoReceitaGeral cbointOrcamentariaCancelar, _
    '                       cbostrOrcamentariaCancelar, strSQL
    mblnLeLancamento = False
End Sub

Private Sub lvw_Orcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tab_3DArecadarAnular_Click(PreviousTab As Integer)
    If tab_3DArecadarAnular.Tab = 1 And mblnLeLancamento Then
        LeLancamento
    End If
End Sub

Private Sub tab_3DArecadarAnular_GotFocus()
    EnviaTeclaTab 13
End Sub

Private Sub tab_3DPasta_GotFocus()
    EnviaTeclaTab 13
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_FilterChange()
    gblnFilraCampos tdb_Lista
End Sub

Private Sub cmd_Convenio_Click()
    CarregaForm frmCadConvenio, cbointConvenio
End Sub

Private Sub cmd_Fundo_Click()
    CarregaForm frmCadFundo, cbointFundo
End Sub

Private Sub cmd_Historico_Click()
    CarregaForm frmCadHistorico, cbo_intHistorico
End Sub

Private Sub cmd_BancoArrecadador_Click()
    mblnCarregaFormConta = True
    CarregaForm frmCadPlanoConta, dbc_strContaContabil, strQueryAplicar
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            mblnAlterandoOrcamentario = True
            txtPKId.Text = .Columns("PKID").Value
            If cbointEvento.ListCount = 0 Then
                preencheCboevento
            End If
            PreencheForm
            If lvw_Orcamentaria.ListItems.Count > 0 Then
                tab_3DPasta.TabEnabled(0) = True
                tab_3DPasta.TabEnabled(1) = False
                tab_3DArecadarAnular.Tab = 0
                tab_3DArecadarAnular.TabEnabled(0) = True
                tab_3DArecadarAnular.TabEnabled(1) = False
                tab_3DPasta.Tab = 0
            ElseIf lvw_ExtraOrcamentaria.ListItems.Count > 0 Then
                tab_3DPasta.TabEnabled(0) = False
                tab_3DPasta.TabEnabled(1) = True
                tab_3DArecadarAnular.Tab = 0
                tab_3DArecadarAnular.TabEnabled(0) = True
                tab_3DArecadarAnular.TabEnabled(1) = False
                tab_3DPasta.Tab = 1
            ElseIf lvw_CodigoOrcamentarioCancelar.ListItems.Count > 0 Then
                tab_3DPastaCacelar.TabEnabled(0) = True
                tab_3DPastaCacelar.TabEnabled(1) = False
                tab_3DArecadarAnular.Tab = 1
                tab_3DArecadarAnular.TabEnabled(0) = False
                tab_3DArecadarAnular.TabEnabled(1) = True
                tab_3DPastaCacelar.Tab = 0
            End If
            
            HabilitaDesabilitaControlesArrecadacao
            mblnLeLancamento = True
            tab_3DArecadarAnular_Click 0
            gCorLinhaSelecionada tdb_Lista
            
            ' If lvw_Orcamentaria.ListItems.Count > 0 Then tab_3DPasta.TabEnabled(1) = False
            'If lvw_ExtraOrcamentaria.ListItems.Count > 0 Then tab_3DPasta.TabEnabled(0) = False
            
            TrocaCorObjeto txt_codEvento, True
            TrocaCorObjeto cbointEvento, True
            TrocaCorObjeto txtdtmData, True
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrNovo, gstrDeletar
            mblnSelecionou = True
        End If
    End With
End Sub

Private Sub Totaliza(lvw_Lista As ListView, txt_total As TextBox)
    Dim intInd          As Integer
    Dim dblTotal        As Double
    Dim dblOrcamentario As Double
    Dim dblExtra        As Double
    
    With lvw_Lista
        For intInd = 1 To .ListItems.Count
            dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(2)))
        Next
        txt_total = gstrConvVrDoSql(CDbl(dblTotal))
        dblOrcamentario = Val(gstrConvVrParaSql(txt_TotalOrcamentario))
        dblExtra = Val(gstrConvVrParaSql(txt_TotalExtraOrcamentario))
        txt_TotalGeral = gstrConvVrDoSql(dblOrcamentario + dblExtra)
        
    End With
End Sub

Sub VerificaTabParaExcluir()
    Select Case tab_3DArecadarAnular.Tab
    Case 0 'Arrecadação
        Select Case tab_3DPasta.Tab
        Case 0
            ExcluirItemDaLista lvw_Orcamentaria, cbointOrcamentaria, _
            mblnAlterandoOrcamentario, _
            txt_TotalOrcamentario, _
            txt_dblValorOrcamentario
        Case 1
            ExcluirItemDaLista lvw_ExtraOrcamentaria, _
            cbointExtraOrcamentario, _
            mblnAlterandoExtra, _
            txt_TotalExtraOrcamentario, _
            txt_dblValorExtraOrcamentaria
        End Select
    Case 1 'Cancelamento
        Select Case tab_3DPastaCacelar.Tab
        Case 0
            ExcluirItemDaLista lvw_CodigoOrcamentarioCancelar, _
            cbointOrcamentariaCancelar, _
            mblnAlterandoOrcamentarioCancelar, _
            txt_TotalGeralCancelado
            dblTotalOrcamentario = CDbl(gstrConvVrDoSql(txt_TotalGeralCancelado))
            txt_TotalGeralCancelado = gstrConvVrDoSql(dblTotalOrcamentario + dblTotalExOrcamentario)
            txt_TotalCancelado = txt_TotalGeralCancelado
            txt_TotalGeral = gstrConvVrDoSql(CDbl(txt_TotalGeral) - CDbl(txt_TotalCancelado))
        Case 1
            ExcluirItemDaLista lvw_ContaContabilCancelar, _
            cbointExtraOrcamentarioCancelar, _
            mblnAlterandoExtraOrcamentariaCancelar, _
            txt_TotalGeralCancelado
            dblTotalExOrcamentario = CDbl(gstrConvVrDoSql(txt_TotalGeralCancelado))
            txt_TotalGeralCancelado = gstrConvVrDoSql(dblTotalOrcamentario + dblTotalExOrcamentario)
            txt_TotalCancelado = txt_TotalGeralCancelado
            txt_TotalGeral = gstrConvVrDoSql(CDbl(txt_TotalGeral) - CDbl(txt_TotalCancelado))
            
        End Select
    End Select
End Sub

Private Sub VerificaTabParaIncluir()
    
    Select Case tab_3DArecadarAnular.Tab
    Case 0 'Arrecadação
        Select Case tab_3DPasta.Tab
        Case 0
            txt_dblValorOrcamentario = gstrConvVrDoSql(txt_dblValorOrcamentario)
            IncluiItemNaLista cbointOrcamentaria, _
            cbostrOrcamentaria, _
            txt_dblValorOrcamentario, lvw_Orcamentaria, _
            txt_TotalOrcamentario, mblnAlterandoOrcamentario
        Case 1
            txt_dblValorExtraOrcamentaria = gstrConvVrDoSql(txt_dblValorExtraOrcamentaria)
            IncluiItemNaLista cbointExtraOrcamentario, _
            cbostrExtraOrcamentaria, _
            txt_dblValorExtraOrcamentaria, lvw_ExtraOrcamentaria, _
            txt_TotalExtraOrcamentario, mblnAlterandoExtra
        End Select
    Case 1 'Cancelamento
        Select Case tab_3DPastaCacelar.Tab
            
        Case 0
            
            txt_dblValorCancelamentoOrcamentario = gstrConvVrDoSql(txt_dblValorCancelamentoOrcamentario)
            IncluiItemNaListaCancelar cbointOrcamentariaCancelar, _
            cbostrOrcamentariaCancelar, _
            txt_dblValorCancelamentoOrcamentario, _
            txt_TotalGeralCancelado, _
            lvw_CodigoOrcamentarioCancelar, _
            mblnAlterandoOrcamentarioCancelar
            'strValorLancamento(0),
            dblTotalOrcamentario = CDbl(gstrConvVrDoSql(txt_TotalGeralCancelado))
            txt_TotalGeralCancelado = gstrConvVrDoSql(dblTotalOrcamentario + dblTotalExOrcamentario)
            txt_TotalCancelado = txt_TotalGeralCancelado
            txt_TotalGeral = gstrConvVrDoSql(CDbl(txt_TotalGeral) - CDbl(txt_TotalCancelado))
        Case 1
            txt_dblValorExtraCancelamentoOrcamentaria = gstrConvVrDoSql(txt_dblValorExtraCancelamentoOrcamentaria)
            IncluiItemNaListaCancelar cbointExtraOrcamentarioCancelar, _
            cbostrExtraOrcamentariaCancelar, _
            txt_dblValorExtraCancelamentoOrcamentaria, _
            txt_TotalGeralCancelado, _
            lvw_ContaContabilCancelar, _
            mblnAlterandoExtraOrcamentariaCancelar
            'strValorLancamento(1)
            dblTotalExOrcamentario = CDbl(gstrConvVrDoSql(txt_TotalGeralCancelado))
            txt_TotalGeralCancelado = gstrConvVrDoSql(dblTotalOrcamentario + dblTotalExOrcamentario)
            txt_TotalCancelado = txt_TotalGeralCancelado
            txt_TotalGeral = gstrConvVrDoSql(CDbl(txt_TotalGeral) - CDbl(txt_TotalCancelado))
            
        End Select
    End Select
End Sub

Private Function strValorLancamento(bytTipo As Byte) As String
    Dim strConta        As String
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Set gobjBanco = New clsBanco
    If bytTipo = 0 Then
        strConta = gstrItemData(cbostrOrcamentariaCancelar)
    Else
        strConta = gstrItemData(cbostrExtraOrcamentariaCancelar)
    End If
    strSQL = ""
    strSQL = strSQL & "SELECT SUM(dblValorOrcamentario) AS Total "
    strSQL = strSQL & "FROM " & gstrContaArrecadacaoReceita & " "
    strSQL = strSQL & "WHERE bytTipo = " & bytTipo & " "
    strSQL = strSQL & "AND intConta = " & strConta & " "
    strSQL = strSQL & "AND intArrecadacao = " & Val(txtPKId)
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            strValorLancamento = gstrConvVrDoSql(adoResultado!Total)
        End If
    End If
End Function

Private Function blnDadosContaOk(cboConta As ComboBox, _
    txtValor As TextBox)
    If cboConta.ListIndex = -1 Then
        ExibeMensagem "A conta não foi informada corretamente."
        cboConta.SetFocus
    ElseIf Val(gstrConvVrParaSql(txtValor)) = 0 Then
        ExibeMensagem "O valor não foi informado corretamente."
        txtValor.SetFocus
    Else
        blnDadosContaOk = True
    End If
End Function

Sub ExcluirItemDaLista(lvw_Lista As ListView, _
    cboConta As ComboBox, _
    blnAlterando As Boolean, _
    Optional txtTotal As TextBox, _
    Optional txtValor As TextBox)
    With lvw_Lista
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
            blnAlterando = False
            cboConta.ListIndex = -1
            If txtTotal Is Nothing = False Then
                Totaliza lvw_Lista, txtTotal
                LimpaDados cboConta, cboConta, txtValor, blnAlterando
            End If
        End If
    End With
End Sub

Private Sub IncluiItemNaLista(cboConta As ComboBox, _
    cboDescricao As ComboBox, _
    txtValor As TextBox, _
    lvw_Lista As ListView, _
    txtTotal As TextBox, _
    blnAlterando As Boolean)
    If blnDadosContaOk(cboDescricao, txtValor) Then
        If blnAlterando Then
            lvw_Lista.SelectedItem.Text = cboConta.Text
            lvw_Lista.SelectedItem.SubItems(1) = cboDescricao.Text
            lvw_Lista.SelectedItem.SubItems(2) = txtValor
            lvw_Lista.Tag = gstrItemData(cboDescricao)
        Else
            Set mobjLista = lvw_Lista.ListItems.Add(, , cboConta.Text)
            mobjLista.SubItems(1) = cboDescricao.Text
            mobjLista.SubItems(2) = txtValor
            mobjLista.Tag = gstrItemData(cboDescricao)
        End If
        Totaliza lvw_Lista, txtTotal
        LimpaDados cboDescricao, cboConta, txtValor, blnAlterando
    End If
End Sub

Private Sub IncluiItemNaListaCancelar(cboConta As ComboBox, _
    cboDescricao As ComboBox, _
    txtValor As TextBox, _
    txtTotal As TextBox, _
    lvw_Lista As ListView, _
    blnAlterando As Boolean)
    
    If cboDescricao.ListIndex > -1 Then
        If blnAlterando Then
            lvw_Lista.SelectedItem.Text = cboConta.Text
            lvw_Lista.SelectedItem.SubItems(1) = cboDescricao.Text
            'lvw_Lista.SelectedItem.SubItems(2) = vntValor
            lvw_Lista.SelectedItem.SubItems(2) = txtValor
            lvw_Lista.Tag = gstrItemData(cboDescricao)
        Else
            Set mobjLista = lvw_Lista.ListItems.Add(, , cboConta.Text)
            mobjLista.SubItems(1) = cboDescricao.Text
            'mobjLista.SubItems(2) = vntValor
            mobjLista.SubItems(2) = txtValor
            mobjLista.Tag = gstrItemData(cboDescricao)
        End If
        cbo_intCodigoReduzidoCancelar.Text = ""
        cboConta.ListIndex = -1
        Totaliza lvw_Lista, txtTotal
        txtValor = ""
    End If
End Sub

Sub LimpaDados(cboDescricao As ComboBox, _
    cboConta As ComboBox, _
    Optional txtValor As TextBox, _
    Optional blnAlterando As Boolean)
    If txtValor Is Nothing = False Then
        txtValor = ""
    End If
    cboDescricao.ListIndex = -1
    blnAlterando = False
    cboConta.SetFocus
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    'EnviaTeclaTab vbKeyReturn
    Select Case UCase(strModoOperacao)
    Case UCase(gstrNovo)
        LimpaDadosArrecadacao
    Case UCase(gstrSalvar)
        GravaArrecadacao
    Case UCase(gstrImprimir)
        If Len(Trim(strDataInicial)) = 0 Or Len(Trim(strDataFinal)) = 0 Then
            If blnAtivaFormImprime Then
                frmIntervaloDataArrecadacao.MantemForm (gstrImprimir)
            Else
                CarregaForm frmIntervaloDataArrecadacao
            End If
            
        Else
            ImprimeRelatorio rptReceitaArrecadadaAnulada, strQuerryRelatorio
            strDataInicial = Space$(0)
            strDataFinal = Space$(0)
        End If
    Case UCase(gstrIncluirItem)
        VerificaTabParaIncluir
    Case UCase(gstrExcluirItem)
        VerificaTabParaExcluir
    Case UCase(gstrPreencherLista)
        ''             If Me.ActiveControl.Name = cbointConvenio.Name Then
        ''                LeDaTabelaParaObj gstrConvenio, cbointConvenio
        ''             ElseIf Me.ActiveControl.Name = cbointFundo.Name Then
        ''                LeDaTabelaParaObj gstrFundo, cbointFundo
        '             ElseIf Me.ActiveControl.Name = cbo_intHistorico.Name Then
        '                LeDaTabelaParaObj gstrHistorico, cbo_intHistorico
        ''             ElseIf Me.ActiveControl.Name = cbointOrcamentaria.Name Then
        ''                LePrevisaoReceitaGeral cbointOrcamentaria, cbostrOrcamentaria, strQueryBuscaByEvento
        ''                LeDaTabelaParaObj "", cbo_intCodigoReduzido, strQueryCodigoReduzido
        ''             ElseIf Me.ActiveControl.Name = cbo_intCodigoReduzido.Name Then
        ''                LePrevisaoReceitaGeral cbointOrcamentaria, cbostrOrcamentaria, strQueryBuscaByEvento
        ''                LeDaTabelaParaObj "", cbo_intCodigoReduzido, strQueryCodigoReduzido
        ''             ElseIf Me.ActiveControl.Name = cbointOrcamentariaCancelar.Name Or Me.ActiveControl.Name = cbostrOrcamentariaCancelar.Name Then
        ''                LePrevisaoReceitaGeral cbointOrcamentariaCancelar, cbostrOrcamentariaCancelar, strQueryBuscaByEvento
        ''                LeDaTabelaParaObj "", cbo_intCodigoReduzidoCancelar, strQueryCodigoReduzido
        ''             ElseIf Me.ActiveControl.Name = cbo_intCodigoReduzidoCancelar.Name Then
        ''                LePrevisaoReceitaGeral cbointOrcamentariaCancelar, cbostrOrcamentariaCancelar, strQueryBuscaByEvento
        ''                LeDaTabelaParaObj "", cbo_intCodigoReduzidoCancelar, strQueryCodigoReduzido
        ''             ElseIf Me.ActiveControl.Name = cbointContaContabil.Name Or Me.ActiveControl.Name = dbc_strContaContabil.Name Then
        ''                LePlanoContaGeralBanco cbointContaContabil, dbc_strContaContabil, "DC"
        ''             ElseIf Me.ActiveControl.Name = cbointExtraOrcamentario.Name Or Me.ActiveControl.Name = cbostrExtraOrcamentaria.Name Then
        ''                LePlanoContaGeral cbointExtraOrcamentario, cbostrExtraOrcamentaria, "EO"
        ''             ElseIf Me.ActiveControl.Name = cbointExtraOrcamentarioCancelar.Name Or Me.ActiveControl.Name = cbostrExtraOrcamentariaCancelar.Name Then
        ''                LePlanoContaGeral cbointExtraOrcamentarioCancelar, cbostrExtraOrcamentariaCancelar, "EO"
        ''             ElseIf Me.ActiveControl.Name = cbointEvento.Name Then
        ''                preencheCboevento
        ''             End If
        Select Case Me.ActiveControl.Name
        Case cbointConvenio.Name
            LeDaTabelaParaObj gstrConvenio, cbointConvenio
        Case cbointFundo.Name
            LeDaTabelaParaObj gstrFundo, cbointFundo
        Case cbo_intHistorico.Name
            LeDaTabelaParaObj gstrHistorico, cbo_intHistorico
        Case cbointOrcamentaria.Name
            LePrevisaoReceitaGeral cbointOrcamentaria, cbostrOrcamentaria, strQueryBuscaByEvento
            LeDaTabelaParaObj "", cbo_intCodigoReduzido, strQueryCodigoReduzido
        Case cbo_intCodigoReduzido.Name
            LePrevisaoReceitaGeral cbointOrcamentaria, cbostrOrcamentaria, strQueryBuscaByEvento
            LeDaTabelaParaObj "", cbo_intCodigoReduzido, strQueryCodigoReduzido
        Case cbointOrcamentariaCancelar.Name, cbostrOrcamentariaCancelar.Name
            LePrevisaoReceitaGeral cbointOrcamentariaCancelar, cbostrOrcamentariaCancelar, strQueryBuscaByEvento
            LeDaTabelaParaObj "", cbo_intCodigoReduzidoCancelar, strQueryCodigoReduzido
        Case cbo_intCodigoReduzidoCancelar.Name
            LePrevisaoReceitaGeral cbointOrcamentariaCancelar, cbostrOrcamentariaCancelar, strQueryBuscaByEvento
            LeDaTabelaParaObj "", cbo_intCodigoReduzidoCancelar, strQueryCodigoReduzido
        Case cbointContaContabil.Name, dbc_strContaContabil.Name
            LePlanoContaGeralBanco cbointContaContabil, dbc_strContaContabil, "DC"
        Case cbointExtraOrcamentario.Name, cbostrExtraOrcamentaria.Name
            LePlanoContaGeral cbointExtraOrcamentario, cbostrExtraOrcamentaria, "EO"
        Case cbointExtraOrcamentarioCancelar.Name, cbostrExtraOrcamentariaCancelar.Name
            LePlanoContaGeral cbointExtraOrcamentarioCancelar, cbostrExtraOrcamentariaCancelar, "EO"
        Case cbointEvento.Name
            preencheCboevento
        End Select
        
    Case UCase(gstrDeletar)
        DeletaArrecadacao
    Case UCase(gstrLocalizar)
        ToolBarGeral strModoOperacao, gstrArrecadacaoReceita, mblnAlterandoOrcamentario, tdb_Lista, Me, , strQuery, strQuery
    Case UCase(gstrImportarDados)
        CarregaForm frmImportarDados
    End Select
End Sub

'Private Function gblnDadosCancelamentoOK()
'    If gblnDataValida(txtdtmDataAnulacao) = False Then
'        ExibeMensagem "A data do cancelamento está incorreta."
'        txtdtmDataAnulacao.SetFocus
'    ElseIf CVDate(txtdtmDataAnulacao) < CVDate(tdb_Lista.Columns("dtmData")) Then
'        ExibeMensagem "A data do cancelamento não pode ser inferior a data da arrecadação."
'        txtdtmDataAnulacao.SetFocus
'    ElseIf lvw_CodigoOrcamentarioCancelar.ListItems.Count = 0 _
'    And lvw_ContaContabilCancelar.ListItems.Count = 0 Then
'        ExibeMensagem "Não há lançamento para cancelar."
'        lvw_CodigoOrcamentarioCancelar.SetFocus
'    Else
'        gblnDadosCancelamentoOK = True
'    End If
'End Function

Private Function blnDadosExclusaoOK() As Boolean
    If Val(txtPKId) = 0 Then
        ExibeMensagem "É necessário selecionar um registro para exclusão."
        Exit Function
    End If
    
    If Not CDate(txtdtmData) > VerificaDataEncerramento("EF", gintExercicio) Then
        ExibeMensagem "Não é possivel excluir registros anteriores a data de encerramento financeiro (" & VerificaDataEncerramento("EF", gintExercicio) & ")."
        Exit Function
    End If
    
    blnDadosExclusaoOK = True
    
End Function


Private Sub CancelaArrecadacao()
    
    '******************************************************************************************
    ' Data: 11/06/2003
    ' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL
    '            permitindo, assim, a execução de múltiplos comandos SQL de uma única vez.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    '
    '    Dim strSQL          As String
    '    If gblnDadosCancelamentoOK Then
    '        If gblnExclusaoGravacaoOk("I", "Confirma cancelamento da arrecadação?", True) Then
    '            strSQL = ""
    '
    '            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    '
    '            strSQL = strSQL & strQueryCancelar(lvw_CodigoOrcamentarioCancelar, 0)
    '            strSQL = strSQL & strQueryCancelar(lvw_ContaContabilCancelar, 1)
    '
    '            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
    '
    '            Set gobjBanco = New clsBanco
    '            gobjBanco.ExecutaBeginTrans
    '            Set gobjBanco = New clsBanco
    '            If gobjBanco.Execute(strSQL) Then
    '                Set gobjBanco = New clsBanco
    '                gobjBanco.ExecutaCommitTrans
    '                LeDaTabelaParaObj gstrArrecadacaoReceita, tdb_Lista, strQuery
    '                LimpaDadosArrecadacao
    '            Else
    '                Set gobjBanco = New clsBanco
    '                gobjBanco.ExecutaRollbackTrans
    '            End If
    '        End If
    '    End If
End Sub

Private Sub DeletaArrecadacao()
    
    Dim strSQL          As String
    If blnDadosExclusaoOK Then
        If gblnExclusaoGravacaoOk("I", "Confirma Exclusão da Arrecadação?", True) Then
            
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            
            strSQL = "Delete " & gstrContaArrecadacaoReceita & " where intarrecadacao = " & txtPKId
            If Not gobjBanco.Execute(strSQL) Then GoTo saida
            
            strSQL = "Delete " & gstrArrecadacaoReceita & " where pkid = " & txtPKId
            If Not gobjBanco.Execute(strSQL) Then GoTo saida
            
            strSQL = "Delete " & gstrLancamentoContabil
            strSQL = strSQL & " where intProcesso in ( "
            strSQL = strSQL & "select pkid FROM " & gstrProcessoPagamento
            strSQL = strSQL & " where intlancamento = " & txtintNumero
            strSQL = strSQL & " and intOrigem = 6 "
            strSQL = strSQL & " and dtmdata = " & gstrConvDtParaSql(txtdtmData) & ")"
            If Not gobjBanco.Execute(strSQL) Then GoTo saida
            
            strSQL = " Delete " & gstrProcessoPagamento
            strSQL = strSQL & " where intlancamento = " & txtintNumero
            strSQL = strSQL & " and intOrigem = 6 "
            strSQL = strSQL & " and dtmdata = " & gstrConvDtParaSql(txtdtmData)
            If Not gobjBanco.Execute(strSQL) Then GoTo saida
            
            
            gobjBanco.ExecutaCommitTrans
            MantemForm gstrLocalizar
            MantemForm gstrNovo
            
            
        End If
    End If
    
    Exit Sub
saida:     'ocorre quando algum dos comandos SQL sofre exeption ao executar
    gobjBanco.ExecutaRollbackTrans
    ExibeMensagem "Devido a um erro ao deletar, nenhum registro foi alterado."
    
    
    
End Sub



Private Sub LimpaDadosArrecadacao()
    mblnAlterandoOrcamentario = False
    HabilitaDesabilitaControlesArrecadacao
    txtPKId.Text = ""
    txtintNumero = ""
    txtdtmData = ""
    '    txtdtmDataAnulacao = ""
    txtstrHistorico = ""
    txt_TotalOrcamentario = ""
    txt_TotalExtraOrcamentario = ""
    txt_TotalGeral = ""
    txt_TotalGeralCancelado = ""
    txt_TotalCancelado = ""
    cbo_intHistorico.ListIndex = -1
    cbointConvenio.ListIndex = -1
    cbointFundo.ListIndex = -1
    cbointContaContabil.ListIndex = -1
    TrocaCorObjeto cbointFundo, False
    TrocaCorObjeto cbointConvenio, False
    TrocaCorObjeto cbointContaContabil, False
    TrocaCorObjeto dbc_strContaContabil, False
    TrocaCorObjeto txt_codEvento, False
    TrocaCorObjeto txtdtmData, False
    txt_codEvento.BackColor = vbWindowBackground
    txt_codEvento.Enabled = True
    TrocaCorObjeto cbointEvento, False
    lvw_Orcamentaria.ListItems.Clear
    lvw_ExtraOrcamentaria.ListItems.Clear
    lvw_CodigoOrcamentarioCancelar.ListItems.Clear
    lvw_ContaContabilCancelar.ListItems.Clear
    cbointEvento.Text = ""
    txt_codEvento.Text = ""
    
    txtdtmData = VerificaDataEncerramento("EF", gintExercicio) + 1
    
    dblTotalOrcamentario = 0
    dblTotalExOrcamentario = 0
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    LimpaDados cbostrExtraOrcamentaria, cbointExtraOrcamentario, _
    txt_dblValorExtraOrcamentaria, mblnAlterandoExtra
    LimpaDados cbostrOrcamentaria, cbointOrcamentaria, _
    txt_dblValorOrcamentario, mblnAlterandoOrcamentario
    
    cbointOrcamentariaCancelar.Clear
    cbostrOrcamentariaCancelar.Clear
    
    txtintNumero.SetFocus
    
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar, gstrDeletar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
    
    tab_3DPasta.TabEnabled(0) = True
    tab_3DPasta.TabEnabled(1) = True
    tab_3DPastaCacelar.TabEnabled(0) = True
    tab_3DPastaCacelar.TabEnabled(1) = True
    tab_3DArecadarAnular.TabEnabled(0) = True
    tab_3DArecadarAnular.TabEnabled(1) = True
    
End Sub

Private Sub txt_Arrecadado_GotFocus()
    EnviaTeclaTab 13
End Sub

Private Sub txt_codEvento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_codEvento
End Sub

Private Sub txt_codEvento_LostFocus()
    PreencheEventobyCodigo txt_codEvento, cbointEvento, "1"
    
    DoEvents
    If cbointEvento.ListIndex = -1 Then
        VerificaControlesByEvento (False)
    Else
        VerificaControlesByEvento (True), True
    End If
End Sub

Private Sub txt_CodHistorico_GotFocus()
    MarcaCampo txt_CodHistorico
End Sub

Private Sub txt_CodHistorico_LostFocus()
    Dim adoResultado As ADODB.Recordset
    Set gobjBanco = New clsBanco
    Dim strSQL As String
    strSQL = "SELECT h.StrDescricao FROM " & gstrHistorico & " H WHERE H.STRCODIGO = '" & Me.txt_CodHistorico.Text & "'"
    gobjBanco.CriaADO strSQL, 10, adoResultado
    With adoResultado
        If Not .EOF Then
            Me.cbo_intHistorico.Text = gstrENulo(!strDescricao)
            Me.txtstrHistorico.Text = gstrENulo(!strDescricao)
        Else
            Me.cbo_intCodigoReduzido.Text = ""
        End If
    End With
End Sub

Private Sub txt_dblValorExtraOrcamentaria_GotFocus()
    MarcaCampo txt_dblValorExtraOrcamentaria
End Sub

Private Sub txt_dblValorExtraOrcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorExtraOrcamentaria
End Sub

Private Sub txt_dblValorExtraOrcamentaria_LostFocus()
    txt_dblValorExtraOrcamentaria = gstrConvVrDoSql(txt_dblValorExtraOrcamentaria)
    'AtivaPastaDeObjeto tab_3DArecadarAnular, 0, tab_3DPasta, 1
End Sub
Private Sub txt_dblValorExtraCancelamentoOrcamentaria_GotFocus()
    MarcaCampo txt_dblValorExtraCancelamentoOrcamentaria
End Sub

Private Sub txt_dblValorExtraCancelamentoOrcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorExtraOrcamentaria
End Sub

Private Sub txt_dblValorExtraCancelamentoOrcamentaria_LostFocus()
    txt_dblValorExtraCancelamentoOrcamentaria = gstrConvVrDoSql(txt_dblValorExtraCancelamentoOrcamentaria)
End Sub

Private Sub txt_dblValorOrcamentario_GotFocus()
    MarcaCampo txt_dblValorOrcamentario
    AtivaPastaDeObjeto tab_3DArecadarAnular, 0, tab_3DPasta, 0
End Sub

Private Sub txt_dblValorOrcamentario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorOrcamentario
End Sub

Private Sub txt_dblValorOrcamentario_LostFocus()
    txt_dblValorOrcamentario = gstrConvVrDoSql(txt_dblValorOrcamentario)
End Sub
Private Sub txt_dblValorCancelamentoOrcamentario_GotFocus()
    MarcaCampo txt_dblValorCancelamentoOrcamentario
End Sub

Private Sub txt_dblValorCancelamentoOrcamentario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorCancelamentoOrcamentario
End Sub

Private Sub txt_dblValorCancelamentoOrcamentario_LostFocus()
    txt_dblValorCancelamentoOrcamentario = gstrConvVrDoSql(txt_dblValorCancelamentoOrcamentario)
End Sub

Private Sub txt_Saldo_GotFocus()
    EnviaTeclaTab 13
End Sub

Private Sub txt_ValorConvenio_GotFocus()
    EnviaTeclaTab 13
End Sub

Private Sub txt_ValorConvenio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtdtmData_GotFocus()
    MarcaCampo txtdtmData
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmData
End Sub

Private Sub txtdtmData_LostFocus()
    txtdtmData = gstrDataFormatada(txtdtmData)
    
    'Orc677
    If IsDate(txtdtmData) Then
        If Year(CDate(txtdtmData)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data tem que estar no exercício de " & gintExercicio & "."
            If txtdtmData.Enabled Then txtdtmData.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

'Private Sub txtdtmDataAnulacao_GotFocus()
'    MarcaCampo txtdtmDataAnulacao
'    AtivaPastaDeObjeto tab_3DArecadarAnular, 1
'End Sub

'Private Sub txtdtmDataAnulacao_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "D", txtdtmDataAnulacao
'End Sub

'Private Sub txtdtmDataAnulacao_LostFocus()
'    txtdtmDataAnulacao = gstrDataFormatada(txtdtmDataAnulacao)
'End Sub

Private Sub txtintNumero_GotFocus()
    gstrProximoCodigo txtintNumero, gstrArrecadacaoReceita, "intNumero", gintCodSeguranca, "intExercicio", Val(gintExercicio), , , , , "intExercicio", Val(gintExercicio)
    MarcaCampo txtintNumero
End Sub

Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrHistorico_GotFocus()
    MarcaCampo txtstrHistorico
End Sub

Private Sub txtstrHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Function strQuerryRelatorio() As String
    
    '******************************************************************************************
    ' Data: 09/06/2003
    ' Alteração: - Substituição da estrutura CASE do SQL Server pela função gstrCASEWHEN.
    ' Responsável: Everton Bianchini
    '------------------------------------------------------------------------------------------
    ' Data: 09/06/2003
    ' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
    '            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
    '            representado pela variável strOUTJOracle.
    ' Responsável: Everton Bianchini
    '******************************************************************************************
    
    Dim strSQL As String
    strSQL = ""
    
    strSQL = strSQL & " SELECT DISTINCT "
    strSQL = strSQL & " A.PKId AS PKIdArrecadacaoReceita, "
    strSQL = strSQL & " A.intContaContabil AS INTCONTA,"
    strSQL = strSQL & " A.intNumero AS NumeroDaGuia, "
    strSQL = strSQL & " A.dtmData, "
    strSQL = strSQL & " D.strContaContabil AS NumeroBanco, "
    strSQL = strSQL & " D.strDescricao AS DescricaoBanco, "
    strSQL = strSQL & " C.strDescricao AS Fundo, "
    strSQL = strSQL & " B.strDescricao AS Convenio,"
    strSQL = strSQL & " B.dblValor AS ValorConvenio, "
    
    '    strSql = strSql & " CASE E.bytTipo "
    '    strSql = strSql & " WHEN 0 THEN 'Sim' "
    '    strSql = strSql & " WHEN 1 THEN 'Não' "
    '    strSql = strSql & " END AS Tipo, "
    strSQL = strSQL & gstrCASEWHEN("E.bytTipo", _
    "0, 'Sim', 1, 'Não'") & " AS Tipo, "
    
    '    strSql = strSql & " CASE E.bytCancelado "
    '    strSql = strSql & " WHEN 0 THEN 'Não' "
    '    strSql = strSql & " WHEN 1 THEN 'Sim' "
    '    strSql = strSql & " END AS Cancelado, "
    strSQL = strSQL & gstrCASEWHEN("E.bytCancelado", _
    "0, 'Não', 1, 'Sim'") & " AS Cancelado, "
    
    strSQL = strSQL & " E.dblValorOrcamentario AS ValorDetail, "
    strSQL = strSQL & " E.intConta AS CONTADETAIL, "
    strSQL = strSQL & " E.PKId AS PKIdArrecadado "
    
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrArrecadacaoReceita & " A, "
    strSQL = strSQL & gstrConvenio & " B, "
    strSQL = strSQL & gstrFundo & " C, "
    strSQL = strSQL & gstrPlanoConta & " D, "
    strSQL = strSQL & gstrContaArrecadacaoReceita & " E "
    
    strSQL = strSQL & " WHERE "
    '    strSql = strSql & " A.intConvenio *= B.PKId "
    strSQL = strSQL & " A.intConvenio " & strOUTJSQLServer & "= B.PKId " & strOUTJOracle
    '    strSql = strSql & " AND A.intFundo *= C.PKId "
    strSQL = strSQL & " AND A.intFundo " & strOUTJSQLServer & "= C.PKId " & strOUTJOracle
    strSQL = strSQL & " AND E.intArrecadacao = A.PKId "
    strSQL = strSQL & " AND A.intContaContabil = D.PKId "
    strSQL = strSQL & " AND A.intExercicio = " & gintExercicio
    strSQL = strSQL & " AND A.dtmData >= " & gstrConvDtParaSql(strDataInicial)
    strSQL = strSQL & " AND A.dtmData <= " & gstrConvDtParaSql(strDataFinal)
    
    strSQL = strSQL & " ORDER BY D.strDescricao, A.intNumero "
    
    strQuerryRelatorio = strSQL
End Function
Private Function strQueryIncluiReceita() As String
    Dim strSQL As String
    strSQL = ""
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    strSQL = strSQL & "INSERT INTO " & gstrArrecadacaoReceita & " "
    strSQL = strSQL & "(intNumero, dtmData, intConvenio, intFundo, "
    strSQL = strSQL & "intContaContabil, intExercicio, "
    strSQL = strSQL & "strHistorico, dtmDtAtualizacao, lngCodUsr, intEvento) "
    '  strSql = strSql & "(SELECT ISNULL(MAX(intNumero), 0) + 1, "
    strSQL = strSQL & "VALUES (" & txtintNumero & ", "
    strSQL = strSQL & gstrConvDtParaSql(txtdtmData.Text) & ", "
    strSQL = strSQL & gstrItemData(cbointConvenio, True) & ", "
    strSQL = strSQL & gstrItemData(cbointFundo, True) & ", "
    strSQL = strSQL & gstrItemData(cbointContaContabil) & ", "
    strSQL = strSQL & gintExercicio & ", "
    strSQL = strSQL & " '" & Trim(txtstrHistorico.Text) & "', "
    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & ", "
    strSQL = strSQL & gstrItemData(cbointEvento) & " );"
    strSQL = strSQL & strQueryConta(lvw_Orcamentaria, 0, 0)
    strSQL = strSQL & strQueryConta(lvw_CodigoOrcamentarioCancelar, 1, 0)
    strSQL = strSQL & strQueryConta(lvw_ExtraOrcamentaria, 0, 1)
    strSQL = strSQL & strQueryConta(lvw_ContaContabilCancelar, 1, 1)
    
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
    
    strQueryIncluiReceita = strSQL
    
End Function
Private Sub HabilitaDesabilitaControlesArrecadacao()
    
    TrocaCorObjeto cbointConvenio, mblnAlterandoOrcamentario
    TrocaCorObjeto cmd_Convenio, mblnAlterandoOrcamentario
    TrocaCorObjeto cbointFundo, mblnAlterandoOrcamentario
    TrocaCorObjeto cmd_Fundo, mblnAlterandoOrcamentario
    TrocaCorObjeto cbointContaContabil, mblnAlterandoOrcamentario
    TrocaCorObjeto dbc_strContaContabil, mblnAlterandoOrcamentario
    TrocaCorObjeto cmd_BancoArrecadador, mblnAlterandoOrcamentario
    
    TrocaCorObjeto cbointEvento, mblnAlterandoOrcamentario
    TrocaCorObjeto cmd_Evento, mblnAlterandoOrcamentario
    
End Sub
Private Function strQueryAlteraReceita() As String
    Dim strSQL As String
    
    strSQL = "UPDATE " & gstrArrecadacaoReceita
    strSQL = strSQL & " SET strHistorico = '" & txtstrHistorico & "'"
    ' strSql = strSql & " , dtmData = " & gstrConvDtParaSql(txtdtmData)
    strSQL = strSQL & " WHERE PKID = " & Val(txtPKId)
    
    strQueryAlteraReceita = strSQL
    
End Function
Private Function strQueryAplicar() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " "
    strSQL = strSQL & "WHERE ABS(blnFinanceira) = 1 "
    strSQL = strSQL & "AND ABS(blnAnalitica) = 1"
    strQueryAplicar = strSQL
End Function

Private Function strQueryAplicarEvento() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEvento & " "
    strSQL = strSQL & "WHERE intTipoEvento = 1 "
    strQueryAplicarEvento = strSQL
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra
End Function


Private Function VerificaBancoConvenio()
    Dim strSQL       As String
    Dim adoResultado As New ADODB.Recordset
    
    
    strSQL = "SELECT PC.PKID, CV.intContaContabil FROM " & gstrConvenio & " CV, "
    strSQL = strSQL & gstrPlanoConta & " PC "
    strSQL = strSQL & " WHERE CV.PKID = " & cbointConvenio.ItemData(cbointConvenio.ListIndex)
    strSQL = strSQL & " AND PC.PKID = CV.intContaContabil"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
        
        With adoResultado
            If .EOF = False Then
                LePlanoContaGeralBanco cbointContaContabil, dbc_strContaContabil, "DC"
                cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, gstrVerificaCampoNulo(!Pkid))
                TrocaCorObjeto cbointContaContabil, True
                TrocaCorObjeto dbc_strContaContabil, True
                cmd_BancoArrecadador.Enabled = False
                DoEvents
            End If
        End With
        
    End If
End Function
Private Function VerificaBancoFundo()
    Dim strSQL       As String
    Dim adoResultado As New ADODB.Recordset
    
    
    strSQL = "SELECT PC.PKID, FU.intContaContabil FROM " & gstrFundo & " FU, "
    strSQL = strSQL & gstrPlanoConta & " PC "
    strSQL = strSQL & " WHERE FU.PKID = " & cbointFundo.ItemData(cbointFundo.ListIndex)
    strSQL = strSQL & " AND PC.PKID = FU.intContaContabil"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
        
        With adoResultado
            If .EOF = False Then
                LePlanoContaGeralBanco cbointContaContabil, dbc_strContaContabil, "DC"
                cbointContaContabil.ListIndex = gintIndiceCBO(cbointContaContabil, gstrVerificaCampoNulo(!Pkid))
                TrocaCorObjeto cbointContaContabil, True
                TrocaCorObjeto dbc_strContaContabil, True
                cmd_BancoArrecadador.Enabled = False
                DoEvents
            End If
        End With
        
    End If
    
End Function

Private Function VerificaDtIntervaloConvenio(strData As String) As Boolean
    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset
    Dim dtDataInicial As Date
    Dim dtDataFinal   As Date
    
    VerificaDtIntervaloConvenio = True
    
    strSQL = "SELECT dtmDataAplicacaoInicial, dtmDataAplicacaoFinal FROM " & gstrConvenio
    strSQL = strSQL & " WHERE PKID = " & cbointConvenio.ItemData(cbointConvenio.ListIndex)
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                dtDataInicial = CDate(!dtmDataAplicacaoInicial)
                dtDataFinal = CDate(!dtmDataAplicacaoFinal)
                If CDate(strData) < dtDataInicial Or CDate(strData) > dtDataFinal Then
                    VerificaDtIntervaloConvenio = False
                End If
            End If
        End With
    End If
End Function


Private Sub preencheCboevento()
    LeDaTabelaParaObj gstrEvento, cbointEvento, "SELECT PKID, strDescricao FROM " & gstrEvento & " WHERE intTipoEvento=1"
    'Tipo Evento : 0-orcamento, 1-arrecadado, 2-empenho, 3-Pagto Empenho
    '              4-Pagto Resto a Pagar, 5-Pagto-Extra
End Sub
Private Sub VerificaControlesByEvento(blnHabilitar As Boolean, Optional blnVerificaDados As Boolean = False)
    Dim intIdx As Integer
    
    
    tab_3DPasta.TabEnabled(0) = blnHabilitar
    tab_3DPastaCacelar.TabEnabled(0) = blnHabilitar
    tab_3DPasta.TabEnabled(1) = Not blnHabilitar
    tab_3DPastaCacelar.TabEnabled(1) = Not blnHabilitar
    
    If blnVerificaDados = False Then
        lvw_Orcamentaria.ListItems.Clear
        lvw_CodigoOrcamentarioCancelar.ListItems.Clear
        cbointOrcamentaria.Clear
        cbostrOrcamentaria.Clear
        cbointOrcamentariaCancelar.Clear
        cbostrOrcamentariaCancelar.Clear
        cbo_intCodigoReduzido.Clear
        cbo_intCodigoReduzidoCancelar.Clear
        txt_TotalOrcamentario.Text = gstrConvVrDoSql("0")
        txt_TotalExtraOrcamentario.Text = gstrConvVrDoSql("0")
        txt_TotalCancelado.Text = gstrConvVrDoSql("0")
        txt_TotalGeral.Text = gstrConvVrDoSql("0")
        txt_TotalGeralCancelado.Text = gstrConvVrDoSql("0")
    End If
    
    If blnHabilitar Then
        tab_3DPasta.Tab = 0
        tab_3DPastaCacelar.Tab = 0
        
        With lvw_ContaContabilCancelar
            For intIdx = 1 To .ListItems.Count
                .ListItems(intIdx).Selected = True
                ExcluirItemDaLista lvw_ContaContabilCancelar, _
                cbointExtraOrcamentarioCancelar, _
                mblnAlterandoExtraOrcamentariaCancelar, _
                txt_TotalGeralCancelado
                dblTotalExOrcamentario = CDbl(gstrConvVrDoSql(txt_TotalGeralCancelado))
                txt_TotalGeralCancelado = gstrConvVrDoSql(dblTotalOrcamentario + dblTotalExOrcamentario)
                txt_TotalCancelado = txt_TotalGeralCancelado
                txt_TotalGeral = gstrConvVrDoSql(CDbl(txt_TotalGeral) - CDbl(txt_TotalCancelado))
            Next
        End With
        
        With lvw_ExtraOrcamentaria
            For intIdx = 1 To .ListItems.Count
                .ListItems(intIdx).Selected = True
                ExcluirItemDaLista lvw_ExtraOrcamentaria, _
                cbointExtraOrcamentario, _
                mblnAlterandoExtra, _
                txt_TotalExtraOrcamentario, _
                txt_dblValorExtraOrcamentaria
            Next
        End With
        
    Else
        tab_3DPasta.Tab = 1
        tab_3DPastaCacelar.Tab = 1
        
        With lvw_CodigoOrcamentarioCancelar
            For intIdx = 1 To .ListItems.Count
                .ListItems(intIdx).Selected = True
                ExcluirItemDaLista lvw_CodigoOrcamentarioCancelar, _
                cbointOrcamentariaCancelar, _
                mblnAlterandoOrcamentarioCancelar, _
                txt_TotalGeralCancelado
                dblTotalOrcamentario = CDbl(gstrConvVrDoSql(txt_TotalGeralCancelado))
                txt_TotalGeralCancelado = gstrConvVrDoSql(dblTotalOrcamentario + dblTotalExOrcamentario)
                txt_TotalCancelado = txt_TotalGeralCancelado
                txt_TotalGeral = gstrConvVrDoSql(CDbl(txt_TotalGeral) - CDbl(txt_TotalCancelado))
            Next
        End With
        
        With lvw_Orcamentaria
            For intIdx = 1 To .ListItems.Count
                .ListItems(intIdx).Selected = True
                ExcluirItemDaLista lvw_Orcamentaria, cbointOrcamentaria, _
                mblnAlterandoOrcamentario, _
                txt_TotalOrcamentario, _
                txt_dblValorOrcamentario
                
            Next
        End With
        
    End If
    
End Sub
Private Function strQueryBuscaByEvento() As String
    Dim strSQL As String
    Dim adoResultado As New ADODB.Recordset
    Dim blnConvenio As Boolean
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT intPrevisaoDaReceita FROM " & gstrConvenio & " WHERE PKID = " & gstrItemData(cbointConvenio, True), 10, adoResultado) Then
        If Not IsNull(adoResultado!intPrevisaoDaReceita) And cbointConvenio.ListIndex <> -1 Then
            blnConvenio = True
        Else
            blnConvenio = False
        End If
    End If
    
    strSQL = "SELECT CO.PKId, CO.strCodigoOrcamentario, CO.strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrCodigoOrcamentario & " CO, "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR "
    
    If blnConvenio Then
        strSQL = strSQL & ", " & gstrConvenio & " CV "
    End If
    
    strSQL = strSQL & "WHERE CO.PKId = PR.intCodigoOrcamentario "
    strSQL = strSQL & "AND PR.intExercicio = " & gintExercicio & " "
    strSQL = strSQL & "AND " & strSUBSTRING & "(CO.strCodigoOrcamentario,1," & Len(BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoReceita, "C", 1)) & ") = '" & _
    BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoReceita, "C", 1) & "'"
    
    If blnConvenio Then
        strSQL = strSQL & " AND CV.intPrevisaoDaReceita = PR.PKID"
        strSQL = strSQL & " AND CV.PKID = " & gstrItemData(cbointConvenio, True) & " "
    End If
    
    strSQL = strSQL & "ORDER BY CO.strDescricao"
    
    strQueryBuscaByEvento = strSQL
    
End Function

Private Function strQueryCodigoReduzido() As String
    Dim strSQL As String
    Dim adoResultado As New ADODB.Recordset
    Dim blnConvenio As Boolean
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO("SELECT intPrevisaoDaReceita FROM " & gstrConvenio & " WHERE PKID = " & gstrItemData(cbointConvenio, True), 10, adoResultado) Then
        If Not IsNull(adoResultado!intPrevisaoDaReceita) And cbointConvenio.ListIndex <> -1 Then
            blnConvenio = True
        Else
            blnConvenio = False
        End If
    End If
    
    strSQL = "SELECT CO.PKId, PR.intCodigoReduzido"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrCodigoOrcamentario & " CO, "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR "
    
    If blnConvenio Then
        strSQL = strSQL & ", " & gstrConvenio & " CV "
    End If
    
    strSQL = strSQL & "WHERE CO.PKId = PR.intCodigoOrcamentario "
    strSQL = strSQL & "AND PR.intExercicio = " & gintExercicio & " "
    strSQL = strSQL & "AND " & strSUBSTRING & "(CO.strCodigoOrcamentario,1," & Len(BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoReceita, "C", 1)) & ") = '" & _
    BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoReceita, "C", 1) & "'"
    
    If blnConvenio Then
        strSQL = strSQL & " AND CV.intPrevisaoDaReceita = PR.PKID"
        strSQL = strSQL & " AND CV.PKID = " & gstrItemData(cbointConvenio, True) & " "
    End If
    
    
    strSQL = strSQL & " ORDER BY CO.strDescricao "
    
    strQueryCodigoReduzido = strSQL
    
End Function

Private Function VerificaValorConvenio() As Integer
    Dim i As Integer
    Dim dblValorTotal As Double
    Dim strSQL As String
    Dim adoResultado As New ADODB.Recordset
    
    VerificaValorConvenio = 0
    
    '0 ---> Valor da tblConvenio eh menor que o que esta sendo gravado
    '1 ---> Valor da tblConvenio eh igual  ao  que esta sendo gravado
    '2 ---> Valor da tblConvenio eh maior que o que esta sendo gravado
    
    For i = 1 To lvw_Orcamentaria.ListItems.Count
        dblValorTotal = dblValorTotal + CDbl(lvw_Orcamentaria.ListItems(i).ListSubItems(2))
    Next
    
    For i = 1 To lvw_ExtraOrcamentaria.ListItems.Count
        dblValorTotal = dblValorTotal + CDbl(lvw_ExtraOrcamentaria.ListItems(i).ListSubItems(2))
    Next
    
    strSQL = "SELECT dblValor FROM " & gstrConvenio & " WHERE PKID = " & gstrItemData(cbointConvenio, True)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If adoResultado.RecordCount >= 1 Then
            If adoResultado!dblValor < dblValorTotal Then
                VerificaValorConvenio = 0
            ElseIf adoResultado!dblValor = dblValorTotal Then
                VerificaValorConvenio = 1
            ElseIf adoResultado!dblValor = dblValorTotal Then
                VerificaValorConvenio = 2
            End If
        End If
    End If
    
End Function
Private Function strRegistroAutenticacao(intNumeroGuia As String) As String
    '********************************************************************************************************************************************************************************
    ' Create by:         Éder Henrique
    ' Módulos:           Orçamentário
    ' Data:              02/05/2006
    ' Ficha:             orc1340
    ' Comentários:       Concatena o registro a ser impresso
    '********************************************************************************************************************************************************************************
    Dim strSQL                          As String
    Dim strAux                          As String
    Dim adoResult                       As New ADODB.Recordset
    
    Err.Clear
    
    On Local Error GoTo ERRO_strRegistroAutenticacao
    
    Set gobjBanco = New clsBanco
    
    'Consulto autenticação
    strSQL = "SELECT AR.intNumero, AR.dtmData, CB.intNumeroConta" & _
    "  FROM tblPlanoConta PC, tblContaBancaria CB, tblarrecadacaoreceita AR" & _
    "  WHERE CB.PKId = PC.intContaBancaria" & _
    "  AND AR.intContaContabil = PC.Pkid" & _
    "  AND PC.Pkid = " & gstrItemData(cbointContaContabil) & _
    "  AND AR.intNumero = " & intNumeroGuia
    
    If gobjBanco.CriaADO(strSQL, 5, adoResult) Then
        If Not adoResult.EOF Then
            strAux = Right(String(6, "0") & adoResult.Fields("intNumero"), 6) & " " & Format(adoResult.Fields("dtmData"), "dd/mm/yyyy") & " " & Right(String(6, "0") & adoResult.Fields("intNumeroConta"), 6) & " "
        End If
        adoResult.Close
    End If
    
    strRegistroAutenticacao = strAux
    
ERRO_strRegistroAutenticacao:
    If Err.Number <> 0 Then
        ExibeMensagem "Ocorreu o erro: " + Str(Err) + vbCrLf + Err.Description + vbCrLf + "Em: Função strRegistroAutenticacao "
        Err.Clear
        strRegistroAutenticacao = ""
    End If
End Function


