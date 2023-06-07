VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadContribuinteN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contribuintes"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   HelpContextID   =   109
   Icon            =   "CadContribuinteN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9675
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   7410
      TabIndex        =   0
      Top             =   -150
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6645
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   11721
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Contribuinte"
      TabPicture(0)   =   "CadContribuinteN.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Codigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrNome"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrCNPJCPF"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbldtmDataCadastro"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_Lista"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "mskstrCNPJCPF"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tab_3DDadosGerais"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_Codigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstrNome"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_bytNaturezaJuridica"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtdtmDataCadastro"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkblnResidenteNoMunicipio"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.CheckBox chkblnResidenteNoMunicipio 
         Caption         =   "Residente no município"
         Height          =   195
         Left            =   7410
         TabIndex        =   69
         Top             =   420
         Width           =   1995
      End
      Begin VB.TextBox txtdtmDataCadastro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6120
         OLEDragMode     =   1  'Automatic
         TabIndex        =   67
         Top             =   390
         Width           =   1005
      End
      Begin VB.Frame fra_bytNaturezaJuridica 
         Caption         =   " Natureza jurídica "
         Height          =   555
         Left            =   90
         TabIndex        =   59
         Top             =   1050
         Width           =   9315
         Begin VB.OptionButton optbytNaturezaJuridica 
            Caption         =   "Outros"
            Height          =   195
            Index           =   3
            Left            =   6930
            TabIndex        =   63
            Top             =   270
            Width           =   795
         End
         Begin VB.OptionButton optbytNaturezaJuridica 
            Caption         =   "SC"
            Height          =   195
            Index           =   2
            Left            =   5150
            TabIndex        =   62
            Top             =   270
            Width           =   705
         End
         Begin VB.OptionButton optbytNaturezaJuridica 
            Caption         =   "Física"
            Height          =   195
            Index           =   0
            Left            =   1050
            TabIndex        =   61
            Top             =   270
            Width           =   915
         End
         Begin VB.OptionButton optbytNaturezaJuridica 
            Caption         =   "Jurídica"
            Height          =   195
            Index           =   1
            Left            =   3040
            TabIndex        =   60
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.TextBox txtstrNome 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   100
         TabIndex        =   58
         Top             =   720
         Width           =   8445
      End
      Begin VB.TextBox txt_Codigo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Left            =   960
         MaxLength       =   19
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   390
         Width           =   1605
      End
      Begin TabDlg.SSTab tab_3DDadosGerais 
         Height          =   2415
         HelpContextID   =   13
         Left            =   60
         TabIndex        =   2
         Top             =   1650
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4260
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Dados Gerais"
         TabPicture(0)   =   "CadContribuinteN.frx":105E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblstrNomeFantasia"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblstrInscricaoEstadual"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblstrCarteiraTrabalho"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbldtmDataNascimento"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblstrTituloEleitoral"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblstrIdentidade"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtstrNomeFantasia"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtstrInscricaoEstadual"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtstrCarteiraTrabalho"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtdtmDataNascimento"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtstrTituloEleitoral"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtstrIdentidade"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "Endereço"
         TabPicture(1)   =   "CadContribuinteN.frx":107A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "tab_3DCorrespondencia"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Comunicações"
         TabPicture(2)   =   "CadContribuinteN.frx":1096
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lbl_DescricaoConteudo"
         Tab(2).Control(1)=   "lbl_TipoComunicacao"
         Tab(2).Control(2)=   "lvw_TipoComunicacao"
         Tab(2).Control(3)=   "txt_DescricaoConteudo"
         Tab(2).Control(4)=   "txt_Conteudo"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Conta Bancária"
         TabPicture(3)   =   "CadContribuinteN.frx":10B2
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lbl_DV"
         Tab(3).Control(1)=   "lbl_Agencia"
         Tab(3).Control(2)=   "lbl_Banco"
         Tab(3).Control(3)=   "lbl_Conta"
         Tab(3).Control(4)=   "lbl_dtmDebito"
         Tab(3).Control(5)=   "lvw_Contas"
         Tab(3).Control(6)=   "txt_DigitoVerificador"
         Tab(3).Control(7)=   "txt_Conta"
         Tab(3).Control(8)=   "cbointBanco"
         Tab(3).Control(9)=   "cbo_strBanco"
         Tab(3).Control(10)=   "Combo1"
         Tab(3).Control(11)=   "Combo2"
         Tab(3).Control(12)=   "cmd_Banco"
         Tab(3).Control(12).Enabled=   0   'False
         Tab(3).Control(13)=   "cmd_Agencia"
         Tab(3).Control(13).Enabled=   0   'False
         Tab(3).Control(14)=   "txt_dtmDebito"
         Tab(3).Control(15)=   "chk_DebitoAutomatico"
         Tab(3).ControlCount=   16
         TabCaption(4)   =   "Histórico"
         TabPicture(4)   =   "CadContribuinteN.frx":10CE
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "tdb_Historico"
         Tab(4).ControlCount=   1
         Begin VB.CheckBox chk_DebitoAutomatico 
            Caption         =   "Débito Automático"
            Height          =   195
            Left            =   -67500
            TabIndex        =   83
            Top             =   1140
            Width           =   1725
         End
         Begin VB.TextBox txt_dtmDebito 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   -69450
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   1110
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Agencia 
            Height          =   300
            Left            =   -66120
            Picture         =   "CadContribuinteN.frx":10EA
            Style           =   1  'Graphical
            TabIndex        =   80
            TabStop         =   0   'False
            ToolTipText     =   "Cliqui aqui para cadastrar agência"
            Top             =   750
            Width           =   360
         End
         Begin VB.CommandButton cmd_Banco 
            Height          =   300
            Left            =   -66120
            Picture         =   "CadContribuinteN.frx":1474
            Style           =   1  'Graphical
            TabIndex        =   79
            TabStop         =   0   'False
            ToolTipText     =   "Cliqui aqui para cadastrar banco"
            Top             =   390
            Width           =   360
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   -73320
            TabIndex        =   78
            Top             =   750
            Width           =   7185
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   -74250
            TabIndex        =   77
            Top             =   750
            Width           =   975
         End
         Begin VB.ComboBox cbo_strBanco 
            Height          =   315
            Left            =   -73320
            TabIndex        =   76
            Top             =   390
            Width           =   7185
         End
         Begin VB.ComboBox cbointBanco 
            Height          =   315
            Left            =   -74250
            TabIndex        =   75
            Top             =   390
            Width           =   975
         End
         Begin VB.TextBox txtstrIdentidade 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   7635
            MaxLength       =   20
            TabIndex        =   12
            Top             =   750
            Width           =   1605
         End
         Begin VB.TextBox txtstrTituloEleitoral 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1455
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1110
            Width           =   1605
         End
         Begin VB.TextBox txtdtmDataNascimento 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4815
            TabIndex        =   10
            Top             =   1110
            Width           =   1605
         End
         Begin VB.TextBox txtstrCarteiraTrabalho 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4815
            MaxLength       =   20
            TabIndex        =   9
            Top             =   750
            Width           =   1605
         End
         Begin VB.TextBox txtstrInscricaoEstadual 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1455
            MaxLength       =   20
            TabIndex        =   8
            Top             =   750
            Width           =   1605
         End
         Begin VB.TextBox txtstrNomeFantasia 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1455
            MaxLength       =   50
            TabIndex        =   7
            Top             =   420
            Width           =   7725
         End
         Begin VB.TextBox txt_Conteudo 
            Height          =   285
            Left            =   -74520
            MaxLength       =   50
            TabIndex        =   6
            Top             =   420
            Width           =   3705
         End
         Begin VB.TextBox txt_DescricaoConteudo 
            Height          =   285
            Left            =   -69690
            MaxLength       =   50
            TabIndex        =   5
            Top             =   420
            Width           =   3945
         End
         Begin VB.TextBox txt_Conta 
            Height          =   285
            Left            =   -74250
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1110
            Width           =   2055
         End
         Begin VB.TextBox txt_DigitoVerificador 
            Height          =   285
            Left            =   -71700
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1110
            Width           =   525
         End
         Begin TabDlg.SSTab tab_3DCorrespondencia 
            Height          =   1935
            Left            =   -74940
            TabIndex        =   13
            Top             =   390
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   3413
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Residencial"
            TabPicture(0)   =   "CadContribuinteN.frx":17FE
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblintCep"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lblintUF"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lblintLogradouro"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lblintBairro"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lblintMunicipio"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lblstrComplemento"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "DataCombo1"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "cbointLogradouro"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "cbointUF"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "cbointBairro"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "cbointMunicipio"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "cmd_Municipio"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "cmd_Bairro"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "cmd_Logradouro"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "txtintCep"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "txtstrComplemento"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "txtintNumero"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "cmd_TipoLog"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "cmd_UFResidencial"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).ControlCount=   19
            TabCaption(1)   =   "Correspondência"
            TabPicture(1)   =   "CadContribuinteN.frx":181A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lblintMunicipioC"
            Tab(1).Control(1)=   "lblintBairroC"
            Tab(1).Control(2)=   "lblintLogradouroC"
            Tab(1).Control(3)=   "lblstrComplementoC"
            Tab(1).Control(4)=   "lblintUFC"
            Tab(1).Control(5)=   "lblintCepC"
            Tab(1).Control(6)=   "DataCombo2"
            Tab(1).Control(7)=   "DataCombo3"
            Tab(1).Control(8)=   "cbointLogradouroCorresp"
            Tab(1).Control(9)=   "cbointUFC"
            Tab(1).Control(10)=   "cbointMunicipioC"
            Tab(1).Control(11)=   "cmd_MunicipioC"
            Tab(1).Control(11).Enabled=   0   'False
            Tab(1).Control(12)=   "txtintNumeroC"
            Tab(1).Control(13)=   "txtstrComplementoC"
            Tab(1).Control(14)=   "txtintCepC"
            Tab(1).Control(15)=   "cmd_TipoLogCorrespendencia"
            Tab(1).Control(15).Enabled=   0   'False
            Tab(1).Control(16)=   "cmd_LogrCorrespondencia"
            Tab(1).Control(16).Enabled=   0   'False
            Tab(1).Control(17)=   "Command1"
            Tab(1).Control(17).Enabled=   0   'False
            Tab(1).Control(18)=   "Command7"
            Tab(1).Control(18).Enabled=   0   'False
            Tab(1).ControlCount=   19
            TabCaption(2)   =   "Domiciliar"
            TabPicture(2)   =   "CadContribuinteN.frx":1836
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label1"
            Tab(2).Control(1)=   "Label2"
            Tab(2).Control(2)=   "Label3"
            Tab(2).Control(3)=   "Label4"
            Tab(2).Control(4)=   "Label5"
            Tab(2).Control(5)=   "Label6"
            Tab(2).Control(6)=   "DataCombo8"
            Tab(2).Control(7)=   "DataCombo7"
            Tab(2).Control(8)=   "DataCombo6"
            Tab(2).Control(9)=   "DataCombo5"
            Tab(2).Control(10)=   "DataCombo4"
            Tab(2).Control(11)=   "Command2"
            Tab(2).Control(11).Enabled=   0   'False
            Tab(2).Control(12)=   "Command3"
            Tab(2).Control(12).Enabled=   0   'False
            Tab(2).Control(13)=   "Text1"
            Tab(2).Control(14)=   "Text2"
            Tab(2).Control(15)=   "Text3"
            Tab(2).Control(16)=   "Command4"
            Tab(2).Control(16).Enabled=   0   'False
            Tab(2).Control(17)=   "Command5"
            Tab(2).Control(17).Enabled=   0   'False
            Tab(2).Control(18)=   "Command6"
            Tab(2).Control(18).Enabled=   0   'False
            Tab(2).ControlCount=   19
            Begin VB.CommandButton Command6 
               Height          =   300
               Left            =   -69210
               Picture         =   "CadContribuinteN.frx":1852
               Style           =   1  'Graphical
               TabIndex        =   98
               TabStop         =   0   'False
               Tag             =   "53"
               ToolTipText     =   "Ativa cadastro de Município"
               Top             =   1155
               Width           =   330
            End
            Begin VB.CommandButton Command5 
               Height          =   300
               Left            =   -66180
               Picture         =   "CadContribuinteN.frx":1BDC
               Style           =   1  'Graphical
               TabIndex        =   97
               TabStop         =   0   'False
               Tag             =   "581"
               ToolTipText     =   "Ativa cadastro de Bairros"
               Top             =   780
               Width           =   330
            End
            Begin VB.CommandButton Command4 
               Height          =   300
               Left            =   -67050
               Picture         =   "CadContribuinteN.frx":1F66
               Style           =   1  'Graphical
               TabIndex        =   96
               TabStop         =   0   'False
               Tag             =   "584"
               ToolTipText     =   "Ativa cadastro de logradouro"
               Top             =   435
               Width           =   330
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   -66945
               MaxLength       =   9
               TabIndex        =   95
               Top             =   1185
               Width           =   1080
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   -73890
               MaxLength       =   20
               TabIndex        =   94
               Top             =   795
               Width           =   3510
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   -66720
               MaxLength       =   8
               TabIndex        =   93
               Top             =   450
               Width           =   855
            End
            Begin VB.CommandButton Command3 
               Height          =   300
               Left            =   -73140
               Picture         =   "CadContribuinteN.frx":22F0
               Style           =   1  'Graphical
               TabIndex        =   92
               TabStop         =   0   'False
               ToolTipText     =   "Clique para cadastar o tipo do logradouro"
               Top             =   435
               Width           =   330
            End
            Begin VB.CommandButton Command2 
               Height          =   300
               Left            =   -67740
               Picture         =   "CadContribuinteN.frx":267A
               Style           =   1  'Graphical
               TabIndex        =   91
               TabStop         =   0   'False
               Tag             =   "581"
               ToolTipText     =   "Ativa cadastrar UF"
               Top             =   1170
               Width           =   330
            End
            Begin VB.CommandButton Command7 
               Height          =   300
               Left            =   -67740
               Picture         =   "CadContribuinteN.frx":2A04
               Style           =   1  'Graphical
               TabIndex        =   90
               TabStop         =   0   'False
               Tag             =   "581"
               ToolTipText     =   "Ativa cadastrar UF"
               Top             =   1170
               Width           =   330
            End
            Begin VB.CommandButton Command1 
               Height          =   300
               Left            =   -66180
               Picture         =   "CadContribuinteN.frx":2D8E
               Style           =   1  'Graphical
               TabIndex        =   88
               TabStop         =   0   'False
               Tag             =   "581"
               ToolTipText     =   "Ativa cadastro de Bairros"
               Top             =   780
               Width           =   330
            End
            Begin VB.CommandButton cmd_LogrCorrespondencia 
               Height          =   300
               Left            =   -67050
               Picture         =   "CadContribuinteN.frx":3118
               Style           =   1  'Graphical
               TabIndex        =   85
               TabStop         =   0   'False
               Tag             =   "584"
               ToolTipText     =   "Ativa cadastro de logradouro"
               Top             =   435
               Width           =   330
            End
            Begin VB.CommandButton cmd_TipoLogCorrespendencia 
               Height          =   300
               Left            =   -73140
               Picture         =   "CadContribuinteN.frx":34A2
               Style           =   1  'Graphical
               TabIndex        =   84
               TabStop         =   0   'False
               ToolTipText     =   "Clique para cadastar o tipo do logradouro"
               Top             =   435
               Width           =   330
            End
            Begin VB.CommandButton cmd_UFResidencial 
               Height          =   300
               Left            =   7260
               Picture         =   "CadContribuinteN.frx":382C
               Style           =   1  'Graphical
               TabIndex        =   74
               TabStop         =   0   'False
               Tag             =   "581"
               ToolTipText     =   "Ativa cadastrar UF"
               Top             =   1170
               Width           =   330
            End
            Begin VB.CommandButton cmd_TipoLog 
               Height          =   300
               Left            =   1860
               Picture         =   "CadContribuinteN.frx":3BB6
               Style           =   1  'Graphical
               TabIndex        =   71
               TabStop         =   0   'False
               ToolTipText     =   "Clique para cadastar o tipo do logradouro"
               Top             =   435
               Width           =   330
            End
            Begin VB.TextBox txtintCepC 
               Height          =   285
               Left            =   -66930
               MaxLength       =   9
               TabIndex        =   25
               Top             =   1185
               Width           =   1080
            End
            Begin VB.TextBox txtstrComplementoC 
               Height          =   285
               Left            =   -73890
               MaxLength       =   20
               TabIndex        =   24
               Top             =   802
               Width           =   3510
            End
            Begin VB.TextBox txtintNumeroC 
               Height          =   285
               Left            =   -66720
               MaxLength       =   8
               TabIndex        =   23
               Top             =   450
               Width           =   855
            End
            Begin VB.TextBox txtintNumero 
               Height          =   285
               Left            =   8280
               MaxLength       =   8
               TabIndex        =   22
               Top             =   450
               Width           =   855
            End
            Begin VB.TextBox txtstrComplemento 
               Height          =   285
               Left            =   1110
               MaxLength       =   20
               TabIndex        =   21
               Top             =   802
               Width           =   3510
            End
            Begin VB.TextBox txtintCep 
               Height          =   285
               Left            =   8070
               MaxLength       =   9
               TabIndex        =   20
               Top             =   1185
               Width           =   1080
            End
            Begin VB.CommandButton cmd_Logradouro 
               Height          =   300
               Left            =   7950
               Picture         =   "CadContribuinteN.frx":3F40
               Style           =   1  'Graphical
               TabIndex        =   19
               TabStop         =   0   'False
               Tag             =   "584"
               ToolTipText     =   "Ativa cadastro de logradouro"
               Top             =   435
               Width           =   330
            End
            Begin VB.CommandButton cmd_Bairro 
               Height          =   300
               Left            =   8820
               Picture         =   "CadContribuinteN.frx":42CA
               Style           =   1  'Graphical
               TabIndex        =   18
               TabStop         =   0   'False
               Tag             =   "581"
               ToolTipText     =   "Ativa cadastro de Bairros"
               Top             =   780
               Width           =   330
            End
            Begin VB.CommandButton cmd_Municipio 
               Height          =   300
               Left            =   5790
               Picture         =   "CadContribuinteN.frx":4654
               Style           =   1  'Graphical
               TabIndex        =   17
               TabStop         =   0   'False
               Tag             =   "53"
               ToolTipText     =   "Ativa cadastro de Município"
               Top             =   1155
               Width           =   330
            End
            Begin VB.CommandButton cmd_MunicipioC 
               Height          =   300
               Left            =   -69210
               Picture         =   "CadContribuinteN.frx":49DE
               Style           =   1  'Graphical
               TabIndex        =   16
               TabStop         =   0   'False
               ToolTipText     =   "Ativa cadastro de Município"
               Top             =   1155
               Width           =   330
            End
            Begin MSDataListLib.DataCombo cbointMunicipio 
               Height          =   315
               Left            =   1110
               TabIndex        =   14
               Top             =   1155
               Width           =   4665
               _ExtentX        =   8229
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cbointBairro 
               Height          =   315
               Left            =   5295
               TabIndex        =   15
               Top             =   787
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cbointUF 
               Height          =   315
               Left            =   6615
               TabIndex        =   26
               Top             =   1155
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cbointLogradouro 
               Height          =   315
               Left            =   2160
               TabIndex        =   27
               Top             =   420
               Width           =   5805
               _ExtentX        =   10239
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cbointMunicipioC 
               Height          =   315
               Left            =   -73890
               TabIndex        =   28
               Top             =   1155
               Width           =   4665
               _ExtentX        =   8229
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cbointUFC 
               Height          =   315
               Left            =   -68385
               TabIndex        =   29
               Top             =   1155
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Left            =   1110
               TabIndex        =   72
               Top             =   420
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cbointLogradouroCorresp 
               Height          =   315
               Left            =   -72840
               TabIndex        =   86
               Top             =   420
               Width           =   5805
               _ExtentX        =   10239
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DataCombo3 
               Height          =   315
               Left            =   -73890
               TabIndex        =   87
               Top             =   420
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DataCombo2 
               Height          =   315
               Left            =   -69705
               TabIndex        =   89
               Top             =   787
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   -73890
               TabIndex        =   99
               Top             =   1155
               Width           =   4665
               _ExtentX        =   8229
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DataCombo5 
               Height          =   315
               Left            =   -69705
               TabIndex        =   100
               Top             =   780
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DataCombo6 
               Height          =   315
               Left            =   -68385
               TabIndex        =   101
               Top             =   1155
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DataCombo7 
               Height          =   315
               Left            =   -72840
               TabIndex        =   102
               Top             =   420
               Width           =   5805
               _ExtentX        =   10239
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DataCombo8 
               Height          =   315
               Left            =   -73890
               TabIndex        =   103
               Top             =   420
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   -67275
               TabIndex        =   109
               Top             =   1215
               Width           =   285
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   -68685
               TabIndex        =   108
               Top             =   1215
               Width           =   210
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   -74730
               TabIndex        =   107
               Top             =   450
               Width           =   810
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   -70200
               TabIndex        =   106
               Top             =   825
               Width           =   405
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   -74625
               TabIndex        =   105
               Top             =   1215
               Width           =   705
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Complemento"
               Height          =   195
               Left            =   -74880
               TabIndex        =   104
               Top             =   825
               Width           =   960
            End
            Begin VB.Label lblstrComplemento 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Complemento"
               Height          =   195
               Left            =   120
               TabIndex        =   73
               Top             =   832
               Width           =   960
            End
            Begin VB.Label lblintCepC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   -67275
               TabIndex        =   40
               Top             =   1215
               Width           =   285
            End
            Begin VB.Label lblintUFC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   -68685
               TabIndex        =   39
               Top             =   1215
               Width           =   210
            End
            Begin VB.Label lblstrComplementoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Complemento"
               Height          =   195
               Left            =   -74880
               TabIndex        =   38
               Top             =   832
               Width           =   960
            End
            Begin VB.Label lblintLogradouroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   -74730
               TabIndex        =   37
               Top             =   450
               Width           =   810
            End
            Begin VB.Label lblintBairroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   -70200
               TabIndex        =   36
               Top             =   832
               Width           =   405
            End
            Begin VB.Label lblintMunicipioC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   -74625
               TabIndex        =   35
               Top             =   1215
               Width           =   705
            End
            Begin VB.Label lblintMunicipio 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   375
               TabIndex        =   34
               Top             =   1215
               Width           =   705
            End
            Begin VB.Label lblintBairro 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   4800
               TabIndex        =   33
               Top             =   832
               Width           =   405
            End
            Begin VB.Label lblintLogradouro 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   270
               TabIndex        =   32
               Top             =   450
               Width           =   810
            End
            Begin VB.Label lblintUF 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   6315
               TabIndex        =   31
               Top             =   1215
               Width           =   210
            End
            Begin VB.Label lblintCep 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   7725
               TabIndex        =   30
               Top             =   1215
               Width           =   285
            End
         End
         Begin MSComctlLib.ListView lvw_TipoComunicacao 
            Height          =   1470
            Left            =   -74910
            TabIndex        =   41
            Top             =   780
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   2593
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Descrição"
               Object.Width           =   52917
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Contas 
            Height          =   900
            Left            =   -74910
            TabIndex        =   42
            Top             =   1440
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   1588
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Descrição"
               Object.Width           =   52917
            EndProperty
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Historico 
            Height          =   1905
            Left            =   -74940
            TabIndex        =   43
            Top             =   420
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3360
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PKID"
            Columns(0).DataField=   "PKId"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "strCodigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Data / Hora"
            Columns(2).DataField=   "dtmDataHora"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo de Transação"
            Columns(3).DataField=   "strTransacao"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nome do Sistema"
            Columns(4).DataField=   "strNomeSistema"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Valor"
            Columns(5).DataField=   "dblValor"
            Columns(5).NumberFormat=   "Standard"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1984"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1984"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1905"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2619"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2540"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2).AllowSizing=0"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=4842"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4763"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3).AllowSizing=0"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=4604"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=4524"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4).AllowSizing=0"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2064"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1984"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5).AllowSizing=0"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
         Begin VB.Label lbl_dtmDebito 
            AutoSize        =   -1  'True
            Caption         =   "Data Início Débito"
            Height          =   195
            Left            =   -70860
            TabIndex        =   82
            Top             =   1140
            Width           =   1305
         End
         Begin VB.Label lblstrIdentidade 
            AutoSize        =   -1  'True
            Caption         =   "Identidade"
            Height          =   195
            Left            =   6765
            TabIndex        =   55
            Top             =   810
            Width           =   750
         End
         Begin VB.Label lblstrTituloEleitoral 
            AutoSize        =   -1  'True
            Caption         =   "Título Eleitoral"
            Height          =   195
            Left            =   315
            TabIndex        =   54
            Top             =   1170
            Width           =   1020
         End
         Begin VB.Label lbldtmDataNascimento 
            AutoSize        =   -1  'True
            Caption         =   "Nascimento"
            Height          =   195
            Left            =   3885
            TabIndex        =   53
            Top             =   1170
            Width           =   840
         End
         Begin VB.Label lblstrCarteiraTrabalho 
            AutoSize        =   -1  'True
            Caption         =   "Carteira de Trabalho"
            Height          =   195
            Left            =   3255
            TabIndex        =   52
            Top             =   810
            Width           =   1440
         End
         Begin VB.Label lblstrInscricaoEstadual 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Estadual"
            Height          =   195
            Left            =   90
            TabIndex        =   51
            Top             =   810
            Width           =   1305
         End
         Begin VB.Label lblstrNomeFantasia 
            AutoSize        =   -1  'True
            Caption         =   "Nome Fantasia"
            Height          =   195
            Left            =   330
            TabIndex        =   50
            Top             =   450
            Width           =   1065
         End
         Begin VB.Label lbl_TipoComunicacao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   -74910
            TabIndex        =   49
            Top             =   480
            Width           =   315
         End
         Begin VB.Label lbl_DescricaoConteudo 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   -70455
            TabIndex        =   48
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lbl_Conta 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   -74715
            TabIndex        =   47
            Top             =   1140
            Width           =   420
         End
         Begin VB.Label lbl_Banco 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   -74760
            TabIndex        =   46
            Top             =   450
            Width           =   465
         End
         Begin VB.Label lbl_Agencia 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Agência"
            Height          =   195
            Left            =   -74880
            TabIndex        =   45
            Top             =   780
            Width           =   585
         End
         Begin VB.Label lbl_DV 
            AutoSize        =   -1  'True
            Caption         =   "DV"
            Height          =   195
            Left            =   -72030
            TabIndex        =   44
            Top             =   1140
            Width           =   225
         End
      End
      Begin MSMask.MaskEdBox mskstrCNPJCPF 
         Height          =   285
         Left            =   3540
         TabIndex        =   56
         Top             =   390
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12632256
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2415
         Left            =   60
         TabIndex        =   70
         Top             =   4140
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4260
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nome"
         Columns(1).DataField=   "strNome"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "CPF / CNPJ"
         Columns(2).DataField=   "strCNPJCPF"
         Columns(2).NumberFormat=   "FormatText Event"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1984"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=9869"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=9790"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4128"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4048"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Named:id=33:Normal"
         _StyleDefs(43)  =   ":id=33,.parent=0"
         _StyleDefs(44)  =   "Named:id=34:Heading"
         _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(46)  =   ":id=34,.wraptext=-1"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lbldtmDataCadastro 
         AutoSize        =   -1  'True
         Caption         =   "Cadastro"
         Height          =   195
         Left            =   5430
         TabIndex        =   68
         Top             =   420
         Width           =   630
      End
      Begin VB.Label lblstrCNPJCPF 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF"
         Height          =   195
         Left            =   2700
         TabIndex        =   66
         Top             =   420
         Width           =   780
      End
      Begin VB.Label lblstrNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   450
         TabIndex        =   65
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lbl_Codigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   375
         TabIndex        =   64
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.Menu mnu_TipoComunicacao 
      Caption         =   "mnuTipoComunicacao"
      Visible         =   0   'False
      Begin VB.Menu mnu_Deletar 
         Caption         =   "Deletar"
      End
      Begin VB.Menu mnu_Traco 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Lista 
         Caption         =   "Lista"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmCadContribuinteN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'    Dim mblnAlterando            As Boolean
'    Dim mobjAux                  As Object
'    Dim oList                    As Object
'    Dim mblnClickOk              As Boolean
'    Dim blnNaturezaJuridicaClick As Boolean
'
'    Dim X                        As XArrayDB 'Grid Contribuintes
'    Dim mblnSelecionou           As Boolean
'    Dim mblnPrimeiraVez          As Boolean
'
'Private Sub cbointBairro_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub cbointLogradouro_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub cbointMunicipio_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub cbointMunicipioC_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'
'Private Sub cbointUFC_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub cbointUFD_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub chk_ContaPublica_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub chk_DebitoAutomatico_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub chkblnResidenteNoMunicipio_Click()
'    If chkblnResidenteNoMunicipio.Value = 0 Then
'        tab_3DCorrespondencia.TabEnabled(0) = False
'        tab_3DCorrespondencia.Tab = 1
'
'        cbointLogradouro.BoundText = ""
'        cbointBairro.BoundText = ""
'        cbointMunicipio.BoundText = ""
'        cbointUF.BoundText = ""
'        txtintNumero = ""
'        txtstrComplemento = ""
'        txtintCep = ""
'    Else
'        tab_3DCorrespondencia.TabEnabled(0) = True
'        tab_3DCorrespondencia.Tab = 0
'    End If
'End Sub
'
'Private Sub chkblnResidenteNoMunicipio_KeyPress(KeyAscii As Integer)
'CaracterValido KeyAscii, "A", chkblnResidenteNoMunicipio
'End Sub
'
'Private Sub cmd_ContasBancarias_Click()
'    If Trim(txtPKId) = "" Then
'        ExibeMensagem "O contribuinte tem que ser salvo."
'        Exit Sub
'    End If
'    frmCadContasBancarias.Show
'    frmCadContasBancarias.dbcintContribuinte.BoundText = txtPKId
'    TrocaCorObjeto frmCadContasBancarias.dbcintContribuinte, True
'    frmCadContasBancarias.cmd_Contribuinte.Enabled = False
'    frmCadContasBancarias.Visible = True
'    frmCadContasBancarias.RotinaAuxiliarClickCombo
'End Sub
'
'Private Sub dbcintTipoLogradouro_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", dbcintTipoLogradouro
'End Sub
'
'Private Sub dbcintTipoLogradouroD_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", dbcintTipoLogradouroD
'End Sub
'
'Private Sub dbcintTituloLogradouro_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", dbcintTituloLogradouro
'End Sub
'
'Private Sub dbcintTituloLogradouroD_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", dbcintTituloLogradouroD
'End Sub
'
'Private Sub cmd_TipoLogradouro_Click()
''    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouro
'End Sub
'
'Private Sub cmd_TituloLogradouro_Click()
''    CarregaForm frmCadTituloLogradouro, dbcintTituloLogradouro
'End Sub
'
'Private Sub cmd_TipoLogradouroD_Click()
''    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouroD
'End Sub
'
'Private Sub cmd_TituloLogradouroD_Click()
''    CarregaForm frmCadTituloLogradouro, dbcintTituloLogradouroD
'End Sub
'
'Private Sub cmd_Logradouro_Click()
'    ChamaFormCadastro frmCadLogradouro, cbointLogradouro
'End Sub
'
'Private Sub cmd_Down_Click()
'    If lvw_TipoComunicacao.ListItems.Count <> 0 Then
'        MoveItemNoListView lvw_TipoComunicacao, True
'    End If
'End Sub
'
'Private Sub cmd_Up_Click()
'    If lvw_TipoComunicacao.ListItems.Count <> 0 Then
'        MoveItemNoListView lvw_TipoComunicacao, False
'    End If
'End Sub
'
'Private Sub cmd_MunicipioC_Click()
'    ChamaFormCadastro frmCadCidade, cbointMunicipioC
'End Sub
'
'Private Sub txtdtmDataCadastro_GotFocus()
'    MarcaCampo txtdtmDataCadastro
'End Sub
'
'Private Sub txtdtmDataCadastro_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "D", txtdtmDataCadastro
'End Sub
'
'Private Sub Form_Activate()
'    gintCodSeguranca = 15
'    VirificaGradeListView Me
'    If mblnSelecionou Then
'        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
'    Else
'        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
'    End If
'    If mobjAux Is Nothing Then
'        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
'    Else
'        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
'    End If
'End Sub
'
'Private Sub Form_Deactivate()
'    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
'End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
'  If KeyCode = vbKeyF1 Then
'     Call_HtmlHelp Me.HelpContextID
'  End If
'End Sub
'
'Private Sub Form_Load()
'
'    LeDaTabelaParaObj gstrBairro, cbointBairro
'    LeDaTabelaParaObj gstrCidade, cbointMunicipio
'    LeDaTabelaParaObj gstrLogradouro, cbointLogradouro, gstrQueryLogradouro
'    LeDaTabelaParaObj gstrUF, cbointUF, gstrQueryUF
'    LeDaTabelaParaObj gstrCidade, cbointMunicipioC
'    LeDaTabelaParaObj gstrUF, cbointUFC, gstrQueryUF
''    LeDaTabelaParaObj gstrUF, cbointUFD, gstrQueryUF
''    LeDaTabelaParaObj gstrTipoLogradouro, dbcintTipoLogradouro, "PKID, strSigla"
''    LeDaTabelaParaObj gstrTituloLogradouro, dbcintTituloLogradouro
''    LeDaTabelaParaObj gstrTipoLogradouro, dbcintTipoLogradouroD, "PKID, strSigla"
''    LeDaTabelaParaObj gstrTituloLogradouro, dbcintTituloLogradouroD
'    LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strQuerryGridContribuinte
'
'    PreencheMenuPopup
'
'    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
''    txtintCodigoLogradouro.Enabled = False
''    TrocaCorObjeto txtintCodigoLogradouro, True
''    txtintCodigoLogradouroD.Enabled = False
''    TrocaCorObjeto txtintCodigoLogradouroD, True
'    VerificaObjParaAplicar mobjAux
'    NovoContribuinte
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
'    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo, gstrSalvar
'    mblnSelecionou = False
'    mblnPrimeiraVez = False
'End Sub
'
'Private Sub lvw_Contas_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub lvw_TipoComunicacao_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    With lvw_TipoComunicacao
'        lbl_TipoComunicacao = .SelectedItem.Text
'        txt_Conteudo = .SelectedItem.SubItems(1)
'        txt_DescricaoConteudo = .SelectedItem.SubItems(2)
'    End With
'End Sub
'
'Private Sub lvw_TipoComunicacao_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub lvw_TipoComunicacao_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then
'        PopupMenu mnu_TipoComunicacao
'    End If
'End Sub
'
'Private Sub mnu_Deletar_Click()
'    With lvw_TipoComunicacao
'        If .ListItems.Count = 0 Then Exit Sub
'        If .SelectedItem.Selected = False Then Exit Sub
'        .ListItems.Remove .SelectedItem.Index
'        txt_Conteudo = ""
'        txt_DescricaoConteudo = ""
'        lbl_TipoComunicacao = "Tipo"
'    End With
'End Sub
'
'Private Sub mnu_Lista_Click(Index As Integer)
'    With lvw_TipoComunicacao
'        .Sorted = False
'        Set oList = .ListItems.Add(, , mnu_Lista(Index).Caption)
'        oList.SubItems(1) = ""
'        oList.SubItems(2) = ""
'        oList.Tag = mnu_Lista(Index).Tag
'        .ListItems(.ListItems.Count).Selected = True
'        .ListItems(.ListItems.Count).EnsureVisible
'        lvw_TipoComunicacao_ItemClick .SelectedItem
'        txt_Conteudo.SetFocus
'    End With
'End Sub
'
'Private Sub mskstrCNPJCPF_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", mskstrCNPJCPF
'End Sub
'
'Private Sub optbytNaturezaJuridica_Click(Index As Integer)
'    Select Case Index
'        Case 0  'Física
'            chkblnResidenteNoMunicipio.Caption = "Residente no município"
'            HabilitaDesabilitaObjeto txtstrIdentidade, True
'            HabilitaDesabilitaObjeto txtstrTituloEleitoral, True
'            HabilitaDesabilitaObjeto txtdtmDataNascimento, True
'            HabilitaDesabilitaObjeto txtstrCarteiraTrabalho, True
'            HabilitaDesabilitaObjeto txtstrInscricaoEstadual, False
'            HabilitaDesabilitaObjeto txtstrNomeFantasia, False
'            If mblnAlterando = False Then
'                HabilitaDesabilitaObjeto mskstrCNPJCPF, True
'            End If
'            If blnNaturezaJuridicaClick Then
'                mskstrCNPJCPF.Mask = "###\.###\.###\-##"
'            End If
'
'        Case 1, 2, 3    'Jurídica, SC , Outros
'            chkblnResidenteNoMunicipio.Caption = "Estabelecido no município"
'            HabilitaDesabilitaObjeto txtstrIdentidade, False
'            HabilitaDesabilitaObjeto txtstrTituloEleitoral, False
'            HabilitaDesabilitaObjeto txtdtmDataNascimento, False
'            HabilitaDesabilitaObjeto txtstrCarteiraTrabalho, False
'            HabilitaDesabilitaObjeto txtstrInscricaoEstadual, True
'            HabilitaDesabilitaObjeto txtstrNomeFantasia, True
'            If mblnAlterando = False Then
'                HabilitaDesabilitaObjeto mskstrCNPJCPF, True
'            End If
'            If blnNaturezaJuridicaClick Then
'                mskstrCNPJCPF.Mask = "##\.###\.###\/####\-##"
'            End If
'
'    End Select
'    HabilitaDesabilitaObjeto txtstrNome, True
'    HabilitaDesabilitaObjeto chkblnResidenteNoMunicipio, True
'    HabilitaDesabilitaObjeto txtdtmDataCadastro, True
'    tab_3DDadosGerais.TabEnabled(1) = True
'    tab_3DDadosGerais.TabEnabled(2) = True
'    tab_3DDadosGerais.TabEnabled(3) = True
'    tab_3DDadosGerais.TabEnabled(4) = True
'
'    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
'
'    blnNaturezaJuridicaClick = False
'End Sub
'
'Private Sub optbytNaturezaJuridica_KeyPress(Index As Integer, KeyAscii As Integer)
'CaracterValido KeyAscii, "A", optbytNaturezaJuridica(Index)
'End Sub
'
'Private Sub optbytNaturezaJuridica_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    blnNaturezaJuridicaClick = True
'End Sub
'
'Private Sub optbytNaturezaJuridica_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    blnNaturezaJuridicaClick = False
'End Sub
'
'Private Sub tab_3DCorrespondencia_Click(PreviousTab As Integer)
'    Select Case tab_3DCorrespondencia.Tab
'        Case 1
'            'If mblnAlterando = False And tab_3DCorrespondencia.TabEnabled(0) = True And txtstrLogradouroC.Text <> cbointLogradouro.Text Then
'                If MsgBox("O endereço de correspondência é o mesmo que o endereço residencial?", vbYesNo + vbQuestion) = vbYes Then
'
''                    dbcintTipoLogradouro.Enabled = False
''                    TrocaCorObjeto dbcintTipoLogradouro, True
'                    'dbcintTituloLogradouro.Enabled = False
'                    'TrocaCorObjeto dbcintTituloLogradouro, True
''                    cmd_TipoLogradouro.Enabled = False
''                    TrocaCorObjeto cmd_TipoLogradouro, True
''                    cmd_TituloLogradouro.Enabled = False
''                    TrocaCorObjeto cmd_TituloLogradouro, True
'
'
''                    txtintCodigoLogradouro = cbointLogradouro.BoundText
''                    txtstrLogradouroC = cbointLogradouro.Text
''                    txtstrBairroC = cbointBairro.Text
''                    cbointMunicipioC.BoundText = cbointMunicipio.BoundText
''                    cbointUFC.BoundText = cbointUF.BoundText
''                    txtintNumeroC = txtintNumero
''                    txtstrComplementoC = txtstrComplemento
''                    txtintCepC = txtintCep
'                End If
'            'End If
'
'        Case 2
''            If mblnAlterando = False And tab_3DCorrespondencia.TabEnabled(0) = True And txtstrLogradouroD.Text <> cbointLogradouro.Text Then
'                If MsgBox("O endereço do domicílio fiscal é o mesmo que o endereço residencial?", vbYesNo + vbQuestion) = vbYes Then
'
''                    dbcintTipoLogradouroD.Enabled = False
''                    TrocaCorObjeto dbcintTipoLogradouroD, True
''                    dbcintTituloLogradouroD.Enabled = False
''                    TrocaCorObjeto dbcintTituloLogradouroD, True
''                    cmd_TipoLogradouroD.Enabled = False
''                    TrocaCorObjeto cmd_TipoLogradouroD, True
''                    cmd_TituloLogradouroD.Enabled = False
''                    TrocaCorObjeto cmd_TituloLogradouroD, True
''
''                    txtintCodigoLogradouroD = cbointLogradouro.BoundText
''                    txtstrLogradouroD = cbointLogradouro.Text
''                    txtstrBairroD = cbointBairro.Text
''                    txtstrMunicipioD = cbointMunicipio.Text
''                    cbointUFD.BoundText = cbointUF.BoundText
''                    txtintNumeroD = txtintNumero
''                    txtstrComplementoD = txtstrComplemento
''                    txtintCepD = txtintCep
'                'End If
'            End If
'
'            If MDIMenu.Tag = "MATERIAL" Then
''                If mblnAlterando = False And tab_3DCorrespondencia.TabEnabled(0) = False And txtstrLogradouroD.Text <> txtstrLogradouroC.Text Then
''                    If MsgBox("O endereço do domicílio fiscal é o mesmo que o endereço de correspondência?", vbYesNo + vbQuestion) = vbYes Then
''
''                        dbcintTipoLogradouroD.Enabled = False
''                        TrocaCorObjeto dbcintTipoLogradouroD, True
''                        dbcintTituloLogradouroD.Enabled = False
''                        TrocaCorObjeto dbcintTituloLogradouroD, True
''                        cmd_TipoLogradouroD.Enabled = False
''                        TrocaCorObjeto cmd_TipoLogradouroD, True
''                        cmd_TituloLogradouroD.Enabled = False
''                        TrocaCorObjeto cmd_TituloLogradouroD, True
''
''                        dbcintTipoLogradouroD.BoundText = dbcintTipoLogradouro.BoundText
''                        dbcintTituloLogradouroD.BoundText = dbcintTituloLogradouro.BoundText
''                        txtstrLogradouroD.Text = txtstrLogradouroC.Text
''                        txtstrBairroD.Text = txtstrBairroC.Text
''                        txtstrMunicipioD.Text = cbointMunicipioC.Text
''                        cbointUFD.BoundText = cbointUFC.BoundText
''                        txtintNumeroD.Text = txtintNumeroC.Text
''                        txtstrComplementoD.Text = txtstrComplementoC.Text
''                        txtintCepD.Text = txtintCepC.Text
''                    End If
''                End If
'            End If
'    End Select
'End Sub
'
'Private Sub tab_3DDadosGerais_Click(PreviousTab As Integer)
'    Select Case tab_3DDadosGerais.Tab
'        Case 1
'            If chkblnResidenteNoMunicipio.Value = 1 Then
'                cbointMunicipio.BoundText = gintMunicipioEmpresa
'                cbointMunicipioC.BoundText = gintMunicipioEmpresa
'                TrocaCorObjeto cbointMunicipio, True
'                TrocaCorObjeto cbointMunicipioC, True
'                cmd_Municipio.Enabled = False
'
'                cbointUF.BoundText = gintUFEmpresa
'                cbointUFC.BoundText = gintUFEmpresa
'                TrocaCorObjeto cbointUF, True
'                TrocaCorObjeto cbointUFC, True
'                cmd_MunicipioC.Enabled = False
'            Else
'                TrocaCorObjeto cbointMunicipio, False
'                TrocaCorObjeto cbointMunicipioC, False
'                TrocaCorObjeto cbointUF, False
'                TrocaCorObjeto cbointUFC, False
'                cmd_Municipio.Enabled = True
'                cmd_MunicipioC.Enabled = True
'            End If
'    End Select
'End Sub
'
'Private Sub tdb_Historico_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub tdb_Lista_Click()
'    mblnPrimeiraVez = True
'    With tdb_Lista
'       If Not .BOF And Not .EOF Then
'            If .Bookmark = 0 Then
''                tdb_Lista_RowColChange 0, 0
'            End If
'       End If
'    End With
'
'End Sub
'
'Private Sub tdb_Lista_DblClick()
'    MantemForm gstrAplicar
'End Sub
'
'Private Sub tdb_Lista_FilterChange()
'    PreencheGridContribuinte
'End Sub
'
'Sub PreencheGridContribuinte()
'    Dim col    As TrueOleDBGrid70.Column
'    Dim c      As Integer
'    Dim tmp    As String
'    Dim n      As Integer
'    Dim strAux As String
'    Dim strSql As String
'    Dim ADOTemp As ADODB.Recordset
'
'    c = tdb_Lista.col
'    tdb_Lista.HoldFields
'
'    tmp = ""
'
'    strSql = ""
'    strSql = strSql & " SELECT PKID, strNome, strCNPJCPF, bytNaturezaJuridica FROM " & gstrContribuinte
'
'    strAux = strSql
'    strAux = strAux & " ORDER BY strNome"
'
'    strSql = strSql & " WHERE "
'
'
'    For Each col In tdb_Lista.Columns
'        If Trim(col.FilterText) <> "" Then
'            n = n + 1
'
'            If tmp <> "" Then
'                tmp = tmp & " AND "
'            End If
'
'            Select Case UCase(col.DataField)
'                Case "PKID"
'                    tmp = tmp & col.DataField & " = " & col.FilterText
'                Case "STRNOME"
'                    tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "%'"
'                Case "STRCNPJCPF"
'                    tmp = tmp & col.DataField & " LIKE '" & gstrValorSemMascara(col.FilterText) & "%'"
'            End Select
'        End If
'    Next
'
'    If tmp <> "" Then
'        strSql = strSql & tmp
'        strSql = strSql & " ORDER BY strNome"
'        Set gobjBanco = New clsBanco
'        If gobjBanco.CriaADO(strSql, 5, ADOTemp) Then
'            If Not ADOTemp.EOF Then
'                'MontaArray strSql
'                LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strSql
'            Else
'                'MontaArray
'                LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strQuerryGridContribuinte
'            End If
'        End If
'    Else
'        'MontaArray
'        LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strQuerryGridContribuinte
'    End If
'
'    tdb_Lista.col = c
'    tdb_Lista.EditActive = True
'    tdb_Lista.CurrentCellModified = True
'
'    NovoContribuinte
'End Sub
'
'Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyUp, vbKeyPageUp, vbKeyPageDown, vbKeyDown
'        mblnClickOk = True
'    Case Else
'        mblnClickOk = False
'    End Select
'End Sub
'
'Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
'    Select Case tdb_Lista.col
'        Case 0
'            CaracterValido KeyAscii, "N", tdb_Lista
'        Case Else
'            CaracterValido KeyAscii, "A", tdb_Lista
'    End Select
'End Sub
'
'Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    mblnClickOk = True
'End Sub
'
'Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'Dim i As Integer
'    On Error GoTo err_tdb_Lista_RowColChange
'    With tdb_Lista
'        If mblnClickOk Then
'            If Not .EOF And Not .BOF Then
'                If mblnPrimeiraVez Then
'                    If Trim(tdb_Lista.Columns("PKId").Value) = "" Then
'                        Exit Sub
'                    End If
'                    mblnClickOk = False
'                    mblnAlterando = True
'            '        mskstrCNPJCPF.Mask = ""
'                    HabilitaDesabilitaObjeto mskstrCNPJCPF
'
'                    txtPKId = Val(tdb_Lista.Columns("PKId").Value)
'                    LeDaTabelaParaObj gstrContribuinte, Me
'                    gCorLinhaSelecionada tdb_Lista
'                    If mobjAux Is Nothing Then
'                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
'                    Else
'                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
'                    End If
'
'                    CarregaTipoComunicacao txtPKId
'                    CarregaContasBancarias txtPKId
'                    CarregaHistorico txtPKId
'
'                    txt_Codigo = Format(txtPKId, "00000000")
'                    tab_3DDadosGerais.TabEnabled(1) = True
'                    tab_3DDadosGerais.TabEnabled(2) = True
'                    tab_3DDadosGerais.TabEnabled(3) = True
'                    tab_3DDadosGerais.TabEnabled(4) = True
'
'                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar, gstrDeletar
'
'                    fra_bytNaturezaJuridica.Enabled = False
'
'                    For i = 0 To 3
'                        If optbytNaturezaJuridica(i).Value = True Then
'                            Select Case i
'                                Case 0
'                                    mskstrCNPJCPF.Mask = "###\.###\.###\-##"
'                                Case 1, 2, 3
'                                    mskstrCNPJCPF.Mask = "##\.###\.###\/####\-##"
'                            End Select
'                            Exit For
'                        End If
'                    Next
'
'                    tab_3DDadosGerais.Tab = 0
'                    mskstrCNPJCPF = .Columns("strCNPJCPF")
'            '        txtstrNome.SetFocus
'                End If
'            End If
'        End If
'    End With
'err_tdb_Lista_RowColChange:
'End Sub
'
'Private Function blnDadosOk() As Boolean
'    On Error GoTo err_blnDadosOK
'    If Trim(txtstrNome) = "" Then
'        ExibeMensagem "O nome tem que ser digitado."
'        txtstrNome.SetFocus
'        Exit Function
'    End If
'
'    If mskstrCNPJCPF.ClipText <> "" Then
'        If optbytNaturezaJuridica(0).Value = True Then
'            If Not gblnCPFOk(mskstrCNPJCPF) Then
'                ExibeMensagem "CPF inválido."
'                mskstrCNPJCPF.SetFocus
'                Exit Function
'            End If
'        Else
'            If Not gblnCGCOk(mskstrCNPJCPF) Then
'                ExibeMensagem "CNPJ / CPF inválido."
'                mskstrCNPJCPF.SetFocus
'                Exit Function
'            End If
'        End If
'    End If
'
'    If blnDadosDuplicados Then
'        Exit Function
'    End If
'
'    If tab_3DCorrespondencia.TabEnabled(0) = True Then
'        If cbointLogradouro.BoundText = "" Then
'            ExibeMensagem "O logradouro residencial tem que ser informado."
'            cbointLogradouro.SetFocus
'            Exit Function
'        End If
'        If Not gblnCepValido(txtintCep, cbointLogradouro) Then
'            ExibeMensagem "O cep residencial é inválido para o logradouro informado."
'            txtintCep.SetFocus
'            Exit Function
'        End If
'    End If
'    If chkblnResidenteNoMunicipio.Value = 0 Then
'        If cbointMunicipioC.BoundText = "" Then
'            ExibeMensagem "O município do endereço de correspondência tem que ser informado."
'            cbointMunicipioC.SetFocus
'            Exit Function
'        End If
'    End If
'    If Not gblnCepValido(txtintCepC, , cbointMunicipioC) Then
'        ExibeMensagem "O cep do endereço de correspondência é inválido para o município informado."
'        txtintCepC.SetFocus
'        Exit Function
'    End If
'    blnDadosOk = True
'err_blnDadosOK:
'End Function
'
'Function blnDadosDuplicados() As Boolean
'    Dim adoResultado As ADODB.Recordset
'    Dim strSql       As String
'
'    strSql = ""
'    strSql = strSql & "SELECT strCNPJCPF "
'    strSql = strSql & "FROM " & gstrContribuinte & " "
'    strSql = strSql & "WHERE strCNPJCPF = '" & gstrValorSemMascara(mskstrCNPJCPF) & "' "
'    If mblnAlterando Then
'        strSql = strSql & "AND PKID <> " & txtPKId
'    End If
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        If Not adoResultado.EOF Then
'            ExibeMensagem "CNPJ / CPF já cadastrado para outro contribuinte."
'        Else
'            Exit Function
'        End If
'    End If
'
'    blnDadosDuplicados = True
'End Function
'
'Sub NovoContribuinte()
'    Dim i As Integer
'
''    mblnAlterando = False
''    lvw_TipoComunicacao.ListItems.Clear
''    lvw_Contas.ListItems.Clear
''    txt_Codigo = Format(glngPegaProximaChave(gstrContribuinte, "PKId"), "00000000")
''    txt_Conteudo = ""
''    txt_DescricaoConteudo = ""
''    lbl_TipoComunicacao = "Tipo"
''
''    txt_CodBanco = ""
''    txt_CodAgencia = ""
''    txt_Banco = ""
''    txt_Agencia = ""
''    txt_Conta = ""
''    txt_DigitoVerificador = ""
''    chk_ContaPublica.Value = 0
''    chk_DebitoAutomatico.Value = 0
''    txt_dtmDebito = ""
''
''    HabilitaDesabilitaObjeto txtstrNome, False
''    HabilitaDesabilitaObjeto mskstrCNPJCPF, False
''    HabilitaDesabilitaObjeto chkblnResidenteNoMunicipio, False
''    HabilitaDesabilitaObjeto txtdtmDataCadastro, False
''    HabilitaDesabilitaObjeto txtstrNomeFantasia, False
''    HabilitaDesabilitaObjeto txtstrInscricaoEstadual, False
''    HabilitaDesabilitaObjeto txtstrIdentidade, False
''    HabilitaDesabilitaObjeto txtstrTituloEleitoral, False
''    HabilitaDesabilitaObjeto txtdtmDataNascimento, False
''    HabilitaDesabilitaObjeto txtstrCarteiraTrabalho, False
''
''    fra_bytNaturezaJuridica.Enabled = True
''    chkblnResidenteNoMunicipio.Value = 0
''
''    tab_3DCorrespondencia.Tab = 1
''    tab_3DCorrespondencia.TabEnabled(0) = False
''    tab_3DDadosGerais.Tab = 0
''    tab_3DDadosGerais.TabEnabled(1) = False
''    tab_3DDadosGerais.TabEnabled(2) = False
''    tab_3DDadosGerais.TabEnabled(3) = False
''    tab_3DDadosGerais.TabEnabled(4) = False
''
''    For i = 0 To 3
''        optbytNaturezaJuridica(i).Value = False
''    Next
'
'    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrSalvar, gstrDeletar
'End Sub
'
'Private Sub tlb_TipoComunicacao_ButtonClick(ByVal Button As MSComctlLib.Button)
'    On Error Resume Next
'    Select Case UCase(Button.Key)
'        Case gstrSalvar
'            If lvw_TipoComunicacao.ListItems.Count = 0 Then
'                Exit Sub
'            End If
'            If lvw_TipoComunicacao.SelectedItem.Selected = False Then
'                Exit Sub
'            End If
'            lvw_TipoComunicacao.SelectedItem.Selected = False
'        Case gstrNovo
'            mnu_Deletar.Visible = False
'            mnu_Traco.Visible = False
'            PopupMenu mnu_TipoComunicacao
'            mnu_Deletar.Visible = True
'            mnu_Traco.Visible = True
'        Case gstrDeletar
'            If lvw_TipoComunicacao.ListItems.Count = 0 Then Exit Sub
'            If lvw_TipoComunicacao.SelectedItem.Selected = False Then Exit Sub
'            lvw_TipoComunicacao.ListItems.Remove lvw_TipoComunicacao.SelectedItem.Index
'    End Select
'    txt_Conteudo = ""
'    txt_DescricaoConteudo = ""
'End Sub
'
'Private Sub txt_Agencia_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", txt_Agencia
'End Sub
'
'Private Sub txt_Banco_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", txt_Banco
'End Sub
'
'Private Sub txt_CodAgencia_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txt_CodBanco_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txt_Codigo
'End Sub
'
'Private Sub txt_Codigo1_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txt_Conta_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txt_Conta
'End Sub
'
'Private Sub txt_Conteudo_Change()
'    If lvw_TipoComunicacao.ListItems.Count = 0 Then Exit Sub
'    If lvw_TipoComunicacao.SelectedItem.Selected = False Then Exit Sub
'    lvw_TipoComunicacao.SelectedItem.SubItems(1) = Trim(txt_Conteudo)
'End Sub
'
'Private Sub txt_Conteudo_GotFocus()
'    MarcaCampo txt_Conteudo
'End Sub
'
'Private Sub txt_Conteudo_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txt_Conteudo
'End Sub
'
'Private Sub txt_DescricaoConteudo_Change()
'    If lvw_TipoComunicacao.ListItems.Count = 0 Then Exit Sub
'    If lvw_TipoComunicacao.SelectedItem.Selected = False Then Exit Sub
'    lvw_TipoComunicacao.SelectedItem.SubItems(2) = Trim(txt_DescricaoConteudo)
'End Sub
'
'Private Sub txt_DescricaoConteudo_GotFocus()
'    MarcaCampo txt_DescricaoConteudo
'End Sub
'
'Private Sub txt_DescricaoConteudo_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txt_DescricaoConteudo
'End Sub
'
'Private Sub txt_DigitoVerificador_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txt_DigitoVerificador
'End Sub
'
'Private Sub txt_dtmDebito_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "D", txt_dtmDebito
'End Sub
'
'Private Sub txt_Nome1_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txtdtmDataCadastro_LostFocus()
'    txtdtmDataCadastro = gstrDataFormatada(txtdtmDataCadastro)
'End Sub
'
'Private Sub txtintCEP_GotFocus()
'    MarcaCampo txtintCep
'End Sub
'
'Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "E", txtintCep
'End Sub
'
'Private Sub txtintCEP_LostFocus()
'    txtintCep = gstrCEPFormatado(txtintCep)
'End Sub
'
'Private Sub txtintCepC_GotFocus()
'    MarcaCampo txtintCepC
'End Sub
'
'Private Sub txtintCepC_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "E", txtintCepC
'End Sub
'
'Private Sub txtintCepC_LostFocus()
'    txtintCepC = gstrCEPFormatado(txtintCepC)
'End Sub
'
'Private Sub mskstrCNPJCPF_GotFocus()
'    MarcaCampo mskstrCNPJCPF
'End Sub
'
'Sub MontaColumnHeaders()
'    With lvw_TipoComunicacao
'        .ColumnHeaders.Clear
'        .ColumnHeaders.Add 1, , "Tipo", 2000
'        .ColumnHeaders.Add 2, , "Conteúdo", 3000
'        .ColumnHeaders.Add 3, , "Descrição", 3210
'    End With
'    With lvw_Contas
'        .ColumnHeaders.Clear
'        .ColumnHeaders.Add 1, , "CodBanco", 0
'        .ColumnHeaders.Add 2, , "Banco", 2700
'        .ColumnHeaders.Add 3, , "codAgência", 0
'        .ColumnHeaders.Add 4, , "Agência", 2000
'        .ColumnHeaders.Add 5, , "Conta", 1500
'        .ColumnHeaders.Add 6, , "DV", 500
'        .ColumnHeaders.Add 7, , "Pública", 800
'        .ColumnHeaders.Add 8, , "Débito Automático", 1500
'        .ColumnHeaders.Add 9, , "Data Início Débito", 1500
'    End With
''    With lvw_Historico
''        .ColumnHeaders.Clear
''        .ColumnHeaders.Add 1, , "Código", 1000
''        .ColumnHeaders.Add 2, , "Data / Hora", 2000
''        .ColumnHeaders.Add 3, , "Tipo da transação", 3000
''        .ColumnHeaders.Add 4, , "Valor", 1500
''    End With
'End Sub
'
'Private Sub cbointUF_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "U", cbointUF
'End Sub
'
'Private Sub cmd_Bairro_Click()
'    ChamaFormCadastro frmCadBairro, cbointBairro
'End Sub
'
'Private Sub cmd_Municipio_Click()
'    ChamaFormCadastro frmCadCidade, cbointMunicipio
'End Sub
'
'Private Sub HabilitaDesabilitaObjeto(mobjObjeto As Object, Optional blnFlag As Boolean)
'    If Not blnFlag Then  'Desabilita
'        If TypeOf mobjObjeto Is TextBox Then
'            mobjObjeto.Text = ""
'            mobjObjeto.BackColor = &HC0C0C0
'        ElseIf TypeOf mobjObjeto Is MaskEdBox Then
'            mobjObjeto.Mask = ""
'            mobjObjeto.Text = ""
'            mobjObjeto.BackColor = &HC0C0C0
'        ElseIf TypeOf mobjObjeto Is DTPicker Then
'            mobjObjeto = Format(Date, "dd/mm/yyyy")
'        ElseIf TypeOf mobjObjeto Is CheckBox Then
'            mobjObjeto.Value = 0
'        End If
'        mobjObjeto.Enabled = False
'    Else    'Habilita
'        If TypeOf mobjObjeto Is TextBox Then
''            mobjObjeto.Text = ""
'            mobjObjeto.BackColor = &H80000005
'        ElseIf TypeOf mobjObjeto Is MaskEdBox Then
'            mobjObjeto.Mask = ""
'            mobjObjeto.Text = ""
'            mobjObjeto.BackColor = &H80000005
'        ElseIf TypeOf mobjObjeto Is DTPicker Then
''            mobjObjeto.Date = Format(Date, "dd/mm/yyyy")
'        ElseIf TypeOf mobjObjeto Is CheckBox Then
''            mobjObjeto.Value = 0
'        End If
'        mobjObjeto.Enabled = True
'    End If
'End Sub
'
'Sub CarregaTipoComunicacao(intCodContribuinte As Long)
'
''******************************************************************************************
'' Data: 09/06/2003
'' Alteração: - Deslocamento dos JOINS utilizados na cláusula FROM para a cláusula WHERE.
'' Responsável: Everton Bianchini
''******************************************************************************************
'
'    Dim strSql       As String
'    Dim adoResultado As ADODB.Recordset
'
'    lvw_TipoComunicacao.ListItems.Clear
'    lbl_TipoComunicacao = "Tipo"
'    txt_Conteudo = ""
'    txt_DescricaoConteudo = ""
'
'    strSql = ""
'    strSql = strSql & "Select TP.strDescricao TipoComunicacao, "
'    strSql = strSql & "FC.intTipoDeComunicacao, FC.strDescricao, FC.strConteudo "
'    strSql = strSql & "From " & gstrTipoDeComunicacao & " TP "
''    strSql = strSql & "Left Join " & gstrFormaDeComunicacao & " FC "
'    strSql = strSql & ", " & gstrFormaDeComunicacao & " FC "
''    strSql = strSql & "On TP.PKId = FC.intTipoDeComunicacao "
'    strSql = strSql & "Where FC.intContribuinte = " & intCodContribuinte & " "
'    strSql = strSql & " AND TP.PKId " & strOUTJOracle & strOUTJSQLServer & "= FC.intTipoDeComunicacao "
'    strSql = strSql & "Order By FC.intSequencia"
'
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        With adoResultado
'            Do While Not .EOF
'                Set oList = lvw_TipoComunicacao.ListItems.Add(, , Trim(!TipoComunicacao))
'                oList.SubItems(1) = gstrVerificaCampoNulo(!strConteudo)
'                oList.SubItems(2) = gstrVerificaCampoNulo(!strDescricao)
'                oList.Tag = gstrVerificaCampoNulo(!intTipoDeComunicacao)
'                .MoveNext
'            Loop
'        End With
'    End If
'    If lvw_TipoComunicacao.ListItems.Count <> 0 Then
'        lvw_TipoComunicacao.SelectedItem.Selected = False
'    End If
'End Sub
'
'Sub CarregaContasBancarias(intCodContribuinte As Long)
'    Dim strSql As String
'
'    lvw_Contas.ListItems.Clear
'
'    strSql = ""
'    strSql = strSql & "Select CB.PKId, BO.intBanco, BO.strDescricao AS Banco, AG.intNumero, AG.strDescricao AS Agencia, CB.strConta, CB.strDigitoVerificador, CB.blnContaPublica, blnDebitoAutomatico, dtmDebito "
'    strSql = strSql & "From " & gstrContaBancaria & " CB, "
'    strSql = strSql & gstrBanco & " BO, "
'    strSql = strSql & gstrAgencia & " AG "
'    strSql = strSql & "Where intContribuinte = " & intCodContribuinte & " "
'    strSql = strSql & "AND CB.intAgencia = AG.PKId "
'    strSql = strSql & "AND CB.intBanco = BO.PKId "
'
'    LeDaTabelaParaObj gstrContaBancaria, lvw_Contas, strSql
'
'    If lvw_Contas.ListItems.Count <> 0 Then
'        lvw_Contas.SelectedItem.Selected = False
'    End If
'End Sub
'
'Sub CarregaHistorico(intCodContribuinte As Long)
'    Dim strSql As String
'
'    strSql = ""
'    strSql = strSql & " Select * "
'    strSql = strSql & " FROM " & gstrHistoricoContribuinte
'    strSql = strSql & " WHERE intContribuinte = " & txtPKId.Text
'    strSql = strSql
'
'    LeDaTabelaParaObj gstrHistoricoContribuinte, tdb_Historico, strSql
'
'End Sub
'
'Function blnGravaTipoComunicacao(intCodContribuinte As Integer) As Boolean
'    Dim strSql As String
'    Dim intI   As Integer
'
'    DeletaTipoComunicacao intCodContribuinte
'
'    With lvw_TipoComunicacao
'        For intI = 1 To .ListItems.Count
'            strSql = ""
'            strSql = strSql & "Insert Into " & gstrFormaDeComunicacao & " "
'            strSql = strSql & "(intContribuinte, intTipoDeComunicacao, strConteudo, strDescricao, "
'            strSql = strSql & "intSequencia) Values ("
'            strSql = strSql & intCodContribuinte & ", "
'            strSql = strSql & .ListItems(intI).Tag & ", '"
'            strSql = strSql & .ListItems(intI).SubItems(1) & "', '"
'            strSql = strSql & .ListItems(intI).SubItems(2) & "', "
'            strSql = strSql & intI & ")"
'            Set gobjBanco = New clsBanco
'            gobjBanco.Execute strSql
'        Next
'    End With
'    blnGravaTipoComunicacao = True
'End Function
'
'Sub DeletaTipoComunicacao(intCodContribuinte As Integer)
'    Dim strSql As String
'
'    strSql = ""
'    strSql = strSql & "Delete From " & gstrFormaDeComunicacao & " "
'    strSql = strSql & "Where intContribuinte = " & intCodContribuinte
'
'    Set gobjBanco = New clsBanco
'    gobjBanco.Execute strSql
'End Sub
'
'Sub PreencheMenuPopup()
'    Dim strSql       As String
'    Dim adoResultado As ADODB.Recordset
'    Dim intI         As Integer
'
'    On Error GoTo Err_Handle
'    intI = 0
'
'    strSql = ""
'    strSql = strSql & "Select TP.PKId, TP.strDescricao TipoComunicacao "
'    strSql = strSql & "From " & gstrTipoDeComunicacao & " TP "
'
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        With adoResultado
'            Do While Not .EOF
'                intI = intI + 1
'                Load mnu_Lista(intI)
'                mnu_Lista(intI).Caption = Trim(!TipoComunicacao)
'                mnu_Lista(intI).Tag = !Pkid
'                .MoveNext
'            Loop
'        End With
'    End If
'    mnu_Lista(0).Visible = False
'
'Err_Handle:
'End Sub
'
'Private Sub txtintCodigoLogradouro_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txtintCodigoLogradouroD_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txtstrBairroC_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", txtstrBairroC
'End Sub
'
'
'
'Private Sub txtstrBairroD_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txtstrCarteiraTrabalho_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txtstrCarteiraTrabalho
'End Sub
'
'Private Sub txtstrComplemento_GotFocus()
'    MarcaCampo txtstrComplemento
'End Sub
'
'Private Sub txtstrComplemento_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txtstrComplemento
'End Sub
'
'Private Sub txtdtmDataNascimento_GotFocus()
'    MarcaCampo txtdtmDataNascimento
'End Sub
'
'Private Sub txtdtmDataNascimento_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "D", txtdtmDataNascimento
'End Sub
'
'Private Sub txtdtmDataNascimento_LostFocus()
'    txtdtmDataNascimento = gstrDataFormatada(txtdtmDataNascimento)
'End Sub
'
'
'Private Sub txtstrDistritoC_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", txtstrDistritoC
'End Sub
'
'Private Sub txtstrIdentidade_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txtstrIdentidade
'End Sub
'
'Private Sub txtstrInscricaoEstadual_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txtstrInscricaoEstadual
'End Sub
'
'Private Sub txtstrLogradouroC_GotFocus()
''    MarcaCampo txtstrLogradouroC
'End Sub
'
'Private Sub txtstrLogradouroC_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", txtstrLogradouroC
'End Sub
'
'Private Sub txtstrLogradouroD_GotFocus()
''    MarcaCampo txtstrLogradouroD
'End Sub
'
'Private Sub txtstrLogradouroD_KeyPress(KeyAscii As Integer)
''    CaracterValido KeyAscii, "A", txtstrLogradouroD
'End Sub
'
'Private Sub txtstrLoteD_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txtstrMunicipioD_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txtstrNome_GotFocus()
'    MarcaCampo txtstrNome
'End Sub
'
'Private Sub txtstrNome_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txtstrNome
'End Sub
'
'Private Sub txtintNumero_GotFocus()
'    MarcaCampo txtintNumero
'End Sub
'
'Private Sub txtintNumero_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "N", txtintNumero
'End Sub
'
'Private Sub txtstrComplementoC_GotFocus()
'    MarcaCampo txtstrComplementoC
'End Sub
'
'Private Sub txtstrComplementoC_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txtstrComplementoC
'End Sub
'
'Private Sub txtintNumeroC_GotFocus()
'    MarcaCampo txtintNumeroC
'End Sub
'
'Private Sub txtintNumeroC_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "N", txtintNumeroC
'End Sub
'
'Private Function strQuerryGridContribuinte() As String
'    Dim strSql As String
'    strSql = ""
'    strSql = strSql & "SELECT PKId, strNome, strCNPJCPF "
'    strSql = strSql & "FROM " & gstrContribuinte & " "
'    strSql = strSql & "ORDER BY strNome"
'    strQuerryGridContribuinte = strSql
'End Function
'
'Function blnDeletaContribuinte(intAuxPKId As Integer) As Boolean
'    Dim strSql As String
'    If Trim(txtPKId) = "" Then
'        Exit Function
'    End If
'    If MsgBox("Confirma a exclusão do contribuinte '" & tdb_Lista.Columns("strNome").Value & "' ?", vbYesNo + vbQuestion) = vbYes Then
'        Set gobjBanco = New clsBanco
'        gobjBanco.ExecutaBeginTrans
'
'        DeletaTipoComunicacao intAuxPKId
'
'        strSql = ""
'        strSql = strSql & "DELETE FROM " & gstrContribuinte & " "
'        strSql = strSql & "WHERE PKId = " & txtPKId
'
'        If Not gobjBanco.Execute(strSql) Then
'            gobjBanco.ExecutaRollbackTrans
'        End If
'
'        gobjBanco.ExecutaCommitTrans
'        blnDeletaContribuinte = True
'    End If
'End Function
'
'Private Function strQueryAplicar() As String
'Dim strSql As String
'    strSql = ""
'    strSql = strSql & "SELECT PKId, strNome "
'    strSql = strSql & "FROM " & gstrContribuinte
'    strSql = strSql & " ORDER BY strNome "
'strQueryAplicar = strSql
'End Function
'
'Public Sub MantemForm(ByVal strModoOperacao As String)
'    Dim varBookMark As Variant
'    Dim strSql      As String
'    Dim lngLinha    As Long
'    Dim blnFlag     As Boolean
'    Dim blnAlterando As Boolean
'
'    On Error GoTo err_MantemForm
'
'    blnFlag = gblnListagemAutomatica
'    gblnListagemAutomatica = False
'
'    If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
'        mblnPrimeiraVez = False
'    End If
'
'    strSql = ""
'
'    Dim intCodContribuinte As Integer
'
'    Select Case UCase(strModoOperacao)
'        Case gstrNovo
'            LimpaObjeto Me, mblnAlterando
'            NovoContribuinte
'
'        Case gstrSalvar
'            If Not mblnAlterando Then
'                If gblnSistemaDemonstracao(gstrContribuinte, 50) Then
'                    Exit Sub
'                End If
'            End If
'            If blnDadosOk Then
'                blnAlterando = mblnAlterando
'                If ToolBarGeral(strModoOperacao, gstrContribuinte, mblnAlterando, tdb_Lista, Me, mobjAux, , strQueryAplicar, , , False) Then
'                    gblnListagemAutomatica = blnFlag
'                    If blnAlterando Then
'                        intCodContribuinte = tdb_Lista.Columns("PKId").Value
'                    Else: intCodContribuinte = glngPegaUltimaChave(gstrContribuinte, "PKId")
'                    End If
'                    If blnGravaTipoComunicacao(intCodContribuinte) Then
'                    End If
'                    'fra_bytNaturezaJuridica.Enabled = False
''                    MontaArray
'                    LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strQuerryGridContribuinte
'                    NovoContribuinte
'                    If Not blnAlterando Then
'                        lngLinha = X.Find(0, 0, Format(intCodContribuinte, "00000000"))
'                        If lngLinha < 0 Then
'                            LimpaObjeto Me, mblnAlterando
'                            NovoContribuinte
'                            Exit Sub
'                        Else
'                            mblnClickOk = True
'                            mblnPrimeiraVez = True
'                            tdb_Lista.MoveFirst
'                            tdb_Lista.MoveRelative lngLinha
'                            tdb_Lista_RowColChange 0, 0
'                            mblnClickOk = False
'                            mblnPrimeiraVez = False
'                        End If
'                    End If
'                   ' mblnAlterando = True
'                End If
'            End If
'
'        Case gstrDeletar
'            If blnDeletaContribuinte(txtPKId) Then
'                LimpaObjeto Me, mblnAlterando
'                NovoContribuinte
'                'MontaArray
'                LeDaTabelaParaObj gstrContribuinte, tdb_Lista, strQuerryGridContribuinte
'            End If
'
'        Case Else
'            ToolBarGeral strModoOperacao, gstrContribuinte, mblnAlterando, tdb_Lista, Me, mobjAux, , strQueryAplicar
'    End Select
'
'    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
'
'
''    dbcintTipoLogradouro.Enabled = True
''    TrocaCorObjeto dbcintTipoLogradouro, False
''    dbcintTituloLogradouro.Enabled = True
''    TrocaCorObjeto dbcintTituloLogradouro, False
''    dbcintTipoLogradouroD.Enabled = True
''    TrocaCorObjeto dbcintTipoLogradouroD, False
''    dbcintTituloLogradouroD.Enabled = True
''    TrocaCorObjeto dbcintTituloLogradouroD, False
''    cmd_TipoLogradouro.Enabled = True
''    TrocaCorObjeto cmd_TipoLogradouro, False
''    cmd_TituloLogradouro.Enabled = True
''    TrocaCorObjeto cmd_TituloLogradouro, False
''    cmd_TipoLogradouroD.Enabled = True
''    TrocaCorObjeto cmd_TipoLogradouroD, False
''    cmd_TituloLogradouroD.Enabled = True
''    TrocaCorObjeto cmd_TituloLogradouroD, False
'
'err_MantemForm:
'End Sub
'
'Private Sub txtstrNomeFantasia_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txtstrNomeFantasia
'End Sub
'
'Private Sub txtstrQuadraD_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txtstrSetorD_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii
'End Sub
'
'Private Sub txtstrTituloEleitoral_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", txtstrTituloEleitoral
'End Sub
