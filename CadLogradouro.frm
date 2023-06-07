VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadLogradouro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logradouros"
   ClientHeight    =   8085
   ClientLeft      =   2670
   ClientTop       =   2055
   ClientWidth     =   7290
   HelpContextID   =   20
   Icon            =   "CadLogradouro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   7290
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2340
      TabIndex        =   22
      Top             =   135
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5160
      Left            =   60
      TabIndex        =   30
      Top             =   30
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   9102
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Logradouro"
      TabPicture(0)   =   "CadLogradouro.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintTituloLogradouro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintTipoLogradouro"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintCodigo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintDistritoFiscal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblintSetorFiscal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblIntCep"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbldbcintBairro"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_TiposDeVias"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_LogradouroInicial"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl_LogradouroFinal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl_QuadraInicial"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl_QuadraFinal"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl_SetorInicial"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl_SetorFinal"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblCancelamento"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl_LeiDeAprovacao"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl_TipoLei"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl_strProcesso"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblExercicio"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "dbcintLogradouroFinal"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dbcintLogradouroInicial"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "dbcintTipoDeVia"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "dbcintBairro"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "dbcintSetorFiscal"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "dbcintDistritoFiscal"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtstrDescricao"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmd_Titulo"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmd_TipoLogradouro"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "dbcintTipoLogradouro"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "dbcintTituloLogradouro"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtstrCodigo"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmd_intDistritoFiscal"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmd_intSetorFiscal"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtintCep"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmd_Bairro"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmd_TiposDeVias"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtstrQuadraInicial"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtstrQuadraFinal"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtstrSetorInicial"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtstrSetorFinal"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt_dtmdtexclusao"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtstrLeiDeAprovacao"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cbobitTipoLei"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtbitDigProcesso"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtintExerProcesso"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtstrCodProcesso"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtintExercicio"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "cmd_Processo"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "CadLogradouro.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_NomeAnterior"
      Tab(1).Control(1)=   "lbl_Observacao"
      Tab(1).Control(2)=   "lvw_Historicos"
      Tab(1).Control(3)=   "txt_strObservacao"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmd_Processo 
         Height          =   315
         Left            =   6180
         Picture         =   "CadLogradouro.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Ativa Cadastro de Processos"
         Top             =   4320
         Width           =   360
      End
      Begin VB.TextBox txtintExercicio 
         Height          =   285
         Left            =   5580
         MaxLength       =   4
         TabIndex        =   21
         Top             =   4740
         Width           =   555
      End
      Begin VB.TextBox txtstrCodProcesso 
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
         Left            =   4545
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   4335
         Width           =   825
      End
      Begin VB.TextBox txtintExerProcesso 
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
         Left            =   5385
         MaxLength       =   4
         TabIndex        =   18
         Top             =   4335
         Width           =   465
      End
      Begin VB.TextBox txtbitDigProcesso 
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
         Left            =   5865
         MaxLength       =   4
         TabIndex        =   19
         Top             =   4335
         Width           =   285
      End
      Begin VB.ComboBox cbobitTipoLei 
         Height          =   315
         Left            =   1140
         TabIndex        =   16
         Top             =   4320
         Width           =   1875
      End
      Begin VB.TextBox txtstrLeiDeAprovacao 
         Height          =   285
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   20
         Top             =   4740
         Width           =   2925
      End
      Begin VB.TextBox txt_strObservacao 
         Enabled         =   0   'False
         Height          =   915
         Left            =   -73725
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   2850
         Width           =   5595
      End
      Begin VB.TextBox txt_dtmdtexclusao 
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
         Left            =   5625
         MaxLength       =   10
         TabIndex        =   1
         Top             =   420
         Width           =   1005
      End
      Begin VB.TextBox txtstrSetorFinal 
         Height          =   300
         Left            =   6270
         MaxLength       =   3
         TabIndex        =   15
         Top             =   3915
         Width           =   555
      End
      Begin VB.TextBox txtstrSetorInicial 
         Height          =   300
         Left            =   4635
         MaxLength       =   3
         TabIndex        =   14
         Top             =   3915
         Width           =   555
      End
      Begin VB.TextBox txtstrQuadraFinal 
         Height          =   300
         Left            =   2940
         MaxLength       =   3
         TabIndex        =   13
         Top             =   3915
         Width           =   555
      End
      Begin VB.TextBox txtstrQuadraInicial 
         Height          =   300
         Left            =   1140
         MaxLength       =   3
         TabIndex        =   12
         Top             =   3915
         Width           =   555
      End
      Begin VB.CommandButton cmd_TiposDeVias 
         Height          =   315
         Left            =   2295
         Picture         =   "CadLogradouro.frx":1404
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "582"
         ToolTipText     =   "Ativa Cadastro de Tipo de Via"
         Top             =   1155
         Width           =   360
      End
      Begin VB.CommandButton cmd_Bairro 
         Height          =   315
         Left            =   4785
         Picture         =   "CadLogradouro.frx":1522
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "585"
         ToolTipText     =   "Ativa Cadastro de Bairro"
         Top             =   1905
         Width           =   360
      End
      Begin VB.TextBox txtintCep 
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
         Left            =   1155
         MaxLength       =   20
         TabIndex        =   2
         Top             =   780
         Width           =   915
      End
      Begin VB.CommandButton cmd_intSetorFiscal 
         Height          =   315
         Left            =   4785
         Picture         =   "CadLogradouro.frx":1640
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "586"
         ToolTipText     =   "Ativa Cadastro de Setor Fiscal"
         Top             =   2670
         Width           =   360
      End
      Begin VB.CommandButton cmd_intDistritoFiscal 
         Height          =   315
         Left            =   4785
         Picture         =   "CadLogradouro.frx":175E
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "585"
         ToolTipText     =   "Ativa Cadastro de Distrito Fiscal"
         Top             =   2280
         Width           =   360
      End
      Begin VB.TextBox txtstrCodigo 
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
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   0
         Top             =   435
         Width           =   915
      End
      Begin MSDataListLib.DataCombo dbcintTituloLogradouro 
         Height          =   315
         Left            =   5085
         TabIndex        =   5
         Top             =   1155
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintTipoLogradouro 
         Height          =   315
         Left            =   3150
         TabIndex        =   4
         Top             =   1155
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.CommandButton cmd_TipoLogradouro 
         Height          =   315
         Left            =   4155
         Picture         =   "CadLogradouro.frx":187C
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "582"
         ToolTipText     =   "Ativa Cadastro de Tipo de Logradouro"
         Top             =   1155
         Width           =   360
      End
      Begin VB.CommandButton cmd_Titulo 
         Height          =   315
         Left            =   6585
         Picture         =   "CadLogradouro.frx":199A
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "583"
         ToolTipText     =   "Ativa Cadastro de Título de Logradouro"
         Top             =   1170
         Width           =   360
      End
      Begin VB.TextBox txtstrDescricao 
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
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1545
         Width           =   5805
      End
      Begin MSDataListLib.DataCombo dbcintDistritoFiscal 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Top             =   2280
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintSetorFiscal 
         Height          =   315
         Left            =   1155
         TabIndex        =   9
         Top             =   2670
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSComctlLib.ListView lvw_Historicos 
         Height          =   1980
         Left            =   -74730
         TabIndex        =   31
         Top             =   630
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   3493
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
            Text            =   "Descrição"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Via"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Título"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lei de aprovação"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Tipo da lei"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Exercicío"
            Object.Width           =   1517
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Processo"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Alteração"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Observação"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbcintBairro 
         Height          =   315
         Left            =   1170
         TabIndex        =   7
         Top             =   1905
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintTipoDeVia 
         Height          =   315
         Left            =   1155
         TabIndex        =   3
         Top             =   1155
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintLogradouroInicial 
         Height          =   315
         Left            =   1620
         TabIndex        =   10
         Top             =   3075
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintLogradouroFinal 
         Height          =   315
         Left            =   1620
         TabIndex        =   11
         Top             =   3510
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercicío"
         Height          =   195
         Left            =   4830
         TabIndex        =   54
         Top             =   4800
         Width           =   675
      End
      Begin VB.Label lbl_strProcesso 
         AutoSize        =   -1  'True
         Caption         =   "Nº do Processo"
         Height          =   195
         Left            =   3360
         TabIndex        =   53
         Top             =   4335
         Width           =   1110
      End
      Begin VB.Label lbl_TipoLei 
         AutoSize        =   -1  'True
         Caption         =   "Tipo da Lei"
         Height          =   195
         Left            =   210
         TabIndex        =   52
         Top             =   4395
         Width           =   795
      End
      Begin VB.Label lbl_LeiDeAprovacao 
         AutoSize        =   -1  'True
         Caption         =   "Lei de aprovação"
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   4785
         Width           =   1245
      End
      Begin VB.Label lbl_Observacao 
         AutoSize        =   -1  'True
         Caption         =   "Observação"
         Height          =   195
         Left            =   -74670
         TabIndex        =   50
         Top             =   2850
         Width           =   900
      End
      Begin VB.Label lblCancelamento 
         AutoSize        =   -1  'True
         Caption         =   "Cancelamento"
         Height          =   195
         Left            =   4500
         TabIndex        =   48
         Top             =   495
         Width           =   1020
      End
      Begin VB.Label lbl_SetorFinal 
         AutoSize        =   -1  'True
         Caption         =   "Setor Final"
         Height          =   195
         Left            =   5340
         TabIndex        =   47
         Top             =   3975
         Width           =   750
      End
      Begin VB.Label lbl_SetorInicial 
         AutoSize        =   -1  'True
         Caption         =   "Setor Inicial"
         Height          =   195
         Left            =   3675
         TabIndex        =   46
         Top             =   3960
         Width           =   825
      End
      Begin VB.Label lbl_QuadraFinal 
         AutoSize        =   -1  'True
         Caption         =   "Quadra Final"
         Height          =   195
         Left            =   1905
         TabIndex        =   45
         Top             =   3960
         Width           =   900
      End
      Begin VB.Label lbl_QuadraInicial 
         AutoSize        =   -1  'True
         Caption         =   "Quadra Inicial"
         Height          =   195
         Left            =   90
         TabIndex        =   44
         Top             =   3945
         Width           =   975
      End
      Begin VB.Label lbl_LogradouroFinal 
         Caption         =   "Logradouro Final"
         Height          =   300
         Left            =   195
         TabIndex        =   43
         Top             =   3570
         Width           =   1275
      End
      Begin VB.Label lbl_LogradouroInicial 
         Caption         =   "Logradouro Inicial"
         Height          =   300
         Left            =   180
         TabIndex        =   42
         Top             =   3120
         Width           =   1275
      End
      Begin VB.Label lbl_TiposDeVias 
         AutoSize        =   -1  'True
         Caption         =   "Via"
         Height          =   195
         Left            =   675
         TabIndex        =   41
         Top             =   1200
         Width           =   225
      End
      Begin VB.Label lbldbcintBairro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bairro"
         Height          =   195
         Left            =   660
         TabIndex        =   40
         Top             =   2025
         Width           =   405
      End
      Begin VB.Label lblIntCep 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Left            =   705
         TabIndex        =   39
         Top             =   870
         Width           =   360
      End
      Begin VB.Label lbl_NomeAnterior 
         AutoSize        =   -1  'True
         Caption         =   "Nome atual"
         Height          =   195
         Left            =   -74355
         TabIndex        =   38
         Top             =   1605
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblintSetorFiscal 
         AutoSize        =   -1  'True
         Caption         =   "Setor Fiscal"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   2790
         Width           =   825
      End
      Begin VB.Label lblintDistritoFiscal 
         AutoSize        =   -1  'True
         Caption         =   "Distrito Fiscal"
         Height          =   195
         Left            =   135
         TabIndex        =   36
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label lblintCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   570
         TabIndex        =   35
         Top             =   510
         Width           =   495
      End
      Begin VB.Label lblintTipoLogradouro 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   2760
         TabIndex        =   34
         Top             =   1215
         Width           =   315
      End
      Begin VB.Label lblintTituloLogradouro 
         AutoSize        =   -1  'True
         Caption         =   "Título"
         Height          =   195
         Left            =   4560
         TabIndex        =   33
         Top             =   1215
         Width           =   420
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   345
         TabIndex        =   32
         Top             =   1635
         Width           =   720
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   2805
      Left            =   0
      TabIndex        =   29
      Top             =   5280
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4948
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
      Columns(1).Caption=   "CodNum"
      Columns(1).DataField=   "CodNum"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Código"
      Columns(2).DataField=   "strCodigo"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Via"
      Columns(3).DataField=   "TipoDeVia"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Tipo"
      Columns(4).DataField=   "Tipo"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Título"
      Columns(5).DataField=   "Titulo"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Nome do Logradouro"
      Columns(6).DataField=   "Logradouro"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Logradouro por Extenso"
      Columns(7).DataField=   "LogradouroExtenso"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Bairro"
      Columns(8).DataField=   "Bairro"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "strCodProcesso"
      Columns(9).DataField=   "strCodProcesso"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "intExerProcesso"
      Columns(10).DataField=   "intExerProcesso"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "bitDigProcesso"
      Columns(11).DataField=   "bitDigProcesso"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "strLeiDeAprovacao"
      Columns(12).DataField=   "strLeiDeAprovacao"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "intExercicio"
      Columns(13).DataField=   "intExercicio"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "bitTipoLei"
      Columns(14).DataField=   "bitTipoLei"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "dtmdtExclusao"
      Columns(15).DataField=   "dtmdtExclusao"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=2037"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1958"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1085"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1005"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=979"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=900"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=2170"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2090"
      Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=7408"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=7329"
      Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(40)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(43)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(44)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(46)=   "Column(8).Width=4842"
      Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=4763"
      Splits(0)._ColumnProps(49)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(51)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(52)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(54)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(9).Visible=0"
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
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=67"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=67"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=68"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=69"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=71"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=86,.parent=67,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=83,.parent=68"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=84,.parent=69"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=85,.parent=71"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=90,.parent=67"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=87,.parent=68"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=88,.parent=69"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=89,.parent=71"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=94,.parent=67"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=68"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=69"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=71"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=98,.parent=67"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=95,.parent=68"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=96,.parent=69"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=97,.parent=71"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=102,.parent=67"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=99,.parent=68"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=100,.parent=69"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=71"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=106,.parent=67"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=103,.parent=68"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=104,.parent=69"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=105,.parent=71"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=110,.parent=67"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=107,.parent=68"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=108,.parent=69"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=109,.parent=71"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=16,.parent=67"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=13,.parent=68"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=14,.parent=69"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=15,.parent=71"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=24,.parent=67"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=21,.parent=68"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=22,.parent=69"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=23,.parent=71"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=20,.parent=67"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=17,.parent=68"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=18,.parent=69"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=19,.parent=71"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=28,.parent=67"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=25,.parent=68"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=26,.parent=69"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=27,.parent=71"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=32,.parent=67"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=29,.parent=68"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=30,.parent=69"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=31,.parent=71"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=46,.parent=67"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=43,.parent=68"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=44,.parent=69"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=45,.parent=71"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=50,.parent=67"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=47,.parent=68"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=48,.parent=69"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=49,.parent=71"
      _StyleDefs(100) =   "Named:id=33:Normal"
      _StyleDefs(101) =   ":id=33,.parent=0"
      _StyleDefs(102) =   "Named:id=34:Heading"
      _StyleDefs(103) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(104) =   ":id=34,.wraptext=-1"
      _StyleDefs(105) =   "Named:id=35:Footing"
      _StyleDefs(106) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(107) =   "Named:id=36:Selected"
      _StyleDefs(108) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(109) =   "Named:id=37:Caption"
      _StyleDefs(110) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(111) =   "Named:id=38:HighlightRow"
      _StyleDefs(112) =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(113) =   "Named:id=39:EvenRow"
      _StyleDefs(114) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(115) =   "Named:id=40:OddRow"
      _StyleDefs(116) =   ":id=40,.parent=33"
      _StyleDefs(117) =   "Named:id=41:RecordSelector"
      _StyleDefs(118) =   ":id=41,.parent=34"
      _StyleDefs(119) =   "Named:id=42:FilterBar"
      _StyleDefs(120) =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadLogradouro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando           As Boolean
    Dim mblnClickOk             As Boolean
    Dim mobjAux                 As Object
    Dim oList                   As Object
    Dim mblnPrimeiraVez         As Boolean
    Dim blnObservacao           As Boolean
    
  ' TIMTIM - 10/02/2003 - Pendência nº 1
    Dim bytOrdenacao            As Byte
    Dim blnOrdenacaoAsc         As Boolean
    Dim blnPertenceAoMunicipio  As Boolean

Private Sub cbobitTipoLei_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmd_Bairro_Click()
    CarregaForm frmCadBairro, dbcintBairro
End Sub

Private Sub cmd_intDistritoFiscal_Click()
    CarregaForm frmCadDistritoFiscal, dbcintDistritoFiscal
End Sub

Private Sub cmd_intSetorFiscal_Click()
    CarregaForm frmCadSetorFiscal, dbcintSetorFiscal
End Sub

Private Sub cmd_Processo_Click()
    CarregaForm frmCadProtocolizacaoProcesso, txtstrcodprocesso
End Sub

Private Sub cmd_TiposDeVias_Click()
    CarregaForm frmCadTiposDeVias, dbcintTipoDeVia
End Sub

Private Sub dbcintBairro_Click(Area As Integer)
    DropDownDataCombo dbcintBairro, Me, Area
End Sub

Private Sub dbcintBairro_GotFocus()
    MarcaCampo dbcintBairro
End Sub

Private Sub dbcintBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintBairro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintBairro
End Sub

Private Sub dbcintDistritoFiscal_Click(Area As Integer)
   DropDownDataCombo dbcintDistritoFiscal, Me, Area
End Sub

Private Sub dbcintDistritoFiscal_GotFocus()
    MarcaCampo dbcintDistritoFiscal
End Sub

Private Sub dbcintDistritoFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintDistritoFiscal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouroFinal_Click(Area As Integer)
    DropDownDataCombo dbcintLogradouroFinal, Me, Area
End Sub

Private Sub dbcintLogradouroFinal_GotFocus()
    MarcaCampo dbcintLogradouroFinal
End Sub

Private Sub dbcintLogradouroFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLogradouroFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouroFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintLogradouroFinal
End Sub


Private Sub dbcintLogradouroInicial_Click(Area As Integer)
    DropDownDataCombo dbcintLogradouroInicial, Me, Area
End Sub

Private Sub dbcintLogradouroInicial_GotFocus()
    MarcaCampo dbcintLogradouroInicial
End Sub

Private Sub dbcintLogradouroInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLogradouroInicial, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouroInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintLogradouroInicial
End Sub

Private Sub dbcintSetorFiscal_Click(Area As Integer)
   DropDownDataCombo dbcintSetorFiscal, Me, Area
End Sub

Private Sub dbcintSetorFiscal_GotFocus()
    MarcaCampo dbcintSetorFiscal
End Sub

Private Sub dbcintSetorFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintSetorFiscal, Me, , KeyCode, Shift
End Sub

Private Sub dbcintSetorFiscal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintSetorFiscal
End Sub

Private Sub dbcintTipoDeVia_Click(Area As Integer)
    DropDownDataCombo dbcintTipoDeVia, Me, Area
End Sub

Private Sub dbcintTipoDeVia_GotFocus()
    MarcaCampo dbcintTipoDeVia
End Sub

Private Sub dbcintTipoDeVia_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTipoDeVia, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoDeVia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTipoDeVia
End Sub

Private Sub dbcintTipoLogradouro_Click(Area As Integer)
   DropDownDataCombo dbcintTipoLogradouro, Me, Area
End Sub

Private Sub dbcintTipoLogradouro_GotFocus()
    MarcaCampo dbcintTipoLogradouro
End Sub

Private Sub dbcintTipoLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTipoLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintTiposDeVias_Click(Area As Integer)
    DropDownDataCombo dbcintTipoLogradouro, Me, Area
End Sub

Private Sub dbcintTiposDeVias_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTipoLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTiposDeVias_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintTipoDeVia
End Sub

Private Sub dbcintTituloLogradouro_Click(Area As Integer)
   DropDownDataCombo dbcintTituloLogradouro, Me, Area
End Sub

Private Sub dbcintTituloLogradouro_GotFocus()
    MarcaCampo dbcintTituloLogradouro
End Sub

Private Sub dbcintTituloLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTituloLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTituloLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cmd_TipoLogradouro_Click()
    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouro
End Sub

Private Sub cmd_Titulo_Click()
    CarregaForm frmCadTituloLogradouro, dbcintTituloLogradouro
End Sub

Private Sub Form_Activate()
    Select Case UCase(App.ProductName)
        Case "PROTOCOLO"
            gintCodSeguranca = 1264
            'strSistema = "H"
        Case "MENOR"
            gintCodSeguranca = 123
            'strSistema = "D"
        Case Else
            gintCodSeguranca = 584
    End Select

    VirificaGradeListView Me
    'If mblnAlterando Then
    '    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    'Else
    '    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    'End If

    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    
    If mblnAlterando Then
       If tdb_Lista.Columns("dtmdtExclusao") = "" Then
          HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
          HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
       Else
          HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
          HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
       End If
    Else
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
    End If

End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()

 ' TIMTIM - 10/02/2003 - Pendência nº 1
   bytOrdenacao = 5: blnOrdenacaoAsc = True
         
   mblnAlterando = False
   dbcintDistritoFiscal.Tag = strQueryDataComboDistritoFiscal & ";strDescricao"
   dbcintBairro.Tag = strQueryDataComboBairro & ";strDescricao"
   dbcintSetorFiscal.Tag = strQueryDataComboSetorFiscal & ";strDescricao"
   dbcintTipoDeVia.Tag = strQueryTiposDeVias & ";strSigla"
   dbcintTipoLogradouro.Tag = gstrQueryTipoLogradouro & ";strSigla"
   dbcintTituloLogradouro.Tag = gstrQueryTituloLogradouro(True) & ";strSigla"
   dbcintLogradouroInicial.Tag = strQueryLogradouro & ";lo.strDescricao"
   dbcintLogradouroFinal.Tag = strQueryLogradouro & ";lo.strDescricao"
   
   blnPertenceAoMunicipio = True
   
   VerificaObjParaAplicar mobjAux
   
   cbobitTipoLei.AddItem "Lei"
   cbobitTipoLei.ItemData(cbobitTipoLei.NewIndex) = 0
   cbobitTipoLei.AddItem "Decreto"
   cbobitTipoLei.ItemData(cbobitTipoLei.NewIndex) = 1
   
   TrocaCorObjeto txt_dtmdtexclusao, True
   TrocaCorObjeto txt_strObservacao, True
   blnObservacao = False
   
End Sub

Private Function strQuery() As String
    
Dim strSql  As String

   strSql = ""
   strSql = strSql & "SELECT L.PKId, /*L.INTCEP,*/ "
   
   If bytDBType = SQLServer Then
      strSql = strSql & " REPLICATE('0', 10 - LEN(L.strCodigo)) + L.strCodigo CodNum, "
   Else
      strSql = strSql & " LPAD(L.strCodigo, 10, '0') CodNum, "
   End If
   
   strSql = strSql & " L.strCodigo, TL.strSigla AS Tipo, U.strDescricao AS Titulo, B.strDescricao as Bairro, "
   strSql = strSql & " RTRIM(LTRIM(L.strDescricao)) " & strCONCAT & "' '" & strCONCAT & " RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & _
            strCONCAT & "' '" & strCONCAT & gstrISNULL("U.strDescricao", "''") & ")) AS LogradouroExtenso,  "
   strSql = strSql & " RTRIM(LTRIM(L.strDescricao)) AS Logradouro, "
   strSql = strSql & " TV.strSigla TipoDeVia, "
   strSql = strSql & "L.Dtmdtexclusao, "
   
   strSql = strSql & "L.intExercicio, "
   strSql = strSql & "L.strCodProcesso, "
   strSql = strSql & "L.intExerProcesso, "
   strSql = strSql & "L.bitDigProcesso, "
   strSql = strSql & "L.strLeiDeAprovacao, "
   strSql = strSql & "L.bitTipoLei "
   
   strSql = strSql & "FROM " & gstrLogradouro & " L, "
   strSql = strSql & gstrTituloLogradouro & " U, "
   strSql = strSql & gstrTipoLogradouro & " TL, "
   strSql = strSql & gstrBairro & " B, "
   strSql = strSql & gstrTiposDeVias & " TV "
   strSql = strSql & " WHERE "
   strSql = strSql & "L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId " & strOUTJOracle & " AND "
   strSql = strSql & "L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId " & strOUTJOracle & " AND "
   strSql = strSql & "L.intTipoDeVia " & strOUTJSQLServer & "=" & " TV.Pkid " & strOUTJOracle & " AND "
   strSql = strSql & "L.intBairro " & strOUTJSQLServer & "= B.PKId " & strOUTJOracle
   
   Select Case bytOrdenacao
      
      Case Is = 1
         strSql = strSql & "ORDER BY " & gstrCONVERT(CDT_INT, "L.strCodigo") & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 2
         strSql = strSql & "ORDER BY TipoDeVia" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 3
         strSql = strSql & "ORDER BY Tipo" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 4
         strSql = strSql & "ORDER BY Titulo" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
      Case Is = 5
          strSql = strSql & "ORDER BY Logradouro" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
   End Select
   
   strQuery = strSql
    
End Function

Private Function strQueryDataComboDistritoFiscal()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrDistritoFiscal & " "
    strSql = strSql & "ORDER BY strDescricao"
    strQueryDataComboDistritoFiscal = strSql
End Function

Private Function strQueryDataComboSetorFiscal()
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao "
    strSql = strSql & "FROM " & gstrSetorFiscal & " "
    strSql = strSql & "ORDER BY strDescricao"
    strQueryDataComboSetorFiscal = strSql
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub lvw_Historicos_Click()
    If lvw_Historicos.ListItems.Count > 0 Then
       txt_strObservacao.Text = lvw_Historicos.SelectedItem.SubItems(9)
    End If
End Sub

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        gOrdenaGrid tdb_Lista, 1
    Else
        gOrdenaGrid tdb_Lista, ColIndex
    End If
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            blnPertenceAoMunicipio = True
            mblnClickOk = False
            mblnAlterando = True
            txtPKId.Text = .Columns("PKID").Value
                If mblnPrimeiraVez Then
                NovoLogradouro
                LeDaTabelaParaObj gstrLogradouro, Me
                PreencheProcesso
                PreencheHistorico txtPKId.Text
                gCorLinhaSelecionada tdb_Lista
                txt_dtmdtexclusao.Text = tdb_Lista.Columns("dtmdtExclusao")
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                If tdb_Lista.Columns("dtmdtExclusao") = "" Then
                   HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                   HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar
                Else
                   HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
                   HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
                End If
                'CarregaListView_Ceps_Historicos .Columns("PKID").Value
                'txt_NomeAnterior = .Columns("Logradouro").Text
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim intCodLogradouro        As Long
    Dim blnAuxAlterando         As Boolean
    
        
    If txtintCep <> "" Then
        txtintCep = Int(Replace(txtintCep, "-", ""))
    End If
    
    If mblnAlterando Then
       intCodLogradouro = IIf(IsNull(tdb_Lista.Columns("PKID").Value), 0, tdb_Lista.Columns("PKID").Value)
    
    End If
    blnAuxAlterando = mblnAlterando
    
    Select Case UCase(strModoOperacao)
    
    Case "NOVO"
        LimpaObjeto Me, mblnAlterando
        NovoLogradouro
    
    Case "SALVAR"
        If blnDadosOk(intCodLogradouro) Then
           Set gobjBanco = New clsBanco
           gobjBanco.ExecutaBeginTrans
           
           If ToolBarGeral(strModoOperacao, gstrLogradouro, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, , , , False) Then
              If (dbcintTipoDeVia.Text <> tdb_Lista.Columns("Via").Value Or dbcintTipoLogradouro.Text <> tdb_Lista.Columns("Tipo").Value Or txt_strObservacao <> "" Or _
                  dbcintTituloLogradouro.Text <> tdb_Lista.Columns("Título").Value Or txtstrdescricao.Text <> tdb_Lista.Columns("Logradouro").Value) Then
                 If GravaHistorico(intCodLogradouro) = False Then
                    ExibeMensagem "Ocorreu um erro ao gravar o Histórico do Logradouro. " & _
                                  "Os dados não foram gravados."
                    gobjBanco.ExecutaRollbackTrans
                    mblnAlterando = blnAuxAlterando
                    Exit Sub
                 End If
              End If
              Set gobjBanco = New clsBanco
              gobjBanco.ExecutaCommitTrans
              LimpaObjeto Me
              NovoLogradouro
              LeDaTabelaParaObj gstrLogradouro, tdb_Lista, strQuery
           Else
              gobjBanco.ExecutaRollbackTrans
              mblnAlterando = blnAuxAlterando
           End If
        End If
    
    Case "DELETAR"
        If gblnExclusaoGravacaoOk("", "Deseja cancelar o logradouro '" & tdb_Lista.Columns("Logradouro") & "' ?", True) = True Then
           If CancelaLogradouro(txtPKId.Text) Then
              LimpaObjeto Me
              NovoLogradouro
              LeDaTabelaParaObj gstrLogradouro, tdb_Lista, strQuery
           End If
        End If
    
    Case "APLICAR"
        ToolBarGeral strModoOperacao, gstrLogradouro, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, , rptLogradouro, strQueryRelatorio
    
    Case Else
        ToolBarGeral strModoOperacao, gstrLogradouro, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, , rptLogradouro, strQueryRelatorio
    
    End Select
    
    If Not gobjGeral Is Nothing Then HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    
    
    txtintCep = gstrCEPFormatado(txtintCep)
    
End Sub

Private Function blnDadosOk(intLogradouro As Long) As Boolean
    Dim adoAux       As ADODB.Recordset
    Dim strAuxNome   As String
    Dim strSql       As String
    Dim i As Integer
    
    If Trim(txtstrCodigo) = "" Then
        ExibeMensagem "O Código do logradouro deve ser informado."
        txtstrCodigo.SetFocus
        Exit Function
    End If
    
    If txtintCep.Text = "" Then
        ExibeMensagem "O CEP deve ser informado."
        txtintCep.SetFocus
        Exit Function
    End If
    
    If Not dbcintTipoLogradouro.MatchedWithList Then
        ExibeMensagem "Informe o Tipo do logradouro."
        If dbcintTipoLogradouro.Enabled Then dbcintTipoLogradouro.SetFocus
        Exit Function
    End If
    
    If Trim(txtstrdescricao) = "" Then
        ExibeMensagem "A descrição do logradouro deve ser informada."
        txtstrdescricao.SetFocus
        Exit Function
    End If
    
    If Not dbcintBairro.MatchedWithList Then
        ExibeMensagem "O Bairro deve ser informado."
        If dbcintBairro.Enabled Then dbcintBairro.SetFocus
        Exit Function
    End If
    
    If Not mblnAlterando Then
        
        strSql = ""
        strSql = strSql & "SELECT PkID FROM "
        strSql = strSql & gstrLogradouro
        strSql = strSql & " WHERE strCodigo = '" & txtstrCodigo.Text & "'"
        strSql = strSql & " AND dtmDtExclusao is null "
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSql, 10, adoAux) Then
            If adoAux.RecordCount > 0 Then
               ExibeMensagem "Código de logradouro " & txtstrCodigo & " já se encontra cadastrado."
               txtstrCodigo.SetFocus
               Exit Function
            End If
        Else
            Exit Function
        End If
        
        strSql = ""
        strSql = strSql & "SELECT PkID FROM "
        strSql = strSql & gstrLogradouro
        strSql = strSql & " WHERE "
        strSql = strSql & " strDescricao = '" & txtstrdescricao & "'"
        
        If dbcintTipoLogradouro.BoundText <> "" Then
            strSql = strSql & " AND intTipoLogradouro = '" & dbcintTipoLogradouro.BoundText & "'"
        Else
            strSql = strSql & " AND intTipoLogradouro is null "
        End If
        
        If dbcintTituloLogradouro.BoundText <> "" Then
            strSql = strSql & " AND intTituloLogradouro = '" & dbcintTituloLogradouro.BoundText & "'"
        Else
            strSql = strSql & " AND intTituloLogradouro is null "
        End If
        
        If Replace(txtintCep, "-", "") <> "" Then
            strSql = strSql & " AND INTCEP = '" & Replace(txtintCep, "-", "") & "'"
        Else
            strSql = strSql & " AND INTCEP is null "
        End If
        
        If dbcintBairro.BoundText <> "" Then
            strSql = strSql & " AND intBairro = '" & dbcintBairro.BoundText & "'"
        Else
            strSql = strSql & " AND intBairro is null "
        End If
        
        If dbcintLogradouroInicial.BoundText <> "" Then
            strSql = strSql & " AND intLogradouroInicial = '" & dbcintLogradouroInicial.BoundText & "'"
        Else
            strSql = strSql & " AND intLogradouroInicial is null "
        End If
        
        If dbcintLogradouroFinal.BoundText <> "" Then
            strSql = strSql & " AND intLogradouroFinal = '" & dbcintLogradouroFinal.BoundText & "'"
        Else
            strSql = strSql & " AND intLogradouroFinal is null "
        End If
            
        If dbcintDistritoFiscal.BoundText <> "" Then
            strSql = strSql & " AND intDistritoFiscal = '" & dbcintDistritoFiscal.BoundText & "'"
        Else
            strSql = strSql & " AND intDistritoFiscal is null "
        End If
            
        If dbcintSetorFiscal.BoundText <> "" Then
            strSql = strSql & " AND intSetorFiscal = '" & dbcintSetorFiscal.BoundText & "'"
        Else
            strSql = strSql & " AND intSetorFiscal  is null "
        End If
            
        
        If dbcintTipoDeVia.BoundText <> "" Then
            strSql = strSql & " AND intTipoDeVia = '" & dbcintTipoDeVia.BoundText & "'"
        Else
            strSql = strSql & " AND intTipoDeVia  is null "
        End If
            
        If txtstrSetorInicial <> "" Then
            strSql = strSql & " AND strSetorInicial = '" & txtstrSetorInicial & "'"
        Else
            strSql = strSql & " AND strSetorInicial is null "
        End If
        
        If txtstrSetorFinal <> "" Then
            strSql = strSql & " AND strSetorFinal = '" & txtstrSetorFinal & "'"
        Else
            strSql = strSql & " AND strSetorFinal is null "
        End If
        
        If txtstrQuadraInicial <> "" Then
            strSql = strSql & " AND strQuadraInicial = '" & txtstrQuadraInicial & "'"
        Else
            strSql = strSql & " AND strQuadraInicial is null "
        End If
        
        If txtstrQuadraFinal <> "" Then
            strSql = strSql & " AND strQuadraFinal = '" & txtstrQuadraFinal & "'"
        Else
            strSql = strSql & " AND strQuadraFinal is null "
        End If
        
        If txtintExercicio <> "" Then
            strSql = strSql & " AND intExercicio = '" & txtintExercicio & "'"
        Else
            strSql = strSql & " AND intExercicio  is null "
        End If
        
        If Not gstrItemData(cbobitTipoLei, True) = "NULL" Then
            strSql = strSql & " AND bitTipoLei = '" & gstrItemData(cbobitTipoLei) & "'"
        Else
            strSql = strSql & " AND bitTipoLei is null "
        End If
            
        If txtstrcodprocesso <> "" Then
            strSql = strSql & " AND strCodProcesso = '" & txtstrcodprocesso & "'"
        Else
            strSql = strSql & " AND strCodProcesso is null "
        End If
        
        If txtintexerprocesso <> "" Then
            strSql = strSql & " AND intExerProcesso = '" & txtintexerprocesso & "'"
        Else
            strSql = strSql & " AND intExerProcesso  is null "
        End If
            
        If txtbitdigprocesso <> "" Then
            strSql = strSql & " AND bitDigProcesso = '" & txtbitdigprocesso & "'"
        Else
            strSql = strSql & " AND bitDigProcesso is null "
        End If
            
        If txtstrLeiDeAprovacao <> "" Then
            strSql = strSql & " AND strLeiDeAprovacao = '" & txtstrLeiDeAprovacao & "'"
        Else
            strSql = strSql & " AND strLeiDeAprovacao is null "
        End If
        
        strSql = strSql & " AND dtmDtExclusao is null "
                
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSql, 10, adoAux) Then
            If adoAux.RecordCount > 0 Then
               ExibeMensagem "Logradouro já cadastrado."
               txtstrCodigo.SetFocus
               Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    If Not blnPertenceAoMunicipio Then
        ExibeMensagem "O logradouro selecionado não pertence ao municipio."
        Exit Function
    End If
    
    If blnCamposObrigatorios Then
        If cbobitTipoLei.ListIndex = -1 Then
           ExibeMensagem "O Tipo da Lei deve ser informado."
           tab_3dPasta.Tab = 0
           cbobitTipoLei.SetFocus
           Exit Function
        End If
    
        If (Trim(txtstrcodprocesso.Text) <> "" And Trim(txtintexerprocesso.Text) <> "" And Trim(txtbitdigprocesso.Text) <> "") Then
           If gblnExisteCodigo(2, gstrProtocolizacaoProcesso, "strCodigo", "'" & Trim(txtstrcodprocesso.Text) & "'", _
              "intExercicio", Trim(txtintexerprocesso.Text), "bitDigito", Trim(txtbitdigprocesso.Text)) = False Then
              ExibeMensagem "O Processo informado não existe."
              tab_3dPasta.Tab = 0
              txtstrcodprocesso.SetFocus
              Exit Function
           End If
        Else
           If (Trim(txtstrcodprocesso.Text) = "" And Trim(txtintexerprocesso.Text) = "" And Trim(txtbitdigprocesso.Text) = "") Then
              ExibeMensagem "O Processo deve ser informado."
              tab_3dPasta.Tab = 0
              txtstrcodprocesso.SetFocus
              Exit Function
           Else
              ExibeMensagem "O Processo deve ser preenchido corretamente."
              tab_3dPasta.Tab = 0
              txtstrcodprocesso.SetFocus
              Exit Function
           End If
        End If
        
        If Trim(txtstrLeiDeAprovacao.Text) = "" Then
           ExibeMensagem "A Lei deve ser preenchida corretamente."
           tab_3dPasta.Tab = 0
           txtstrLeiDeAprovacao.SetFocus
           Exit Function
        End If
        
        If Trim(txtintExercicio.Text) = "" Then
           ExibeMensagem "O Exercicío da Lei deve ser preenchido corretamente."
           tab_3dPasta.Tab = 0
           txtintExercicio.SetFocus
           Exit Function
        End If
        
        'Verifica se os campos abaixo foram alterados
        If mblnAlterando = True And (dbcintTipoDeVia.Text <> tdb_Lista.Columns("Via").Value Or dbcintTipoLogradouro.Text <> tdb_Lista.Columns("Tipo").Value Or _
                                    dbcintTituloLogradouro.Text <> tdb_Lista.Columns("Título").Value Or txtstrdescricao.Text <> tdb_Lista.Columns("Logradouro").Value) Then
           'Verifica se os campos de Lei e Processo são válidos
           If blnLeiOK(intLogradouro) = False Then
              Exit Function
           End If
        End If
    End If
    

    If mblnAlterando Then
        If blnObservacao = False Then
           blnObservacao = True
           txt_strObservacao.Text = ""
           TrocaCorObjeto txt_strObservacao, False
           If gblnExclusaoGravacaoOk("", "Deseja fazer alguma observação para o logradouro antigo?", True) = True Then
              tab_3dPasta.Tab = 1
              txt_strObservacao.SetFocus
              Exit Function
           Else
              blnDadosOk = True
              Exit Function
           End If
        End If
    End If
    
    blnDadosOk = True
End Function

Private Sub PreencheProcesso()
Dim adoProcesso As ADODB.Recordset
Dim strSql As String

    'Rafael 05/10/2004
    'Essa função foi desenvolvida, pois quando trazia os digito do processo,
    'se fosse nulo, ele colocava zero (isso usando LeDaTabelaParaObj)
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "strCodProcesso, intExerProcesso, bitDigProcesso "
    strSql = strSql & "FROM " & gstrLogradouro & " "
    strSql = strSql & "WHERE pkID = " & txtPKId.Text
    
    Set adoProcesso = New ADODB.Recordset
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoProcesso) Then
       With adoProcesso
         txtstrcodprocesso.Text = gstrENulo(!strCodProcesso)
         txtintexerprocesso.Text = gstrENulo(!intExerProcesso)
         txtbitdigprocesso.Text = gstrENulo(!bitDigProcesso)
       End With
    End If
    Set adoProcesso = Nothing
    
End Sub

Private Function CancelaLogradouro(Pkid As Long) As Boolean
Dim adoProcesso As ADODB.Recordset
Dim strSql As String
    strSql = ""
    strSql = strSql & "UPDATE "
    strSql = strSql & gstrLogradouro & " "
    strSql = strSql & "SET dtmdtExclusao = " & Format$(strGETDATE, "yyyy/mm/dd")
    strSql = strSql & "WHERE pkID = " & Pkid
    
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSql) = False Then
       ExibeMensagem "Ocorreu um erro ao cancelar o logradouro. Logradouro não cancelado."
       Exit Function
    End If
    CancelaLogradouro = True
End Function

Private Function blnLeiOK(intLogradouro As Long) As Boolean
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    If txtstrcodprocesso.Text = tdb_Lista.Columns("strCodProcesso").Value And _
       txtintexerprocesso.Text = tdb_Lista.Columns("intExerProcesso").Value And _
       txtbitdigprocesso.Text = tdb_Lista.Columns("bitDigProcesso").Value Then
       ExibeMensagem "O Processo deve ser diferente do atual."
       tab_3dPasta.Tab = 0
       txtstrcodprocesso.SetFocus
       Exit Function
    End If
    
    strSql = strSql & "SELECT pkID "
    strSql = strSql & "FROM "
    strSql = strSql & gstrHistoricoLogradouro & " "
    strSql = strSql & "WHERE "
    strSql = strSql & "strCodProcesso = '" & Trim(txtstrcodprocesso.Text) & "' AND "
    strSql = strSql & "intExerProcesso = " & Trim(txtintexerprocesso.Text) & " AND "
    strSql = strSql & "bitDigProcesso = " & Trim(txtbitdigprocesso.Text)
    
    If intLogradouro <> 0 Then
       strSql = strSql & " AND intLogradouro = " & intLogradouro & ""
    End If
    
    strSql = strSql & "ORDER BY dtmdtAlteracao "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
       If Not adoResultado.EOF Then
          ExibeMensagem "O Processo informado já existe no histórico."
          tab_3dPasta.Tab = 0
          txtstrcodprocesso.SetFocus
          Set adoResultado = Nothing
          Exit Function
       End If
    End If
   
    If txtstrLeiDeAprovacao.Text = tdb_Lista.Columns("strLeiDeAprovacao").Value Then
       ExibeMensagem "A Lei de Aprovação deve ser diferente da atual."
       tab_3dPasta.Tab = 0
       txtstrLeiDeAprovacao.SetFocus
       Exit Function
    End If
    
    If gblnExisteCodigo(2, gstrHistoricoLogradouro, "strLeiDeAprovacao", "'" & Trim(txtstrLeiDeAprovacao.Text) & "'", _
       "intLogradouro", Str(intLogradouro)) = True Then
       ExibeMensagem "A Lei de Aprovação informada já existe no histórico."
       tab_3dPasta.Tab = 0
       txtstrLeiDeAprovacao.SetFocus
       Exit Function
    End If
    
    blnLeiOK = True
    
End Function

Private Function GravaHistorico(intLogradouro As Long) As Boolean
    Dim strSql As String
    
    With tdb_Lista
      strSql = ""
      strSql = strSql & "INSERT INTO " & gstrHistoricoLogradouro & " "
      strSql = strSql & "(intLogradouro, strDescricao, strLeiDeAprovacao, "
      strSql = strSql & "dtmDtAtualizacao, lngCodUsr, intExercicio, bitTipoLei, "
      strSql = strSql & "dtmdtAlteracao , strObservacao, strCodProcesso, intExerProcesso, "
      strSql = strSql & "bitDigProcesso, "
      strSql = strSql & "strTipoDeVia, "
      strSql = strSql & "strTipoLogradouro, "
      strSql = strSql & "strTituloLogradouro "
      strSql = strSql & ") VALUES ("
      strSql = strSql & intLogradouro & ", '"
      strSql = strSql & .Columns("Logradouro").Value & "', '"
      strSql = strSql & IIf(.Columns("strLeiDeAprovacao").Value = "", " ", .Columns("strLeiDeAprovacao").Value) & "', "
      strSql = strSql & strGETDATE & ", "
      strSql = strSql & glngCodUsr & ", "
      strSql = strSql & IIf(.Columns("intExercicio").Value = "", "NULL", .Columns("intExercicio").Value) & ", "
      strSql = strSql & IIf(.Columns("bitTipoLei").Value = "", "NULL", .Columns("bitTipoLei").Value) & ", "
      strSql = strSql & strGETDATE & ", '"
      strSql = strSql & Trim(txt_strObservacao.Text) & "', '"
      strSql = strSql & .Columns("strCodProcesso").Value & "', "
      strSql = strSql & IIf(.Columns("intExerProcesso").Value = "", "NULL", .Columns("intExerProcesso").Value) & ", "
      strSql = strSql & IIf(.Columns("bitDigProcesso").Value = "", "NULL", .Columns("bitDigProcesso").Value) & ", '"
      strSql = strSql & .Columns("Via").Value & "', '"
      strSql = strSql & .Columns("Tipo").Value & "', '"
      strSql = strSql & .Columns("Título").Value & "') "
    End With
    
    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSql) Then
       GravaHistorico = True
    End If
    
End Function

Private Sub PreencheHistorico(intLogradouro As Long)
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    Dim mobjLista As Object

    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "strDescricao, "
    strSql = strSql & "strTipoDeVia, "
    strSql = strSql & "strTipoLogradouro, "
    strSql = strSql & "strTituloLogradouro, "
    strSql = strSql & "strLeideAprovacao, "
    strSql = strSql & "intExercicio, "
    strSql = strSql & gstrCASEWHEN("bitTipoLei", "0,'Lei',1,'Decreto'") & " bitTipoLei, "
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "strCodProcesso") & strCONCAT & "'/'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "intExerProcesso") & strCONCAT & "'-'" & strCONCAT
    strSql = strSql & gstrCONVERT(CDT_VARCHAR, "bitDigProcesso") & " strProcesso , "
    strSql = strSql & "dtmdtAlteracao, "
    strSql = strSql & "strObservacao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrHistoricoLogradouro & " "
    strSql = strSql & "WHERE "
    strSql = strSql & "intLogradouro = " & intLogradouro & " "
    strSql = strSql & "ORDER BY dtmdtAlteracao DESC"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
       With adoResultado
         Do While Not adoResultado.EOF
            Set mobjLista = lvw_Historicos.ListItems.Add(, , gstrENulo(!strDescricao))
            mobjLista.SubItems(1) = gstrENulo(!strTipoDeVia)
            mobjLista.SubItems(2) = gstrENulo(!strTipoLogradouro)
            mobjLista.SubItems(3) = gstrENulo(!strTituloLogradouro)
            mobjLista.SubItems(4) = gstrENulo(!strLeiDeAprovacao)
            mobjLista.SubItems(5) = gstrENulo(!bitTipoLei)
            mobjLista.SubItems(6) = gstrENulo(!intExercicio)
            mobjLista.SubItems(7) = gstrENulo(!strProcesso)
            mobjLista.SubItems(8) = gstrENulo(!DTMDTALTERACAO)
            mobjLista.SubItems(9) = gstrENulo(!strObservacao)
            
            adoResultado.MoveNext
         Loop
       End With
    End If

    
End Sub

Sub NovoLogradouro()
    'lvw_Ceps.ListItems.Clear
    lvw_Historicos.ListItems.Clear
    txt_strObservacao.Text = ""
    TrocaCorObjeto txt_strObservacao, True
    blnObservacao = False
    'txt_Cep = ""
    'txt_NomeAnterior = ""
    'txt_Lei = ""
    tab_3dPasta.Tab = 0
    txt_dtmdtexclusao.Text = ""
    blnPertenceAoMunicipio = True
End Sub

Private Sub txt_strobservacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strObservacao, True
End Sub

Private Sub txtbitDigProcesso_GotFocus()
    MarcaCampo txtbitdigprocesso
End Sub

Private Sub txtbitDigProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitdigprocesso
End Sub

Private Sub txtintCEP_LostFocus()
    CepLogradouro txtintCep, txtstrdescricao, dbcintBairro, , , , , blnPertenceAoMunicipio, False, True
    txtintCep = gstrCEPFormatado(txtintCep)
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintExerProcesso_GotFocus()
    MarcaCampo txtintexerprocesso
End Sub

Private Sub txtintExerProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintexerprocesso
End Sub

Private Sub txtstrCodigo_GotFocus()
    MarcaCampo txtstrCodigo
    gstrProximoCodigo txtstrCodigo, gstrLogradouro, "strCodigo", gintCodSeguranca
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub

Private Sub txtstrCodProcesso_GotFocus()
    MarcaCampo txtstrcodprocesso
End Sub

Private Sub txtstrCodProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrcodprocesso
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrdescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrdescricao
End Sub

Function strQueryRelatorio() As String

'******************************************************************************************
' Data: 05/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 10/03/2003
' Alteração: - Alteração do comando SELECT devido a incompatibilidades de estrutura dos
'            outer joins entre o SQL Server e o Oracle. Os joins da cláusula FROM foram
'            substituídos por joins correspondentes na cláusula WHERE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 26/03/2003
' Alteração: - Substituição da variável strISNULL pela função gstrISNULL. Ver definição da
'            função para maiores esclarecimentos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 27/03/2003
' Alteração: - Substituição do comando CONVERT do SQL Server pela função gstrCONVERT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 14/04/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql  As String

   strSql = ""
       
 ' TIMTIM - 10/02/2003 - Pendência nº 1
 ' strSql = strSql & "SELECT CL.intCep, L.PKId, L.strCodigo, RTRIM(LTRIM(ISNULL(TL.strSigla, '') + ' ' + ISNULL(U.strDescricao,'') + ' ' + L.strDescricao)) AS Logradouro "
'   strSql = strSql & "SELECT CL.intCep, L.PKId, L.strCodigo, RTRIM(LTRIM(ISNULL(TL.strSigla, ''))) + CASE WHEN LEN(RTRIM(LTRIM(U.strDescricao))) > 0 THEN ' ' ELSE '' END + RTRIM(LTRIM(ISNULL(U.strDescricao,'')))  + ' ' + RTRIM(LTRIM(L.strDescricao)) AS Logradouro  "
'   strSql = strSql & "SELECT CL.intCep, L.PKId, L.strCodigo, RTRIM(LTRIM(" & strISNULL & "(TL.strSigla, ''))) "
   strSql = strSql & "SELECT CL.intCep, L.PKId, L.strCodigo, RTRIM(LTRIM(" & gstrISNULL("TL.strSigla", "''") & ")) " & _
        strCONCAT & gstrCASEWHEN(strLen & "(RTRIM(LTRIM(U.strDescricao)))", "0,''", "''") & strCONCAT
'        " RTRIM(LTRIM(" & strISNULL & "(U.strDescricao,'')))  " & strCONCAT & " ' ' " & strCONCAT
   strSql = strSql & " RTRIM(LTRIM(" & gstrISNULL("U.strDescricao", "''") & "))  " & strCONCAT & " ' ' " & strCONCAT
'        "RTRIM(LTRIM(L.strDescricao)) AS Logradouro  "
   strSql = strSql & "RTRIM(LTRIM(L.strDescricao)) " & " AS Logradouro  "

'   strSql = strSql & "FROM (" & gstrLogradouro & " L "
   strSql = strSql & "FROM " & gstrLogradouro & " L, "
'   strSql = strSql & "LEFT JOIN  " & gstrTituloLogradouro & " U "
   strSql = strSql & gstrTituloLogradouro & " U, "
'   strSql = strSql & "ON L.intTituloLogradouro = U.PKId) "
'   strSql = strSql & "LEFT JOIN " & gstrTipoLogradouro & " TL "
   strSql = strSql & gstrTipoLogradouro & " TL, "
'   strSql = strSql & "ON L.intTipoLogradouro = TL.PKId "
'   strSql = strSql & "LEFT JOIN " & gstrCepsLogradouro & " CL "
   strSql = strSql & gstrCepsLogradouro & " CL "
'   strSql = strSql & "ON CL.intLogradouro = L.PKId "
   strSql = strSql & "WHERE L.intTituloLogradouro " & strOUTJSQLServer & "= U.PKId" & strOUTJOracle & " AND "
   strSql = strSql & "L.intTipoLogradouro " & strOUTJSQLServer & "= TL.PKId" & strOUTJOracle & "  AND "
   strSql = strSql & "CL.intLogradouro " & strOUTJOracle & " =" & strOUTJSQLServer & " L.PKId "
       
 ' TIMTIM - 10/02/2003 - Pendência nº 1
 ' strSql = strSql & " ORDER BY Logradouro "
       
   Select Case bytOrdenacao
         
      Case Is = 1
'         strSql = strSql & "ORDER BY CONVERT(int, L.strCodigo)" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
         strSql = strSql & "ORDER BY " & gstrCONVERT(CDT_INT, "L.strCodigo") & IIf(blnOrdenacaoAsc, " ASC", " DESC")
              
      Case Is = 2, 3, 4
          strSql = strSql & "ORDER BY Logradouro" & IIf(blnOrdenacaoAsc, " ASC", " DESC")
 
   End Select
   
   strQueryRelatorio = strSql
   
End Function

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub

Private Function strQueryDataComboBairro() As String
    Dim strSql          As String
    
    strSql = "SELECT PKId, strDescricao FROM " & gstrBairro
    strSql = strSql & " WHERE bytPertenceAoMunicipo = 1"

    strQueryDataComboBairro = strSql
    
End Function

Private Function strQueryTiposDeVias() As String
Dim strSql As String
strSql = strSql & "SELECT Pkid, strSigla"
strSql = strSql & " FROM "
strSql = strSql & gstrTiposDeVias
strSql = strSql & " ORDER BY strSigla"
strQueryTiposDeVias = strSql
End Function


Private Function strQueryLogradouro() As String
Dim strSql As String
strSql = "SELECT LO.Pkid,"
strSql = strSql & gstrISNULL("RTRIM(TL.strSigla)" & strCONCAT & "' '", "' '") & strCONCAT
strSql = strSql & gstrISNULL("RTRIM(TTL.strSigla)" & strCONCAT & "' '", "' '") & strCONCAT
strSql = strSql & "LO.strDescricao AS strDescricao"
strSql = strSql & " FROM "
strSql = strSql & gstrLogradouro & " LO, "
strSql = strSql & gstrTipoLogradouro & " TL, "
strSql = strSql & gstrTituloLogradouro & " TTL"
strSql = strSql & " WHERE "
strSql = strSql & " LO.dtmdtExclusao IS NULL AND "
strSql = strSql & " LO.intTipoLogradouro " & strOUTJSQLServer & "=" & " TL.Pkid " & strOUTJOracle & " AND"
strSql = strSql & " LO.intTituloLogradouro " & strOUTJSQLServer & "=" & " TTL.Pkid " & strOUTJOracle


strSql = strSql & " ORDER BY LO.strDescricao"
strQueryLogradouro = strSql
End Function

Private Sub txtstrLeiDeAprovacao_GotFocus()
    MarcaCampo txtstrLeiDeAprovacao
End Sub

Private Sub txtstrLeiDeAprovacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrLeiDeAprovacao
End Sub

Private Sub txtstrQuadraFinal_GotFocus()
    MarcaCampo txtstrQuadraFinal
End Sub

Private Sub txtstrQuadraInicial_GotFocus()
    MarcaCampo txtstrQuadraInicial
End Sub

Private Sub txtstrSetorFinal_GotFocus()
    MarcaCampo txtstrSetorFinal
End Sub

Private Sub txtstrSetorInicial_GotFocus()
    MarcaCampo txtstrSetorInicial
End Sub

Private Function blnCamposObrigatorios() As Boolean
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & "Select * From " & gstrEmpresa
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            blnCamposObrigatorios = gstrENulo(adoResultado!bytobrigatoriologradouro)
        End If
    End If
End Function
