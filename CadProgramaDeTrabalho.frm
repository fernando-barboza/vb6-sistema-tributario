VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadProgramaDeTrabalho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programas de Trabalho"
   ClientHeight    =   6930
   ClientLeft      =   1395
   ClientTop       =   2505
   ClientWidth     =   10920
   HelpContextID   =   15
   Icon            =   "CadProgramaDeTrabalho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10920
   Begin VB.TextBox txt_dblValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8745
      TabIndex        =   92
      Top             =   5310
      Visible         =   0   'False
      Width           =   1320
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4485
      Left            =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7911
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Programas de Trabalho"
      TabPicture(0)   =   "CadProgramaDeTrabalho.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintPrograma"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintSubPrograma"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintFuncao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintOrgao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintProjetoAtividade"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblintElementoDespesa"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_CodigoReduzido"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblstrCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblintUnidadeOrcamentaria"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbldblvalor"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblintSubfuncao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblintTipoCredito"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl_Total"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblintVinculo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblintFonteRecurso"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblintGFRecurso"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl_Evento"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblintSubunidade"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblintModalidadeAplicacao"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmd_Funcao"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmd_Programa"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmd_SubPrograma"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmd_Orgao"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmd_ElementoDespesa"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtPKId"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtstrCodigo"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmd_UnidadeOrcamentaria"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtdblValor"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmd_SubFuncao"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmd_ProjetoAtividade"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmd_TipoCredito"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txt_Total"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtintCodigoReduzido"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "dbcintFuncao"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "dbcintPrograma"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "dbcintProjetoAtividade"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "dbcintUnidadeOrcamentaria"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "dbcintTipoCredito"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "dbcintSubFuncao"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "dbcintSubPrograma"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "dbcintElementoDespesa"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "dbcintOrgao"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cbointVinculo"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmd_Vinculo"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cbointFonteRecurso"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmd_FonteRecurso"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt_Grupo"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtintExercicio"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt_intUnidadeOrcamentaria"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txt_intTipoCredito"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txt_intSubFuncao"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt_intSubPrograma"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txt_intElementoDespesa"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txt_intFonteRecurso"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txt_intOrgao"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txt_intFuncao"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txt_intPrograma"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txt_intProjetoAtividade"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txt_intVinculo"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txt_codEvento"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cbointEvento"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cmd_Evento"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "chkbytSituacao"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "dbcintSubunidade"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txt_intSubunidade"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "cmd_Subunidade"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txt_intModalidadeAplicacao"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "dbcintModalidade"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "cmd_ModalidadeAplicacao"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "frm_Integrante"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txt_intConvenio"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).ControlCount=   71
      TabCaption(1)   =   "Saldos"
      TabPicture(1)   =   "CadProgramaDeTrabalho.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(8)=   "Label9"
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(10)=   "txt_saldoini"
      Tab(1).Control(11)=   "txt_empenhado"
      Tab(1).Control(12)=   "txt_liquidado"
      Tab(1).Control(13)=   "txt_Pago"
      Tab(1).Control(14)=   "txt_anulado"
      Tab(1).Control(15)=   "txt_suplementado"
      Tab(1).Control(16)=   "txt_Reservado"
      Tab(1).Control(17)=   "txt_SaldoParaEmpenho"
      Tab(1).Control(18)=   "txt_strCodigoSaldo"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txt_TotalBloqueado"
      Tab(1).ControlCount=   20
      Begin VB.TextBox txt_intConvenio 
         Height          =   315
         Left            =   6900
         MaxLength       =   2
         TabIndex        =   59
         Top             =   3180
         Width           =   525
      End
      Begin VB.Frame frm_Integrante 
         Caption         =   "Integrante"
         Height          =   555
         Left            =   1050
         TabIndex        =   67
         Top             =   3840
         Width           =   4095
         Begin VB.CheckBox chkbytEducacao 
            Caption         =   "Educação"
            Height          =   195
            Left            =   2040
            TabIndex        =   69
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox chkbytSaude 
            Caption         =   "Saúde"
            Height          =   255
            Left            =   600
            TabIndex        =   68
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmd_ModalidadeAplicacao 
         Height          =   315
         Left            =   10380
         Picture         =   "CadProgramaDeTrabalho.frx":107A
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Tag             =   "200"
         ToolTipText     =   "Clique aqui para cadastar elemento de despesa"
         Top             =   3195
         Width           =   360
      End
      Begin VB.ComboBox dbcintModalidade 
         Height          =   315
         Left            =   7440
         Sorted          =   -1  'True
         TabIndex        =   60
         Top             =   3180
         Width           =   2925
      End
      Begin VB.TextBox txt_intModalidadeAplicacao 
         Height          =   315
         Left            =   6360
         MaxLength       =   3
         TabIndex        =   58
         Top             =   3180
         Width           =   525
      End
      Begin VB.CommandButton cmd_Subunidade 
         Height          =   315
         Left            =   4770
         Picture         =   "CadProgramaDeTrabalho.frx":1404
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "197"
         ToolTipText     =   "Clique aqui para cadastar subunidade"
         Top             =   1410
         Width           =   360
      End
      Begin VB.TextBox txt_intSubunidade 
         Height          =   315
         Left            =   1065
         TabIndex        =   14
         Top             =   1410
         Width           =   825
      End
      Begin VB.ComboBox dbcintSubunidade 
         Height          =   315
         Left            =   1890
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   1410
         Width           =   2865
      End
      Begin VB.TextBox txt_TotalBloqueado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68100
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   94
         Top             =   2745
         Width           =   3165
      End
      Begin VB.TextBox txt_strCodigoSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   -73185
         LinkItem        =   "0"
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   90
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   540
         Width           =   5115
      End
      Begin VB.TextBox txt_SaldoParaEmpenho 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68100
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   88
         Top             =   3330
         Width           =   3165
      End
      Begin VB.TextBox txt_Reservado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68100
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   86
         Top             =   2160
         Width           =   3165
      End
      Begin VB.TextBox txt_suplementado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68100
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   84
         Top             =   1575
         Width           =   3165
      End
      Begin VB.TextBox txt_anulado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68100
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   82
         Top             =   990
         Width           =   3165
      End
      Begin VB.TextBox txt_Pago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73170
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   80
         Top             =   2745
         Width           =   3165
      End
      Begin VB.TextBox txt_liquidado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73170
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   78
         Top             =   2160
         Width           =   3165
      End
      Begin VB.TextBox txt_empenhado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73170
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   76
         Top             =   1575
         Width           =   3165
      End
      Begin VB.TextBox txt_saldoini 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73170
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   74
         Top             =   990
         Width           =   3165
      End
      Begin VB.CheckBox chkbytSituacao 
         Caption         =   "Situacao"
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   6390
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   73
         Top             =   390
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_Evento 
         Height          =   315
         Left            =   10380
         Picture         =   "CadProgramaDeTrabalho.frx":178E
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Tag             =   "247"
         ToolTipText     =   "Clique para cadastar convênio"
         Top             =   2490
         Width           =   360
      End
      Begin VB.ComboBox cbointEvento 
         Height          =   315
         Left            =   7440
         TabIndex        =   43
         Top             =   2490
         Width           =   2925
      End
      Begin VB.TextBox txt_codEvento 
         Height          =   315
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   42
         Top             =   2490
         Width           =   1065
      End
      Begin VB.TextBox txt_intVinculo 
         Height          =   315
         Left            =   1065
         TabIndex        =   46
         Top             =   2850
         Width           =   825
      End
      Begin VB.TextBox txt_intProjetoAtividade 
         Height          =   315
         Left            =   1065
         TabIndex        =   38
         Top             =   2490
         Width           =   825
      End
      Begin VB.TextBox txt_intPrograma 
         Height          =   315
         Left            =   1065
         TabIndex        =   30
         Top             =   2130
         Width           =   825
      End
      Begin VB.TextBox txt_intFuncao 
         Height          =   315
         Left            =   1065
         TabIndex        =   22
         Top             =   1770
         Width           =   825
      End
      Begin VB.TextBox txt_intOrgao 
         Height          =   315
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1050
         Width           =   825
      End
      Begin VB.TextBox txt_intFonteRecurso 
         Height          =   315
         Left            =   1065
         TabIndex        =   54
         Top             =   3180
         Width           =   825
      End
      Begin VB.TextBox txt_intElementoDespesa 
         Height          =   315
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   50
         Top             =   2835
         Width           =   1065
      End
      Begin VB.TextBox txt_intSubPrograma 
         Height          =   315
         Left            =   6360
         TabIndex        =   34
         Top             =   2130
         Width           =   1065
      End
      Begin VB.TextBox txt_intSubFuncao 
         Height          =   315
         Left            =   6360
         TabIndex        =   26
         Top             =   1770
         Width           =   1065
      End
      Begin VB.TextBox txt_intTipoCredito 
         Height          =   315
         Left            =   6360
         TabIndex        =   18
         Top             =   1410
         Width           =   1065
      End
      Begin VB.TextBox txt_intUnidadeOrcamentaria 
         Height          =   315
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1050
         Width           =   1065
      End
      Begin VB.TextBox txtintExercicio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   7980
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   72
         Top             =   330
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txt_Grupo 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   1050
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   63
         Top             =   3540
         Width           =   4035
      End
      Begin VB.CommandButton cmd_FonteRecurso 
         Height          =   315
         Left            =   4770
         Picture         =   "CadProgramaDeTrabalho.frx":1B18
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Tag             =   "207"
         ToolTipText     =   "Clique aqui para cadastar fonte de recurso"
         Top             =   3180
         Width           =   360
      End
      Begin VB.ComboBox cbointFonteRecurso 
         Height          =   315
         Left            =   1890
         Sorted          =   -1  'True
         TabIndex        =   55
         Top             =   3180
         Width           =   2865
      End
      Begin VB.CommandButton cmd_Vinculo 
         Height          =   315
         Left            =   4770
         Picture         =   "CadProgramaDeTrabalho.frx":1EA2
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Tag             =   "211"
         ToolTipText     =   "Clique aqui para cadastar vínculo"
         Top             =   2850
         Width           =   360
      End
      Begin VB.ComboBox cbointVinculo 
         Height          =   315
         Left            =   1890
         Sorted          =   -1  'True
         TabIndex        =   47
         Top             =   2850
         Width           =   2865
      End
      Begin VB.ComboBox dbcintOrgao 
         Height          =   315
         Left            =   1890
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1050
         Width           =   2865
      End
      Begin VB.ComboBox dbcintElementoDespesa 
         Height          =   315
         Left            =   7440
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   2835
         Width           =   2925
      End
      Begin VB.ComboBox dbcintSubPrograma 
         Height          =   315
         Left            =   7440
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   2130
         Width           =   2925
      End
      Begin VB.ComboBox dbcintSubFuncao 
         Height          =   315
         Left            =   7440
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   1770
         Width           =   2925
      End
      Begin VB.ComboBox dbcintTipoCredito 
         Height          =   315
         Left            =   7440
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   1410
         Width           =   2925
      End
      Begin VB.ComboBox dbcintUnidadeOrcamentaria 
         Height          =   315
         Left            =   7440
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1050
         Width           =   2925
      End
      Begin VB.ComboBox dbcintProjetoAtividade 
         Height          =   315
         Left            =   1890
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   2490
         Width           =   2865
      End
      Begin VB.ComboBox dbcintPrograma 
         Height          =   315
         Left            =   1890
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   2130
         Width           =   2865
      End
      Begin VB.ComboBox dbcintFuncao 
         Height          =   315
         Left            =   1890
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   1770
         Width           =   2865
      End
      Begin VB.TextBox txtintCodigoReduzido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1065
         MultiLine       =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   690
         Width           =   1005
      End
      Begin VB.TextBox txt_Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   8880
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   70
         Top             =   3540
         Width           =   1545
      End
      Begin VB.CommandButton cmd_TipoCredito 
         Height          =   315
         Left            =   10380
         Picture         =   "CadProgramaDeTrabalho.frx":222C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Clique aqui para cadastar tipo de crédito"
         Top             =   1410
         Width           =   360
      End
      Begin VB.CommandButton cmd_ProjetoAtividade 
         Height          =   315
         Left            =   4770
         Picture         =   "CadProgramaDeTrabalho.frx":25B6
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Tag             =   "185"
         ToolTipText     =   "Clique aqui para cadastar projeto/atividade"
         Top             =   2490
         Width           =   360
      End
      Begin VB.CommandButton cmd_SubFuncao 
         Height          =   315
         Left            =   10380
         Picture         =   "CadProgramaDeTrabalho.frx":2940
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "199"
         ToolTipText     =   "Clique aqui para cadastar subfunção do governo"
         Top             =   1785
         Width           =   360
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   3510
         Width           =   1545
      End
      Begin VB.CommandButton cmd_UnidadeOrcamentaria 
         Height          =   315
         Left            =   10380
         Picture         =   "CadProgramaDeTrabalho.frx":2CCA
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "194"
         ToolTipText     =   "Clique aqui para cadastar unidade orçamentária"
         Top             =   1050
         Width           =   360
      End
      Begin VB.TextBox txtstrCodigo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   6360
         LinkItem        =   "0"
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   690
         Width           =   4035
      End
      Begin VB.TextBox txtPKId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8850
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmd_ElementoDespesa 
         Height          =   315
         Left            =   10380
         Picture         =   "CadProgramaDeTrabalho.frx":3054
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Tag             =   "200"
         ToolTipText     =   "Clique aqui para cadastar elemento de despesa"
         Top             =   2850
         Width           =   360
      End
      Begin VB.CommandButton cmd_Orgao 
         Height          =   315
         Left            =   4770
         Picture         =   "CadProgramaDeTrabalho.frx":33DE
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "193"
         ToolTipText     =   "Clique aqui para cadastar órgão"
         Top             =   1050
         Width           =   360
      End
      Begin VB.CommandButton cmd_SubPrograma 
         Height          =   315
         Left            =   10380
         Picture         =   "CadProgramaDeTrabalho.frx":3768
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Tag             =   "202"
         ToolTipText     =   "Clique aqui para cadastar subprograma "
         Top             =   2130
         Width           =   360
      End
      Begin VB.CommandButton cmd_Programa 
         Height          =   315
         Left            =   4770
         Picture         =   "CadProgramaDeTrabalho.frx":3AF2
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "186"
         ToolTipText     =   "Clique aqui para cadastar programa de governo"
         Top             =   2130
         Width           =   360
      End
      Begin VB.CommandButton cmd_Funcao 
         Height          =   315
         Left            =   4770
         Picture         =   "CadProgramaDeTrabalho.frx":3E7C
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "198"
         ToolTipText     =   "Clique aqui para cadastar função"
         Top             =   1785
         Width           =   360
      End
      Begin VB.Label lblintModalidadeAplicacao 
         AutoSize        =   -1  'True
         Caption         =   "Mod. Aplicação"
         Height          =   195
         Left            =   5190
         TabIndex        =   57
         ToolTipText     =   "Elemento da despesa"
         Top             =   3225
         Width           =   1110
      End
      Begin VB.Label lblintSubunidade 
         AutoSize        =   -1  'True
         Caption         =   "Subunid."
         Height          =   195
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "Unidade orçamentária"
         Top             =   1470
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total Bloqueado"
         Height          =   180
         Left            =   -69345
         TabIndex        =   95
         Top             =   2797
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Funcional Programática"
         Height          =   195
         Left            =   -74910
         TabIndex        =   91
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Saldo para Empenho"
         Height          =   195
         Left            =   -69660
         TabIndex        =   89
         Top             =   3375
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Reservado"
         Height          =   195
         Left            =   -68955
         TabIndex        =   87
         Top             =   2205
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Suplementado"
         Height          =   195
         Left            =   -69195
         TabIndex        =   85
         Top             =   1620
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Anulado"
         Height          =   195
         Left            =   -68760
         TabIndex        =   83
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pago"
         Height          =   195
         Left            =   -73620
         TabIndex        =   81
         Top             =   2790
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Liquidação"
         Height          =   195
         Left            =   -74025
         TabIndex        =   79
         Top             =   2205
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Empenhado"
         Height          =   195
         Left            =   -74505
         TabIndex        =   77
         Top             =   1620
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dotação inicial"
         Height          =   195
         Left            =   -74295
         TabIndex        =   75
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label lbl_Evento 
         AutoSize        =   -1  'True
         Caption         =   "Evento Contábil"
         Height          =   195
         Left            =   5175
         TabIndex        =   41
         Top             =   2550
         Width           =   1125
      End
      Begin VB.Label lblintGFRecurso 
         AutoSize        =   -1  'True
         Caption         =   "G.F. Recurso"
         Height          =   195
         Left            =   45
         TabIndex        =   62
         ToolTipText     =   "Grupo da Fonte de Recurso"
         Top             =   3570
         Width           =   945
      End
      Begin VB.Label lblintFonteRecurso 
         AutoSize        =   -1  'True
         Caption         =   "F. Recurso"
         Height          =   195
         Left            =   210
         TabIndex        =   53
         ToolTipText     =   "Fonte de Recurso"
         Top             =   3240
         Width           =   780
      End
      Begin VB.Label lblintVinculo 
         AutoSize        =   -1  'True
         Caption         =   "Vínculo"
         Height          =   195
         Left            =   435
         TabIndex        =   45
         Top             =   2910
         Width           =   555
      End
      Begin VB.Label lbl_Total 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   8430
         TabIndex        =   66
         Top             =   3570
         Width           =   360
      End
      Begin VB.Label lblintTipoCredito 
         AutoSize        =   -1  'True
         Caption         =   "Tipo do Crédito"
         Height          =   195
         Left            =   5220
         TabIndex        =   17
         Top             =   1470
         Width           =   1080
      End
      Begin VB.Label lblintSubfuncao 
         AutoSize        =   -1  'True
         Caption         =   "Subfunções"
         Height          =   195
         Left            =   5445
         TabIndex        =   25
         ToolTipText     =   "Subfunções de governo"
         Top             =   1815
         Width           =   855
      End
      Begin VB.Label lbldblvalor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   5940
         TabIndex        =   64
         Top             =   3570
         Width           =   360
      End
      Begin VB.Label lblintUnidadeOrcamentaria 
         AutoSize        =   -1  'True
         Caption         =   "U.Orçamentária"
         Height          =   195
         Left            =   5190
         TabIndex        =   9
         ToolTipText     =   "Unidade orçamentária"
         Top             =   1110
         Width           =   1110
      End
      Begin VB.Label lblstrCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Funcional Programática"
         Height          =   195
         Left            =   4635
         TabIndex        =   3
         Top             =   750
         Width           =   1665
      End
      Begin VB.Label lbl_CodigoReduzido 
         AutoSize        =   -1  'True
         Caption         =   "Reduzido"
         Height          =   195
         Left            =   315
         TabIndex        =   1
         Top             =   750
         Width           =   675
      End
      Begin VB.Label lblintElementoDespesa 
         AutoSize        =   -1  'True
         Caption         =   "E.da Despesa"
         Height          =   195
         Left            =   5295
         TabIndex        =   49
         ToolTipText     =   "Elemento da despesa"
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label lblintProjetoAtividade 
         AutoSize        =   -1  'True
         Caption         =   "Proj/Ativ"
         Height          =   195
         Left            =   375
         TabIndex        =   37
         ToolTipText     =   "Projeto/Atividade"
         Top             =   2550
         Width           =   615
      End
      Begin VB.Label lblintOrgao 
         AutoSize        =   -1  'True
         Caption         =   "Orgão"
         Height          =   195
         Left            =   555
         TabIndex        =   5
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label lblintFuncao 
         AutoSize        =   -1  'True
         Caption         =   "Funções"
         Height          =   195
         Left            =   375
         TabIndex        =   21
         ToolTipText     =   "Funções de governo"
         Top             =   1815
         Width           =   615
      End
      Begin VB.Label lblintSubPrograma 
         AutoSize        =   -1  'True
         Caption         =   "Subprograma"
         Height          =   195
         Left            =   5355
         TabIndex        =   33
         Top             =   2175
         Width           =   945
      End
      Begin VB.Label lblintPrograma 
         AutoSize        =   -1  'True
         Caption         =   "Programa"
         Height          =   195
         Left            =   315
         TabIndex        =   29
         Top             =   2175
         Width           =   675
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   2175
      Left            =   60
      Negotiate       =   -1  'True
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   4620
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   3836
      _LayoutType     =   4
      _RowHeight      =   13
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Pkid"
      Columns(0).DataField=   "PKID"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Cod.Reduzido"
      Columns(1).DataField=   "intCodigoReduzido"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Órgão"
      Columns(2).DataField=   "Orgao"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "U.Orçamentária"
      Columns(3).DataField=   "UnidadeOrcamentaria"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Subunidade"
      Columns(4).DataField=   "Subunidade"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Proj/Ativ"
      Columns(5).DataField=   "ProjetoAtividade"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Cod.E.Despesa"
      Columns(6).DataField=   "strCodigoElementoDespesa"
      Columns(6).NumberFormat=   "FormatText Event"
      Columns(6).EditMaskUpdate=   -1  'True
      Columns(6).EditMaskRight=   -1  'True
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Elemento de Despesa"
      Columns(7).DataField=   "Elemento"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "F. de Recurso"
      Columns(8).DataField=   "FonteRecurso"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Valor"
      Columns(9).DataField=   "Valor"
      Columns(9).NumberFormat=   "FormatText Event"
      Columns(9).EditMaskUpdate=   -1  'True
      Columns(9).EditMaskRight=   -1  'True
      Columns(9).ConvertEmptyCell=   1
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Total"
      Columns(10).DataField=   "dblTotal"
      Columns(10).NumberFormat=   "Standard"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   13160664
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1931"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1852"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=260"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=953"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=873"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=260"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2143"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2064"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=260"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1667"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1588"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=260"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=1296"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1217"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=260"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=1349"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1270"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=4921"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=4842"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=260"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(49)=   "Column(8).Width=1905"
      Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1826"
      Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(54)=   "Column(9).Width=2170"
      Splits(0)._ColumnProps(55)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(9)._WidthInPix=2090"
      Splits(0)._ColumnProps(57)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(58)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(59)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(60)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(61)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(63)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(64)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(65)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      EditDropDown    =   0   'False
      HeadLines       =   1
      FootLines       =   1
      MarqueeUnique   =   0   'False
      TabAction       =   2
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTips        =   2
      CellTipsWidth   =   0
      InsertMode      =   0   'False
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=3,.bold=0,.fontsize=825"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=87,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=88,.alignment=0"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=118,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=115,.parent=88,.alignment=0"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=116,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=117,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=122,.parent=87"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=119,.parent=88,.alignment=0"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=120,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=121,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=126,.parent=87"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=123,.parent=88,.alignment=0"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=124,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=125,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=142,.parent=87"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=139,.parent=88,.alignment=0"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=140,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=141,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=20,.parent=87"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=146,.parent=87"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=143,.parent=88,.alignment=0"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=144,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=145,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=28,.parent=87"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=150,.parent=87,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=147,.parent=88,.alignment=2"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=148,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=149,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=24,.parent=87"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=21,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=22,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=23,.parent=91"
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
Attribute VB_Name = "frmCadProgramaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnProgramaAprovado    As Boolean
Dim mblnAlterando           As Boolean
Dim mblnClickOk             As Boolean
Dim mblnselecionou          As Boolean
Dim mobjAux                 As Object
Dim intPkIdRow              As Variant
Dim blnFilterBar            As Boolean
Dim dblValorAnt             As Double
Public blnOrcamento         As Boolean
Dim strChamadaForm          As String
Dim intIndiceCombo          As Integer

Dim strElementoDespesa  As String

Public mIntCodSeguranca     As Integer

Private Sub LeGrupo()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT GF.strDescricao FROM "
    strSQL = strSQL & gstrGrupoDeFonteRecurso & " GF, "
    strSQL = strSQL & gstrFonteRecurso & " FR "
    strSQL = strSQL & "WHERE GF.PKId = FR.intGrupo "
    strSQL = strSQL & "AND FR.PKId = " & gstrItemData(cbointFonteRecurso)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txt_Grupo = gstrENulo(!strDescricao)
            Else
                txt_Grupo = ""
            End If
        End With
    End If
    adoResultado.Close
End Sub

Private Sub FormataCodigoTrabalho()
    
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim ncount As Integer
    
    txtstrCodigo = ""
    strSQL = ""
    
    strSQL = strSQL & gstrStoredProcedure("sp_CodigoProgTrabalho", _
    gstrItemData(dbcintOrgao) & ", " & gstrItemData(dbcintUnidadeOrcamentaria) & ", " & _
    gstrItemData(dbcintFuncao) & ", " & gstrItemData(dbcintSubFuncao) & ", " & _
    gstrItemData(dbcintPrograma) & ", " & gstrItemData(dbcintSubPrograma) & ", " & _
    gstrItemData(dbcintProjetoAtividade) & ", " & gstrItemData(dbcintElementoDespesa) & ", " & _
    gstrItemData(dbcintSubunidade) & ", " & gstrItemData(cbointEvento) & _
    ", " & gstrItemData(cbointFonteRecurso), True)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                If !STRORGAO > 0 Then
                    'txtstrCodigo = txtstrCodigo & Format(gstrValorSemMascara(!STRORGAO), "00") & "."
                    txtstrCodigo = txtstrCodigo & gstrValorSemMascara(!STRORGAO) & "."
                    txt_intOrgao.Text = gstrValorSemMascara(!STRORGAO)
                Else
                    txt_intOrgao.Text = Space$(0)
                End If
                If !strUnidadeOrcamentaria > 0 Then
                    'txtstrCodigo = txtstrCodigo & Format(gstrValorSemMascara(!strUnidadeOrcamentaria), "0000") & "."
                    txtstrCodigo = txtstrCodigo & gstrValorSemMascara(!strUnidadeOrcamentaria) & "."
                    txt_intUnidadeOrcamentaria.Text = gstrValorSemMascara(!strUnidadeOrcamentaria)
                Else
                    txt_intUnidadeOrcamentaria.Text = Space$(0)
                End If
                If !strSubUnidade > 0 Then
                    'txtstrCodigo = txtstrCodigo & Format(gstrValorSemMascara(!strSubunidade), "0000") & "."
                    txtstrCodigo = txtstrCodigo & gstrValorSemMascara(!strSubUnidade) & "."
                    txt_intSubunidade.Text = gstrValorSemMascara(!strSubUnidade)
                Else
                    
                    LeDaTabelaParaObj gstrSubUnidade, dbcintSubunidade, "SELECT Pkid, strDescricao FROM " & gstrSubUnidade & " where intOrgao = '" & gstrItemData(dbcintOrgao) & "' and intUnidadeOrcamentaria = '" & gstrItemData(dbcintUnidadeOrcamentaria) & "'"
                    
                    dbcintSubunidade.Text = ""
                    For ncount = 0 To dbcintSubunidade.ListCount - 1
                        If Trim(txt_intSubunidade) = dbcintSubunidade.ItemData(ncount) Then
                            dbcintSubunidade = dbcintSubunidade.list(ncount)
                        End If
                    Next
                    If Len(Trim(dbcintSubunidade)) = 0 Then txt_intSubunidade.Text = Space$(0)
                End If
                If !strFuncao > 0 Then
                    'txtstrCodigo = txtstrCodigo & Format(gstrValorSemMascara(!strFuncao), "00") & "."
                    txtstrCodigo = txtstrCodigo & gstrValorSemMascara(!strFuncao) & "."
                    txt_intFuncao.Text = gstrValorSemMascara(!strFuncao)
                Else
                    txt_intFuncao.Text = Space$(0)
                End If
                If !strSubFuncao > 0 Then
                    'txtstrCodigo = txtstrCodigo & Format(gstrValorSemMascara(!strSubFuncao), "000") & "."
                    txtstrCodigo = txtstrCodigo & gstrValorSemMascara(!strSubFuncao) & "."
                    txt_intSubFuncao.Text = gstrValorSemMascara(!strSubFuncao)
                Else
                    txt_intSubFuncao.Text = Space$(0)
                End If
                If !strPrograma > 0 Then
                    'txtstrCodigo = txtstrCodigo & Format(gstrValorSemMascara(!strPrograma), "000") & "."
                    txtstrCodigo = txtstrCodigo & gstrValorSemMascara(!strPrograma) & "."
                    txt_intPrograma.Text = gstrValorSemMascara(!strPrograma)
                Else
                    txt_intPrograma.Text = Space$(0)
                End If
                If !strSubprograma > 0 Then
                    'txtstrCodigo = txtstrCodigo & Format(gstrValorSemMascara(!strSubprograma), "0000") & "."
                    txtstrCodigo = txtstrCodigo & gstrValorSemMascara(!strSubprograma) & "."
                    txt_intSubPrograma.Text = gstrValorSemMascara(!strSubprograma)
                Else
                    txt_intSubPrograma.Text = Space$(0)
                End If
                If !STRPROJETO > 0 Then
                    'txtstrCodigo = txtstrCodigo & Format(gstrValorSemMascara(!STRPROJETO), "0000") & "."
                    txtstrCodigo = txtstrCodigo & gstrValorSemMascara(!STRPROJETO) & "."
                    txt_intProjetoAtividade.Text = gstrValorSemMascara(!STRPROJETO)
                Else
                    txt_intProjetoAtividade.Text = Space$(0)
                End If
                If !strCodigoElementoDespesa > 0 Then
                    txtstrCodigo = txtstrCodigo & _
                    gstrValorSemMascara(gvntFormatacaoEspecifica(!strCodigoElementoDespesa)) & "."
                    'Format(gstrValorSemMascara(gvntFormatacaoEspecifica(!strCodigoElementoDespesa)), "00000000") & "."
                    txt_intElementoDespesa.Text = gstrValorSemMascara(!strCodigoElementoDespesa)
                Else
                    txt_intElementoDespesa.Text = Space$(0)
                End If
                If !strEvento > 0 Then
                    txt_codEvento.Text = gstrValorSemMascara(!strEvento)
                Else
                    txt_codEvento.Text = Space$(0)
                End If
                
                If adoResultado!STRFONTERECURSO <> "0" And adoResultado!STRFONTERECURSO <> "" Then 'Orc1599 - Fernando - Inclusão de "adoResultado!STRFONTERECURSO <> "0""
                'txtstrCodigo = txtstrCodigo & Format(gstrValorSemMascara(!STRFONTERECURSO), "0000")
                txtstrCodigo = txtstrCodigo & gstrValorSemMascara(!STRFONTERECURSO)
                txt_intFonteRecurso.Text = gstrValorSemMascara(!STRFONTERECURSO)
            Else
                txt_intFonteRecurso.Text = Space$(0)
            End If
            
        End If
    End With
End If
End Sub

Private Sub cbointFonteRecurso_Click()
    
    LeGrupo
    
    If cbointFonteRecurso.ListIndex > -1 Then
        txt_intFonteRecurso.Text = RetornaCodigo("strCodigo", "PkID", gstrFonteRecurso, cbointFonteRecurso.ItemData(cbointFonteRecurso.ListIndex), " AND intExercicio = " & txtintExercicio)
        With Me.ActiveControl
            If .Name = "cbointFonteRecurso" Or Mid(.Name, 1, 4) = "cmd_" Then
                FormataCodigoTrabalho
            End If
        End With
    Else
        txt_intFonteRecurso.Text = Space$(0)
    End If
    
End Sub

Private Sub cbointFonteRecurso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointFonteRecurso_LostFocus()
    If cbointFonteRecurso.ListIndex = -1 Then txt_intFonteRecurso = Space$(0)
End Sub

Private Sub cbointVinculo_Click()
    If cbointVinculo.ListIndex > -1 Then
        txt_intVinculo.Text = RetornaCodigo("strCodigo", "strDescricao", gstrVinculo, cbointVinculo.Text)
    Else
        txt_intVinculo.Text = Space$(0)
    End If
End Sub

Private Sub cbointVinculo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointVinculo_LostFocus()
    If cbointVinculo.ListIndex = -1 Then txt_intVinculo = Space$(0)
End Sub

Private Sub cmd_Evento_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadEvento, cbointEvento, strQueryEvento
    If Not blnOrcamento Then frmCadEvento.Caption = frmCadEvento.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_FonteRecurso_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadFonteRecurso, cbointFonteRecurso
    If Not blnOrcamento Then frmCadFonteRecurso.Caption = frmCadFonteRecurso.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_ModalidadeAplicacao_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadModalidadeAplicacao, dbcintModalidade
    If Not blnOrcamento Then frmCadModalidadeAplicacao.Caption = frmCadModalidadeAplicacao.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_Vinculo_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadVinculo, cbointVinculo
    If Not blnOrcamento Then frmCadVinculo.Caption = frmCadVinculo.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub dbcintElementoDespesa_Click()
    With Me.ActiveControl
        If .Name = "dbcintElementoDespesa" Or Mid(.Name, 1, 4) = "cmd_" Then
            FormataCodigoTrabalho
        End If
    End With
End Sub

Private Sub dbcintElementoDespesa_GotFocus()
    If dbcintElementoDespesa.ListIndex = -1 Then
        cbointEvento_LostFocus
    End If
End Sub

Private Sub dbcintElementoDespesa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintElementoDespesa_LostFocus()
    If dbcintElementoDespesa.ListIndex = -1 Then
        txt_intElementoDespesa = Space$(0)
        FormataCodigoTrabalho
    End If
End Sub

Private Sub cbointEvento_Click()
    FormataCodigoTrabalho
End Sub

Private Sub cbointEvento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cbointEvento_LostFocus()
    Dim strSQL As String
    
    If cbointEvento.ListIndex = -1 Then
        txt_codEvento = Space$(0)
        FormataCodigoTrabalho
    Else
        strElementoDespesa = txt_intElementoDespesa
        If gintExercicio > 2006 Or Not blnOrcamento Then
            LeDaTabelaParaObj gstrElementoDespesa, dbcintElementoDespesa, "SELECT PKId, strDescricao FROM " & gstrElementoDespesa & " WHERE intExercicio=" & IIf(blnOrcamento = True, gintExercicio, gintExercicio + 1)
        Else
            LeDaTabelaParaObj gstrElementoDespesa, dbcintElementoDespesa, "SELECT PKId, strDescricao FROM " & gstrElementoDespesa & " WHERE intExercicio=" & IIf(blnOrcamento = True, gintExercicio, gintExercicio + 1) & _
            " AND " & strSUBSTRING & "(strCodigoElementoDespesa,1," & Len(BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoDespesa, "C", 0)) & ") = '" & _
            BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoDespesa, "C", 0) & "'"
        End If
        txt_intElementoDespesa = strElementoDespesa
        txt_intElementoDespesa_LostFocus
    End If
    
End Sub

Private Sub dbcintFuncao_Click()
    With Me.ActiveControl
        If .Name = "dbcintFuncao" Or Mid(.Name, 1, 4) = "cmd_" Then
            FormataCodigoTrabalho
        End If
    End With
End Sub

Private Sub dbcintFuncao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintFuncao_LostFocus()
    If dbcintFuncao.ListIndex = -1 Then
        txt_intFuncao = Space$(0)
        FormataCodigoTrabalho
    End If
End Sub

Private Sub dbcintModalidade_Click()
    If dbcintModalidade.ListIndex > -1 Then
        txt_intModalidadeAplicacao.Text = Format(RetornaCodigo("strCodigo", "strDescricao", gstrModalidade, dbcintModalidade.Text), "000")
        txt_intConvenio.Text = CarregaConvenio(dbcintModalidade.ItemData(dbcintModalidade.ListIndex))
    Else
        txt_intModalidadeAplicacao.Text = Space$(0)
    End If
End Sub

Private Sub dbcintModalidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintModalidade_LostFocus()
    If dbcintModalidade.ListIndex = -1 Then
        txt_intModalidadeAplicacao = Space$(0)
        txt_intConvenio.Text = Space$(0)
    End If
End Sub

Private Sub dbcintOrgao_Click()
    'LimpaObjetos
    LeDaTabelaParaObj gstrUnidadeOrcamentaria, dbcintUnidadeOrcamentaria, strQueryUO
    dbcintSubunidade.Clear
    With Me.ActiveControl
        If .Name = "dbcintOrgao" Or Mid(.Name, 1, 4) = "cmd_" Then
            FormataCodigoTrabalho
        End If
    End With
    
End Sub

Private Function strQueryUO() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao FROM "
    strSQL = strSQL & gstrUnidadeOrcamentaria
    strSQL = strSQL & " WHERE intOrgao = '" & gstrItemData(dbcintOrgao) & "'"
    strQueryUO = strSQL
End Function

Private Function strQuerySubunidade() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao FROM "
    strSQL = strSQL & gstrSubUnidade & " "
    strSQL = strSQL & "WHERE intUnidadeOrcamentaria = " & gstrItemData(dbcintUnidadeOrcamentaria)
    strQuerySubunidade = strSQL
End Function

Private Sub dbcintOrgao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintOrgao_LostFocus()
    If dbcintOrgao.ListIndex = -1 Then
        txt_intOrgao = Space$(0)
        FormataCodigoTrabalho
    End If
End Sub

Private Sub dbcintPrograma_Click()
    LeDaTabelaParaObj gstrSubPrograma, dbcintSubPrograma, strQuerySubprograma
    With Me.ActiveControl
        If .Name = "dbcintPrograma" Or Mid(.Name, 1, 4) = "cmd_" Then
            FormataCodigoTrabalho
        End If
    End With
End Sub

Private Function strQuerySubprograma() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao FROM "
    strSQL = strSQL & gstrSubPrograma & " "
    strSQL = strSQL & "WHERE intPrograma = " & gstrItemData(dbcintPrograma)
    strQuerySubprograma = strSQL
End Function

Private Sub dbcintPrograma_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintPrograma_LostFocus()
    If dbcintPrograma.ListIndex = -1 Then
        txt_intPrograma = Space$(0)
        FormataCodigoTrabalho
    End If
End Sub

Private Sub dbcintProjetoAtividade_Click()
    With Me.ActiveControl
        If .Name = "dbcintProjetoAtividade" Or Mid(.Name, 1, 4) = "cmd_" Then
            FormataCodigoTrabalho
        End If
    End With
End Sub

Private Sub dbcintProjetoAtividade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintProjetoAtividade_LostFocus()
    If dbcintProjetoAtividade.ListIndex = -1 Then
        txt_intProjetoAtividade = Space$(0)
        FormataCodigoTrabalho
    End If
End Sub

Private Sub dbcintSubFuncao_Click()
    With Me.ActiveControl
        If .Name = "dbcintSubFuncao" Or Mid(.Name, 1, 4) = "cmd_" Then
            FormataCodigoTrabalho
        End If
    End With
End Sub

Private Sub dbcintSubFuncao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintSubFuncao_LostFocus()
    If dbcintSubFuncao.ListIndex = -1 Then
        txt_intSubFuncao = Space$(0)
        FormataCodigoTrabalho
    End If
End Sub

Private Sub dbcintSubPrograma_Click()
    With Me.ActiveControl
        If .Name = "dbcintSubPrograma" Or Mid(.Name, 1, 4) = "cmd_" Then
            FormataCodigoTrabalho
        End If
    End With
End Sub

Private Sub dbcintSubPrograma_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintSubPrograma_LostFocus()
    If dbcintSubPrograma.ListIndex = -1 Then
        txt_intSubPrograma = Space$(0)
        FormataCodigoTrabalho
    End If
End Sub

Private Sub dbcintSubunidade_Click()
    With Me.ActiveControl
        If .Name = "dbcintSubunidade" Or Mid(.Name, 1, 4) = "cmd_" Then
            FormataCodigoTrabalho
        End If
    End With
End Sub

Private Sub dbcintSubunidade_GotFocus()
    '    If Len(Trim(txt_intOrgao)) <> 0 And Len(Trim(txt_intUnidadeOrcamentaria)) <> 0 Then
    '        LeDaTabelaParaObj gstrSubUnidade, dbcintSubunidade, "SELECT PKId, strDescricao FROM " & gstrSubUnidade & " where intOrgao = '" & gstrItemData(dbcintOrgao) & "' and intUnidadeOrcamentaria = '" & gstrItemData(dbcintUnidadeOrcamentaria) & "'"
    '       ' LeDaTabelaParaObj gstrSubUnidade, dbcintSubunidade, "SELECT PKId, strDescricao FROM " & gstrSubUnidade & " where intOrgao = " & txt_intOrgao & " "
    '    End If
End Sub

Private Sub dbcintSubunidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintSubunidade_LostFocus()
    If dbcintSubunidade.ListIndex = -1 Then
        txt_intSubunidade = Space$(0)
        FormataCodigoTrabalho
    End If
End Sub

Private Sub dbcintTipoCredito_Click()
    If dbcintTipoCredito.ListIndex > -1 Then
        txt_intTipoCredito.Text = RetornaCodigo("strCodigo", "strDescricao", gstrTipoCredito, dbcintTipoCredito.Text)
    Else
        txt_intTipoCredito.Text = Space$(0)
    End If
End Sub

Private Sub dbcintTipoCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintTipoCredito_LostFocus()
    If dbcintTipoCredito.ListIndex = -1 Then
        txt_intTipoCredito = Space$(0)
    End If
End Sub

Private Sub dbcintUnidadeOrcamentaria_Click()
    
    dbcintSubunidade.Clear
    With Me.ActiveControl
        If .Name = "dbcintUnidadeOrcamentaria" Or Mid(.Name, 1, 4) = "cmd_" Then
            FormataCodigoTrabalho
        End If
    End With
    
End Sub

Private Sub dbcintUnidadeOrcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub cmd_ElementoDespesa_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadElementosDespesa, dbcintElementoDespesa
    If Not blnOrcamento Then frmCadElementosDespesa.Caption = frmCadElementosDespesa.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_Orgao_Click()
    If blnOrcamento = True Then
        gbytMenu = gbytMenuCadastro
    Else
        gbytMenu = gbytMenuProposta
    End If
    CarregaForm frmCadOrgao, dbcintOrgao
    If Not blnOrcamento Then frmCadOrgao.Caption = frmCadOrgao.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_ProjetoAtividade_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadProjetoAtividade, dbcintProjetoAtividade
    strChamadaForm = "ProjetoAtividade"
    If Not blnOrcamento Then frmCadProjetoAtividade.Caption = frmCadProjetoAtividade.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_SubFuncao_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadSubfuncaoDoGoverno, dbcintSubFuncao
    If Not blnOrcamento Then frmCadSubfuncaoDoGoverno.Caption = frmCadSubfuncaoDoGoverno.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_SubPrograma_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadSubPrograma, dbcintSubPrograma
    If dbcintPrograma.ListIndex >= 0 Then
        frmCadSubPrograma.Tag = dbcintPrograma.ItemData(dbcintPrograma.ListIndex)
    Else
        frmCadSubPrograma.Tag = ""
    End If
    If Not blnOrcamento Then frmCadSubPrograma.Caption = frmCadSubPrograma.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_Programa_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadPrograma, dbcintPrograma
    If Not blnOrcamento Then frmCadPrograma.Caption = frmCadPrograma.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_Funcao_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadFuncaoDoGoverno, dbcintFuncao
    If Not blnOrcamento Then frmCadFuncaoDoGoverno.Caption = frmCadFuncaoDoGoverno.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_Subunidade_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadSubunidade, dbcintSubunidade
    strChamadaForm = "SubUnidade"
    If dbcintUnidadeOrcamentaria.ListIndex >= 0 Then
        frmCadSubunidade.Tag = dbcintUnidadeOrcamentaria.ItemData(dbcintUnidadeOrcamentaria.ListIndex)
    Else
        frmCadSubunidade.Tag = ""
    End If
    If Not blnOrcamento Then frmCadSubunidade.Caption = frmCadSubunidade.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_TipoCredito_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadTipoCredito, dbcintTipoCredito, gstrQueryTipoCredito("0,1,2,3")
    If Not blnOrcamento Then frmCadTipoCredito.Caption = frmCadTipoCredito.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub cmd_UnidadeOrcamentaria_Click()
    gbytMenu = IIf(blnOrcamento, 0, 1)
    CarregaForm frmCadUnidadeOrcamentaria, dbcintUnidadeOrcamentaria
    strChamadaForm = "UnidadeOrcamentaria"
    If dbcintOrgao.ListIndex >= 0 Then
        frmCadUnidadeOrcamentaria.Tag = dbcintOrgao.ItemData(dbcintOrgao.ListIndex)
    Else
        frmCadUnidadeOrcamentaria.Tag = ""
    End If
    If Not blnOrcamento Then frmCadUnidadeOrcamentaria.Caption = frmCadUnidadeOrcamentaria.Caption & " (Proposta Orçamentária)"
End Sub

Private Sub dbcintUnidadeOrcamentaria_LostFocus()
    If dbcintUnidadeOrcamentaria.ListIndex = -1 Then
        txt_intUnidadeOrcamentaria = Space$(0)
        FormataCodigoTrabalho
    End If
End Sub

Private Sub Form_Activate()
    
    gintCodSeguranca = mIntCodSeguranca
    
    VirificaGradeListView Me
    
    If blnOrcamento = True Then
        txtintExercicio = gintExercicio
        txtintCodigoReduzido.OLEDropMode = 0
        txtintCodigoReduzido.Enabled = True
        txtintCodigoReduzido.SetFocus
        txtdblValor.OLEDropMode = 1
        TrocaCorObjeto txtdblValor, True
    Else
        txtintExercicio = gintExercicio + 1
        txtintCodigoReduzido.OLEDropMode = 1
    End If
    
    If mblnselecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    End If
    
    If blnOrcamento = True Then
        HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuCadastro), gstrBtnArquivo, gstrIncluiElementoDespesa
        HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuCadastro), gstrBtnArquivo, gstrIncluiProjetoAtividade
    Else
        HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuProposta), gstrBtnArquivo, gstrIncluiElementoDespesa
        HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuProposta), gstrBtnArquivo, gstrIncluiProjetoAtividade
    End If
    
    Select Case strChamadaForm
    Case Is = "UnidadeOrcamentaria"
        intIndiceCombo = dbcintUnidadeOrcamentaria.ListIndex
        LeDaTabelaParaObj gstrUnidadeOrcamentaria, dbcintUnidadeOrcamentaria, strQueryUO
        dbcintUnidadeOrcamentaria.ListIndex = intIndiceCombo
    Case Is = "SubUnidade"
        intIndiceCombo = dbcintSubunidade.ListIndex
        LeDaTabelaParaObj "", dbcintSubunidade, strQuerySubunidade
        dbcintSubunidade.ListIndex = intIndiceCombo
    Case Is = "ProjetoAtividade"
        intIndiceCombo = dbcintProjetoAtividade.ListIndex
        If blnOrcamento = True Then
            LeDaTabelaParaObj gstrProjeto, dbcintProjetoAtividade, "SELECT PKId, strDescricao FROM " & gstrProjeto & " WHERE intExercicio=" & gintExercicio
        Else
            LeDaTabelaParaObj gstrProjeto, dbcintProjetoAtividade, "SELECT PKId, strDescricao FROM " & gstrProjeto & " WHERE intExercicio=" & gintExercicio + 1
        End If
        
        dbcintProjetoAtividade.ListIndex = intIndiceCombo
    End Select
    
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, _
    gstrAplicar, _
    gstrDeletar, _
    gstrSalvar, _
    gstrNovo
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, _
    gstrIncluiElementoDespesa, _
    gstrIncluiProjetoAtividade
    
    
    
End Sub

Private Sub Form_GotFocus()
    txt_intOrgao.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    mblnProgramaAprovado = blnProgramaAprovado
    
    If blnOrcamento = True Then
        LeDaTabelaParaObj gstrOrgao, dbcintOrgao, "SELECT PKId, strDescricao FROM " & gstrOrgao & " WHERE intExercicio = " & gintExercicio
        LeDaTabelaParaObj gstrFuncaoDoGoverno, dbcintFuncao, "SELECT PKId, strDescricao FROM " & gstrFuncaoDoGoverno & " WHERE intExercicio=" & gintExercicio
        LeDaTabelaParaObj gstrSubFuncaoGoverno, dbcintSubFuncao, "SELECT PKId, strDescricao FROM " & gstrSubFuncaoGoverno & " WHERE intExercicio=" & gintExercicio
        LeDaTabelaParaObj gstrPrograma, dbcintPrograma, "SELECT PKId, strDescricao FROM " & gstrPrograma & " WHERE intExercicio=" & gintExercicio
        LeDaTabelaParaObj gstrSubPrograma, dbcintSubPrograma, "SELECT PKId, strDescricao FROM " & gstrSubPrograma & " WHERE intExercicio=" & gintExercicio
        LeDaTabelaParaObj gstrProjeto, dbcintProjetoAtividade, "SELECT PKId, strDescricao FROM " & gstrProjeto & " WHERE intExercicio=" & gintExercicio
        LeDaTabelaParaObj gstrElementoDespesa, dbcintElementoDespesa, "SELECT PKId, strDescricao FROM " & gstrElementoDespesa & " WHERE intExercicio=" & gintExercicio
        LeDaTabelaParaObj "", dbcintTipoCredito, gstrQueryTipoCredito("0,1,2,3")
        LeDaTabelaParaObj gstrVinculo, cbointVinculo
        LeDaTabelaParaObj gstrFonteRecurso, cbointFonteRecurso, "SELECT PKId, strDescricao FROM " & gstrFonteRecurso & " WHERE intExercicio=" & gintExercicio
        LeDaTabelaParaObj gstrSubUnidade, dbcintSubunidade, "SELECT PKId, strDescricao FROM " & gstrSubUnidade
        LeDaTabelaParaObj gstrModalidade, dbcintModalidade
        
        txtintCodigoReduzido.OLEDragMode = 1
        
        'DESABILITADO POR M4RCELØ 11/08/2004
        'tdb_Lista.Columns("strCodigoElementoDespesa").AllowFocus = False
        
        tdb_Lista.Columns("Valor").AllowFocus = False
    Else
        LeDaTabelaParaObj gstrOrgao, dbcintOrgao, "SELECT PKId, strDescricao FROM " & gstrOrgao & " WHERE intExercicio = " & gintExercicio + 1
        LeDaTabelaParaObj gstrFuncaoDoGoverno, dbcintFuncao, "SELECT PKId, strDescricao FROM " & gstrFuncaoDoGoverno & " WHERE intExercicio=" & gintExercicio + 1
        LeDaTabelaParaObj gstrSubFuncaoGoverno, dbcintSubFuncao, "SELECT PKId, strDescricao FROM " & gstrSubFuncaoGoverno & " WHERE intExercicio=" & gintExercicio + 1
        LeDaTabelaParaObj gstrPrograma, dbcintPrograma, "SELECT PKId, strDescricao FROM " & gstrPrograma & " WHERE intExercicio=" & gintExercicio + 1
        LeDaTabelaParaObj gstrSubPrograma, dbcintSubPrograma, "SELECT PKId, strDescricao FROM " & gstrSubPrograma & " WHERE intExercicio=" & gintExercicio + 1
        LeDaTabelaParaObj gstrProjeto, dbcintProjetoAtividade, "SELECT PKId, strDescricao FROM " & gstrProjeto & " WHERE intExercicio=" & gintExercicio + 1
        LeDaTabelaParaObj gstrElementoDespesa, dbcintElementoDespesa, "SELECT PKId, strDescricao FROM " & gstrElementoDespesa & " WHERE intExercicio=" & gintExercicio + 1
        LeDaTabelaParaObj "", dbcintTipoCredito, gstrQueryTipoCredito("0,1,2,3")
        LeDaTabelaParaObj gstrVinculo, cbointVinculo
        LeDaTabelaParaObj gstrFonteRecurso, cbointFonteRecurso, "SELECT PKId, strDescricao FROM " & gstrFonteRecurso & " WHERE intExercicio=" & gintExercicio + 1
        LeDaTabelaParaObj gstrSubUnidade, dbcintSubunidade, "SELECT PKId, strDescricao FROM " & gstrSubUnidade
        LeDaTabelaParaObj gstrModalidade, dbcintModalidade
        'DESABILITADO POR M4RCELØ 11/08/2004
        'tdb_Lista.Columns("strCodigoElementoDespesa").AllowFocus = True
        
        tdb_Lista.Columns("Valor").AllowFocus = True
    End If
    
    LeDaTabelaParaObj "", cbointEvento, strQueryEvento
    If cbointEvento.ListCount = 1 Then
        cbointEvento.ListIndex = 0
    End If
    
    VerificaListaAutomatica "", tdb_Lista, strQuery
    txt_total = tdb_Lista.Columns("dblTotal")
    VerificaObjParaAplicar mobjAux
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub EditaValor(Pkid As Variant, Col As Integer, Row As Variant)
    Dim adoResultado As ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo TrataErroLocal
    
PosGravaValor:
    
    DoEvents
    
    If blnOrcamento = True Then Exit Sub
    
    strSQL = "Select " & gstrISNULL("intExercicio", "0") & " intExercicio FROM " & gstrProgramaDeTrabalho & " WHERE PKID =" & IIf(Val(intPkIdRow) = 0, "0", intPkIdRow)
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            If adoResultado!intExercicio = gintExercicio Then Exit Sub
        End If
    End If
    
    
    'Vamos tornar o campo valor editavel
    If Col = 9 And Not txt_dblValor.Visible And Not tdb_Lista.FilterActive Then
        
        'Vamos passar os atributos a caixa de texto de edicao
        txt_dblValor.Width = tdb_Lista.Columns("Valor").Width
        txt_dblValor.Height = tdb_Lista.RowHeight
        txt_dblValor.Top = tdb_Lista.Top + tdb_Lista.RowTop(Row)
        txt_dblValor.Left = tdb_Lista.Left + tdb_Lista.Columns("Valor").Left
        
        txt_dblValor.Text = gstrConvVrDoSql(tdb_Lista.Columns("Valor").Value)
        
        txt_dblValor.Visible = True
        
        txt_dblValor.SetFocus
        
        dblValorAnt = tdb_Lista.Columns("Valor").Value 'Alfred 18/07/2003
        
    ElseIf txt_dblValor.Visible Then
        
        txt_dblValor.Visible = False
        
        If Val(gstrConvVrParaSql(txt_dblValor.Text)) <> dblValorAnt Then 'Alfred 18/07/2003
        
        'Vamos atualizar o valor na tabela de nprograma de trabalho
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        gobjBanco.Execute "UPDATE " & gstrProgramaDeTrabalho & " SET dblValor = " & gstrConvVrParaSql(txt_dblValor.Text) & _
        " WHERE PkId = " & intPkIdRow
        gobjBanco.ExecutaCommitTrans
        Set gobjBanco = Nothing
        
        LimpaObjetos
        
        LeDaTabelaParaObj "", tdb_Lista, strQuery
        
    End If
    
    txt_total = tdb_Lista.Columns("dblTotal")
    
    If intPkIdRow > 0 Then
        Set adoResultado = tdb_Lista.DataSource
        adoResultado.Find "PKId = '" & Pkid & "'"
        tdb_Lista.MarqueeStyle = dbgHighlightRow
    End If
    
    GoTo PosGravaValor
    
End If

Exit Sub

TrataErroLocal:
Resume Next

End Sub

Private Sub tdb_Lista_BeforeRowColChange(Cancel As Integer)
    blnFilterBar = False
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
    'ORC1599 - F3RN4NDØ
    mblnAlterando = True
End Sub

Private Sub tdb_Lista_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    txt_dblValor.Width = tdb_Lista.Columns("Valor").Width
    txt_dblValor.Height = tdb_Lista.RowHeight
    txt_dblValor.Top = tdb_Lista.Top + tdb_Lista.RowTop(tdb_Lista.Row)
    txt_dblValor.Left = tdb_Lista.Left + tdb_Lista.Columns("Valor").Left
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub


Private Sub tdb_Lista_FilterChange()
    blnFilterBar = True
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 6 Then
        Value = gvntFormatacaoEspecifica(Value, 3)
    ElseIf ColIndex = 9 Then
        Value = gstrConvVrDoSql(Value)
    End If
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        EditaValor tdb_Lista.Columns("PkId").Value, 9, tdb_Lista.Row
        KeyAscii = 0
    Else
        CaracterValido KeyAscii
    End If
End Sub

Private Sub tdb_Lista_LostFocus()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluiProjetoAtividade
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
End Sub
Private Sub Preenche_camposaldo()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    If Trim(Len(txtPKID)) <> 0 Then
        strSQL = ""
        strSQL = strSQL & "SELECT saldoIni , SUM(Empenhado) dblEmpenhado,"
        strSQL = strSQL & "SUM(Suplementado) dblSuplementado , SUM(Anulado) dblAnulado,"
        strSQL = strSQL & "SUM(Reservado) dblReservado, SUM(Bloqueado) dblBloqueado,"
        strSQL = strSQL & "SUM(Pago)dblPago, 0 dblLiquidado"
        strSQL = strSQL & " FROM " & gstrContaValoresAcumulados
        strSQL = strSQL & " WHERE INTPROGRAMADETRABALHO = " & txtPKID
        strSQL = strSQL & " GROUP BY saldoIni "
        
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If adoResultado.EOF = False Then
                With adoResultado
                    'liquidado será colocado posteriormente
                    txt_strCodigoSaldo = txtstrCodigo
                    txt_saldoini = gstrConvVrDoSql(!SaldoIni)
                    txt_Empenhado = gstrConvVrDoSql(!dblEmpenhado)
                    txt_liquidado = gstrConvVrDoSql(!dblliquidado)
                    txt_Pago = gstrConvVrDoSql(!dblPago)
                    txt_anulado = gstrConvVrDoSql(!dblAnulado)
                    txt_Suplementado = gstrConvVrDoSql(!dblSuplementado)
                    txt_Reservado = gstrConvVrDoSql(!dblReservado)
                    txt_TotalBloqueado = gstrConvVrDoSql(!dblBloqueado)
                    txt_SaldoParaEmpenho = gstrConvVrDoSql((!SaldoIni + !dblSuplementado) - !dblAnulado - !dblEmpenhado - !dblBloqueado - !dblReservado)
                End With
                Exit Sub
            End If
        End If
    End If
    
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim strSQL As String
    
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            If blnFilterBar Then Exit Sub
            mblnClickOk = False
'            TrocaCorObjeto txt_codEvento, True
'            TrocaCorObjeto cbointEvento, True
'            TrocaCorObjeto cmd_Evento, True
            txtPKID = tdb_Lista.Columns(0).Value
            txt_intConvenio.Text = ""
            LeDaTabelaParaObj gstrProgramaDeTrabalho, Me, mblnProgramaAprovado
            
            txt_total = tdb_Lista.Columns("dblTotal")
            
            If Not IsNull(.Columns("UnidadeOrcamentaria").Value) Then
                txt_intFonteRecurso = .Columns("FonteRecurso").Value
                txt_intFonteRecurso_LostFocus
                txt_intOrgao = .Columns("Orgao").Value
                txt_intOrgao_LostFocus
                txt_intUnidadeOrcamentaria = .Columns("UnidadeOrcamentaria").Value
                txt_intUnidadeOrcamentaria_LostFocus
                
                LeDaTabelaParaObj gstrSubUnidade, dbcintSubunidade, "SELECT Pkid, strDescricao FROM " & gstrSubUnidade & " WHERE intOrgao = " & gstrItemData(dbcintOrgao) & _
                " AND intUnidadeOrcamentaria = " & gstrItemData(dbcintUnidadeOrcamentaria)
                
                txt_intSubunidade = .Columns("Subunidade").Value
                txt_intSubunidade_LostFocus
            End If
            
            txt_intElementoDespesa = .Columns("strCodigoElementoDespesa").Value
            txt_intElementoDespesa_LostFocus
            
            gCorLinhaSelecionada tdb_Lista
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
            
            EditaValor tdb_Lista.Columns("PkId").Value, tdb_Lista.Col, tdb_Lista.Row
            
            intPkIdRow = .Columns("PkId").Value
            mblnAlterando = True
            If blnOrcamento = True Then
                HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuCadastro), gstrBtnArquivo, gstrIncluiProjetoAtividade
                HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuCadastro), gstrBtnArquivo, gstrIncluiElementoDespesa
            Else
                HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuProposta), gstrBtnArquivo, gstrIncluiElementoDespesa
                HabilitaDesabilitaBotao1 (gbytMenu = gbytMenuProposta), gstrBtnArquivo, gstrIncluiProjetoAtividade
            End If
            FormataCodigoTrabalho
            Screen.MousePointer = vbHourglass
            Preenche_camposaldo
            Screen.MousePointer = vbDefault
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim gintExercicioAux As Integer
    'Cláudio - Bit para gravar na tabela
    '1 - Proposta Orcamentaria : deverá gravar 0
    '2 - Orcamento : deverá gravar 1
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then chkbytSituacao.Value = Abs(blnOrcamento)
    
    If UCase(strModoOperacao) = gstrSalvar And mblnAlterando = False Then
        If Not blnDadosOk Then Exit Sub
    End If
    If UCase(strModoOperacao) = gstrIncluiElementoDespesa Then
        If blnOrcamento = True Then
            gbytMenu = gbytMenuCadastro
        Else
            gbytMenu = gbytMenuProposta
        End If
        CarregaForm frmCadElementosDespesa
        frmCadElementosDespesa.blnProgramaDeTrabalho = True
        Exit Sub
    ElseIf UCase(strModoOperacao) = gstrIncluiProjetoAtividade Then
        If blnOrcamento = True Then
            gbytMenu = gbytMenuCadastro
        Else
            gbytMenu = gbytMenuProposta
        End If
        CarregaForm frmCadProjetoAtividade
        frmCadProjetoAtividade.strValorFonteRecurso = STRFONTERECURSO
        frmCadProjetoAtividade.blnProgramaDeTrabalho = True
        txt_total = tdb_Lista.Columns("dblTotal")
        Exit Sub
    ElseIf UCase(strModoOperacao) = gstrPreencherLista Then
        '        Dim strvdbcintOrgao, strvdbcintFuncao, strvdbcintSubFuncao, strvdbcintPrograma, strvdbcintSubPrograma, strvdbcintProjetoAtividade As String 'usadas para guadar os valores atuais das combos
        '        Dim strvdbcintElementoDespesa, strvdbcintTipoCredito, strvcbointVinculo, strvcbointFonteRecurso, strvdbcintSubunidade, strvcbointEvento As String 'usadas para guadar os valores atuais das combos
        '        'If UCase(Me.ActiveControl.Name) = UCase(dbcintOrgao.Name) Then
        '        'PreencherListaDeOpcoes dbcintOrgao
        '            'guarda os valores
        '            strvdbcintOrgao = dbcintOrgao.Text
        '            strvdbcintFuncao = dbcintFuncao.Text
        '            strvdbcintSubFuncao = dbcintSubFuncao.Text
        '            strvdbcintPrograma = dbcintPrograma.Text
        '            strvdbcintSubPrograma = dbcintSubPrograma.Text
        '            strvdbcintProjetoAtividade = dbcintProjetoAtividade.Text
        '            strvdbcintElementoDespesa = dbcintElementoDespesa.Text
        '            strvdbcintTipoCredito = dbcintTipoCredito.Text
        '            strvcbointVinculo = cbointVinculo.Text
        '            strvcbointFonteRecurso = cbointFonteRecurso.Text
        '            strvdbcintSubunidade = dbcintSubunidade.Text
        '            strvcbointEvento = cbointEvento.Text
        '            'atualiza os combos
        '            If blnOrcamento = False Then
        '                gintExercicioAux = gintExercicio + 1
        '            Else
        '                gintExercicioAux = gintExercicio
        '            End If
        '            LeDaTabelaParaObj gstrOrgao, dbcintOrgao, "SELECT PKId, strDescricao FROM " & gstrOrgao & " WHERE intExercicio = " & gintExercicioAux & " AND strDescricao LIKE '" & dbcintOrgao.Text & "%'"
        '            LeDaTabelaParaObj gstrOrgao, dbcintOrgao, "SELECT PKId, strDescricao FROM " & gstrOrgao & " WHERE intExercicio = " & gintExercicioAux
        '            LeDaTabelaParaObj gstrFuncaoDoGoverno, dbcintFuncao, "SELECT PKId, strDescricao FROM " & gstrFuncaoDoGoverno & " WHERE intExercicio=" & gintExercicioAux
        '            LeDaTabelaParaObj gstrSubFuncaoGoverno, dbcintSubFuncao, "SELECT PKId, strDescricao FROM " & gstrSubFuncaoGoverno & " WHERE intExercicio=" & gintExercicioAux
        '            LeDaTabelaParaObj gstrPrograma, dbcintPrograma, "SELECT PKId, strDescricao FROM " & gstrPrograma & " WHERE intExercicio=" & gintExercicioAux
        '            LeDaTabelaParaObj gstrSubPrograma, dbcintSubPrograma, "SELECT PKId, strDescricao FROM " & gstrSubPrograma & " WHERE intExercicio=" & gintExercicioAux
        '            LeDaTabelaParaObj gstrProjeto, dbcintProjetoAtividade, "SELECT PKId, strDescricao FROM " & gstrProjeto & " WHERE intExercicio=" & gintExercicioAux
        '            LeDaTabelaParaObj gstrElementoDespesa, dbcintElementoDespesa, "SELECT PKId, strDescricao FROM " & gstrElementoDespesa & " WHERE intExercicio=" & gintExercicioAux
        '            LeDaTabelaParaObj "", dbcintTipoCredito, gstrQueryTipoCredito("0,1,2,3")
        '            LeDaTabelaParaObj gstrVinculo, cbointVinculo
        '            LeDaTabelaParaObj gstrFonteRecurso, cbointFonteRecurso, "SELECT PKId, strDescricao FROM " & gstrFonteRecurso & " WHERE intExercicio=" & gintExercicioAux
        '            LeDaTabelaParaObj gstrSubUnidade, dbcintSubunidade, "SELECT PKId, strDescricao FROM " & gstrSubUnidade
        If UCase(Me.ActiveControl.Name) = UCase(cbointEvento.Name) Then
            LeDaTabelaParaObj "", cbointEvento, strQueryEvento
        End If
        '            'devolve os valores
        '            dbcintOrgao.Text = strvdbcintOrgao
        '            dbcintFuncao.Text = strvdbcintFuncao
        '            dbcintSubFuncao.Text = strvdbcintSubFuncao
        '            dbcintPrograma.Text = strvdbcintPrograma
        '            dbcintSubPrograma.Text = strvdbcintSubPrograma
        '            dbcintProjetoAtividade.Text = strvdbcintProjetoAtividade
        '            dbcintElementoDespesa.Text = strvdbcintElementoDespesa
        '            dbcintTipoCredito.Text = strvdbcintTipoCredito
        '            cbointVinculo.Text = strvcbointVinculo
        '            cbointFonteRecurso.Text = strvcbointFonteRecurso
        '            dbcintSubunidade.Text = strvdbcintSubunidade
        '            cbointEvento.Text = strvcbointEvento
        '        'End If
        '        Exit Sub
    Else
        
        'Cláudio - Rotina de Verificação para que um projeto de atividade não seja cadastrado para mais de um órgão.
        'Como segue na pendência 183 (blnCadProjetoAtividade)
        If UCase(strModoOperacao) = gstrSalvar And mblnAlterando = False Then
            
            If blnVerificaParamProj = False Then
                
                If Not blnCadProjetoAtividade Then
                    ExibeMensagem "Este Projeto/Atividade já existe em outro órgão."
                    dbcintProjetoAtividade.SetFocus 'ORC1599 - F3RN4NDØ
                    Exit Sub
                End If
                
            End If
            
        End If
        
        If UCase(strModoOperacao) = gstrDeletar And blnOrcamento = True Then
            ExibeMensagem "Não é possível excluir uma dotação já cadastrada."
            Exit Sub
        End If
        
        If UCase(strModoOperacao) = gstrSalvar And mblnAlterando = True And blnOrcamento = True Then
            If blnDadosOk Then
                If txtdblValor = "0,00" Then
                    ToolBarGeral strModoOperacao, gstrProgramaDeTrabalho, mblnAlterando, tdb_Lista, Me, mobjAux, "", strQueryAplicar, rptProgramaDeTrabalho, strQueryRelatorio
                    txt_total = tdb_Lista.Columns("dblTotal")
                    LimpaObjetos
                Else
                    ExibeMensagem "A alteração só é permitida para dotações com Saldo Inicial igual a zero."
                End If
            End If
            
            Exit Sub
        End If
        
        If UCase(strModoOperacao) = gstrLocalizar And Len(Trim(txtstrCodigo)) > 0 Then
            txtstrCodigo = Replace(txtstrCodigo, ".", ",")
        End If
        
        If UCase(strModoOperacao) = gstrLocalizar And Len(Trim(txtstrCodigo)) > 0 Then
            txtstrCodigo = Replace(txtstrCodigo, ",", ".")
        End If
        
        If UCase(strModoOperacao) = gstrNovo And blnOrcamento = True Then
            TrocaCorObjeto txtintCodigoReduzido, False
            txtintCodigoReduzido_GotFocus
            txtintCodigoReduzido.SetFocus
            chkbytEducacao.Value = 0
            chkbytSaude.Value = 0
            tab_3dPasta.Tab = 0
        End If
        
        If UCase(strModoOperacao) = gstrNovo Then
            TrocaCorObjeto txt_codEvento, False
            TrocaCorObjeto cbointEvento, False
            TrocaCorObjeto cmd_Evento, False
            LeDaTabelaParaObj "", cbointEvento, strQueryEvento
            If cbointEvento.ListCount = 1 Then
                cbointEvento.ListIndex = 0
            End If
        End If
        
        If UCase(strModoOperacao) = gstrSalvar And mblnAlterando = False And blnOrcamento = True Then
            If Not GeraMovimentosByEvento(gstrItemData(cbointEvento), gstrDataDoSistema, Str(CDbl(IIf(Len(Trim(txtdblValor)) = 0, "0,00", txtdblValor))), "", IIf(Len(Trim(txtintCodigoReduzido)) = 0, "0", txtintCodigoReduzido), "1") Then
                ExibeMensagem "Ocorreram erros durante a gravação deste evento.Entre em contato com o fornecedor."
            End If
        End If
        
        If UCase(strModoOperacao) = gstrLocalizar Then
            LeDaTabelaParaObj "", tdb_Lista, strQuery
        End If
        
        If strModoOperacao <> gstrLocalizar Then
            ToolBarGeral strModoOperacao, gstrProgramaDeTrabalho, mblnAlterando, tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar, rptProgramaDeTrabalho, strQueryRelatorio, False
            txt_total = tdb_Lista.Columns("dblTotal")
        End If
        
        If Not strModoOperacao = gstrAplicar And Not strModoOperacao = gstrLocalizar And Not strModoOperacao = gstrSalvar Then LimpaObjetos
        
    End If
End Sub

Private Function blnProgramaAprovado() As Boolean
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT bytSituacao "
    strSQL = strSQL & "FROM " & gstrProgramaDeTrabalho & " PT "
    If blnOrcamento = True Then
        strSQL = strSQL & "WHERE PT.intExercicio = " & gintExercicio & " "
    Else
        strSQL = strSQL & "WHERE PT.intExercicio = " & gintExercicio + 1 & " "
    End If
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            blnProgramaAprovado = adoResultado!bytSituacao
            HabilitaDesabilitaBotao1 Abs(blnProgramaAprovado) - 1, _
            gstrBtnArquivo, gstrNovo, gstrSalvar, _
            gstrDeletar
        End If
    End If
End Function


Private Function blnVerificaParamProj() As Boolean
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT bytProjAtvVariosOrgaos "
    strSQL = strSQL & "FROM " & gstrConfiguracaoGeral
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            blnVerificaParamProj = adoResultado!bytProjAtvVariosOrgaos
        End If
    End If
End Function



Private Function strQueryAplicar() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strCodigo "
    strSQL = strSQL & "FROM " & gstrProgramaDeTrabalho & " "
    strSQL = strSQL & "WHERE PT.intExercicio = " & Val(txtintExercicio) & " "
    strQueryAplicar = strSQL
End Function

Private Function strQuery() As String
    
    Dim strSQL  As String
    strSQL = ""
    
    strSQL = strSQL & "SELECT PT.PKId, PT.intCodigoReduzido, OG.strCodigo AS Orgao, "
    strSQL = strSQL & "UO.strCodigo AS UnidadeOrcamentaria, SU.strCodigo AS SubUnidade, "
    strSQL = strSQL & "PA.strCodigo AS ProjetoAtividade, MO.strDescricao AS Modalidade, "
    strSQL = strSQL & "ED.strCodigoElementoDespesa, ED.strDescricao AS Elemento, FR.strCodigo AS FonteRecurso, "
    strSQL = strSQL & "PT.dblValor AS Valor, (SELECT " & gstrISNULL("SUM(dblValor)", "0")
    strSQL = strSQL & "FROM " & gstrProgramaDeTrabalho & " PT "
    If blnOrcamento = True Then
        strSQL = strSQL & "WHERE PT.intExercicio = " & (gintExercicio) & ") AS dblTotal "
    Else
        strSQL = strSQL & "WHERE PT.intExercicio = " & (gintExercicio + 1) & ") AS dblTotal "
    End If
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT, " & gstrOrgao & " OG, "
    strSQL = strSQL & gstrUnidadeOrcamentaria & " UO, " & gstrSubUnidade & " SU, "
    strSQL = strSQL & gstrProjeto & " PA, " & gstrElementoDespesa & " ED, " & gstrFonteRecurso & " FR, " & gstrModalidade & " MO "
    strSQL = strSQL & "WHERE PT.intOrgao " & strOUTJSQLServer & "= OG.PKId " & strOUTJOracle
    strSQL = strSQL & "AND PT.intUnidadeOrcamentaria " & strOUTJSQLServer & "= UO.PKId " & strOUTJOracle
    strSQL = strSQL & "AND PT.intProjetoAtividade " & strOUTJSQLServer & "= PA.PKId " & strOUTJOracle
    strSQL = strSQL & "AND PT.intSubunidade " & strOUTJSQLServer & "= SU.PKId " & strOUTJOracle
    strSQL = strSQL & "AND PT.intElementoDespesa " & strOUTJSQLServer & "= ED.PKId " & strOUTJOracle
    strSQL = strSQL & "AND PT.intFonteRecurso " & strOUTJSQLServer & "= FR.PKId " & strOUTJOracle
    strSQL = strSQL & "AND PT.intModalidade " & strOUTJSQLServer & "= MO.PKId " & strOUTJOracle
    If blnOrcamento = True Then
        strSQL = strSQL & "AND PT.intExercicio = " & (gintExercicio) & " "
    Else
        strSQL = strSQL & "AND PT.intExercicio = " & (gintExercicio + 1) & " "
    End If
    
    'Codigo Reduzido
    If Trim(txtintCodigoReduzido.Text) <> "" Then
        strSQL = strSQL & " AND PT.intCodigoReduzido = " & Trim(txtintCodigoReduzido)
    End If
    'Orgão
    If Trim(dbcintOrgao.Text) <> "" Then
        strSQL = strSQL & " AND PT.intOrgao = " & gstrItemData(dbcintOrgao)
    End If
    'Unidade Orçamentaria
    If Trim(dbcintUnidadeOrcamentaria.Text) <> "" Then
        strSQL = strSQL & " AND PT.Intunidadeorcamentaria = " & gstrItemData(dbcintUnidadeOrcamentaria)
    End If
    'SubUnidade
    If Trim(dbcintSubunidade.Text) <> "" Then
        'strSql = strSql & " AND SU.Pkid = " & gstrItemData(dbcintSubunidade)
        strSQL = strSQL & " AND pt.intsubunidade = " & gstrItemData(dbcintSubunidade)
    End If
    'Projeto Atividade
    If Trim(dbcintProjetoAtividade.Text) <> "" Then
        'strSql = strSql & " AND PA.Pkid = " & gstrItemData(dbcintProjetoAtividade)
        strSQL = strSQL & " AND pt.intProjetoAtividade = " & gstrItemData(dbcintProjetoAtividade)
    End If
    'Elemento da Despesa
    If Trim(dbcintElementoDespesa.Text) <> "" Then
        strSQL = strSQL & " AND pt.intelementodespesa = " & gstrItemData(dbcintElementoDespesa)
    End If
    'Fonte de Recurso
    If Trim(cbointFonteRecurso.Text) <> "" Then
        strSQL = strSQL & " AND pt.intfonterecurso = " & gstrItemData(cbointFonteRecurso)
    End If
    'Tipo de Crédito
    If Trim(dbcintTipoCredito.Text) <> "" Then
        strSQL = strSQL & " AND PT.intTipoCredito = " & gstrItemData(dbcintTipoCredito)
    End If
    'Função
    If Trim(dbcintFuncao.Text) <> "" Then
        strSQL = strSQL & " AND PT.intFuncao = " & gstrItemData(dbcintFuncao)
    End If
    'SubFunção
    If Trim(dbcintSubFuncao.Text) <> "" Then
        strSQL = strSQL & " AND PT.intSubFuncao = " & gstrItemData(dbcintSubFuncao)
    End If
    'Programa
    If Trim(dbcintPrograma.Text) <> "" Then
        strSQL = strSQL & " AND PT.intPrograma = " & gstrItemData(dbcintPrograma)
    End If
    'SubPrograma
    If Trim(dbcintSubPrograma.Text) <> "" Then
        strSQL = strSQL & " AND PT.intSubPrograma = " & gstrItemData(dbcintSubPrograma)
    End If
    'EventoContabil
    If Trim(cbointEvento.Text) <> "" Then
        strSQL = strSQL & " AND PT.intEvento = " & gstrItemData(cbointEvento)
    End If
    'Vinculo
    If Trim(cbointVinculo.Text) <> "" Then
        strSQL = strSQL & " AND PT.intVinculo = " & gstrItemData(cbointVinculo)
    End If
    'Modalidade
    If Trim(dbcintModalidade.Text) <> "" Then
        strSQL = strSQL & " AND PT.intModalidade = " & gstrItemData(dbcintModalidade)
    End If
    
    '- Cláudio
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(cdt_numeric, "OG.strCodigo") & " , "
    strSQL = strSQL & gstrCONVERT(cdt_numeric, "UO.strCodigo") & " , "
    strSQL = strSQL & gstrCONVERT(cdt_numeric, "SU.strCodigo") & " , "
    strSQL = strSQL & gstrCONVERT(cdt_numeric, "PA.strCodigo") & " , "
    strSQL = strSQL & gstrCONVERT(cdt_numeric, "ED.strCodigoElementoDespesa")
    
    strQuery = strSQL
    
End Function

Private Sub tdb_Lista_Scroll(Cancel As Integer)
    EditaValor tdb_Lista.Columns("PkId").Value, 0, tdb_Lista.Row
End Sub

Private Sub txt_codEvento_GotFocus()
MarcaCampo txt_codEvento
End Sub

Private Sub txt_dblValor_GotFocus()
    MarcaCampo txt_dblValor
End Sub

Private Sub txt_DBLVALOR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        EditaValor tdb_Lista.Columns("PkId").Value, 0, tdb_Lista.Row
        tdb_Lista.SetFocus
    Else
        CaracterValido KeyAscii, "V", txt_dblValor
    End If
End Sub

Private Sub txt_intConvenio_GotFocus()
    MarcaCampo txt_intConvenio
End Sub

Private Sub txt_intConvenio_KeyPress(KeyAscii As Integer)
    'ORC1599 - F3RN4NDØ
    'CaracterValido KeyAscii, "A", txt_intConvenio
    gstrLimitaCampoValor txt_intConvenio, KeyAscii, 2, 0
End Sub

Private Sub txt_intConvenio_LostFocus()
    Dim strDescricao As String
    Dim intI As Integer
    'If Len(Trim(txt_intConvenio.Text)) > 0 And Len(Trim(txt_intModalidadeAplicacao.Text)) > 0 Then
    If Len(Trim(txt_intModalidadeAplicacao.Text)) > 0 Then
        strDescricao = CarregaModalidade(txt_intModalidadeAplicacao.Text, txt_intConvenio.Text)
        If Len(strDescricao) = 0 Then
            dbcintModalidade.ListIndex = -1
            txt_intModalidadeAplicacao.SetFocus
        Else
            For intI = 0 To dbcintModalidade.ListCount - 1
                If strDescricao = dbcintModalidade.ItemData(intI) Then
                    dbcintModalidade.ListIndex = intI
                    Exit For
                End If
            Next

        End If

    Else
        dbcintModalidade.ListIndex = -1
    End If
End Sub

Private Sub txt_intElementoDespesa_GotFocus()
    MarcaCampo txt_intElementoDespesa
    cbointEvento_LostFocus
End Sub

Private Sub txt_intElementoDespesa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intElementoDespesa
End Sub

Private Sub txt_intElementoDespesa_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    'If cbointEvento.ListIndex > -1 Then
    If Len(txt_intElementoDespesa) > 0 Then
        
        txt_intElementoDespesa = Trim(txt_intElementoDespesa) & String(15 - Len(Trim(txt_intElementoDespesa)), "0")
        
        strDescricao = RetornaCodigo("Pkid", "strCodigoElementoDespesa", gstrElementoDespesa, txt_intElementoDespesa.Text, " AND intExercicio = " & txtintExercicio)
        
        If Len(strDescricao) = 0 Then
            dbcintElementoDespesa.ListIndex = -1
            txt_intElementoDespesa.SetFocus
            cbointEvento.Clear
            txt_codEvento = ""
        Else
            
            For i = 0 To dbcintElementoDespesa.ListCount - 1
                If strDescricao = dbcintElementoDespesa.ItemData(i) Then
                    dbcintElementoDespesa.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintElementoDespesa.ListIndex = -1
    End If
    'End If
End Sub

Private Sub txt_codEvento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_codEvento
End Sub

Private Sub txt_codEvento_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Len(txt_codEvento) > 0 Then
        
        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrEvento, txt_codEvento.Text)
        
        If Len(strDescricao) = 0 Then
            cbointEvento.ListIndex = -1
            txt_codEvento.SetFocus
        Else
            
            For i = 0 To cbointEvento.ListCount - 1
                If strDescricao = cbointEvento.ItemData(i) Then
                    cbointEvento.ListIndex = i
                    Exit For
                End If
            Next
            
            strElementoDespesa = txt_intElementoDespesa
            If gintExercicio > 2006 Or Not blnOrcamento Then
                LeDaTabelaParaObj gstrElementoDespesa, dbcintElementoDespesa, "SELECT PKId, strDescricao FROM " & gstrElementoDespesa & " WHERE intExercicio=" & IIf(blnOrcamento = True, gintExercicio, gintExercicio + 1)
            Else
                LeDaTabelaParaObj gstrElementoDespesa, dbcintElementoDespesa, "SELECT PKId, strDescricao FROM " & gstrElementoDespesa & " WHERE intExercicio=" & IIf(blnOrcamento = True, gintExercicio, gintExercicio + 1) & _
                " AND " & strSUBSTRING & "(strCodigoElementoDespesa,1," & Len(BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoDespesa, "C", 0)) & ") = '" & _
                BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoDespesa, "C", 0) & "'"
            End If
            txt_intElementoDespesa = strElementoDespesa
            txt_intElementoDespesa_LostFocus
            
        End If
        
    Else
        cbointEvento.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intFonteRecurso_GotFocus()
MarcaCampo txt_intFonteRecurso
End Sub

Private Sub txt_intFonteRecurso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intFonteRecurso
End Sub

Private Sub txt_intFonteRecurso_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Val(txt_intFonteRecurso) > 0 Then
        
        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrFonteRecurso, txt_intFonteRecurso.Text, " AND intExercicio = " & txtintExercicio)
        
        DoEvents
        
        If Len(strDescricao) = 0 Then
            cbointFonteRecurso.ListIndex = -1
            txt_intFonteRecurso.SetFocus
        Else
            
            For i = 0 To cbointFonteRecurso.ListCount - 1
                If strDescricao = cbointFonteRecurso.ItemData(i) Then
                    cbointFonteRecurso.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        cbointFonteRecurso.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intFuncao_GotFocus()
MarcaCampo txt_intFuncao
End Sub

Private Sub txt_intFuncao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intFuncao
End Sub

Private Sub txt_intFuncao_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Len(txt_intFuncao) > 0 Then
        
        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrFuncaoDoGoverno, txt_intFuncao.Text, " AND intExercicio = " & txtintExercicio)
        
        If Len(strDescricao) = 0 Then
            dbcintFuncao.ListIndex = -1
            txt_intFuncao.SetFocus
        Else
            
            For i = 0 To dbcintFuncao.ListCount - 1
                If strDescricao = dbcintFuncao.ItemData(i) Then
                    dbcintFuncao.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintFuncao.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intModalidadeAplicacao_GotFocus()
    MarcaCampo txt_intModalidadeAplicacao
    dbcintModalidade_LostFocus
End Sub

Private Sub txt_intModalidadeAplicacao_KeyPress(KeyAscii As Integer)
    'ORC1599 - F3RN4NDØ
    'CaracterValido KeyAscii, "A", txt_intModalidadeAplicacao
    gstrLimitaCampoValor txt_intModalidadeAplicacao, KeyAscii, 3, 0
End Sub

Private Sub txt_intOrgao_GotFocus()
MarcaCampo txt_intOrgao
End Sub

'Private Sub txt_intModalidadeAplicacao_LostFocus()
'    Dim strDescricao As String
'    Dim i As Integer
'
'    If Len(txt_intModalidadeAplicacao) > 0 Then
'
'        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrModalidade, txt_intModalidadeAplicacao.Text)
'
'        If Len(strDescricao) = 0 Then
'            dbcintModalidade.ListIndex = -1
'            txt_intModalidadeAplicacao.SetFocus
'        Else
'            txt_intConvenio.Text = CarregaConvenio(strDescricao)
'            For i = 0 To dbcintModalidade.ListCount - 1
'                If strDescricao = dbcintModalidade.ItemData(i) Then
'                    dbcintModalidade.ListIndex = i
'                    Exit For
'                End If
'            Next
'
'        End If
'
'    Else
'        dbcintModalidade.ListIndex = -1
'    End If
'End Sub

Private Sub txt_intOrgao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intOrgao
End Sub

Private Sub txt_intOrgao_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Val(txt_intOrgao) > 0 Then
        
        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrOrgao, txt_intOrgao.Text, " AND intExercicio = " & txtintExercicio)
        
        If Len(strDescricao) = 0 Then
            dbcintOrgao.ListIndex = -1
            txt_intOrgao.SetFocus
        Else
            
            For i = 0 To dbcintOrgao.ListCount - 1
                If strDescricao = dbcintOrgao.ItemData(i) Then
                    dbcintOrgao.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintOrgao.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intPrograma_GotFocus()
MarcaCampo txt_intPrograma
End Sub

Private Sub txt_intPrograma_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intPrograma
End Sub

Private Sub txt_intPrograma_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Len(txt_intPrograma) > 0 Then
        
        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrPrograma, txt_intPrograma.Text, " AND intExercicio = " & txtintExercicio)
        
        If Len(strDescricao) = 0 Then
            dbcintPrograma.ListIndex = -1
            txt_intPrograma.SetFocus
        Else
            
            For i = 0 To dbcintPrograma.ListCount - 1
                If strDescricao = dbcintPrograma.ItemData(i) Then
                    dbcintPrograma.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintPrograma.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intProjetoAtividade_GotFocus()
MarcaCampo txt_intProjetoAtividade
End Sub

Private Sub txt_intProjetoAtividade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intProjetoAtividade
End Sub

Private Sub txt_intProjetoAtividade_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Len(txt_intProjetoAtividade) > 0 Then
        
        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrProjeto, txt_intProjetoAtividade.Text, " AND intExercicio = " & txtintExercicio)
        
        If Len(strDescricao) = 0 Then
            dbcintProjetoAtividade.ListIndex = -1
            txt_intProjetoAtividade.SetFocus
        Else
            
            For i = 0 To dbcintProjetoAtividade.ListCount - 1
                If strDescricao = dbcintProjetoAtividade.ItemData(i) Then
                    dbcintProjetoAtividade.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintProjetoAtividade.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intSubFuncao_GotFocus()
MarcaCampo txt_intSubFuncao
End Sub

Private Sub txt_intSubFuncao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intSubFuncao
End Sub

Private Sub txt_intSubFuncao_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Len(txt_intSubFuncao) > 0 Then
        
        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrSubFuncaoGoverno, txt_intSubFuncao.Text, " AND intExercicio = " & txtintExercicio)
        
        If Len(strDescricao) = 0 Then
            dbcintSubFuncao.ListIndex = -1
            txt_intSubFuncao.SetFocus
        Else
            
            For i = 0 To dbcintSubFuncao.ListCount - 1
                If strDescricao = dbcintSubFuncao.ItemData(i) Then
                    dbcintSubFuncao.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintSubFuncao.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intSubPrograma_GotFocus()
MarcaCampo txt_intSubPrograma
End Sub

Private Sub txt_intSubPrograma_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intSubPrograma
End Sub

Private Sub txt_intSubPrograma_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Len(txt_intSubPrograma) > 0 Then
        
        If Len(txt_intPrograma) > 0 Then
            strDescricao = RetornaCodigo("PkID", "strCodigo", gstrSubPrograma, txt_intSubPrograma.Text, " AND intPrograma = " & dbcintPrograma.ItemData(dbcintPrograma.ListIndex))
        End If
        
        If Len(strDescricao) = 0 Then
            dbcintSubPrograma.ListIndex = -1
            txt_intSubPrograma.SetFocus
        Else
            
            For i = 0 To dbcintSubPrograma.ListCount - 1
                If strDescricao = dbcintSubPrograma.ItemData(i) Then
                    dbcintSubPrograma.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintSubPrograma.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intSubunidade_GotFocus()
MarcaCampo txt_intSubunidade
End Sub

Private Sub txt_intSubunidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intSubunidade
End Sub

Private Sub txt_intSubunidade_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    
    
    
    If Val(txt_intSubunidade) > 0 Then
        
        
        If dbcintOrgao.ListIndex >= 0 And dbcintUnidadeOrcamentaria.ListIndex >= 0 Then
            strDescricao = RetornaCodigo("PkID", "strCodigo", gstrSubUnidade, txt_intSubunidade.Text, " AND intOrgao = " & dbcintOrgao.ItemData(dbcintOrgao.ListIndex) & " AND intUnidadeOrcamentaria = " & dbcintUnidadeOrcamentaria.ItemData(dbcintUnidadeOrcamentaria.ListIndex))
        End If
        
        If Len(strDescricao) = 0 Then
            dbcintSubunidade.ListIndex = -1
            txt_intSubunidade.SetFocus
        Else
            
            For i = 0 To dbcintSubunidade.ListCount - 1
                If strDescricao = dbcintSubunidade.ItemData(i) Then
                    dbcintSubunidade.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintSubunidade.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intTipoCredito_GotFocus()
MarcaCampo txt_intTipoCredito
End Sub

Private Sub txt_intTipoCredito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intTipoCredito
End Sub

Private Sub txt_intTipoCredito_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Len(txt_intTipoCredito) > 0 Then
        
        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrTipoCredito, txt_intTipoCredito.Text)
        
        If Len(strDescricao) = 0 Then
            dbcintTipoCredito.ListIndex = -1
            txt_intTipoCredito.SetFocus
        Else
            
            For i = 0 To dbcintTipoCredito.ListCount - 1
                If strDescricao = dbcintTipoCredito.ItemData(i) Then
                    dbcintTipoCredito.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintTipoCredito.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intUnidadeOrcamentaria_Change()
    'txt_intSubunidade.Text = ""
End Sub

Private Sub txt_intUnidadeOrcamentaria_GotFocus()
MarcaCampo txt_intUnidadeOrcamentaria
End Sub

Private Sub txt_intUnidadeOrcamentaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intUnidadeOrcamentaria
End Sub

Private Sub txt_intUnidadeOrcamentaria_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset

    
    If Val(txt_intUnidadeOrcamentaria) > 0 Then
        
        If dbcintOrgao.ListIndex >= 0 Then
            strDescricao = RetornaCodigo("PkID", "strCodigo", gstrUnidadeOrcamentaria, txt_intUnidadeOrcamentaria.Text, " AND intOrgao = " & dbcintOrgao.ItemData(dbcintOrgao.ListIndex))
        Else
        'If Len(Trim(txtPKID.Text)) > 0 Then
            strSQL = ""
            strSQL = strSQL & " SELECT intUnidadeOrcamentaria "
            strSQL = strSQL & "   FROM " & gstrProgramaDeTrabalho
            strSQL = strSQL & "  WHERE pkid = " & txtPKID.Text
            
            Set gobjBanco = New clsBanco
    
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        
            If Not adoResultado.EOF Then
                strDescricao = adoResultado!INTUNIDADEORCAMENTARIA
            End If
        
        End If
    End If
        
        If Len(strDescricao) = 0 Then
            dbcintUnidadeOrcamentaria.ListIndex = -1
            txt_intUnidadeOrcamentaria.SetFocus
        Else
            
            For i = 0 To dbcintUnidadeOrcamentaria.ListCount - 1
                If strDescricao = dbcintUnidadeOrcamentaria.ItemData(i) Then
                    dbcintUnidadeOrcamentaria.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        dbcintUnidadeOrcamentaria.ListIndex = -1
    End If
    
End Sub

Private Sub txt_intVinculo_GotFocus()
MarcaCampo txt_intVinculo
End Sub

Private Sub txt_intVinculo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_intVinculo
End Sub

Private Sub txt_intVinculo_LostFocus()
    Dim strDescricao As String
    Dim i As Integer
    
    If Len(txt_intVinculo) > 0 Then
        
        strDescricao = RetornaCodigo("PkID", "strCodigo", gstrVinculo, txt_intVinculo.Text)
        
        If Len(strDescricao) = 0 Then
            cbointVinculo.ListIndex = -1
            txt_intVinculo.SetFocus
        Else
            
            For i = 0 To cbointVinculo.ListCount - 1
                If strDescricao = cbointVinculo.ItemData(i) Then
                    cbointVinculo.ListIndex = i
                    Exit For
                End If
            Next
            
        End If
        
    Else
        cbointVinculo.ListIndex = -1
    End If
    
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

Public Function strQueryRelatorio()
    
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PT.intCodigoReduzido AS CodigoReduzido, PT.strCodigo AS CódigoProgramaTrabalho, OG.strDescricao AS Orgao, "
    strSQL = strSQL & "UO.strDescricao AS UnidadeOrcamentaria, SU.strDescricao AS Subunidade, TC.strDescricao AS TipoCredito, FG.strDescricao AS FuncaoGoverno, MO.strDescricao Modalidade, "
    strSQL = strSQL & "SF.strDescricao AS Subfuncao, PG.strDescricao AS Programa, SP.strDescricao AS Subprograma, PJ.strDescricao AS ProjetoAtividade, ED.strDescricao AS ElementoDespesa, "
    strSQL = strSQL & "PT.dblValor AS Valor "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT, "
    strSQL = strSQL & gstrOrgao & " OG, "
    strSQL = strSQL & gstrUnidadeOrcamentaria & " UO, "
    strSQL = strSQL & gstrSubUnidade & " SU, "
    strSQL = strSQL & gstrTipoCredito & " TC, "
    strSQL = strSQL & gstrFuncaoDoGoverno & " FG, "
    strSQL = strSQL & gstrSubFuncaoGoverno & " SF, "
    strSQL = strSQL & gstrPrograma & " PG, "
    strSQL = strSQL & gstrSubPrograma & " SP, "
    strSQL = strSQL & gstrProjeto & " PJ, "
    strSQL = strSQL & gstrModalidade & " MO, "
    strSQL = strSQL & gstrElementoDespesa & " ED "
    
    strSQL = strSQL & "WHERE PT.intOrgao " & strOUTJSQLServer & "= OG.PKId " & strOUTJOracle
    
    strSQL = strSQL & "AND PT.intUnidadeOrcamentaria " & strOUTJSQLServer & "= UO.PKId " & strOUTJOracle
    
    strSQL = strSQL & "AND PT.intSubunidade " & strOUTJSQLServer & "= SU.PKId " & strOUTJOracle
    
    strSQL = strSQL & "AND PT.intTipoCredito " & strOUTJSQLServer & "= TC.PKId " & strOUTJOracle
    
    strSQL = strSQL & "AND PT.intFuncao " & strOUTJSQLServer & "= FG.PKId " & strOUTJOracle
    
    strSQL = strSQL & "AND PT.intSubFuncao " & strOUTJSQLServer & "= SF.PKId " & strOUTJOracle
    
    strSQL = strSQL & "AND PT.intPrograma " & strOUTJSQLServer & "= PG.PKId " & strOUTJOracle
    
    strSQL = strSQL & "AND PT.intSubPrograma " & strOUTJSQLServer & "= SP.PKId " & strOUTJOracle
    
    strSQL = strSQL & "AND PT.intProjetoAtividade " & strOUTJSQLServer & "= PJ.PKId " & strOUTJOracle
    strSQL = strSQL & "AND PT.intModalidade " & strOUTJSQLServer & "= MO.PKId " & strOUTJOracle
    strSQL = strSQL & "AND PT.intElementoDespesa " & strOUTJSQLServer & "= ED.PKId " & strOUTJOracle
    strSQL = strSQL & "AND PT.intExercicio = " & Val(txtintExercicio) & " "
    
    'Codigo Reduzido
    If Trim(txtintCodigoReduzido.Text) <> "" Then
        strSQL = strSQL & " AND PT.intCodigoReduzido = " & Trim(txtintCodigoReduzido)
    End If
    
    strSQL = strSQL & "ORDER BY PT.intCodigoReduzido, PT.strCodigo, OG.strDescricao, UO.strDescricao, "
    strSQL = strSQL & "SU.strDescricao, TC.strDescricao, FG.strDescricao, SF.strDescricao, PG.strDescricao, "
    strSQL = strSQL & "SP.strDescricao, PJ.strDescricao, ED.strDescricao, PT.dblValor"
    
    strQueryRelatorio = strSQL
    
End Function

Private Function RetornaCodigo(strCampoRetorno As String, strCampoPesquisa As String, strTabela As String, strValorDePesquisa As String, Optional strANDQuery As String) As String
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT " & strCampoRetorno & " Retorno FROM " & strTabela
    strSQL = strSQL & " WHERE " & strCampoPesquisa & " = '" & strValorDePesquisa & "'"
    
    If Len(strANDQuery) > 0 Then strSQL = strSQL & strANDQuery
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        
        If Not adoResultado.EOF Then
            RetornaCodigo = adoResultado("Retorno")
        Else
            RetornaCodigo = Space$(0)
        End If
        
    End If
    
    adoResultado.Close: Set adoResultado = Nothing
    
End Function

Public Sub AtualizaListaPosElementoDespesa()
    Dim adoResultado As ADODB.Recordset
    
    VerificaListaAutomatica "", tdb_Lista, strQuery
    txt_total = tdb_Lista.Columns("dblTotal")
    
    LeDaTabelaParaObj gstrElementoDespesa, dbcintElementoDespesa
    
    If intPkIdRow > 0 Then
        
        Set adoResultado = tdb_Lista.DataSource
        adoResultado.Find "PKId = '" & intPkIdRow & "'"
        tdb_Lista.MarqueeStyle = dbgHighlightRow
        
    End If
    
End Sub

Public Sub AtualizaListaPosProjetoAtividade()
    Dim adoResultado As ADODB.Recordset
    
    VerificaListaAutomatica "", tdb_Lista, strQuery
    txt_total = tdb_Lista.Columns("dblTotal")
    
    LeDaTabelaParaObj gstrProjeto, dbcintProjetoAtividade
    
    If intPkIdRow > 0 Then
        
        Set adoResultado = tdb_Lista.DataSource
        adoResultado.Find "PKId = '" & intPkIdRow & "'"
        tdb_Lista.MarqueeStyle = dbgHighlightRow
        
    End If
    
End Sub

Private Sub LimpaObjetos()
    mblnAlterando = False
    txt_intOrgao = ""
    txt_intSubunidade = ""
    txt_intFuncao = ""
    txt_intPrograma = ""
    txt_intProjetoAtividade = ""
    txt_intVinculo = ""
    txt_intUnidadeOrcamentaria = ""
    txt_intTipoCredito = ""
    txt_intSubFuncao = ""
    txt_intSubPrograma = ""
    txt_intElementoDespesa = ""
    'txt_codEvento = ""
    txt_intFonteRecurso = ""
    txtstrCodigo = ""
    txt_Grupo = ""
    dbcintSubunidade.Text = ""
    dbcintOrgao.Text = ""
    dbcintUnidadeOrcamentaria.Text = ""
    dbcintSubunidade.Text = ""
    dbcintTipoCredito.Text = ""
    dbcintFuncao.Text = ""
    dbcintSubFuncao.Text = ""
    dbcintPrograma.Text = ""
    dbcintSubPrograma.Text = ""
    dbcintProjetoAtividade.Text = ""
    dbcintElementoDespesa.Text = ""
    'cbointEvento.Text = ""
    cbointVinculo.Text = ""
    cbointFonteRecurso.Text = ""
    txt_strCodigoSaldo = ""
    txt_saldoini = ""
    txt_Empenhado = ""
    txt_anulado = ""
    txt_Suplementado = ""
    txt_liquidado = ""
    txt_Pago = ""
    txt_Reservado = ""
    txt_SaldoParaEmpenho = ""
    txt_intConvenio.Text = ""
End Sub

Private Function STRFONTERECURSO() As String
    Dim adoResultado As ADODB.Recordset
    If tdb_Lista.Columns("F. de Recurso") <> "" Then
        Dim strSQL As String
        strSQL = ""
        strSQL = strSQL & "SELECT  Pkid"
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrFonteRecurso
        strSQL = strSQL & " WHERE  strCodigo = '" & tdb_Lista.Columns("F. de Recurso") & "'"
        strSQL = strSQL & " AND intExercicio = " & txtintExercicio
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                STRFONTERECURSO = adoResultado.Fields("Pkid")
            End If
        End If
    Else
        STRFONTERECURSO = "NULL"
    End If
    
End Function

Private Function blnCadProjetoAtividade() As Boolean
    Dim strSQL       As String
    Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT " & gstrTOPnSQLServer(1) & " Pkid"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrProgramaDeTrabalho
    strSQL = strSQL & " WHERE intOrgao <> " & dbcintOrgao.ItemData(dbcintOrgao.ListIndex) & " AND"
    strSQL = strSQL & " intProjetoAtividade = " & dbcintProjetoAtividade.ItemData(dbcintProjetoAtividade.ListIndex) & " AND"
    If blnOrcamento = True Then
        strSQL = strSQL & " intExercicio = " & gintExercicio
    Else
        strSQL = strSQL & " intExercicio = " & gintExercicio + 1
    End If
    
    strSQL = gstrTOPnOracle(strSQL, 1)
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            blnCadProjetoAtividade = False
        Else
            blnCadProjetoAtividade = True
        End If
    End If
    
    adoResultado.Close: Set adoResultado = Nothing
    
End Function

Private Sub txtintCodigoReduzido_GotFocus()
    gstrProximoCodigo txtintCodigoReduzido, gstrProgramaDeTrabalho, "intCodigoReduzido", gintCodSeguranca, "intExercicio", Val(gintExercicio), , , , , "intExercicio", Val(gintExercicio)
    MarcaCampo txtintCodigoReduzido
End Sub

Private Sub txtintCodigoReduzido_KeyPress(KeyAscii As Integer)
    'ORC1599 - F3RN4NDØ
    'CaracterValido KeyAscii, "N"
    If KeyAscii = 45 Then KeyAscii = 0
    gstrLimitaCampoValor txtintCodigoReduzido, KeyAscii, 9, 0
End Sub

Private Sub txtintCodigoReduzido_LostFocus()
    
    Dim adoResultado As New ADODB.Recordset
    Dim strSQL       As String
    
    If Len(Trim(txtintCodigoReduzido)) > 0 Then
        strSQL = "SELECT * FROM " & gstrProgramaDeTrabalho
        strSQL = strSQL & " WHERE intCodigoReduzido = " & Trim(txtintCodigoReduzido)
        strSQL = strSQL & " AND intExercicio = " & txtintExercicio
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                If .EOF = False Then
                    
                    TrocaCorObjeto txt_codEvento, True
                    TrocaCorObjeto cbointEvento, True
                    TrocaCorObjeto cmd_Evento, True
                    txtPKID = (!Pkid)
                    
                    dbcintOrgao.ListIndex = gintIndiceCBO(dbcintOrgao, !intOrgao)
                    dbcintUnidadeOrcamentaria.ListIndex = gintIndiceCBO(dbcintUnidadeOrcamentaria, !INTUNIDADEORCAMENTARIA)
                    LeDaTabelaParaObj gstrSubUnidade, dbcintSubunidade, "SELECT Pkid, strDescricao FROM " & gstrSubUnidade & " WHERE intOrgao = " & gstrItemData(dbcintOrgao) & _
                    " AND intUnidadeOrcamentaria = " & gstrItemData(dbcintUnidadeOrcamentaria)
                    dbcintSubunidade.ListIndex = gintIndiceCBO(dbcintSubunidade, !INTSUBUNIDADE)
                    dbcintElementoDespesa.ListIndex = gintIndiceCBO(dbcintElementoDespesa, !intElementoDespesa)
                    dbcintFuncao.ListIndex = gintIndiceCBO(dbcintFuncao, !INTFUNCAO)
                    dbcintPrograma.ListIndex = gintIndiceCBO(dbcintPrograma, !INTPROGRAMA)
                    dbcintProjetoAtividade.ListIndex = gintIndiceCBO(dbcintProjetoAtividade, !INTPROJETOATIVIDADE)
                    dbcintSubFuncao.ListIndex = gintIndiceCBO(dbcintSubFuncao, !INTSUBFUNCAO)
                    dbcintSubPrograma.ListIndex = gintIndiceCBO(dbcintSubPrograma, !intSubPrograma)
                    dbcintTipoCredito.ListIndex = gintIndiceCBO(dbcintTipoCredito, !intTipoCredito)
                    cbointEvento.ListIndex = gintIndiceCBO(cbointEvento, !intEvento)
                    cbointFonteRecurso.ListIndex = gintIndiceCBO(cbointFonteRecurso, !INTFONTERECURSO)
                    cbointVinculo.ListIndex = gintIndiceCBO(cbointVinculo, !intVinculo)
                    dbcintModalidade.ListIndex = gintIndiceCBO(dbcintModalidade, !intModalidade)
                    
                    FormataCodigoTrabalho
                    txtdblValor = gstrConvVrDoSql(!dblValor)
                    If Len(Trim(txtdblValor)) = 0 Then
                        txtdblValor = "0,00"
                    End If
                    'ORC1599 - F3RN4NDØ
                    'mblnAlterando = True
                End If
            End With
        End If
    End If
    
End Sub
Private Function blnDadosOk() As Boolean
    
    'CÓDIGO REDUZIDO
    If Len(Trim(txtintCodigoReduzido)) = 0 And blnOrcamento = True Then
        ExibeMensagem "É necessário informar o Código Reduzido."
        txtintCodigoReduzido.SetFocus
        Exit Function
    'ÓRGÃO
    ElseIf dbcintOrgao.ListIndex = -1 Then
        ExibeMensagem "É necessário informar o Orgão."
        dbcintOrgao.SetFocus
        Exit Function
    'UNIDADE ORÇAMENTÁRIA
    ElseIf dbcintUnidadeOrcamentaria.ListIndex = -1 Then
        ExibeMensagem "É necessário informar a Unidade Orçamentária."
        dbcintUnidadeOrcamentaria.SetFocus
        Exit Function
    'SUB UNIDADE
    ElseIf dbcintSubunidade.ListIndex = -1 And Len(Trim(dbcintSubunidade.Text)) <> 0 Then
       'ExibeMensagem "É necessário informar a Subunidade."
       txt_intSubunidade.Text = ""
       dbcintSubunidade.Text = ""
       'Exit Function
    'TIPO CRÉDITO
    ElseIf dbcintTipoCredito.ListIndex = -1 Then
        ExibeMensagem "É necessário informar o Tipo de Credito."
        dbcintTipoCredito.SetFocus
        Exit Function
    'FUNÇÃO
    ElseIf dbcintFuncao.ListIndex = -1 Then
        ExibeMensagem "É necessário informar a Função."
        dbcintFuncao.SetFocus
        Exit Function
    'SUB FUNÇÃO
    ElseIf dbcintSubFuncao.ListIndex = -1 Then
        ExibeMensagem "É necessário informar a SubFunção."
        dbcintSubFuncao.SetFocus
        Exit Function
    'PROGRAMA
    ElseIf dbcintPrograma.ListIndex = -1 Then
        ExibeMensagem "É necessário informar o Programa."
        dbcintPrograma.SetFocus
        Exit Function
    'SUB PROGRAMA
    ElseIf dbcintSubPrograma.ListIndex = -1 And Len(Trim(dbcintSubPrograma.Text)) <> 0 Then
        Me.txt_intSubPrograma.Text = ""
        dbcintSubPrograma.Text = ""
    'PROJETO/ATIVIDADE
    ElseIf dbcintProjetoAtividade.ListIndex = -1 Then
        ExibeMensagem "É necessário informar o Projeto/Atividade."
        dbcintProjetoAtividade.SetFocus
        Exit Function
    'EVENTO CONTÁBIL
    ElseIf cbointEvento.ListIndex = -1 And blnOrcamento = True Then
        'Verifica o Evento Somente no Cadastro de Novos Programas de Trabalho
        If mblnAlterando = False Then
            ExibeMensagem "É necessário informar o Evento Contábil."
            If cbointEvento.Enabled Then cbointEvento.SetFocus
            Exit Function
        End If
    'VÍNCULO
    ElseIf cbointVinculo.ListIndex = -1 And Len(Trim(cbointVinculo.Text)) <> 0 Then
        txt_intVinculo.Text = ""
        cbointVinculo.Text = ""
        
    'ELEMENTO DA DESPESA
    ElseIf dbcintElementoDespesa.ListIndex = -1 Then
        ExibeMensagem "É necessário informar o Elemento de Despesa."
        dbcintElementoDespesa.SetFocus
        Exit Function
    ElseIf blnOrcamento = True And Not VerificaEvento Then
        ExibeMensagem "Este Elemento de Despesa não é compatível com o Evento Contábil."
        txt_intElementoDespesa = ""
        dbcintElementoDespesa.ListIndex = -1
        txt_intElementoDespesa.SetFocus
        Exit Function
    'FONTE DE RECURSO
    ElseIf cbointFonteRecurso.ListIndex = -1 Then
        ExibeMensagem "É necessário informar a Fonte de Recurso."
        cbointFonteRecurso.SetFocus
        Exit Function
    'MODALIDADE DE APLICAÇÃO
    ElseIf dbcintModalidade.ListIndex = -1 Then
        ExibeMensagem "É necessário informar a Modalidade de Aplicação."
        dbcintModalidade.SetFocus
        Exit Function
        
    End If
    
    
    If Not mblnAlterando Then
        If Not VerificaCodigoReduzido(Trim(txtintCodigoReduzido), IIf(blnOrcamento = True, gintExercicio, gintExercicio + 1)) Then
            If blnOrcamento Then
                ExibeMensagem "Código reduzido já utilizado."
                If txtintCodigoReduzido.Enabled = True Then txtintCodigoReduzido.SetFocus
                Exit Function
            End If
        End If
        If Not VerificaProgramaDeTrabalho(IIf(blnOrcamento = True, gintExercicio, gintExercicio + 1)) Then
            ExibeMensagem "Programa de Trabalho já existente."
            If dbcintOrgao.Enabled = True Then dbcintOrgao.SetFocus
            Exit Function
        End If
    End If
    
    'FORMATAÇÃO DO VALOR
    If Len(Trim(txtdblValor)) = 0 Then
        txtdblValor = "0,00"
    End If
    
    blnDadosOk = True
    
End Function

Private Function strQueryEvento() As String
    
    Dim strSQL As String
    
    
    strSQL = "SELECT EV.PKID, EV.strDescricao, EVC.intContaContabil, PC.strContaContabil"
    strSQL = strSQL & " FROM " & gstrEvento & " EV, " & gstrEventoContaContabilCredito & " EVC,"
    strSQL = strSQL & gstrPlanoConta & " PC"
    strSQL = strSQL & " WHERE EV.Pkid = EVC.intEvento AND "
    strSQL = strSQL & " EVC.intContaContabil = PC.PKid AND "
    strSQL = strSQL & strSUBSTRING & "(PC.strContaContabil,1,1) = '3' AND "
    strSQL = strSQL & " EV.intTipoEvento = 0"
    
    strQueryEvento = strSQL
    
End Function

Private Function VerificaEvento() As Boolean
    Dim strSQL       As String
    Dim adoResultado As New ADODB.Recordset
    
    VerificaEvento = True
    
    strSQL = "SELECT EV.PKID, EV.strDescricao, EVC.intContaContabil, PC.strContaContabil"
    strSQL = strSQL & " FROM " & gstrEvento & " EV, " & gstrEventoContaContabilCredito & " EVC,"
    strSQL = strSQL & gstrPlanoConta & " PC"
    strSQL = strSQL & " WHERE  EV.Pkid = EVC.intEvento AND EV.PKID = " & gstrItemData(cbointEvento) & " AND "
    strSQL = strSQL & " EVC.intContaContabil = PC.PKid AND "
    strSQL = strSQL & strSUBSTRING & "(PC.strContaContabil,1,1) = '3' AND "
    'strSQL = strSQL & strSUBSTRING & "(PC.strContaContabil,2,1) = '" & Mid(Trim(txt_intElementoDespesa), 1, 1) & "' AND "
    strSQL = strSQL & strSUBSTRING & "(PC.strContaContabil,2," & Len(BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoDespesa, "C", 0)) & ") = '" & _
    Mid(txt_intElementoDespesa, 1, Len(BuscaCodigosPeloEvento(gstrItemData(cbointEvento), gstrDigitoDespesa, "C", 0))) & "' AND "
    strSQL = strSQL & " EV.intTipoEvento = 0"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            If .EOF = True Then
                VerificaEvento = False
            End If
        End With
    End If
End Function

Private Function VerificaCodigoReduzido(intCodigoReduzido As String, intExercicio As Integer) As Boolean
    
    Dim strSQL         As String
    Dim adoResultado   As ADODB.Recordset
    
    If intCodigoReduzido <> "" Then
        VerificaCodigoReduzido = True
        
        
        strSQL = ""
        
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "intCodigoReduzido, "
        strSQL = strSQL & "intExercicio "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "intCodigoReduzido = " & intCodigoReduzido
        strSQL = strSQL & "AND PT.intexercicio = " & intExercicio
        
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If adoResultado.RecordCount > 0 Then
                If adoResultado!intCodigoReduzido = intCodigoReduzido And adoResultado!intExercicio = gintExercicio Then
                    VerificaCodigoReduzido = False
                End If
            End If
            
        End If
    Else
        VerificaCodigoReduzido = False
    End If
    
End Function

Private Function VerificaProgramaDeTrabalho(intExercicio As Integer) As Boolean
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim strSubprograma  As String
    Dim strSubUnidade   As String
    Dim strVinculo      As String
    
    VerificaProgramaDeTrabalho = True
    
    'Tratamento de campos nao obrigatorios
    If dbcintSubPrograma.ListIndex = -1 Then
        strSubprograma = "IS NULL"
    Else
        strSubprograma = " = " & dbcintSubPrograma.ItemData(dbcintSubPrograma.ListIndex)
    End If
    If dbcintSubunidade.ListIndex = -1 Then
        strSubUnidade = "IS NULL"
    Else
        strSubUnidade = " = " & dbcintSubunidade.ItemData(dbcintSubunidade.ListIndex)
    End If
    If cbointVinculo.ListIndex = -1 Then
        strVinculo = "IS NULL"
    Else
        strVinculo = " = " & cbointVinculo.ItemData(cbointVinculo.ListIndex)
    End If
    
    strSQL = ""
    
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "intCodigoReduzido, "
    strSQL = strSQL & "intExercicio "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "PT.intOrgao = " & gstrENulo(dbcintOrgao.ItemData(dbcintOrgao.ListIndex), , True)
    strSQL = strSQL & " AND PT.intUnidadeOrcamentaria = " & gstrENulo(dbcintUnidadeOrcamentaria.ItemData(dbcintUnidadeOrcamentaria.ListIndex), , True)
    strSQL = strSQL & " AND PT.intFuncao = " & gstrENulo(dbcintFuncao.ItemData(dbcintFuncao.ListIndex), , True)
    strSQL = strSQL & " AND PT.intSubFuncao = " & gstrENulo(dbcintSubFuncao.ItemData(dbcintSubFuncao.ListIndex), , True)
    strSQL = strSQL & " AND PT.intPrograma = " & gstrENulo(dbcintPrograma.ItemData(dbcintPrograma.ListIndex), , True)
    strSQL = strSQL & " AND PT.intSubPrograma " & strSubprograma
    strSQL = strSQL & " AND PT.intProjetoAtividade = " & gstrENulo(dbcintProjetoAtividade.ItemData(dbcintProjetoAtividade.ListIndex), , True)
    strSQL = strSQL & " AND PT.intElementoDespesa = " & gstrENulo(dbcintElementoDespesa.ItemData(dbcintElementoDespesa.ListIndex), , True)
    strSQL = strSQL & " AND PT.intSubUnidade " & strSubUnidade
    strSQL = strSQL & " AND PT.intTipoCredito = " & gstrENulo(dbcintTipoCredito.ItemData(dbcintTipoCredito.ListIndex), , True)
    strSQL = strSQL & " AND PT.intFonteRecurso = " & gstrENulo(cbointFonteRecurso.ItemData(cbointFonteRecurso.ListIndex), , True)
    strSQL = strSQL & " AND PT.intVinculo " & strVinculo
    If cbointEvento.ListIndex > -1 Then
        strSQL = strSQL & " AND PT.intEvento = " & gstrENulo(cbointEvento.ItemData(cbointEvento.ListIndex), , True)
    End If
    strSQL = strSQL & " AND PT.intModalidade = " & gstrENulo(dbcintModalidade.ItemData(dbcintModalidade.ListIndex), , True)
    strSQL = strSQL & " AND PT.intexercicio = " & intExercicio
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            VerificaProgramaDeTrabalho = False
        End If
    End If
    
End Function

Private Function CarregaConvenio(strPKId As String) As String
    Dim strSQL As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT CV.strCodigo "
    strSQL = strSQL & " From " & gstrModalidade & " MO,"
    strSQL = strSQL & gstrConvenio & " CV"
    strSQL = strSQL & " WHERE MO.Pkid = " & strPKId
    strSQL = strSQL & " AND MO.intConvenio = CV.PKID "
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            If Len(Trim(gstrENulo(adoResultado.Fields("strCodigo").Value))) > 0 Then
                CarregaConvenio = Format(adoResultado.Fields("strCodigo").Value, "00")
            End If
        Else
            CarregaConvenio = "00"
        End If
    End If
End Function
Private Function CarregaModalidade(strCodigoModalidade As String, strCodigoConvenio As String) As String
    Dim strSQL As String
    Dim adoResultado    As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT  MO.PKID, MO.intConvenio"
    strSQL = strSQL & " FROM " & gstrModalidade & " MO, "
    strSQL = strSQL & gstrConvenio & " CV"
    strSQL = strSQL & " WHERE " & gstrCONVERT(CDT_INT, "MO.strcodigo") & " = " & Val(strCodigoModalidade)
    If Len(Trim(txt_intConvenio.Text)) > 0 Then
        If txt_intConvenio.Text <> "00" Then
            strSQL = strSQL & " AND " & gstrCONVERT(CDT_INT, "CV.strcodigo") & " = " & Val(strCodigoConvenio)
            strSQL = strSQL & " AND MO.intConvenio = CV.PKID"
        End If
    End If
    strSQL = strSQL & " GROUP BY MO.PKID, MO.intConvenio"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            CarregaModalidade = gstrENulo(adoResultado.Fields("PKID").Value)
            If Len(Trim(gstrENulo(adoResultado.Fields("intConvenio").Value))) > 0 Then
                txt_intConvenio.Text = Format(adoResultado.Fields("intConvenio").Value, 0)
            Else
                txt_intConvenio.Text = "00"
            End If
        End If
    End If
End Function


