VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadCalculoIssqnArbitrado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de ISSQN Arbitrado"
   ClientHeight    =   6915
   ClientLeft      =   1455
   ClientTop       =   2160
   ClientWidth     =   8595
   Icon            =   "CadCalculoIssqnArbitrado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8595
   Begin VB.TextBox txt_PKId 
      Height          =   285
      Left            =   7590
      TabIndex        =   49
      Top             =   105
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6645
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   120
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   11721
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lançamento de ISSQN Arbitrado"
      TabPicture(0)   =   "CadCalculoIssqnArbitrado.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strComposicaoReceita"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_strInscricaoCadastral"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_intOcorrencia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbc_intOcorrencia"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbc_intComposicaoDaReceita"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdb_Proprietarios"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dbc_strInscricaoCadastral"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra_Proprietario"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Composição da Receita"
      TabPicture(1)   =   "CadCalculoIssqnArbitrado.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdb_Atividades"
      Tab(1).Control(1)=   "fra_Frame"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Emissão de Guias de Arrecadação"
      TabPicture(2)   =   "CadCalculoIssqnArbitrado.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_EmissaoDeGuias"
      Tab(2).ControlCount=   1
      Begin VB.Frame fra_Frame 
         Height          =   1785
         Left            =   -74790
         TabIndex        =   55
         Top             =   420
         Width           =   7995
         Begin VB.TextBox txt_intParcelaFinal 
            Height          =   285
            Left            =   2610
            MaxLength       =   15
            TabIndex        =   7
            Top             =   990
            Width           =   465
         End
         Begin VB.TextBox txt_intParcelaInicial 
            Height          =   285
            Left            =   1770
            MaxLength       =   15
            TabIndex        =   6
            Top             =   990
            Width           =   465
         End
         Begin VB.TextBox txt_intExercício 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6285
            MaxLength       =   4
            TabIndex        =   9
            Top             =   645
            Width           =   735
         End
         Begin VB.TextBox txt_dblValorArbitrado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1770
            MaxLength       =   15
            TabIndex        =   4
            Top             =   270
            Width           =   1305
         End
         Begin VB.TextBox txt_dblAliquota 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   8
            Top             =   1365
            Width           =   1305
         End
         Begin VB.TextBox txt_intIntervalo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6285
            MaxLength       =   3
            TabIndex        =   11
            Top             =   1365
            Width           =   1005
         End
         Begin VB.TextBox txt_dtmDataVencimento 
            Height          =   285
            Left            =   6285
            MaxLength       =   15
            TabIndex        =   10
            Top             =   1005
            Width           =   1005
         End
         Begin VB.TextBox txt_dtmDataPagamento 
            Height          =   285
            Left            =   1770
            MaxLength       =   15
            TabIndex        =   5
            Top             =   645
            Width           =   1305
         End
         Begin VB.Label lbl_ParcelaFinal 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   2280
            TabIndex        =   66
            Top             =   1050
            Width           =   225
         End
         Begin VB.Label lbl_ParcelaIncial 
            AutoSize        =   -1  'True
            Caption         =   "Parcela"
            Height          =   195
            Left            =   1080
            TabIndex        =   65
            Top             =   1035
            Width           =   540
         End
         Begin VB.Label lbl_p1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   3180
            TabIndex        =   63
            Top             =   1440
            Width           =   120
         End
         Begin VB.Label lbl_dtmDataPagamento 
            AutoSize        =   -1  'True
            Caption         =   "Data de Lançamento"
            Height          =   195
            Left            =   165
            TabIndex        =   62
            Top             =   705
            Width           =   1500
         End
         Begin VB.Label lbl_dtmDataVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vencimento"
            Height          =   195
            Left            =   4680
            TabIndex        =   61
            Top             =   1065
            Width           =   1455
         End
         Begin VB.Label lbl_intExercício 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   5460
            TabIndex        =   60
            Top             =   720
            Width           =   675
         End
         Begin VB.Label lbl_dblValorArbitrado 
            AutoSize        =   -1  'True
            Caption         =   "Valor Arbitrado"
            Height          =   195
            Left            =   630
            TabIndex        =   59
            Top             =   345
            Width           =   1035
         End
         Begin VB.Label lbl_dblAliquota 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota sobre Valor"
            Height          =   195
            Left            =   225
            TabIndex        =   58
            Top             =   1440
            Width           =   1440
         End
         Begin VB.Label lbl_intIntervalo 
            AutoSize        =   -1  'True
            Caption         =   "Intervalo entre Parcelas"
            Height          =   195
            Left            =   4455
            TabIndex        =   57
            Top             =   1440
            Width           =   1680
         End
         Begin VB.Label lbl_dias 
            AutoSize        =   -1  'True
            Caption         =   "dias."
            Height          =   195
            Left            =   7350
            TabIndex        =   56
            Top             =   1410
            Width           =   330
         End
      End
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5445
         Left            =   -74790
         TabIndex        =   50
         Top             =   390
         Width           =   7935
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1695
            Left            =   480
            TabIndex        =   53
            Top             =   3030
            Width           =   6945
            Begin VB.TextBox txt_Mensagem2 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   31
               Top             =   0
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   15
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
               TabIndex        =   54
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1695
            Left            =   480
            TabIndex        =   51
            Top             =   1200
            Width           =   6945
            Begin VB.TextBox txt_Mensagem1 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
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
               TabIndex        =   13
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
               TabIndex        =   52
               Top             =   390
               Width           =   780
            End
         End
      End
      Begin VB.Frame fra_Proprietario 
         Height          =   3105
         Left            =   120
         TabIndex        =   32
         Top             =   810
         Width           =   8070
         Begin VB.TextBox txt_strAtividadePrincipal 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1260
            Width           =   4110
         End
         Begin VB.TextBox txt_strAtividadeBasica 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   930
            Width           =   4110
         End
         Begin VB.TextBox txt_strNome 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2235
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   210
            Width           =   5595
         End
         Begin VB.Frame fra_Endereco 
            Caption         =   "Endereço"
            Height          =   1395
            Left            =   120
            TabIndex        =   33
            Top             =   1590
            Width           =   7845
            Begin VB.TextBox txt_UF 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5070
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   990
               Width           =   510
            End
            Begin VB.TextBox txt_Municipio 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5070
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   630
               Width           =   2625
            End
            Begin VB.TextBox txt_Cep 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   6615
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   990
               Width           =   1080
            End
            Begin VB.TextBox txt_Complemento 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   6810
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   270
               Width           =   870
            End
            Begin VB.TextBox txt_Numero 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   270
               Width           =   795
            End
            Begin VB.TextBox txt_Logradouro 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   270
               Width           =   4005
            End
            Begin VB.TextBox txt_Bairro 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   630
               Width           =   3105
            End
            Begin VB.TextBox txt_Distrito 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   990
               Width           =   3525
            End
            Begin VB.Label lbl_strDistritoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               Height          =   195
               Left            =   450
               TabIndex        =   42
               Top             =   1050
               Width           =   480
            End
            Begin VB.Label lbl_intMunicipioC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   4290
               TabIndex        =   41
               Top             =   690
               Width           =   705
            End
            Begin VB.Label lbl_intBairroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   525
               TabIndex        =   40
               Top             =   690
               Width           =   405
            End
            Begin VB.Label lbl_intLogradouroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   330
               Width           =   810
            End
            Begin VB.Label lbl_intNumeroC 
               AutoSize        =   -1  'True
               Caption         =   "Nº"
               Height          =   195
               Left            =   5145
               TabIndex        =   37
               Top             =   330
               Width           =   180
            End
            Begin VB.Label lbl_strComplementoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6270
               TabIndex        =   36
               Top             =   330
               Width           =   480
            End
            Begin VB.Label lbl_intUFC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   4770
               TabIndex        =   35
               Top             =   1065
               Width           =   210
            End
            Begin VB.Label lbl_intCepC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   6240
               TabIndex        =   34
               Top             =   1050
               Width           =   285
            End
         End
         Begin VB.TextBox txt_strCNPJCPFP 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   570
            Width           =   1845
         End
         Begin VB.TextBox txt_intContribuinte 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   210
            Width           =   735
         End
         Begin VB.Label lbl_strAtividadePrincipal 
            AutoSize        =   -1  'True
            Caption         =   "Atividade principal"
            Height          =   195
            Left            =   105
            TabIndex        =   48
            Top             =   1305
            Width           =   1290
         End
         Begin VB.Label lbl_strAtividadeBasica 
            AutoSize        =   -1  'True
            Caption         =   "Atividade básica"
            Height          =   195
            Left            =   225
            TabIndex        =   47
            Top             =   975
            Width           =   1170
         End
         Begin VB.Label lbl_strNome 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   600
            TabIndex        =   44
            Top             =   255
            Width           =   795
         End
         Begin VB.Label lbl_strCNPJCPFP 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   615
            TabIndex        =   43
            Top             =   615
            Width           =   780
         End
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastral 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Proprietarios 
         Height          =   1605
         Left            =   150
         TabIndex        =   3
         Top             =   4860
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   2831
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3043"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2963"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=10557"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=10478"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
         Height          =   315
         Left            =   2010
         TabIndex        =   1
         Top             =   4020
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
         Height          =   3465
         Left            =   -74850
         TabIndex        =   12
         Top             =   2340
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   6112
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=450"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=370"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=12039"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=11959"
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
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H0&"
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
      Begin MSDataListLib.DataCombo dbc_intOcorrencia 
         Height          =   315
         Left            =   2010
         TabIndex        =   2
         Top             =   4440
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lbl_intOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   1095
         TabIndex        =   64
         Top             =   4560
         Width           =   780
      End
      Begin VB.Label lbl_strInscricaoCadastral 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Top             =   570
         Width           =   1350
      End
      Begin VB.Label lbl_strComposicaoReceita 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   4125
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCadCalculoIssqnArbitrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnAlterando                   As Boolean
Dim mblnPrimeiraVez                 As Boolean
Dim mobjAux                         As Object
Dim mblnSelecionou                  As Boolean
Dim blnSelecionouReceita            As Boolean
Dim xarReceita                      As XArrayDB
Dim adoResultado                    As ADODB.Recordset

Private Sub dbc_intComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intMensagem1_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intMensagem1, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intMensagem2_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intMensagem2, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intOcorrencia_Click(Area As Integer)
   DropDownDataCombo dbc_intOcorrencia, Me, Area
End Sub

Private Sub dbc_intOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoCadastral_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoCadastral, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 668
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrSalvar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrCalcularReajuste
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
    mblnAlterando = False
    Set xarReceita = New XArrayDB
    xarReceita.Clear
    xarReceita.ReDim 0, 0, 0, 3
    
'    LeDaTabelaParaObj gstrImobiliario, dbc_strInscricaoCadastral, strQueryEconomico(0)
    dbc_strInscricaoCadastral.Tag = strQueryEconomico(0) & ";E.strInscricaoCadastral"
    
    LeDaTabelaParaObj gstrImobiliario, tdb_Proprietarios, strQueryEconomico(1)
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoDaReceita, " SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita & " WHERE intUtilizacao = 2 ORDER BY strDescricao "
    VerificaMascaraInscricao
    LeDaTabelaParaObj gstrOcorrencia, dbc_intOcorrencia, strQuerryOcorrencia
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem1, strQueryMensagem
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem2, strQueryMensagem
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Function strQueryEconomico(bytObjeto As Byte) As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String

    If bytObjeto = 0 Then 'Data Combo   dbc_strInscricaoCadastral
'        strSQL = " SELECT E.PKId AS PKId, (E.strInscricaoCadastral + ' - ' + C.strNome ) AS Inscricao "
        strSql = " SELECT E.PKId AS PKId, (E.strInscricaoCadastral " & strCONCAT & " ' - ' " & strCONCAT & " C.strNome ) AS Inscricao "
    Else                  'True DBGrid  tdb_Proprietarios
        strSql = " SELECT E.PKId AS PKId, E.strInscricaoCadastral AS strInscricaoCadastral, C.strNome AS strNome "
    End If

    strSql = strSql & " FROM " & gstrEconomico & " E," & _
            gstrContribuinte & " C " & _
            " WHERE C.PKId = E.intContribuinte " & _
            " AND  E.dtmDataBaixa IS NULL "
'            " ORDER BY CONVERT(NUMERIC, E.strInscricaoCadastral) "
    strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "E.strInscricaoCadastral")
    strQueryEconomico = strSql
End Function

Sub VerificaMascaraInscricao()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
    
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
    dbc_strInscricaoCadastral.Text = strMascara
    
End Sub

Private Sub dbc_strInscricaoCadastral_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, Area
    If Area = 2 And dbc_strInscricaoCadastral.MatchedWithList Then
        txt_PKId.Text = dbc_strInscricaoCadastral.BoundText
        BuscarDadosProprietario
    End If
End Sub

Private Function BuscarDadosProprietario()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT A.PKId, CO.PKId AS CodigoContribuinte, CO.strNome, LG.strDescricao AS Logradouro, A.intNumero, "
    strSql = strSql & " A.strComplemento , A.intCep, BA.strDescricao AS Bairro, UF.strSigla AS UF, CO.strCNPJCPF, "
    strSql = strSql & " CO.strDistritoC, CI.strDescricao AS Municipio "
    strSql = strSql & " FROM "
    
'    strSQL = strSQL & gstrEconomico & " AS A, "
    strSql = strSql & gstrEconomico & " A, "
    
'    strSQL = strSQL & gstrContribuinte & " AS CO, "
    strSql = strSql & gstrContribuinte & " CO, "
'    strSQL = strSQL & gstrComposicaoDaReceita & " AS CR, "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
'    strSQL = strSQL & gstrLogradouro & " AS LG, "
    strSql = strSql & gstrLogradouro & " LG, "
'    strSQL = strSQL & gstrCidade & " AS CI, "
    strSql = strSql & gstrCidade & " CI, "
'    strSQL = strSQL & gstrBairro & " AS BA, "
    strSql = strSql & gstrBairro & " BA, "
'    strSQL = strSQL & gstrUF & " AS UF "
    strSql = strSql & gstrUF & " UF "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = A.intContribuinte "
    strSql = strSql & " AND CI.PKId = CO.intMunicipioC "
    strSql = strSql & " AND LG.PKId = A.intLogradouro "
    strSql = strSql & " AND BA.PKId = A.intBairro "
    strSql = strSql & " AND UF.PKId = CO.intUF "
    strSql = strSql & " AND A.PKID = " & txt_PKId.Text
    
    LimpaEndereco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                dbc_strInscricaoCadastral.BoundText = !Pkid
                txt_strCNPJCPFP = gstrCGCCPFFormatado(!StrCnpjCpf)
                txt_strNome.Text = !STRNOME
                txt_intContribuinte.Text = !CodigoContribuinte
                txt_Logradouro.Text = !Logradouro
                txt_Numero.Text = !INTNUMERO
                txt_Complemento.Text = gstrENulo(!STRCOMPLEMENTO)
                txt_Bairro.Text = gstrENulo(!Bairro)
                txt_Municipio.Text = gstrENulo(!Municipio)
                txt_Distrito.Text = gstrENulo(!strDistritoC)
                txt_UF.Text = gstrENulo(!UF)
                txt_Cep.Text = gstrCEPFormatado(!INTCEP)
            End With
        End If
    End If
    
    
    Set gobjBanco = New clsBanco

    strSql = ""
'    strSql = strSql & " SELECT CONVERT(VARCHAR, AEC.intCodigo) + ' - ' + AEC.strDescricao AS strAtividadePrincipal "
    strSql = strSql & " SELECT " & gstrCONVERT(CDT_VARCHAR, "AEC.intCodigo") & strCONCAT & " ' - ' " & strCONCAT & " AEC.strDescricao AS strAtividadePrincipal "
    strSql = strSql & " FROM "
'    strSQL = strSQL & gstrEconomico & " AS A, "
    strSql = strSql & gstrEconomico & " A, "
'    strSQL = strSQL & gstrAtividadeDaEmpresa & " AS AE, "
    strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
    'StrSql = StrSql & gstrAtividade & " AS ATI, "
'    strSQL = strSQL & gstrAtividadeEC & " AS AEC "
    strSql = strSql & gstrAtividadeEC & " AEC "
    strSql = strSql & " WHERE "
    strSql = strSql & " AE.intEconomico = A.PKId AND AE.blnPrincipal = 1 "
    'StrSql = StrSql & " AND ATI.PKId = AE.intAtividade AND ATI.intUtilizacao IN (4,5,6) "
    strSql = strSql & " AND AEC.PKId = AE.intAtividade "
    strSql = strSql & " AND A.PKID = " & txt_PKId.Text
    
    txt_strAtividadePrincipal.Text = ""
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txt_strAtividadePrincipal.Text = !strAtividadePrincipal
            End With
        End If
    End If
    
    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT AB.strDescricao AS strAtividadeBasica "
    strSql = strSql & " FROM "
'    strSQL = strSQL & gstrEconomico & " AS A, "
    strSql = strSql & gstrEconomico & " A, "
'    strSQL = strSQL & gstrAtividadeBasica & " AS AB "
    strSql = strSql & gstrAtividadeBasica & " AB "
    strSql = strSql & " WHERE AB.PKId = A.intAtividadeBasica "
    strSql = strSql & " AND A.PKID = " & txt_PKId.Text
    
    txt_strAtividadeBasica.Text = ""
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txt_strAtividadeBasica.Text = !strAtividadeBasica
            End With
        End If
    End If
    
End Function

Private Function LimpaEndereco()
    txt_strCNPJCPFP = ""
    txt_strNome.Text = ""
    txt_intContribuinte.Text = ""
    txt_Logradouro.Text = ""
    txt_Numero.Text = ""
    txt_Complemento.Text = ""
    txt_Bairro.Text = ""
    txt_Municipio.Text = ""
    txt_Distrito.Text = ""
    txt_UF.Text = ""
    txt_Cep.Text = ""
End Function

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
    If Area = 2 And dbc_intComposicaoDaReceita.MatchedWithList Then
        MontaAtividade dbc_intComposicaoDaReceita.BoundText
    End If
End Sub

Private Sub MontaAtividade(intComposicaoReceita As Integer)
Dim strSql As String
Dim adoRec As ADODB.Recordset
Dim varAux As String

On Error GoTo Err_Handle

Set xarReceita = New XArrayDB
xarReceita.Clear

xarReceita.ReDim 0, 0, 0, 2

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
            xarReceita.ReDim 0, .RecordCount - 1, 0, 2
            Do While Not .EOF
                varAux = !Pkid
                xarReceita(.AbsolutePosition - 1, 0) = varAux
                
                varAux = False
                xarReceita(.AbsolutePosition - 1, 1) = varAux
            
                varAux = !strDescricao
                xarReceita(.AbsolutePosition - 1, 2) = varAux
                
                .MoveNext
            Loop
        End If
    End With
End If

Set tdb_Atividades.Array = xarReceita
tdb_Atividades.Rebind
tdb_Atividades.Refresh

Exit Sub
Err_Handle:

End Sub

Private Sub txt_dblAliquota_GotFocus()
    MarcaCampo txt_dblAliquota
End Sub

Private Sub txt_dblAliquota_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_dblAliquota
End Sub

Private Sub txt_dblValorArbitrado_GotFocus()
    MarcaCampo txt_dblValorArbitrado
End Sub

Private Sub txt_dblValorArbitrado_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_dblValorArbitrado
End Sub

Private Sub txt_dblValorArbitrado_LostFocus()
    txt_dblValorArbitrado = gstrConvVrDoSql(txt_dblValorArbitrado)
End Sub

Private Sub txt_dtmDataPagamento_GotFocus()
    MarcaCampo txt_dtmDataPagamento
End Sub

Private Sub txt_dtmDataPagamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataPagamento
End Sub

Private Sub txt_dtmDataPagamento_LostFocus()
    txt_dtmDataPagamento.Text = gstrDataFormatada(txt_dtmDataPagamento.Text)
End Sub

Private Sub txt_dtmDataVencimento_GotFocus()
    MarcaCampo txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_LostFocus()
    txt_dtmDataVencimento.Text = gstrDataFormatada(txt_dtmDataVencimento.Text)
End Sub

Private Sub txt_intExercício_GotFocus()
    MarcaCampo txt_intExercício
End Sub

Private Sub txt_intExercício_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercício
End Sub

Private Sub txt_intIntervalo_GotFocus()
    MarcaCampo txt_intIntervalo
End Sub

Private Sub txt_intIntervalo_KeyPress(KeyAscii As Integer)
        CaracterValido KeyAscii, "N", txt_intIntervalo
End Sub

Private Sub tdb_Atividades_AfterColUpdate(ByVal ColIndex As Integer)
    tdb_Atividades.Update
End Sub

Private Sub tdb_Atividades_AfterUpdate()
    tdb_Atividades.Update
End Sub

Private Sub tdb_Proprietarios_Click()
    mblnPrimeiraVez = True
    With tdb_Proprietarios
        If Not .EOF And Not .BOF Then
            If .Bookmark = 1 Then
                tdb_Proprietarios_RowColChange 0, 0
            End If
        End If
    End With
End Sub

Private Sub tdb_Proprietarios_FilterChange()
    gblnFilraCampos tdb_Proprietarios
End Sub

Private Sub tdb_Proprietarios_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Proprietarios
        If Not .EOF And Not .BOF Then
            txt_PKId.Text = .Columns("PKID").Value

            If mblnPrimeiraVez Then

                BuscarDadosProprietario

                gCorLinhaSelecionada tdb_Proprietarios

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar

                mblnAlterando = True
            End If

        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSql              As String
    Dim blnExiteLancamento  As Boolean
    Dim strInscricao        As String
    If tab_3DPasta.Tab = 2 Then
        If UCase(strModoOperacao) = UCase(gstrNovo) Then
            LimpaObjetos
        End If
        If UCase(strModoOperacao) = UCase(gstrFechar) Then
            Unload Me
        End If
    End If
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosGuiaOK = True Then
            strInscricao = Trim(Left(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, "-", vbTextCompare) - 1))
            strSql = gstrQueryRelatorioGuiaDeArrecadacao(blnExiteLancamento, strInscricao, strInscricao, txt_intExercício.Text, dbc_intComposicaoDaReceita.BoundText, , , Val(txt_intParcelaInicial), Val(txt_intParcelaFinal))
            If blnExiteLancamento Then
                Set gfrmFormularioQueEstaImprimindoGuia = Me
                rptGuiaDeArrecadacaoMunicipal.strImposto = dbc_intComposicaoDaReceita.Text
                ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, strSql
            End If
        End If
    ElseIf UCase(strModoOperacao) = gstrCalcularReajuste Then
            CalculoLancamento
    ElseIf strModoOperacao = gstrPreencherLista Then
        PreencherListaDeOpcoes dbc_strInscricaoCadastral
    End If
    
End Sub

'====================

Private Sub CalculoLancamento()

'******************************************************************************************
' Data: 07/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 07/05/2003
' Alteração: - Substituição da chamada à função CriaADO por uma chamada à função
'            ExecuteStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql            As String
'    Dim adoResultado      As ADODB.Recordset
    Dim adoParameters      As ADODB.Parameters
    Dim intNumeroParcelas As Integer
    Dim strInscricao      As String
    
    Screen.MousePointer = vbHourglass
    If blnDadosOk Then
        strInscricao = Trim(Left(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, "-", vbTextCompare) - 1))
'        strSql = "sp_CalculoParaUsuario '" & strInscricao & "','" & strPKId & _
'                 "',2," & dbc_intComposicaoDaReceita.BoundText & _
'                 ",0,0,0,'" & gstrConvVrParaSql(Abs(txt_dblValorArbitrado)) & ","
'                 gstrConvVrParaSql(txt_dblAliquota.Text) & ", @dblValor OUTPUT'"
        strSql = gstrStoredProcedure("sp_CalculoParaUsuario", "'" & strInscricao & "','" & strPKId & _
                 "',2," & dbc_intComposicaoDaReceita.BoundText & _
                 ",0,0,0,'" & gstrConvVrParaSql(Abs(txt_dblValorArbitrado)) & "," & _
                 gstrConvVrParaSql(txt_dblAliquota.Text) & ", " & _
                 IIf((bytDBType = EDatabases.Oracle), " :dblValor ", "@dblValor OUTPUT") & "'")
        
        
        Set gobjBanco = New clsBanco
'        If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
        If gobjBanco.ExecuteStoredProcedure(strSql, 10, , adoParameters) Then
'            If Not (adoResultado.BOF And adoResultado.EOF) Then
            If Not (adoParameters Is Nothing) Then
                'Mostra Valores para Usuário
'                strSql = "Confirma o cálculo de " & gstrConvVrDoSql(adoResultado!dblValorAparcelar) & _
'                        Chr(10) & " + " & gstrConvVrDoSql(adoResultado!dblValorNaoParcelado) & " em " & _
'                        (Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1) & " parcela(s) ?"
                strSql = "Confirma o cálculo de " & gstrConvVrDoSql(adoParameters("dblValorAparcelar")) & _
                        Chr(10) & " + " & gstrConvVrDoSql(adoParameters("dblValorNaoParcelado")) & " em " & _
                        (Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1) & " parcela(s) ?"
                Screen.MousePointer = vbNormal
                If MsgBox(strSql, vbYesNo, "Tributário") = vbYes Then
                    Screen.MousePointer = vbHourglass
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaBeginTrans
                    If Not gBlnVerificaLancamentos(txt_intExercício, dbc_intComposicaoDaReceita.BoundText, _
                                                    dbc_intComposicaoDaReceita.Text, (Val(txt_intParcelaFinal) - Val(txt_intParcelaInicial) + 1), _
                                                    gstrConvDtParaSql(txt_dtmDataPagamento), 0, strInscricao) Then
                        gobjBanco.ExecutaRollbackTrans
                        Screen.MousePointer = vbNormal
                        Exit Sub
                    End If
'                    strSql = strLancamentos(adoResultado!dblValorAparcelar, adoResultado!dblValorNaoParcelado)
                    strSql = strLancamentos(adoParameters("dblValorAparcelar"), adoParameters("dblValorNaoParcelado"))
                    'Executa o Lançamento
                    If gobjBanco.Execute(strSql, False) Then
                        gobjBanco.ExecutaCommitTrans
                        Screen.MousePointer = vbNormal
                        ExibeMensagem "Cálculo efetuado com sucesso!"
                    Else
                        gobjBanco.ExecutaRollbackTrans
                    End If
                End If
            End If
        End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Function strContribuintes() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql              As String
    Dim strInscricao As String
    strInscricao = Trim(Left(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, "-", vbTextCompare) - 1))
    strSql = " SELECT A.intContribuinte, A.strInscricaoCadastral " & _
            " FROM "
'            gstrEconomico & " AS A, "
    strSql = strSql & gstrEconomico & " A, "
'            gstrContribuinte & " AS CO "
    strSql = strSql & gstrContribuinte & " CO " & _
            " WHERE dtmDataBaixa IS NULL " & _
            " AND CO.PKId = A.intContribuinte " & _
            " AND A.strInscricaoCadastral = """ & strInscricao & """"
    strContribuintes = strSql
End Function

Private Function strLancamentos(dblValorAparcelar As String, dblValorNaoParcelado As String) As String

'******************************************************************************************
' Data: 08/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
'    strSql = " sp_CalculoLancamentoReceitas 1, " & gstrConvVrParaSql(dblValorAparcelar) & " , " & gstrConvVrParaSql(dblValorNaoParcelado) & ",'" & _
'            strPKId & "','" & strContribuintes & "'," & Val(txt_intExercício.Text) & _
'            "," & dbc_intComposicaoDaReceita.BoundText & ", NULL, " & _
'            gstrConvDtParaSql(txt_dtmDataPagamento.Text) & _
'            "," & gstrConvDtParaSql(txt_dtmDataVencimento) & "," & Val(txt_intParcelaInicial.Text) & "," & Val(txt_intParcelaFinal.Text) & "," & _
'            Val(txt_intIntervalo.Text) & ",3," & Val(dbc_intOcorrencia.BoundText) & ",2," & glngCodUsr & ",0,0," & gstrConvVrParaSql(txt_dblAliquota) & _
'             ",'?" & gstrConvVrParaSql(Abs(txt_dblValorArbitrado)) & "," & gstrConvVrParaSql(txt_dblAliquota.Text) & ", @dblValor OUTPUT'"
    strSql = gstrStoredProcedure("sp_CalculoLancamentoReceitas", "1, " & gstrConvVrParaSql(dblValorAparcelar) & " , " & gstrConvVrParaSql(dblValorNaoParcelado) & ",'" & _
            strPKId & "','" & strContribuintes & "'," & Val(txt_intExercício.Text) & _
            "," & dbc_intComposicaoDaReceita.BoundText & ", NULL, " & _
            gstrConvDtParaSql(txt_dtmDataPagamento.Text) & _
            "," & gstrConvDtParaSql(txt_dtmDataVencimento) & "," & Val(txt_intParcelaInicial.Text) & "," & Val(txt_intParcelaFinal.Text) & "," & _
            Val(txt_intIntervalo.Text) & ",3," & Val(dbc_intOcorrencia.BoundText) & ",2," & glngCodUsr & ",0,0," & gstrConvVrParaSql(txt_dblAliquota) & _
             ",'?" & gstrConvVrParaSql(Abs(txt_dblValorArbitrado)) & "," & gstrConvVrParaSql(txt_dblAliquota.Text) & _
             ", " & IIf((bytDBType = EDatabases.Oracle), ":dblValor", "@dblValor OUTPUT") & "'")
    strLancamentos = strSql
End Function
'Fim da Nova Versão
'====================

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3DPasta.Tab = 1 Then txt_dblValorArbitrado.SetFocus
End Sub

'==Para Efetuar Cálculo==
Private Function blnDadosOk() As Boolean
blnDadosOk = False
    Dim i As Integer
    
    If Val(Trim(txt_dblValorArbitrado.Text)) = 0 Then
        ExibeMensagem "O campo " & lbl_dblValorArbitrado.Caption & " não pode ser zero nem nulo."
        tab_3DPasta.Tab = 1
        txt_dblValorArbitrado.SetFocus
        Exit Function
    End If
    If dbc_intOcorrencia.MatchedWithList = False Then
        ExibeMensagem "O campo Ocorrência não pode ser Nulo"
        tab_3DPasta.Tab = 0
        dbc_intOcorrencia.SetFocus
        Exit Function
    End If
    If txt_dtmDataPagamento.Text = "" Then
        ExibeMensagem "O campo " & lbl_dtmDataPagamento.Caption & " não pode ser nulo."
        tab_3DPasta.Tab = 1
        txt_dtmDataPagamento.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_dtmDataPagamento.Text, True) Then
        tab_3DPasta.Tab = 1
        txt_dtmDataPagamento.SetFocus
        Exit Function
    End If
    If txt_dtmDataVencimento.Text = "" Then
        ExibeMensagem "O campo " & lbl_dtmDataVencimento.Caption & " não pode ser nulo."
        tab_3DPasta.Tab = 1
        txt_dtmDataVencimento.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_dtmDataVencimento.Text, True) Then
        tab_3DPasta.Tab = 1
        txt_dtmDataVencimento.SetFocus
        Exit Function
    ElseIf CVDate(txt_dtmDataPagamento.Text) > CVDate(txt_dtmDataVencimento.Text) Then
        ExibeMensagem "A " & lbl_dtmDataVencimento.Caption & " deve ser posterior a " & lbl_dtmDataPagamento.Caption & "."
        tab_3DPasta.Tab = 1
        txt_dtmDataVencimento.SetFocus
        Exit Function
    End If
    If txt_intIntervalo.Text = "" Then
        ExibeMensagem "O campo " & lbl_intIntervalo.Caption & " não pode ser nulo."
        tab_3DPasta.Tab = 1
        txt_intIntervalo.SetFocus
        Exit Function
    End If
    If txt_dblAliquota.Text = "" Then
        ExibeMensagem "O campo " & lbl_dblAliquota.Caption & " não pode ser nulo."
        tab_3DPasta.Tab = 1
        txt_dblAliquota.SetFocus
        Exit Function
    End If
    If txt_intExercício.Text = "" Then
        ExibeMensagem "O campo " & lbl_intExercício.Caption & " não pode ser nulo."
        tab_3DPasta.Tab = 1
        txt_intExercício.SetFocus
        Exit Function
    End If
    If (txt_intParcelaInicial.Text = "") Or (txt_intParcelaFinal.Text = "") Then
        ExibeMensagem "O intervalo de parcelas deve ser definido "
        tab_3DPasta.Tab = 1
        If txt_intParcelaInicial.Text = "" Then
            txt_intParcelaInicial.SetFocus
        Else
            txt_intParcelaFinal.SetFocus
        End If
        Exit Function
    End If
    If Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal) Then
        ExibeMensagem "O número da parcela final deve ser maior que o número da parcela inicial"
        tab_3DPasta.Tab = 1
        txt_intParcelaFinal.SetFocus
        Exit Function
    End If
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 Then
            blnDadosOk = True
            Exit Function
        End If
    Next
    ExibeMensagem "Selecione uma receita para efetuar o cálculo!"
    tab_3DPasta.Tab = 1
End Function

Private Function strPKId() As String
    Dim strSql As String
    Dim i As Integer
    strSql = ""
    For i = 0 To xarReceita.Count(1) - 1

        If xarReceita(i, 2) = -1 Then
            If strSql <> "" Then
               strSql = strSql & ","
            End If
            strSql = strSql & xarReceita(i, 0)
        End If
    Next
    strPKId = strSql
End Function

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
    
    If (txt_intParcelaInicial.Text = "") Or (txt_intParcelaFinal.Text = "") Then
        ExibeMensagem "O intervalo de parcelas deve ser definido "
        tab_3DPasta.Tab = 1
        If txt_intParcelaInicial.Text = "" Then
            txt_intParcelaInicial.SetFocus
        Else
            txt_intParcelaFinal.SetFocus
        End If
        Exit Function
    End If
    If Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal) Then
        ExibeMensagem "O número da parcela final deve ser maior que o número da parcela inicial"
        tab_3DPasta.Tab = 1
        txt_intParcelaFinal.SetFocus
        Exit Function
    End If
    If txt_intExercício.Text = "" Then
        ExibeMensagem "O exercício deve ser informado."
        tab_3DPasta.Tab = 1
        txt_intExercício.SetFocus
        Exit Function
    End If
    If dbc_intComposicaoDaReceita.BoundText = "" Then
        tab_3DPasta.Tab = 0
        ExibeMensagem "A Composição da Receita deve ser selecionada."
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If

    If dbc_strInscricaoCadastral.Text = "" Then
        ExibeMensagem "Selecione uma Inscrição Cadastral para gerar a Guia de Arrecadação."
        tab_3DPasta.Tab = 0
        dbc_strInscricaoCadastral.SetFocus
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
    dbc_intMensagem1.BoundText = ""
    dbc_intMensagem2.BoundText = ""
    txt_Mensagem1 = ""
    txt_Mensagem2 = ""
End Sub

Private Sub dbc_intMensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem1
End Sub

Private Sub dbc_intMensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem2
End Sub
