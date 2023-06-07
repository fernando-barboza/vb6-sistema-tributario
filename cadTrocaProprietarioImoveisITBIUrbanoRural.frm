VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadTrocaProprietarioImoveisITBIUrbanoRural 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Troca de Proprietário de Imóveis"
   ClientHeight    =   6255
   ClientLeft      =   1200
   ClientTop       =   2445
   ClientWidth     =   8640
   Icon            =   "cadTrocaProprietarioImoveisITBIUrbanoRural.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6015
      Left            =   135
      TabIndex        =   4
      Top             =   105
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Proprietário do Imóvel"
      TabPicture(0)   =   "cadTrocaProprietarioImoveisITBIUrbanoRural.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strInscricaoCadastral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_strProprietario"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "msk_strInscricaoCadastral(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbc_strProprietario"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_Proprietarios"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_Proprietario"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_TipoITBI"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_PKId"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Troca de Proprietário"
      TabPicture(1)   =   "cadTrocaProprietarioImoveisITBIUrbanoRural.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_Frame"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra_Frame 
         Height          =   3045
         Left            =   -74850
         TabIndex        =   61
         Top             =   390
         Width           =   8025
         Begin VB.CommandButton cmd_Contribuinte 
            Height          =   300
            Left            =   7470
            Picture         =   "cadTrocaProprietarioImoveisITBIUrbanoRural.frx":107A
            Style           =   1  'Graphical
            TabIndex        =   62
            TabStop         =   0   'False
            Tag             =   "15"
            ToolTipText     =   "Ativa Cadastro de Contribuinte"
            Top             =   270
            Width           =   360
         End
         Begin VB.TextBox txt_strDesmembramento 
            Height          =   285
            Left            =   2130
            MaxLength       =   20
            TabIndex        =   7
            Top             =   1020
            Width           =   2070
         End
         Begin VB.TextBox txt_strLoteamento 
            Height          =   285
            Left            =   5790
            MaxLength       =   20
            TabIndex        =   8
            Top             =   1035
            Width           =   2070
         End
         Begin VB.TextBox txt_strEscritura 
            Height          =   285
            Left            =   5790
            MaxLength       =   15
            TabIndex        =   12
            Top             =   1755
            Width           =   2070
         End
         Begin VB.TextBox txt_strHabitese 
            Height          =   285
            Left            =   2130
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1755
            Width           =   2070
         End
         Begin VB.TextBox txt_strLote 
            Height          =   285
            Left            =   5790
            MaxLength       =   20
            TabIndex        =   10
            Top             =   1395
            Width           =   2070
         End
         Begin VB.TextBox txt_strQuadra 
            Height          =   285
            Left            =   2130
            MaxLength       =   20
            TabIndex        =   9
            Top             =   1395
            Width           =   2070
         End
         Begin VB.TextBox txt_strDescricao 
            Height          =   750
            Left            =   2130
            MaxLength       =   16
            TabIndex        =   13
            Top             =   2130
            Width           =   5730
         End
         Begin MSDataListLib.DataCombo dbc_strNovoProprietario 
            Height          =   315
            Left            =   2130
            TabIndex        =   5
            Top             =   270
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Style           =   2
            Text            =   ""
         End
         Begin MSMask.MaskEdBox msk_strInscricaoCadastral 
            Height          =   285
            Index           =   1
            Left            =   2130
            TabIndex        =   6
            Top             =   660
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin VB.Label lbl_strNovoProprietario 
            AutoSize        =   -1  'True
            Caption         =   "Nome do novo Proprietário"
            Height          =   195
            Left            =   150
            TabIndex        =   71
            Top             =   360
            Width           =   1890
         End
         Begin VB.Label lbl_strDesmembramento 
            AutoSize        =   -1  'True
            Caption         =   "Desmembramento"
            Height          =   195
            Left            =   765
            TabIndex        =   70
            Top             =   1080
            Width           =   1275
         End
         Begin VB.Label lbl_strLoteamento 
            AutoSize        =   -1  'True
            Caption         =   "Loteamento"
            Height          =   195
            Left            =   4830
            TabIndex        =   69
            Top             =   1095
            Width           =   840
         End
         Begin VB.Label lbl_strLote 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            Height          =   195
            Left            =   5355
            TabIndex        =   68
            Top             =   1455
            Width           =   315
         End
         Begin VB.Label lbl_strEscritura 
            AutoSize        =   -1  'True
            Caption         =   "Matrícula"
            Height          =   195
            Left            =   4995
            TabIndex        =   67
            Top             =   1800
            Width           =   675
         End
         Begin VB.Label lbl_strHabitese 
            AutoSize        =   -1  'True
            Caption         =   "Habite-se"
            Height          =   195
            Left            =   1365
            TabIndex        =   66
            Top             =   1815
            Width           =   675
         End
         Begin VB.Label lbl_strQuadra 
            AutoSize        =   -1  'True
            Caption         =   "Quadra"
            Height          =   195
            Left            =   1515
            TabIndex        =   65
            Top             =   1455
            Width           =   525
         End
         Begin VB.Label lbl_strInscricaoAnterior 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   690
            TabIndex        =   64
            Top             =   735
            Width           =   1350
         End
         Begin VB.Label lbl_strDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Histórico"
            Height          =   195
            Left            =   1440
            TabIndex        =   63
            Top             =   2355
            Width           =   615
         End
      End
      Begin VB.TextBox txt_Conteudo 
         Height          =   285
         Left            =   -73140
         MaxLength       =   50
         TabIndex        =   41
         Top             =   720
         Width           =   3945
      End
      Begin VB.TextBox txt_DescricaoConteudo 
         Height          =   285
         Left            =   -73140
         MaxLength       =   50
         TabIndex        =   40
         Top             =   1020
         Width           =   3945
      End
      Begin VB.CommandButton cmd_Up 
         Height          =   285
         Left            =   -67470
         Picture         =   "cadTrocaProprietarioImoveisITBIUrbanoRural.frx":1198
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Acima"
         Top             =   1425
         Width           =   300
      End
      Begin VB.CommandButton cmd_Down 
         Height          =   285
         Left            =   -67470
         Picture         =   "cadTrocaProprietarioImoveisITBIUrbanoRural.frx":12E2
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Abaixo"
         Top             =   1725
         Width           =   300
      End
      Begin VB.TextBox txt_PKId 
         Height          =   285
         Left            =   7680
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame fra_TipoITBI 
         Caption         =   "Tipo de ITBI"
         Height          =   615
         Left            =   5850
         TabIndex        =   36
         Top             =   510
         Width           =   2355
         Begin VB.OptionButton opt_Urbano 
            Caption         =   "Urbano"
            Height          =   315
            Left            =   360
            TabIndex        =   1
            Top             =   210
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton opt_Rural 
            Caption         =   "Rural"
            Height          =   255
            Left            =   1350
            TabIndex        =   2
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame fra_Proprietario 
         Height          =   2430
         Left            =   120
         TabIndex        =   14
         Top             =   1590
         Width           =   8070
         Begin VB.TextBox txt_strNome 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   210
            Width           =   6765
         End
         Begin VB.Frame fra_Endereco 
            Caption         =   "Endereço"
            Height          =   1395
            Left            =   120
            TabIndex        =   16
            Top             =   900
            Width           =   7845
            Begin VB.TextBox txt_UF 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5070
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   990
               Width           =   510
            End
            Begin VB.TextBox txt_Municipio 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5070
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   630
               Width           =   2625
            End
            Begin VB.TextBox txt_Cep 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   6615
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   990
               Width           =   1080
            End
            Begin VB.TextBox txt_Complemento 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   6810
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   270
               Width           =   870
            End
            Begin VB.TextBox txt_Numero 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   270
               Width           =   795
            End
            Begin VB.TextBox txt_Logradouro 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   270
               Width           =   4005
            End
            Begin VB.TextBox txt_Bairro 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   630
               Width           =   3105
            End
            Begin VB.TextBox txt_Distrito 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   990
               Width           =   3525
            End
            Begin VB.Label lbl_strDistritoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               Height          =   195
               Left            =   450
               TabIndex        =   32
               Top             =   1050
               Width           =   480
            End
            Begin VB.Label lbl_intMunicipioC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   4290
               TabIndex        =   31
               Top             =   690
               Width           =   705
            End
            Begin VB.Label lbl_intBairroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   525
               TabIndex        =   30
               Top             =   690
               Width           =   405
            End
            Begin VB.Label lbl_intLogradouroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   330
               Width           =   810
            End
            Begin VB.Label lbl_intNumeroC 
               AutoSize        =   -1  'True
               Caption         =   "Nº"
               Height          =   195
               Left            =   5100
               TabIndex        =   28
               Top             =   330
               Width           =   180
            End
            Begin VB.Label lbl_strComplementoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6270
               TabIndex        =   27
               Top             =   330
               Width           =   480
            End
            Begin VB.Label lbl_intUFC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   4770
               TabIndex        =   26
               Top             =   1065
               Width           =   210
            End
            Begin VB.Label lbl_intCepC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   6240
               TabIndex        =   25
               Top             =   1050
               Width           =   285
            End
         End
         Begin VB.TextBox txt_strCNPJCPFP 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   570
            Width           =   1845
         End
         Begin VB.Label lbl_strNome 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   135
            TabIndex        =   35
            Top             =   255
            Width           =   795
         End
         Begin VB.Label lbl_strCNPJCPFP 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   150
            TabIndex        =   34
            Top             =   615
            Width           =   780
         End
      End
      Begin MSComctlLib.ListView lvw_TipoComunicacao 
         Height          =   3210
         Left            =   -74760
         TabIndex        =   42
         Top             =   1410
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   5662
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
      Begin Threed.SSPanel ssp_TipoComunicacao 
         Height          =   390
         Left            =   -69060
         TabIndex        =   43
         Top             =   915
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComctlLib.Toolbar tlb_TipoComunicacao 
            Height          =   330
            Left            =   30
            TabIndex        =   44
            Top             =   30
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Novo"
                  Object.ToolTipText     =   "Novo"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Salvar"
                  Object.ToolTipText     =   "Adicionar / Alterar"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Deletar"
                  Object.ToolTipText     =   "Remover"
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ImageList img_Aux 
         Left            =   -67470
         Top             =   2070
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadTrocaProprietarioImoveisITBIUrbanoRural.frx":142C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadTrocaProprietarioImoveisITBIUrbanoRural.frx":158C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadTrocaProprietarioImoveisITBIUrbanoRural.frx":16E8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Proprietarios 
         Height          =   1755
         Left            =   150
         TabIndex        =   45
         Top             =   4110
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3096
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
         Columns(1).DataField=   "strInscricaoAnterior"
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
         Splits(0)._ColumnProps(12)=   "Column(2).Width=10583"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=10504"
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
         Height          =   3615
         Left            =   -74850
         TabIndex        =   46
         Top             =   2250
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   6376
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
         Columns(1).Caption=   "Código"
         Columns(1).DataField=   "intCodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   68
         Columns(2)._MaxComboItems=   20
         Columns(2).ValueItems(0)._DefaultItem=   0
         Columns(2).ValueItems(0).Value=   ""
         Columns(2).ValueItems(0).Value.vt=   8
         Columns(2).ValueItems(0).DisplayValue=   ""
         Columns(2).ValueItems(0).DisplayValue.vt=   8
         Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(2).ValueItems.Count=   1
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descrição"
         Columns(3).DataField=   "strDescricao"
         Columns(3).DropDown=   "tdd_Atividades"
         Columns(3).DropDown.vt=   8
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   1
         Splits(0).MarqueeStyle=   5
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1111"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1032"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=450"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=370"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=10848"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=10769"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(3).AutoDropDown=1"
         Splits(0)._ColumnProps(25)=   "Column(3).DropDownList=1"
         Splits(0)._ColumnProps(26)=   "Column(3).AutoCompletion=1"
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
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbc_strProprietario 
         Height          =   315
         Left            =   1890
         TabIndex        =   3
         Top             =   1230
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSMask.MaskEdBox msk_strInscricaoCadastral 
         Height          =   285
         Index           =   0
         Left            =   1890
         TabIndex        =   0
         Top             =   870
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin VB.Label lbl_strProprietario 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Proprietário"
         Height          =   195
         Left            =   285
         TabIndex        =   60
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label lbl_DescricaoConteudo 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   -74175
         TabIndex        =   59
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label lbl_TipoComunicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   -73770
         TabIndex        =   58
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lbl_strInscricaoCadastral 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   420
         TabIndex        =   57
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label lbl_dias 
         AutoSize        =   -1  'True
         Caption         =   "dias"
         Height          =   195
         Left            =   -67860
         TabIndex        =   56
         Top             =   1410
         Width           =   285
      End
      Begin VB.Label lbl_dblAliquota 
         AutoSize        =   -1  'True
         Caption         =   "Alíquota sobre Valor"
         Height          =   195
         Left            =   -74355
         TabIndex        =   55
         Top             =   1770
         Width           =   1440
      End
      Begin VB.Label lbl_dblbvalorITBI 
         AutoSize        =   -1  'True
         Caption         =   "Valor de Avaliação do Imóvel"
         Height          =   195
         Left            =   -70920
         TabIndex        =   54
         Top             =   675
         Width           =   2070
      End
      Begin VB.Label lbl_dblDesconto 
         AutoSize        =   -1  'True
         Caption         =   "Desconto"
         Height          =   195
         Left            =   -69540
         TabIndex        =   53
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lbl_intIntervalo 
         AutoSize        =   -1  'True
         Caption         =   "Intervalo entre Parcelas"
         Height          =   195
         Left            =   -70530
         TabIndex        =   52
         Top             =   1380
         Width           =   1680
      End
      Begin VB.Label lbl_intNumeroParcelas 
         AutoSize        =   -1  'True
         Caption         =   "Número de Parcelas"
         Height          =   195
         Left            =   -74355
         TabIndex        =   51
         Top             =   1410
         Width           =   1440
      End
      Begin VB.Label lbl_dtmDataVencimento 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vencimento"
         Height          =   195
         Left            =   -70305
         TabIndex        =   50
         Top             =   1035
         Width           =   1455
      End
      Begin VB.Label lbl_dtmDataPagamento 
         AutoSize        =   -1  'True
         Caption         =   "Data de Lançamento"
         Height          =   195
         Left            =   -74415
         TabIndex        =   49
         Top             =   1020
         Width           =   1500
      End
      Begin VB.Label lbl_p1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   -71460
         TabIndex        =   48
         Top             =   1815
         Width           =   120
      End
      Begin VB.Label lbl_p2 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   -67410
         TabIndex        =   47
         Top             =   1785
         Width           =   120
      End
   End
End
Attribute VB_Name = "frmCadTrocaProprietarioImoveisITBIUrbanoRural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando           As Boolean
Dim mobjAux                 As Object
Dim mblnSelecionou          As Boolean
Dim mblnPrimeiraVez         As Boolean
Dim ConsultaProprietario    As Integer
Dim xarReceita              As XArrayDB
Dim adoResultado            As ADODB.Recordset

Private Sub cmd_Contribuinte_Click()
    CarregaForm frmCadContribuinte, dbc_strNovoProprietario
End Sub

Function blnExisteDebitoEmAberto() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo DATEDIFF() do SQL Server pela função
'            gstrDATEDIFF.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim adoRec As ADODB.Recordset

    strSql = ""
    strSql = strSql & " SELECT PAR.intLancamentoCalculo FROM "
    strSql = strSql & gstrParcelaReceita & " PAR, "
    strSql = strSql & gstrLancamentoCalculo & " LAN "
    strSql = strSql & " WHERE LAN.intContribuinte = " & dbc_strNovoProprietario.BoundText
    strSql = strSql & " AND PAR.intLancamentoCalculo = LAN.PKId "
'    strSql = strSql & " AND DATEDIFF(DAY, PAR.dtmDataVencimento, GETDATE()) > 0 "
    strSql = strSql & " AND " & gstrDATEDIFF("PAR.dtmDataVencimento", strGETDATE) & " > 0 "
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        With adoRec
            If Not .EOF Then
                blnExisteDebitoEmAberto = True
            End If
        End With
    End If
    
End Function

Private Sub dbc_strNovoProprietario_Click(Area As Integer)
    DropDownDataCombo dbc_strNovoProprietario, Me, Area
End Sub

Private Sub dbc_strNovoProprietario_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strNovoProprietario, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strProprietario_Click(Area As Integer)
    DropDownDataCombo dbc_strProprietario, Me, Area
    ConsultaProprietario = 2
    If Area = 2 And dbc_strProprietario.MatchedWithList Then
        Set dbc_strProprietario.DataSource = Nothing
        txt_PKId.Text = dbc_strProprietario.BoundText
        BuscarDadosProprietario
    End If
End Sub

Private Sub dbc_strProprietario_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strProprietario, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 673
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    If mobjAux Is Nothing Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
    msk_strInscricaoCadastral(0).SetFocus
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    TrocaCorObjeto msk_strInscricaoCadastral(1), True
    LeDaTabelaParaObj gstrImobiliario, tdb_Proprietarios, strQueryImobiliario
    LeDaTabelaParaObj gstrImobiliario, dbc_strProprietario, strQueryProprietario
    VerificaMascaraInscricao
    LeDaTabelaParaObj gstrContribuinte, dbc_strNovoProprietario, strQueryNovoProprietario
    VerificaObjParaAplicar mobjAux
    ConsultaProprietario = 1
    
'''GUIA
    'LeDaTabelaParaObj gstrEconomico, dbc_strInscricaoInicial, strQueryInscricao
    'LeDaTabelaParaObj gstrEconomico, dbc_strInscricaoFinal, strQueryInscricao
    'LeDaTabelaParaObj gstrMensagem, dbc_intMensagem1, strQueryMensagem
    'LeDaTabelaParaObj gstrMensagem, dbc_intMensagem2, strQueryMensagem
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub



Private Sub msk_strInscricaoCadastral_LostFocus(Index As Integer)
    If Index = 0 Then
        ConsultaProprietario = 1
        BuscarDadosProprietario
    End If
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



Private Sub tdb_Proprietarios_Click()
    mblnPrimeiraVez = True
End Sub

Private Sub tdb_Proprietarios_FilterChange()
    gblnFilraCampos tdb_Proprietarios
End Sub

Private Sub tdb_Proprietarios_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ConsultaProprietario = 2
    With tdb_Proprietarios
        If Not .EOF And Not .BOF Then
            txt_PKId.Text = .Columns("PKID").Value

            If mblnPrimeiraVez Then

                BuscarDadosProprietario

                gCorLinhaSelecionada tdb_Proprietarios

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

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    

Dim strSql As String
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        dbc_strProprietario.Text = ""
        dbc_strNovoProprietario.Text = ""
        txt_strDesmembramento.Text = ""
        txt_strQuadra.Text = ""
        txt_strHabitese.Text = ""
        txt_strDescricao.Text = ""
        txt_strLoteamento.Text = ""
        txt_strLote.Text = ""
        txt_strEscritura.Text = ""
        msk_strInscricaoCadastral(0).Text = ""
        msk_strInscricaoCadastral(1).Text = ""
        VerificaMascaraInscricao
        LimpaEndereco
        mblnSelecionou = False
    End If
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    If strModoOperacao = gstrSalvar Then
        If gblnExclusaoGravacaoOk("A", " de proprietário ") Then
            If blnDadosDigitadosOK Then
                If blnExisteDebitoEmAberto Then
                    
                    If MsgBox(" O contribuinte " & dbc_strNovoProprietario.Text & " tem débito em aberto " & Chr(13) & " Deseja continuar ? ", vbOKCancel + vbInformation) = vbCancel Then
                        Exit Sub
                    End If
                    
                End If
            
                strSql = ""
                strSql = strSql & " UPDATE " & IIf(opt_Urbano.Value, gstrImobiliario, gstrImobiliarioRural)
                strSql = strSql & "  SET intContribuinte = " & dbc_strNovoProprietario.BoundText
                If opt_Urbano.Value Then
                    strSql = strSql & ", strDesmembramento = '" & txt_strDesmembramento.Text & "'"
                    strSql = strSql & ", strLoteamento = '" & txt_strLoteamento.Text & "'"
                    strSql = strSql & ", strQuadra = '" & txt_strQuadra.Text & "'"
                    strSql = strSql & ", strLote = '" & txt_strLote.Text & "'"
                    strSql = strSql & ", strHabitese = '" & txt_strHabitese.Text & "'"
                    strSql = strSql & ", strMatricula = '" & txt_strEscritura.Text & "'"
                End If
'                strSql = strSql & ", dtmDtAtualizacao = getdate(), lngCodUsr = " & glngCodUsr
                strSql = strSql & ", dtmDtAtualizacao = " & strGETDATE & ", lngCodUsr = " & glngCodUsr
                
                strSql = strSql & " WHERE PKId = " & txt_PKId.Text
                
                strSql = strSql & " INSERT INTO " & IIf(opt_Urbano.Value, gstrHistoricoImobiliario, gstrHistoricoImobiliarioRural)
                strSql = strSql & " (intImobiliario, strDescricao) "
                strSql = strSql & " VALUES (" & txt_PKId.Text & ", '" & txt_strDescricao & "')"
            
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaBeginTrans
                If gobjBanco.Execute(strSql) Then
                    gobjBanco.ExecutaCommitTrans
                    ExibeMensagem " Transferência efetuada com sucesso "
                    LeDaTabelaParaObj gstrImobiliario, tdb_Proprietarios, strQueryImobiliario
                    LeDaTabelaParaObj gstrImobiliario, dbc_strProprietario, strQueryProprietario
                    MantemForm (gstrNovo)
                Else
                    gobjBanco.ExecutaRollbackTrans
                End If
                
            End If
        End If
        
    End If
        

    
End Sub

Private Function blnDadosDigitadosOK() As Boolean
    
    If msk_strInscricaoCadastral(1).Text = "" Then
        ExibeMensagem "Selecione um imóvel. "
        tab_3DPasta.Tab = 0
        msk_strInscricaoCadastral(0).SetFocus
        Exit Function
    End If
    
    If Not dbc_strNovoProprietario.MatchedWithList Then
        ExibeMensagem "O campo " & lbl_strNovoProprietario.Caption & " não pode ser nulo."
        tab_3DPasta.Tab = 1
        dbc_strNovoProprietario.SetFocus
        Exit Function
    ElseIf dbc_strProprietario.Text = dbc_strNovoProprietario.Text Then
        ExibeMensagem "O campo " & lbl_strNovoProprietario.Caption & " não pode ser igual ao do atual proprietário."
        tab_3DPasta.Tab = 1
        dbc_strNovoProprietario.SetFocus
        Exit Function
    End If
    
    blnDadosDigitadosOK = True
End Function

Private Sub opt_Rural_Click()
    mblnPrimeiraVez = False
    LimpaEndereco
    msk_strInscricaoCadastral(0).Text = ""
    msk_strInscricaoCadastral(1).Text = ""
    dbc_strProprietario.Text = ""
    LeDaTabelaParaObj gstrImobiliario, dbc_strProprietario, strQueryProprietario
    tdb_Proprietarios.DataSource = Nothing
    LeDaTabelaParaObj gstrImobiliario, tdb_Proprietarios, strQueryImobiliario
    msk_strInscricaoCadastral(0).SetFocus
    mblnPrimeiraVez = False
    ReposiconaHistorico (2)
End Sub

Private Sub ReposiconaHistorico(intPosicao As Integer)
    Dim blnUrbano  As Boolean
    
    If intPosicao = 1 Then
        txt_strDescricao.Top = 2565 - 400
        lbl_strDescricao.Top = 2775 - 400
        blnUrbano = True
    Else
        txt_strDescricao.Top = msk_strInscricaoCadastral(1).Top
        lbl_strDescricao.Top = 1305 - 500
        blnUrbano = False
    End If
    lbl_strInscricaoAnterior.Visible = blnUrbano
    msk_strInscricaoCadastral(1).Visible = blnUrbano
    lbl_strDesmembramento.Visible = blnUrbano
    txt_strDesmembramento.Visible = blnUrbano
    lbl_strLoteamento.Visible = blnUrbano
    txt_strLoteamento.Visible = blnUrbano
    lbl_strQuadra.Visible = blnUrbano
    txt_strQuadra.Visible = blnUrbano
    lbl_strLote.Visible = blnUrbano
    txt_strLote.Visible = blnUrbano
    lbl_strHabitese.Visible = blnUrbano
    txt_strHabitese.Visible = blnUrbano
    lbl_strEscritura.Visible = blnUrbano
    txt_strEscritura.Visible = blnUrbano
End Sub

Private Sub opt_Urbano_Click()
    mblnPrimeiraVez = False
    LimpaEndereco
    msk_strInscricaoCadastral(0).Text = ""
    msk_strInscricaoCadastral(1).Text = ""
    dbc_strProprietario.Text = ""
    LeDaTabelaParaObj gstrImobiliario, dbc_strProprietario, strQueryProprietario
    tdb_Proprietarios.DataSource = Nothing
    LeDaTabelaParaObj gstrImobiliario, tdb_Proprietarios, strQueryImobiliario
    msk_strInscricaoCadastral(0).SetFocus
    mblnPrimeiraVez = False
    ReposiconaHistorico (1)
End Sub

Private Function strQueryImobiliario() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String

    strSql = ""
    strSql = strSql & " SELECT A.PKId, A.strInscricaoAnterior, CO.strNome "
    strSql = strSql & " FROM "
    If opt_Rural.Value = True Then
'        strSQL = strSQL & gstrImobiliarioRural & " AS A, "
        strSql = strSql & gstrImobiliarioRural & " A, "
    Else
'        strSQL = strSQL & gstrImobiliario & " AS A, "
        strSql = strSql & gstrImobiliario & " A, "
    End If
'    strSQL = strSQL & gstrContribuinte & " AS CO "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = A.intContribuinte "
    strSql = strSql & " ORDER BY strInscricaoAnterior "
    strSql = strSql
    
    strQueryImobiliario = strSql
End Function

Private Function strQueryProprietario() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String

    strSql = ""
    strSql = strSql & " SELECT A.PKId, CO.strNome "
    strSql = strSql & " FROM "
    If opt_Rural.Value = True Then
'        strSQL = strSQL & gstrImobiliarioRural & " AS A, "
        strSql = strSql & gstrImobiliarioRural & " A, "
    Else
'        strSQL = strSQL & gstrImobiliario & " AS A, "
        strSql = strSql & gstrImobiliario & " A, "
    End If
'    strSQL = strSQL & gstrContribuinte & " AS CO "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = A.intContribuinte "
    strSql = strSql & " ORDER BY strNome "
    strSql = strSql
    
    strQueryProprietario = strSql
End Function

Private Function strQueryNovoProprietario() As String
    Dim strSql As String

    strSql = ""
    strSql = strSql & " SELECT PKId, strNome "
    strSql = strSql & " FROM " & gstrContribuinte
    strSql = strSql & " ORDER BY strNome "
    strSql = strSql
    
    strQueryNovoProprietario = strSql
End Function

Private Function BuscarDadosProprietario()

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT A.PKId, " & IIf(opt_Urbano.Value, "A.strDesmembramento, A.strLoteamento, A.strQuadra, A.strLote, A.strHabitese, A.strMatricula,", "")
    strSql = strSql & " CO.strNome, LG.strDescricao AS Logradouro, A.intNumero, "
    strSql = strSql & " A.strComplemento , A.intCep, BA.strDescricao AS Bairro, UF.strSigla AS UF, A.strCNPJCPF, "
    strSql = strSql & " CO.strDistritoC, CI.strDescricao AS Municipio, A.dblValorITBI, "
    strSql = strSql & " A.strInscricaoAnterior "
    strSql = strSql & " FROM "
    If opt_Rural.Value = True Then
'        strSQL = strSQL & gstrImobiliarioRural & " AS A, "
        strSql = strSql & gstrImobiliarioRural & " A, "
    Else
'        strSQL = strSQL & gstrImobiliario & " AS A, "
        strSql = strSql & gstrImobiliario & " A, "
    End If
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
    strSql = strSql & " CO.PKId = A.intContribuinte "
    strSql = strSql & " AND CI.PKId = CO.intMunicipioC "
    strSql = strSql & " AND LG.PKId = A.intLogradouro "
    strSql = strSql & " AND BA.PKId = A.intBairro "
    strSql = strSql & " AND UF.PKId = A.intUF "
    If ConsultaProprietario <> 1 Then
        strSql = strSql & " AND A.PKID = " & txt_PKId.Text
    Else
        strSql = strSql & " AND A.strInscricaoAnterior = '" & msk_strInscricaoCadastral(0).Text & "'"
    End If
    
    LimpaEndereco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                dbc_strProprietario.BoundText = !Pkid
                msk_strInscricaoCadastral(0).Text = !strInscricaoAnterior
                msk_strInscricaoCadastral(1).Text = msk_strInscricaoCadastral(0).Text
                txt_strCNPJCPFP = gstrCGCCPFFormatado(!StrCnpjCpf)
                txt_strNome.Text = !STRNOME
                txt_Logradouro.Text = !Logradouro
                txt_Numero.Text = !INTNUMERO
                txt_Complemento.Text = gstrENulo(!STRCOMPLEMENTO)
                txt_Bairro.Text = gstrENulo(!Bairro)
                txt_Municipio.Text = gstrENulo(!Municipio)
                txt_Distrito.Text = gstrENulo(!strDistritoC)
                txt_UF.Text = gstrENulo(!UF)
                txt_Cep.Text = gstrCEPFormatado(!INTCEP)
                If opt_Urbano.Value Then
                    txt_strDesmembramento.Text = gstrENulo(!strDesmembramento)
                    txt_strLoteamento.Text = gstrENulo(!strLoteamento)
                    txt_strQuadra.Text = gstrENulo(!strQuadra)
                    txt_strLote.Text = gstrENulo(!strLote)
                    txt_strHabitese.Text = gstrENulo(!strHabitese)
                    txt_strEscritura.Text = gstrENulo(!strMatricula)
                End If
            End With
        Else
            dbc_strProprietario.Text = ""
'            msk_strInscricaoCadastral.SetFocus
        End If
    End If
    
End Function

Sub VerificaMascaraInscricao()
Dim strSql As String
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
    msk_strInscricaoCadastral(0).Mask = strMascara
    msk_strInscricaoCadastral(1).Mask = strMascara
End Sub

Private Function LimpaEndereco()
    txt_strCNPJCPFP = ""
    txt_strNome.Text = ""
    txt_Logradouro.Text = ""
    txt_Numero.Text = ""
    txt_Complemento.Text = ""
    txt_Bairro.Text = ""
    txt_Municipio.Text = ""
    txt_Distrito.Text = ""
    txt_UF.Text = ""
    txt_Cep.Text = ""
End Function

'''' ######################   caractervalido e marcacampo   ###################### ''''

'Private Sub dbc_strInscricaoCadastral_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "", dbc_strInscricaoCadastral
'End Sub

Private Sub dbc_strProprietario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "", dbc_strProprietario
End Sub

Private Sub msk_strInscricaoCadastral_GotFocus(Index As Integer)
    If Index = 0 Then
        MarcaCampo msk_strInscricaoCadastral(0)
    End If
End Sub

Private Sub msk_strInscricaoCadastral_KeyPress(Index As Integer, KeyAscii As Integer)
    CaracterValido KeyAscii, "A", msk_strInscricaoCadastral
End Sub

Private Sub txt_strDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strDescricao
End Sub

Private Sub txt_strEscritura_GotFocus()
    MarcaCampo txt_strEscritura
End Sub

Private Sub txt_strEscritura_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strEscritura
End Sub

Private Sub txt_strHabitese_GotFocus()
    MarcaCampo txt_strHabitese
End Sub

Private Sub txt_strHabitese_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strHabitese
End Sub

Private Sub txt_strLote_GotFocus()
    MarcaCampo txt_strLote
End Sub

Private Sub txt_strLote_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strLote
End Sub

Private Sub txt_strLoteamento_GotFocus()
    MarcaCampo txt_strLoteamento
End Sub

Private Sub txt_strLoteamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strLoteamento
End Sub

Private Sub txt_strDesmembramento_GotFocus()
    MarcaCampo txt_strDesmembramento
End Sub

Private Sub txt_strDesmembramento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strDesmembramento
End Sub

Private Sub txt_strQuadra_GotFocus()
    MarcaCampo txt_strQuadra
End Sub

Private Sub txt_strQuadra_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strQuadra
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

'Private Sub chk_EmBranco1_Click()
'    If chk_EmBranco1.Value = 1 Then
'        dbc_intMensagem1.BoundText = ""
'        dbc_intMensagem1.Enabled = False
'        TrocaCorObjeto dbc_intMensagem1, True
'        txt_Mensagem1.Text = ""
'        txt_Mensagem1.Enabled = False
'        TrocaCorObjeto txt_Mensagem1, True
'    Else
'        dbc_intMensagem1.Enabled = True
'        TrocaCorObjeto dbc_intMensagem1, False
'        txt_Mensagem1.Enabled = True
'        TrocaCorObjeto txt_Mensagem1, False
'    End If
'End Sub
'
'Private Sub chk_EmBranco2_Click()
'    If chk_EmBranco2.Value = 1 Then
'        dbc_intMensagem2.BoundText = ""
'        dbc_intMensagem2.Enabled = False
'        TrocaCorObjeto dbc_intMensagem2, True
'        txt_Mensagem2.Text = ""
'        txt_Mensagem2.Enabled = False
'        TrocaCorObjeto txt_Mensagem2, True
'    Else
'        dbc_intMensagem2.Enabled = True
'        TrocaCorObjeto dbc_intMensagem2, False
'        txt_Mensagem2.Enabled = True
'        TrocaCorObjeto txt_Mensagem2, False
'    End If
'End Sub
'
'Private Sub dbc_intMensagem1_Click(Area As Integer)
'    If Area = 2 Then
'        LeDoComboParaTXT1
'    End If
'End Sub
'
'Private Sub dbc_intMensagem2_Click(Area As Integer)
'    If Area = 2 Then
'        LeDoComboParaTXT2
'    End If
'End Sub
'
'Private Function LeDoComboParaTXT1()
'Dim strSql As String
'    strSql = ""
'    strSql = strSql & " SELECT strMensagem "
'    strSql = strSql & " FROM " & gstrMensagem
'    strSql = strSql & " WHERE PKId = " & Val(dbc_intMensagem1.BoundText)
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        If adoResultado.EOF = False Then
'            txt_Mensagem1.Text = adoResultado!strMensagem
'            adoResultado.MoveNext
'        Else
'            txt_Mensagem1.Text = ""
'        End If
'    End If
'End Function
'
'Private Function LeDoComboParaTXT2()
'Dim strSql As String
'    strSql = ""
'    strSql = strSql & " SELECT strMensagem "
'    strSql = strSql & " FROM " & gstrMensagem
'    strSql = strSql & " WHERE PKId = " & Val(dbc_intMensagem2.BoundText)
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        If adoResultado.EOF = False Then
'            txt_Mensagem2.Text = adoResultado!strMensagem
'            adoResultado.MoveNext
'        Else
'            txt_Mensagem2.Text = ""
'        End If
'    End If
'End Function
'
'Private Function strQueryMensagem() As String
'Dim strSql As String
'
'    strSql = ""
'    strSql = strSql & "SELECT PKId, ltrim(rtrim(PKId)) + ' - ' + ltrim(rtrim(strDescricao)) as Descricao "
'    strSql = strSql & " FROM " & gstrMensagem
'    strSql = strSql & " ORDER BY PKId "
'
'strQueryMensagem = strSql
'End Function
'
'Private Function blnDadosGuiaOK() As Boolean
'blnDadosGuiaOK = False
'
'    If dbc_strInscricaoInicial.BoundText = "" Then
'        ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para gerar a Guia de Arrecadação."
'        dbc_strInscricaoInicial.SetFocus
'        Exit Function
'    End If
'
'    If dbc_strInscricaoFinal.BoundText = "" Then
'        ExibeMensagem "Selecione uma Inscrição Cadastral Final para gerar a Guia de Arrecadação."
'        dbc_strInscricaoFinal.SetFocus
'        Exit Function
'    End If
'
'    If txt_intExercicio.Text = "" Then
'        ExibeMensagem "O Exercício deve ser Digitado."
'        txt_intExercicio.SetFocus
'        Exit Function
'    End If
'
'    If txt_DataDeVencimento.Text = "" Then
'        ExibeMensagem "A data de vencimento deve ser digitada."
'        txt_DataDeVencimento.SetFocus
'        Exit Function
'    ElseIf gblnDataValida(txt_DataDeVencimento.Text) = False Then
'        ExibeMensagem "Data de vencimento inválida."
'        txt_DataDeVencimento.SetFocus
'        Exit Function
'    End If
'
'    If chk_EmBranco1.Value = 0 Then
'        If txt_Mensagem1.Text = "" Then
'            ExibeMensagem "A mensagem 1 tem que ser selecionada."
'            Exit Function
'        End If
'    End If
'
'    If chk_EmBranco2.Value = 0 Then
'        If txt_Mensagem2.Text = "" Then
'            ExibeMensagem "A mensagem 2 tem que ser selecionada."
'            Exit Function
'        End If
'    End If
'
'blnDadosGuiaOK = True
'End Function
'
'Private Sub LimpaObjetos()
'    dbc_strInscricaoInicial.BoundText = ""
'    dbc_strInscricaoFinal.BoundText = ""
'    txt_intExercicio.Text = ""
'    txt_DataDeVencimento.Text = ""
'    dbc_intMensagem1.BoundText = ""
'    dbc_intMensagem2.BoundText = ""
'    txt_Mensagem1 = ""
'    txt_Mensagem2 = ""
'    dbc_strInscricaoInicial.SetFocus
'End Sub
'
'Private Sub dbc_intMensagem1_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", dbc_intMensagem1
'End Sub
'
'Private Sub dbc_intMensagem2_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", dbc_intMensagem2
'End Sub
'
'Private Sub txt_DataDeVencimento_GotFocus()
'    MarcaCampo txt_DataDeVencimento
'End Sub
'
'Private Sub txt_DataDeVencimento_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "D", txt_DataDeVencimento
'End Sub
'
'Private Sub txt_intExercicio_GotFocus()
'    MarcaCampo txt_intExercicio
'End Sub
'
'Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "N", txt_intExercicio
'End Sub
'
'Private Sub dbc_strInscricaoInicial_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", dbc_strInscricaoInicial
'End Sub
'
'Private Sub dbc_strInscricaoFinal_KeyPress(KeyAscii As Integer)
'    CaracterValido KeyAscii, "A", dbc_strInscricaoFinal
'End Sub
'
'
