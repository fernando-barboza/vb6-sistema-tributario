VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadCalculoITBIUrbanoRural 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo do ITBI Urbano e Rural"
   ClientHeight    =   6660
   ClientLeft      =   -150
   ClientTop       =   495
   ClientWidth     =   8700
   Icon            =   "cadCalculoITBIUrbanoRural.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   6405
      Left            =   210
      TabIndex        =   21
      Top             =   150
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   11298
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cálculo do ITBI Urbano e Rural"
      TabPicture(0)   =   "cadCalculoITBIUrbanoRural.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strInscricaoCadastral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_intComposicaoDaReceita"
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
      Tab(0).Control(7)=   "txt_PKId"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fra_TipoITBI"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_Proprietario"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Composição da Receita"
      TabPicture(1)   =   "cadCalculoITBIUrbanoRural.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdb_Atividades"
      Tab(1).Control(1)=   "fra_Frame"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Emissão de Guias de Arrecadação"
      TabPicture(2)   =   "cadCalculoITBIUrbanoRural.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_EmissaoDeGuias"
      Tab(2).ControlCount=   1
      Begin VB.Frame fra_Frame 
         Height          =   1695
         Left            =   -74820
         TabIndex        =   63
         Top             =   420
         Width           =   7965
         Begin VB.TextBox txt_intParcelaInicial 
            Height          =   285
            Left            =   2325
            MaxLength       =   15
            TabIndex        =   8
            Top             =   930
            Width           =   540
         End
         Begin VB.TextBox txt_intParcelaFinal 
            Height          =   285
            Left            =   3330
            MaxLength       =   15
            TabIndex        =   9
            Top             =   930
            Width           =   480
         End
         Begin VB.TextBox txt_dblAliquota 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2325
            MaxLength       =   3
            TabIndex        =   10
            Top             =   1290
            Width           =   1305
         End
         Begin VB.TextBox txt_dblbValorITBI 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2325
            MaxLength       =   15
            TabIndex        =   6
            Top             =   210
            Width           =   1500
         End
         Begin VB.TextBox txt_dblDesconto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6255
            MaxLength       =   3
            TabIndex        =   13
            Top             =   1305
            Width           =   1005
         End
         Begin VB.TextBox txt_intIntervalo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6255
            MaxLength       =   3
            TabIndex        =   12
            Top             =   930
            Width           =   1005
         End
         Begin VB.TextBox txt_dtmDataVencimento 
            Height          =   285
            Left            =   6255
            MaxLength       =   15
            TabIndex        =   11
            Top             =   570
            Width           =   1005
         End
         Begin VB.TextBox txt_dtmDataPagamento 
            Height          =   285
            Left            =   2325
            MaxLength       =   15
            TabIndex        =   7
            Top             =   585
            Width           =   1485
         End
         Begin VB.Label lbl_ParcelaIncial 
            AutoSize        =   -1  'True
            Caption         =   "Parcela"
            Height          =   195
            Left            =   1650
            TabIndex        =   75
            Top             =   975
            Width           =   540
         End
         Begin VB.Label lbl_ParcelaFinal 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   2925
            TabIndex        =   74
            Top             =   975
            Width           =   225
         End
         Begin VB.Label lbl_dias 
            AutoSize        =   -1  'True
            Caption         =   "dias."
            Height          =   195
            Left            =   7335
            TabIndex        =   72
            Top             =   1020
            Width           =   330
         End
         Begin VB.Label lbl_dblAliquota 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota sobre Valor"
            Height          =   195
            Left            =   795
            TabIndex        =   71
            Top             =   1380
            Width           =   1440
         End
         Begin VB.Label lbl_dblbvalorITBI 
            AutoSize        =   -1  'True
            Caption         =   "Valor de Avaliação do Imóvel"
            Height          =   195
            Left            =   165
            TabIndex        =   70
            Top             =   285
            Width           =   2070
         End
         Begin VB.Label lbl_dblDesconto 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
            Height          =   195
            Left            =   5475
            TabIndex        =   69
            Top             =   1380
            Width           =   690
         End
         Begin VB.Label lbl_intIntervalo 
            AutoSize        =   -1  'True
            Caption         =   "Intervalo entre Parcelas"
            Height          =   195
            Left            =   4485
            TabIndex        =   68
            Top             =   990
            Width           =   1680
         End
         Begin VB.Label lbl_dtmDataVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vencimento"
            Height          =   195
            Left            =   4710
            TabIndex        =   67
            Top             =   645
            Width           =   1455
         End
         Begin VB.Label lbl_dtmDataPagamento 
            AutoSize        =   -1  'True
            Caption         =   "Data de Lançamento"
            Height          =   195
            Left            =   735
            TabIndex        =   66
            Top             =   660
            Width           =   1500
         End
         Begin VB.Label lbl_p1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   3720
            TabIndex        =   65
            Top             =   1410
            Width           =   120
         End
         Begin VB.Label lbl_p2 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   7335
            TabIndex        =   64
            Top             =   1395
            Width           =   120
         End
      End
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5415
         Left            =   -74820
         TabIndex        =   58
         Top             =   420
         Width           =   7965
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1695
            Left            =   480
            TabIndex        =   61
            Top             =   2970
            Width           =   6945
            Begin VB.TextBox txt_Mensagem2 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   17
               Top             =   0
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   18
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label lbl_Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   62
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1695
            Left            =   480
            TabIndex        =   59
            Top             =   930
            Width           =   6945
            Begin VB.TextBox txt_Mensagem1 
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
            Begin VB.CheckBox chk_EmBranco1 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   14
               Top             =   0
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem1 
               Height          =   315
               Left            =   1080
               TabIndex        =   15
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label lbl_Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   60
               Top             =   390
               Width           =   780
            End
         End
      End
      Begin VB.Frame fra_Proprietario 
         Height          =   2430
         Left            =   120
         TabIndex        =   34
         Top             =   1170
         Width           =   8070
         Begin VB.TextBox txt_intContribuinte 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   210
            Width           =   735
         End
         Begin VB.TextBox txt_strCNPJCPFP 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   570
            Width           =   1845
         End
         Begin VB.Frame fra_Endereco 
            Caption         =   "Endereço"
            Height          =   1395
            Left            =   120
            TabIndex        =   36
            Top             =   900
            Width           =   7845
            Begin VB.TextBox txt_Distrito 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   990
               Width           =   3525
            End
            Begin VB.TextBox txt_Bairro 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   630
               Width           =   3105
            End
            Begin VB.TextBox txt_Logradouro 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   270
               Width           =   4005
            End
            Begin VB.TextBox txt_Numero 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   270
               Width           =   795
            End
            Begin VB.TextBox txt_Complemento 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   6810
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   270
               Width           =   870
            End
            Begin VB.TextBox txt_Cep 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   6615
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   990
               Width           =   1080
            End
            Begin VB.TextBox txt_Municipio 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5070
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   630
               Width           =   2625
            End
            Begin VB.TextBox txt_UF 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5070
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   990
               Width           =   510
            End
            Begin VB.Label lbl_intCepC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   6240
               TabIndex        =   52
               Top             =   1050
               Width           =   285
            End
            Begin VB.Label lbl_intUFC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   4770
               TabIndex        =   51
               Top             =   1065
               Width           =   210
            End
            Begin VB.Label lbl_strComplementoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6300
               TabIndex        =   50
               Top             =   330
               Width           =   480
            End
            Begin VB.Label lbl_intNumeroC 
               AutoSize        =   -1  'True
               Caption         =   "Nº"
               Height          =   195
               Left            =   5130
               TabIndex        =   49
               Top             =   330
               Width           =   180
            End
            Begin VB.Label lbl_intLogradouroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   120
               TabIndex        =   48
               Top             =   330
               Width           =   810
            End
            Begin VB.Label lbl_intBairroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   525
               TabIndex        =   47
               Top             =   690
               Width           =   405
            End
            Begin VB.Label lbl_intMunicipioC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   4290
               TabIndex        =   46
               Top             =   690
               Width           =   705
            End
            Begin VB.Label lbl_strDistritoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               Height          =   195
               Left            =   450
               TabIndex        =   45
               Top             =   1050
               Width           =   480
            End
         End
         Begin VB.TextBox txt_strNome 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   210
            Width           =   6135
         End
         Begin VB.Label lbl_strCNPJCPFP 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   150
            TabIndex        =   54
            Top             =   615
            Width           =   780
         End
         Begin VB.Label lbl_strNome 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   135
            TabIndex        =   53
            Top             =   255
            Width           =   795
         End
      End
      Begin VB.Frame fra_TipoITBI 
         Caption         =   "Tipo de ITBI"
         Height          =   615
         Left            =   6210
         TabIndex        =   32
         Top             =   510
         Width           =   1965
         Begin VB.OptionButton opt_Urbano_Rural 
            Caption         =   "Rural"
            Height          =   255
            Index           =   1
            Left            =   1170
            TabIndex        =   5
            Top             =   240
            Width           =   705
         End
         Begin VB.OptionButton opt_Urbano_Rural 
            Caption         =   "Urbano"
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   1
            Top             =   210
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.TextBox txt_PKId 
         Height          =   285
         Left            =   7680
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton cmd_Down 
         Height          =   285
         Left            =   -67470
         Picture         =   "cadCalculoITBIUrbanoRural.frx":1096
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Abaixo"
         Top             =   1725
         Width           =   300
      End
      Begin VB.CommandButton cmd_Up 
         Height          =   285
         Left            =   -67470
         Picture         =   "cadCalculoITBIUrbanoRural.frx":11E0
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Acima"
         Top             =   1425
         Width           =   300
      End
      Begin VB.TextBox txt_DescricaoConteudo 
         Height          =   285
         Left            =   -73140
         MaxLength       =   50
         TabIndex        =   23
         Top             =   1020
         Width           =   3945
      End
      Begin VB.TextBox txt_Conteudo 
         Height          =   285
         Left            =   -73140
         MaxLength       =   50
         TabIndex        =   20
         Top             =   720
         Width           =   3945
      End
      Begin MSComctlLib.ListView lvw_TipoComunicacao 
         Height          =   3210
         Left            =   -74760
         TabIndex        =   26
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
         TabIndex        =   27
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
            TabIndex        =   28
            Top             =   30
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "img_Aux"
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
               Picture         =   "cadCalculoITBIUrbanoRural.frx":132A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadCalculoITBIUrbanoRural.frx":148A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "cadCalculoITBIUrbanoRural.frx":15E6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastral 
         Height          =   315
         Left            =   1620
         TabIndex        =   0
         Top             =   810
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Proprietarios 
         Height          =   1695
         Left            =   180
         TabIndex        =   4
         Top             =   4530
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   2990
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
      Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
         Height          =   3615
         Left            =   -74850
         TabIndex        =   22
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
         Splits(0)._ColumnProps(14)=   "Column(2).Width=10848"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=10769"
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
      Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
         Height          =   315
         Left            =   1980
         TabIndex        =   2
         Top             =   3720
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intOcorrencia 
         Height          =   315
         Left            =   1980
         TabIndex        =   3
         Top             =   4110
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lbl_intOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   1050
         TabIndex        =   73
         Top             =   4230
         Width           =   780
      End
      Begin VB.Label lbl_intComposicaoDaReceita 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   150
         TabIndex        =   56
         Top             =   3825
         Width           =   1695
      End
      Begin VB.Label lbl_strInscricaoCadastral 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   150
         TabIndex        =   33
         Top             =   930
         Width           =   1350
      End
      Begin VB.Label lbl_TipoComunicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   -73770
         TabIndex        =   30
         Top             =   720
         Width           =   315
      End
      Begin VB.Label lbl_DescricaoConteudo 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   -74175
         TabIndex        =   29
         Top             =   1110
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCadCalculoITBIUrbanoRural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim mblnAlterando                As Boolean
Dim mobjAux                      As Object
Dim mblnSelecionou               As Boolean
Dim mblnPrimeiraVez              As Boolean
Dim xarReceita                   As XArrayDB
Dim adoResultado                 As ADODB.Recordset
Dim bytOPT                       As Byte

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
    If Area = 2 And dbc_intComposicaoDaReceita.MatchedWithList Then
        MontaAtividade dbc_intComposicaoDaReceita.BoundText
    End If
End Sub

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

Private Sub dbc_strInscricaoCadastral_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, Area
    If Area = 2 And dbc_strInscricaoCadastral.MatchedWithList Then
        txt_PKId.Text = dbc_strInscricaoCadastral.BoundText
        BuscarDadosProprietario (0)
    End If
End Sub

Private Sub dbc_strInscricaoCadastral_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoCadastral, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 672
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
    dbc_strInscricaoCadastral.Tag = strQueryImobiliario(0) & ";I.strInscricaoAnterior"
    'LeDaTabelaParaObj gstrImobiliario, tdb_Proprietarios, strQueryImobiliario(1)
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoDaReceita, QueryComposicao
    VerificaMascaraInscricao
    VerificaObjParaAplicar mobjAux
    LeDaTabelaParaObj gstrOcorrencia, dbc_intOcorrencia, strQuerryOcorrencia

'''GUIA
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem1, strQueryMensagem
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem2, strQueryMensagem
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub opt_Urbano_Rural_Click(Index As Integer)
    bytOPT = Index
    If Index = 0 Then
        LimpaEndereco
        dbc_strInscricaoCadastral.Text = ""
        dbc_intComposicaoDaReceita.Text = ""
        MontaAtividade 0 ' limpar o grid com as taxas. informei o valor 0 para não dar pau na query
        tdb_Atividades.DataSource = Nothing
        tdb_Proprietarios.DataSource = Nothing
        LeDaTabelaParaObj gstrImobiliario, dbc_strInscricaoCadastral, strQueryImobiliario(0)
        LeDaTabelaParaObj gstrImobiliario, tdb_Proprietarios, strQueryImobiliario(1)
        mblnPrimeiraVez = False
    Else
        LimpaEndereco
        dbc_strInscricaoCadastral.Text = ""
        dbc_intComposicaoDaReceita.Text = ""
        MontaAtividade 0 ' limpar o grid com as taxas. informei o valor 0 para não dar pau na query
        tdb_Atividades.DataSource = Nothing
        tdb_Proprietarios.DataSource = Nothing
        LeDaTabelaParaObj gstrImobiliarioRural, dbc_strInscricaoCadastral, strQueryImobiliario(0)
        LeDaTabelaParaObj gstrImobiliario, tdb_Proprietarios, strQueryImobiliario(1)
        mblnPrimeiraVez = False
    End If
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3DPasta.Tab = 1 Then txt_dblbValorITBI.SetFocus
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
    mblnPrimeiraVez = False
End Sub

Private Sub tdb_Proprietarios_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Proprietarios
        If Not .EOF And Not .BOF Then
            txt_PKId.Text = .Columns("PKID").Value

            If mblnPrimeiraVez Then

                BuscarDadosProprietario (1)

                gCorLinhaSelecionada tdb_Proprietarios

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar

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
    Dim strSql              As String
    Dim blnExiteLancamento  As Boolean
    Dim strInscricao        As String
    
    If UCase(strModoOperacao) = gstrPreencherLista Or UCase(strModoOperacao) = gstrLocalizar Then
        strSql = strQueryImobiliario(1)
        ToolBarGeral strModoOperacao, IIf((opt_Urbano_Rural(0).Value = True), gstrImobiliario, gstrImobiliarioRural), False, tdb_Proprietarios, Me, mobjAux, strSql
        Exit Sub
    End If
    
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
            strSql = gstrQueryRelatorioGuiaDeArrecadacao(blnExiteLancamento, strInscricao, strInscricao, gintExercicio, dbc_intComposicaoDaReceita.BoundText, , , Val(txt_intParcelaInicial.Text), Val(txt_intParcelaFinal.Text))
            If blnExiteLancamento Then
                Set gfrmFormularioQueEstaImprimindoGuia = Me
                rptGuiaDeArrecadacaoMunicipal.strImposto = dbc_intComposicaoDaReceita.Text
                ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, strSql
            End If
        End If
    Else
        If UCase(strModoOperacao) = gstrCalcularReajuste Then
            CalculoLancamento
        End If
    End If
    
    
End Sub


'====================

Private Sub CalculoLancamento()

'******************************************************************************************
' Data: 08/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 08/05/2003
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
'        strSql = "sp_CalculoParaUsuario '" & strInscricao & "','" & strPKId & "'," & bytOPT & _
'                 "," & dbc_intComposicaoDaReceita.BoundText & ",0,0,0,'" & gstrConvVrParaSql(txt_dblbValorITBI) & _
'                 "," & gstrConvVrParaSql(txt_dblAliquota) & "," & gstrConvVrParaSql(txt_dblDesconto) & _
'                 ", @dblValor OUTPUT'"
        strSql = gstrStoredProcedure("sp_CalculoParaUsuario", "'" & strInscricao & "','" & strPKId & "'," & bytOPT & _
                 "," & dbc_intComposicaoDaReceita.BoundText & ",0,0,0,'" & gstrConvVrParaSql(txt_dblbValorITBI) & _
                 "," & gstrConvVrParaSql(txt_dblAliquota) & "," & gstrConvVrParaSql(txt_dblDesconto) & _
                 ", " & IIf((bytDBType = EDatabases.Oracle), ":dblValor", "@dblValor OUTPUT") & "'")
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
                    If Not gBlnVerificaLancamentos(gintExercicio, dbc_intComposicaoDaReceita.BoundText, _
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

    Dim strSql          As String
    Dim strInscricao    As String

    strInscricao = Trim(Left(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, "-", vbTextCompare) - 1))
    strSql = " SELECT A.intContribuinte, A.strInscricaoAnterior FROM "
    If bytOPT = 1 Then
'        strSQL = strSQL & gstrImobiliarioRural & " AS A, "
        strSql = strSql & gstrImobiliarioRural & " A, "
    Else
'        strSQL = strSQL & gstrImobiliario & " AS A, "
        strSql = strSql & gstrImobiliario & " A, "
    End If
'    strSQL = strSQL & gstrContribuinte & " AS CO "
    strSql = strSql & gstrContribuinte & " CO " & _
            " WHERE " & _
            " CO.PKId = A.intContribuinte " & _
            " AND A.strInscricaoAnterior = """ & strInscricao & """"
    strContribuintes = strSql
End Function

Private Function strLancamentos(dblValorAparcelar As String, dblValorNaoParcelado As String) As String
    Dim strSql As String
    
'    strSql = " sp_CalculoLancamentoReceitas 1, " & gstrConvVrParaSql(dblValorAparcelar) & "," & _
'            gstrConvVrParaSql(dblValorNaoParcelado) & _
'            ",'" & strPKId & "','" & strContribuintes & "'," & gintExercicio & _
'            "," & dbc_intComposicaoDaReceita.BoundText & ", NULL, " & _
'            gstrConvDtParaSql(txt_dtmDataPagamento.Text) & _
'            "," & gstrConvDtParaSql(txt_dtmDataVencimento) & "," & Val(txt_intParcelaInicial.Text) & "," & _
'            Val(txt_intParcelaFinal.Text) & "," & Val(txt_intIntervalo.Text) & ",1," & Val(dbc_intOcorrencia.BoundText) & _
'            "," & bytOPT & "," & glngCodUsr & ",0,0," & gstrConvVrParaSql(txt_dblAliquota) & ",'?" & _
'            gstrConvVrParaSql(txt_dblbValorITBI) & "," & gstrConvVrParaSql(txt_dblAliquota) & _
'            "," & gstrConvVrParaSql(txt_dblDesconto) & ", @dblValor OUTPUT'"
    strSql = gstrStoredProcedure("sp_CalculoLancamentoReceitas", "1, " & gstrConvVrParaSql(dblValorAparcelar) & "," & _
            gstrConvVrParaSql(dblValorNaoParcelado) & _
            ",'" & strPKId & "','" & strContribuintes & "'," & gintExercicio & _
            "," & dbc_intComposicaoDaReceita.BoundText & ", NULL, " & _
            gstrConvDtParaSql(txt_dtmDataPagamento.Text) & _
            "," & gstrConvDtParaSql(txt_dtmDataVencimento) & "," & Val(txt_intParcelaInicial.Text) & "," & _
            Val(txt_intParcelaFinal.Text) & "," & Val(txt_intIntervalo.Text) & ",1," & Val(dbc_intOcorrencia.BoundText) & _
            "," & bytOPT & "," & glngCodUsr & ",0,0," & gstrConvVrParaSql(txt_dblAliquota) & ",'?" & _
            gstrConvVrParaSql(txt_dblbValorITBI) & "," & gstrConvVrParaSql(txt_dblAliquota) & _
            "," & gstrConvVrParaSql(txt_dblDesconto) & _
            ", " & IIf((bytDBType = EDatabases.Oracle), ":dblValor", "@dblValor OUTPUT") & "'")
    strLancamentos = strSql
End Function

'====================

Private Function QueryComposicao() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita
    strSql = strSql & " WHERE intUtilizacao = 1 "
    strSql = strSql & " ORDER BY strDescricao "
    QueryComposicao = strSql
End Function

Private Function strQueryImobiliario(bytObjeto As Byte) As String

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
'        strSQL = " SELECT I.PKId AS PKId, (I.strInscricaoAnterior + ' - ' + C.strNome ) AS Inscricao "
        strSql = " SELECT I.PKId AS PKId, (I.strInscricaoAnterior " & strCONCAT & " ' - ' " & strCONCAT & " C.strNome ) AS Inscricao "
    Else                  'True DBGrid  tdb_Proprietarios
        strSql = " SELECT I.PKId AS PKId, I.strInscricaoAnterior AS strInscricaoAnterior, C.strNome AS strNome "
    End If
    strSql = strSql & " FROM " & IIf((opt_Urbano_Rural(0).Value = True), gstrImobiliario, gstrImobiliarioRural) & _
            " I, " & gstrContribuinte & " C " & _
            " WHERE I.intContribuinte = C.PKId "
'            " ORDER BY CONVERT(NUMERIC, I.strInscricaoAnterior)"
    strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "I.strInscricaoAnterior")
    strQueryImobiliario = strSql
End Function

Private Function BuscarDadosProprietario(bytObjeto As Byte)

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql          As String
    Dim strAux          As String
    Dim adoResultado    As ADODB.Recordset
    
    strSql = " SELECT A.PKId, CO.PKId AS CodigoContribuinte, CO.strNome, " & _
             " LG.strDescricao AS Logradouro, A.intNumero, " & _
             " A.strComplemento , A.intCep, BA.strDescricao AS Bairro, " & _
             " UF.strSigla AS UF, A.strCNPJCPF, " & _
             " CO.strDistritoC, CI.strDescricao AS Municipio, "
'             " ISNULL(A.dblValorITBI,0) AS dblValorITBI "
    strSql = strSql & gstrISNULL("A.dblValorITBI", "0") & " AS dblValorITBI " & _
             " FROM " & IIf(opt_Urbano_Rural(1).Value = True, gstrImobiliarioRural, gstrImobiliario)
'             " AS A, " & gstrContribuinte & " AS CO, "
    strSql = strSql & " A, " & gstrContribuinte & " CO, "
'             gstrLogradouro & " AS LG, "
    strSql = strSql & gstrLogradouro & " LG, "
'             gstrCidade & " AS CI, "
    strSql = strSql & gstrCidade & " CI, "
'             gstrBairro & " AS BA, "
    strSql = strSql & gstrBairro & " BA, "
'             gstrUF & " AS UF "
    strSql = strSql & gstrUF & " UF " & _
             " WHERE CO.PKId = A.intContribuinte "
'             " AND CI.PKId =* CO.intMunicipioC "
    strSql = strSql & " AND CI.PKId =" & strOUTJSQLServer & " CO.intMunicipioC " & strOUTJOracle
'             " AND LG.PKId =* A.intLogradouro "
    strSql = strSql & " AND LG.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & " A.intLogradouro "
'             " AND BA.PKId =* A.intBairro "
    strSql = strSql & " AND BA.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & " A.intBairro "
'             " AND UF.PKId =* A.intUF "
    strSql = strSql & " AND UF.PKId " & strOUTJOracle & "=" & strOUTJSQLServer & " A.intUF " & _
             " AND A.PKID = " & IIf(bytObjeto = 0, dbc_strInscricaoCadastral.BoundText, txt_PKId.Text)
    
    LimpaEndereco
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                dbc_strInscricaoCadastral.BoundText = gstrENulo(!Pkid)
                txt_strCNPJCPFP = gstrCGCCPFFormatado(gstrENulo(!StrCnpjCpf))
                txt_strNome.Text = gstrENulo(!STRNOME)
                txt_intContribuinte.Text = gstrENulo(!CodigoContribuinte)
                txt_Logradouro.Text = gstrENulo(!Logradouro)
                txt_Numero.Text = gstrENulo(!INTNUMERO)
                txt_Complemento.Text = gstrENulo(!STRCOMPLEMENTO)
                txt_Bairro.Text = gstrENulo(!Bairro)
                txt_Municipio.Text = gstrENulo(!Municipio)
                txt_Distrito.Text = gstrENulo(!strDistritoC)
                txt_UF.Text = gstrENulo(!UF)
                txt_Cep.Text = gstrCEPFormatado(gstrENulo(!INTCEP))
                txt_dblbValorITBI = gstrConvVrDoSql(gstrENulo(!dblValorITBI))
            End With
        End If
    End If
    
End Function

Private Sub MontaAtividade(intComposicaoReceita As Integer)
Dim strSql As String
Dim adoRec As ADODB.Recordset
Dim varAux As String

On Error GoTo Err_Handle

Set xarReceita = New XArrayDB
xarReceita.Clear

xarReceita.ReDim 0, 0, 0, 3

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

Sub VerificaMascaraInscricao()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    Dim strMascara   As String
    
    strMascara = ""
    
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & TYP_DIVIDA_ATIVA
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

Private Function dblRetornaIndexador() As Double
Dim strSql As String
Dim adoRec As Recordset
Dim dblIndexador As Double

strSql = ""
strSql = strSql & "Select IE.PKId, IE.dblValor, IE.dtmData, "
strSql = strSql & "E.strSiglaIndexador Sigla "
strSql = strSql & "FROM " & gstrIndiceEconomico & " IE, " & gstrIndexadorEconomico & " E "
strSql = strSql & "WHERE IE.intIndexador = E.PKId "
strSql = strSql & "AND IE.dtmData in (SELECT MAX(IE.dtmData) FROM " & gstrIndiceEconomico & " IE )"

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    With adoRec
        If Not (.BOF And .EOF) Then
            dblIndexador = !DBLVALOR
        Else
            dblIndexador = 0
        End If
    End With
End If
dblRetornaIndexador = dblIndexador
End Function

'==Para Efetuar Cálculo==
Private Function blnDadosOk() As Boolean
blnDadosOk = False
    Dim i As Integer
    If txt_dblbValorITBI.Text = "" Then
        ExibeMensagem "O campo " & lbl_dblbvalorITBI.Caption & " não pode ser nulo."
        tab_3DPasta.Tab = 1
        txt_dblbValorITBI.SetFocus
        Exit Function
    End If
    If Not ValorITBIValido Then
        tab_3DPasta.Tab = 1
        txt_dblbValorITBI.SetFocus
        Exit Function
    End If

    If txt_dtmDataPagamento.Text = "" Then
        ExibeMensagem "O campo " & lbl_dtmDataPagamento.Caption & " não pode ser nulo."
        tab_3DPasta.Tab = 1
        txt_dtmDataPagamento.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_dtmDataPagamento.Text) = False Then
            ExibeMensagem "A data de lançamento não é válida."
            tab_3DPasta.Tab = 1
            txt_dtmDataPagamento.SetFocus
            Exit Function
        End If
    End If
    
    If dbc_intOcorrencia.MatchedWithList = False Then
        ExibeMensagem "O campo Ocorrência não pode ser Nulo"
        tab_3DPasta.Tab = 0
        dbc_intOcorrencia.SetFocus
        Exit Function
    End If
    
    If txt_dtmDataVencimento.Text = "" Then
        ExibeMensagem "O campo " & lbl_dtmDataVencimento.Caption & " não pode ser nulo."
        tab_3DPasta.Tab = 1
        txt_dtmDataVencimento.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_dtmDataVencimento.Text) = False Then
            ExibeMensagem "A data de vencimento não é válida."
            tab_3DPasta.Tab = 1
            txt_dtmDataVencimento.SetFocus
            Exit Function
        End If
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
Private Function ValorITBIValido() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 12/05/2003
' Alteração: - Substituição do bloco Transact-SQL por uma cláusula SELECT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql As String
    Dim adoResultado As ADODB.Recordset
    
    Dim blnITBIValido As Boolean
    Dim dblValorITBI As Double
    
    blnITBIValido = True
    
'    strSql = "IF(SELECT ISNULL(dblValorITBI,0) FROM " & _
'             IIf(bytOPT = 0, gstrImobiliarioRural, gstrImobiliario) & _
'             " WHERE PKId = " & txt_PKId.Text & ") > " & _
'             gstrConvVrParaSql(txt_dblbValorITBI.Text) & _
'             " SELECT 0 AS bitITBIValido ELSE " & _
'             " SELECT 1 AS bitITBIValido "
    strSql = "SELECT " & gstrISNULL("dblValorITBI", "0") & " dblValorITBI " & _
                " FROM " & _
             IIf(bytOPT = 0, gstrImobiliarioRural, gstrImobiliario) & _
             " WHERE PKId = " & txt_PKId.Text
             
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
    
        If Not (adoResultado Is Nothing) Then
            
            With adoResultado
                
                If (Not .BOF) And (Not .EOF) Then
                    
                    If LTrim(txt_dblbValorITBI.Text) <> "" Then
                        dblValorITBI = CDbl(txt_dblbValorITBI.Text)
                    End If
            
                    blnITBIValido = (IIf((IsNull(!dblValorITBI)), 0, !dblValorITBI) <= dblValorITBI)
            
                End If
            
            End With
        End If
    
'        If (adoResultado!bitITBIValido) = 0 Then
        If (Not blnITBIValido) Then
            If MsgBox("O Valor de Avaliação do Imóvel indicado no Cadastro Imobiliário é maior que o valor informado " & _
                      Chr(13) & "no cálculo do ITBI. Deseja continuar com o Cálculo ? ", vbYesNo) = vbYes Then
                ValorITBIValido = True
            End If
        Else
            ValorITBIValido = True
        End If
    End If
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


'''''#######################Caracter valido e Marca Campo #####################'''''

Private Sub txt_dtmDataPagamento_GotFocus()
    MarcaCampo txt_dtmDataPagamento
End Sub

Private Sub txt_dtmDataPagamento_LostFocus()
    txt_dtmDataPagamento = gstrDataFormatada(txt_dtmDataPagamento)
End Sub

Private Sub txt_dtmDataPagamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataPagamento
End Sub

Private Sub txt_dtmDataVencimento_GotFocus()
    MarcaCampo txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_LostFocus()
    txt_dtmDataVencimento = gstrDataFormatada(txt_dtmDataVencimento)
End Sub

Private Sub txt_intIntervalo_GotFocus()
    MarcaCampo txt_intIntervalo
End Sub

Private Sub txt_intIntervalo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intIntervalo
End Sub

Private Sub txt_intParcelaInicial_GotFocus()
    MarcaCampo txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaFinal_GotFocus()
    MarcaCampo txt_intParcelaFinal
End Sub

Private Sub txt_intParcelaInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaFinal
End Sub

Private Sub txt_dblAliquota_GotFocus()
    MarcaCampo txt_dblAliquota
End Sub

Private Sub txt_dblAliquota_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblAliquota
End Sub

Private Sub txt_dblDesconto_GotFocus()
    MarcaCampo txt_dblDesconto
End Sub

Private Sub txt_dblDesconto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblDesconto
End Sub

Private Sub txt_dblbvalorITBI_GotFocus()
    MarcaCampo txt_dblbValorITBI
End Sub

Private Sub txt_dblbvalorITBI_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblbValorITBI
End Sub

Private Sub txt_dblbValorITBI_LostFocus()
    txt_dblbValorITBI = gstrConvVrDoSql(txt_dblbValorITBI)
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
    strSql = strSql & " SELECT intContribuinte, strInscricaoAnterior "
    strSql = strSql & " FROM "
    If bytOPT = 0 Then
        strSql = strSql & gstrImobiliario
    Else
        strSql = strSql & gstrImobiliarioRural
    End If
'    strSql = strSql & " ORDER BY CONVERT(NUMERIC, strInscricaoAnterior) "
    strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoAnterior")
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
    
    If dbc_intComposicaoDaReceita.Text = "" Then
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

Private Sub txt_Mensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Mensagem1
End Sub

Private Sub txt_Mensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Mensagem2
End Sub
