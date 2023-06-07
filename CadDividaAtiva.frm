VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadDividaAtiva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dívida Ativa"
   ClientHeight    =   9000
   ClientLeft      =   1785
   ClientTop       =   1920
   ClientWidth     =   10635
   HelpContextID   =   5
   Icon            =   "CadDividaAtiva.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   2580
      TabIndex        =   41
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   7335
      Left            =   90
      TabIndex        =   40
      Top             =   30
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dívida Ativa"
      TabPicture(0)   =   "CadDividaAtiva.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_Parcelas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_DomicilioFiscal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Notificação"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Fra_Titulo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_PrescricaoDoDebito"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_Valores"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Observação"
      TabPicture(1)   =   "CadDividaAtiva.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_strindexador"
      Tab(1).Control(1)=   "lbl_dblvlindexador"
      Tab(1).Control(2)=   "fra_Historico"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(4)=   "txtstrindexador"
      Tab(1).Control(5)=   "txtdblvlindexador"
      Tab(1).ControlCount=   6
      Begin VB.Frame fra_Valores 
         Caption         =   "Valores"
         Height          =   855
         Left            =   180
         TabIndex        =   89
         Top             =   4740
         Width           =   10035
         Begin VB.TextBox txt_dblvalorimposto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   1290
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   1725
         End
         Begin VB.TextBox txt_dblvalortaxas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   3660
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Impostos"
            Height          =   195
            Left            =   570
            TabIndex        =   91
            Top             =   450
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Taxas"
            Height          =   195
            Left            =   3090
            TabIndex        =   90
            Top             =   450
            Width           =   435
         End
      End
      Begin VB.Frame fra_PrescricaoDoDebito 
         Caption         =   "Contribuinte"
         Height          =   915
         Left            =   180
         TabIndex        =   84
         Top             =   1320
         Width           =   10035
         Begin VB.TextBox txtstrpromissario 
            Height          =   285
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   12
            Top             =   540
            Width           =   8625
         End
         Begin VB.TextBox txtstrnomeproprietario 
            Height          =   285
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   9
            Top             =   180
            Width           =   4635
         End
         Begin VB.TextBox txtstrcnpjcpf 
            Height          =   285
            Left            =   8760
            TabIndex        =   11
            Top             =   180
            Width           =   1155
         End
         Begin VB.TextBox txtstridentidade 
            Height          =   285
            Left            =   6780
            TabIndex        =   10
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label lbl_Prescricao 
            AutoSize        =   -1  'True
            Caption         =   "Promissário"
            Height          =   195
            Left            =   450
            TabIndex        =   88
            Top             =   600
            Width           =   795
         End
         Begin VB.Label lbl_nome 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   450
            TabIndex        =   87
            Top             =   270
            Width           =   795
         End
         Begin VB.Label lbl_CNPJCPF 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ/CPF"
            Height          =   195
            Left            =   7950
            TabIndex        =   86
            Top             =   270
            Width           =   780
         End
         Begin VB.Label lbl_identidade 
            AutoSize        =   -1  'True
            Caption         =   "Identidade"
            Height          =   195
            Left            =   6000
            TabIndex        =   85
            Top             =   270
            Width           =   750
         End
      End
      Begin VB.TextBox txtdblvlindexador 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -71520
         MaxLength       =   25
         TabIndex        =   82
         Top             =   1500
         Width           =   1605
      End
      Begin VB.TextBox txtstrindexador 
         Height          =   285
         Left            =   -73950
         MaxLength       =   20
         TabIndex        =   80
         Top             =   1500
         Width           =   945
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   -74820
         TabIndex        =   70
         Top             =   330
         Width           =   10035
         Begin VB.TextBox txtstrAviso2 
            Height          =   285
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   34
            Top             =   570
            Width           =   945
         End
         Begin VB.TextBox txtdtmdtinscricao2 
            Height          =   285
            Left            =   4080
            TabIndex        =   35
            Top             =   570
            Width           =   1005
         End
         Begin VB.TextBox txtintExercicio2 
            Height          =   285
            Left            =   9420
            MaxLength       =   8
            TabIndex        =   33
            Top             =   150
            Width           =   495
         End
         Begin VB.TextBox txtintlivro2 
            Height          =   285
            Left            =   8970
            MaxLength       =   8
            TabIndex        =   38
            Top             =   570
            Width           =   945
         End
         Begin VB.TextBox txtintfolha2 
            Height          =   285
            Left            =   7740
            MaxLength       =   4
            TabIndex        =   37
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtstrIncricao2 
            Height          =   285
            Left            =   7305
            TabIndex        =   32
            Top             =   150
            Width           =   1305
         End
         Begin VB.TextBox txtcadastro2 
            Height          =   285
            Left            =   5115
            TabIndex        =   31
            Top             =   150
            Width           =   1425
         End
         Begin VB.TextBox txtintcertidao2 
            Height          =   285
            Left            =   5820
            TabIndex        =   36
            Top             =   570
            Width           =   1365
         End
         Begin MSDataListLib.DataCombo dbc_intReceita2 
            Height          =   315
            Left            =   1710
            TabIndex        =   30
            Top             =   150
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl_aviso2 
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Left            =   1230
            TabIndex        =   79
            Top             =   660
            Width           =   390
         End
         Begin VB.Label lbl_inscricao2 
            AutoSize        =   -1  'True
            Caption         =   "Data de Inscrição"
            Height          =   195
            Left            =   2760
            TabIndex        =   78
            Top             =   660
            Width           =   1260
         End
         Begin VB.Label lbl_exercicio2 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   8670
            TabIndex        =   77
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lbl_livro2 
            AutoSize        =   -1  'True
            Caption         =   "Livro"
            Height          =   195
            Left            =   8550
            TabIndex        =   76
            Top             =   660
            Width           =   345
         End
         Begin VB.Label lbl_folha2 
            AutoSize        =   -1  'True
            Caption         =   "Folha"
            Height          =   195
            Left            =   7260
            TabIndex        =   75
            Top             =   660
            Width           =   390
         End
         Begin VB.Label lbl_strInscricao2 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição"
            Height          =   195
            Left            =   6615
            TabIndex        =   74
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lbl_cadastro2 
            AutoSize        =   -1  'True
            Caption         =   "Cadastro"
            Height          =   195
            Left            =   4410
            TabIndex        =   73
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lbl_certidao2 
            AutoSize        =   -1  'True
            Caption         =   "Certidão"
            Height          =   195
            Left            =   5160
            TabIndex        =   72
            Top             =   660
            Width           =   585
         End
         Begin VB.Label lbl_compreceita2 
            AutoSize        =   -1  'True
            Caption         =   "Composição da receita"
            Height          =   195
            Left            =   60
            TabIndex        =   71
            Top             =   240
            Width           =   1620
         End
      End
      Begin VB.Frame Fra_Titulo 
         Height          =   975
         Left            =   180
         TabIndex        =   60
         Top             =   330
         Width           =   10035
         Begin VB.TextBox txtintcertidao 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5820
            TabIndex        =   6
            Top             =   570
            Width           =   1365
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Left            =   5115
            TabIndex        =   1
            Top             =   150
            Width           =   1425
         End
         Begin VB.TextBox txtintfolha 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7740
            MaxLength       =   4
            TabIndex        =   7
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox txtintlivro 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8970
            MaxLength       =   8
            TabIndex        =   8
            Top             =   570
            Width           =   945
         End
         Begin VB.TextBox txtintExercicio 
            Height          =   285
            Left            =   9420
            MaxLength       =   8
            TabIndex        =   3
            Top             =   150
            Width           =   495
         End
         Begin VB.TextBox txtdtmdtinscricao 
            Height          =   285
            Left            =   4080
            TabIndex        =   5
            Top             =   570
            Width           =   1005
         End
         Begin VB.TextBox txtstrAviso 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   4
            Top             =   570
            Width           =   945
         End
         Begin MSDataListLib.DataCombo dbc_intReceita 
            Height          =   315
            Left            =   1710
            TabIndex        =   0
            Top             =   150
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSMask.MaskEdBox mskstrInscricao 
            Height          =   285
            Left            =   7320
            TabIndex        =   2
            Top             =   150
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   24
            PromptChar      =   " "
         End
         Begin VB.Label lbl_compreceita 
            AutoSize        =   -1  'True
            Caption         =   "Composição da receita"
            Height          =   195
            Left            =   60
            TabIndex        =   69
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label lbl_certidao 
            AutoSize        =   -1  'True
            Caption         =   "Certidão"
            Height          =   195
            Left            =   5160
            TabIndex        =   68
            Top             =   660
            Width           =   585
         End
         Begin VB.Label lbl_cadastro 
            AutoSize        =   -1  'True
            Caption         =   "Cadastro"
            Height          =   195
            Left            =   4410
            TabIndex        =   67
            Top             =   240
            Width           =   630
         End
         Begin VB.Label lbl_strInscricao 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição"
            Height          =   195
            Left            =   6615
            TabIndex        =   66
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lbl_folha 
            AutoSize        =   -1  'True
            Caption         =   "Folha"
            Height          =   195
            Left            =   7260
            TabIndex        =   65
            Top             =   660
            Width           =   390
         End
         Begin VB.Label lbl_livro 
            AutoSize        =   -1  'True
            Caption         =   "Livro"
            Height          =   195
            Left            =   8550
            TabIndex        =   64
            Top             =   660
            Width           =   345
         End
         Begin VB.Label lbl_exercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   8670
            TabIndex        =   63
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lbl_inscricao 
            AutoSize        =   -1  'True
            Caption         =   "Data de Inscrição"
            Height          =   195
            Left            =   2760
            TabIndex        =   62
            Top             =   660
            Width           =   1260
         End
         Begin VB.Label lbl_aviso 
            AutoSize        =   -1  'True
            Caption         =   "Aviso"
            Height          =   195
            Left            =   1230
            TabIndex        =   61
            Top             =   660
            Width           =   390
         End
      End
      Begin VB.Frame fra_Historico 
         Caption         =   "Histórico"
         Height          =   3795
         Left            =   -74820
         TabIndex        =   55
         Top             =   1980
         Width           =   10035
         Begin VB.TextBox txtHistorico 
            Height          =   3345
            Left            =   120
            MaxLength       =   3000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   270
            Width           =   9825
         End
      End
      Begin VB.Frame fra_Notificação 
         Caption         =   "Endereço de Notificação"
         Height          =   1245
         Left            =   180
         TabIndex        =   49
         Top             =   3480
         Width           =   10035
         Begin VB.TextBox txtstrComplementoN 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   22
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtstrNumeroN 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   5850
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   21
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txt_UFN 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   26
            Top             =   900
            Width           =   405
         End
         Begin VB.TextBox txt_CepN 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   24
            Top             =   570
            Width           =   1005
         End
         Begin VB.TextBox txt_MunicipioN 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   25
            Top             =   900
            Width           =   5535
         End
         Begin VB.TextBox txt_BairroN 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   23
            Top             =   570
            Width           =   5535
         End
         Begin VB.TextBox txt_LogradouroN 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   20
            Top             =   240
            Width           =   4125
         End
         Begin VB.Label lbl_strComplementoC 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6930
            TabIndex        =   59
            Top             =   270
            Width           =   480
         End
         Begin VB.Label lbl_numeroC 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5520
            TabIndex        =   57
            Top             =   270
            Width           =   180
         End
         Begin VB.Label lbl_UFN 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   7170
            TabIndex        =   54
            Top             =   960
            Width           =   210
         End
         Begin VB.Label lbl_CepN 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   7080
            TabIndex        =   53
            Top             =   630
            Width           =   285
         End
         Begin VB.Label lbl_MunicipioN 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   555
            TabIndex        =   52
            Top             =   930
            Width           =   705
         End
         Begin VB.Label lbl_BairroN 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   855
            TabIndex        =   51
            Top             =   600
            Width           =   405
         End
         Begin VB.Label lbl_LogradouroN 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   195
            Left            =   570
            TabIndex        =   50
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.Frame fra_DomicilioFiscal 
         Caption         =   "Local"
         Height          =   1215
         Left            =   180
         TabIndex        =   42
         Top             =   2250
         Width           =   10035
         Begin VB.TextBox txtstrComplemento 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   15
            Top             =   180
            Width           =   1935
         End
         Begin VB.TextBox txtstrNumero 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   5850
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   14
            Top             =   180
            Width           =   1005
         End
         Begin VB.TextBox txt_Logradouro 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   13
            Top             =   180
            Width           =   4125
         End
         Begin VB.TextBox txt_Bairro 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   16
            Top             =   510
            Width           =   5535
         End
         Begin VB.TextBox txt_Municipio 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   19
            Top             =   840
            Width           =   5535
         End
         Begin VB.TextBox txt_Cep 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   17
            Top             =   510
            Width           =   1005
         End
         Begin VB.TextBox txt_UF 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   18
            Top             =   840
            Width           =   405
         End
         Begin VB.Label lbl_strComplemento 
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6960
            TabIndex        =   58
            Top             =   210
            Width           =   480
         End
         Begin VB.Label lbl_numero 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5520
            TabIndex        =   56
            Top             =   210
            Width           =   180
         End
         Begin VB.Label lbl_Logradouro 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   195
            Left            =   570
            TabIndex        =   47
            Top             =   210
            Width           =   690
         End
         Begin VB.Label lbl_Bairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   855
            TabIndex        =   46
            Top             =   540
            Width           =   405
         End
         Begin VB.Label lbl_Municipio 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   555
            TabIndex        =   45
            Top             =   870
            Width           =   705
         End
         Begin VB.Label lbl_Cep 
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   7080
            TabIndex        =   44
            Top             =   570
            Width           =   285
         End
         Begin VB.Label lbl_UF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   7170
            TabIndex        =   43
            Top             =   900
            Width           =   210
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Parcelas 
         Height          =   1545
         Left            =   180
         TabIndex        =   29
         Top             =   5670
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   2725
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
         Columns(1).Caption=   "Nº"
         Columns(1).DataField=   "intNumeroParcela"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Acordo"
         Columns(2).DataField=   "strAcordo"
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Vencimento"
         Columns(3).DataField=   "dtmVencimento"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Baixa"
         Columns(4).DataField=   "dtmDtPagamento"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Moeda"
         Columns(5).DataField=   "strMoeda"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Valor "
         Columns(6).DataField=   "dblValor"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Juros"
         Columns(7).DataField=   "dblJuros"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Multa"
         Columns(8).DataField=   "dblMulta"
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Correção"
         Columns(9).DataField=   "dblCorrecao"
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Total"
         Columns(10).DataField=   "dblTotal"
         Columns(10).NumberFormat=   "Standard"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Baixa - Descrição"
         Columns(11).DataField=   "strDescricao"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Baixa - Observação"
         Columns(12).DataField=   "strObservacao"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   13
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=13"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=476"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=397"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1905"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1826"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=1693"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1614"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=1746"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1667"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=1005"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=926"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=2143"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2064"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(7).Width=1984"
         Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1905"
         Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(49)=   "Column(8).Width=1905"
         Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1826"
         Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=2"
         Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(55)=   "Column(9).Width=2090"
         Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=2011"
         Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=2"
         Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(61)=   "Column(10).Width=2196"
         Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=2117"
         Splits(0)._ColumnProps(64)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(10)._ColStyle=2"
         Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(67)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(70)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(72)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(73)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(74)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(75)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(76)=   "Column(12).Order=13"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(16)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=78,.parent=13,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(9).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(70)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(10).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(74)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
         _StyleDefs(78)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(82)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(85)  =   "Named:id=33:Normal"
         _StyleDefs(86)  =   ":id=33,.parent=0"
         _StyleDefs(87)  =   "Named:id=34:Heading"
         _StyleDefs(88)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(89)  =   ":id=34,.wraptext=-1"
         _StyleDefs(90)  =   "Named:id=35:Footing"
         _StyleDefs(91)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(92)  =   "Named:id=36:Selected"
         _StyleDefs(93)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(94)  =   "Named:id=37:Caption"
         _StyleDefs(95)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(96)  =   "Named:id=38:HighlightRow"
         _StyleDefs(97)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(98)  =   "Named:id=39:EvenRow"
         _StyleDefs(99)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(100) =   "Named:id=40:OddRow"
         _StyleDefs(101) =   ":id=40,.parent=33"
         _StyleDefs(102) =   "Named:id=41:RecordSelector"
         _StyleDefs(103) =   ":id=41,.parent=34"
         _StyleDefs(104) =   "Named:id=42:FilterBar"
         _StyleDefs(105) =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_dblvlindexador 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Indexador"
         Height          =   195
         Left            =   -72930
         TabIndex        =   83
         Top             =   1590
         Width           =   1335
      End
      Begin VB.Label lbl_strindexador 
         AutoSize        =   -1  'True
         Caption         =   "Indexador"
         Height          =   195
         Left            =   -74700
         TabIndex        =   81
         Top             =   1590
         Width           =   705
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
      Height          =   1485
      Left            =   90
      TabIndex        =   48
      Top             =   7380
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   2619
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "PkidDativa"
      Columns(0).DataField=   "Pkid"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "PkitdAlfa"
      Columns(1).DataField=   "intAlfa"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Comp. Receita"
      Columns(2).DataField=   "strComposicao"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Inscrição"
      Columns(3).DataField=   "strInscricao"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Exercício"
      Columns(4).DataField=   "intExercicio"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Aviso"
      Columns(5).DataField=   "strNumeroAviso"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Proprietário"
      Columns(6).DataField=   "strNomeProprietario"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "intUtilizacao"
      Columns(7).DataField=   "intUtilizacao"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=5609"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5530"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=4180"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=4101"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=1349"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1270"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=2223"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2143"
      Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=7223"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=7144"
      Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(43)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(44)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(69)  =   "Named:id=33:Normal"
      _StyleDefs(70)  =   ":id=33,.parent=0"
      _StyleDefs(71)  =   "Named:id=34:Heading"
      _StyleDefs(72)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   ":id=34,.wraptext=-1"
      _StyleDefs(74)  =   "Named:id=35:Footing"
      _StyleDefs(75)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(76)  =   "Named:id=36:Selected"
      _StyleDefs(77)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(78)  =   "Named:id=37:Caption"
      _StyleDefs(79)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(80)  =   "Named:id=38:HighlightRow"
      _StyleDefs(81)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(82)  =   "Named:id=39:EvenRow"
      _StyleDefs(83)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(84)  =   "Named:id=40:OddRow"
      _StyleDefs(85)  =   ":id=40,.parent=33"
      _StyleDefs(86)  =   "Named:id=41:RecordSelector"
      _StyleDefs(87)  =   ":id=41,.parent=34"
      _StyleDefs(88)  =   "Named:id=42:FilterBar"
      _StyleDefs(89)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mobjAux                   As Object
Dim mblnClickOk               As Boolean
Dim mblnSelecionou            As Boolean
Dim mblnPrimeiraVez           As Boolean

Private Sub dbc_intReceita_Change()
    If dbc_intReceita.MatchedWithList = True And dbc_intReceita.BoundText <> "" Then
        PreencherListaDeOpcoes dbc_intReceita2, dbc_intReceita.BoundText
        PreencheCadastro CLng(dbc_intReceita.BoundText)
    Else
        txtcadastro.Text = ""
        txtcadastro2.Text = ""
        mskstrInscricao.Mask = ""
        mskstrInscricao.Text = ""
    End If
End Sub

Private Sub dbc_intReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intReceita, Me, Area
End Sub

Private Sub dbc_intReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intReceita, Me, , KeyCode, Shift
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnPrimeiraVez = True
    mblnClickOk = True
End Sub

Private Sub tdb_Parcelas_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If ColIndex = 2 And Trim(Value) <> "" Then
        Value = gstrFormataInscricao(Trim(Str(Value)), TYP_ACORDO)
    End If
End Sub

Private Sub tdb_Parcelas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not tdb_Parcelas.EOF And Not tdb_Parcelas.BOF Then
        gCorLinhaSelecionada tdb_Parcelas
    End If
End Sub

Private Sub txtcadastro_GotFocus()
    MarcaCampo txtcadastro
End Sub

Private Sub txtcadastro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtcadastro
End Sub

Private Sub txtdblvlindexador_GotFocus()
    MarcaCampo txtdblvlindexador
End Sub

Private Sub txtDblVlIndexador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblvlindexador
End Sub

Private Sub txtdblvlindexador_LostFocus()
    txtdblvlindexador = gstrConvVrDoSql(txtdblvlindexador, 6)
End Sub

Private Sub txtdtmdtinscricao_GotFocus()
    MarcaCampo txtdtmdtinscricao
End Sub

Private Sub txtdtmdtinscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdtinscricao
End Sub

Private Sub txtdtmdtinscricao_LostFocus()
    txtdtmdtinscricao = gstrDataFormatada(txtdtmdtinscricao)
End Sub

Private Sub txtintCertidao_GotFocus()
    MarcaCampo txtintcertidao
End Sub

Private Sub txtintcertidao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtintcertidao
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub txtintFolha_GotFocus()
    MarcaCampo txtintfolha
End Sub

Private Sub txtintfolha_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintfolha
End Sub

Private Sub txtintLivro_GotFocus()
    MarcaCampo txtintlivro
End Sub

Private Sub txtintlivro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintlivro
End Sub

Private Sub txtstrAviso_GotFocus()
    MarcaCampo txtstrAviso
End Sub

Private Sub txtstrAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrAviso
End Sub

Private Sub mskstrInscricao_GotFocus()
    MarcaCampo mskstrInscricao
End Sub

Private Sub mskstrInscricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricao
End Sub




Private Sub txtstrindexador_GotFocus()
    MarcaCampo txtstrindexador
End Sub

Private Sub txtstrindexador_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrindexador
End Sub

Private Sub txtstrNomeProprietario_GotFocus()
    MarcaCampo txtstrnomeproprietario
End Sub

Private Sub txtstrNomeProprietario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrnomeproprietario
End Sub

Private Sub txtstrPromissario_GotFocus()
    MarcaCampo txtstrpromissario
End Sub

Private Sub txtstrPromissario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrpromissario
End Sub

Private Sub dbc_intReceita_GotFocus()
    MarcaCampo dbc_intReceita
End Sub

Private Sub dbc_intReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intReceita
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1207

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
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    VerificaObjParaAplicar mobjAux
    
    TrocaCorObjeto dbc_intReceita2, True
    TrocaCorObjeto txtcadastro2, True
    TrocaCorObjeto txtstrIncricao2, True
    TrocaCorObjeto txtintExercicio2, True
    TrocaCorObjeto txtstrAviso2, True
    TrocaCorObjeto txtintfolha2, True
    TrocaCorObjeto txtintlivro2, True
    TrocaCorObjeto txtdtmdtinscricao2, True
    TrocaCorObjeto txtintcertidao2, True
    TrocaCorObjeto txtcadastro, True
    TrocaCorObjeto txtstrindexador, True
    TrocaCorObjeto txtdblvlindexador, True
    
    
    dbc_intReceita.Tag = strQueryComposicaoReceita & ";strDescricao"
    dbc_intReceita2.Tag = strQueryComposicaoReceita & ";strDescricao"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Function strQueryComposicaoReceita()
    Dim strsql As String
    
    strsql = ""
    strsql = strsql & "SELECT PKId, strDescricao "
    strsql = strsql & "FROM " & gstrComposicaoDaReceita & " "
    'strSQL = strSQL & "WHERE bytDividaAtiva = 0 "
    strsql = strsql & "ORDER BY strDescricao"
    
    strQueryComposicaoReceita = strsql
    
End Function

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    mblnClickOk = True
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
   gOrdenaGrid tdb_Lista, ColIndex
   mblnPrimeiraVez = False
   mblnClickOk = False
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If Not .EOF And Not .BOF Then
            If mblnPrimeiraVez Then
                If mblnClickOk Then
                    gCorLinhaSelecionada tdb_Lista
                    mblnClickOk = False
                    mblnSelecionou = True
                    txtPKId = .Columns(0).Value
                    gCorLinhaSelecionada tdb_Lista
                    PreencheCampos
                    LeDaTabelaParaObj "", tdb_Parcelas, strQueryParcela
                    If mobjAux Is Nothing Then
                        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                    Else
                        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(gstrImprimir) = UCase(strModoOperacao) Then

    ElseIf UCase(strModoOperacao) = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    ElseIf UCase(strModoOperacao) = gstrLocalizar Then
        LeDaTabelaParaObj "", tdb_Lista, strQuery(True)
    ElseIf UCase(strModoOperacao) = gstrRefresh Then
        LeDaTabelaParaObj "", tdb_Lista, strQuery(False)
    ElseIf UCase(gstrFechar) = UCase(strModoOperacao) Then
        Unload Me
    ElseIf UCase(gstrNovo) = UCase(strModoOperacao) Then
        Limpa_Controles Me, True, True, True, True, True
        Set tdb_Parcelas.DataSource = Nothing
        dbc_intReceita.SetFocus
        mskstrInscricao.Text = ""
    Else
        
    End If
End Sub

Private Sub PreencheCampos()
    Dim strsql As String
    Dim adoResultado As ADODB.Recordset
    
    strsql = ""
    strsql = strsql & "Select "
    strsql = strsql & "LA.Pkid, "
    strsql = strsql & "LA.Intcomposicaodareceita, "
    strsql = strsql & "CR.INTUTILIZACAO, "
    strsql = strsql & "lA.strInscricao, "
    strsql = strsql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, "
    strsql = strsql & "LA.intExercicio, "
    strsql = strsql & "A.intFolha, "
    strsql = strsql & "A.intlivro, "
    strsql = strsql & "A.dtmdtinscricao, "
    strsql = strsql & "A.strnomeproprietario, "
    strsql = strsql & "A.strcnpjcpf, "
    strsql = strsql & "A.stridentidade, "
    strsql = strsql & "A.strlogradouro, "
    strsql = strsql & "A.strnumero, "
    strsql = strsql & "A.strcomplemento, "
    strsql = strsql & "A.strbairro, "
    strsql = strsql & "A.strmunicipio, "
    strsql = strsql & "A.struf, "
    strsql = strsql & "A.intcep, "
    strsql = strsql & "A.strlogradouroc, "
    strsql = strsql & "A.strnumeroc, "
    strsql = strsql & "A.strcomplementoc, "
    strsql = strsql & "A.strbairroc, "
    strsql = strsql & "A.strmunicipioc, "
    strsql = strsql & "A.strufc, "
    strsql = strsql & "A.intcepc, "
    strsql = strsql & "A.strpromissario, "
    strsql = strsql & "A.strobservacao, "
    strsql = strsql & "A.strindexador, "
    strsql = strsql & "A.dblvlindexador, "
    strsql = strsql & "A.intcertidao, "
    strsql = strsql & "A.dblvalorimposto, "
    strsql = strsql & "A.dblvalortaxas "
    strsql = strsql & "From "
    strsql = strsql & gstrLancamentoAlfa & " LA, "
    strsql = strsql & gstrComposicaoDaReceita & " CR, "
    strsql = strsql & gstrDativa & " A "
    strsql = strsql & "Where "
    strsql = strsql & "LA.Pkid = A.Intlancamentoalfa AND "
    strsql = strsql & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strsql = strsql & "A.pkid = " & Trim(txtPKId)
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                
                PreencherListaDeOpcoes dbc_intReceita, gstrENulo(!intComposicaoDaReceita)
                Select Case gstrENulo(!intUtilizacao)
                    Case 1
                        txtcadastro = "Imobiliário"
                        txtcadastro2 = "Imobiliário"
                    Case 2
                        txtcadastro = "Econômico"
                        txtcadastro2 = "Econômico"
                    Case 3
                        txtcadastro = "Dívida Ativa"
                        txtcadastro2 = "Dívida Ativa"
                    Case 4
                        txtcadastro = "Acordo"
                        txtcadastro2 = "Acordo"
                    Case 5
                        txtcadastro = "Preco Público"
                        txtcadastro2 = "Preco Público"
                    Case Else
                        txtcadastro = ""
                        txtcadastro2 = ""
                End Select
                
                mskstrInscricao = gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao)
                txtintExercicio = gstrENulo(!intExercicio)
                txtstrAviso = gstrENulo(!strNumeroAviso)
                txtintfolha = gstrENulo(!intFolha)
                txtintlivro = gstrENulo(!intLivro)
                txtdtmdtinscricao = gstrDataFormatada(gstrENulo(!dtmdtinscricao))
                txtintcertidao = gstrENulo(!intCertidao)
                
                PreencherListaDeOpcoes dbc_intReceita2, gstrENulo(!intComposicaoDaReceita)
                txtstrIncricao2 = gstrFormataInscricao(Right(!strInscricao, gintRetornaTamanhoMascara(!intUtilizacao)), !intUtilizacao)
                txtintExercicio2 = gstrENulo(!intExercicio)
                txtstrAviso2 = gstrENulo(!strNumeroAviso)
                txtintfolha2 = gstrENulo(!intFolha)
                txtintlivro2 = gstrENulo(!intLivro)
                txtdtmdtinscricao2 = gstrDataFormatada(gstrENulo(!dtmdtinscricao))
                txtintcertidao2 = gstrENulo(!intCertidao)
                
                txtstrnomeproprietario = gstrENulo(!strnomeproprietario)
                txtstrcnpjcpf = gstrCGCCPFFormatado(gstrENulo(!StrCnpjCpf))
                txtstridentidade = gstrENulo(!STRIDENTIDADE)

                txt_Logradouro = gstrENulo(!strLogradouro)
                txtstrNumero = gstrENulo(!strNumero)
                txtstrComplemento = gstrENulo(!STRCOMPLEMENTO)
                txt_Bairro = gstrENulo(!strBairro)
                txt_Municipio = gstrENulo(!STRMUNICIPIO)
                txt_UF = gstrENulo(!STRUF)
                txt_Cep = gstrCEPFormatado(gstrENulo(!INTCEP))
                
                txt_LogradouroN = gstrENulo(!strlogradouroc)
                txtstrNumeroN = gstrENulo(!strNumeroC)
                txtstrComplementoN = gstrENulo(!strComplementoC)
                txt_BairroN = gstrENulo(!strBairroC)
                txt_MunicipioN = gstrENulo(!strMunicipioC)
                txt_UFN = gstrENulo(!strUFC)
                txt_CepN = gstrCEPFormatado(gstrENulo(!intcepc))
                
                txtstrpromissario = gstrENulo(!strpromissario)
                txtHistorico = gstrENulo(!strObservacao)
                txtstrindexador = gstrENulo(!Strindexador)
                txtdblvlindexador = gstrConvVrDoSql(gstrENulo(!dblvlIndexador), 2, , True)
                txt_dblvalorimposto = gstrConvVrDoSql(gstrENulo(!dblValorImposto), 2, , True)
                txt_dblvalortaxas = gstrConvVrDoSql(gstrENulo(!dblvalortaxas), 2, , True)
            End If
        End With
    End If
    
End Sub

Private Function strQuery(blnFiltro As Boolean) As String
    Dim strsql As String
    
    strsql = ""
    strsql = strsql & "Select "
    If bytDBType = Oracle Then
        strsql = strsql & "/*+ index(A) */ " 'Parâmetro adicional inserido para otimizar a consulta à pedido do DBA
    End If
    strsql = strsql & "A.Pkid, "
    strsql = strsql & "LA.Pkid AS intAlfa, "
    strsql = strsql & "CR.Strdescricao AS strComposicao, "
    strsql = strsql & "CR.intUtilizacao, "
    strsql = strsql & "LA.Strinscricao, "
    strsql = strsql & "LA.Intexercicio, "
    strsql = strsql & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, "
    strsql = strsql & "A.Strnomeproprietario "
    strsql = strsql & "from "
    strsql = strsql & gstrLancamentoAlfa & " LA, "
    strsql = strsql & gstrComposicaoDaReceita & " CR, "
    strsql = strsql & gstrDativa & " A "
    strsql = strsql & "WHERE "
    strsql = strsql & "CR.PKID = LA.Intcomposicaodareceita AND "
    'strSQL = strSQL & "bytDividaAtiva = 0 AND "
    strsql = strsql & "La.pkid = A.Intlancamentoalfa "
    
    If blnFiltro Then
        If dbc_intReceita.MatchedWithList = True Then strsql = strsql & " AND LA.intComposicaoDaReceita = " & dbc_intReceita.BoundText
        If Trim(txtcadastro) <> "" Then strsql = strsql & ""
        If Trim(mskstrInscricao) <> "" Then strsql = strsql & " AND LA.strInscricao ='" & (String(gintLenInscricao - Len(Trim(mskstrInscricao.Text)), "0") & Trim(mskstrInscricao.Text)) & "'"
        If Trim(txtintExercicio) <> "" Then strsql = strsql & " AND LA.intExercicio = " & txtintExercicio.Text
        If Trim(txtstrAviso) <> "" Then strsql = strsql & " AND LA.strNumeroAviso = '" & (String(gintLenNumAviso - Len(Trim(txtstrAviso)), "0") & Trim(txtstrAviso)) & "'"
        If Trim(txtdtmdtinscricao) <> "" Then strsql = strsql & " AND A.dtmdtinscricao = '" & txtdtmdtinscricao.Text & "'"
        If Trim(txtintcertidao) <> "" Then strsql = strsql & " AND A.intcertidao = " & txtintcertidao.Text
        If Trim(txtintfolha) <> "" Then strsql = strsql & " AND A.intfolha = " & txtintfolha.Text
        If Trim(txtintlivro) <> "" Then strsql = strsql & " AND A.intlivro = " & txtintlivro.Text
        If Trim(txtstrnomeproprietario) <> "" Then strsql = strsql & " AND UPPER(A.strnomeproprietario) Like '" & UCase(txtstrnomeproprietario.Text) & "%'"
        If Trim(txtstrpromissario) <> "" Then strsql = strsql & " AND UPPER(A.strpromissario) Like '" & UCase(txtstrpromissario.Text) & "%'"
    End If
    
    strsql = strsql & " ORDER BY strComposicao ASC, LA.Strinscricao ASC, LA.Intexercicio DESC"
    strQuery = strsql
    
End Function

Private Function strQueryParcela() As String
    Dim strsql As String
    
    strsql = strsql & "SELECT DP.Pkid, DP.INTPARCELA as intNumeroParcela, "
    strsql = strsql & "(SELECT Strabreviatura FROM " & gstrMoedas & " M WHERE M.pkid = DP.intMoeda) strMoeda, "
    strsql = strsql & "DP.DTMDTVENCIMENTO as dtmVencimento, "
    strsql = strsql & "(" & gstrISNULL("DP.DBLVALOR", "0") & "+" & gstrISNULL("DP.DBLJUROS", "0") & "+" & gstrISNULL("DP.DBLMULTA", "0") & "+" & gstrISNULL("DP.Dblcorrecaomonet", "0") & ") as dblTotal, "
    strsql = strsql & "DP.DBLVALOR as dblValor, "
    strsql = strsql & "DP.DBLJUROS as dblJuros, "
    strsql = strsql & "DP.DBLMULTA as dblMulta, "
    strsql = strsql & "DP.Dblcorrecaomonet as dblCorrecao, "
    strsql = strsql & gstrRIGHT("Ltrim(RTrim(LA.strInscricao))", gintRetornaTamanhoMascara(TYP_ACORDO)) & " strAcordo, "
    'strSQL = strSQL & "LA.strInscricao strAcordo, "
    strsql = strsql & "LP.DTMDTPAGAMENTO, "
    strsql = strsql & "LP.strObservacao, "
    strsql = strsql & "(SELECT strDescricao FROM tblCodigoBaixa CB WHERE CB.pkid = LP.intCodigoBaixa) strDescricao "
    strsql = strsql & "FROM " & gstrDativa & " DA, "
    strsql = strsql & gstrDaParcel & " DP, "
    strsql = strsql & gstrLancamentoValor & " LV, "
    strsql = strsql & gstrLancamentoAlfa & " LA, "
    strsql = strsql & gstrLancamentoPagamento & " LP "
    strsql = strsql & "WHERE Da.Pkid = DP.Intdativa AND "
    strsql = strsql & "LV.Intlancamentoalfa = DA.INTLANCAMENTOALFA AND "
    strsql = strsql & "DA.Pkid = " & Trim(txtPKId) & " AND "
    strsql = strsql & "LA.Pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " LV.INTLANCAMENTOALFAACORDO AND "
    strsql = strsql & "LV.Intparcela = DP.INTPARCELA AND "
    strsql = strsql & "LP.INTLANCAMENTOVALOR " & strOUTJOracle & " =" & strOUTJSQLServer & " LV.PKID "
    strsql = strsql & " Order By DP.intParcela "
    
    strQueryParcela = strsql
    
End Function

Private Sub txtstridentidade_GotFocus()
    MarcaCampo txtstridentidade
End Sub

Private Sub txtstrIdentidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstridentidade
End Sub

Private Sub txtstrcnpjcpf_GotFocus()
    MarcaCampo txtstrcnpjcpf
End Sub

Private Sub txtstrcnpjcpf_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrcnpjcpf
End Sub

Private Sub PreencheCadastro(lngPkid As Long)
    Dim strsql          As String
    Dim adoResultado    As ADODB.Recordset
     
    strsql = ""
    strsql = strsql & "Select * From " & gstrComposicaoDaReceita & " Where pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Select Case gstrENulo(adoResultado!intUtilizacao)
                Case 1
                    txtcadastro = "Imobiliário"
                    txtcadastro2 = "Imobiliário"
                Case 2
                    txtcadastro = "Econômico"
                    txtcadastro2 = "Econômico"
                Case 3
                    txtcadastro = "Dívida Ativa"
                    txtcadastro2 = "Dívida Ativa"
                Case 4
                    txtcadastro = "Acordo"
                    txtcadastro2 = "Acordo"
                Case 5
                    txtcadastro = "Preco Público"
                    txtcadastro2 = "Preco Público"
                Case Else
                    txtcadastro = ""
                    txtcadastro2 = ""
            End Select
            VerificaMascaraInscricao CInt(gstrENulo(adoResultado!intUtilizacao))
        End If
    End If
End Sub

Sub VerificaMascaraInscricao(intTipoComposicao As Integer)
Dim strsql As String
Dim adoResultado As ADODB.Recordset
Dim strMascara   As String
    
    strMascara = ""
    strsql = ""
    strsql = strsql & "Select * From " & gstrCampoDeInscricao & " "
    strsql = strsql & "Where intTipoDeInscricao = " & intTipoComposicao
    strsql = strsql & " Order By intSequencia"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strsql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    mskstrInscricao.Mask = strMascara
    
End Sub

