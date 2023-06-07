VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmIssNotaFiscal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Fiscais de ISS"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   9465
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   7665
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13520
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Nota Fiscal de ISS"
      TabPicture(0)   =   "frmIssNotaFiscal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strInscricaoCadastral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_strRazaoSocial"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_strNomeFantasia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_strAtividadePrincipal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_strISSTipo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_strListaServico"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt_strInscricaoCadastral"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tdb_NotasFiscais"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tdb_Lista"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_Endereco"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_strRazaoSocial"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_strNomeFantasia"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_strAtividadePrincipal"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_strISSTipo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_strListaServico"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "fra_NotasFiscais"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.Frame fra_NotasFiscais 
         Caption         =   "Notas Fiscais"
         Height          =   1335
         Left            =   180
         TabIndex        =   28
         Top             =   3510
         Width           =   8925
         Begin VB.TextBox txtPKId 
            Height          =   285
            Left            =   0
            TabIndex        =   50
            Top             =   -120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtstrObservacao 
            Height          =   285
            Left            =   1020
            MaxLength       =   160
            TabIndex        =   46
            Top             =   960
            Width           =   7785
         End
         Begin VB.TextBox txtdtmDtCancelamento 
            Height          =   285
            Left            =   7770
            TabIndex        =   44
            Top             =   615
            Width           =   1035
         End
         Begin VB.TextBox txtdtmDtNotaFiscalBaixa 
            Height          =   285
            Left            =   5250
            TabIndex        =   42
            Top             =   615
            Width           =   1035
         End
         Begin VB.TextBox txtdblNotaFiscalValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2970
            TabIndex        =   40
            Top             =   615
            Width           =   1365
         End
         Begin VB.TextBox txtdtmDtLimite 
            Height          =   285
            Left            =   1020
            TabIndex        =   38
            Top             =   615
            Width           =   1035
         End
         Begin VB.TextBox txtstrNotaFiscalSerie 
            Height          =   285
            Left            =   7770
            TabIndex        =   36
            Top             =   285
            Width           =   1035
         End
         Begin VB.TextBox txtstrNotaFiscalNr 
            Height          =   285
            Left            =   5250
            TabIndex        =   34
            Top             =   285
            Width           =   1035
         End
         Begin VB.TextBox txtdtmDtBase 
            Height          =   285
            Left            =   2970
            TabIndex        =   32
            Top             =   285
            Width           =   1035
         End
         Begin VB.TextBox txtintControleNr 
            Height          =   285
            Left            =   1020
            TabIndex        =   30
            Top             =   285
            Width           =   1035
         End
         Begin VB.Label lblstrObservacao 
            AutoSize        =   -1  'True
            Caption         =   "Observação"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   1005
            Width           =   870
         End
         Begin VB.Label lbldtmDtCancelamento 
            AutoSize        =   -1  'True
            Caption         =   "Data Cancelamento"
            Height          =   195
            Left            =   6330
            TabIndex        =   43
            Top             =   660
            Width           =   1410
         End
         Begin VB.Label lbldtmDtNotaFiscalBaixa 
            AutoSize        =   -1  'True
            Caption         =   "Data Baixa"
            Height          =   195
            Left            =   4410
            TabIndex        =   41
            Top             =   660
            Width           =   780
         End
         Begin VB.Label lbldblNotaFiscalValor 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   2160
            TabIndex        =   39
            Top             =   660
            Width           =   360
         End
         Begin VB.Label lbldtmDtLimite 
            AutoSize        =   -1  'True
            Caption         =   "Data Limite"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   660
            Width           =   795
         End
         Begin VB.Label lblstrNotaFiscalSerie 
            AutoSize        =   -1  'True
            Caption         =   "Série"
            Height          =   195
            Left            =   6330
            TabIndex        =   35
            Top             =   330
            Width           =   360
         End
         Begin VB.Label lblstrNotaFiscalNr 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   4410
            TabIndex        =   33
            Top             =   330
            Width           =   180
         End
         Begin VB.Label lbldtmDtBase 
            AutoSize        =   -1  'True
            Caption         =   "Data Base"
            Height          =   195
            Left            =   2130
            TabIndex        =   31
            Top             =   330
            Width           =   750
         End
         Begin VB.Label lblintControleNr 
            AutoSize        =   -1  'True
            Caption         =   "Controle"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   330
            Width           =   585
         End
      End
      Begin VB.TextBox txt_strListaServico 
         Height          =   285
         Left            =   5280
         TabIndex        =   12
         Top             =   1740
         Width           =   1815
      End
      Begin VB.TextBox txt_strISSTipo 
         Height          =   285
         Left            =   1830
         TabIndex        =   10
         Top             =   1740
         Width           =   1965
      End
      Begin VB.TextBox txt_strAtividadePrincipal 
         Height          =   285
         Left            =   1830
         TabIndex        =   8
         Top             =   1410
         Width           =   6285
      End
      Begin VB.TextBox txt_strNomeFantasia 
         Height          =   285
         Left            =   1830
         TabIndex        =   6
         Top             =   1080
         Width           =   6285
      End
      Begin VB.TextBox txt_strRazaoSocial 
         Height          =   285
         Left            =   1830
         TabIndex        =   4
         Top             =   750
         Width           =   6285
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   "Endereço"
         Height          =   1395
         Left            =   180
         TabIndex        =   13
         Top             =   2100
         Width           =   8925
         Begin VB.TextBox txt_UF 
            Height          =   285
            Left            =   1020
            TabIndex        =   25
            Top             =   990
            Width           =   510
         End
         Begin VB.TextBox txt_Municipio 
            Height          =   285
            Left            =   5070
            TabIndex        =   23
            Top             =   630
            Width           =   3015
         End
         Begin VB.TextBox txt_Cep 
            Height          =   285
            Left            =   1995
            TabIndex        =   27
            Top             =   990
            Width           =   1080
         End
         Begin VB.TextBox txt_Complemento 
            Height          =   285
            Left            =   6795
            TabIndex        =   19
            Top             =   270
            Width           =   1290
         End
         Begin VB.TextBox txt_Numero 
            Height          =   285
            Left            =   5400
            TabIndex        =   17
            Top             =   270
            Width           =   795
         End
         Begin VB.TextBox txt_Logradouro 
            Height          =   285
            Left            =   1020
            TabIndex        =   15
            Top             =   270
            Width           =   4005
         End
         Begin VB.TextBox txt_Bairro 
            Height          =   285
            Left            =   1020
            TabIndex        =   21
            Top             =   630
            Width           =   3105
         End
         Begin VB.Label lbl_intMunicipioC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   4290
            TabIndex        =   22
            Top             =   690
            Width           =   705
         End
         Begin VB.Label lbl_intBairroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   525
            TabIndex        =   20
            Top             =   690
            Width           =   405
         End
         Begin VB.Label lbl_intLogradouroC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lbl_intNumeroC 
            AutoSize        =   -1  'True
            Caption         =   "Nº"
            Height          =   195
            Left            =   5145
            TabIndex        =   16
            Top             =   330
            Width           =   180
         End
         Begin VB.Label lbl_strComplementoC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl."
            Height          =   195
            Left            =   6270
            TabIndex        =   18
            Top             =   330
            Width           =   480
         End
         Begin VB.Label lbl_intUFC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   720
            TabIndex        =   24
            Top             =   1065
            Width           =   210
         End
         Begin VB.Label lbl_intCepC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cep"
            Height          =   195
            Left            =   1620
            TabIndex        =   26
            Top             =   1050
            Width           =   285
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   1365
         Left            =   180
         TabIndex        =   48
         Top             =   6150
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2408
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
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Razão Social"
         Columns(2).DataField=   "strNome"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Atividade Principal"
         Columns(3).DataField=   "strAtividade"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nome Fantasia"
         Columns(4).DataField=   "strNomeFantasia"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "ISS Tipo"
         Columns(5).DataField=   "strTipoIss"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Lista Serviço"
         Columns(6).DataField=   "strListaServico"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
         Splits(0)._ColumnProps(12)=   "Column(2).Width=5636"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=5556"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=6456"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=6376"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(34)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(41)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
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
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_NotasFiscais 
         Height          =   1155
         Left            =   180
         TabIndex        =   47
         Top             =   4920
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Controle"
         Columns(1).DataField=   "intControleNr"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data Base"
         Columns(2).DataField=   "dtmDtBase"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Número N.F."
         Columns(3).DataField=   "strNotaFiscalNr"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Série N.F."
         Columns(4).DataField=   "strNotaFiscalSerie"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Data Limite"
         Columns(5).DataField=   "dtmDtLimite"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Valor"
         Columns(6).DataField=   "dblNotaFiscalValor"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Data Baixa"
         Columns(7).DataField=   "dtmDtNotaFiscalBaixa"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Data Cancelamento"
         Columns(8).DataField=   "dtmdtcancelamento"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Observação"
         Columns(9).DataField=   "strObservacao"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=4101"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4022"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=3731"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3651"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=4022"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3942"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=3307"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3228"
         Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(30)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(39)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(41)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(45)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(46)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(48)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(49)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(51)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(52)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(53)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(55)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(60)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
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
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Named:id=33:Normal"
         _StyleDefs(77)  =   ":id=33,.parent=0"
         _StyleDefs(78)  =   "Named:id=34:Heading"
         _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   ":id=34,.wraptext=-1"
         _StyleDefs(81)  =   "Named:id=35:Footing"
         _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(83)  =   "Named:id=36:Selected"
         _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=37:Caption"
         _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(87)  =   "Named:id=38:HighlightRow"
         _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(89)  =   "Named:id=39:EvenRow"
         _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(91)  =   "Named:id=40:OddRow"
         _StyleDefs(92)  =   ":id=40,.parent=33"
         _StyleDefs(93)  =   "Named:id=41:RecordSelector"
         _StyleDefs(94)  =   ":id=41,.parent=34"
         _StyleDefs(95)  =   "Named:id=42:FilterBar"
         _StyleDefs(96)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox txt_strInscricaoCadastral 
         Height          =   285
         Left            =   1830
         TabIndex        =   2
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_strListaServico 
         AutoSize        =   -1  'True
         Caption         =   "ISS Lista Serviço"
         Height          =   195
         Left            =   3990
         TabIndex        =   11
         Top             =   1785
         Width           =   1215
      End
      Begin VB.Label lbl_strISSTipo 
         AutoSize        =   -1  'True
         Caption         =   "ISS Tipo"
         Height          =   195
         Left            =   315
         TabIndex        =   9
         Top             =   1785
         Width           =   615
      End
      Begin VB.Label lbl_strAtividadePrincipal 
         AutoSize        =   -1  'True
         Caption         =   "Atividade Principal"
         Height          =   195
         Left            =   315
         TabIndex        =   7
         Top             =   1470
         Width           =   1305
      End
      Begin VB.Label lbl_strNomeFantasia 
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia"
         Height          =   195
         Left            =   315
         TabIndex        =   5
         Top             =   1125
         Width           =   1065
      End
      Begin VB.Label lbl_strRazaoSocial 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social"
         Height          =   195
         Left            =   315
         TabIndex        =   3
         Top             =   795
         Width           =   945
      End
      Begin VB.Label lbl_strInscricaoCadastral 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   315
         TabIndex        =   1
         Top             =   465
         Width           =   1350
      End
   End
   Begin VB.TextBox txt_PkidEc 
      Height          =   285
      Left            =   60
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   765
   End
End
Attribute VB_Name = "frmIssNotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnAlterando        As Boolean
Private mblnPrimeiraVezNF    As Boolean
Private mblnPrimeiraVezLista As Boolean

Private Function blnDadosOK() As Boolean

    blnDadosOK = False
    
    If Len(Trim$(txt_PkidEc)) = 0 Then
        ExibeMensagem "Selecione uma inscrição."
        txt_strInscricaoCadastral.SetFocus
        Exit Function
    End If
    
    If Len(Trim$(txtPKId)) = 0 Then
        ExibeMensagem "Selecione uma nota fiscal."
        txtdblNotaFiscalValor.SetFocus
        Exit Function
    End If
    
'    If Len(Trim$(txtdblNotaFiscalValor)) = 0 Then
'        ExibeMensagem "É necessário informar um valor para a nota fiscal."
'        txtdblNotaFiscalValor.SetFocus
'        Exit Function
'    End If
'
'    If Len(Trim$(txtdtmDtNotaFiscalBaixa)) = 0 Then
'        ExibeMensagem "É necessário informar uma data de baixa para a nota fiscal."
'        txtdtmDtNotaFiscalBaixa.SetFocus
'        Exit Function
'    End If
    
    blnDadosOK = True
    
End Function

Private Sub LimpaCampos()

    txt_PkidEc = Space$(0)
    txt_strInscricaoCadastral = Space$(0)
    txt_strRazaoSocial = Space$(0)
    txt_strNomeFantasia = Space$(0)
    txt_strAtividadePrincipal = Space$(0)
    txt_strISSTipo = Space$(0)
    txt_strListaServico = Space$(0)
    
    txt_Logradouro = Space$(0)
    txt_Numero = Space$(0)
    txt_Bairro = Space$(0)
    txt_Municipio = Space$(0)
    txt_UF = Space$(0)
    txt_Cep = Space$(0)
    txt_Complemento = Space$(0)
    
    Set tdb_NotasFiscais.DataSource = Nothing

    txt_strInscricaoCadastral.SetFocus
    
End Sub

Private Sub LimpaCamposNota()

    txtPKId = Space$(0)
    txtintControleNr = Space$(0)
    txtdtmDtBase = Space$(0)
    txtstrNotaFiscalNr = Space$(0)
    txtstrNotaFiscalSerie = Space$(0)
    txtdtmDtLimite = Space$(0)
    txtdblNotaFiscalValor = Space$(0)
    txtdtmDtNotaFiscalBaixa = Space$(0)
    txtdtmDtCancelamento = Space$(0)
    txtstrObservacao = Space$(0)
    
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim intLinha As Integer

    Select Case strModoOperacao
        
        Case gstrNovo
            LimpaCampos
            LimpaCamposNota
            HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar, gstrImprimir
            
        Case gstrSalvar
        
            If blnDadosOK Then
                intLinha = tdb_NotasFiscais.Bookmark
                ToolBarGeral strModoOperacao, "tblnotafiscaliss", mblnAlterando, tdb_NotasFiscais, Me, , strQueryNotasFiscais, , , , False
                LeDaTabelaParaObj "", tdb_NotasFiscais, strQueryNotasFiscais
                tdb_NotasFiscais.Bookmark = intLinha
            End If
            
        Case gstrDeletar
            If tdb_NotasFiscais.ApproxCount > 0 Then
                ToolBarGeral strModoOperacao, "tblnotafiscaliss", mblnAlterando, tdb_NotasFiscais, Me, , strQueryNotasFiscais
                LeDaTabelaParaObj "", tdb_NotasFiscais, strQueryNotasFiscais
            End If
            
        Case gstrLocalizar
            mblnPrimeiraVezLista = True
            LeDaTabelaParaObj "", tdb_Lista, strQueryEconomico
            
        Case gstrPreencherLista
        
        Case gstrImprimir
    
    End Select

End Sub

Private Sub strPreencheEndereco()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
    
    If Len(Trim$(txt_PkidEc)) > 0 Then
    
'        strSQL = ""
'        strSQL = strSQL & "SELECT "
'        strSQL = strSQL & " Ltrim(Rtrim(CO.strnome)) strnome "
'        strSQL = strSQL & "FROM " & gstrEconomico & " EC, "
'        strSQL = strSQL & gstrContribuinte & " CO "
'        strSQL = strSQL & " Where EC.intContribuinte = CO.Pkid And EC.Pkid = " & dbc_strInscricao.BoundText
        
        strSQL = _
        "select " & _
            "tl.strdescricao strtitulo, " & _
            "lo.strdescricao strlogradouro, " & _
            "ba.strdescricao strbairro, " & _
            "ec.intnumero, " & _
            "mu.strdescricao strmunicipio, " & _
            "uf.strsigla struf, " & _
            "ec.intcep, " & _
            "ec.strcomplemento "
        strSQL = strSQL & _
        "from " & _
            gstrEconomico & " ec, " & _
            gstrContribuinte & " co, " & _
            gstrLogradouro & " lo, " & _
            gstrTituloLogradouro & " tl, " & _
            gstrBairro & " ba, " & _
            gstrCidade & " mu, " & _
            gstrUF & " uf, " & _
            gstrFormaDeComunicacao & " fc "
        strSQL = strSQL & _
        "Where " & _
            "ec.intcontribuinte = co.pkid and " & _
            "ec.intlogradouro = lo.pkid and " & _
            "lo.inttitulologradouro " & strOUTJSQLServer & "= tl.pkid " & strOUTJOracle & " and " & _
            "ec.intbairro = ba.pkid and " & _
            "ba.intmunicipio = mu.pkid and " & _
            "mu.intuf = uf.pkid and " & _
            "co.pkid " & strOUTJSQLServer & "= fc.intcontribuinte " & strOUTJOracle & " and " & _
            "ec.pkid = " & txt_PkidEc
    
        Set gobjBanco = New clsBanco
        
        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
            
            If Not adoResultado.EOF Then
                txt_Logradouro = gstrENulo(adoResultado!strTitulo) & " " & gstrENulo(adoResultado!strLogradouro)
                txt_Numero = gstrENulo(adoResultado!INTNUMERO)
                txt_Bairro = gstrENulo(adoResultado!strBairro)
                txt_Municipio = gstrENulo(adoResultado!STRMUNICIPIO)
                txt_UF = gstrENulo(adoResultado!STRUF)
                txt_Cep = gstrCEPFormatado(gstrENulo(adoResultado!INTCEP))
                txt_Complemento = gstrENulo(adoResultado!STRCOMPLEMENTO)
            Else
                txt_Logradouro = Space$(0)
                txt_Numero = Space$(0)
                txt_Bairro = Space$(0)
                txt_Municipio = Space$(0)
                txt_UF = Space$(0)
                txt_Cep = Space$(0)
                txt_Complemento = Space$(0)
            End If
        End If
    
    End If
    
End Sub

Private Function strQueryEconomico() As String
Dim strSQL As String

    strSQL = "select " & _
        "ec.pkid, " & _
        gstrRIGHT("ec.strinscricaocadastral", gintRetornaTamanhoMascara(TYP_ECONOMICA)) & " strinscricaocadastral, " & _
        "co.strnome, " & _
        "co.strnomefantasia, " & _
        "ac.strdescricao stratividade, " & _
        "ti.strdescricao strtipoiss, " & _
        "ls.strDescricao strlistaservico " & _
    "from " & _
        gstrEconomico & " ec, " & _
        gstrContribuinte & " co, " & _
        gstrAtividadeDaEmpresa & " ae, " & _
        gstrAtividadeEC & " ac, " & _
        gstrTipoIss & " ti, " & _
        gstrListaServico & " ls " & _
    "Where " & _
        "co.pkid = ec.intcontribuinte and " & _
        "ae.inteconomico = ec.pkid and " & _
        "ac.pkid = ae.intatividade and " & _
        "ti.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " ec.inttipoiss and " & _
        "ls.pkid " & strOUTJOracle & " =" & strOUTJSQLServer & " ec.intlistaservico and " & _
        "ae.blnPrincipal = 1 "

    If Len(Trim$(txt_strInscricaoCadastral)) > 0 Then
        strSQL = strSQL & " and ec.strinscricaocadastral = '" & String$(20 - Len(txt_strInscricaoCadastral), "0") & txt_strInscricaoCadastral & "'"
    End If
    
    If Len(Trim$(txt_strRazaoSocial)) > 0 Then
        strSQL = strSQL & " and co.strnome like '" & txt_strRazaoSocial & "%'"
    End If
    
    strQueryEconomico = strSQL
    
End Function

Private Function strQueryNotasFiscais() As String
Dim strSQL As String

    strSQL = "select " & _
        "nf.* " & _
    "from " & _
        gstrEconomico & " ec, " & _
        "tblnotafiscaliss nf " & _
    "Where " & _
        "nf.inteconomico = ec.pkid and " & _
        "ec.pkid = " & txt_PkidEc
    
    strQueryNotasFiscais = strSQL
    
End Function

Private Sub VerificaMascaraInscricao()
Dim strSQL       As String
Dim adoResultado As ADODB.Recordset
Dim strMascara   As String
    
    strMascara = ""
    
    strSQL = ""
    strSQL = strSQL & "Select * From " & gstrCampoDeInscricao & " "
    strSQL = strSQL & "Where intTipoDeInscricao = " & TYP_ECONOMICA
    strSQL = strSQL & "Order By intSequencia"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    txt_strInscricaoCadastral.Mask = strMascara
    
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1425
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar, gstrDeletar, gstrImprimir
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar, gstrImprimir
End Sub

Private Sub Form_Load()

'    TrocaCorObjeto txt_strInscricaoCadastral, True
'    TrocaCorObjeto txt_strRazaoSocial, True
    TrocaCorObjeto txt_strNomeFantasia, True
    TrocaCorObjeto txt_strAtividadePrincipal, True
    TrocaCorObjeto txt_strISSTipo, True
    TrocaCorObjeto txt_strListaServico, True
    
    TrocaCorObjeto txt_Logradouro, True
    TrocaCorObjeto txt_Bairro, True
    TrocaCorObjeto txt_Complemento, True
    TrocaCorObjeto txt_Municipio, True
    TrocaCorObjeto txt_Numero, True
    TrocaCorObjeto txt_Cep, True
    TrocaCorObjeto txt_UF, True
    
    TrocaCorObjeto txtintControleNr, True
    TrocaCorObjeto txtdtmDtBase, True
    TrocaCorObjeto txtstrNotaFiscalNr, True
    TrocaCorObjeto txtstrNotaFiscalSerie, True
    TrocaCorObjeto txtdtmDtLimite, True
        
    mblnAlterando = False
    mblnPrimeiraVezNF = True
    mblnPrimeiraVezLista = True
    
    VerificaMascaraInscricao
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar, gstrDeletar, gstrImprimir
End Sub

Private Sub tdb_Lista_Click()
    
    mblnPrimeiraVezLista = True
    mblnPrimeiraVezNF = True
    
    LimpaCamposNota
    
End Sub

Private Sub tdb_Lista_FilterChange()
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Value = gstrFormataInscricao(CStr(Value), TYP_ECONOMICA)
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Lista
        
        If Not .EOF And Not .BOF Then
            
            txt_PkidEc.Text = .Columns("PKID").Value

            If mblnPrimeiraVezLista Then
                
                mblnPrimeiraVezLista = False
                
                txt_strInscricaoCadastral = .Columns("strinscricaocadastral").Value
                txt_strRazaoSocial = .Columns("strnome").Value
                txt_strNomeFantasia = .Columns("strnomefantasia").Value
                txt_strAtividadePrincipal = .Columns("strAtividade").Value
                txt_strISSTipo = .Columns("strtipoiss").Value
                txt_strListaServico = .Columns("strlistaservico").Value
                
                strPreencheEndereco
                
                If Len(Trim$(txt_PkidEc)) > 0 Then
                    LeDaTabelaParaObj "", tdb_NotasFiscais, strQueryNotasFiscais
                Else
                    Set tdb_NotasFiscais.DataSource = Nothing
                End If
                
                gCorLinhaSelecionada tdb_Lista

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar

                mblnAlterando = True
            
            End If

        End If
    
    End With

End Sub

Private Sub txt_strInscricaoCadastral_GotFocus()
    MarcaCampo txt_strInscricaoCadastral
End Sub

Private Sub txt_strInscricaoCadastral_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strInscricaoCadastral
End Sub

Private Sub txt_strRazaoSocial_GotFocus()
    MarcaCampo txt_strRazaoSocial
End Sub

Private Sub txt_strRazaoSocial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strRazaoSocial
End Sub

Private Sub txt_strNomeFantasia_Click()
    MarcaCampo txt_strNomeFantasia
End Sub

Private Sub txt_strNomeFantasia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strNomeFantasia
End Sub

Private Sub txt_strAtividadePrincipal_Click()
    MarcaCampo txt_strAtividadePrincipal
End Sub

Private Sub txt_strAtividadePrincipal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strAtividadePrincipal
End Sub

Private Sub txt_strISSTipo_Click()
    MarcaCampo txt_strISSTipo
End Sub

Private Sub txt_strISSTipo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strISSTipo
End Sub

Private Sub txt_strListaServico_Click()
    MarcaCampo txt_strListaServico
End Sub

Private Sub txt_strListaServico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_strListaServico
End Sub

Private Sub txtdblNotaFiscalValor_GotFocus()
    MarcaCampo txtdblNotaFiscalValor
End Sub

Private Sub txtdblNotaFiscalValor_LostFocus()
    txtdblNotaFiscalValor = gstrConvVrDoSql(txtdblNotaFiscalValor)
End Sub

Private Sub txtdtmDtCancelamento_GotFocus()
    MarcaCampo txtdtmDtCancelamento
End Sub

Private Sub txtdtmDtCancelamento_LostFocus()
    txtdtmDtCancelamento = gstrDataFormatada(txtdtmDtCancelamento)
End Sub

Private Sub txtdtmDtNotaFiscalBaixa_GotFocus()
    MarcaCampo txtdtmDtNotaFiscalBaixa
End Sub

Private Sub txtdtmDtNotaFiscalBaixa_LostFocus()
    txtdtmDtNotaFiscalBaixa = gstrDataFormatada(txtdtmDtNotaFiscalBaixa)
End Sub

Private Sub txtintControleNr_Click()
    MarcaCampo txtintControleNr
End Sub

Private Sub txtintControleNr_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtintControleNr
End Sub

Private Sub txtdtmDtBase_Click()
    MarcaCampo txtdtmDtBase
End Sub

Private Sub txtdtmDtBase_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtdtmDtBase
End Sub

Private Sub txtstrNotaFiscalNr_GotFocus()
    MarcaCampo txtstrNotaFiscalNr
End Sub

Private Sub txtstrNotaFiscalNr_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNotaFiscalNr
End Sub

Private Sub txtstrNotaFiscalSerie_Click()
    MarcaCampo txtstrNotaFiscalSerie
End Sub

Private Sub txtstrNotaFiscalSerie_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrNotaFiscalSerie
End Sub

Private Sub txtdtmDtLimite_Click()
    MarcaCampo txtdtmDtLimite
End Sub

Private Sub txtdtmDtLimite_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtdtmDtLimite
End Sub

Private Sub txtdblNotaFiscalValor_Click()
    MarcaCampo txtdblNotaFiscalValor
End Sub

Private Sub txtdblNotaFiscalValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblNotaFiscalValor
End Sub

Private Sub txtdtmDtNotaFiscalBaixa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtNotaFiscalBaixa
End Sub

Private Sub txtdtmDtCancelamento_Click()
    MarcaCampo txtdtmDtCancelamento
End Sub

Private Sub txtdtmDtCancelamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtCancelamento
End Sub

Private Sub txtstrObservacao_Click()
    MarcaCampo txtstrObservacao
End Sub

Private Sub txtstrObservacao_GotFocus()
    MarcaCampo txtstrObservacao
End Sub

Private Sub txtstrObservacao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrObservacao
End Sub

Private Sub tdb_NotasFiscais_Click()
    mblnPrimeiraVezNF = True
End Sub

Private Sub tdb_NotasFiscais_FilterChange()
    gblnFilraCampos tdb_NotasFiscais
End Sub

Private Sub tdb_NotasFiscais_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_NotasFiscais, ColIndex
End Sub

Private Sub tdb_NotasFiscais_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_NotasFiscais
        
        If Not .EOF And Not .BOF Then
            
            txtPKId.Text = .Columns("PKID").Value

            If mblnPrimeiraVezNF Then
                
                mblnPrimeiraVezNF = False
                
                txtintControleNr = .Columns("intcontrolenr").Value
                txtdtmDtBase = .Columns("dtmdtbase").Value
                txtstrNotaFiscalNr = .Columns("strnotafiscalnr").Value
                txtstrNotaFiscalSerie = .Columns("strnotafiscalserie").Value
                txtdtmDtLimite = .Columns("dtmdtlimite").Value
                txtdblNotaFiscalValor = gstrConvVrDoSql(.Columns("dblnotafiscalvalor").Value)
                txtdtmDtNotaFiscalBaixa = .Columns("dtmdtnotafiscalbaixa").Value
                txtdtmDtCancelamento = .Columns("dtmDtCancelamento").Value
                txtstrObservacao = .Columns("strObservacao").Value
                
                gCorLinhaSelecionada tdb_NotasFiscais

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar, gstrDeletar

                mblnAlterando = True
            
            End If

        End If
    
    End With

End Sub

